/* AK Excel Workbook - main.js (StartupScript)
 * Builds the UI shell, runs the Univer spreadsheet, and talks to AL.
 *
 * Flow:    main.js loads -> ControlAddInReady -> AL Init(config) -> Univer created -> AddInFrameReady
 *          -> AL may send a saved workbook with ReceiveChunk('load', ...).
 * Toolbar: Save, Refresh, Get data, Combine, Import, Export all live inside the add-in, so the BC
 *          page needs no fields or actions and the spreadsheet gets the whole page.
 * Data:    Get data panel -> OnTablesRequested / OnFieldsRequested / OnDataRequested -> AL answers with
 *          SetTables / SetFields / AppendData (paged) -> rows written to a sheet linked to that source.
 * Combine: joins two sheets in the browser (left/inner/full/differences, sum/first/all matches) into a
 *          new sheet that is re-calculated on every Refresh.
 * Reconcile: matches books against a statement / GSTR-2B / party ledger with recon.js, writes an
 *          exceptions-first sheet plus a live summary sheet (COUNTIF/SUMIF formulas).
 * Edit:    a sheet loaded with "Allow editing" tracks changes against the loaded rows; Save to BC sends
 *          them to AL (purpose 'write'), which validates each field and answers per row ('writeresult').
 */
(function () {
  'use strict';

  var H = window.AKHelper;
  var MUTATION = 2; // Univer CommandType.MUTATION
  var NUMERIC_TYPES = { Decimal: 1, Integer: 1, BigInteger: 1, Number: 1 };

  var state = {
    config: { defaultMaxRows: 10000, maxRowsLimit: 100000, chunkSize: 400000, workbookCode: '', description: '' },
    univer: null,
    api: null,
    disposables: [],
    resizeObserver: null,
    sources: {},            // sheetId -> BC data source (table, fields, filters)
    joins: {},              // output sheetId -> combine definition
    recons: {},             // output sheetId -> reconciliation definition
    recon: { keys: [], left: [], right: [], preset: 'gst' },
    autosave: true,
    autosaveTimer: null,
    importMode: 'replace',
    pending: {},            // requestId -> { source, columns, rows, resolve, reject }
    writes: {},             // write requestId -> { source, changes, resolve }
    applying: 0,            // > 0 while the add-in itself changes cells (not a user edit)
    requestSeq: 0,
    tables: [],
    table: null,            // { tableNo, tableCaption, canRead, fields[] }
    fieldState: {},         // fieldNo -> { checked, filter }
    fieldSearch: '',
    combine: { keys: [{ left: '', right: '' }], rightCols: {} },
    dirty: false,
    suppressDirty: 0,
    busy: false,
    assembler: new H.ChunkAssembler(),
    destroyed: false
  };

  var ui = {};

  // =====================================================================================
  // Shell
  // =====================================================================================
  function buildShell() {
    var host = document.getElementById('controlAddIn') || document.body;
    host.classList.add('ak-host');
    host.innerHTML =
      '<div class="ak-app" id="ak-app">' +
      '  <header class="ak-bar">' +
      '    <div class="ak-bar__brand">' +
      '      <label class="ak-autosave" title="Save automatically about 30 seconds after a change">' +
      '        <input type="checkbox" id="ak-autosave"><span class="ak-switch" aria-hidden="true"></span><span>AutoSave</span></label>' +
      '      <span class="ak-bar__logo">' + H.icon('sheet') + '</span>' +
      '      <div class="ak-bar__names">' +
      '        <input class="ak-title" id="ak-title" maxlength="100" placeholder="Untitled workbook" aria-label="Workbook name" spellcheck="false">' +
      '        <span class="ak-bar__code" id="ak-code"></span>' +
      '      </div>' +
      '      <span class="ak-chip" id="ak-dirty" data-state="saved">Saved</span>' +
      '    </div>' +
      '    <div class="ak-bar__sources" id="ak-bar-sources" aria-live="polite"></div>' +
      '    <div class="ak-bar__actions" role="toolbar" aria-label="Workbook">' +
      '      <button class="ak-btn ak-btn--write" data-act="write" id="ak-write" hidden disabled title="Write the changed and new rows of editable sheets to Business Central">' + H.icon('upload') + '<span>Save to BC</span><span class="ak-count" id="ak-write-count" hidden></span></button>' +
      '      <button class="ak-btn ak-btn--subtle" data-act="save" title="Save (Ctrl+S)">' + H.icon('save') + '<span>Save</span></button>' +
      '      <button class="ak-btn ak-btn--subtle" data-act="refresh" title="Reload every linked sheet from Business Central">' + H.icon('refresh') + '<span>Refresh</span></button>' +
      '      <span class="ak-bar__sep" aria-hidden="true"></span>' +
      '      <button class="ak-btn ak-btn--subtle" data-act="import" title="Open an .xlsx file">' + H.icon('upload') + '<span>Import</span></button>' +
      '      <button class="ak-btn ak-btn--subtle" data-act="export" title="Download as .xlsx">' + H.icon('download') + '<span>Export</span></button>' +
      '      <span class="ak-bar__sep" aria-hidden="true"></span>' +
      '      <button class="ak-btn ak-btn--outline" data-act="view-combine" aria-controls="ak-panel" title="Join two sheets on matching keys">' + H.icon('merge') + '<span>Combine</span></button>' +
      '      <button class="ak-btn ak-btn--outline" data-act="view-recon" aria-controls="ak-panel" title="Reconcile books against a statement, GSTR-2B or party ledger">' + H.icon('scale') + '<span>Reconcile</span></button>' +
      '      <button class="ak-btn ak-btn--primary" data-act="view-data" aria-controls="ak-panel" title="Load a Business Central table">' + H.icon('database') + '<span>Get data</span></button>' +
      '    </div>' +
      '    <input type="file" id="ak-file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden>' +
      '  </header>' +
      '  <div class="ak-progress" id="ak-progress" hidden><div class="ak-progress__bar" id="ak-progress-bar"></div></div>' +
      '  <div class="ak-body">' +
      '    <main class="ak-sheet" id="ak-sheet"></main>' +
      '    <aside class="ak-panel" id="ak-panel" data-view="data">' + dataViewHtml() + combineViewHtml() + reconViewHtml() + '</aside>' +
      '  </div>' +
      '  <div class="ak-toasts" id="ak-toasts" aria-live="polite"></div>' +
      '  <div class="ak-boot" id="ak-boot"><div class="ak-spinner"></div><div>Starting spreadsheet&hellip;</div></div>' +
      '</div>';

    var ids = ['app', 'sheet', 'panel', 'dirty', 'title', 'code', 'file', 'bar-sources', 'progress', 'progress-bar', 'boot',
      'table-search', 'table-results', 'table-picked', 'step-fields', 'field-search', 'field-list', 'field-count',
      'max-rows', 'sheet-name', 'load', 'load-hint', 'source-list', 'edit', 'edit-wrap', 'write', 'write-count',
      'j-left', 'j-right', 'j-keys', 'j-cols', 'j-name', 'j-run', 'j-hint', 'j-status', 'j-case', 'autosave',
      'r-left', 'r-right', 'r-llabel', 'r-rlabel', 'r-keys', 'r-norm', 'r-group', 'r-lamt', 'r-lminus', 'r-ramt', 'r-rminus',
      'r-flip', 'r-tol', 'r-ldate', 'r-rdate', 'r-ad', 'r-days', 'r-lextra', 'r-rextra', 'r-name', 'r-summary', 'r-run', 'r-hint', 'r-tip'];
    ids.forEach(function (id) {
      ui[id.replace(/-([a-z])/g, function (m, c) { return c.toUpperCase(); })] = document.getElementById('ak-' + id);
    });
    ui.loadBtn = ui.load;
    wireShell();
  }

  function panelHead(title, sub) {
    return '<div class="ak-panel__head">' +
      '<div><div class="ak-panel__title">' + title + '</div><div class="ak-panel__sub">' + sub + '</div></div>' +
      '<button class="ak-icon-btn" data-act="panel-close" aria-label="Close panel">' + H.icon('x') + '</button></div>';
  }

  function dataViewHtml() {
    return '<section class="ak-view" data-view="data" aria-label="Get data from Business Central">' +
      panelHead('Get data', 'Any table you can read in Business Central') +
      '<div class="ak-panel__scroll">' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">1</span>Table</div>' +
      '    <div class="ak-input-icon">' + H.icon('search') +
      '      <input id="ak-table-search" class="ak-input" type="search" placeholder="Search by name or number, e.g. Customer, 17" autocomplete="off" spellcheck="false">' +
      '    </div>' +
      '    <ul class="ak-list" id="ak-table-results" role="listbox" aria-label="Tables"></ul>' +
      '    <div class="ak-picked" id="ak-table-picked" hidden></div>' +
      '  </section>' +
      '  <section class="ak-step is-disabled" id="ak-step-fields">' +
      '    <div class="ak-step__label"><span class="ak-step__num">2</span>Fields and filters<span class="ak-step__meta" id="ak-field-count"></span></div>' +
      '    <div class="ak-field-tools">' +
      '      <div class="ak-input-icon ak-input-icon--sm">' + H.icon('search') +
      '        <input id="ak-field-search" class="ak-input ak-input--sm" type="search" placeholder="Find field" autocomplete="off">' +
      '      </div>' +
      '      <button class="ak-link" data-act="fields-key">Key</button>' +
      '      <button class="ak-link" data-act="fields-all">All</button>' +
      '      <button class="ak-link" data-act="fields-none">None</button>' +
      '    </div>' +
      '    <div class="ak-field-list" id="ak-field-list"></div>' +
      '    <p class="ak-help">Filters use Business Central syntax: <code>&gt;0</code>, <code>10000..20000</code>, <code>&lt;&gt;&#39;&#39;</code>, <code>A*|B*</code>, <code>01-04-26..31-03-27</code></p>' +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">3</span>Output</div>' +
      '    <div class="ak-grid2">' +
      '      <label class="ak-field"><span>Sheet name</span><input id="ak-sheet-name" class="ak-input" maxlength="31" placeholder="Table name"></label>' +
      '      <label class="ak-field"><span>Max rows</span><input id="ak-max-rows" class="ak-input" type="number" min="1" step="1000"></label>' +
      '    </div>' +
      '    <div class="ak-seg" role="radiogroup" aria-label="Destination">' +
      '      <label><input type="radio" name="ak-dest" value="new" checked><span>New sheet</span></label>' +
      '      <label><input type="radio" name="ak-dest" value="active"><span>Replace active sheet</span></label>' +
      '    </div>' +
      '    <label class="ak-check ak-edit-opt" id="ak-edit-wrap" hidden><input type="checkbox" id="ak-edit">' +
      '<span><b>Allow editing</b><br>Change values in the sheet and use <b>Save to BC</b>. Key and calculated fields stay locked.</span></label>' +
      '  </section>' +
      '  <section class="ak-step ak-step--sources">' +
      '    <div class="ak-step__label">' + H.icon('link') + 'Linked sheets</div>' +
      '    <ul class="ak-sources" id="ak-source-list"></ul>' +
      '  </section>' +
      '</div>' +
      '<div class="ak-panel__foot">' +
      '  <span class="ak-panel__hint" id="ak-load-hint">Pick a table to start</span>' +
      '  <button class="ak-btn ak-btn--primary ak-btn--lg" id="ak-load" disabled>' + H.icon('database') + '<span>Load data</span></button>' +
      '</div></section>';
  }

  function seg(name, items, checked) {
    return '<div class="ak-seg ak-seg--wrap" role="radiogroup">' + items.map(function (it) {
      return '<label title="' + H.escapeHtml(it[2] || '') + '"><input type="radio" name="' + name + '" value="' + it[0] + '"' +
        (it[0] === checked ? ' checked' : '') + '><span>' + it[1] + '</span></label>';
    }).join('') + '</div>';
  }

  function combineViewHtml() {
    return '<section class="ak-view" data-view="combine" aria-label="Combine sheets">' +
      panelHead('Combine sheets', 'Join two sheets on matching keys, like XLOOKUP for whole tables') +
      '<div class="ak-panel__scroll">' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">1</span>Sheets</div>' +
      '    <div class="ak-pair">' +
      '      <label class="ak-field"><span>Main sheet</span><select id="ak-j-left" class="ak-input ak-select"></select></label>' +
      '      <span class="ak-pair__icon">' + H.icon('merge') + '</span>' +
      '      <label class="ak-field"><span>Lookup sheet</span><select id="ak-j-right" class="ak-input ak-select"></select></label>' +
      '    </div>' +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">2</span>Match on</div>' +
      '    <div id="ak-j-keys" class="ak-keys"></div>' +
      '    <div class="ak-row-between">' +
      '      <button class="ak-link" data-act="j-add-key">+ Add key column</button>' +
      '      <label class="ak-check"><input type="checkbox" id="ak-j-case" checked><span>Ignore case and spaces</span></label>' +
      '    </div>' +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">3</span>Rows to keep</div>' +
      seg('ak-j-type', [
        ['left', 'All main rows', 'Every row of the main sheet, with lookup values where a match exists'],
        ['inner', 'Matched only', 'Only rows found in both sheets'],
        ['full', 'Both sheets', 'All rows from both sheets, matched where possible'],
        ['anti', 'Differences', 'Only rows that exist on one side - ideal for reconciliation']
      ], 'left') +
      '    <div class="ak-step__sub">When several lookup rows match</div>' +
      seg('ak-j-mode', [
        ['sum', 'Sum numbers', 'One row per key; numeric lookup columns are added up'],
        ['first', 'First match', 'Take the first matching lookup row'],
        ['all', 'Every match', 'Repeat the main row for every matching lookup row']
      ], 'sum') +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">4</span>Columns from lookup sheet' +
      '      <span class="ak-step__meta"><button class="ak-link" data-act="j-cols-all">All</button><button class="ak-link" data-act="j-cols-none">None</button></span></div>' +
      '    <div id="ak-j-cols" class="ak-chips"></div>' +
      '    <label class="ak-check ak-mt"><input type="checkbox" id="ak-j-status" checked><span>Add a <b>Match</b> column (Matched / Main only / Lookup only)</span></label>' +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">5</span>Output</div>' +
      '    <label class="ak-field"><span>New sheet name</span><input id="ak-j-name" class="ak-input" maxlength="31"></label>' +
      '  </section>' +
      '</div>' +
      '<div class="ak-panel__foot">' +
      '  <span class="ak-panel__hint" id="ak-j-hint">Load or open two sheets first</span>' +
      '  <button class="ak-btn ak-btn--primary ak-btn--lg" id="ak-j-run" disabled>' + H.icon('merge') + '<span>Combine</span></button>' +
      '</div></section>';
  }

  function reconViewHtml() {
    var tile = function (id, icon, title, sub) {
      return '<label class="ak-tile"><input type="radio" name="ak-r-preset" value="' + id + '"' + (id === 'gst' ? ' checked' : '') + '>' +
        '<span class="ak-tile__box"><span class="ak-tile__icon">' + H.icon(icon) + '</span>' +
        '<span class="ak-tile__title">' + title + '</span><span class="ak-tile__sub">' + sub + '</span></span></label>';
    };
    var sideRow = function (side, label) {
      return '<div class="ak-side">' +
        '<input id="ak-r-' + side[0] + 'label" class="ak-input ak-side__label" maxlength="20" value="' + label + '" aria-label="Name for this side">' +
        '<select id="ak-r-' + side + '" class="ak-input ak-select" aria-label="Sheet"></select></div>';
    };
    var pair = function (side, l, r) {
      return '<div class="ak-pair2' + (r ? '' : ' ak-pair2--one') + '"><span class="ak-pair2__title" data-r-label="' + side + '"></span>' + l + r + '</div>';
    };
    var sel = function (id, label) { return '<select id="ak-r-' + id + '" class="ak-input ak-input--sm ak-select" aria-label="' + label + '"></select>'; };
    return '<section class="ak-view" data-view="recon" aria-label="Reconcile">' +
      panelHead('Reconcile', 'Match your books against a bank statement, GSTR-2B or a party ledger') +
      '<div class="ak-panel__scroll">' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">1</span>What are you reconciling?</div>' +
      '    <div class="ak-tiles">' +
      tile('gst', 'receipt', 'GST vs GSTR-2B', 'GSTIN + invoice, summed per invoice') +
      tile('bank', 'bank', 'Bank vs G/L', 'Amount and date, debit/credit') +
      tile('party', 'users', 'Party ledger', 'Document no., opposite sign') +
      tile('custom', 'sliders', 'Custom', 'Set every rule yourself') +
      '    </div>' +
      '    <p class="ak-help ak-tip" id="ak-r-tip"></p>' +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">2</span>Sheets</div>' +
      sideRow('left', 'Books') + sideRow('right', 'Statement') +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">3</span>Match keys<span class="ak-step__meta">optional for bank</span></div>' +
      '    <div id="ak-r-keys" class="ak-keys"></div>' +
      '    <button class="ak-link ak-mt-s" data-act="r-add-key">+ Add key column</button>' +
      '    <div class="ak-checks">' +
      '      <label class="ak-check"><input type="checkbox" id="ak-r-norm" checked><span>Normalise document numbers (INV/001 = inv-1)</span></label>' +
      '      <label class="ak-check"><input type="checkbox" id="ak-r-group"><span>Sum lines per key first (one row per invoice)</span></label>' +
      '    </div>' +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">4</span>Amounts</div>' +
      '    <div class="ak-pair2 ak-pair2__legend"><span></span><span>Amount</span><span>Minus (e.g. Debit)</span></div>' +
      pair('left', sel('lamt', 'Amount'), sel('lminus', 'Minus column (optional)')) +
      pair('right', sel('ramt', 'Amount'), sel('rminus', 'Minus column (optional)')) +
      '    <div class="ak-row-between">' +
      '      <label class="ak-check"><input type="checkbox" id="ak-r-flip"><span>Second sheet has the opposite sign</span></label>' +
      '      <label class="ak-inline">Tolerance <input id="ak-r-tol" class="ak-input ak-input--sm ak-num" type="number" min="0" step="0.01" value="1"></label>' +
      '    </div>' +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">5</span>Dates</div>' +
      pair('left', sel('ldate', 'Date'), '') +
      pair('right', sel('rdate', 'Date'), '') +
      '    <label class="ak-check ak-mt-s"><input type="checkbox" id="ak-r-ad"><span>Also match by amount when dates are within ' +
      '<input id="ak-r-days" class="ak-input ak-input--sm ak-num ak-num--inline" type="number" min="0" max="60" value="3"> days</span></label>' +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">6</span>Output</div>' +
      '    <label class="ak-field"><span>Result sheet</span><input id="ak-r-name" class="ak-input" maxlength="22" value="Recon"></label>' +
      '    <div class="ak-step__sub">Extra columns from <span data-r-label="left"></span></div><div id="ak-r-lextra" class="ak-chips"></div>' +
      '    <div class="ak-step__sub">Extra columns from <span data-r-label="right"></span></div><div id="ak-r-rextra" class="ak-chips"></div>' +
      '    <label class="ak-check ak-mt"><input type="checkbox" id="ak-r-summary" checked><span>Create a summary sheet with live totals</span></label>' +
      '  </section>' +
      '</div>' +
      '<div class="ak-panel__foot">' +
      '  <span class="ak-panel__hint" id="ak-r-hint">Load or import two sheets first</span>' +
      '  <button class="ak-btn ak-btn--primary ak-btn--lg" id="ak-r-run" disabled>' + H.icon('scale') + '<span>Reconcile</span></button>' +
      '</div></section>';
  }

  function wireShell() {
    ui.app.addEventListener('click', function (e) {
      var btn = e.target.closest('[data-act]');
      if (!btn) return;
      var act = btn.getAttribute('data-act');
      var sheetId = btn.getAttribute('data-sheet');
      switch (act) {
        case 'save': requestSave(); break;
        case 'write': saveToBC(); break;
        case 'refresh': refreshAll(); break;
        case 'import': chooseImportFile(); break;
        case 'export': requestExport(); break;
        case 'view-data': togglePanel('data'); break;
        case 'view-combine': togglePanel('combine'); break;
        case 'view-recon': togglePanel('recon'); break;
        case 'r-add-key': state.recon.keys.push({ left: state.recon.left[0] || '', right: state.recon.right[0] || '' }); renderReconKeys(); break;
        case 'r-remove-key': state.recon.keys.splice(+btn.getAttribute('data-index'), 1); renderReconKeys(); break;
        case 'panel-close': togglePanel(false); break;
        case 'fields-all': setAllFields(true); break;
        case 'fields-none': setAllFields(false); break;
        case 'fields-key': selectKeyFields(); break;
        case 'change-table': resetTablePick(); break;
        case 'source-refresh': refreshOne(sheetId); break;
        case 'source-unlink': unlinkSource(sheetId); break;
        case 'source-open': activateSheet(sheetId); break;
        case 'j-add-key': state.combine.keys.push({ left: '', right: '' }); renderCombineKeys(); break;
        case 'j-remove-key': state.combine.keys.splice(+btn.getAttribute('data-index'), 1); renderCombineKeys(); break;
        case 'j-cols-all': setAllCombineCols(true); break;
        case 'j-cols-none': setAllCombineCols(false); break;
      }
    });

    // Workbook name -> BC Description
    ui.title.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); ui.title.blur(); }
      if (e.key === 'Escape') { ui.title.value = state.config.description || ''; ui.title.blur(); }
    });
    ui.title.addEventListener('blur', function () {
      var v = ui.title.value.trim().substring(0, 100);
      if (v === (state.config.description || '')) return;
      state.config.description = v;
      H.invoke('OnDescriptionChanged', [v]);
    });

    ui.file.addEventListener('change', function () {
      var f = ui.file.files && ui.file.files[0];
      ui.file.value = '';
      if (!f) return;
      if (f.size > 20 * 1024 * 1024) { H.toast('error', 'The file is larger than 20 MB.'); return; }
      setProgress(-1);
      f.arrayBuffer().then(function (buf) { return importXlsxBuffer(buf, state.importMode); }, function (e) { setProgress(null); reportError(e); });
    });

    // Get data view
    ui.tableSearch.addEventListener('input', H.debounce(function () { requestTables(ui.tableSearch.value); }, 250));
    ui.tableSearch.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { var first = ui.tableResults.querySelector('[data-table]'); if (first) first.click(); }
    });
    ui.tableResults.addEventListener('click', function (e) {
      var li = e.target.closest('[data-table]');
      if (li) pickTable(+li.getAttribute('data-table'));
    });
    ui.fieldSearch.addEventListener('input', function () {
      state.fieldSearch = ui.fieldSearch.value.toLowerCase();
      renderFields();
    });
    ui.fieldList.addEventListener('change', function (e) {
      var row = e.target.closest('[data-field]');
      if (!row || e.target.type !== 'checkbox') return;
      var fs = state.fieldState[+row.getAttribute('data-field')];
      fs.checked = e.target.checked;
      row.classList.toggle('is-checked', fs.checked);
      updateLoadState();
    });
    ui.fieldList.addEventListener('input', function (e) {
      if (!e.target.classList.contains('ak-filter')) return;
      var row = e.target.closest('[data-field]');
      var fs = state.fieldState[+row.getAttribute('data-field')];
      fs.filter = e.target.value;
      if (fs.filter && !fs.checked) { // filtering on a field implies you want it in the output
        fs.checked = true;
        row.classList.add('is-checked');
        row.querySelector('input[type=checkbox]').checked = true;
      }
      row.classList.toggle('has-filter', !!fs.filter);
      updateLoadState();
    });
    ui.loadBtn.addEventListener('click', loadFromPanel);
    ui.edit.addEventListener('change', function () { renderFields(); updateLoadState(); });

    // Combine view
    ui.jLeft.addEventListener('change', function () { onCombineSheetsChanged(); });
    ui.jRight.addEventListener('change', function () { onCombineSheetsChanged(); });
    ui.jKeys.addEventListener('change', function (e) {
      var sel = e.target.closest('select[data-side]');
      if (!sel) return;
      state.combine.keys[+sel.getAttribute('data-index')][sel.getAttribute('data-side')] = sel.value;
      renderCombineCols();
      updateCombineState();
    });
    ui.jCols.addEventListener('change', function (e) {
      if (e.target.type !== 'checkbox') return;
      state.combine.rightCols[e.target.value] = e.target.checked;
      e.target.parentNode.classList.toggle('is-checked', e.target.checked);
      updateCombineState();
    });
    ui.jRun.addEventListener('click', runCombineFromPanel);

    // Reconcile view
    ui.panel.addEventListener('change', function (e) {
      if (e.target.name === 'ak-r-preset') applyReconPreset(e.target.value, true);
      if (e.target.closest('.ak-chipbox')) e.target.parentNode.classList.toggle('is-checked', e.target.checked);
    });
    [ui.rLeft, ui.rRight].forEach(function (el) { el.addEventListener('change', function () { onReconSheetsChanged(); }); });
    [ui.rLlabel, ui.rRlabel].forEach(function (el) { el.addEventListener('input', renderReconLabels); });
    ui.rKeys.addEventListener('change', function (e) {
      var el = e.target.closest('select[data-side]');
      if (!el) return;
      state.recon.keys[+el.getAttribute('data-index')][el.getAttribute('data-side')] = el.value;
      updateReconState();
    });
    [ui.rLamt, ui.rRamt, ui.rAd].forEach(function (el) { el.addEventListener('change', updateReconState); });
    ui.rRun.addEventListener('click', runReconFromPanel);

    // AutoSave switch (remembered per user in this browser)
    var pref = H.pref('autosave');
    state.autosave = pref === undefined ? true : !!pref;
    ui.autosave.checked = state.autosave;
    ui.autosave.addEventListener('change', function () {
      state.autosave = ui.autosave.checked;
      H.pref('autosave', state.autosave);
      if (state.autosave && state.dirty) scheduleAutosave();
      H.toast('info', state.autosave ? 'AutoSave is on.' : 'AutoSave is off. Use Save or Ctrl+S.');
    });

    // Ctrl+S saves through AL, instead of the browser "save page" dialog.
    document.addEventListener('keydown', function (e) {
      if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
        e.preventDefault();
        requestSave();
      }
    }, true);
  }

  // =====================================================================================
  // Univer lifecycle
  // =====================================================================================
  function startUniver() {
    if (state.univer) return;
    var created = window.AKUniver.create('ak-sheet');
    state.univer = created.univer;
    state.api = created.univerAPI;
    // Ignore start-up mutations (render setup, default styles) so a new workbook opens as "Saved".
    state.suppressDirty++;
    setTimeout(function () { state.suppressDirty--; }, 1200);
    newEmptyWorkbook();

    state.disposables.push(state.api.addEvent(state.api.Event.CommandExecuted, function (evt) {
      if (evt && evt.type === MUTATION && state.suppressDirty === 0) setDirty(true);
      if (evt && evt.type === MUTATION) onSheetMutation(evt);
    }));

    // Univer listens to window resize; BC resizes the iframe but the panel changes our own width.
    var kick = H.debounce(function () { window.dispatchEvent(new Event('resize')); }, 60);
    if (window.ResizeObserver) {
      state.resizeObserver = new ResizeObserver(kick);
      state.resizeObserver.observe(ui.sheet);
    }
    ui.boot.classList.add('is-done');
    setTimeout(function () { if (ui.boot.parentNode) ui.boot.parentNode.removeChild(ui.boot); }, 300);
  }

  function workbook() {
    return state.api && state.api.getActiveWorkbook();
  }

  function newEmptyWorkbook() {
    withoutDirty(function () {
      state.api.createWorkbook({ id: H.uid('wb'), name: 'Workbook', locale: window.AKUniver.LocaleType.EN_US });
    });
  }

  function replaceWorkbook(snapshot) {
    withoutDirty(function () {
      var old = workbook();
      if (old) old.dispose();
      // Always a fresh unit id: the saved id may still be registered while the old unit is disposed.
      snapshot.id = H.uid('wb');
      snapshot.locale = window.AKUniver.LocaleType.EN_US;
      state.api.createWorkbook(snapshot);
    });
  }

  function withoutDirty(fn) {
    state.suppressDirty++;
    // Univer applies some mutations asynchronously after a unit is created, so keep the guard briefly.
    try { return fn(); } finally { setTimeout(function () { state.suppressDirty--; }, 400); }
  }

  function destroy() {
    if (state.destroyed) return;
    state.destroyed = true;
    try { if (state.resizeObserver) state.resizeObserver.disconnect(); } catch (e) { /* ignore */ }
    state.disposables.forEach(function (d) { try { d.dispose(); } catch (e) { /* ignore */ } });
    state.disposables = [];
    try { if (state.univer) state.univer.dispose(); } catch (e) { /* ignore */ }
    state.univer = null;
    state.api = null;
  }

  // =====================================================================================
  // Header state
  // =====================================================================================
  function setDirty(isDirty) {
    if (state.dirty === isDirty) return;
    state.dirty = isDirty;
    ui.dirty.setAttribute('data-state', isDirty ? 'dirty' : 'saved');
    ui.dirty.textContent = isDirty ? 'Unsaved changes' : 'Saved';
    ui.app.classList.toggle('is-dirty', isDirty);
    H.invoke('OnDirtyChanged', [isDirty]);
    if (isDirty) scheduleAutosave();
  }

  /** AutoSave: one quiet save about 30 s after the first unsaved change (retries while busy). */
  function scheduleAutosave() {
    if (!state.autosave || state.autosaveTimer) return;
    state.autosaveTimer = setTimeout(function tick() {
      state.autosaveTimer = null;
      if (!state.autosave || !state.dirty) return;
      if (state.busy || ui.app.classList.contains('is-saving')) { state.autosaveTimer = setTimeout(tick, 10000); return; }
      requestSave(true);
    }, state.config.autosaveMs || 30000);
  }

  function renderBarSources() {
    var nSrc = sourceList().length, nJoin = joinList().length + reconList().length;
    if (!nSrc && !nJoin) { ui.barSources.innerHTML = ''; return; }
    var total = sourceList().reduce(function (n, s) { return n + (s.rowCount || 0); }, 0);
    ui.barSources.innerHTML = H.icon('link') + '<span>' + nSrc + ' linked ' + (nSrc === 1 ? 'sheet' : 'sheets') +
      ' &middot; ' + H.formatInt(total) + ' rows' + (nJoin ? ' &middot; ' + nJoin + ' derived' : '') + '</span>';
  }

  function setProgress(fraction) {
    if (fraction == null) { ui.progress.hidden = true; ui.progressBar.style.width = '0'; return; }
    ui.progress.hidden = false;
    ui.progress.classList.toggle('is-indeterminate', fraction < 0);
    ui.progressBar.style.width = fraction < 0 ? '' : Math.max(3, Math.round(fraction * 100)) + '%';
  }

  function setBusy(b) {
    state.busy = b;
    ui.app.classList.toggle('is-busy', b);
    if (!b) setProgress(null);
    updateLoadState();
    updateCombineState();
    updateReconState();
    updateWriteButton();
  }

  // =====================================================================================
  // Panel
  // =====================================================================================
  /** togglePanel('data' | 'combine') opens that view (or closes it if already showing); false closes. */
  function togglePanel(view) {
    var isOpen = ui.app.classList.contains('is-panel-open');
    var current = ui.panel.getAttribute('data-view');
    var open = view !== false && !(isOpen && (view === current || view === undefined));
    var target = typeof view === 'string' ? view : current;
    ui.app.classList.toggle('is-panel-open', open);
    ui.panel.setAttribute('data-view', target);
    ['data', 'combine', 'recon'].forEach(function (v) {
      var b = ui.app.querySelector('[data-act=view-' + v + ']');
      b.setAttribute('aria-expanded', String(open && target === v));
      b.classList.toggle('is-active', open && target === v);
    });
    if (!open) return;
    if (target === 'data') {
      renderSources();
      if (!state.table) {
        if (!state.tables.length) requestTables('');
        setTimeout(function () { ui.tableSearch.focus(); }, 200);
      }
    } else if (target === 'combine') {
      prepareCombine();
    } else {
      prepareRecon();
    }
  }

  // =====================================================================================
  // Get data
  // =====================================================================================
  function requestTables(text) {
    ui.tableResults.innerHTML = '<li class="ak-list__empty"><span class="ak-spinner ak-spinner--sm"></span>Searching&hellip;</li>';
    H.invoke('OnTablesRequested', [text || '']);
  }

  function renderTables() {
    if (!state.tables.length) {
      ui.tableResults.innerHTML = '<li class="ak-list__empty">No tables match. Try the table number.</li>';
      return;
    }
    ui.tableResults.innerHTML = state.tables.map(function (t) {
      return '<li class="ak-list__item" role="option" tabindex="0" data-table="' + t.id + '">' +
        '<span class="ak-list__icon">' + H.icon('table') + '</span>' +
        '<span class="ak-list__text"><span class="ak-list__title">' + H.escapeHtml(t.caption) + '</span>' +
        '<span class="ak-list__meta">' + t.id + ' &middot; ' + H.escapeHtml(t.name) + '</span></span></li>';
    }).join('');
  }

  function pickTable(tableNo) {
    var t = state.tables.filter(function (x) { return x.id === tableNo; })[0];
    ui.tableResults.hidden = true;
    ui.tableSearch.parentNode.hidden = true;
    ui.tablePicked.hidden = false;
    ui.tablePicked.innerHTML = '<span class="ak-list__icon">' + H.icon('table') + '</span>' +
      '<span class="ak-list__text"><span class="ak-list__title">' + H.escapeHtml(t ? t.caption : 'Table ' + tableNo) + '</span>' +
      '<span class="ak-list__meta">Table ' + tableNo + '</span></span>' +
      '<button class="ak-link" data-act="change-table">Change</button>';
    ui.fieldList.innerHTML = '<div class="ak-list__empty"><span class="ak-spinner ak-spinner--sm"></span>Loading fields&hellip;</div>';
    ui.stepFields.classList.remove('is-disabled');
    state.table = null;
    updateLoadState();
    H.invoke('OnFieldsRequested', [tableNo]);
  }

  function resetTablePick() {
    state.table = null;
    state.fieldState = {};
    ui.tablePicked.hidden = true;
    ui.tableResults.hidden = false;
    ui.tableSearch.parentNode.hidden = false;
    ui.stepFields.classList.add('is-disabled');
    ui.fieldList.innerHTML = '';
    ui.fieldCount.textContent = '';
    ui.sheetName.value = '';
    ui.editWrap.hidden = true;
    ui.edit.checked = false;
    updateLoadState();
    ui.tableSearch.focus();
  }

  function applyFieldsResponse(data) {
    state.table = data;
    state.fieldState = {};
    var defaults = 0;
    data.fields.forEach(function (f) {
      // Default selection: primary key + the first useful normal fields, like a list page would show.
      var pick = f.pk || (!f.flow && defaults < 8);
      if (pick && !f.pk) defaults++;
      state.fieldState[f.no] = { checked: pick, filter: '' };
    });
    ui.sheetName.value = H.safeSheetName(data.tableCaption);
    ui.editWrap.hidden = !data.canEdit;
    ui.edit.checked = false;
    ui.fieldSearch.value = '';
    state.fieldSearch = '';
    renderFields();
    updateLoadState();
    if (!data.canRead) H.toast('warning', 'You can see this table but do not have permission to read its data.');
  }

  function renderFields() {
    if (!state.table) return;
    var q = state.fieldSearch;
    var html = state.table.fields.filter(function (f) {
      return !q || f.caption.toLowerCase().indexOf(q) >= 0 || f.name.toLowerCase().indexOf(q) >= 0 || String(f.no) === q;
    }).map(function (f) {
      var fs = state.fieldState[f.no], editing = isEditing(), forced = editing && f.pk, on = fs.checked || forced;
      return '<div class="ak-frow' + (on ? ' is-checked' : '') + (fs.filter ? ' has-filter' : '') + '" data-field="' + f.no + '">' +
        '<label class="ak-frow__main"><input type="checkbox"' + (on ? ' checked' : '') + (forced ? ' disabled title="Key fields are needed to save changes"' : '') + '>' +
        '<span class="ak-frow__text"><span class="ak-frow__cap">' + H.escapeHtml(f.caption) + '</span>' +
        '<span class="ak-frow__meta">' + f.no + ' &middot; ' + H.escapeHtml(f.type) +
        (f.pk ? ' <span class="ak-badge ak-badge--key">' + H.icon('key') + 'Key</span>' : '') +
        (f.flow ? ' <span class="ak-badge ak-badge--flow">' + H.icon('sigma') + 'Calculated</span>' : '') +
        (editing && f.editable ? ' <span class="ak-badge ak-badge--edit">' + H.icon('pencil') + 'Editable</span>' : '') +
        '</span></span></label>' +
        '<input class="ak-input ak-input--sm ak-filter" placeholder="Filter" value="' + H.escapeHtml(fs.filter) + '" aria-label="Filter for ' + H.escapeHtml(f.caption) + '">' +
        '</div>';
    }).join('');
    ui.fieldList.innerHTML = html || '<div class="ak-list__empty">No fields match.</div>';
  }

  function setAllFields(on) {
    if (!state.table) return;
    state.table.fields.forEach(function (f) { state.fieldState[f.no].checked = on || !!state.fieldState[f.no].filter; });
    renderFields(); updateLoadState();
  }

  function selectKeyFields() {
    if (!state.table) return;
    state.table.fields.forEach(function (f) { state.fieldState[f.no].checked = f.pk || !!state.fieldState[f.no].filter; });
    renderFields(); updateLoadState();
  }

  function isEditing() {
    return !!(state.table && state.table.canEdit && ui.edit && ui.edit.checked);
  }

  function selectedFields() {
    if (!state.table) return [];
    var editing = isEditing();
    return state.table.fields.filter(function (f) { return state.fieldState[f.no].checked || (editing && f.pk); });
  }

  function updateLoadState() {
    if (!ui.loadBtn) return;
    var fields = selectedFields();
    var filters = fields.filter(function (f) { return state.fieldState[f.no].filter; }).length;
    ui.fieldCount.textContent = state.table ? fields.length + ' of ' + state.table.fields.length + ' selected' : '';
    ui.loadBtn.disabled = !(state.table && state.table.canRead && fields.length > 0 && !state.busy);
    ui.loadHint.textContent = !state.table ? 'Pick a table to start'
      : state.busy ? 'Working…'
      : !fields.length ? 'Select at least one field'
      : fields.length + ' fields' + (filters ? ' · ' + filters + (filters === 1 ? ' filter' : ' filters') : ' · no filters');
  }

  function loadFromPanel() {
    var fields = selectedFields();
    if (!state.table || !fields.length) return;
    var maxRows = parseInt(ui.maxRows.value, 10);
    if (!(maxRows > 0)) maxRows = state.config.defaultMaxRows;
    maxRows = Math.min(maxRows, state.config.maxRowsLimit);
    ui.maxRows.value = maxRows;

    var dest = (ui.panel.querySelector('input[name=ak-dest]:checked') || {}).value || 'new';
    var wb = workbook();
    var sheet;
    if (dest === 'active') {
      sheet = wb.getActiveSheet();
      var oldDims = state.sources[sheet.getSheetId()] || state.joins[sheet.getSheetId()] || state.recons[sheet.getSheetId()];
      if (oldDims && oldDims.lastRows) withoutDirty(function () { sheet.getRange(0, 0, oldDims.lastRows, oldDims.lastCols).clearContent(); });
      delete state.joins[sheet.getSheetId()];
      delete state.recons[sheet.getSheetId()];
      if (ui.sheetName.value && sheet.getSheetName() !== ui.sheetName.value) {
        try { sheet.setName(uniqueSheetName(ui.sheetName.value)); } catch (e) { /* keep old name */ }
      }
    } else {
      sheet = wb.insertSheet(uniqueSheetName(ui.sheetName.value || state.table.tableCaption));
    }
    wb.setActiveSheet(sheet);

    var source = {
      sheetId: sheet.getSheetId(),
      tableNo: state.table.tableNo,
      tableCaption: state.table.tableCaption,
      fields: fields.map(function (f) { return { no: f.no, caption: f.caption, type: f.type, pk: !!f.pk }; }),
      filters: fields.filter(function (f) { return state.fieldState[f.no].filter; })
        .map(function (f) { return { fieldNo: f.no, filter: state.fieldState[f.no].filter }; }),
      maxRows: maxRows,
      editable: isEditing(),
      lastRows: 0,
      lastCols: 0
    };
    var createdSheetId = dest === 'new' ? source.sheetId : null;
    runSource(source).then(function (res) {
      H.toast(res.truncated ? 'warning' : 'success', res.truncated
        ? 'Loaded the first ' + H.formatInt(res.loaded) + ' of ' + H.formatInt(res.total) + ' rows. Raise Max rows or add filters.'
        : 'Loaded ' + H.formatInt(res.loaded) + ' rows from ' + source.tableCaption + '.');
      if (window.innerWidth < 900) togglePanel(false);
    }, function () {
      // Error already shown; remove the empty sheet we created for this load.
      if (!createdSheetId) return;
      var w = workbook(), s = w && w.getSheetBySheetId(createdSheetId);
      if (s && w.getSheets().length > 1) withoutDirty(function () { w.deleteSheet(s); });
    });
  }

  /** Requests data for a source and writes it to its sheet. Resolves with {loaded,total,truncated}. */
  function runSource(source) {
    setBusy(true);
    var requestId = 'r' + (++state.requestSeq);
    return new Promise(function (resolve, reject) {
      state.pending[requestId] = { source: source, columns: null, rows: [], resolve: resolve, reject: reject };
      setProgress(-1);
      H.invoke('OnDataRequested', [JSON.stringify({
        requestId: requestId,
        tableNo: source.tableNo,
        fields: source.fields.map(function (f) { return f.no; }),
        filters: source.filters,
        maxRows: source.maxRows,
        editable: !!source.editable
      })]).catch(function (e) { failRequest(requestId, e.message || String(e)); });
    }).then(function (r) { setBusy(false); return r; }, function (e) { setBusy(false); throw e; });
  }

  function onChunk(chunk) {
    var p = state.pending[chunk.requestId];
    if (!p) return;
    if (chunk.chunk === 0) {
      p.columns = chunk.columns;
      p.filterText = chunk.filterText;
      p.editable = !!chunk.editable;
      p.canInsert = !!chunk.canInsert;
      p.versions = [];
    }
    var rows = chunk.rows || [];
    for (var i = 0; i < rows.length; i++) p.rows.push(rows[i]);
    if (chunk.versions) for (var v = 0; v < chunk.versions.length; v++) p.versions.push(chunk.versions[v]);
    var target = Math.min(chunk.total, p.source.maxRows);
    setProgress(target > 0 ? chunk.loaded / target : -1);
    if (!chunk.isLast) return;

    delete state.pending[chunk.requestId];
    try {
      var columns = p.columns || [];
      var types = columns.map(function (c) { return c.type; });
      var values = p.rows.map(function (src) {
        var out = new Array(types.length);
        for (var c = 0; c < types.length; c++) out[c] = convertValue(types[c], src[c]);
        return out;
      });
      var header = columns.map(function (c) { return c.caption; });
      if (p.editable) {
        // Extra last column: BC Status (Changed / New / Saved / BC's error message).
        header.push('BC Status');
        types = types.concat(['Text']);
        values.forEach(function (row) { row.push(''); });
      }
      p.source.editable = p.editable;
      applying(function () { writeTable(p.source, header, types, values); });
      if (p.editable) setupEditable(p.source, columns, values, p.versions, p.canInsert);
      p.source.filterText = p.filterText || '';
      p.source.loadedAt = new Date().toISOString();
      p.source.total = chunk.total;
      p.source.truncated = chunk.truncated;
      state.sources[p.source.sheetId] = p.source;
      renderSources();
      renderBarSources();
      p.resolve({ loaded: chunk.loaded, total: chunk.total, truncated: chunk.truncated });
    } catch (e) {
      reportError(e);
      p.reject(e);
    }
  }

  function failRequest(requestId, message) {
    var p = state.pending[requestId];
    H.toast('error', message || 'Loading data failed.');
    if (!p) { setBusy(false); return; }
    delete state.pending[requestId];
    p.reject(new Error(message));
  }

  function convertValue(type, v) {
    if (v === '' || v == null) return '';
    if (type === 'Date') return H.isoDateToSerial(v);
    if (type === 'DateTime') return H.isoDateTimeToSerial(v);
    return v;
  }

  // =====================================================================================
  // Writing a table into a sheet (shared by BC loads and Combine)
  // =====================================================================================
  var NUMBER_FORMATS = { Decimal: '#,##0.00', Date: 'dd-mmm-yyyy', DateTime: 'dd-mmm-yyyy hh:mm' };

  /**
   * Writes header + rows to target.sheetId. Clears only the block the previous write produced,
   * so formulas next to or below the data survive a refresh. Updates target.lastRows/lastCols/rowCount.
   */
  function writeTable(target, header, types, rows) {
    var wb = workbook();
    var sheet = wb.getSheetBySheetId(target.sheetId);
    if (!sheet) {
      sheet = wb.insertSheet(uniqueSheetName(target.tableCaption || target.name || 'Data'));
      target.sheetId = sheet.getSheetId();
    }
    var nCols = header.length, nRows = rows.length + 1;

    if (target.lastRows && target.lastCols) sheet.getRange(0, 0, target.lastRows, target.lastCols).clearContent();
    if (sheet.getMaxRows() < nRows + 50) sheet.insertRowsAfter(sheet.getMaxRows() - 1, nRows + 50 - sheet.getMaxRows());
    if (sheet.getMaxColumns() < nCols + 5) sheet.insertColumnsAfter(sheet.getMaxColumns() - 1, nCols + 5 - sheet.getMaxColumns());

    var values = new Array(nRows);
    values[0] = header.slice();
    for (var r = 0; r < rows.length; r++) values[r + 1] = rows[r];
    sheet.getRange(0, 0, nRows, nCols).setValues(values);

    var head = sheet.getRange(0, 0, 1, nCols);
    head.setFontWeight('bold');
    head.setBackgroundColor('#EBF3FC');
    head.setFontColor('#0F3D6E');
    sheet.setFrozenRows(1);

    if (rows.length) {
      types.forEach(function (t, c) {
        var fmt = NUMBER_FORMATS[t];
        if (fmt) { try { sheet.getRange(1, c, rows.length, 1).setNumberFormat(fmt); } catch (e) { /* numfmt optional */ } }
        if (NUMERIC_TYPES[t]) { try { sheet.getRange(0, c, 1, 1).setHorizontalAlignment('right'); } catch (e) { /* optional */ } }
      });
    }
    autoWidth(sheet, header, types, rows);

    try {
      var existing = sheet.getFilter && sheet.getFilter();
      if (existing) existing.remove();
      sheet.getRange(0, 0, Math.max(nRows, 2), nCols).createFilter();
    } catch (e) { /* filter preset optional */ }
    try { sheet.getRange(0, 0, 1, 1).activate(); } catch (e) { /* keep selection */ }

    target.lastRows = nRows;
    target.lastCols = nCols;
    target.rowCount = rows.length;
    return sheet;
  }

  function autoWidth(sheet, header, types, rows) {
    var sample = Math.min(rows.length, 200);
    header.forEach(function (cap, c) {
      var t = types[c], len = String(cap).length + 2;
      for (var r = 0; r < sample; r++) {
        var v = rows[r][c];
        var l = t === 'Date' ? 11 : t === 'DateTime' ? 17 : NUMERIC_TYPES[t] ? String(Math.round(+v || 0)).length + 6 : String(v == null ? '' : v).length;
        if (l > len) len = l;
      }
      sheet.setColumnWidths(c, 1, Math.max(64, Math.min(320, len * 7.5 + 18)));
    });
  }

  function uniqueSheetName(base) {
    var wb = workbook(), name = H.safeSheetName(base), i = 2;
    while (wb.getSheetByName(name)) {
      var suffix = ' (' + i++ + ')';
      name = H.safeSheetName(base).substring(0, 31 - suffix.length) + suffix;
    }
    return name;
  }

  // =====================================================================================
  // Linked sheets (BC sources + combined sheets)
  // =====================================================================================
  function pruneMissing(map) {
    var wb = workbook();
    Object.keys(map).forEach(function (id) { if (!wb.getSheetBySheetId(id)) delete map[id]; });
    return Object.keys(map).map(function (id) { return map[id]; });
  }
  function sourceList() { return workbook() ? pruneMissing(state.sources) : []; }
  function joinList() { return workbook() ? pruneMissing(state.joins) : []; }
  function reconList() { return workbook() ? pruneMissing(state.recons) : []; }
  function derivedList() {
    return joinList().map(function (d) { return { kind: 'join', def: d }; })
      .concat(reconList().map(function (d) { return { kind: 'recon', def: d }; }))
      .sort(function (a, b) { return (a.def.createdAt || '') < (b.def.createdAt || '') ? -1 : 1; });
  }
  function runDerived(d) {
    return d.kind === 'join' ? runJoin(d.def).rows : runRecon(d.def).rows;
  }

  function refreshAll() {
    var sources = sourceList(), derived = derivedList();
    if (!sources.length && !derived.length) {
      H.toast('info', 'No linked sheets yet. Use Get data to load a table.');
      return Promise.resolve();
    }
    if (state.busy) return Promise.resolve();
    return confirmDiscard(sources).then(function (ok) { if (ok) return refreshAllNow(sources, derived); });
  }

  function refreshAllNow(sources, derived) {
    var done = 0, rows = 0;
    var chain = Promise.resolve();
    sources.forEach(function (s) {
      chain = chain.then(function () { return runSource(s); }).then(function (r) { done++; rows += r.loaded; });
    });
    // Combined and reconciled sheets are rebuilt after their inputs, in the order they were created.
    derived.forEach(function (d) {
      chain = chain.then(function () { rows += runDerived(d); done++; });
    });
    return chain.then(function () {
      H.toast('success', 'Refreshed ' + done + (done === 1 ? ' sheet' : ' sheets') + ' (' + H.formatInt(rows) + ' rows).');
    }, function (e) { if (e && !e.shown) reportError(e); });
  }

  function refreshOne(sheetId) {
    if (state.busy) return;
    var s = state.sources[sheetId], j = state.joins[sheetId], rc = state.recons[sheetId];
    if (s) confirmDiscard([s]).then(function (ok) {
      if (ok) runSource(s).then(function (r) { H.toast('success', s.tableCaption + ': ' + H.formatInt(r.loaded) + ' rows.'); }, function () { });
    });
    else if (j) {
      try { var r = runJoin(j); H.toast('success', 'Recombined: ' + H.formatInt(r.rows) + ' rows.'); } catch (e) { reportError(e); }
    } else if (rc) {
      try { var x = runRecon(rc); H.toast('success', 'Reconciled again: ' + H.formatInt(x.matched) + ' matched, ' + H.formatInt(x.exceptions) + ' exceptions.'); } catch (e) { reportError(e); }
    }
  }

  function unlinkSource(sheetId) {
    delete state.sources[sheetId];
    updateWriteButton();
    delete state.joins[sheetId];
    delete state.recons[sheetId];
    renderSources(); renderBarSources(); setDirty(true);
    H.toast('info', 'Sheet unlinked. Its data stays but will not refresh.');
  }

  function activateSheet(sheetId) {
    var wb = workbook(), sheet = wb && wb.getSheetBySheetId(sheetId);
    if (sheet) wb.setActiveSheet(sheet);
  }

  function sheetName(sheetId) {
    var s = workbook().getSheetBySheetId(sheetId);
    return s ? s.getSheetName() : '?';
  }

  function renderSources() {
    var sources = sourceList(), joins = joinList(), recons = reconList();
    if (!sources.length && !joins.length && !recons.length) {
      ui.sourceList.innerHTML = '<li class="ak-list__empty">Sheets you load from Business Central appear here and refresh with one click.</li>';
      return;
    }
    var item = function (id, name, meta, badge) {
      return '<li class="ak-source">' +
        '<button class="ak-source__main" data-act="source-open" data-sheet="' + id + '" title="Go to sheet">' +
        '<span class="ak-source__name">' + badge + H.escapeHtml(name) + '</span>' +
        '<span class="ak-source__meta">' + meta + '</span></button>' +
        '<button class="ak-icon-btn" data-act="source-refresh" data-sheet="' + id + '" aria-label="Refresh" title="Refresh">' + H.icon('refresh') + '</button>' +
        '<button class="ak-icon-btn" data-act="source-unlink" data-sheet="' + id + '" aria-label="Unlink" title="Unlink (keep data)">' + H.icon('unlink') + '</button>' +
        '</li>';
    };
    var when = function (iso) {
      return iso ? ' &middot; ' + new Date(iso).toLocaleString([], { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' }) : '';
    };
    ui.sourceList.innerHTML = sources.map(function (s) {
      return item(s.sheetId, sheetName(s.sheetId),
        H.escapeHtml(s.tableCaption) + ' &middot; ' + H.formatInt(s.rowCount) + ' rows' +
        (s.filters.length ? ' &middot; ' + s.filters.length + ' filter' + (s.filters.length > 1 ? 's' : '') : '') + when(s.loadedAt), '');
    }).concat(joins.map(function (j) {
      return item(j.sheetId, sheetName(j.sheetId),
        H.escapeHtml(sheetName(j.leftSheetId)) + ' + ' + H.escapeHtml(sheetName(j.rightSheetId)) + ' &middot; ' + H.formatInt(j.rowCount) + ' rows' + when(j.builtAt),
        '<span class="ak-badge ak-badge--join">' + H.icon('merge') + '</span>');
    })).concat(recons.map(function (r) {
      return item(r.sheetId, sheetName(r.sheetId),
        H.escapeHtml(sheetName(r.leftSheetId)) + ' vs ' + H.escapeHtml(sheetName(r.rightSheetId)) + ' &middot; ' +
        H.formatInt(r.matched || 0) + ' matched, ' + H.formatInt(r.exceptions || 0) + ' exceptions' + when(r.builtAt),
        '<span class="ak-badge ak-badge--recon">' + H.icon('scale') + '</span>');
    })).join('');
  }

  // =====================================================================================
  // Stage 1: editable sheets (write-back to Business Central)
  //
  // A sheet loaded with "Allow editing" keeps a copy of every loaded row by primary key (orig) and
  // the record's SystemModifiedAt (ver). On each edit the sheet is compared with that copy:
  //   - an existing key with different values in editable columns -> 'modify'
  //   - a row whose key is not in the copy (typed below the data)  -> 'insert'
  // Key, calculated and system columns are locked: an edit there is put back at once.
  // Save to BC sends the changes to AL, which validates each field like the standard page and
  // answers per row; the BC Status column shows "Saved" or BC's own error message.
  // =====================================================================================
  var EDIT_COLORS = { locked: '#F3F2F1', lockedText: '#605E5C', changed: '#FFF4CE', added: '#E7F6EC', head: '#EBF3FC', statusHead: '#323130' };
  var STATUS_CHANGED = 'Changed', STATUS_NEW = 'New';

  function editableList() { return sourceList().filter(function (s) { return s.editable; }); }

  function sourceBySheet(sheetId) {
    var s = state.sources[sheetId];
    return s && s.editable ? s : null;
  }

  function keyOf(source, row) {
    return source.keyIdx.map(function (i) { return toBC(source.columns[i].type, row[i]); }).join('\u0001');
  }

  /** Sheet value -> the text AL expects for that field type. */
  function toBC(type, v) {
    if (v == null) return '';
    if (type === 'Date') return H.serialToIsoDate(v);
    if (type === 'DateTime') return H.serialToIsoDateTime(v);
    if (type === 'Boolean') {
      if (v === '') return '';
      var t = String(v).trim().toLowerCase();
      return (v === true || t === 'true' || t === 'yes' || t === '1') ? 'true' : 'false';
    }
    if (NUMERIC_TYPES[type]) {
      if (v === '') return '';
      if (typeof v === 'number') return String(Math.round(v * 1e9) / 1e9);
      var n = String(v).replace(/,/g, '').trim();
      return n;
    }
    return String(v).trim();
  }

  function sameValue(type, a, b) {
    var x = toBC(type, a), y = toBC(type, b);
    if (NUMERIC_TYPES[type] && x !== '' && y !== '' && isFinite(+x) && isFinite(+y)) return Math.abs(+x - +y) < 1e-9;
    if (type === 'Boolean') return (x || 'false') === (y || 'false');
    return x === y;
  }

  /** Called after an editable load: remembers the loaded rows and paints locked columns. */
  function setupEditable(source, columns, rows, versions, canInsert) {
    source.columns = columns.map(function (c) { return { no: c.no, caption: c.caption, type: c.type, pk: !!c.pk, editable: !!c.editable }; });
    source.keyIdx = [];
    source.columns.forEach(function (c, i) { if (c.pk) source.keyIdx.push(i); });
    source.dataCols = source.columns.length;
    source.statusCol = source.columns.length;
    source.canInsert = !!canInsert;
    source.orig = {};
    source.ver = {};
    source.posKeys = [];
    rows.forEach(function (r, i) {
      var k = keyOf(source, r);
      source.orig[k] = r.slice(0, source.dataCols);
      source.ver[k] = (versions && versions[i]) || '';
      source.posKeys[i + 1] = k;
    });
    source.changeCount = 0;
    paintBase(source, true);
    updateWriteButton();
  }

  /** Base look: locked columns grey, editable columns white, status column header dark. Clears old change marks. */
  function paintBase(source, includeHeader) {
    var sheet = workbook().getSheetBySheetId(source.sheetId);
    if (!sheet) return;
    var n = source.dataCols;
    var rows = Math.max(sheet.getLastRow(), source.paintedRows || 0, 1) + 20;
    applying(function () {
      source.columns.forEach(function (c, i) {
        var range = sheet.getRange(1, i, rows, 1);
        range.setBackgroundColor(c.editable ? '#FFFFFF' : EDIT_COLORS.locked);
        range.setFontColor(c.editable ? '#242424' : EDIT_COLORS.lockedText);
        // Number formats also cover the empty rows below, so new rows look like the loaded ones.
        if (NUMBER_FORMATS[c.type]) { try { range.setNumberFormat(NUMBER_FORMATS[c.type]); } catch (e) { /* optional */ } }
        if (includeHeader) {
          var h = sheet.getRange(0, i, 1, 1);
          h.setBackgroundColor(c.editable ? EDIT_COLORS.head : '#E1DFDD');
        }
      });
      var st = sheet.getRange(1, n, rows, 1);
      st.setBackgroundColor('#FFFFFF');
      if (includeHeader) sheet.setColumnWidths(n, 1, 340);
      if (includeHeader) {
        var sh = sheet.getRange(0, n, 1, 1);
        sh.setBackgroundColor(EDIT_COLORS.statusHead);
        sh.setFontColor('#FFFFFF');
        sh.setFontWeight('bold');
      }
    });
    source.paintedRows = rows;
  }

  /** Runs fn while edit tracking ignores the mutations it causes (styles, status text, saved values). */
  function applying(fn) {
    state.applying = (state.applying || 0) + 1;
    try { return fn(); } finally { state.applying--; }
  }

  /** Univer mutation hook: protects locked cells and schedules a change check. */
  function onSheetMutation(evt) {
    if (state.applying || !evt || !evt.params) return;
    var id = evt.id, params = evt.params;
    var source = sourceBySheet(params.subUnitId);
    if (!source) return;
    if (id === 'sheet.mutation.set-range-values' && params.cellValue) {
      var reverts = [];
      Object.keys(params.cellValue).forEach(function (rk) {
        var r = +rk, rowCells = params.cellValue[rk];
        if (r < 1) return;
        Object.keys(rowCells).forEach(function (ck) {
          var c = +ck, cell = rowCells[ck];
          if (c >= source.dataCols || !cell || !('v' in cell || 'f' in cell || 'p' in cell)) return;
          var k = source.posKeys[r];
          if (k != null && source.orig[k] && !source.columns[c].editable) reverts.push({ r: r, c: c, v: source.orig[k][c] });
        });
      });
      if (reverts.length) setTimeout(function () { revertCells(source, reverts); }, 0);
      scheduleDiff(source);
    } else if (/reorder-range|move-rows|remove-rows|insert-row|move-range/.test(id)) {
      // Rows moved (sort, insert, delete): marks moved with them, so repaint and check again.
      setTimeout(function () { paintBase(source, false); diffSource(source); }, 0);
    }
  }

  function revertCells(source, reverts) {
    var sheet = workbook().getSheetBySheetId(source.sheetId);
    if (!sheet) return;
    applying(function () {
      reverts.forEach(function (x) { sheet.getRange(x.r, x.c, 1, 1).setValue(x.v == null ? '' : x.v); });
    });
    var cols = reverts.map(function (x) { return source.columns[x.c].caption; })
      .filter(function (v, i, a) { return a.indexOf(v) === i; }).slice(0, 3).join(', ');
    H.toast('warning', cols + (reverts.length > 1 ? ' are' : ' is') + ' read-only here (key, calculated or system field). The value was put back.');
    scheduleDiff(source);
  }

  function scheduleDiff(source) {
    clearTimeout(source._diffTimer);
    source._diffTimer = setTimeout(function () { diffSource(source); }, 350);
  }

  /** Compares the sheet with the loaded rows. Returns the list of changes and marks them in the sheet. */
  function diffSource(source) {
    var sheet = workbook() && workbook().getSheetBySheetId(source.sheetId);
    if (!sheet || !source.columns) return [];
    var n = source.dataCols, last = sheet.getLastRow();
    var raw = last >= 1 ? sheet.getRange(1, 0, last, n + 1).getRawValues() : [];
    var changes = [], seen = {}, posKeys = [];
    raw.forEach(function (full, i) {
      var r = i + 1, row = full.slice(0, n), status = full[n];
      var empty = row.every(function (v) { return v == null || v === ''; });
      if (empty) {
        if (status === STATUS_NEW || status === STATUS_CHANGED) changes.push({ r: r, op: 'clear' });
        return;
      }
      var k = keyOf(source, row), orig = source.orig[k];
      if (orig && !seen[k]) {
        seen[k] = true;
        posKeys[r] = k;
        var cols = [];
        source.columns.forEach(function (c, ci) { if (c.editable && !sameValue(c.type, row[ci], orig[ci])) cols.push(ci); });
        if (cols.length) changes.push({ r: r, op: 'modify', key: k, row: row, cols: cols, status: status });
        else if (status === STATUS_CHANGED) changes.push({ r: r, op: 'clear' });
      } else {
        var filled = [];
        source.columns.forEach(function (c, ci) { if ((c.editable || c.pk) && row[ci] !== '' && row[ci] != null) filled.push(ci); });
        changes.push({ r: r, op: 'insert', row: row, cols: filled, status: status });
      }
    });
    source.posKeys = posKeys;
    markChanges(source, sheet, changes);
    var real = changes.filter(function (x) { return x.op !== 'clear'; });
    source.changeCount = real.length;
    updateWriteButton();
    return real;
  }

  function markChanges(source, sheet, changes) {
    var n = source.dataCols;
    applying(function () {
      (source.marks || []).forEach(function (m) {
        var c = source.columns[m.c];
        if (c) sheet.getRange(m.r, m.c, 1, 1).setBackgroundColor(c.editable ? '#FFFFFF' : EDIT_COLORS.locked);
      });
      source.marks = [];
      changes.forEach(function (x) {
        if (x.op === 'clear') { sheet.getRange(x.r, n, 1, 1).setValue(''); return; }
        if (source.marks.length > 3000) return;
        var color = x.op === 'insert' ? EDIT_COLORS.added : EDIT_COLORS.changed;
        var cols = x.op === 'insert' ? source.columns.map(function (c, i) { return i; }) : x.cols;
        cols.forEach(function (ci) {
          sheet.getRange(x.r, ci, 1, 1).setBackgroundColor(color);
          source.marks.push({ r: x.r, c: ci });
        });
        var label = x.op === 'insert' ? STATUS_NEW : STATUS_CHANGED;
        if (!x.status || x.status === STATUS_CHANGED || x.status === STATUS_NEW || /^Saved/.test(String(x.status))) {
          var cell = sheet.getRange(x.r, n, 1, 1);
          cell.setValue(label);
          cell.setFontColor(x.op === 'insert' ? '#0E700E' : '#835B00');
        }
      });
    });
  }

  function updateWriteButton() {
    if (!ui.write) return;
    var total = editableList().reduce(function (t, s) { return t + (s.changeCount || 0); }, 0);
    ui.write.hidden = editableList().length === 0;
    ui.write.disabled = total === 0 || state.busy;
    ui.writeCount.textContent = total ? String(total) : '';
    ui.writeCount.hidden = !total;
  }

  /** Toolbar "Save to BC": writes every editable sheet that has changes, one sheet after the other. */
  function saveToBC() {
    if (state.busy) return;
    var work = editableList().map(function (s) { return { source: s, changes: diffSource(s) }; })
      .filter(function (w) { return w.changes.length; });
    if (!work.length) { H.toast('info', 'No changes to save to Business Central.'); return; }
    var nMod = 0, nNew = 0;
    work.forEach(function (w) { w.changes.forEach(function (x) { if (x.op === 'insert') nNew++; else nMod++; }); });
    var parts = [];
    if (nMod) parts.push(nMod + (nMod === 1 ? ' changed row' : ' changed rows'));
    if (nNew) parts.push(nNew + (nNew === 1 ? ' new row' : ' new rows'));
    H.choose('Write ' + parts.join(' and ') + ' to Business Central? Each field is checked by BC like on the normal page. Rows with an error are not saved.',
      [{ label: 'Save to BC', value: true, primary: true }], 'Save to Business Central')
      .then(function (ok) {
        if (!ok) return;
        setBusy(true);
        setProgress(-1);
        var totals = { saved: 0, failed: 0 };
        var chain = Promise.resolve();
        work.forEach(function (w) {
          chain = chain.then(function () { return writeSource(w.source, w.changes); })
            .then(function (r) { totals.saved += r.saved; totals.failed += r.failed; });
        });
        return chain.then(function () {
          setBusy(false);
          if (totals.failed) H.toast('warning', 'Saved ' + totals.saved + (totals.saved === 1 ? ' row' : ' rows') + '. ' + totals.failed + (totals.failed === 1 ? ' row has' : ' rows have') + ' an error: see the BC Status column.');
          else H.toast('success', 'Saved ' + totals.saved + (totals.saved === 1 ? ' row' : ' rows') + ' to Business Central.');
        }, function (e) { setBusy(false); reportError(e); });
      });
  }

  function writeSource(source, changes) {
    var limit = 1000;
    if (changes.length > limit) {
      H.toast('warning', 'Saving the first ' + limit + ' rows of ' + sheetName(source.sheetId) + '. Click Save to BC again for the rest.');
      changes = changes.slice(0, limit);
    }
    var keyNos = source.keyIdx.map(function (i) { return source.columns[i].no; });
    var payload = {
      requestId: 'w' + (++state.requestSeq),
      sheetId: source.sheetId,
      tableNo: source.tableNo,
      fields: source.columns.map(function (c) { return c.no; }),
      keyFields: keyNos,
      changes: changes.map(function (x) {
        return {
          row: x.r,
          op: x.op,
          key: source.keyIdx.map(function (i) { return toBC(source.columns[i].type, x.row[i]); }),
          version: x.op === 'modify' ? (source.ver[x.key] || '') : '',
          values: x.cols.filter(function (ci) { return !source.columns[ci].pk; }).map(function (ci) {
            return { no: source.columns[ci].no, value: toBC(source.columns[ci].type, x.row[ci]) };
          })
        };
      })
    };
    return new Promise(function (resolve, reject) {
      state.writes[payload.requestId] = { source: source, resolve: resolve, changes: changes };
      H.sendChunked('write', JSON.stringify(payload), state.config.chunkSize).catch(function (e) {
        delete state.writes[payload.requestId];
        reject(e);
      });
    });
  }

  /** AL answer to a write: shows the saved values (Validate may have changed other fields) and each row's status. */
  function onWriteResult(text) {
    var res = JSON.parse(text);
    var w = state.writes[res.requestId];
    if (!w) return;
    delete state.writes[res.requestId];
    var source = w.source;
    if (res.refused) {
      H.toast('error', res.refused);
      w.resolve({ saved: 0, failed: w.changes.length });
      return;
    }
    var sheet = workbook().getSheetBySheetId(source.sheetId);
    var n = source.dataCols, time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    applying(function () {
      (res.results || []).forEach(function (r) {
        var st = sheet.getRange(r.row, n, 1, 1);
        if (r.ok) {
          var values = source.columns.map(function (c, i) { return convertValue(c.type, r.values[i]); });
          sheet.getRange(r.row, 0, 1, n).setValues([values]);
          var change = w.changes.filter(function (x) { return x.r === r.row; })[0];
          if (change && change.op === 'modify') { delete source.orig[change.key]; delete source.ver[change.key]; }
          var k = keyOf(source, values);
          source.orig[k] = values;
          source.ver[k] = r.version || '';
          st.setValue('Saved ' + time);
          st.setFontColor('#0E700E');
        } else {
          st.setValue(r.error || 'Not saved');
          st.setFontColor('#A4262C');
        }
      });
    });
    setDirty(true);
    diffSource(source);
    w.resolve({ saved: res.saved || 0, failed: res.failed || 0 });
  }

  /** Before a refresh replaces an editable sheet: ask when it still has unsaved BC changes. */
  function confirmDiscard(sources) {
    var pending = sources.filter(function (s) { return s.editable && diffSource(s).length; });
    if (!pending.length) return Promise.resolve(true);
    var n = pending.reduce(function (t, s) { return t + s.changeCount; }, 0);
    return H.choose(n + (n === 1 ? ' row has' : ' rows have') + ' changes that are not saved to Business Central. Refresh replaces them with the data from BC.',
      [{ label: 'Refresh and discard', value: true, primary: true }], 'Unsaved BC changes')
      .then(function (v) { return v === true; });
  }

  // =====================================================================================
  // Combine (Phase 2): join two sheets in the browser
  // =====================================================================================
  /** Reads a sheet as { name, header[], types[], rows[][] } using raw (unformatted) values. */
  function readSheetTable(sheetId) {
    var wb = workbook(), sheet = wb.getSheetBySheetId(sheetId);
    if (!sheet) throw new Error('A sheet used by this combine no longer exists.');
    var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    var name = sheet.getSheetName();
    if (lastRow < 0 || lastCol < 0) return { name: name, header: [], types: [], rows: [] };
    var raw = sheet.getRange(0, 0, lastRow + 1, lastCol + 1).getRawValues();
    var hr = isDerivedOrBC(sheetId) ? 0 : detectHeaderRow(raw);
    var top = raw[hr], sub = !isDerivedOrBC(sheetId) && isSubHeaderRow(raw[hr + 1]) ? raw[hr + 1] : null;
    var clean = function (v) { return v == null || v === '' ? '' : String(v).replace(/\s+/g, ' ').trim(); };
    // Two-row grouped headers ("Invoice details" over "Invoice number | Invoice Date"): the lower label wins.
    var header = top.map(function (h, i) {
      var label = (sub && clean(sub[i])) || clean(h);
      return label || 'Column ' + H.colLetter(i);
    });
    var seen = {};
    header = header.map(function (h) { seen[h] = (seen[h] || 0) + 1; return seen[h] > 1 ? h + ' (' + seen[h] + ')' : h; });
    var rows = raw.slice(hr + (sub ? 2 : 1)).filter(function (r) { return r.some(function (v) { return v != null && v !== ''; }); });
    return { name: name, header: header, types: columnTypes(sheetId, header, rows), rows: rows, headerRow: hr };
  }

  /** A second header line: at least two text labels and no numbers or dates. */
  function isSubHeaderRow(row) {
    if (!row) return false;
    var texts = 0;
    for (var i = 0; i < row.length; i++) {
      var v = row[i];
      if (v == null || v === '') continue;
      if (typeof v !== 'string' || !isNaN(+v.replace(/,/g, '')) || window.AKRecon.toDate(v) != null) return false;
      texts++;
    }
    return texts >= 2;
  }

  function isDerivedOrBC(sheetId) {
    return !!(state.sources[sheetId] || state.joins[sheetId] || state.recons[sheetId]);
  }

  /**
   * Imported files (bank statements, GSTR-2B) often start with title rows and grouped two-row headers.
   * The header is the row among the first 15 with the most text cells that has data under it.
   */
  function detectHeaderRow(raw) {
    var best = 0, bestScore = -1, limit = Math.min(raw.length - 1, 15);
    for (var r = 0; r < limit; r++) {
      var texts = 0, filled = 0;
      raw[r].forEach(function (v) {
        if (v == null || v === '') return;
        filled++;
        if (typeof v === 'string' && isNaN(+v)) texts++;
      });
      var below = raw[r + 1] ? raw[r + 1].filter(function (v) { return v != null && v !== ''; }).length : 0;
      var score = texts * 2 + filled - (below === 0 ? 100 : 0);
      if (score > bestScore) { bestScore = score; best = r; }
    }
    return best;
  }

  /** Column types: from the BC source or combine definition when known, otherwise inferred. */
  function columnTypes(sheetId, header, rows) {
    var known = {};
    var src = state.sources[sheetId], jn = state.joins[sheetId];
    if (src) src.fields.forEach(function (f) { known[f.caption] = f.type; });
    if (jn && jn.outTypes) jn.outHeader.forEach(function (h, i) { known[h] = jn.outTypes[i]; });
    var rn = state.recons[sheetId];
    if (rn && rn.outTypes) rn.outHeader.forEach(function (h, i) { known[h] = rn.outTypes[i]; });
    return header.map(function (h, c) {
      if (known[h]) return known[h];
      var seen = 0;
      for (var r = 0; r < rows.length && r < 500; r++) {
        var v = rows[r][c];
        if (v == null || v === '') continue;
        if (typeof v !== 'number') return 'Text';
        seen++;
      }
      return seen ? 'Number' : 'Text';
    });
  }

  function sheetOptions(selectedId) {
    return workbook().getSheets().map(function (s) {
      var id = s.getSheetId();
      var tag = state.sources[id] ? '  \u00b7  ' + state.sources[id].tableCaption : state.joins[id] ? '  \u00b7  combined' : state.recons[id] ? '  \u00b7  recon' : '';
      if (state.sources[id] && state.sources[id].tableCaption === s.getSheetName()) tag = '  \u00b7  BC';
      return '<option value="' + id + '"' + (id === selectedId ? ' selected' : '') + '>' + H.escapeHtml(s.getSheetName() + tag) + '</option>';
    }).join('');
  }

  function prepareCombine() {
    var sheets = workbook().getSheets().filter(function (s) { return s.getLastRow() >= 1; });
    var ids = sheets.map(function (s) { return s.getSheetId(); });
    var left = ids.indexOf(ui.jLeft.value) >= 0 ? ui.jLeft.value : null;
    var right = ids.indexOf(ui.jRight.value) >= 0 ? ui.jRight.value : null;
    if (!left) {
      var active = workbook().getActiveSheet().getSheetId();
      left = ids.indexOf(active) >= 0 ? active : ids[0];
    }
    if (!right || right === left) right = ids.filter(function (id) { return id !== left; })[0] || '';
    ui.jLeft.innerHTML = sheetOptions(left);
    ui.jRight.innerHTML = sheetOptions(right);
    onCombineSheetsChanged(true);
  }

  function onCombineSheetsChanged(keepKeys) {
    var L = safeHeader(ui.jLeft.value), R = safeHeader(ui.jRight.value);
    state.combine.left = L; state.combine.right = R;
    if (!keepKeys || !keysValid()) state.combine.keys = [suggestKey(L, R, ui.jLeft.value, ui.jRight.value)];
    state.combine.rightCols = {};
    R.forEach(function (h) { state.combine.rightCols[h] = true; });
    var name = H.safeSheetName(sheetName(ui.jLeft.value) + ' + ' + sheetName(ui.jRight.value));
    ui.jName.value = name;
    renderCombineKeys();
  }

  function safeHeader(sheetId) {
    try { return sheetId ? readSheetTable(sheetId).header : []; } catch (e) { return []; }
  }

  function keysValid() {
    var c = state.combine;
    return c.keys.length && c.keys.every(function (k) { return c.left.indexOf(k.left) >= 0 && c.right.indexOf(k.right) >= 0; });
  }

  /**
   * Suggests the key pair by scoring every column combination:
   *   same caption 100 | "<table name> + caption" on the other side 90 ("No." vs "Customer No.")
   *   one caption ends with the other 40 | +15 when the main column is a primary-key field.
   */
  function suggestKey(L, R, leftId, rightId) {
    var norm = function (x) { return String(x).toLowerCase().replace(/[^a-z0-9]/g, ''); };
    var tableName = function (id) {
      var src = id && state.sources[id];
      return norm(src ? src.tableCaption : id ? sheetName(id) : '');
    };
    var pkOf = function (id) {
      var src = id && state.sources[id];
      return src ? src.fields.filter(function (f) { return f.pk; }).map(function (f) { return f.caption; }) : [];
    };
    var lt = tableName(leftId), rt = tableName(rightId), lpk = pkOf(leftId), rpk = pkOf(rightId);
    var best = null, bestScore = 0;
    L.forEach(function (l) {
      R.forEach(function (r) {
        var a = norm(l), b = norm(r), score = 0;
        if (!a || !b) return;
        if (a === b) score = 100;
        else if (b === lt + a || a === rt + b) score = 90;
        else if ((b.length > a.length && b.slice(-a.length) === a) || (a.length > b.length && a.slice(-b.length) === b)) score = 40;
        if (!score) return;
        if (lpk.indexOf(l) >= 0 || rpk.indexOf(r) >= 0) score += 15;
        if (score > bestScore) { bestScore = score; best = { left: l, right: r }; }
      });
    });
    return best || { left: L[0] || '', right: R[0] || '' };
  }

  function options(list, selected) {
    return list.map(function (h) {
      return '<option' + (h === selected ? ' selected' : '') + '>' + H.escapeHtml(h) + '</option>';
    }).join('');
  }

  function renderCombineKeys() {
    var c = state.combine;
    ui.jKeys.innerHTML = c.keys.map(function (k, i) {
      return '<div class="ak-keyrow">' +
        '<select class="ak-input ak-input--sm ak-select" data-side="left" data-index="' + i + '" aria-label="Main sheet key column">' + options(c.left, k.left) + '</select>' +
        '<span class="ak-keyrow__eq">=</span>' +
        '<select class="ak-input ak-input--sm ak-select" data-side="right" data-index="' + i + '" aria-label="Lookup sheet key column">' + options(c.right, k.right) + '</select>' +
        (c.keys.length > 1 ? '<button class="ak-icon-btn" data-act="j-remove-key" data-index="' + i + '" aria-label="Remove key">' + H.icon('x') + '</button>' : '<span class="ak-keyrow__pad"></span>') +
        '</div>';
    }).join('');
    // Selects show the first option when the stored value is empty; sync state with what is displayed.
    ui.jKeys.querySelectorAll('select').forEach(function (s) {
      c.keys[+s.getAttribute('data-index')][s.getAttribute('data-side')] = s.value;
    });
    renderCombineCols();
    updateCombineState();
  }

  function renderCombineCols() {
    var c = state.combine;
    var keyCols = c.keys.map(function (k) { return k.right; });
    ui.jCols.innerHTML = c.right.filter(function (h) { return keyCols.indexOf(h) < 0; }).map(function (h) {
      var on = c.rightCols[h] !== false;
      return '<label class="ak-chipbox' + (on ? ' is-checked' : '') + '"><input type="checkbox" value="' + H.escapeHtml(h) + '"' + (on ? ' checked' : '') + '>' +
        '<span>' + H.escapeHtml(h) + '</span></label>';
    }).join('') || '<span class="ak-list__empty">No other columns</span>';
  }

  function setAllCombineCols(on) {
    state.combine.right.forEach(function (h) { state.combine.rightCols[h] = on; });
    renderCombineCols(); updateCombineState();
  }

  function updateCombineState() {
    if (!ui.jRun) return;
    var c = state.combine;
    var ok = !!(ui.jLeft.value && ui.jRight.value && c.left && c.left.length && c.right && c.right.length && keysValid() && !state.busy);
    var same = ui.jLeft.value && ui.jLeft.value === ui.jRight.value;
    ui.jRun.disabled = !ok || same;
    ui.jHint.textContent = workbook() && workbook().getSheets().filter(function (s) { return s.getLastRow() >= 1; }).length < 2 ? 'Load or open two sheets first'
      : same ? 'Pick two different sheets'
      : !ok ? 'Choose the key columns'
      : 'Match on ' + c.keys.map(function (k) { return k.left; }).join(' + ');
  }

  function radioValue(name) {
    var el = ui.panel.querySelector('input[name=' + name + ']:checked');
    return el ? el.value : null;
  }

  function runCombineFromPanel() {
    var c = state.combine;
    var keyCols = c.keys.map(function (k) { return k.right; });
    var def = {
      leftSheetId: ui.jLeft.value,
      rightSheetId: ui.jRight.value,
      keys: c.keys.map(function (k) { return { left: k.left, right: k.right }; }),
      joinType: radioValue('ak-j-type') || 'left',
      mode: radioValue('ak-j-mode') || 'sum',
      rightColumns: c.right.filter(function (h) { return keyCols.indexOf(h) < 0 && c.rightCols[h] !== false; }),
      status: ui.jStatus.checked,
      ignoreCase: ui.jCase.checked,
      createdAt: new Date().toISOString(),
      lastRows: 0,
      lastCols: 0
    };
    var wb = workbook();
    var sheet = wb.insertSheet(uniqueSheetName(ui.jName.value || 'Combined'));
    def.sheetId = sheet.getSheetId();
    try {
      var r = runJoin(def);
      wb.setActiveSheet(sheet);
      H.toast('success', 'Combined into ' + H.formatInt(r.rows) + ' rows: ' + r.matched + ' matched, ' + r.leftOnly + ' main only, ' + r.rightOnly + ' lookup only.');
      if (window.innerWidth < 900) togglePanel(false);
    } catch (e) {
      if (wb.getSheets().length > 1) withoutDirty(function () { wb.deleteSheet(sheet); });
      reportError(e);
    }
  }

  /** Runs a combine definition and writes the result. Returns counts. Throws on missing sheets/columns. */
  function runJoin(def) {
    var L = readSheetTable(def.leftSheetId), R = readSheetTable(def.rightSheetId);
    var colIndex = function (t, name) {
      var i = t.header.indexOf(name);
      if (i < 0) throw new Error('Column "' + name + '" was not found in sheet "' + t.name + '".');
      return i;
    };
    var lk = def.keys.map(function (k) { return colIndex(L, k.left); });
    var rk = def.keys.map(function (k) { return colIndex(R, k.right); });
    var rc = def.rightColumns.filter(function (h) { return R.header.indexOf(h) >= 0; }).map(function (h) { return R.header.indexOf(h); });

    var norm = function (v) {
      if (v == null) return '';
      if (typeof v === 'number') return String(Math.round(v * 1e6) / 1e6);
      var s = String(v);
      return def.ignoreCase ? s.trim().replace(/\s+/g, ' ').toLowerCase() : s;
    };
    var keyOf = function (row, idx) { return idx.map(function (i) { return norm(row[i]); }).join('\u0001'); };

    var groups = new Map();
    R.rows.forEach(function (row) {
      var k = keyOf(row, rk);
      if (k.replace(/\u0001/g, '') === '') return; // rows without a key cannot match
      var g = groups.get(k);
      if (!g) groups.set(k, g = []);
      g.push(row);
    });

    // Output header (rename lookup columns that clash with main columns).
    var rightNames = rc.map(function (i) {
      var h = R.header[i];
      return L.header.indexOf(h) >= 0 ? h + ' (' + R.name + ')' : h;
    });
    // Match status goes first (column A) so a reconciler sees it without scrolling.
    var header = (def.status ? ['Match'] : []).concat(L.header, rightNames);
    var types = (def.status ? ['Text'] : []).concat(L.types, rc.map(function (i) { return R.types[i]; }));
    if (def.mode === 'sum') { header.push('Match count'); types.push('Integer'); }

    // Only amounts are summed; Integer columns in BC are mostly IDs (Entry No.) and take the first value.
    var numericRight = rc.map(function (i) { return R.types[i] === 'Decimal' || R.types[i] === 'Number'; });
    var pick = function (rows) {
      // Returns a list of "right parts" (arrays aligned with rc) according to the match mode.
      if (def.mode === 'first') return [rc.map(function (i) { return rows[0][i]; })];
      if (def.mode === 'all') return rows.map(function (row) { return rc.map(function (i) { return row[i]; }); });
      return [rc.map(function (i, n) {
        if (numericRight[n]) return rows.reduce(function (s, row) { return s + (typeof row[i] === 'number' ? row[i] : 0); }, 0);
        return rows[0][i];
      })];
    };
    var blankRight = rc.map(function () { return ''; });
    var counts = { matched: 0, leftOnly: 0, rightOnly: 0 };
    var out = [];
    var used = new Set();
    var boolCols = [];
    types.forEach(function (t, i) { if (t === 'Boolean') boolCols.push(i); });
    var emit = function (leftPart, rightPart, count, status) {
      var row = (def.status ? [status] : []).concat(leftPart, rightPart);
      if (def.mode === 'sum') row.push(count);
      // Raw sheet values store booleans as 1/0; write them back as TRUE/FALSE.
      boolCols.forEach(function (i) { if (row[i] === 1 || row[i] === 0) row[i] = row[i] === 1; });
      out.push(row);
    };

    L.rows.forEach(function (row) {
      var k = keyOf(row, lk);
      var matches = groups.get(k);
      var lp = L.header.map(function (h, i) { return row[i] == null ? '' : row[i]; });
      if (matches) {
        used.add(k);
        counts.matched++;
        if (def.joinType === 'anti') return;
        pick(matches).forEach(function (rp) { emit(lp, rp, matches.length, 'Matched'); });
      } else {
        counts.leftOnly++;
        if (def.joinType === 'inner') return;
        emit(lp, blankRight, 0, 'Main only');
      }
    });

    if (def.joinType === 'full' || def.joinType === 'anti') {
      groups.forEach(function (rows, k) {
        if (used.has(k)) return;
        counts.rightOnly++;
        // Put the key values into the main key columns so the row is identifiable.
        var lp = L.header.map(function () { return ''; });
        lk.forEach(function (li, n) { lp[li] = rows[0][rk[n]]; });
        pick(rows).forEach(function (rp) { emit(lp, rp, rows.length, 'Lookup only'); });
      });
    }

    withoutDirty(function () { writeTable(def, header, types, out); });
    if (def.status) colorStatus(def.sheetId, 0, out.length);
    def.outHeader = header;
    def.outTypes = types;
    def.builtAt = new Date().toISOString();
    state.joins[def.sheetId] = def;
    setDirty(true);
    renderSources(); renderBarSources();
    return { rows: out.length, matched: counts.matched, leftOnly: counts.leftOnly, rightOnly: counts.rightOnly };
  }

  var JOIN_STATUS_COLORS = [['Matched', '#DFF6DD', '#0E700E'], ['Main only', '#FFF4CE', '#835B00'], ['Lookup only', '#FDE7E9', '#A4262C']];

  /** Conditional colouring of a status column, using Univer conditional formatting. rules: [[text, bg, font]]. */
  function colorStatus(sheetId, col, n, rules, startRow) {
    if (!n) return;
    rules = rules || JOIN_STATUS_COLORS;
    var sheet = workbook().getSheetBySheetId(sheetId);
    try {
      var range = sheet.getRange(startRow || 1, col, n, 1);
      (sheet.getConditionalFormattingRules() || []).forEach(function (rule) {
        if (rule.ranges && rule.ranges.some(function (rg) { return rg.startColumn === col; })) sheet.deleteConditionalFormattingRule(rule.cfId);
      });
      rules.forEach(function (s) {
        var rule = sheet.newConditionalFormattingRule().whenTextEqualTo(s[0]).setBackground(s[1]).setFontColor(s[2]).setRanges([range.getRange()]).build();
        sheet.addConditionalFormattingRule(rule);
      });
    } catch (e) { /* conditional formatting optional */ }
  }


  // =====================================================================================
  // Reconcile (Phase 3): books vs statement / GSTR-2B / party ledger, engine in recon.js
  // =====================================================================================
  var RECON_PRESETS = {
    gst: {
      left: 'Books', right: 'GSTR-2B', normalize: true, group: true, flip: false, tol: 1, ad: false, days: 3,
      leftSheet: ['gst'],
      tip: 'Tip: load Detailed GST Ledger Entry with GST Component Code filter IGST|CGST so each line\'s taxable value counts once, then import the GSTR-2B Excel as new sheets.',
      keys: [[['gstin', 'buyersellerregno', 'gstregistration', 'registrationno', 'regno'], ['gstinofsupplier', 'suppliergstin', 'gstin']],
        [['externaldocument', 'invoiceno', 'documentno', 'invoice'], ['invoicenumber', 'invoiceno', 'documentnumber', 'documentno', 'invoice']]],
      lamt: ['gstbaseamount', 'taxable', 'baseamount', 'amount'], ramt: ['taxablevalue', 'taxable', 'amount'],
      ldate: ['documentdate', 'postingdate', 'date'], rdate: ['invoicedate', 'documentdate', 'date']
    },
    bank: {
      left: 'G/L', right: 'Bank', normalize: true, group: false, flip: false, tol: 0, ad: true, days: 3,
      tip: 'Tip: load Bank Account Ledger Entry filtered on one bank account and period, then import the bank statement Excel as new sheets.',
      leftSheet: ['bankaccountledger', 'glentry', 'ledger'], keys: [],
      lamt: ['amount', 'debitamount'], ramt: ['credit', 'deposit'], rminus: ['debit', 'withdrawal'], ramtFallback: ['amount'],
      ldate: ['postingdate', 'documentdate', 'date'], rdate: ['txndate', 'transactiondate', 'valuedate', 'date']
    },
    party: {
      left: 'Ours', right: 'Theirs', normalize: true, group: false, flip: true, tol: 1, ad: true, days: 5,
      tip: 'Tip: load Vendor (or Customer) Ledger Entry filtered on one party, then import their statement. Your payable is their receivable, so signs are opposite.',
      leftSheet: ['vendorledger', 'customerledger', 'ledger'],
      keys: [[['externaldocument', 'documentno'], ['documentno', 'invoiceno', 'billno', 'reference', 'document']]],
      lamt: ['amountlcy', 'amount'], ramt: ['amount', 'balance', 'debit'],
      ldate: ['postingdate', 'documentdate'], rdate: ['documentdate', 'invoicedate', 'date']
    }
  };
  var RECON_COLORS = function (labels) {
    return [[labels.match, '#DFF6DD', '#0E700E'], [labels.diff, '#FFF4CE', '#835B00'],
      [labels.left, '#FDE7E9', '#A4262C'], [labels.right, '#EFE6FB', '#5B2C9A']];
  };

  function dataSheets() {
    return workbook().getSheets().filter(function (s) { return s.getLastRow() >= 1; });
  }

  function prepareRecon() {
    var ids = dataSheets().map(function (s) { return s.getSheetId(); });
    var keep = function (v) { return ids.indexOf(v) >= 0 ? v : ''; };
    var l = keep(ui.rLeft.value), r = keep(ui.rRight.value);
    ui.rLeft.innerHTML = sheetOptions(l);
    ui.rRight.innerHTML = sheetOptions(r);
    if (!state.recon.inited || !l || !r) {
      state.recon.inited = true;
      applyReconPreset(state.recon.preset, true);
    } else {
      onReconSheetsChanged(true);
    }
  }

  /** Picks the sheets a template most likely needs: a BC table for the books, an imported sheet for the other side. */
  function pickReconSheets(preset) {
    var sheets = dataSheets(), norm = function (x) { return String(x).toLowerCase().replace(/[^a-z0-9]/g, ''); };
    var bc = sheets.filter(function (s) { return state.sources[s.getSheetId()]; });
    var other = sheets.filter(function (s) { var id = s.getSheetId(); return !state.sources[id] && !state.joins[id] && !state.recons[id]; });
    var left = null;
    (preset && preset.leftSheet || []).some(function (pat) {
      left = bc.filter(function (s) { return norm(state.sources[s.getSheetId()].tableCaption).indexOf(pat) >= 0; })[0] || null;
      return !!left;
    });
    left = left || bc[0] || sheets[0];
    // Other side: the imported sheet whose columns fit the template best (ties: the most recently added).
    var find = window.AKRecon.findColumn;
    var pats = preset ? [preset.ramt, preset.rminus, preset.rdate].concat((preset.keys || []).map(function (k) { return k[1]; })) : [];
    var right = null, bestScore = -1;
    other.filter(function (s) { return s !== left; }).forEach(function (s) {
      var hdr = safeHeader(s.getSheetId()), score = 0;
      pats.forEach(function (pt) { if (pt && find(hdr, pt)) score++; });
      if (score >= bestScore) { bestScore = score; right = s; }
    });
    right = right || sheets.filter(function (s) { return s !== left; })[0];
    if (left) ui.rLeft.value = left.getSheetId();
    if (right) ui.rRight.value = right.getSheetId();
  }

  function applyReconPreset(name, pickSheets) {
    state.recon.preset = name;
    var p = RECON_PRESETS[name];
    var radio = ui.panel.querySelector('input[name=ak-r-preset][value=' + name + ']');
    if (radio) radio.checked = true;
    ui.rTip.textContent = p ? p.tip : 'Choose the sheets, keys, amount and date columns yourself.';
    if (p) {
      ui.rLlabel.value = p.left;
      ui.rRlabel.value = p.right;
      ui.rNorm.checked = p.normalize;
      ui.rGroup.checked = p.group;
      ui.rFlip.checked = p.flip;
      ui.rTol.value = p.tol;
      ui.rAd.checked = p.ad;
      ui.rDays.value = p.days;
      if (pickSheets) pickReconSheets(p);
    }
    onReconSheetsChanged(!p);
  }

  /** Re-reads both headers, fills the column pickers and (for templates) guesses every column. */
  function onReconSheetsChanged(keepChoices) {
    var L = safeHeader(ui.rLeft.value), R = safeHeader(ui.rRight.value);
    var c = state.recon, p = keepChoices === true ? null : RECON_PRESETS[c.preset];
    var find = window.AKRecon.findColumn;
    var prev = {
      lamt: ui.rLamt.value, lminus: ui.rLminus.value, ramt: ui.rRamt.value, rminus: ui.rRminus.value,
      ldate: ui.rLdate.value, rdate: ui.rRdate.value
    };
    c.left = L; c.right = R;
    var pickL = function (k, pats) { return p ? (pats ? find(L, pats) : '') : (L.indexOf(prev[k]) >= 0 ? prev[k] : ''); };
    var pickR = function (k, pats) { return p ? (pats ? find(R, pats) : '') : (R.indexOf(prev[k]) >= 0 ? prev[k] : ''); };

    var ramt = pickR('ramt', p && p.ramt), rminus = '';
    if (p && p.rminus) rminus = ramt ? find(R, p.rminus) : '';
    if (p && !ramt && p.ramtFallback) ramt = find(R, p.ramtFallback);
    if (!p) rminus = pickR('rminus');

    fillColumnSelect(ui.rLamt, L, pickL('lamt', p && p.lamt) || L[L.length - 1], false);
    fillColumnSelect(ui.rLminus, L, p ? '' : pickL('lminus'), true);
    fillColumnSelect(ui.rRamt, R, ramt || R[R.length - 1], false);
    fillColumnSelect(ui.rRminus, R, rminus, true);
    fillColumnSelect(ui.rLdate, L, pickL('ldate', p && p.ldate), true);
    fillColumnSelect(ui.rRdate, R, pickR('rdate', p && p.rdate), true);

    if (p) {
      c.keys = [];
      p.keys.forEach(function (k) {
        var l = find(L, k[0]), r = find(R, k[1]);
        if (l && r) c.keys.push({ left: l, right: r });
      });
    } else {
      c.keys = c.keys.filter(function (k) { return L.indexOf(k.left) >= 0 && R.indexOf(k.right) >= 0; });
    }
    renderReconKeys();
    renderReconExtras();
    renderReconLabels();
  }

  function fillColumnSelect(el, list, value, optional) {
    el.innerHTML = (optional ? '<option value="">&mdash;</option>' : '') + list.map(function (h) {
      return '<option value="' + H.escapeHtml(h) + '"' + (h === value ? ' selected' : '') + '>' + H.escapeHtml(h) + '</option>';
    }).join('');
    if (!optional && !value && list.length) el.value = list[0];
  }

  function renderReconKeys() {
    var c = state.recon;
    ui.rKeys.innerHTML = c.keys.length ? c.keys.map(function (k, i) {
      return '<div class="ak-keyrow">' +
        '<select class="ak-input ak-input--sm ak-select" data-side="left" data-index="' + i + '" aria-label="Key column, first sheet">' + options(c.left, k.left) + '</select>' +
        '<span class="ak-keyrow__eq">=</span>' +
        '<select class="ak-input ak-input--sm ak-select" data-side="right" data-index="' + i + '" aria-label="Key column, second sheet">' + options(c.right, k.right) + '</select>' +
        '<button class="ak-icon-btn" data-act="r-remove-key" data-index="' + i + '" aria-label="Remove key">' + H.icon('x') + '</button></div>';
    }).join('') : '<div class="ak-list__empty">No key columns: rows are matched by amount and date.</div>';
    ui.rKeys.querySelectorAll('select').forEach(function (el) {
      c.keys[+el.getAttribute('data-index')][el.getAttribute('data-side')] = el.value;
    });
    updateReconState();
  }

  function renderReconExtras() {
    var chips = function (list) {
      return list.map(function (h) {
        return '<label class="ak-chipbox"><input type="checkbox" value="' + H.escapeHtml(h) + '"><span>' + H.escapeHtml(h) + '</span></label>';
      }).join('') || '<span class="ak-list__empty">No columns</span>';
    };
    ui.rLextra.innerHTML = chips(state.recon.left);
    ui.rRextra.innerHTML = chips(state.recon.right);
  }

  function renderReconLabels() {
    var l = ui.rLlabel.value.trim() || 'Books', r = ui.rRlabel.value.trim() || 'Statement';
    ui.panel.querySelectorAll('[data-r-label=left]').forEach(function (e) { e.textContent = l; });
    ui.panel.querySelectorAll('[data-r-label=right]').forEach(function (e) { e.textContent = r; });
  }

  function updateReconState() {
    if (!ui.rRun) return;
    var c = state.recon;
    var two = workbook() && dataSheets().length >= 2;
    var same = ui.rLeft.value && ui.rLeft.value === ui.rRight.value;
    var hasRule = c.keys.length > 0 || ui.rAd.checked;
    var ok = two && !same && ui.rLamt.value && ui.rRamt.value && hasRule && !state.busy;
    ui.rRun.disabled = !ok;
    ui.rHint.textContent = !two ? 'Load or import two sheets first'
      : same ? 'Pick two different sheets'
      : !hasRule ? 'Add a key or match by amount and date'
      : c.keys.length ? 'Match on ' + c.keys.map(function (k) { return k.left; }).join(' + ') + (ui.rAd.checked ? ', then amount + date' : '')
      : 'Match on amount + date';
  }

  function checkedValues(container) {
    return Array.prototype.map.call(container.querySelectorAll('input:checked'), function (i) { return i.value; });
  }

  function runReconFromPanel() {
    var wb = workbook();
    var def = {
      preset: state.recon.preset,
      leftSheetId: ui.rLeft.value,
      rightSheetId: ui.rRight.value,
      leftLabel: ui.rLlabel.value.trim() || 'Books',
      rightLabel: ui.rRlabel.value.trim() || 'Statement',
      keys: state.recon.keys.map(function (k) { return { left: k.left, right: k.right }; }),
      normalize: ui.rNorm.checked,
      group: ui.rGroup.checked,
      amount: { left: ui.rLamt.value, leftMinus: ui.rLminus.value, right: ui.rRamt.value, rightMinus: ui.rRminus.value },
      flip: ui.rFlip.checked,
      tolerance: parseFloat(ui.rTol.value) || 0,
      date: { left: ui.rLdate.value, right: ui.rRdate.value },
      useAmountDate: ui.rAd.checked,
      days: parseInt(ui.rDays.value, 10) || 0,
      extraLeft: checkedValues(ui.rLextra),
      extraRight: checkedValues(ui.rRextra),
      summary: ui.rSummary.checked,
      createdAt: new Date().toISOString(),
      lastRows: 0,
      lastCols: 0
    };
    var base = H.safeSheetName(ui.rName.value || 'Recon');
    var created = [];
    try {
      var sheet = wb.insertSheet(uniqueSheetName(base));
      created.push(sheet);
      def.sheetId = sheet.getSheetId();
      if (def.summary) {
        var sum = wb.insertSheet(uniqueSheetName(base + ' summary'));
        created.push(sum);
        def.summarySheetId = sum.getSheetId();
      }
      var r = runRecon(def);
      wb.setActiveSheet(def.summary ? wb.getSheetBySheetId(def.summarySheetId) : sheet);
      H.toast(r.exceptions ? 'warning' : 'success', r.exceptions
        ? H.formatInt(r.matched) + ' matched, ' + H.formatInt(r.exceptions) + ' to review (' + r.detail + ').'
        : 'Fully reconciled: all ' + H.formatInt(r.matched) + ' rows matched.');
      if (window.innerWidth < 900) togglePanel(false);
    } catch (e) {
      withoutDirty(function () { created.forEach(function (s) { if (wb.getSheets().length > 1) wb.deleteSheet(s); }); });
      delete state.recons[def.sheetId];
      reportError(e);
    }
  }

  /** Runs a reconciliation definition, writes the result (and summary) sheets. Throws on missing sheets/columns. */
  function runRecon(def) {
    var L = readSheetTable(def.leftSheetId), R = readSheetTable(def.rightSheetId);
    var res = window.AKRecon.run(L, R, def);
    withoutDirty(function () { writeTable(def, res.header, res.types, res.rows); });
    colorStatus(def.sheetId, 0, res.rows.length, RECON_COLORS(res.statusLabels));
    if (def.summary) writeReconSummary(def, res, L.name, R.name);
    var st = res.stats;
    def.outHeader = res.header;
    def.outTypes = res.types;
    def.builtAt = new Date().toISOString();
    def.matched = st.match;
    def.exceptions = st.diff + st.left + st.right;
    state.recons[def.sheetId] = def;
    setDirty(true);
    renderSources(); renderBarSources();
    return {
      rows: res.rows.length, matched: st.match, exceptions: def.exceptions,
      detail: st.diff + ' amount differences, ' + st.left + ' only in ' + def.leftLabel + ', ' + st.right + ' only in ' + def.rightLabel
    };
  }

  /** Summary sheet with live COUNTIF / SUMIF formulas over the result sheet, so manual edits update the totals. */
  function writeReconSummary(def, res, leftName, rightName) {
    var wb = workbook();
    var sheet = def.summarySheetId && wb.getSheetBySheetId(def.summarySheetId);
    if (!sheet) {
      sheet = wb.insertSheet(uniqueSheetName(sheetName(def.sheetId) + ' summary'));
      def.summarySheetId = sheet.getSheetId();
    }
    var lr = sheet.getLastRow(), lc = sheet.getLastColumn();
    withoutDirty(function () {
      if (lr >= 0 && lc >= 0) sheet.getRange(0, 0, lr + 1, lc + 1).clearContent();
    });
    var q = "'" + sheetName(def.sheetId).replace(/'/g, "''") + "'";
    var last = Math.max(res.rows.length + 1, 2);
    var rng = function (colIndex) { var c = H.colLetter(colIndex); return q + '!$' + c + '$2:$' + c + '$' + last; };
    var lab = res.statusLabels, statuses = [lab.match, lab.diff, lab.left, lab.right];
    var esc = function (t) { return String(t).replace(/"/g, '""'); };
    var LL = def.leftLabel, RL = def.rightLabel;
    var cells = [
      [{ v: 'Reconciliation summary' }, {}, {}, {}, {}],
      [{ v: leftName + ' (' + LL + ') vs ' + rightName + ' (' + RL + ')  ·  ' + new Date().toLocaleString() }, {}, {}, {}, {}],
      [{}, {}, {}, {}, {}],
      [{ v: 'Status' }, { v: 'Count' }, { v: LL + ' amount' }, { v: RL + ' amount' }, { v: 'Difference' }]
    ];
    statuses.forEach(function (st) {
      cells.push([
        { v: st },
        { f: '=COUNTIF(' + rng(0) + ',"' + esc(st) + '")' },
        { f: '=SUMIF(' + rng(0) + ',"' + esc(st) + '",' + rng(res.leftAmountCol) + ')' },
        { f: '=SUMIF(' + rng(0) + ',"' + esc(st) + '",' + rng(res.rightAmountCol) + ')' },
        { f: '=SUMIF(' + rng(0) + ',"' + esc(st) + '",' + rng(res.diffCol) + ')' }
      ]);
    });
    cells.push([{ v: 'Total' }, { f: '=SUM(B5:B8)' }, { f: '=SUM(C5:C8)' }, { f: '=SUM(D5:D8)' }, { f: '=SUM(E5:E8)' }]);
    cells.push([{}, {}, {}, {}, {}]);
    cells.push([{ v: 'Matched %' }, { f: '=IFERROR(B5/B9,0)' }, {}, {}, {}]);
    cells.push([{ v: 'Rows to review' }, { f: '=B9-B5' }, {}, {}, {}]);

    withoutDirty(function () {
      if (sheet.getMaxRows() < cells.length + 2) sheet.insertRowsAfter(sheet.getMaxRows() - 1, cells.length + 2 - sheet.getMaxRows());
      sheet.getRange(0, 0, cells.length, 5).setValues(cells);
      // Each formatting step is independent: one unsupported call must not skip the rest.
      var tryFmt = function (fn) { try { fn(); } catch (e) { /* formatting optional */ } };
      tryFmt(function () { sheet.setColumnWidths(0, 1, 240); sheet.setColumnWidths(1, 4, 150); });
      tryFmt(function () { sheet.getRange(0, 0, 1, 1).setFontWeight('bold').setFontSize(16); });
      tryFmt(function () { sheet.getRange(1, 0, 1, 1).setFontColor('#616161'); });
      tryFmt(function () { var hd = sheet.getRange(3, 0, 1, 5); hd.setFontWeight('bold'); hd.setBackgroundColor('#EBF3FC'); hd.setFontColor('#0F3D6E'); });
      tryFmt(function () { var tt = sheet.getRange(8, 0, 1, 5); tt.setFontWeight('bold'); tt.setBackgroundColor('#F5F5F5'); });
      tryFmt(function () { sheet.getRange(4, 1, 5, 1).setNumberFormat('#,##0'); });
      tryFmt(function () { sheet.getRange(4, 2, 5, 3).setNumberFormat('#,##0.00'); });
      tryFmt(function () { sheet.getRange(10, 1, 1, 1).setNumberFormat('0.0%'); sheet.getRange(11, 1, 1, 1).setNumberFormat('#,##0'); });
      tryFmt(function () { sheet.getRange(3, 1, 1, 4).setHorizontalAlignment('right'); });
    });
    colorStatus(def.summarySheetId, 0, 4, RECON_COLORS(lab), 4);
  }

  // =====================================================================================
  // Save / load / import / export
  // =====================================================================================
  /** silent = AutoSave: AL stores it without a toast. */
  function requestSave(silent) {
    var wb = workbook();
    if (!wb) return Promise.resolve();
    clearTimeout(state.autosaveTimer);
    state.autosaveTimer = null;
    if (ui.title === document.activeElement) ui.title.blur(); // commit a pending rename first
    var sources = sourceList(), joins = joinList(), recons = reconList();
    var envelope = {
      format: 'ak-workbook',
      version: 3,
      savedAt: new Date().toISOString(),
      snapshot: wb.save(),
      sources: sources.map(function (src) {
        var copy = {};
        Object.keys(src).forEach(function (k) { if (k.charAt(0) !== '_' && k !== 'posKeys' && k !== 'marks') copy[k] = src[k]; });
        return copy;
      }),
      joins: joins,
      recons: recons
    };
    var linked = sources.map(function (s) { return s.tableCaption; })
      .filter(function (v, i, a) { return a.indexOf(v) === i; }).join(', ');
    ui.app.classList.add('is-saving');
    ui.dirty.textContent = 'Saving…';
    return H.invoke('OnWorkbookSaving', [linked])
      .then(function () { return H.sendChunked(silent === true ? 'autosave' : 'save', JSON.stringify(envelope), state.config.chunkSize); })
      .then(function () { ui.app.classList.remove('is-saving'); state.dirty = true; setDirty(false); },
        function (e) { ui.app.classList.remove('is-saving'); ui.dirty.textContent = 'Unsaved changes'; reportError(e); });
  }

  function loadEnvelope(text) {
    var env = JSON.parse(text);
    if (!env || env.format !== 'ak-workbook' || !env.snapshot) throw new Error('The saved workbook could not be read.');
    replaceWorkbook(env.snapshot);
    state.sources = {};
    state.joins = {};
    state.recons = {};
    (env.sources || []).forEach(function (s) { state.sources[s.sheetId] = s; });
    // Editable sheets: rebuild the row positions and change marks from the saved copy of the loaded rows.
    setTimeout(function () { editableList().forEach(function (s) { s.marks = []; diffSource(s); }); updateWriteButton(); }, 600);
    (env.joins || []).forEach(function (j) { state.joins[j.sheetId] = j; });
    (env.recons || []).forEach(function (r) { state.recons[r.sheetId] = r; });
    renderSources(); renderBarSources();
    setTimeout(function () { state.dirty = true; setDirty(false); }, 450);
  }

  function chooseImportFile() {
    var go = function (mode) { state.importMode = mode; ui.file.click(); };
    var hasContent = workbook().getSheets().some(function (s) { return s.getLastRow() >= 0; });
    if (!hasContent) { go('replace'); return; }
    H.choose('Add the file\'s sheets next to your current ones (for example a bank statement or GSTR-2B to reconcile), or replace the whole workbook?',
      [{ label: 'Replace workbook', value: 'replace' }, { label: 'Add as new sheets', value: 'add', primary: true }], 'Import Excel file')
      .then(function (mode) { if (mode) go(mode); });
  }

  function requestExport() {
    var wb = workbook();
    if (!wb) return Promise.resolve();
    setProgress(-1);
    return Promise.resolve().then(function () {
      return buildXlsx(wb);
    }).then(function (buffer) {
      return H.sendChunked('export', H.arrayBufferToBase64(buffer), state.config.chunkSize);
    }).then(function () { setProgress(null); }, function (e) { setProgress(null); reportError(e); });
  }

  /** Univer snapshot -> ExcelJS workbook (values, formulas, styles, widths, merges, freeze, filter). */
  function buildXlsx(fwb) {
    var snap = fwb.save();
    var xl = new window.ExcelJS.Workbook();
    xl.creator = 'Business Central - AK Excel Workbook';
    xl.created = new Date();
    var styles = snap.styles || {};

    (snap.sheetOrder || Object.keys(snap.sheets)).forEach(function (sheetId) {
      var sd = snap.sheets[sheetId];
      if (!sd) return;
      var fsheet = fwb.getSheetBySheetId(sheetId);
      var ySplit = sd.freeze && sd.freeze.ySplit > 0 ? sd.freeze.ySplit : 0;
      var xSplit = sd.freeze && sd.freeze.xSplit > 0 ? sd.freeze.xSplit : 0;
      var ws = xl.addWorksheet(H.safeSheetName(sd.name), {
        views: (ySplit || xSplit) ? [{ state: 'frozen', xSplit: xSplit, ySplit: ySplit }] : [],
        state: sd.hidden === 1 ? 'hidden' : 'visible'
      });

      // Formulas: read through the facade so shared/copied formulas come back fully expanded.
      var formulas = null;
      if (fsheet) {
        var lastRow = fsheet.getLastRow(), lastCol = fsheet.getLastColumn();
        if (lastRow >= 0 && lastCol >= 0) {
          try { formulas = fsheet.getRange(0, 0, lastRow + 1, lastCol + 1).getFormulas(); } catch (e) { formulas = null; }
        }
      }

      var cellData = sd.cellData || {};
      Object.keys(cellData).forEach(function (rKey) {
        var row = cellData[rKey], r = +rKey;
        Object.keys(row).forEach(function (cKey) {
          var cd = row[cKey], c = +cKey;
          if (!cd) return;
          var cell = ws.getCell(r + 1, c + 1);
          var f = (formulas && formulas[r] && formulas[r][c]) || cd.f;
          var v = cd.v;
          if (cd.t === 3 && (v === 1 || v === 0)) v = v === 1;
          if (f) cell.value = { formula: String(f).replace(/^=/, ''), result: v == null ? undefined : v };
          else if (v != null && v !== '') cell.value = v;
          var st = typeof cd.s === 'string' ? styles[cd.s] : cd.s;
          if (st) {
            var xs = H.univerStyleToExcel(st);
            if (xs.font) cell.font = xs.font;
            if (xs.fill) cell.fill = xs.fill;
            if (xs.numFmt) cell.numFmt = xs.numFmt;
            if (xs.alignment) cell.alignment = xs.alignment;
            if (xs.border) cell.border = xs.border;
          }
        });
      });

      var colData = sd.columnData || {};
      Object.keys(colData).forEach(function (cKey) {
        var cd = colData[cKey];
        if (cd && cd.w) ws.getColumn(+cKey + 1).width = Math.max(4, Math.round(cd.w / 7));
        if (cd && cd.hd === 1) ws.getColumn(+cKey + 1).hidden = true;
      });
      var rowData = sd.rowData || {};
      Object.keys(rowData).forEach(function (rKey) {
        var rd = rowData[rKey];
        if (rd && rd.h && rd.h !== sd.defaultRowHeight) ws.getRow(+rKey + 1).height = Math.round(rd.h * 0.75);
        if (rd && rd.hd === 1) ws.getRow(+rKey + 1).hidden = true;
      });
      (sd.mergeData || []).forEach(function (m) {
        try { ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1); } catch (e) { /* overlapping merge */ }
      });
      if (fsheet && fsheet.getFilter) {
        try {
          var flt = fsheet.getFilter();
          var fr = flt && flt.getRange && flt.getRange();
          if (fr) ws.autoFilter = { from: { row: fr.getRow() + 1, column: fr.getColumn() + 1 }, to: { row: fr.getRow() + 1, column: fr.getColumn() + fr.getWidth() } };
        } catch (e) { /* optional */ }
      }
    });
    if (!xl.worksheets.length) xl.addWorksheet('Sheet1');
    return xl.xlsx.writeBuffer();
  }

  /** .xlsx bytes -> Univer sheets. mode 'add' appends the sheets; otherwise the workbook is replaced. */
  function importXlsxBuffer(buffer, mode) {
    setProgress(-1);
    var xl = new window.ExcelJS.Workbook();
    return xl.xlsx.load(buffer).then(function () {
      var snapshot = { id: H.uid('wb'), name: 'Workbook', locale: 'enUS', sheetOrder: [], sheets: {}, styles: {} };
      xl.worksheets.forEach(function (ws) {
        var id = H.uid('sh');
        var cellData = {}, maxR = 0, maxC = 0;
        ws.eachRow({ includeEmpty: false }, function (row, rn) {
          row.eachCell({ includeEmpty: false }, function (cell, cn) {
            var r = rn - 1, c = cn - 1, cd = {};
            var v = cell.value;
            var formula = cell.formula || (v && v.formula) || null;
            if (formula) {
              cd.f = '=' + formula;
              var res = v && v.result;
              if (res != null && typeof res !== 'object') cd.v = res;
              else if (res instanceof Date) cd.v = H.dateToSerial(res);
            } else if (v instanceof Date) {
              cd.v = H.dateToSerial(v);
            } else if (v && typeof v === 'object') {
              if (v.richText) cd.v = v.richText.map(function (p) { return p.text; }).join('');
              else if (v.text != null) cd.v = typeof v.text === 'object' && v.text.richText ? v.text.richText.map(function (p) { return p.text; }).join('') : v.text;
              else if (v.error) cd.v = v.error;
            } else if (v != null) {
              cd.v = v;
            }
            if (typeof cd.v === 'boolean') { cd.v = cd.v ? 1 : 0; cd.t = 3; }
            var st = H.excelStyleToUniver(cell);
            if (v instanceof Date && !(st && st.n)) { st = st || {}; st.n = { pattern: 'dd-mmm-yyyy' }; }
            if (st) cd.s = st;
            if (cd.v == null && !cd.f && !cd.s) return;
            (cellData[r] = cellData[r] || {})[c] = cd;
            if (r > maxR) maxR = r;
            if (c > maxC) maxC = c;
          });
        });
        var columnData = {};
        (ws.columns || []).forEach(function (col, i) {
          if (col && col.width) columnData[i] = { w: Math.round(col.width * 7 + 5) };
          if (col && col.hidden) (columnData[i] = columnData[i] || {}).hd = 1;
        });
        var mergeData = [];
        ((ws.model && ws.model.merges) || []).forEach(function (ref) { var m = H.decodeA1Range(ref); if (m) mergeData.push(m); });
        var view = (ws.views || [])[0] || {};
        var freeze = view.state === 'frozen'
          ? { xSplit: view.xSplit || 0, ySplit: view.ySplit || 0, startRow: view.ySplit || -1, startColumn: view.xSplit || -1 }
          : undefined;
        snapshot.sheetOrder.push(id);
        snapshot.sheets[id] = {
          id: id,
          name: H.safeSheetName(ws.name),
          rowCount: Math.max(maxR + 101, 200),
          columnCount: Math.max(maxC + 11, 26),
          cellData: cellData,
          columnData: columnData,
          mergeData: mergeData,
          freeze: freeze,
          hidden: ws.state === 'hidden' ? 1 : 0
        };
      });
      if (!snapshot.sheetOrder.length) throw new Error('The file has no worksheets.');
      var n = snapshot.sheetOrder.length;
      if (mode === 'add') {
        var wb = workbook(), first = null;
        snapshot.sheetOrder.forEach(function (id) {
          var sd = snapshot.sheets[id];
          var sheet = wb.create(uniqueSheetName(sd.name), sd.rowCount, sd.columnCount, {
            sheet: { cellData: sd.cellData, columnData: sd.columnData, mergeData: sd.mergeData, freeze: sd.freeze, hidden: sd.hidden }
          });
          first = first || sheet;
        });
        if (first) wb.setActiveSheet(first);
        setDirty(true);
        setProgress(null);
        H.toast('success', 'Added ' + n + (n === 1 ? ' sheet' : ' sheets') + ' from the file.');
        return;
      }
      replaceWorkbook(snapshot);
      state.sources = {};
      state.joins = {};
      state.recons = {};
      renderSources(); renderBarSources();
      setTimeout(function () { setDirty(true); }, 450);
      setProgress(null);
      H.toast('success', 'Imported ' + n + (n === 1 ? ' sheet' : ' sheets') + '. Save to keep it.');
    }).catch(function (e) { setProgress(null); reportError(e); });
  }

  // =====================================================================================
  // Error boundary
  // =====================================================================================
  var lastErrorAt = 0;
  function reportError(e) {
    var msg = (e && (e.message || e.reason && e.reason.message)) || String(e);
    console.error('[AK]', e);
    H.toast('error', msg);
    if (e && typeof e === 'object') e.shown = true;
    var now = Date.now();
    if (now - lastErrorAt > 2000) {
      lastErrorAt = now;
      H.invoke('OnError', [String(msg).substring(0, 2000)]).catch(function () { });
    }
  }

  function guard(fn) {
    return function () {
      try { return fn.apply(this, arguments); } catch (e) { reportError(e); }
    };
  }

  window.addEventListener('error', function (e) {
    if (e && e.message && e.message.indexOf('ResizeObserver') >= 0) return; // benign browser warning
    reportError(e.error || e.message);
  });
  window.addEventListener('unhandledrejection', function (e) { reportError(e.reason || 'Unexpected error'); });
  window.addEventListener('pagehide', destroy);
  window.addEventListener('beforeunload', destroy);

  // =====================================================================================
  // Procedures called by AL (names match the controladdin definition)
  // =====================================================================================
  window.Init = guard(function (configJson) {
    if (configJson) {
      var cfg = JSON.parse(configJson);
      Object.keys(cfg).forEach(function (k) { state.config[k] = cfg[k]; });
    }
    ui.maxRows.value = state.config.defaultMaxRows;
    ui.maxRows.max = state.config.maxRowsLimit;
    ui.title.value = state.config.description || '';
    ui.code.textContent = state.config.workbookCode || '';
    ui.code.hidden = !state.config.workbookCode;
    startUniver();
    H.invoke('AddInFrameReady', []);
  });

  window.Refresh = guard(function () { refreshAll(); });

  window.TogglePanel = guard(function () { togglePanel('data'); });

  window.SetTables = guard(function (json) {
    state.tables = JSON.parse(json || '[]');
    renderTables();
  });

  window.SetFields = guard(function (json) { applyFieldsResponse(JSON.parse(json)); });

  window.AppendData = guard(function (json) { onChunk(JSON.parse(json)); });

  window.DataError = guard(function (requestId, message) { failRequest(requestId, message); });

  window.ReceiveChunk = guard(function (purpose, index, count, data) {
    var full = state.assembler.add(purpose, index, count, data);
    if (full === null) return;
    if (purpose === 'load') loadEnvelope(full);
    else if (purpose === 'import') importXlsxBuffer(H.base64ToArrayBuffer(full), 'replace');
    else if (purpose === 'writeresult') onWriteResult(full);
  });

  window.RequestSave = guard(function () { requestSave(); });

  window.RequestExport = guard(function () { requestExport(); });

  window.ShowToast = guard(function (level, message) { H.toast(level, message); });

  // Exposed for automated tests only.
  window.AKWorkbookDebug = { state: state, diffSource: diffSource, saveToBC: saveToBC, buildXlsx: buildXlsx, runJoin: runJoin, runRecon: runRecon, importXlsxBuffer: importXlsxBuffer };

  // =====================================================================================
  // Boot
  // =====================================================================================
  try {
    buildShell();
    H.invoke('ControlAddInReady', []);
  } catch (e) {
    reportError(e);
  }
})();
