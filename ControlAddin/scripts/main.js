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

  var ADDIN_VERSION = '2.2.4.0';  // shown on the logo tooltip, so testers can see which script BC loaded
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
    budgetPending: {},      // budget load requestId -> { source, resolve, reject }
    slicers: {},            // period pickers: id -> { sheetId, row, col, kind }
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
      '      <span class="ak-bar__logo" title="AK Excel Workbook, add-in ' + ADDIN_VERSION + '">' + H.icon('sheet') + '</span>' +
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
      '      <button class="ak-btn ak-btn--outline" data-act="insert-chart" title="Insert a chart or KPI card from the selected cells (also on the right-click menu)">' + H.icon('chart') + '<span>Insert chart</span></button>' +
      '      <button class="ak-btn ak-btn--outline" data-act="view-combine" aria-controls="ak-panel" title="Join two sheets on matching keys">' + H.icon('merge') + '<span>Combine</span></button>' +
      '      <button class="ak-btn ak-btn--outline" data-act="view-recon" aria-controls="ak-panel" title="Reconcile books against a statement, GSTR-2B or party ledger">' + H.icon('scale') + '<span>Reconcile</span></button>' +
      '      <button class="ak-btn ak-btn--primary" data-act="insert-region" aria-controls="ak-panel" title="Put a Business Central table at the selected cell (also on the right-click menu)">' + H.icon('database') + '<span>Insert region</span></button>' +
      '    </div>' +
      '    <input type="file" id="ak-file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden>' +
      '  </header>' +
      '  <div class="ak-regionbar" id="ak-regionbar" hidden></div>' +
      '  <div class="ak-progress" id="ak-progress" hidden><div class="ak-progress__bar" id="ak-progress-bar"></div></div>' +
      '  <div class="ak-body">' +
      '    <main class="ak-sheet" id="ak-sheet"></main>' +
      '    <aside class="ak-panel" id="ak-panel" data-view="data">' + dataViewHtml() + combineViewHtml() + reconViewHtml() + budgetViewHtml() + '</aside>' +
      '  </div>' +
      '  <div class="ak-toasts" id="ak-toasts" aria-live="polite"></div>' +
      '  <div class="ak-boot" id="ak-boot"><div class="ak-spinner"></div><div>Starting spreadsheet&hellip;</div></div>' +
      '</div>';

    var ids = ['app', 'sheet', 'panel', 'dirty', 'title', 'code', 'file', 'bar-sources', 'progress', 'progress-bar', 'boot',
      'table-search', 'table-results', 'table-picked', 'step-fields', 'field-search', 'field-list', 'field-count',
      'max-rows', 'sheet-name', 'load', 'load-hint', 'source-list', 'edit', 'edit-wrap', 'edit-why', 'regionbar', 'dest-cell', 'dest-wrap', 'region-target', 'load-label', 'quick',
      'b-name', 'b-start', 'b-periods', 'b-acc', 'b-d1', 'b-d2', 'b-d1cap', 'b-d2cap', 'b-ly', 'b-dest-wrap', 'b-dest-cell', 'b-why', 'b-hint', 'b-run', 'b-run-label', 'write', 'write-count',
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
      '<div class="ak-panel__head"><div><div class="ak-panel__title">Business Central region</div>' +
      '<div class="ak-panel__sub" id="ak-region-target">New region</div></div>' +
      '<button class="ak-icon-btn" data-act="panel-close" aria-label="Close panel">' + H.icon('x') + '</button></div>' +
      '<div class="ak-panel__scroll">' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">1</span>Table</div>' +
      '    <div class="ak-quick" id="ak-quick" hidden></div>' +
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
      '    <p class="ak-help">Filters use Business Central syntax: <code>&gt;0</code>, <code>10000..20000</code>, <code>&lt;&gt;&#39;&#39;</code>, <code>A*|B*</code>, <code>01-04-26..31-03-27</code>. ' +
      'Put <code>[B1]</code> to take the filter from a cell (also <code>[B1]..[C1]</code> or <code>[Sheet1!B1]</code>): change the cell and the region refreshes. ' +
      'For dates, point to a <b>period picker</b> cell (Quick start or right-click).</p>' +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">3</span>Output</div>' +
      '    <div class="ak-grid2">' +
      '      <label class="ak-field"><span>Sheet name</span><input id="ak-sheet-name" class="ak-input" maxlength="31" placeholder="Table name"></label>' +
      '      <label class="ak-field"><span>Max rows</span><input id="ak-max-rows" class="ak-input" type="number" min="1" step="1000"></label>' +
      '    </div>' +
      '    <div class="ak-seg" role="radiogroup" aria-label="Where" id="ak-dest-wrap">' +
      '      <label><input type="radio" name="ak-dest" value="here" checked><span>At <b id="ak-dest-cell">A1</b></span></label>' +
      '      <label><input type="radio" name="ak-dest" value="new"><span>New sheet</span></label>' +
      '    </div>' +
      '    <label class="ak-check ak-edit-opt" id="ak-edit-wrap" hidden><input type="checkbox" id="ak-edit">' +
      '<span><b>Allow editing</b><br>Change values in the sheet and use <b>Save to BC</b>. Key and calculated fields stay locked.' +
      '<span class="ak-edit-why" id="ak-edit-why" hidden></span></span></label>' +
      '  </section>' +
      '  <section class="ak-step ak-step--sources">' +
      '    <div class="ak-step__label">' + H.icon('link') + 'Regions and linked sheets</div>' +
      '    <ul class="ak-sources" id="ak-source-list"></ul>' +
      '  </section>' +
      '</div>' +
      '<div class="ak-panel__foot">' +
      '  <span class="ak-panel__hint" id="ak-load-hint">Pick a table to start</span>' +
      '  <button class="ak-btn ak-btn--primary ak-btn--lg" id="ak-load" disabled>' + H.icon('database') + '<span id="ak-load-label">Insert region</span></button>' +
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
        case 'insert-region': openRegionPanel(null); break;
        case 'quick-start': startQuick(btn.getAttribute('data-quick')); break;
        case 'open-budget': openBudgetPanel(null); break;
        case 'period-picker': insertPeriodPicker(); break;
        case 'insert-chart': openChartPicker(); break;
        case 'region-props': if (activeRegion()) { if (isBudget(activeRegion())) openBudgetPanel(activeRegion()); else openRegionPanel(activeRegion()); } break;
        case 'region-refresh': if (activeRegion()) refreshOne(activeRegion().id); break;
        case 'region-open': drillAt(selectedCell()); break;
        case 'region-select': if (activeRegion()) selectRegion(activeRegion(), 'toggle'); break;
        case 'region-remove': if (activeRegion()) askRemoveRegion(activeRegion()); break;
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
      if (e.target.hasAttribute('data-list')) { fs.list = e.target.checked; return; }
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
    ui.loadBtn.addEventListener('click', function () { loadFromPanel(false); });
    ui.edit.addEventListener('change', function () { renderFields(); updateLoadState(); });
    ui.regionbar.addEventListener('change', function (e) {
      if (e.target.hasAttribute('data-region-edit') && activeRegion()) setRegionEditing(activeRegion(), e.target.checked);
    });

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
    ui.bRun.addEventListener('click', guard(insertBudgetFromPanel));

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
      if (evt && evt.type === MUTATION) { onSheetMutation(evt); redrawCharts(); }
      // Clicking a cell runs this operation; it is the reliable signal that the selection moved.
      if (evt && evt.id === 'sheet.operation.set-selections') onSelectionMoved();
    }));

    // Ctrl+click (Cmd+click on Mac) on a cell opens it in BC (drill-down); the selection operation that follows tells which cell.
    ui.sheet.addEventListener('pointerdown', function (e) { state.ctrlClick = (e.ctrlKey || e.metaKey) && e.button === 0 ? Date.now() : 0; }, true);

    // Univer listens to window resize; BC resizes the iframe but the panel changes our own width.
    try {
      state.disposables.push(state.api.addEvent(state.api.Event.SelectionChanged, onSelectionMoved));
      state.disposables.push(state.api.addEvent(state.api.Event.ActiveSheetChanged, function () { setTimeout(guard(updateRegionBar), 50); }));
    } catch (e) { /* events optional */ }
    registerMenus();
    registerBCFunctions();
    registerChartComponent();

    var kick = H.debounce(function () { window.dispatchEvent(new Event('resize')); }, 60);
    if (window.ResizeObserver) {
      state.resizeObserver = new ResizeObserver(kick);
      state.resizeObserver.observe(ui.sheet);
    }
    ui.boot.classList.add('is-done');
    setTimeout(function () { if (ui.boot.parentNode) ui.boot.parentNode.removeChild(ui.boot); }, 300);
  }

  /** The selection moved: show the region bar for the cell and let open panels follow it. */
  var onSelectionMoved = H.debounce(function () {
    guard(updateRegionBar)();
    guard(followSelectionInPanels)();
    if (state.ctrlClick && Date.now() - state.ctrlClick < 1500) {
      state.ctrlClick = 0;
      guard(function () { drillAt(selectedCell()); })();
    }
  }, 80);

  function workbook() {
    return state.api && state.api.getActiveWorkbook();
  }

  function newEmptyWorkbook() {
    withoutDirty(function () {
      state.api.createWorkbook({ id: H.uid('wb'), name: 'Workbook', locale: window.AKUniver.LocaleType.EN_US });
      excelLook(workbook().getActiveSheet());
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

  /** Excel-like defaults for a new sheet: Calibri 11 and light gridlines. */
  function excelLook(sheet) {
    if (!sheet) return;
    try { sheet.setDefaultStyle({ ff: 'Calibri', fs: 11 }); } catch (e) { /* optional */ }
    try { sheet.setGridLinesColor('#D4D4D4'); } catch (e) { /* optional */ }
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
      var b = ui.app.querySelector(v === 'data' ? '[data-act=insert-region]' : '[data-act=view-' + v + ']');
      b.setAttribute('aria-expanded', String(open && target === v));
      b.classList.toggle('is-active', open && target === v);
    });
    if (!open) return;
    if (target === 'data') {
      if (!state.regionEdit) state.regionAnchor = selectedCell();
      renderRegionTarget();
      renderSources();
      if (!state.table) {
        if (!state.tables.length) requestTables('');
        setTimeout(function () { ui.tableSearch.focus(); }, 200);
      }
    } else if (target === 'budget') {
      prepareBudgetPanel();
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
    state.fieldOrder = null;
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
      state.fieldState[f.no] = { checked: pick, filter: '', list: true };
    });
    ui.sheetName.value = H.safeSheetName(data.tableCaption);
    // Always shown once a table is picked; greyed out with the reason when this table cannot be edited.
    ui.editWrap.hidden = false;
    ui.edit.checked = false;
    ui.edit.disabled = !data.canEdit;
    ui.editWrap.classList.toggle('is-disabled', !data.canEdit);
    ui.editWhy.hidden = !!data.canEdit;
    ui.editWhy.textContent = data.canEdit ? '' : (data.editReason || 'This table cannot be edited from a workbook.');
    ui.fieldSearch.value = '';
    state.fieldSearch = '';
    // Region properties: start from the region's own fields, filters, rows and editing choice.
    var pre = state.prefill;
    state.fieldOrder = pre && pre.tableNo === data.tableNo ? pre.fields.map(function (f) { return f.no; }) : null;
    if (pre && pre.tableNo === data.tableNo) {
      data.fields.forEach(function (f) { state.fieldState[f.no].checked = false; });
      pre.fields.forEach(function (f) { if (state.fieldState[f.no]) state.fieldState[f.no].checked = true; });
      (pre.filters || []).forEach(function (f) { if (state.fieldState[f.fieldNo]) state.fieldState[f.fieldNo].filter = f.filter; });
      (pre.noList || []).forEach(function (no) { if (state.fieldState[no]) state.fieldState[no].list = false; });
      ui.maxRows.value = pre.maxRows || state.config.defaultMaxRows;
      ui.edit.checked = !!(pre.editable && data.canEdit);
    }
    state.prefill = null;
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
        // Dropdown tick: only for editable Code and Option fields while editing.
        (editing && f.editable && LOOKUP_TYPES[f.type]
          ? '<label class="ak-frow__list" title="Show a dropdown list from BC in this column"><input type="checkbox" data-list' + (fs.list !== false ? ' checked' : '') + '><span>Dropdown</span></label>'
          : '') +
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
    var list = state.table.fields.filter(function (f) { return state.fieldState[f.no].checked || (editing && f.pk); });
    // Regions from a quick start or an existing region keep their own column order (it is also the Validate order).
    var order = state.fieldOrder;
    if (order && order.length) {
      var pos = function (f) { var i = order.indexOf(f.no); return i < 0 ? 10000 + f.no : i; };
      list.sort(function (a, b) { return pos(a) - pos(b); });
    }
    return list;
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

  function loadFromPanel(confirmed) {
    var fields = selectedFields();
    if (!state.table || !fields.length) return;
    if (!confirmed && !(state.regionEdit && state.sources[state.regionEdit])) { confirmNewRegion(); return; }
    var maxRows = parseInt(ui.maxRows.value, 10);
    if (!(maxRows > 0)) maxRows = state.config.defaultMaxRows;
    maxRows = Math.min(maxRows, state.config.maxRowsLimit);
    ui.maxRows.value = maxRows;

    var wb = workbook();
    var existing = state.regionEdit && state.sources[state.regionEdit];
    var dest = (ui.panel.querySelector('input[name=ak-dest]:checked') || {}).value || 'here';
    var sheet, anchor, createdSheetId = null;
    if (existing) {
      sheet = wb.getSheetBySheetId(existing.sheetId);
      anchor = anchorOf(existing);
    } else if (dest === 'new') {
      sheet = wb.insertSheet(uniqueSheetName(ui.sheetName.value || state.table.tableCaption));
      excelLook(sheet);
      anchor = { row: 0, col: 0 };
      createdSheetId = sheet.getSheetId();
    } else {
      var t = state.regionAnchor || selectedCell();
      sheet = wb.getSheetBySheetId(t.sheetId) || wb.getActiveSheet();
      anchor = { row: t.row, col: t.col };
      var clash = regionAt(sheet.getSheetId(), anchor.row, anchor.col);
      if (clash) { H.toast('error', 'This cell is inside region ' + clash.tableCaption + '. Pick an empty cell, or use Properties to change that region.'); return; }
      if (slicerAt(sheet.getSheetId(), anchor.row, anchor.col)) { H.toast('error', cellName(anchor.row, anchor.col) + ' is a period picker. Click the cell where the region should start (for example A4).'); return; }
      if (state.joins[sheet.getSheetId()] || state.recons[sheet.getSheetId()]) { H.toast('error', 'This sheet is a combined or reconciled result. Put the region on another sheet.'); return; }
    }
    wb.setActiveSheet(sheet);

    var source = existing || { id: H.uid('rg'), sheetId: sheet.getSheetId(), anchor: anchor, lastRows: 0, lastCols: 0 };
    var wasEditable = !!source.editable;
    source.tableNo = state.table.tableNo;
    source.tableCaption = state.table.tableCaption;
    source.fields = fields.map(function (f) { return { no: f.no, caption: f.caption, type: f.type, pk: !!f.pk }; });
    source.filters = fields.filter(function (f) { return state.fieldState[f.no].filter; })
      .map(function (f) { return { fieldNo: f.no, filter: state.fieldState[f.no].filter }; });
    source.maxRows = maxRows;
    source.keyCount = state.table.fields.filter(function (f) { return f.pk; }).length;
    source.noList = fields.filter(function (f) { return state.fieldState[f.no].list === false && LOOKUP_TYPES[f.type]; }).map(function (f) { return f.no; });
    source.editable = isEditing();
    if (source.editable && !wasEditable) source.editing = true;
    if (!source.editable) { delete source.orig; delete source.ver; delete source.columns; source.editing = undefined; }
    try { parseRefs(source); } catch (e) { H.toast('error', e.message); return; }

    var finish = function () {
      runSource(source).then(function (res) {
        H.toast(res.truncated ? 'warning' : 'success', res.truncated
          ? 'Loaded the first ' + H.formatInt(res.loaded) + ' of ' + H.formatInt(res.total) + ' rows. Raise Max rows or add filters.'
          : 'Loaded ' + H.formatInt(res.loaded) + ' rows from ' + source.tableCaption + ' at ' + regionLabel(source) + '.');
        state.regionEdit = source.id;
        renderRegionTarget();
        updateRegionBar();
        if (window.innerWidth < 900) togglePanel(false);
      }, function () {
        if (!createdSheetId) return;
        var w = workbook(), sh = w && w.getSheetBySheetId(createdSheetId);
        if (sh && w.getSheets().length > 1) withoutDirty(function () { w.deleteSheet(sh); });
      });
    };
    if (existing) confirmDiscard([existing]).then(function (ok) { if (ok) finish(); });
    else finish();
  }

  /**
   * Before a NEW region is placed, ask once: this stops a second click from adding the same table twice.
   * When the workbook already has a region of the same table, offer to update that one instead.
   */
  function confirmNewRegion() {
    var dest = (ui.panel.querySelector('input[name=ak-dest]:checked') || {}).value || 'here';
    var t = state.regionAnchor || selectedCell();
    var where = dest === 'new' ? 'a new sheet' : sheetName(t.sheetId) + '!' + cellName(t.row, t.col);
    var same = sourceList().filter(function (o) { return o.tableNo === state.table.tableNo && !isBudget(o); });
    // Prefer a region of the same table on the sheet the user is on.
    same.sort(function (x, y) { return (x.sheetId === t.sheetId ? 0 : 1) - (y.sheetId === t.sheetId ? 0 : 1); });
    var buttons = [], msg = 'Insert a new ' + state.table.tableCaption + ' region at ' + where + '?';
    if (same.length) {
      var o = same[0];
      msg += ' This workbook already has ' + (same.length === 1 ? (/^[aeiou]/i.test(state.table.tableCaption) ? 'an ' : 'a ') + state.table.tableCaption + ' region at ' + regionLabel(o) : same.length + ' ' + state.table.tableCaption + ' regions (the first is at ' + regionLabel(o) + ')') +
        '. "Update existing" loads the fields and filters from this panel into that region instead.';
      buttons.push({ label: 'Update existing', value: 'update' });
    }
    buttons.push({ label: 'Insert new region', value: 'insert', primary: true });
    H.choose(msg, buttons, 'Insert region').then(guard(function (v) {
      if (v === 'insert') loadFromPanel(true);
      if (v === 'update') {
        state.regionEdit = same[0].id;
        state.regionAnchor = { sheetId: same[0].sheetId, row: anchorOf(same[0]).row, col: anchorOf(same[0]).col };
        renderRegionTarget();
        loadFromPanel(true);
      }
    }));
  }

  /** Requests data for a source and writes it to its sheet. Resolves with {loaded,total,truncated}. */
  function runSource(source) {
    if (isBudget(source)) return runBudget(source).then(function (r) { return { loaded: r.accounts, total: r.accounts, truncated: false }; });
    setBusy(true);
    var requestId = 'r' + (++state.requestSeq);
    return new Promise(function (resolve, reject) {
      state.pending[requestId] = { source: source, columns: null, rows: [], resolve: resolve, reject: reject };
      setProgress(-1);
      var filters;
      try { filters = resolveFilters(source); } catch (e) { failRequest(requestId, e.message); return; }
      H.invoke('OnDataRequested', [JSON.stringify({
        requestId: requestId,
        tableNo: source.tableNo,
        fields: source.fields.map(function (f) { return f.no; }),
        filters: filters,
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
      // A region never grows into the next region below it.
      var room = roomBelow(p.source);
      if (values.length + 1 > room) {
        H.toast('warning', 'Only ' + H.formatInt(Math.max(room - 1, 0)) + ' rows fit above the next region. Move that region down or add filters.');
        values = values.slice(0, Math.max(room - 1, 0));
        if (p.versions) p.versions = p.versions.slice(0, values.length);
      }
      applying(function () {
        var sheet = writeTable(p.source, header, types, values);
        styleRegion(p.source, sheet, values.length + 1, header.length);
      });
      if (p.editable) setupEditable(p.source, columns, values, p.versions, p.canInsert);
      p.source.filterText = p.filterText || '';
      p.source.loadedAt = new Date().toISOString();
      p.source.total = chunk.total;
      p.source.truncated = chunk.truncated;
      state.sources[p.source.id] = p.source;
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
      excelLook(sheet);
      target.sheetId = sheet.getSheetId();
    }
    var a = target.anchor || { row: 0, col: 0 };
    var nCols = header.length, nRows = rows.length + 1;

    if (target.lastRows && target.lastCols) sheet.getRange(a.row, a.col, target.lastRows, target.lastCols).clearContent();
    if (target._end && target._end > a.row + (target.lastRows || 0) - 1) {
      // Rows saved below an editable region since the last load belong to the region too.
      var extra = target._end - (a.row + (target.lastRows || 1) - 1);
      if (extra > 0) sheet.getRange(a.row + target.lastRows, a.col, extra, target.lastCols || nCols).clearContent();
    }
    if (sheet.getMaxRows() < a.row + nRows + 50) sheet.insertRowsAfter(sheet.getMaxRows() - 1, a.row + nRows + 50 - sheet.getMaxRows());
    if (sheet.getMaxColumns() < a.col + nCols + 5) sheet.insertColumnsAfter(sheet.getMaxColumns() - 1, a.col + nCols + 5 - sheet.getMaxColumns());

    var values = new Array(nRows);
    values[0] = header.slice();
    for (var r = 0; r < rows.length; r++) values[r + 1] = rows[r];
    sheet.getRange(a.row, a.col, nRows, nCols).setValues(values);

    var head = sheet.getRange(a.row, a.col, 1, nCols);
    head.setFontWeight('bold');
    head.setBackgroundColor('#4472C4');
    head.setFontColor('#FFFFFF');
    if (a.row === 0) sheet.setFrozenRows(1);

    if (rows.length) {
      types.forEach(function (t, c) {
        var fmt = NUMBER_FORMATS[t];
        if (fmt) { try { sheet.getRange(a.row + 1, a.col + c, rows.length, 1).setNumberFormat(fmt); } catch (e) { /* numfmt optional */ } }
        if (NUMERIC_TYPES[t]) { try { sheet.getRange(a.row, a.col + c, 1, 1).setHorizontalAlignment('right'); } catch (e) { /* optional */ } }
      });
    }
    autoWidth(sheet, header, types, rows, a.col);

    // Univer allows one auto-filter per sheet: give it to the region when it is alone on its sheet.
    var alone = !target.id || regionsOn(sheet.getSheetId()).filter(function (o) { return o.id !== target.id; }).length === 0;
    if (alone) {
      try {
        var existing = sheet.getFilter && sheet.getFilter();
        if (existing) existing.remove();
        sheet.getRange(a.row, a.col, Math.max(nRows, 2), nCols).createFilter();
      } catch (e) { /* filter preset optional */ }
    }
    try { sheet.getRange(a.row, a.col, 1, 1).activate(); } catch (e) { /* keep selection */ }

    target.lastRows = nRows;
    target.lastCols = nCols;
    target.rowCount = rows.length;
    target._end = a.row + rows.length;
    return sheet;
  }

  function autoWidth(sheet, header, types, rows, startCol) {
    startCol = startCol || 0;
    var sample = Math.min(rows.length, 200);
    header.forEach(function (cap, c) {
      var t = types[c], len = String(cap).length + 2;
      for (var r = 0; r < sample; r++) {
        var v = rows[r][c];
        var l = t === 'Date' ? 11 : t === 'DateTime' ? 17 : NUMERIC_TYPES[t] ? String(Math.round(+v || 0)).length + 6 : String(v == null ? '' : v).length;
        if (l > len) len = l;
      }
      sheet.setColumnWidths(startCol + c, 1, Math.max(64, Math.min(320, len * 7.5 + 18)));
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
  /** BC regions (keyed by region id; several may share a sheet). Regions whose sheet was deleted are dropped. */
  function sourceList() {
    var wb = workbook();
    if (!wb) return [];
    Object.keys(state.sources).forEach(function (id) { if (!wb.getSheetBySheetId(state.sources[id].sheetId)) delete state.sources[id]; });
    return Object.keys(state.sources).map(function (id) { return state.sources[id]; });
  }
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
      if (refreshBCFunctions()) { H.toast('success', 'BC functions recalculated.'); return Promise.resolve(); }
      H.toast('info', 'Nothing to refresh yet. Use Insert region to load a table, or BC functions such as =BCGLSUM("40000..49999", "This Year").');
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
      refreshBCFunctions();
      H.toast('success', 'Refreshed ' + done + (done === 1 ? ' sheet' : ' sheets') + ' (' + H.formatInt(rows) + ' rows).');
    }, function (e) { if (e && !e.shown) reportError(e); });
  }

  function refreshOne(sheetId) {
    if (state.busy) return;
    // id is a region id (BC regions) or a sheet id (combined / reconciled sheets).
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
    var region = state.sources[sheetId];
    if (region && region.bandRuleId) {
      try { workbook().getSheetBySheetId(region.sheetId).deleteConditionalFormattingRule(region.bandRuleId); } catch (e) { /* optional */ }
    }
    delete state.sources[sheetId];
    updateWriteButton(); updateRegionBar();
    delete state.joins[sheetId];
    delete state.recons[sheetId];
    renderSources(); renderBarSources(); setDirty(true);
    H.toast('info', 'Sheet unlinked. Its data stays but will not refresh.');
  }

  /** Remove a region: clear its cells (data, format, filter, dropdowns) or keep the data and only unlink it. */
  function askRemoveRegion(s) {
    var g = regionRange(s, 'all'), where = sheetName(s.sheetId) + '!' + cellName(g.r, g.c) + ':' + cellName(g.r + g.rows - 1, g.c + g.cols - 1);
    var pending = s.changeCount ? ' It has ' + s.changeCount + ' change(s) not saved to Business Central; they will be lost.' : '';
    H.choose('Remove the ' + s.tableCaption + ' region at ' + where + '?' + pending +
      ' "Remove region" clears these cells. "Keep data" leaves the values on the sheet as plain cells that no longer refresh. Nothing is deleted in Business Central.',
      [{ label: 'Keep data, unlink', value: 'unlink' }, { label: 'Remove region', value: 'remove', primary: true }], 'Remove region')
      .then(guard(function (v) {
        if (v === 'unlink') unlinkSource(s.id);
        if (v === 'remove') removeRegion(s);
      }));
  }

  function removeRegion(s) {
    var sheet = workbook().getSheetBySheetId(s.sheetId);
    var g = regionRange(s, 'all');
    if (sheet) {
      withoutDirty(function () {
        state.applying++;
        try {
          // The editable line under the rows and Univer's border edges one row/column outside belong to the region too.
          var area = sheet.getRange(g.r, g.c, g.rows + (s.editable ? 2 : 1), g.cols + 1);
          try { var f = sheet.getFilter && sheet.getFilter(); var fr = f && f.getRange && f.getRange();
            if (f && (!fr || (fr.getRow() === g.r && fr.getColumn() === g.c))) f.remove(); } catch (e) { /* no filter */ }
          // Dropdowns reach below the rows (room for new lines), so clear them over that height too.
          try { sheet.getRange(g.r, g.c, Math.max(g.rows + 2, (s.columns ? lookupRows(s) : 0) + 1), g.cols + 1).setDataValidation(null); } catch (e) { /* no dropdowns */ }
          area.clear();
        } finally { state.applying--; }
      });
    }
    if (s.bandRuleId && sheet) { try { sheet.deleteConditionalFormattingRule(s.bandRuleId); } catch (e) { /* optional */ } }
    delete state.sources[s.id];
    delete state.joins[s.id];
    delete state.recons[s.id];
    updateWriteButton(); updateRegionBar(); renderSources(); renderBarSources(); setDirty(true);
    H.toast('success', s.tableCaption + ' region removed from ' + sheetName(s.sheetId) + '. Use Undo (Ctrl+Z) to bring the cells back as plain data.');
  }

  function activateSheet(id) {
    var region = state.sources[id];
    var wb = workbook(), sheet = wb && wb.getSheetBySheetId(region ? region.sheetId : id);
    if (!sheet) return;
    wb.setActiveSheet(sheet);
    if (region) { try { sheet.getRange(anchorOf(region).row, anchorOf(region).col, 1, 1).activate(); } catch (e) { /* optional */ } }
  }

  function sheetName(sheetId) {
    var s = workbook().getSheetBySheetId(sheetId);
    return s ? s.getSheetName() : '?';
  }

  function renderSources() {
    var sources = sourceList(), joins = joinList(), recons = reconList();
    if (!sources.length && !joins.length && !recons.length) {
      ui.sourceList.innerHTML = '<li class="ak-list__empty">Regions you insert from Business Central appear here and refresh with one click.</li>';
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
      return item(s.id, regionLabel(s) + ' \u00b7 ' + s.tableCaption,
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
  // Regions (Stage 1b): a Business Central table placed at any cell, many per sheet
  //
  // A region is a "source" with an anchor cell. Its header is at the anchor, its rows below it.
  // Regions move when rows or columns are inserted or deleted above / left of them.
  // Filters may point to a cell, e.g. [B1] or [Sheet 1!B1]; changing that cell refreshes the region.
  // =====================================================================================
  var XL = { head: '#4472C4', headText: '#FFFFFF', lockedHead: '#8EA9DB', band: '#D9E1F2', border: '#8EA9DB', statusHead: '#305496' };

  function anchorOf(s) { return s.anchor || { row: 0, col: 0 }; }
  function regionWidth(s) { return s.lastCols || ((s.fields || []).length + (s.editable ? 1 : 0)); }
  function regionEndRow(s) { return Math.max(anchorOf(s).row + Math.max(s.lastRows || 1, 1) - 1, s._end || 0); }
  function regionsOn(sheetId) {
    return sourceList().filter(function (s) { return s.sheetId === sheetId; }).sort(function (a, b) {
      return anchorOf(a).row - anchorOf(b).row || anchorOf(a).col - anchorOf(b).col;
    });
  }
  /** The first region on a sheet. Combine and Reconcile read this one when they use a sheet. */
  function srcOf(sheetId) { return regionsOn(sheetId)[0] || null; }
  function regionAt(sheetId, row, col) {
    return regionsOn(sheetId).filter(function (s) {
      var a = anchorOf(s);
      return row >= a.row && row <= regionEndRow(s) + (s.editable ? 1 : 0) && col >= a.col && col < a.col + regionWidth(s);
    })[0] || null;
  }
  function cellName(row, col) { return H.colLetter(col) + (row + 1); }
  function regionLabel(s) { var a = anchorOf(s); return sheetName(s.sheetId) + '!' + cellName(a.row, a.col); }

  /** Rows available before the next region below (that overlaps in columns), so a load never overwrites it. */
  function roomBelow(s) {
    var a = anchorOf(s), w = regionWidth(s) || 1, limit = Infinity;
    regionsOn(s.sheetId).forEach(function (o) {
      if (o === s) return;
      var b = anchorOf(o), ow = regionWidth(o) || 1;
      var overlap = b.col < a.col + w && a.col < b.col + ow;
      if (overlap && b.row > a.row) limit = Math.min(limit, b.row - a.row - 1); // keep one empty row
    });
    return limit;
  }

  function selectedCell() {
    var wb = workbook(), sheet = wb && wb.getActiveSheet(), sel = sheet && sheet.getSelection && sheet.getSelection();
    var range = sel && sel.getActiveRange && sel.getActiveRange();
    if (!range) return { sheetId: sheet ? sheet.getSheetId() : '', row: 0, col: 0 };
    return { sheetId: sheet.getSheetId(), row: range.getRow(), col: range.getColumn() };
  }

  // ---------------------------------------------------------------- cell-driven filters
  var REF_RE = /\[(?:'?([^\]!']+)'?!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})\]/g;

  function parseRefs(source) {
    var refs = [];
    (source.filters || []).forEach(function (f) {
      String(f.filter).replace(REF_RE, function (m, sheet, col, row) {
        var target = sheet ? workbook().getSheetByName(sheet) : workbook().getSheetBySheetId(source.sheetId);
        if (!target) throw new Error('The filter ' + m + ' points to a sheet that does not exist.');
        refs.push({ sheetId: target.getSheetId(), row: +row - 1, col: colIndex(col) });
        return m;
      });
    });
    return refs;
  }

  function colIndex(letters) {
    var c = 0;
    letters = letters.toUpperCase();
    for (var i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
    return c - 1;
  }

  /** Filters with every [cell] replaced by that cell's displayed value (so dates follow the cell's format). */
  function resolveFilters(source) {
    return (source.filters || []).map(function (f) {
      var text = String(f.filter).replace(REF_RE, function (m, sheet) {
        var target = sheet ? workbook().getSheetByName(sheet) : workbook().getSheetBySheetId(source.sheetId);
        if (!target) throw new Error('The filter ' + m + ' points to a sheet that does not exist.');
        var ref = /\$?([A-Za-z]{1,3})\$?(\d{1,7})\]$/.exec(m);
        var row = +ref[2] - 1, col = colIndex(ref[1]);
        var cell = target.getRange(row, col, 1, 1);
        var shown = cell.getDisplayValue ? cell.getDisplayValue() : cell.getValue();
        var picked = slicerFilter(target.getSheetId(), row, col, shown);
        if (picked !== null) return picked;
        return shown == null ? '' : String(shown).trim();
      });
      return { fieldNo: f.fieldNo, filter: text };
    }).filter(function (f) { return f.filter !== ''; });
  }

  function onRefCellsChanged(sheetId, cells) {
    sourceList().forEach(function (s) {
      var refs;
      try { refs = parseRefs(s); } catch (e) { return; }
      var hit = refs.some(function (r) { return r.sheetId === sheetId && cells.some(function (c) { return c.r === r.row && c.c === r.col; }); });
      if (!hit) return;
      clearTimeout(s._refTimer);
      s._refTimer = setTimeout(function tick() {
        if (state.busy) { s._refTimer = setTimeout(tick, 800); return; }
        if ((isTracking(s) && diffSource(s).length) || (budgetTracking(s) && budgetDiff(s).length)) {
          H.toast('warning', regionLabel(s) + ' was not refreshed: it has changes that are not saved to BC yet.');
          return;
        }
        runSource(s).then(function (r) {
          updateRegionBar();
          H.toast('info', s.tableCaption + ' refreshed for the new filter value (' + H.formatInt(r.loaded) + ' rows).');
        }, function () { });
      }, 700);
    });
  }

  /** Inserting or deleting rows/columns above or left of a region moves it, like an Excel table. */
  function shiftRegions(id, params) {
    var range = params.range || (params.ranges && params.ranges[0]);
    if (!range || !params.subUnitId) return;
    var isRow = /row/.test(id), isInsert = /insert/.test(id);
    var start = isRow ? range.startRow : range.startColumn, end = isRow ? range.endRow : range.endColumn;
    var count = end - start + 1;
    regionsOn(params.subUnitId).forEach(function (s) {
      var a = anchorOf(s), pos = isRow ? a.row : a.col;
      if (isInsert) {
        if (start <= pos) pos += count;
      } else if (end < pos) {
        pos -= count;
      } else if (start <= pos) {
        delete state.sources[s.id];
        H.toast('warning', 'The header of region ' + s.tableCaption + ' was deleted, so the region was removed. Its data stays as plain cells.');
        return;
      }
      s.anchor = isRow ? { row: pos, col: a.col } : { row: a.row, col: pos };
    });
    Object.keys(state.slicers).forEach(function (k) {
      var p = state.slicers[k];
      if (p.sheetId !== params.subUnitId) return;
      var pos = isRow ? p.row : p.col;
      if (isInsert) { if (start <= pos) pos += count; }
      else if (end < pos) pos -= count;
      else if (start <= pos) { delete state.slicers[k]; return; }
      if (isRow) p.row = pos; else p.col = pos;
    });
    renderSources(); renderBarSources(); updateRegionBar();
  }

  // =====================================================================================
  // Editable regions (write-back to Business Central)
  //
  // A region loaded with "Allow editing" keeps a copy of every loaded row by primary key (orig) and
  // the record's SystemModifiedAt (ver). On each edit the region is compared with that copy:
  //   - an existing key with different values in editable columns -> 'modify'
  //   - a row typed directly below the data whose key is not in the copy -> 'insert'
  // Rows below that contain a formula (totals) end the region and are never sent to BC.
  // Key, calculated and system columns are locked: an edit there is put back at once.
  // The region's Editing switch turns all of this off, so the user can just view and use formulas.
  // =====================================================================================
  var EDIT_COLORS = { locked: '#F2F2F2', lockedText: '#595959', changed: '#FFF4CE', added: '#E7F6EC' };
  var STATUS_CHANGED = 'Changed', STATUS_NEW = 'New';

  function isTracking(s) { return !!(s && s.editable && s.editing !== false && s.columns); }
  function editableList() { return sourceList().filter(isTracking); }

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
      return String(v).replace(/,/g, '').trim();
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
    var a = anchorOf(source);
    source.columns = columns.map(function (c) { return { no: c.no, caption: c.caption, type: c.type, pk: !!c.pk, editable: !!c.editable }; });
    source.keyIdx = [];
    source.columns.forEach(function (c, i) { if (c.pk) source.keyIdx.push(i); });
    source.dataCols = source.columns.length;
    source.canInsert = !!canInsert;
    if (source.editing === undefined) source.editing = true;
    source.orig = {};
    source.ver = {};
    source.posKeys = [];
    rows.forEach(function (r, i) {
      var k = keyOf(source, r);
      source.orig[k] = r.slice(0, source.dataCols);
      source.ver[k] = (versions && versions[i]) || '';
      source.posKeys[a.row + 1 + i] = k;
    });
    source._end = a.row + rows.length;
    source.changeCount = 0;
    source.marks = [];
    paintBase(source, true);
    var sheetNow = workbook().getSheetBySheetId(source.sheetId);
    if (sheetNow) growRegionBox(source, sheetNow);
    updateWriteButton();
    requestLookups(source, true);
  }

  /** Base look of an editable region: locked columns grey, editable white, status column header dark. */
  function paintBase(source, includeHeader) {
    var sheet = workbook().getSheetBySheetId(source.sheetId);
    if (!sheet || !source.columns) return;
    var a = anchorOf(source), n = source.dataCols;
    // The loaded rows plus a few empty rows for new records (when inserting is allowed); old paint is cleared.
    var rows = boxRows(source) - 1;
    var old = source.paintedRows || 0;
    if (old > rows) applying(function () { sheet.getRange(a.row + 1 + rows, a.col, old - rows, n + 1).setBackgroundColor('#FFFFFF'); });
    applying(function () {
      source.columns.forEach(function (c, i) {
        var range = sheet.getRange(a.row + 1, a.col + i, rows, 1);
        range.setBackgroundColor(c.editable ? '#FFFFFF' : EDIT_COLORS.locked);
        range.setFontColor(c.editable ? '#000000' : EDIT_COLORS.lockedText);
        // Number formats also cover the empty rows below, so new rows look like the loaded ones.
        if (NUMBER_FORMATS[c.type]) { try { range.setNumberFormat(NUMBER_FORMATS[c.type]); } catch (e) { /* optional */ } }
        if (includeHeader) sheet.getRange(a.row, a.col + i, 1, 1).setBackgroundColor(c.editable ? XL.head : XL.lockedHead);
      });
      sheet.getRange(a.row + 1, a.col + n, rows, 1).setBackgroundColor('#FFFFFF');
      if (includeHeader) {
        var sh = sheet.getRange(a.row, a.col + n, 1, 1);
        sh.setBackgroundColor(XL.statusHead);
        sh.setFontColor('#FFFFFF');
        sh.setFontWeight('bold');
        sheet.setColumnWidths(a.col + n, 1, 300);
      }
    });
    source.paintedRows = rows;
  }

  /** Runs fn while edit tracking ignores the mutations it causes (styles, status text, saved values). */
  function applying(fn) {
    state.applying = (state.applying || 0) + 1;
    try { return fn(); } finally { state.applying--; }
  }

  /** Univer mutation hook: moves regions, follows filter cells, protects locked cells, schedules change checks. */
  function onSheetMutation(evt) {
    if (state.applying || !evt || !evt.params) return;
    // Formula results are written back as mutations too; they are not user edits.
    if (evt.options && (evt.options.fromFormula || evt.options.applyFormulaCalculationResult)) return;
    var id = evt.id, params = evt.params, sheetId = params.subUnitId;
    if (!sheetId) return;
    if (/sheet\.mutation\.(insert-row|remove-rows|insert-col|remove-col)$/.test(id)) {
      shiftRegions(id, params);
      regionsOn(sheetId).filter(isTracking).forEach(function (s) { setTimeout(function () { paintBase(s, false); diffSource(s); }, 0); });
      return;
    }
    if (id === 'sheet.mutation.set-range-values' && params.cellValue) {
      var touched = [], perRegion = {}, budgetCells = {};
      Object.keys(params.cellValue).forEach(function (rk) {
        var r = +rk, rowCells = params.cellValue[rk];
        Object.keys(rowCells).forEach(function (ck) {
          var c = +ck, cell = rowCells[ck];
          if (!cell || !('v' in cell || 'f' in cell || 'p' in cell || 'si' in cell)) return;
          touched.push({ r: r, c: c });
          var s = regionAt(sheetId, r, c);
          if (isBudget(s)) { (budgetCells[s.id] = budgetCells[s.id] || { s: s, cells: [] }).cells.push({ r: r, c: c }); return; }
          if (!isTracking(s)) return;
          (perRegion[s.id] = perRegion[s.id] || { s: s, cells: [] }).cells.push({ r: r, c: c });
        });
      });
      if (touched.length) { onSlicerCells(sheetId, touched); onRefCellsChanged(sheetId, touched); }
      Object.keys(budgetCells).forEach(function (k) { onBudgetCells(budgetCells[k].s, budgetCells[k].cells); });
      Object.keys(perRegion).forEach(function (k) {
        var s = perRegion[k].s, a = anchorOf(s), reverts = [];
        perRegion[k].cells.forEach(function (x) {
          var ci = x.c - a.col, key = s.posKeys[x.r];
          if (x.r <= a.row || ci >= s.dataCols) return;
          if (key != null && s.orig[key] && !s.columns[ci].editable) reverts.push({ r: x.r, c: x.c, ci: ci, v: s.orig[key][ci] });
          if (s.lookupTypeCols && s.lookupTypeCols.indexOf(ci) >= 0) scheduleLookups(s);
        });
        if (reverts.length) setTimeout(function () { revertCells(s, reverts); }, 0);
        scheduleDiff(s);
      });
      return;
    }
    if (/reorder-range|move-rows|move-range/.test(id)) {
      // Rows moved (sort): marks moved with them, so repaint and check again.
      regionsOn(sheetId).filter(isTracking).forEach(function (s) { setTimeout(function () { paintBase(s, false); diffSource(s); }, 0); scheduleLookups(s); });
    }
  }

  function revertCells(source, reverts) {
    var sheet = workbook().getSheetBySheetId(source.sheetId);
    if (!sheet) return;
    applying(function () {
      reverts.forEach(function (x) { sheet.getRange(x.r, x.c, 1, 1).setValue(x.v == null ? '' : x.v); });
    });
    var cols = reverts.map(function (x) { return source.columns[x.ci].caption; })
      .filter(function (v, i, arr) { return arr.indexOf(v) === i; }).slice(0, 3).join(', ');
    H.toast('warning', cols + (reverts.length > 1 ? ' are' : ' is') + ' read-only here (key, calculated or system field). The value was put back.');
    scheduleDiff(source);
  }

  function scheduleDiff(source) {
    clearTimeout(source._diffTimer);
    source._diffTimer = setTimeout(function () { diffSource(source); }, 350);
  }

  /**
   * Compares the region with the loaded rows and marks the differences. Returns the list of changes.
   * Below the loaded rows the region continues while rows have typed values; it stops at the first
   * empty row, a row with a formula (a totals row) or the next region.
   */
  function diffSource(source) {
    var sheet = workbook() && workbook().getSheetBySheetId(source.sheetId);
    if (!sheet || !isTracking(source)) { if (source) { source.changeCount = 0; updateWriteButton(); } return []; }
    var a = anchorOf(source), n = source.dataCols;
    var loadedEnd = a.row + (source.rowCount || 0);
    var lastAllowed = Math.min(sheet.getLastRow(), a.row + roomBelow(source));
    var count = Math.max(0, lastAllowed - a.row);
    var raw = count ? sheet.getRange(a.row + 1, a.col, count, n + 1).getRawValues() : [];
    var formulas = [];
    try { formulas = count ? sheet.getRange(a.row + 1, a.col, count, n).getFormulas() : []; } catch (e) { formulas = []; }
    var changes = [], seen = {}, posKeys = [], end = a.row;
    for (var i = 0; i < raw.length; i++) {
      var r = a.row + 1 + i, full = raw[i], row = full.slice(0, n), status = full[n];
      var hasFormula = (formulas[i] || []).some(function (f) { return !!f; });
      var empty = row.every(function (v) { return v == null || v === ''; });
      if (r > loadedEnd && (empty || hasFormula)) {
        if (empty && (status === STATUS_NEW || status === STATUS_CHANGED)) changes.push({ r: r, op: 'clear' });
        break;
      }
      if (empty) {
        if (status === STATUS_NEW || status === STATUS_CHANGED) changes.push({ r: r, op: 'clear' });
        continue;
      }
      end = r;
      var k = keyOf(source, row), orig = source.orig[k];
      if (orig && !seen[k]) {
        seen[k] = true;
        posKeys[r] = k;
        var cols = [];
        source.columns.forEach(function (c, ci) { if (c.editable && !sameValue(c.type, row[ci], orig[ci])) cols.push(ci); });
        if (cols.length) changes.push({ r: r, op: 'modify', key: k, row: row, cols: cols, status: status });
        else if (status === STATUS_CHANGED) changes.push({ r: r, op: 'clear' });
      } else if (r > loadedEnd && !hasFormula) {
        var filled = [];
        source.columns.forEach(function (c, ci) { if ((c.editable || c.pk) && row[ci] !== '' && row[ci] != null) filled.push(ci); });
        changes.push(source.canInsert
          ? { r: r, op: 'insert', row: row, cols: filled, status: status }
          : { r: r, op: 'blocked', status: status });
      }
    }
    source.posKeys = posKeys;
    source._end = end;
    growRegionBox(source, sheet);
    markChanges(source, sheet, changes);
    var real = changes.filter(function (x) { return x.op === 'modify' || x.op === 'insert'; });
    // The region grew (a new row typed) or shrank: move the dropdowns with it.
    if (source._lkRows && lk.cache[source.tableNo] && lookupRows(source) !== source._lkRows) scheduleLookups(source);
    source.changeCount = real.length;
    updateWriteButton();
    return real;
  }

  /**
   * The blue box is the region. When a new row is typed on the line under an editable region, the box
   * (border) grows to take it in, and shrinks back when that row is cleared. It never reaches into another
   * region below, because diffSource only reads rows up to roomBelow().
   */
  /**
   * Rows of the blue box: header + data rows + (editable and inserting allowed) one empty line to type a new
   * row. Typing on that line makes it a data row, so the box grows by one and a new empty line appears.
   * The box never reaches the row just above another region.
   */
  function boxRows(source) {
    var a = anchorOf(source);
    var rows = 1 + Math.max(regionEndRow(source) - a.row, 0) + (source.editable && source.canInsert ? 1 : 0);
    return Math.max(1, Math.min(rows, roomBelow(source) + 1));
  }

  function growRegionBox(source, sheet) {
    if (!source.styledArea || !source.editable || !source.columns) return;
    var want = boxRows(source);
    if (want === source.styledArea.rows) return;
    applying(function () {
      styleRegion(source, sheet, want, source.styledArea.cols);
      // The sheet's filter buttons belong to this region: keep the filter range equal to the box.
      try {
        var a = anchorOf(source), f = sheet.getFilter && sheet.getFilter(), fr = f && f.getRange();
        if (f && fr && fr.getRow() === a.row && fr.getColumn() === a.col && fr.getHeight() !== Math.max(want, 2)) {
          var cols = fr.getWidth();
          f.remove();
          sheet.getRange(a.row, a.col, Math.max(want, 2), cols).createFilter();
        }
      } catch (e) { /* filter optional */ }
    });
    paintBase(source, false);
  }

  function markChanges(source, sheet, changes) {
    var a = anchorOf(source), n = source.dataCols;
    applying(function () {
      (source.marks || []).forEach(function (m) {
        var c = source.columns[m.ci];
        if (c) sheet.getRange(m.r, a.col + m.ci, 1, 1).setBackgroundColor(c.editable ? '#FFFFFF' : EDIT_COLORS.locked);
      });
      source.marks = [];
      changes.forEach(function (x) {
        var cell = sheet.getRange(x.r, a.col + n, 1, 1);
        if (x.op === 'clear') { cell.setValue(''); return; }
        if (x.op === 'blocked') {
          cell.setValue('Not saved: new records are not allowed for this table');
          cell.setFontColor('#A4262C');
          return;
        }
        if (source.marks.length > 3000) return;
        var color = x.op === 'insert' ? EDIT_COLORS.added : EDIT_COLORS.changed;
        var cols = x.op === 'insert' ? source.columns.map(function (c, i) { return i; }) : x.cols;
        cols.forEach(function (ci) {
          sheet.getRange(x.r, a.col + ci, 1, 1).setBackgroundColor(color);
          source.marks.push({ r: x.r, ci: ci });
        });
        if (!x.status || x.status === STATUS_CHANGED || x.status === STATUS_NEW || /^Saved/.test(String(x.status))) {
          cell.setValue(x.op === 'insert' ? STATUS_NEW : STATUS_CHANGED);
          cell.setFontColor(x.op === 'insert' ? '#0E700E' : '#835B00');
        }
      });
    });
  }

  function updateWriteButton() {
    if (!ui.write) return;
    var list = editableList().concat(budgetList());
    var total = list.reduce(function (t, s) { return t + (s.changeCount || 0); }, 0);
    ui.write.hidden = list.length === 0;
    ui.write.disabled = total === 0 || state.busy;
    ui.writeCount.textContent = total ? String(total) : '';
    ui.writeCount.hidden = !total;
  }

  /** Region Editing switch: off = view and formulas only (nothing tracked or sent); on = changes are tracked again. */
  function setRegionEditing(source, on) {
    if (!source || !source.editable) return;
    if (isBudget(source)) {
      var bs = workbook().getSheetBySheetId(source.sheetId);
      if (!on && bs) {
        applying(function () { (source.marks || []).forEach(function (mk) { bs.getRange(mk.r, mk.c, 1, 1).setBackgroundColor('#FFFFFF'); }); });
        source.marks = []; source.changeCount = 0;
      }
      source.editing = on;
      if (on) budgetDiff(source);
      updateWriteButton(); updateRegionBar(); setDirty(true);
      H.toast('info', on ? 'Editing is on for ' + source.tableCaption + '.' : 'Editing is off for ' + source.tableCaption + '. Nothing is sent to BC.');
      return;
    }
    source.editing = on;
    var sheet = workbook().getSheetBySheetId(source.sheetId);
    if (!on && sheet) {
      // Hide the marks; the copy of the loaded rows is kept, so turning it on again shows any differences.
      applying(function () {
        (source.marks || []).forEach(function (m) {
          var c = source.columns[m.ci];
          if (c) sheet.getRange(m.r, anchorOf(source).col + m.ci, 1, 1).setBackgroundColor(c.editable ? '#FFFFFF' : EDIT_COLORS.locked);
        });
      });
      source.marks = [];
      source.changeCount = 0;
      clearStatusLabels(source, sheet);
    }
    if (on) diffSource(source);
    updateWriteButton(); updateRegionBar(); renderSources(); setDirty(true);
    H.toast('info', on ? 'Editing is on for ' + source.tableCaption + '. Changes are marked and can be saved to BC.'
      : 'Editing is off for ' + source.tableCaption + '. Use the data and formulas freely; nothing is sent to BC.');
  }

  /** Removes the Changed / New labels from the BC Status column (other statuses stay). */
  function clearStatusLabels(source, sheet) {
    var a = anchorOf(source), n = source.dataCols, rows = Math.max(regionEndRow(source) - a.row, 0) + 3;
    if (!rows) return;
    var st = sheet.getRange(a.row + 1, a.col + n, rows, 1).getRawValues();
    applying(function () {
      st.forEach(function (v, i) {
        if (v[0] === STATUS_CHANGED || v[0] === STATUS_NEW) sheet.getRange(a.row + 1 + i, a.col + n, 1, 1).setValue('');
      });
    });
  }

  /** Toolbar "Save to BC": writes every editable region that has changes, one after the other. */
  function saveToBC() {
    if (state.busy) return;
    var work = editableList().map(function (s) { return { source: s, changes: diffSource(s) }; })
      .concat(budgetList().map(function (s) { return { source: s, changes: budgetDiff(s), budget: true }; }))
      .filter(function (w) { return w.changes.length; });
    if (!work.length) { H.toast('info', 'No changes to save to Business Central.'); return; }
    var nMod = 0, nNew = 0, nCells = 0;
    work.forEach(function (w) { w.changes.forEach(function (x) { if (w.budget) nCells++; else if (x.op === 'insert') nNew++; else nMod++; }); });
    var parts = [];
    if (nMod) parts.push(nMod + (nMod === 1 ? ' changed row' : ' changed rows'));
    if (nNew) parts.push(nNew + (nNew === 1 ? ' new row' : ' new rows'));
    if (nCells) parts.push(nCells + (nCells === 1 ? ' budget amount' : ' budget amounts'));
    H.choose('Write ' + parts.join(' and ') + ' to Business Central? Each field is checked by BC like on the normal page. Rows with an error are not saved.',
      [{ label: 'Save to BC', value: true, primary: true }], 'Save to Business Central')
      .then(function (ok) {
        if (!ok) return;
        setBusy(true);
        setProgress(-1);
        var totals = { saved: 0, failed: 0 };
        var chain = Promise.resolve();
        work.forEach(function (w) {
          chain = chain.then(function () { return w.budget ? writeBudget(w.source, w.changes) : writeSource(w.source, w.changes); })
            .then(function (r) { totals.saved += r.saved; totals.failed += r.failed; });
        });
        return chain.then(function () {
          setBusy(false);
          var unit = nCells && !nMod && !nNew ? ['amount', 'amounts'] : ['row', 'rows'];
          if (unit[0] === 'amount') {
            if (totals.failed) H.toast('warning', 'Saved ' + totals.saved + ' budget ' + (totals.saved === 1 ? unit[0] : unit[1]) + '. ' + totals.failed + ' could not be saved: see the BC Status column.');
            else H.toast('success', 'Saved ' + totals.saved + ' budget ' + (totals.saved === 1 ? unit[0] : unit[1]) + ' to Business Central.');
            return;
          }
          if (totals.failed) H.toast('warning', 'Saved ' + totals.saved + (totals.saved === 1 ? ' row' : ' rows') + '. ' + totals.failed + (totals.failed === 1 ? ' row has' : ' rows have') + ' an error: see the BC Status column.');
          else H.toast('success', 'Saved ' + totals.saved + (totals.saved === 1 ? ' row' : ' rows') + ' to Business Central.');
        }, function (e) { setBusy(false); reportError(e); });
      });
  }

  function writeSource(source, changes) {
    var limit = 1000;
    if (changes.length > limit) {
      H.toast('warning', 'Saving the first ' + limit + ' rows of ' + regionLabel(source) + '. Click Save to BC again for the rest.');
      changes = changes.slice(0, limit);
    }
    var payload = {
      requestId: 'w' + (++state.requestSeq),
      sheetId: source.sheetId,
      regionId: source.id,
      tableNo: source.tableNo,
      fields: source.columns.map(function (c) { return c.no; }),
      keyFields: source.keyIdx.map(function (i) { return source.columns[i].no; }),
      changes: changes.map(function (x) {
        return {
          row: x.r,
          op: x.op,
          key: source.keyIdx.map(function (i) {
            var v = toBC(source.columns[i].type, x.row[i]);
            return v === '' && x.op === 'insert' ? singleFilterValue(source, source.columns[i].no) : v;
          }),
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

  /**
   * New rows in line regions: a key cell left empty takes the region's filter value when that filter is one
   * plain value (journal template and batch, document type and no.). Line No. stays empty: AL gives the next one.
   */
  function singleFilterValue(source, fieldNo) {
    var f;
    try { f = resolveFilters(source).filter(function (x) { return x.fieldNo === fieldNo; })[0]; } catch (e) { return ''; }
    if (!f) return '';
    var v = String(f.filter).trim();
    if (/\.\.|[|&<>*@()=]/.test(v)) return '';
    return v.replace(/^'(.*)'$/, '$1');
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
    var a = anchorOf(source), n = source.dataCols, time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    applying(function () {
      (res.results || []).forEach(function (r) {
        var st = sheet.getRange(r.row, a.col + n, 1, 1);
        if (r.ok) {
          var values = source.columns.map(function (c, i) { return convertValue(c.type, r.values[i]); });
          sheet.getRange(r.row, a.col, 1, n).setValues([values]);
          var change = w.changes.filter(function (x) { return x.r === r.row; })[0];
          if (change && change.op === 'modify') { delete source.orig[change.key]; delete source.ver[change.key]; }
          if (change && change.op === 'insert' && r.row > a.row + (source.rowCount || 0)) source.rowCount = r.row - a.row;
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
    source.lastRows = Math.max(source.lastRows || 0, (source.rowCount || 0) + 1);
    setDirty(true);
    diffSource(source);
    updateRegionBar(); renderSources(); renderBarSources();
    w.resolve({ saved: res.saved || 0, failed: res.failed || 0 });
  }

  /** Before a refresh replaces an editable region: ask when it still has unsaved BC changes. */
  function confirmDiscard(sources) {
    var pending = sources.filter(function (s) { return (isTracking(s) && diffSource(s).length) || (budgetTracking(s) && budgetDiff(s).length); });
    if (!pending.length) return Promise.resolve(true);
    var n = pending.reduce(function (t, s) { return t + s.changeCount; }, 0);
    return H.choose(n + (n === 1 ? ' row has' : ' rows have') + ' changes that are not saved to Business Central. Refresh replaces them with the data from BC.',
      [{ label: 'Refresh and discard', value: true, primary: true }], 'Unsaved BC changes')
      .then(function (v) { return v === true; });
  }

  // ---------------------------------------------------------------- region bar (shown when the selection is in a region)
  function updateRegionBar() {
    if (!ui.regionbar) return;
    var cell = workbook() ? selectedCell() : null;
    var s = cell && regionAt(cell.sheetId, cell.row, cell.col);
    state.activeRegion = s ? s.id : null;
    if (!s) { ui.regionbar.hidden = true; ui.app.classList.remove('has-regionbar'); return; }
    var editSwitch = s.editable
      ? '<label class="ak-rswitch" title="Off: view and use formulas only, nothing is sent to BC"><input type="checkbox" data-region-edit' + (s.editing !== false ? ' checked' : '') + '>' +
        '<span class="ak-switch" aria-hidden="true"></span><span>Editing</span></label>'
      : '<span class="ak-rtag ak-rtag--ro">' + H.icon('lock') + 'Read-only</span>';
    var refs = '';
    try { refs = parseRefs(s).length ? '<span class="ak-rtag" title="A filter takes its value from a cell; change the cell and the region refreshes">' + H.icon('link') + 'Cell filter</span>' : ''; } catch (e) { refs = ''; }
    ui.regionbar.innerHTML =
      '<span class="ak-rbar__icon">' + H.icon('table') + '</span>' +
      '<span class="ak-rbar__name">' + H.escapeHtml(s.tableCaption) + '</span>' +
      '<span class="ak-rbar__meta">' + H.escapeHtml(regionLabel(s)) + ' &middot; ' + H.formatInt(s.rowCount || 0) + (isBudget(s) ? ' accounts &middot; ' + (s.budget.periods) + ' ' + (s.budget.periodType === 'quarter' ? 'quarters' : 'months') : (s.rowCount === 1 ? ' row' : ' rows')) +
      (s.filters && s.filters.length ? ' &middot; ' + s.filters.length + (s.filters.length === 1 ? ' filter' : ' filters') : '') + '</span>' +
      refs + editSwitch +
      // Editable but new rows not allowed: say so, otherwise the box has no empty line and does not grow.
      (s.editable && !isBudget(s) && !s.canInsert
        ? '<span class="ak-rtag ak-rtag--warn" title="To add new rows, an admin ticks Allow Insert for this table on the Edit Tables page, and you need insert permission on it">New rows off</span>'
        : '') +
      '<span class="ak-rbar__sp"></span>' +
      '<span class="ak-rbar__hint" title="Ctrl+click a row or amount to open it in Business Central">Ctrl+click to open in BC</span>' +
      '<button class="ak-btn ak-btn--subtle ak-btn--sm" data-act="region-select" title="Select the whole region (header and rows). Click again to select only the data rows.">' + H.icon('table') + '<span>Select</span></button>' +
      '<button class="ak-btn ak-btn--subtle ak-btn--sm" data-act="region-open" title="Open the selected row or amount in Business Central">' + H.icon('link') + '<span>Open</span></button>' +
      '<button class="ak-btn ak-btn--subtle ak-btn--sm" data-act="region-props">' + H.icon('sliders') + '<span>Properties</span></button>' +
      '<button class="ak-btn ak-btn--subtle ak-btn--sm" data-act="region-refresh">' + H.icon('refresh') + '<span>Refresh</span></button>' +
      '<button class="ak-btn ak-btn--subtle ak-btn--sm" data-act="region-remove" title="Remove the region from the sheet, or keep the data and only stop linking it to BC">' + H.icon('x') + '<span>Remove</span></button>';
    ui.regionbar.hidden = false;
    ui.app.classList.add('has-regionbar');
  }

  function activeRegion() { return state.activeRegion ? state.sources[state.activeRegion] : null; }

  /** Range of a region: 'all' = header + rows, 'data' = rows only (the empty new-row line is not included). */
  function regionRange(s, part) {
    var a = anchorOf(s), first = part === 'data' ? a.row + 1 : a.row, last = Math.max(regionEndRow(s), first);
    return { r: first, c: a.col, rows: last - first + 1, cols: Math.max(regionWidth(s), 1) };
  }

  /** Selects a region like clicking the corner of an Excel table. 'toggle': whole region first, then only the data rows. */
  function selectRegion(s, part) {
    var wb = workbook(), sheet = wb.getSheetBySheetId(s.sheetId);
    if (!sheet) return;
    if (part === 'toggle') {
      var all = regionRange(s, 'all'), cur = selectionRect();
      part = cur.sheetId === s.sheetId && cur.r === all.r && cur.c === all.c && cur.rows === all.rows && cur.cols === all.cols ? 'data' : 'all';
    }
    var g = regionRange(s, part);
    if (wb.getActiveSheet().getSheetId() !== s.sheetId) wb.setActiveSheet(sheet);
    sheet.getRange(g.r, g.c, g.rows, g.cols).activate();
    H.toast('info', (part === 'data' ? 'Data rows selected: ' : 'Region selected: ') + cellName(g.r, g.c) + ':' + cellName(g.r + g.rows - 1, g.c + g.cols - 1) +
      '. Copy it, format it or click Insert chart.');
  }

  // ---------------------------------------------------------------- Excel-style region look
  /** Excel "Table Style Medium 2" look: blue header, thin border, banded rows for read-only regions. */
  function styleRegion(source, sheet, nRows, nCols) {
    var a = anchorOf(source), api = state.api;
    try {
      var head = sheet.getRange(a.row, a.col, 1, nCols);
      head.setBackgroundColor(XL.head);
      head.setFontColor(XL.headText);
      head.setFontWeight('bold');
    } catch (e) { /* optional */ }
    try {
      // Clear the border of the previous load first (the region may have shrunk), then draw the new one.
      if (source.styledArea) {
        // Univer also stores an edge on the neighbouring cell, so clear one row below and one column right as well.
        sheet.getRange(a.row, a.col, source.styledArea.rows + 1, source.styledArea.cols + 1)
          .setBorder(api.Enum.BorderType.NONE, api.Enum.BorderStyleTypes.NONE);
      }
      var all = sheet.getRange(a.row, a.col, Math.max(nRows, 1), nCols);
      all.setBorder(api.Enum.BorderType.OUTSIDE, api.Enum.BorderStyleTypes.THIN, XL.border);
      source.styledArea = { rows: Math.max(nRows, 1), cols: nCols };
    } catch (e) { /* borders optional */ }
    // Banding through one conditional-format rule, removed and re-added on every load.
    try {
      if (source._bandRule) { sheet.deleteConditionalFormattingRule(source._bandRule); source._bandRule = null; }
      (sheet.getConditionalFormattingRules() || []).forEach(function (rule) {
        if (source.bandRuleId && rule.cfId === source.bandRuleId) sheet.deleteConditionalFormattingRule(rule.cfId);
      });
      source.bandRuleId = null;
      if (!source.editable && nRows > 2) {
        var body = sheet.getRange(a.row + 1, a.col, nRows - 1, nCols);
        var rule = sheet.newConditionalFormattingRule().whenFormulaSatisfied('=MOD(ROW()-' + (a.row + 1) + ',2)=1')
          .setBackground(XL.band).setRanges([body.getRange()]).build();
        sheet.addConditionalFormattingRule(rule);
        source.bandRuleId = rule.cfId;
      }
    } catch (e) { /* banding optional */ }
  }

  // ---------------------------------------------------------------- right-click menu
  function registerMenus() {
    var api = state.api;
    if (!api || !api.createMenu) return;
    var add = function (id, title, fn, order) {
      try { api.createMenu({ id: id, title: title, tooltip: title, order: order, action: function () { guard(fn)(); return true; } }).appendTo('contextMenu.mainArea|contextMenu.others'); }
      catch (e) { /* menus optional */ }
    };
    add('ak.drill', 'Open in Business Central (Ctrl+click)', function () { drillAt(selectedCell()); }, 0);
    add('ak.chart.insert', 'Insert chart from selection', function () { openChartPicker(); }, 4);
    add('ak.region.insert', 'Insert Business Central region here', function () { openRegionPanel(null); }, 0);
    add('ak.budget.insert', 'Insert budget grid here', function () { openBudgetPanel(null); }, 1);
    add('ak.period.insert', 'Insert period picker here', function () { insertPeriodPicker(); }, 1);
    add('ak.region.props', 'Region properties', function () { var s = regionAtSelection(); if (s) { if (isBudget(s)) openBudgetPanel(s); else openRegionPanel(s); } else H.toast('info', 'Select a cell inside a region first.'); }, 2);
    add('ak.region.select', 'Select region', function () { var s = regionAtSelection(); if (s) selectRegion(s, 'all'); else H.toast('info', 'Select a cell inside a region first.'); }, 2);
    add('ak.region.selectdata', 'Select region data rows', function () { var s = regionAtSelection(); if (s) selectRegion(s, 'data'); else H.toast('info', 'Select a cell inside a region first.'); }, 2);
    add('ak.region.remove', 'Remove region', function () { var s = regionAtSelection(); if (s) askRemoveRegion(s); else H.toast('info', 'Select a cell inside a region first.'); }, 3);
    add('ak.region.refresh', 'Refresh region', function () { var s = regionAtSelection(); if (s) refreshOne(s.id); else H.toast('info', 'Select a cell inside a region first.'); }, 3);
  }

  function regionAtSelection() {
    var c = selectedCell();
    return regionAt(c.sheetId, c.row, c.col);
  }

  /** Opens the side panel for a new region at the selected cell (source = null) or to change an existing region. */
  function openRegionPanel(source) {
    var wasEditing = !!state.regionEdit;
    state.regionEdit = source ? source.id : null;
    if (!source) setDest('ak-dest', 'here');
    if (!source && wasEditing && state.table) resetTablePick();
    var c = selectedCell();
    state.regionAnchor = source ? { sheetId: source.sheetId, row: anchorOf(source).row, col: anchorOf(source).col } : c;
    if (!ui.app.classList.contains('is-panel-open') || ui.panel.getAttribute('data-view') !== 'data') togglePanel('data');
    renderRegionTarget();
    if (source) {
      state.prefill = source;
      ui.tableSearch.value = '';
      if (!state.tables.some(function (t) { return t.id === source.tableNo; })) {
        state.tables.push({ id: source.tableNo, name: source.tableCaption, caption: source.tableCaption });
      }
      pickTable(source.tableNo);
    }
  }

  /**
   * While the region or budget panel is open, it follows the cell the user clicks: a click outside the region
   * being changed switches the panel back to "new region at <that cell>", so the button never updates the
   * previous region by surprise.
   */
  function followSelectionInPanels() {
    if (!ui.app.classList.contains('is-panel-open')) return;
    var view = ui.panel.getAttribute('data-view'), c = selectedCell();
    if (view === 'data') {
      var editing = state.regionEdit && state.sources[state.regionEdit];
      if (editing && regionAt(c.sheetId, c.row, c.col) === editing) return;
      var wasEditing = !!editing;
      state.regionEdit = null;
      state.regionAnchor = c;
      // A click on a cell means "put the next region here": switch back from New sheet to At <cell>.
      if (wasEditing) setDest('ak-dest', 'here');
      renderRegionTarget();
    } else if (view === 'budget') {
      var b = state.budgetEdit && state.sources[state.budgetEdit];
      if (b && regionAt(c.sheetId, c.row, c.col) === b) return;
      if (b) { state.budgetEdit = null; prepareBudgetPanel(); return; }
      state.budgetAnchor = c;
      ui.bDestCell.textContent = sheetName(c.sheetId) + '!' + cellName(c.row, c.col);
    }
  }

  function setDest(name, value) {
    var r = ui.panel.querySelector('input[name=' + name + '][value=' + value + ']');
    if (r) r.checked = true;
  }

  function renderRegionTarget() {
    var editing = state.regionEdit && state.sources[state.regionEdit];
    var t = state.regionAnchor;
    var here = t ? sheetName(t.sheetId) + '!' + cellName(t.row, t.col) : 'A1';
    ui.destCell.textContent = here;
    ui.regionTarget.textContent = editing ? 'Changing region ' + editing.tableCaption + ' at ' + regionLabel(editing) : 'New region at ' + here;
    ui.destWrap.hidden = !!editing;
    ui.loadLabel.textContent = editing ? 'Update region' : 'Insert region';
  }

  // =====================================================================================
  // Stage 2: journal and document lines (quick starts in the region panel, regions opened from BC pages)
  // =====================================================================================
  function renderQuickStarts() {
    var list = state.config.quickStarts || [];
    ui.quick.hidden = !list.length;
    ui.quick.innerHTML = list.length ? '<span class="ak-quick__label">Quick start</span>' + list.map(function (q) {
      var r = q.region || {};
      var meta = (r.filters || []).map(function (f) { return f.filter; }).filter(Boolean).join(' / ');
      return '<button class="ak-quick__btn" data-act="quick-start" data-quick="' + H.escapeHtml(q.id) + '" title="' + H.escapeHtml(r.tableCaption || '') + '">' +
        H.icon(/jnl/.test(q.id) ? 'receipt' : 'table') + '<span><b>' + H.escapeHtml(q.title) + '</b>' +
        (meta ? '<small>' + H.escapeHtml(meta) + '</small>' : '') + '</span></button>';
    }).join('') : '';
    ui.quick.hidden = false;
    if (!list.length) ui.quick.innerHTML = '<span class="ak-quick__label">Quick start</span>';
    ui.quick.innerHTML += '<button class="ak-quick__btn" data-act="period-picker" title="A dropdown cell that drives date filters">' + H.icon('calendar') +
      '<span><b>Period picker</b><small>This Month, YTD, Last Year&hellip;</small></span></button>';
    var b = state.config.budget;
    if (b && (b.names || []).length) {
      ui.quick.innerHTML += '<button class="ak-quick__btn" data-act="open-budget" title="G/L accounts by period">' + H.icon('sliders') +
        '<span><b>Budget grid</b><small>' + H.escapeHtml(b.names.map(function (x) { return x.name; }).slice(0, 3).join(', ')) + '</small></span></button>';
    }
  }

  /** A quick start fills the panel with the region's table, fields (in page order) and filters; the user can still change them. */
  function startQuick(id) {
    var q = (state.config.quickStarts || []).filter(function (x) { return x.id === id; })[0];
    if (!q) return;
    var r = q.region;
    state.prefill = { tableNo: r.tableNo, fields: r.fields, filters: r.filters, editable: r.editable, maxRows: state.config.defaultMaxRows };
    if (!state.tables.some(function (t) { return t.id === r.tableNo; })) state.tables.push({ id: r.tableNo, name: r.tableCaption, caption: r.tableCaption });
    ui.sheetName.value = H.safeSheetName(q.title);
    pickTable(r.tableNo);
  }

  /** Workbook opened from a journal or document page: show that region (refresh it) or create it. */
  function applyAutoRegion(def) {
    if (!def || !workbook()) return;
    var sig = function (filters) {
      return JSON.stringify((filters || []).map(function (f) { return [f.fieldNo, String(f.filter)]; }).sort());
    };
    var found = sourceList().filter(function (s) { return s.tableNo === def.tableNo && sig(s.filters) === sig(def.filters); })[0];
    if (found) {
      activateSheet(found.id);
      updateRegionBar();
      confirmDiscard([found]).then(function (ok) { if (ok) runSource(found).then(function () { updateRegionBar(); }, function () { }); });
      return;
    }
    var wb = workbook(), sheet = wb.getActiveSheet();
    var a1 = sheet.getRange(0, 0, 1, 1).getValue();
    var empty = sheet.getLastRow() <= 0 && sheet.getLastColumn() <= 0 && (a1 == null || a1 === '');
    if (!empty || regionsOn(sheet.getSheetId()).length) {
      sheet = wb.insertSheet(uniqueSheetName(def.tableCaption));
      excelLook(sheet);
    }
    wb.setActiveSheet(sheet);
    var source = {
      id: H.uid('rg'), sheetId: sheet.getSheetId(), anchor: { row: 0, col: 0 }, lastRows: 0, lastCols: 0,
      tableNo: def.tableNo, tableCaption: def.tableCaption,
      fields: def.fields.map(function (f) { return { no: f.no, caption: f.caption, type: f.type, pk: !!f.pk }; }),
      filters: def.filters || [], maxRows: state.config.defaultMaxRows, editable: !!def.editable
    };
    if (source.editable) source.editing = true;
    runSource(source).then(function (r) {
      updateRegionBar();
      H.toast('success', def.tableCaption + ': ' + H.formatInt(r.loaded) + ' lines loaded.' +
        (source.editable ? ' Type or paste lines below, then Save to BC.' : ''));
      if (!source.editable && def.editReason) H.toast('info', 'Read-only: ' + def.editReason);
    }, function () { });
  }

  // =====================================================================================
  // Stage 3: budget grid region (G/L accounts x periods, written as G/L Budget Entries)
  //
  // A budget region is a source with kind 'budget'. Columns: No., Name, one column per period, Total
  // (a SUM formula), optionally last year's actuals per period and their total, and BC Status.
  // Only period cells of posting accounts are editable. Typing a number into Total spreads it over the
  // periods (by last year's actuals when shown, otherwise evenly). Save to BC sends changed cells to AL,
  // which adds one G/L Budget Entry for each cell's difference.
  // =====================================================================================
  var BUDGET_FILTERS = { account: -1, dim1: -2, dim2: -3 };

  function isBudget(s) { return !!(s && s.kind === 'budget'); }
  function budgetTracking(s) { return isBudget(s) && s.editable && s.editing !== false && s.budgetOrig; }
  function budgetList() { return sourceList().filter(budgetTracking); }

  function budgetViewHtml() {
    return '<section class="ak-view" data-view="budget" aria-label="Budget grid">' +
      panelHead('Budget grid', 'G/L accounts by period, saved as G/L Budget Entries') +
      '<div class="ak-panel__scroll">' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">1</span>Budget</div>' +
      '    <label class="ak-field"><span>Budget name</span><select id="ak-b-name" class="ak-input ak-select"></select></label>' +
      '    <div class="ak-grid2">' +
      '      <label class="ak-field"><span>Starting date</span><input id="ak-b-start" class="ak-input" type="date"></label>' +
      '      <label class="ak-field"><span>Periods</span><input id="ak-b-periods" class="ak-input" type="number" min="1" max="24" value="12"></label>' +
      '    </div>' +
      '    <div class="ak-seg" role="radiogroup" aria-label="Period length">' +
      '      <label><input type="radio" name="ak-b-type" value="month" checked><span>Months</span></label>' +
      '      <label><input type="radio" name="ak-b-type" value="quarter"><span>Quarters</span></label>' +
      '    </div>' +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">2</span>Filters</div>' +
      '    <label class="ak-field"><span>G/L accounts</span><input id="ak-b-acc" class="ak-input" placeholder="e.g. 40000..69999 (empty = all)"></label>' +
      '    <label class="ak-field"><span id="ak-b-d1cap">Global Dimension 1</span><input id="ak-b-d1" class="ak-input" placeholder="empty = all, or one value, or [B1]"></label>' +
      '    <label class="ak-field"><span id="ak-b-d2cap">Global Dimension 2</span><input id="ak-b-d2" class="ak-input" placeholder="empty = all, or one value, or [B2]"></label>' +
      '    <p class="ak-help">Filters can point to a cell, like <code>[B1]</code>: change the cell and the grid reloads. To save changes, dimension filters must be empty or one value.</p>' +
      '  </section>' +
      '  <section class="ak-step">' +
      '    <div class="ak-step__label"><span class="ak-step__num">3</span>Output</div>' +
      '    <label class="ak-check"><input type="checkbox" id="ak-b-ly" checked><span>Show last year\'s actuals next to the budget</span></label>' +
      '    <div class="ak-seg ak-mt" role="radiogroup" aria-label="Where" id="ak-b-dest-wrap">' +
      '      <label><input type="radio" name="ak-b-dest" value="here" checked><span>At <b id="ak-b-dest-cell">A1</b></span></label>' +
      '      <label><input type="radio" name="ak-b-dest" value="new"><span>New sheet</span></label>' +
      '    </div>' +
      '    <p class="ak-help ak-edit-why" id="ak-b-why" hidden></p>' +
      '  </section>' +
      '</div>' +
      '<div class="ak-panel__foot">' +
      '  <span class="ak-panel__hint" id="ak-b-hint">Type amounts in the period columns, or a yearly amount in Total to spread it.</span>' +
      '  <button class="ak-btn ak-btn--primary ak-btn--lg" id="ak-b-run">' + H.icon('table') + '<span id="ak-b-run-label">Insert budget grid</span></button>' +
      '</div></section>';
  }

  function prepareBudgetPanel() {
    var cfg = state.config.budget || {};
    var editing = state.budgetEdit && state.sources[state.budgetEdit];
    var def = editing ? editing.budget : null;
    ui.bName.innerHTML = (cfg.names || []).map(function (n) {
      return '<option value="' + H.escapeHtml(n.name) + '">' + H.escapeHtml(n.name + (n.description ? ' · ' + n.description : '') + (n.blocked ? ' (blocked)' : '')) + '</option>';
    }).join('');
    ui.bD1cap.textContent = cfg.dim1Caption || 'Global Dimension 1';
    ui.bD2cap.textContent = cfg.dim2Caption || 'Global Dimension 2';
    if (def) {
      ui.bName.value = def.budgetName;
      ui.bStart.value = def.startDate;
      ui.bPeriods.value = def.periods;
      ui.bAcc.value = def.accountFilter || '';
      ui.bD1.value = def.dim1Filter || '';
      ui.bD2.value = def.dim2Filter || '';
      ui.bLy.checked = !!def.showLY;
      (ui.panel.querySelector('input[name=ak-b-type][value=' + (def.periodType || 'month') + ']') || {}).checked = true;
    } else if (!ui.bStart.value) {
      ui.bStart.value = cfg.startDate || new Date().toISOString().slice(0, 10);
    }
    var t = editing ? { sheetId: editing.sheetId, row: anchorOf(editing).row, col: anchorOf(editing).col } : selectedCell();
    state.budgetAnchor = t;
    ui.bDestCell.textContent = sheetName(t.sheetId) + '!' + cellName(t.row, t.col);
    ui.bDestWrap.hidden = !!editing;
    ui.bRunLabel.textContent = editing ? 'Update budget grid' : 'Insert budget grid';
    ui.bWhy.hidden = !!cfg.canWrite;
    ui.bWhy.textContent = cfg.canWrite ? '' : 'Read-only: ' + (cfg.writeReason || 'budgets cannot be changed from a workbook.');
    ui.bRun.disabled = !(cfg.names || []).length;
    if (!(cfg.names || []).length) ui.bHint.textContent = 'No G/L budgets yet. Create one on the G/L Budgets page first.';
  }

  function openBudgetPanel(source) {
    state.budgetEdit = source ? source.id : null;
    if (!ui.app.classList.contains('is-panel-open') || ui.panel.getAttribute('data-view') !== 'budget') togglePanel('budget');
    else prepareBudgetPanel();
  }

  function insertBudgetFromPanel() {
    var cfg = state.config.budget || {};
    var wb = workbook();
    var existing = state.budgetEdit && state.sources[state.budgetEdit];
    var dest = (ui.panel.querySelector('input[name=ak-b-dest]:checked') || {}).value || 'here';
    var def = {
      budgetName: ui.bName.value,
      startDate: ui.bStart.value,
      periods: Math.max(1, Math.min(24, parseInt(ui.bPeriods.value, 10) || 12)),
      periodType: (ui.panel.querySelector('input[name=ak-b-type]:checked') || {}).value || 'month',
      accountFilter: ui.bAcc.value.trim(),
      dim1Filter: ui.bD1.value.trim(),
      dim2Filter: ui.bD2.value.trim(),
      showLY: ui.bLy.checked
    };
    if (!def.budgetName || !def.startDate) { H.toast('warning', 'Choose a budget and a starting date.'); return; }
    var sheet, anchor;
    if (existing) {
      sheet = wb.getSheetBySheetId(existing.sheetId); anchor = anchorOf(existing);
    } else if (dest === 'new') {
      sheet = wb.insertSheet(uniqueSheetName('Budget ' + def.budgetName)); excelLook(sheet); anchor = { row: 0, col: 0 };
    } else {
      var t = state.budgetAnchor || selectedCell();
      sheet = wb.getSheetBySheetId(t.sheetId) || wb.getActiveSheet(); anchor = { row: t.row, col: t.col };
      var clash = regionAt(sheet.getSheetId(), anchor.row, anchor.col);
      if (clash) { H.toast('error', 'This cell is inside region ' + clash.tableCaption + '. Pick an empty cell.'); return; }
      if (slicerAt(sheet.getSheetId(), anchor.row, anchor.col)) { H.toast('error', cellName(anchor.row, anchor.col) + ' is a period picker. Click the cell where the grid should start.'); return; }
    }
    wb.setActiveSheet(sheet);
    var source = existing || { id: H.uid('rg'), kind: 'budget', sheetId: sheet.getSheetId(), anchor: anchor, lastRows: 0, lastCols: 0, fields: [] };
    source.budget = def;
    source.tableCaption = 'Budget ' + def.budgetName;
    source.filters = [
      { fieldNo: BUDGET_FILTERS.account, filter: def.accountFilter },
      { fieldNo: BUDGET_FILTERS.dim1, filter: def.dim1Filter },
      { fieldNo: BUDGET_FILTERS.dim2, filter: def.dim2Filter }
    ].filter(function (f) { return f.filter; });
    source.editable = !!cfg.canWrite;
    if (source.editing === undefined) source.editing = true;
    try { parseRefs(source); } catch (e) { H.toast('error', e.message); return; }
    var go = function () {
      runBudget(source).then(function (r) {
        H.toast('success', 'Budget ' + def.budgetName + ': ' + H.formatInt(r.accounts) + ' accounts, ' + r.periods + ' periods at ' + regionLabel(source) + '.');
        state.budgetEdit = source.id;
        prepareBudgetPanel();
        updateRegionBar();
      }, function () { });
    };
    if (existing) confirmDiscard([existing]).then(function (ok) { if (ok) go(); }); else go();
  }

  function budgetFilterValue(source, which) {
    var no = BUDGET_FILTERS[which];
    var f = resolveFilters(source).filter(function (x) { return x.fieldNo === no; })[0];
    return f ? f.filter : '';
  }

  function runBudget(source) {
    setBusy(true);
    setProgress(-1);
    var requestId = 'b' + (++state.requestSeq);
    return new Promise(function (resolve, reject) {
      var payload;
      try {
        payload = {
          requestId: requestId, budgetName: source.budget.budgetName, startDate: source.budget.startDate,
          periods: source.budget.periods, periodType: source.budget.periodType, showLY: !!source.budget.showLY,
          accountFilter: budgetFilterValue(source, 'account'), dim1Filter: budgetFilterValue(source, 'dim1'), dim2Filter: budgetFilterValue(source, 'dim2')
        };
      } catch (e) { H.toast('error', e.message); reject(e); return; }
      state.budgetPending[requestId] = { source: source, resolve: resolve, reject: reject };
      H.sendChunked('budget', JSON.stringify(payload), state.config.chunkSize).catch(function (e) {
        delete state.budgetPending[requestId]; reject(e);
      });
    }).then(function (r) { setBusy(false); return r; }, function (e) { setBusy(false); throw e; });
  }

  /** AL answer to a budget load: writes the grid, the Total formulas, last year's block and the look. */
  function onBudgetData(text) {
    var res = JSON.parse(text);
    var p = state.budgetPending[res.requestId];
    if (!p) return;
    delete state.budgetPending[res.requestId];
    if (res.error) { H.toast('error', res.error); p.reject(new Error(res.error)); return; }
    var source = p.source, def = source.budget, periods = res.periods, n = periods.length, ly = !!def.showLY;
    var header = ['No.', 'Name'].concat(periods.map(function (x) { return x.label; }), ['Total']);
    if (ly) header = header.concat(periods.map(function (x) { return 'LY ' + x.label; }), ['LY Total']);
    header.push('BC Status');
    var types = header.map(function (h, i) { return i < 2 || i === header.length - 1 ? 'Text' : 'Decimal'; });
    var rows = res.rows.map(function (r) {
      var line = [r.no, new Array((r.indent || 0) + 1).join('   ') + r.name].concat(r.budget, [0]);
      if (ly) line = line.concat(r.ly, [0]);
      line.push('');
      return line;
    });
    var room = roomBelow(source);
    if (rows.length + 1 > room) { H.toast('warning', 'Only ' + Math.max(room - 1, 0) + ' accounts fit above the next region.'); rows = rows.slice(0, Math.max(room - 1, 0)); }
    source.editable = !!res.canWrite && !!(state.config.budget || {}).canWrite;
    source.periodsDef = periods;
    source.budgetMeta = { n: n, ly: ly, totalCol: 2 + n, lyStart: 3 + n, lyTotalCol: 3 + 2 * n, statusCol: header.length - 1 };
    source.dataCols = header.length - 1;
    source.budgetOrig = {};
    source.budgetPosting = {};
    res.rows.slice(0, rows.length).forEach(function (r) { source.budgetOrig[r.no] = r.budget.slice(); source.budgetPosting[r.no] = !!r.posting; });
    source.rowCount = rows.length;
    source.budgetGrid = rows.map(function (r) { return r.slice(0, source.dataCols); });
    applying(function () {
      var sheet = writeTable(source, header, types, rows);
      paintBudget(source, sheet, res.rows.slice(0, rows.length));
      styleRegion(source, sheet, rows.length + 1, header.length);
    });
    source.loadedAt = new Date().toISOString();
    state.sources[source.id] = source;
    budgetDiff(source);
    renderSources(); renderBarSources(); updateRegionBar(); setDirty(true);
    p.resolve({ accounts: rows.length, periods: n });
  }

  function paintBudget(source, sheet, meta) {
    var a = anchorOf(source), m = source.budgetMeta, n = m.n, width = source.dataCols + 1;
    var col = function (c) { return H.colLetter(a.col + c); };
    meta.forEach(function (r, i) {
      var row = a.row + 1 + i, rn = row + 1;
      sheet.getRange(row, a.col + m.totalCol, 1, 1).setFormula('=SUM(' + col(2) + rn + ':' + col(1 + n) + rn + ')');
      if (m.ly) sheet.getRange(row, a.col + m.lyTotalCol, 1, 1).setFormula('=SUM(' + col(m.lyStart) + rn + ':' + col(m.lyStart + n - 1) + rn + ')');
      if (!r.posting) {
        var line = sheet.getRange(row, a.col, 1, width - 1);
        line.setFontWeight('bold');
        line.setBackgroundColor('#F2F2F2');
      }
    });
    var body = Math.max(meta.length, 1);
    try { sheet.getRange(a.row + 1, a.col, body, 2).setFontColor('#262626'); } catch (e) { /* optional */ }
    try {
      var totals = sheet.getRange(a.row + 1, a.col + m.totalCol, body, 1);
      totals.setFontWeight('bold'); totals.setBackgroundColor('#EEF3FA');
    } catch (e) { /* optional */ }
    if (m.ly) {
      try {
        var lyRange = sheet.getRange(a.row + 1, a.col + m.lyStart, body, n + 1);
        lyRange.setFontColor('#595959'); lyRange.setBackgroundColor('#F7F7F7');
        sheet.getRange(a.row, a.col + m.lyStart, 1, n + 1).setBackgroundColor('#8EA9DB');
      } catch (e) { /* optional */ }
    }
    try {
      var st = sheet.getRange(a.row, a.col + m.statusCol, 1, 1);
      st.setBackgroundColor(XL.statusHead);
      sheet.setColumnWidths(a.col + m.statusCol, 1, 260);
      sheet.setColumnWidths(a.col + 1, 1, 240);
      sheet.setColumnWidths(a.col + 2, n, 96);
      sheet.setColumnWidths(a.col + m.totalCol, 1, 118);
      if (m.ly) sheet.setColumnWidths(a.col + m.lyTotalCol, 1, 118);
      if (m.ly) sheet.setColumnWidths(a.col + m.lyStart, n, 92);
    } catch (e) { /* optional */ }
    if (a.row === 0) { try { sheet.setFrozenColumns(a.col + 2); } catch (e) { /* optional */ } }
  }

  /** Cell edits inside a budget region: period cells of posting rows are tracked, Total spreads, the rest is put back. */
  function onBudgetCells(source, cells) {
    if (!budgetTracking(source)) return;
    var sheet = workbook().getSheetBySheetId(source.sheetId);
    var a = anchorOf(source), m = source.budgetMeta, reverts = [], spreads = [];
    cells.forEach(function (x) {
      var r = x.r - a.row - 1, c = x.c - a.col;
      if (r < 0 || r >= source.rowCount || c >= source.dataCols) return;
      var no = String(sheet.getRange(x.r, a.col, 1, 1).getValue());
      var posting = source.budgetPosting[no];
      if (c >= 2 && c < 2 + m.n && posting) return; // editable period cell
      if (c === m.totalCol && posting) { spreads.push({ r: x.r, no: no }); return; }
      reverts.push(x);
    });
    if (spreads.length) setTimeout(function () { spreadTotals(source, sheet, spreads); }, 0);
    if (reverts.length) setTimeout(function () { revertBudgetCells(source, sheet, reverts); }, 0);
    clearTimeout(source._diffTimer);
    source._diffTimer = setTimeout(function () { budgetDiff(source); }, 350);
  }

  /** A number typed into Total is spread over the periods: by last year's actuals when shown and not zero, else evenly. */
  function spreadTotals(source, sheet, spreads) {
    var a = anchorOf(source), m = source.budgetMeta, col = function (c) { return H.colLetter(a.col + c); };
    spreads.forEach(function (x) {
      var cell = sheet.getRange(x.r, a.col + m.totalCol, 1, 1);
      // The typed number can come back as formatted text ("240,000.00"), so strip the separators.
      var raw = cell.getRawValues()[0][0];
      var total = typeof raw === 'number' ? raw : +String(raw == null ? '' : raw).replace(/[^0-9.\-]/g, '');
      if (!isFinite(total) || String(raw).trim() === '') return;
      var weights = [];
      if (m.ly) weights = sheet.getRange(x.r, a.col + m.lyStart, 1, m.n).getRawValues()[0].map(function (v) { return Math.abs(+v || 0); });
      var sum = weights.reduce(function (t, v) { return t + v; }, 0);
      if (!sum) { weights = []; for (var i = 0; i < m.n; i++) weights.push(1); sum = m.n; }
      var parts = weights.map(function (w) { return Math.round(total * w / sum * 100) / 100; });
      parts[m.n - 1] = Math.round((total - parts.slice(0, -1).reduce(function (t, v) { return t + v; }, 0)) * 100) / 100;
      sheet.getRange(x.r, a.col + 2, 1, m.n).setValues([parts]);
      applying(function () { cell.setFormula('=SUM(' + col(2) + (x.r + 1) + ':' + col(1 + m.n) + (x.r + 1) + ')'); });
    });
    H.toast('info', 'Spread over ' + m.n + ' periods' + (m.ly ? ' by last year\'s actuals.' : ' evenly.') + ' Save to BC when ready.');
    budgetDiff(source);
  }

  function revertBudgetCells(source, sheet, cells) {
    var a = anchorOf(source), m = source.budgetMeta, values = 0;
    applying(function () {
      cells.forEach(function (x) {
        var c = x.c - a.col, i = x.r - a.row - 1, rn = x.r + 1;
        if (c === m.totalCol || c === m.lyTotalCol) {
          var from = c === m.totalCol ? 2 : m.lyStart;
          sheet.getRange(x.r, x.c, 1, 1).setFormula('=SUM(' + H.colLetter(a.col + from) + rn + ':' + H.colLetter(a.col + from + m.n - 1) + rn + ')');
          return;
        }
        var line = source.budgetGrid && source.budgetGrid[i];
        if (line) { sheet.getRange(x.r, x.c, 1, 1).setValue(line[c] == null ? '' : line[c]); values++; }
      });
    });
    H.toast('warning', values
      ? 'Only the period amounts of posting accounts can be changed here. The value was put back.'
      : 'Total is a formula. Type a number in Total to spread it over the periods.');
  }

  function budgetDiff(source) {
    var sheet = workbook() && workbook().getSheetBySheetId(source.sheetId);
    if (!sheet || !budgetTracking(source)) { if (source) { source.changeCount = 0; updateWriteButton(); } return []; }
    var a = anchorOf(source), m = source.budgetMeta, n = m.n;
    var raw = source.rowCount ? sheet.getRange(a.row + 1, a.col, source.rowCount, source.dataCols + 1).getRawValues() : [];
    var changes = [];
    raw.forEach(function (line, i) {
      var no = String(line[0]), orig = source.budgetOrig[no];
      if (!orig || !source.budgetPosting[no]) return;
      for (var k = 0; k < n; k++) {
        var v = line[2 + k];
        var num = v === '' || v == null ? 0 : +String(v).replace(/,/g, '');
        if (!isFinite(num)) num = NaN;
        if (isNaN(num) || Math.abs(num - (+orig[k] || 0)) > 0.0001) changes.push({ r: a.row + 1 + i, c: a.col + 2 + k, k: k, account: no, amount: num, status: line[m.statusCol] });
      }
    });
    applying(function () {
      (source.marks || []).forEach(function (mk) { sheet.getRange(mk.r, mk.c, 1, 1).setBackgroundColor('#FFFFFF'); });
      source.marks = [];
      var rowsChanged = {};
      changes.forEach(function (x) {
        if (source.marks.length < 3000) { sheet.getRange(x.r, x.c, 1, 1).setBackgroundColor(EDIT_COLORS.changed); source.marks.push({ r: x.r, c: x.c }); }
        rowsChanged[x.r] = true;
      });
      for (var i = 0; i < raw.length; i++) {
        var r = a.row + 1 + i, st = raw[i][m.statusCol], cell = sheet.getRange(r, a.col + m.statusCol, 1, 1);
        if (rowsChanged[r] && (!st || st === STATUS_CHANGED || /^Saved/.test(String(st)))) { cell.setValue(STATUS_CHANGED); cell.setFontColor('#835B00'); }
        else if (!rowsChanged[r] && st === STATUS_CHANGED) cell.setValue('');
      }
    });
    source.changeCount = changes.length;
    updateWriteButton();
    return changes;
  }

  function writeBudget(source, changes) {
    var p = source.periodsDef;
    var payload = {
      requestId: 'bw' + (++state.requestSeq),
      budgetName: source.budget.budgetName,
      dim1Filter: budgetFilterValue(source, 'dim1'),
      dim2Filter: budgetFilterValue(source, 'dim2'),
      changes: changes.map(function (x) {
        return { row: x.r, col: x.c, account: x.account, start: p[x.k].start, end: p[x.k].end, amount: isNaN(x.amount) ? 'x' : String(Math.round(x.amount * 100) / 100) };
      })
    };
    return new Promise(function (resolve, reject) {
      state.writes[payload.requestId] = { source: source, resolve: resolve, changes: changes, budget: true };
      H.sendChunked('budgetwrite', JSON.stringify(payload), state.config.chunkSize).catch(function (e) {
        delete state.writes[payload.requestId]; reject(e);
      });
    });
  }

  function onBudgetResult(text) {
    var res = JSON.parse(text);
    var w = state.writes[res.requestId];
    if (!w) return;
    delete state.writes[res.requestId];
    var source = w.source;
    if (res.error) { H.toast('error', res.error); w.resolve({ saved: 0, failed: w.changes.length }); return; }
    var sheet = workbook().getSheetBySheetId(source.sheetId);
    var a = anchorOf(source), m = source.budgetMeta, time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    var errorsByRow = {}, savedRows = {};
    applying(function () {
      (res.results || []).forEach(function (r) {
        var ch = w.changes.filter(function (x) { return x.r === r.row && x.c === r.col; })[0];
        if (!ch) return;
        if (r.ok) {
          source.budgetOrig[ch.account][ch.k] = r.amount;
          sheet.getRange(r.row, r.col, 1, 1).setValue(r.amount);
          savedRows[r.row] = true;
        } else {
          (errorsByRow[r.row] = errorsByRow[r.row] || []).push(source.periodsDef[ch.k].label + ': ' + r.error);
          sheet.getRange(r.row, r.col, 1, 1).setBackgroundColor('#FDE7E9');
        }
      });
      Object.keys(savedRows).forEach(function (row) {
        if (errorsByRow[row]) return;
        var c = sheet.getRange(+row, a.col + m.statusCol, 1, 1); c.setValue('Saved ' + time); c.setFontColor('#0E700E');
      });
      Object.keys(errorsByRow).forEach(function (row) {
        var c = sheet.getRange(+row, a.col + m.statusCol, 1, 1); c.setValue(errorsByRow[row].join(' | ')); c.setFontColor('#A4262C');
      });
    });
    setDirty(true);
    budgetDiff(source);
    w.resolve({ saved: res.saved || 0, failed: res.failed || 0 });
    // Totaling accounts (Total revenue...) come from BC: reload the grid when everything was saved.
    if (!res.failed && res.saved) setTimeout(function () { if (!state.busy) runBudget(source).catch(function () { }); }, 300);
  }

  // =====================================================================================
  // Stage 4: period picker (a slicer cell)
  //
  // A picker is a cell with a dropdown of periods (This Month, Year to Date, Last Year...). The cell to its
  // right shows the dates. A region filter that points to the picker cell, e.g. Posting Date = [B2], gets
  // the period as yyyy-mm-dd..yyyy-mm-dd (AL turns it into the user's date format), so one pick refreshes
  // every region that uses it. A typed BC date filter (01-04-26..30-06-26) is passed through as typed.
  // =====================================================================================
  var PERIODS = ['Today', 'This Week', 'This Month', 'Last Month', 'This Quarter', 'Last Quarter',
    'Year to Date', 'This Year', 'Last Year', 'Fiscal Year to Date', 'This Fiscal Year', 'Last Fiscal Year',
    'Last 7 Days', 'Last 30 Days', 'Last 90 Days', 'Next 30 Days'];

  function isoDate(d) { return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2); }
  function utc(y, m, d) { return new Date(Date.UTC(y, m, d)); }
  function workDate() {
    var w = /^(\d{4})-(\d{2})-(\d{2})/.exec(state.config.workDate || '');
    var now = new Date();
    return w ? utc(+w[1], +w[2] - 1, +w[3]) : utc(now.getFullYear(), now.getMonth(), now.getDate());
  }
  function fiscalStart(d) {
    var f = /^(\d{4})-(\d{2})-(\d{2})/.exec((state.config.budget || {}).startDate || '');
    var month = f ? +f[2] - 1 : 0, day = f ? +f[3] : 1;
    var y = d.getUTCFullYear(), start = utc(y, month, day);
    return start > d ? utc(y - 1, month, day) : start;
  }
  function addDays(d, n) { return new Date(d.getTime() + n * 86400000); }

  /** Period name -> { from, to } (UTC dates), or null for text that is not a period name. */
  function periodRange(name) {
    var d = workDate(), y = d.getUTCFullYear(), m = d.getUTCMonth(), q = Math.floor(m / 3);
    var fs = fiscalStart(d), fy = function (k) { var s = utc(fs.getUTCFullYear() + k, fs.getUTCMonth(), fs.getUTCDate()); return { from: s, to: addDays(utc(s.getUTCFullYear() + 1, s.getUTCMonth(), s.getUTCDate()), -1) }; };
    switch (String(name).trim().toLowerCase()) {
      case 'today': return { from: d, to: d };
      case 'this week': var dow = (d.getUTCDay() + 6) % 7; return { from: addDays(d, -dow), to: addDays(d, 6 - dow) };
      case 'this month': return { from: utc(y, m, 1), to: utc(y, m + 1, 0) };
      case 'last month': return { from: utc(y, m - 1, 1), to: utc(y, m, 0) };
      case 'this quarter': return { from: utc(y, q * 3, 1), to: utc(y, q * 3 + 3, 0) };
      case 'last quarter': return { from: utc(y, q * 3 - 3, 1), to: utc(y, q * 3, 0) };
      case 'year to date': return { from: utc(y, 0, 1), to: d };
      case 'this year': return { from: utc(y, 0, 1), to: utc(y, 11, 31) };
      case 'last year': return { from: utc(y - 1, 0, 1), to: utc(y - 1, 11, 31) };
      case 'fiscal year to date': return { from: fs, to: d };
      case 'this fiscal year': return fy(0);
      case 'last fiscal year': return fy(-1);
      case 'last 7 days': return { from: addDays(d, -6), to: d };
      case 'last 30 days': return { from: addDays(d, -29), to: d };
      case 'last 90 days': return { from: addDays(d, -89), to: d };
      case 'next 30 days': return { from: d, to: addDays(d, 29) };
    }
    return null;
  }

  function slicerAt(sheetId, row, col) {
    return Object.keys(state.slicers).map(function (k) { return state.slicers[k]; })
      .filter(function (p) { return p.sheetId === sheetId && p.row === row && p.col === col; })[0] || null;
  }

  /** Filter text for a reference to a picker cell; null when the cell is not a picker. */
  function slicerFilter(sheetId, row, col, shown) {
    if (!slicerAt(sheetId, row, col)) return null;
    var r = periodRange(shown);
    return r ? isoDate(r.from) + '..' + isoDate(r.to) : String(shown == null ? '' : shown).trim();
  }

  function fmtShort(d) {
    return ('0' + d.getUTCDate()).slice(-2) + '-' + ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getUTCMonth()] + '-' + String(d.getUTCFullYear()).slice(2);
  }

  /** Writes the dates of the picked period into the cell to the right of the picker. */
  function showSlicerDates(p) {
    var sheet = workbook() && workbook().getSheetBySheetId(p.sheetId);
    if (!sheet) return;
    var v = sheet.getRange(p.row, p.col, 1, 1).getValue(), r = periodRange(v);
    applying(function () {
      var cell = sheet.getRange(p.row, p.col + 1, 1, 1);
      cell.setValue(r ? fmtShort(r.from) + ' to ' + fmtShort(r.to) : (v ? 'Custom filter' : ''));
      cell.setFontColor('#605E5C');
      cell.setFontStyle && cell.setFontStyle('italic');
    });
  }

  function insertPeriodPicker() {
    var c = selectedCell(), wb = workbook(), sheet = wb.getSheetBySheetId(c.sheetId);
    if (regionAt(c.sheetId, c.row, c.col)) { H.toast('error', 'Put the period picker outside a region, for example above it.'); return; }
    var p = { id: H.uid('pp'), sheetId: c.sheetId, row: c.row, col: c.col, kind: 'period' };
    state.slicers[p.id] = p;
    applying(function () {
      var cell = sheet.getRange(c.row, c.col, 1, 1);
      var current = cell.getValue();
      if (!current || !periodRange(current)) cell.setValue('This Month');
      cell.setBackgroundColor('#FFF2CC');
      cell.setFontWeight('bold');
      try { cell.setBorder(state.api.Enum.BorderType.OUTSIDE, state.api.Enum.BorderStyleTypes.THIN, '#BF9000'); } catch (e) { /* optional */ }
      try {
        var rule = state.api.newDataValidation().requireValueInList(PERIODS, false, true).setAllowInvalid(true).build();
        cell.setDataValidation(rule);
      } catch (e) { /* dropdown optional: typing a period name still works */ }
      sheet.setColumnWidths(c.col, 1, 150);
      sheet.setColumnWidths(c.col + 1, 1, 170);
    });
    showSlicerDates(p);
    setDirty(true);
    var ref = '[' + cellName(c.row, c.col) + ']';
    H.toast('success', 'Period picker in ' + cellName(c.row, c.col) + '. Use ' + ref + ' as the filter of a date field (for example Posting Date) in a region.');
  }

  /** Picker cells that changed: update their date text (regions using them refresh through the cell references). */
  function onSlicerCells(sheetId, cells) {
    cells.forEach(function (x) { var p = slicerAt(sheetId, x.r, x.c); if (p) showSlicerDates(p); });
  }

  // =====================================================================================
  // Stage 5: drill-down (Ctrl+click or right-click > Open in Business Central)
  // =====================================================================================
  /** Opens the BC page behind a cell: the record of a region row, or the entries behind a budget / actual amount. */
  function drillAt(cell) {
    var s = regionAt(cell.sheetId, cell.row, cell.col);
    if (!s) { H.toast('info', 'Ctrl+click a cell inside a Business Central region to open it in BC.'); return; }
    var a = anchorOf(s), sheet = workbook().getSheetBySheetId(s.sheetId);
    if (cell.row <= a.row) { H.toast('info', 'Ctrl+click a data row (not the header) to open it.'); return; }
    if (isBudget(s)) { drillBudget(s, sheet, cell); return; }
    var fields = s.columns || s.fields || [];
    var keyIdx = [];
    fields.forEach(function (f, i) { if (f.pk) keyIdx.push(i); });
    if (!keyIdx.length || keyIdx.length < (s.keyCount || keyIdx.length)) {
      H.toast('warning', 'Add the key fields of ' + s.tableCaption + ' to the region (Properties > Key) to open its records.');
      return;
    }
    var row = sheet.getRange(cell.row, a.col, 1, fields.length).getRawValues()[0];
    var key = keyIdx.map(function (i) { return toBC(fields[i].type, row[i]); });
    if (key.some(function (v) { return v === ''; })) {
      H.toast('info', s.editable && cell.row > anchorOf(s).row + (s.rowCount || 0)
        ? 'This is a new row. Save it to BC first, then Ctrl+click opens it.'
        : 'This row has no key values, so there is no record to open.');
      return;
    }
    sendDrill({ kind: 'record', tableNo: s.tableNo, keyFields: keyIdx.map(function (i) { return fields[i].no; }), key: key });
  }

  function drillBudget(s, sheet, cell) {
    var a = anchorOf(s), m = s.budgetMeta, p = s.periodsDef, c = cell.col - a.col;
    var account = String(sheet.getRange(cell.row, a.col, 1, 1).getValue() || '');
    if (!account) return;
    var shift = function (iso) { var d = new Date(iso + 'T00:00:00Z'); d.setUTCFullYear(d.getUTCFullYear() - 1); return isoDate(d); };
    var req = { budgetName: s.budget.budgetName, account: account, dim1Filter: budgetFilterValue(s, 'dim1'), dim2Filter: budgetFilterValue(s, 'dim2') };
    if (c >= 2 && c < 2 + m.n) { req.kind = 'budget'; req.start = p[c - 2].start; req.end = p[c - 2].end; }
    else if (c === m.totalCol || c < 2) { req.kind = 'budget'; req.start = p[0].start; req.end = p[m.n - 1].end; }
    else if (m.ly && c >= m.lyStart && c < m.lyStart + m.n) { req.kind = 'actual'; req.start = shift(p[c - m.lyStart].start); req.end = shift(p[c - m.lyStart].end); }
    else if (m.ly && c === m.lyTotalCol) { req.kind = 'actual'; req.start = shift(p[0].start); req.end = shift(p[m.n - 1].end); }
    else return;
    sendDrill(req);
  }

  function sendDrill(req) {
    H.sendChunked('drill', JSON.stringify(req), state.config.chunkSize).catch(reportError);
  }

  // =====================================================================================
  // Stage 6: BC functions in formulas (BCGLSUM, BCBUDGET, BCSUM, BCCOUNT, BCVALUE)
  //
  // Each function is an async Univer function. Calls made while the sheet calculates are collected for a
  // moment and sent to AL in one batch (purpose 'formulas'); the answers are cached per call until Refresh.
  // A period name (This Month, Last Year...) or a period picker value as the period argument becomes
  // yyyy-mm-dd..yyyy-mm-dd; anything else is passed as a BC filter.
  // =====================================================================================
  var BC_FUNCTIONS = [
    ['BCGLSUM', 'Net change of G/L accounts. BCGLSUM(account or filter, [period], [dimension 1 filter], [dimension 2 filter])'],
    ['BCBUDGET', 'Budgeted amount. BCBUDGET(budget name, account or filter, [period], [dimension 1 filter], [dimension 2 filter])'],
    ['BCSUM', 'Sum of a number field. BCSUM(table, field, [filter field, filter]...)'],
    ['BCCOUNT', 'Number of records. BCCOUNT(table, [filter field, filter]...)'],
    ['BCVALUE', 'A field of the first matching record. BCVALUE(table, field, [filter field, filter]...)']
  ];
  var bcf = { cache: {}, queue: {}, timer: null, pending: {} };

  function registerBCFunctions() {
    var f;
    try { f = state.api.getFormula(); } catch (e) { return; }
    if (!f || !f.registerAsyncFunction) return;
    BC_FUNCTIONS.forEach(function (d) {
      try {
        f.registerAsyncFunction(d[0], function () { return bcCall(d[0], Array.prototype.slice.call(arguments)); }, d[1]);
      } catch (e) { /* function registration optional */ }
    });
  }

  /** Univer passes a cell or range as a 2D array; use its first value. Period names become ISO date ranges. */
  function bcArg(fn, index, v) {
    while (Array.isArray(v)) v = v.length ? v[0] : '';
    if (v == null) return '';
    var periodArg = (fn === 'BCGLSUM' && index === 1) || (fn === 'BCBUDGET' && index === 2);
    var text = typeof v === 'number' ? String(Math.round(v * 1e9) / 1e9) : String(v).trim();
    if (periodArg || /^(bcsum|bccount|bcvalue)$/i.test(fn)) {
      var r = periodRange(text);
      if (r) return isoDate(r.from) + '..' + isoDate(r.to);
    }
    return text;
  }

  function bcCall(fn, rawArgs) {
    var args = rawArgs.map(function (v, i) { return bcArg(fn, i, v); });
    while (args.length && args[args.length - 1] === '') args.pop();
    var key = fn + '\u0001' + args.join('\u0001');
    if (Object.prototype.hasOwnProperty.call(bcf.cache, key)) return Promise.resolve(bcf.cache[key]);
    if (bcf.queue[key]) return bcf.queue[key].promise;
    var entry = { fn: fn, args: args };
    entry.promise = new Promise(function (resolve) { entry.resolve = resolve; });
    bcf.queue[key] = entry;
    clearTimeout(bcf.timer);
    bcf.timer = setTimeout(flushBCCalls, 60);
    return entry.promise;
  }

  function flushBCCalls() {
    var keys = Object.keys(bcf.queue).slice(0, 2000);
    if (!keys.length) return;
    var requestId = 'f' + (++state.requestSeq), batch = {};
    keys.forEach(function (k) { batch[k] = bcf.queue[k]; delete bcf.queue[k]; });
    bcf.pending[requestId] = batch;
    var calls = keys.map(function (k) { return { id: k, fn: batch[k].fn, args: batch[k].args }; });
    H.sendChunked('formulas', JSON.stringify({ requestId: requestId, calls: calls }), state.config.chunkSize).catch(function (e) {
      delete bcf.pending[requestId];
      keys.forEach(function (k) { batch[k].resolve('#BC! ' + (e.message || e)); });
    });
    if (Object.keys(bcf.queue).length) bcf.timer = setTimeout(flushBCCalls, 60);
  }

  function onFormulaResult(text) {
    var res = JSON.parse(text), batch = bcf.pending[res.requestId];
    if (!batch) return;
    delete bcf.pending[res.requestId];
    if (res.error) {
      Object.keys(batch).forEach(function (k) { batch[k].resolve('#BC! ' + res.error); });
      return;
    }
    (res.results || []).forEach(function (r) {
      var entry = batch[r.id];
      if (!entry) return;
      var value = r.error ? '#BC! ' + r.error : r.value;
      bcf.cache[r.id] = value;
      entry.resolve(value);
      delete batch[r.id];
    });
    Object.keys(batch).forEach(function (k) { batch[k].resolve('#BC! No answer'); });
  }

  /** Refresh: forget cached BC values and let the sheet calculate again, so every BC function asks AL anew. */
  function refreshBCFunctions() {
    var had = Object.keys(bcf.cache).length;
    bcf.cache = {};
    try { state.api.getFormula().executeCalculation(); } catch (e) { /* optional */ }
    return had;
  }

  // =====================================================================================
  // Stage 8: dropdown lists on editable columns, like the lookup on a BC page
  //
  // After an editable region loads, the add-in asks AL (purpose 'lookups') for the values of its editable
  // Code and Option columns: option captions, codes of the related table (Unit of Measure, Item Category,
  // dimension values...), and for Account No. / No. on journal and document lines one list per type
  // (G/L Account, Customer, Item...), chosen row by row from the type column. Dropdowns allow other values,
  // so BC still has the last word on save. Lists are kept for 5 minutes; Refresh on a region asks again.
  // =====================================================================================
  var lk = { pending: {}, cache: {} };   // pending: requestId -> { regionId, fields }; cache: tableNo -> { resp, fields, at }
  var LOOKUP_TYPES = { Code: 1, Option: 1 };
  var LOOKUP_TTL = 5 * 60 * 1000;

  function lookupFields(source) {
    var off = source.noList || [];
    return (source.columns || []).filter(function (c) { return c.editable && LOOKUP_TYPES[c.type] && off.indexOf(c.no) < 0; }).map(function (c) { return c.no; });
  }

  function requestLookups(source, force) {
    if (!isTracking(source) || !state.api || !state.api.newDataValidation) return;
    var fields = lookupFields(source);
    if (!fields.length) return;
    var hit = lk.cache[source.tableNo];
    if (!force && hit && Date.now() - hit.at < LOOKUP_TTL && fields.every(function (f) { return hit.fields.indexOf(f) >= 0; })) {
      applyLookups(source);
      return;
    }
    var requestId = H.uid('lk');
    lk.pending[requestId] = { regionId: source.id, fields: fields };
    H.sendChunked('lookups', JSON.stringify({ requestId: requestId, tableNo: source.tableNo, fields: fields }), state.config.chunkSize)
      .catch(function () { delete lk.pending[requestId]; });
  }

  function onLookupResult(text) {
    var resp = JSON.parse(text), p = lk.pending[resp.requestId];
    delete lk.pending[resp.requestId];
    if (resp.error) { H.toast('warning', 'Dropdown lists could not be loaded: ' + resp.error); return; }
    if (!p) return;
    lk.cache[resp.tableNo] = { resp: resp, fields: p.fields, at: Date.now() };
    sourceList().filter(function (s) { return s.tableNo === resp.tableNo && isTracking(s); }).forEach(applyLookups);
  }

  function colIndex(source, fieldNo) {
    for (var i = 0; i < source.columns.length; i++) if (source.columns[i].no === fieldNo) return i;
    return -1;
  }

  /**
   * Dropdowns stay inside the region: its rows plus the one empty line under them where a new row is typed
   * (when inserting is allowed). When a new row is typed the region grows and the dropdowns follow (diffSource).
   */
  function lookupRows(source) { return Math.max(boxRows(source) - 1, 1); }

  /** Removes dropdowns that are below the region (it shrank, or an older version put them 20 rows down). */
  function clearLookupsBelow(source, sheet, rows) {
    var a = anchorOf(source), below = Math.max((source._lkRows || 0) - rows, 0);
    if (!source._lkCleaned) below = Math.max(below, 25);
    below = Math.min(below, Math.max(0, roomBelow(source) - 1 - rows));
    if (below > 0) { try { sheet.getRange(a.row + 1 + rows, a.col, below, source.columns.length).setDataValidation(null); } catch (e) { /* none */ } }
    source._lkCleaned = true;
  }

  function listValues(resp, key) {
    var list = key && resp.lists && resp.lists[key];
    return list ? list.values.map(function (v) { return v[0]; }) : null;
  }

  function setList(range, values) {
    try {
      if (!values || !values.length) { range.setDataValidation(null); return; }
      // Arrow display (render mode 1): the cell keeps its normal look with a small arrow at the right, like an Excel list.
      range.setDataValidation(state.api.newDataValidation().requireValueInList(values, false, true).setAllowInvalid(true)
        .setOptions({ renderMode: 1 }).build());
    } catch (e) { /* dropdowns optional */ }
  }

  function applyLookups(source) {
    var hit = lk.cache[source.tableNo], sheet = workbook() && workbook().getSheetBySheetId(source.sheetId);
    if (!hit || !sheet || !isTracking(source)) return;
    var resp = hit.resp, a = anchorOf(source), rows = lookupRows(source);
    source.lookupTypeCols = [];
    source.lookupCount = 0;
    applying(function () {
      clearLookupsBelow(source, sheet, rows);
      source._lkRows = rows;
      // Columns whose Dropdown tick is off in the region panel: no list.
      (source.noList || []).forEach(function (no) {
        var ci = colIndex(source, no);
        if (ci >= 0) setList(sheet.getRange(a.row + 1, a.col + ci, rows, 1), null);
      });
      (resp.lookups || []).forEach(function (l) {
        if ((source.noList || []).indexOf(l.field) >= 0) return;
        var ci = colIndex(source, l.field);
        if (ci < 0 || !source.columns[ci].editable) return;
        if (l.kind === 'byType') {
          // One list per run of rows with the same type (Account Type, Type).
          var ti = colIndex(source, l.typeField);
          if (ti < 0) return;
          source.lookupTypeCols.push(ti);
          var types = sheet.getRange(a.row + 1, a.col + ti, rows, 1).getDisplayValues().map(function (r) { return String(r[0] == null ? '' : r[0]).trim(); });
          var start = 0;
          for (var i = 1; i <= rows; i++) {
            if (i === rows || types[i] !== types[start]) {
              setList(sheet.getRange(a.row + 1 + start, a.col + ci, i - start, 1), listValues(resp, l.map[types[start] || l.blankType]));
              start = i;
            }
          }
        } else {
          setList(sheet.getRange(a.row + 1, a.col + ci, rows, 1), l.values || listValues(resp, l.list));
        }
        source.lookupCount++;
      });
    });
  }

  function scheduleLookups(source) {
    clearTimeout(source._lkTimer);
    source._lkTimer = setTimeout(guard(function () { applyLookups(source); }), 300);
  }

  // =====================================================================================
  // Stage 7: Excel-style charts and KPI cards on the sheet
  //
  // Charts are Univer floating objects (sheets-drawing): they sit over the cells, can be dragged and
  // resized, and are saved in the workbook snapshot with their settings in `data`:
  //   { kind: 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'kpi', sheetId, r, c, rows, cols,
  //     title, labels: bool, legend: bool, cmp: { r, c } (KPI), good: 'up' | 'down' (KPI) }
  // The add-in draws each chart as SVG from the live cell values and redraws after every change
  // (formula results included, so BC functions and regions feed the charts).
  // =====================================================================================
  var XL_PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E'];
  var CHART_TYPES = [
    ['column', 'Clustered Column'], ['bar', 'Clustered Bar'], ['line', 'Line with Markers'],
    ['area', 'Area'], ['pie', 'Pie'], ['doughnut', 'Doughnut'], ['kpi', 'KPI Card']
  ];
  var charts = { mounted: {} };   // floatDomId -> { el, data }

  function registerChartComponent() {
    var R = window.AKUniver && window.AKUniver.React;
    if (!R || !state.api.registerComponent) return;
    var Chart = function (props) {
      var ref = R.useRef(null);
      var data = props.data || {};
      R.useEffect(function () {
        var el = ref.current;
        if (!el) return undefined;
        charts.mounted[props.floatDomId] = { el: el, data: data, id: props.floatDomId };
        drawChartSafe(props.floatDomId);
        var ro = window.ResizeObserver ? new ResizeObserver(function () { drawChartSafe(props.floatDomId); }) : null;
        if (ro) ro.observe(el);
        return function () { if (ro) ro.disconnect(); delete charts.mounted[props.floatDomId]; };
      }, [props.floatDomId, JSON.stringify(data)]);
      return R.createElement('div', { ref: ref, className: 'ak-xchart', 'data-chart': props.floatDomId });
    };
    try { state.api.registerComponent('AKChart', Chart); } catch (e) { /* charts optional */ }
    // The small toolbar on a selected chart. Pointer events on it must not start a drag of the chart.
    ['pointerdown', 'mousedown'].forEach(function (type) {
      document.addEventListener(type, function (e) { if (e.target.closest && e.target.closest('.ak-xchart__tools, .ak-xchart__rename')) e.stopPropagation(); }, true);
    });
    document.addEventListener('click', function (e) {
      var btn = e.target.closest && e.target.closest('[data-chart-act]');
      if (!btn) return;
      e.preventDefault(); e.stopPropagation();
      guard(function () { onChartTool(btn.getAttribute('data-chart-act'), btn.getAttribute('data-id')); })();
    }, true);
  }

  function drawChartSafe(id) {
    var m = charts.mounted[id];
    if (!m || !workbook()) return;
    try { m.el.innerHTML = chartMarkup(m.data, m.el.clientWidth || 480, m.el.clientHeight || 300, id); }
    catch (e) { m.el.innerHTML = '<div class="ak-xchart__err">Chart error: ' + H.escapeHtml(e.message) + '</div>'; }
  }

  var redrawCharts = H.debounce(function () { Object.keys(charts.mounted).forEach(drawChartSafe); }, 250);

  // ---------------------------------------------------------------- reading the data block
  function num(v) {
    if (typeof v === 'number') return v;
    if (v == null) return NaN;
    var t = String(v).trim();
    if (!/^[\s\-−(]*[₹$€]?\s*[\d.,]+\)?\s*%?$/.test(t)) return NaN;
    var n = +t.replace(/[^0-9.]/g, '');
    return /^[\s]*[-−(]/.test(t) ? -n : n;
  }

  /**
   * Block -> { categories[], series[{ name, values[] }] }, read the way Excel reads a chart range:
   * the first row holds series names when it is text; the label column is the first text column (in a BC
   * region with No. and Description, Description becomes the label and No. is skipped); a column is a
   * series only when its cells show numbers, so dates, Yes/No and code columns are left out.
   */
  function chartSeries(d) {
    var sheet = workbook().getSheetBySheetId(d.sheetId);
    if (!sheet) throw new Error('The sheet with the chart data was deleted.');
    var range = sheet.getRange(d.r, d.c, d.rows, d.cols);
    var block = range.getRawValues(), shown = range.getDisplayValues ? range.getDisplayValues() : block;
    var isText = function (v) { return v !== '' && v != null && isNaN(num(v)); };
    var hasHeader = d.rows > 1 && block[0].some(function (v, i) { return (i > 0 || d.cols === 1) && isText(v); });
    var start = hasHeader ? 1 : 0, body = block.slice(start), bodyShown = shown.slice(start);
    var textCol = function (c) { return bodyShown.some(function (r) { return isText(r[c]); }) && !bodyShown.every(function (r) { return /^(true|false|yes|no)?$/i.test(String(r[c] == null ? '' : r[c]).trim()); }); };
    var labelCol = -1;
    for (var lc = 0; lc < Math.min(d.cols - 1, 3); lc++) if (textCol(lc)) { labelCol = lc; break; }
    var series = [];
    for (var c = 0; c < d.cols; c++) {
      if (c <= labelCol) continue;
      // Every filled cell must show a number (a date shows "02-Oct-2026", a Boolean "FALSE").
      var cells = bodyShown.map(function (r) { return r[c]; }).filter(function (v) { return v !== '' && v != null; });
      if (!cells.length || cells.some(function (v) { return isNaN(num(v)); })) continue;
      series.push({ name: hasHeader ? String(block[0][c] == null ? '' : block[0][c]) : 'Series ' + (series.length + 1), values: body.map(function (r) { return num(r[c]); }) });
    }
    var categories = body.map(function (r, i) { return labelCol >= 0 ? String(r[labelCol] == null ? '' : r[labelCol]) : String(i + 1); });
    return { categories: categories, series: series.slice(0, 8), hasHeader: hasHeader };
  }

  // ---------------------------------------------------------------- number formats and axis
  function fmtAxis(n, step) {
    var a = Math.abs(n), s = n < 0 ? '-' : '';
    if (a >= 1e7) return s + trim(a / 1e7, step / 1e7) + ' Cr';
    if (a >= 1e5) return s + trim(a / 1e5, step / 1e5) + ' L';
    if (a >= 1e3) return s + trim(a / 1e3, step / 1e3) + 'K';
    return s + trim(a, step);
  }
  function trim(v, step) { var dec = step < 1 ? Math.min(2, Math.ceil(-Math.log10(step))) : 0; return v.toFixed(dec).replace(/\.0+$/, ''); }
  function fmtFull(n) {
    if (!isFinite(n)) return '—';
    var dec = Math.round(n) === n || Math.abs(n) >= 1000 ? 0 : 2;
    try { return n.toLocaleString('en-IN', { maximumFractionDigits: dec, minimumFractionDigits: dec }); } catch (e) { return n.toFixed(dec); }
  }
  /** Excel-like "nice" axis: 4-6 steps of 1, 2 or 5 x 10^n, always including zero. */
  function niceAxis(min, max) {
    min = Math.min(0, min); max = Math.max(0, max);
    if (min === max) max = min + 1;
    var raw = (max - min) / 5, mag = Math.pow(10, Math.floor(Math.log10(raw))), f = raw / mag;
    var step = (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * mag;
    return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step: step };
  }
  function esc(t) { return H.escapeHtml(t == null ? '' : String(t)); }
  function clip(t, n) { t = String(t); return t.length > n ? t.slice(0, n - 1) + '…' : t; }

  // ---------------------------------------------------------------- drawing
  function chartMarkup(d, W, Hh, id) {
    var tools = '<div class="ak-xchart__tools">' +
      '<button data-chart-act="type" data-id="' + id + '" title="Change chart type">' + H.icon('table') + '</button>' +
      (d.kind !== 'kpi' ? '<button data-chart-act="labels" data-id="' + id + '" title="Data labels on/off">123</button>' : '<button data-chart-act="good" data-id="' + id + '" title="Higher is ' + (d.good === 'down' ? 'worse' : 'better') + '">' + (d.good === 'down' ? '↓' : '↑') + '</button>') +
      '<button data-chart-act="data" data-id="' + id + '" title="Use the selected cells as the chart data">' + H.icon('sliders') + '</button>' +
      '<button data-chart-act="title" data-id="' + id + '" title="Rename">' + H.icon('pencil') + '</button>' +
      '<button data-chart-act="delete" data-id="' + id + '" title="Delete chart">' + H.icon('x') + '</button></div>';
    if (d.kind === 'kpi') return kpiMarkup(d, W, Hh) + tools;
    var data = chartSeries(d);
    var title = d.title || '';
    var top = title ? 34 : 12, legend = d.legend === false || (data.series.length < 2 && d.kind !== 'pie' && d.kind !== 'doughnut') ? null : legendLayout(d, data, W);
    var legendH = legend ? legend.rows * 16 + 10 : 0;
    var svg = '<svg class="ak-xchart__svg" viewBox="0 0 ' + W + ' ' + Hh + '" width="' + W + '" height="' + Hh + '" font-family="Calibri, \'Segoe UI\', Arial, sans-serif">' +
      (title ? '<text x="' + (W / 2) + '" y="22" text-anchor="middle" font-size="14" fill="#595959">' + esc(clip(title, 60)) + '</text>' : '');
    var plot = { x: 12, y: top, w: W - 24, h: Hh - top - legendH - 8 };
    if (d.kind === 'pie' || d.kind === 'doughnut') svg += pieSvg(d, data, plot);
    else svg += axisSvg(d, data, plot);
    if (legend) svg += legendSvg(legend, Hh - legendH + 4);
    return svg + '</svg>' + tools;
  }

  /** Legend items centred under the plot, wrapped onto more lines when they do not fit (as Excel does). */
  function legendLayout(d, data, W) {
    var items = d.kind === 'pie' || d.kind === 'doughnut'
      ? data.categories.slice(0, 12).map(function (c, i) { return { name: c, color: XL_PALETTE[i % 8] }; })
      : data.series.map(function (s, i) { return { name: s.name, color: XL_PALETTE[i % 8] }; });
    var lines = [[]], used = 0, max = W - 16;
    items.forEach(function (it) {
      it.w = 22 + Math.min(clip(it.name, 18).length * 6, 112);
      if (used + it.w > max && lines[lines.length - 1].length) { lines.push([]); used = 0; }
      lines[lines.length - 1].push(it); used += it.w;
    });
    lines = lines.slice(0, 3);
    return { lines: lines, rows: lines.length, W: W };
  }

  function legendSvg(lg, y) {
    var out = '';
    lg.lines.forEach(function (line, li) {
      var total = line.reduce(function (t, it) { return t + it.w; }, 0), x = Math.max(8, (lg.W - total) / 2), yy = y + li * 16;
      line.forEach(function (it) {
        out += '<rect x="' + x + '" y="' + (yy + 2) + '" width="9" height="9" fill="' + it.color + '"/>' +
          '<text x="' + (x + 13) + '" y="' + (yy + 10) + '" font-size="11" fill="#595959">' + esc(clip(it.name, 18)) + '</text>';
        x += it.w;
      });
    });
    return out;
  }

  function axisSvg(d, data, p) {
    var horiz = d.kind === 'bar', all = [];
    data.series.forEach(function (s) { s.values.forEach(function (v) { if (isFinite(v)) all.push(v); }); });
    var ax = niceAxis(Math.min.apply(null, all.concat([0])), Math.max.apply(null, all.concat([0])));
    var ticks = []; for (var t = ax.min; t <= ax.max + ax.step / 2; t += ax.step) ticks.push(t);
    var labelW = Math.max.apply(null, ticks.map(function (v) { return fmtAxis(v, ax.step).length; })) * 6.5 + 8;
    var catW = horiz ? Math.min(110, Math.max.apply(null, data.categories.map(function (c) { return clip(c, 16).length; }).concat([2])) * 6.2 + 10) : 0;
    var x0 = p.x + (horiz ? catW : labelW), y0 = p.y, w = p.w - (horiz ? catW : labelW) - (horiz && d.labels ? 34 : 0), h = p.h - (horiz ? 18 : 20);
    var n = data.categories.length, out = '';
    var val = function (v) { return (v - ax.min) / (ax.max - ax.min); };
    // gridlines and value labels
    ticks.forEach(function (tv) {
      if (horiz) {
        var x = x0 + val(tv) * w;
        out += '<line x1="' + x + '" x2="' + x + '" y1="' + y0 + '" y2="' + (y0 + h) + '" stroke="#D9D9D9" stroke-width="0.75"/>' +
          '<text x="' + x + '" y="' + (y0 + h + 13) + '" text-anchor="middle" font-size="10" fill="#595959">' + fmtAxis(tv, ax.step) + '</text>';
      } else {
        var y = y0 + h - val(tv) * h;
        out += '<line x1="' + x0 + '" x2="' + (x0 + w) + '" y1="' + y + '" y2="' + y + '" stroke="#D9D9D9" stroke-width="0.75"/>' +
          '<text x="' + (x0 - 6) + '" y="' + (y + 3.5) + '" text-anchor="end" font-size="10" fill="#595959">' + fmtAxis(tv, ax.step) + '</text>';
      }
    });
    var zero = horiz ? x0 + val(0) * w : y0 + h - val(0) * h;
    out += horiz ? '<line x1="' + zero + '" x2="' + zero + '" y1="' + y0 + '" y2="' + (y0 + h) + '" stroke="#BFBFBF"/>'
      : '<line x1="' + x0 + '" x2="' + (x0 + w) + '" y1="' + zero + '" y2="' + zero + '" stroke="#BFBFBF"/>';
    var band = (horiz ? h : w) / Math.max(n, 1), S = data.series.length;
    // category labels
    var every = horiz ? 1 : Math.ceil(n / Math.max(1, Math.floor(w / 46)));
    data.categories.forEach(function (c, i) {
      if (i % every) return;
      if (horiz) out += '<text x="' + (x0 - 6) + '" y="' + (y0 + band * (i + 0.5) + 3.5) + '" text-anchor="end" font-size="10" fill="#595959">' + esc(clip(c, 16)) + '</text>';
      else out += '<text x="' + (x0 + band * (i + 0.5)) + '" y="' + (y0 + h + 14) + '" text-anchor="middle" font-size="10" fill="#595959">' + esc(clip(c, Math.max(4, Math.floor(band * every / 6.2)))) + '</text>';
    });
    var label = function (x, y, v, anchor) {
      return d.labels ? '<text x="' + x + '" y="' + y + '" text-anchor="' + (anchor || 'middle') + '" font-size="9.5" fill="#404040">' + fmtAxis(v, ax.step / 10) + '</text>' : '';
    };
    if (d.kind === 'line' || d.kind === 'area') {
      data.series.forEach(function (s, si) {
        var color = XL_PALETTE[si % 8], pts = [];
        s.values.forEach(function (v, i) { if (isFinite(v)) pts.push([x0 + band * (i + 0.5), y0 + h - val(v) * h, v]); });
        if (!pts.length) return;
        if (d.kind === 'area') {
          out += '<path d="M' + pts[0][0] + ',' + zero + ' L' + pts.map(function (q) { return q[0] + ',' + q[1]; }).join(' L') + ' L' + pts[pts.length - 1][0] + ',' + zero + ' Z" fill="' + color + '" fill-opacity="' + (S > 1 ? 0.45 : 0.75) + '"/>';
        } else {
          out += '<polyline fill="none" stroke="' + color + '" stroke-width="2.25" stroke-linejoin="round" points="' + pts.map(function (q) { return q[0] + ',' + q[1]; }).join(' ') + '"/>';
          pts.forEach(function (q, k) { out += '<circle cx="' + q[0] + '" cy="' + q[1] + '" r="3.2" fill="' + color + '" stroke="#fff" stroke-width="1"><title>' + esc(data.categories[k] + ' · ' + s.name + ': ' + fmtFull(q[2])) + '</title></circle>'; });
        }
        pts.forEach(function (q) { out += label(q[0], q[1] - 7, q[2]); });
      });
    } else {
      var gap = band * 0.25, bw = (band - gap) / Math.max(S, 1);
      data.series.forEach(function (s, si) {
        var color = XL_PALETTE[si % 8];
        s.values.forEach(function (v, i) {
          if (!isFinite(v)) return;
          var a = val(Math.max(v, 0)), b = val(Math.min(v, 0)), tip = '<title>' + esc(data.categories[i] + ' · ' + s.name + ': ' + fmtFull(v)) + '</title>';
          if (horiz) {
            var y = y0 + band * i + gap / 2 + bw * si, xa = x0 + b * w, xb = x0 + a * w;
            out += '<rect x="' + xa + '" y="' + y + '" width="' + Math.max(1, xb - xa) + '" height="' + Math.max(1, bw - 1) + '" fill="' + color + '">' + tip + '</rect>';
            out += label(v >= 0 ? xb + 3 : xa - 3, y + bw / 2 + 3, v, v >= 0 ? 'start' : 'end');
          } else {
            var x = x0 + band * i + gap / 2 + bw * si, ya = y0 + h - a * h, yb = y0 + h - b * h;
            out += '<rect x="' + x + '" y="' + ya + '" width="' + Math.max(1, bw - 1) + '" height="' + Math.max(1, yb - ya) + '" fill="' + color + '">' + tip + '</rect>';
            out += label(x + bw / 2, v >= 0 ? ya - 4 : yb + 11, v);
          }
        });
      });
    }
    return out;
  }

  function pieSvg(d, data, p) {
    var s = data.series[0] || { values: [] }, values = s.values.map(function (v) { return isFinite(v) ? Math.abs(v) : 0; });
    var total = values.reduce(function (t, v) { return t + v; }, 0);
    if (!total) return '<text x="' + (p.x + p.w / 2) + '" y="' + (p.y + p.h / 2) + '" text-anchor="middle" font-size="12" fill="#8A8886">No numbers to show</text>';
    var cx = p.x + p.w / 2, cy = p.y + p.h / 2, r = Math.max(10, Math.min(p.w, p.h) / 2 - 6), ri = d.kind === 'doughnut' ? r * 0.55 : 0, a0 = -Math.PI / 2, out = '';
    values.forEach(function (v, i) {
      if (!v) return;
      var a1 = a0 + v / total * Math.PI * 2, large = a1 - a0 > Math.PI ? 1 : 0, color = XL_PALETTE[i % 8];
      var pt = function (rad, ang) { return (cx + rad * Math.cos(ang)) + ',' + (cy + rad * Math.sin(ang)); };
      var path = values.filter(Boolean).length === 1
        ? 'M' + (cx - r) + ',' + cy + ' a' + r + ',' + r + ' 0 1,0 ' + (2 * r) + ',0 a' + r + ',' + r + ' 0 1,0 ' + (-2 * r) + ',0'
        : 'M' + pt(r, a0) + ' A' + r + ',' + r + ' 0 ' + large + ',1 ' + pt(r, a1) + (ri ? ' L' + pt(ri, a1) + ' A' + ri + ',' + ri + ' 0 ' + large + ',0 ' + pt(ri, a0) + ' Z' : ' L' + cx + ',' + cy + ' Z');
      out += '<path d="' + path + '" fill="' + color + '" stroke="#fff" stroke-width="1.5"><title>' + esc(data.categories[i] + ': ' + fmtFull(s.values[i]) + ' (' + (v / total * 100).toFixed(1) + '%)') + '</title></path>';
      var mid = (a0 + a1) / 2, pr = ri ? (r + ri) / 2 : r * 0.62, pct = v / total * 100;
      if (pct >= 4 && d.labels !== false) out += '<text x="' + (cx + pr * Math.cos(mid)) + '" y="' + (cy + pr * Math.sin(mid) + 4) + '" text-anchor="middle" font-size="11" font-weight="600" fill="#fff">' + Math.round(pct) + '%</text>';
      a0 = a1;
    });
    if (ri) out += '<text x="' + cx + '" y="' + (cy + 5) + '" text-anchor="middle" font-size="13" font-weight="600" fill="#404040">' + fmtAxis(total, total / 100) + '</text>';
    return out;
  }

  function kpiMarkup(d, W, Hh) {
    var sheet = workbook().getSheetBySheetId(d.sheetId);
    var v = num(sheet.getRange(d.r, d.c, 1, 1).getRawValues()[0][0]), delta = '';
    if (d.cmp) {
      var base = num(sheet.getRange(d.cmp.r, d.cmp.c, 1, 1).getRawValues()[0][0]);
      if (isFinite(v) && isFinite(base) && base !== 0) {
        var diff = v - base, pct = diff / Math.abs(base) * 100, good = d.good === 'down' ? diff <= 0 : diff >= 0;
        delta = '<div class="ak-xkpi__delta ' + (good ? 'is-good' : 'is-bad') + '">' + (diff >= 0 ? '▲ ' : '▼ ') + Math.abs(pct).toFixed(1) + '%' +
          '<span> vs ' + esc(d.cmpTitle || 'target') + ' ' + fmtAxis(base, Math.abs(base) / 100) + '</span></div>';
      }
    }
    var size = Math.max(22, Math.min(44, Math.min(W / 7, Hh / 3)));
    return '<div class="ak-xkpi"><div class="ak-xkpi__title">' + esc(d.title || 'KPI') + '</div>' +
      '<div class="ak-xkpi__value" style="font-size:' + size + 'px">' + fmtFull(v) + '</div>' + delta + '</div>';
  }

  // ---------------------------------------------------------------- inserting
  function selectionRect() {
    var wb = workbook(), sheet = wb.getActiveSheet(), sel = sheet.getSelection && sheet.getSelection();
    var range = sel && sel.getActiveRange && sel.getActiveRange();
    var g = range && range.getRange ? range.getRange() : { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
    return { sheet: sheet, sheetId: sheet.getSheetId(), r: g.startRow, c: g.startColumn, rows: g.endRow - g.startRow + 1, cols: g.endColumn - g.startColumn + 1 };
  }

  function guessChartTitle(b) {
    var above = b.r > 0 ? cellText(b.sheetId, b.r - 1, b.c) : '';
    if (above && isNaN(num(above))) return above;
    // Like Excel: one series -> its name is the title; several series -> "Chart Title" (rename with the pencil).
    var head = cellText(b.sheetId, b.r, b.c + b.cols - 1);
    var labelCol = b.cols > 1 && isNaN(num(cellText(b.sheetId, b.r + 1, b.c)));
    if (b.cols - (labelCol ? 1 : 0) === 1 && head && isNaN(num(head))) return head;
    return 'Chart Title';
  }

  function cellText(sheetId, r, c) {
    var sheet = workbook().getSheetBySheetId(sheetId);
    if (!sheet || r < 0 || c < 0) return '';
    var v = sheet.getRange(r, c, 1, 1).getValue();
    return v == null ? '' : String(v).trim();
  }

  /** Position of a cell's left/top edge in drawing coordinates (these include the row and column headers). */
  function cellOffset(sheet, r, c) {
    var cfg = {}, x, y, i;
    try { cfg = sheet.getSheet().getConfig() || {}; } catch (e) { /* defaults */ }
    x = cfg.rowHeader && !cfg.rowHeader.hidden ? cfg.rowHeader.width || 46 : 0;
    y = cfg.columnHeader && !cfg.columnHeader.hidden ? cfg.columnHeader.height || 20 : 0;
    for (i = 0; i < c; i++) x += (sheet.getColumnWidth ? sheet.getColumnWidth(i) : 88);
    for (i = 0; i < r; i++) y += (sheet.getRowHeight ? sheet.getRowHeight(i) : 24);
    return { x: x, y: y };
  }

  /** Excel's Insert Chart dialog: one tile per chart type, each a live preview of the selected data. */
  function openChartPicker() {
    var b = selectionRect();
    var inRegion = b.rows === 1 && b.cols === 1 && regionAt(b.sheetId, b.r, b.c);
    if (inRegion) {
      // Like Excel: one cell inside a table means "chart the whole table".
      var g = regionRange(inRegion, 'all');
      b = { sheet: b.sheet, sheetId: b.sheetId, r: g.r, c: g.c, rows: g.rows, cols: g.cols };
      try { b.sheet.getRange(g.r, g.c, g.rows, g.cols).activate(); } catch (e) { /* optional */ }
    }
    if (b.rows === 1 && b.cols === 1 && isNaN(num(b.sheet.getRange(b.r, b.c, 1, 1).getRawValues()[0][0]))) {
      H.toast('info', 'Select the data first: labels in the first column, series names in the first row and numbers in the rest (for example A4:D7), or one number cell for a KPI card.');
      return;
    }
    var single = b.rows * b.cols <= 2;
    var base = { sheetId: b.sheetId, r: b.r, c: b.c, rows: b.rows, cols: b.cols, title: guessChartTitle(b), labels: false, legend: true };
    var tiles = CHART_TYPES.filter(function (t) { return single ? t[0] === 'kpi' || t[0] === 'column' : t[0] !== 'kpi' || b.rows * b.cols <= 2; }).map(function (t) {
      var d = Object.assign({}, base, { kind: t[0] });
      if (t[0] === 'kpi') prepareKpi(d, b);
      var html;
      try { html = chartMarkup(d, 220, 140, 'preview').replace(/<div class="ak-xchart__tools">[\s\S]*$/, ''); } catch (e) { html = '<div class="ak-xchart__err">' + esc(e.message) + '</div>'; }
      return '<button class="ak-picker__tile" data-kind="' + t[0] + '"><span class="ak-picker__prev">' + html + '</span><span class="ak-picker__name">' + t[1] + '</span></button>';
    }).join('');
    var host = document.getElementById('ak-app');
    var back = H.el('div', 'ak-modal');
    back.innerHTML = '<div class="ak-modal__box ak-picker" role="dialog" aria-label="Insert Chart"><div class="ak-modal__title">Insert Chart</div>' +
      '<div class="ak-picker__sub">Data: ' + esc(sheetName(b.sheetId) + '!' + cellName(b.r, b.c) + ':' + cellName(b.r + b.rows - 1, b.c + b.cols - 1)) + '</div>' +
      '<div class="ak-picker__grid">' + tiles + '</div><div class="ak-modal__btns"><button class="ak-btn ak-btn--subtle" data-i="cancel">Cancel</button></div></div>';
    var close = function () { if (back.parentNode) back.parentNode.removeChild(back); document.removeEventListener('keydown', onKey, true); };
    var onKey = function (e) { if (e.key === 'Escape') close(); };
    back.addEventListener('click', function (e) {
      var tile = e.target.closest('.ak-picker__tile');
      if (tile) { close(); insertChart(base, tile.getAttribute('data-kind'), b); return; }
      if (e.target === back || e.target.closest('[data-i=cancel]')) close();
    });
    document.addEventListener('keydown', onKey, true);
    host.appendChild(back);
  }

  function prepareKpi(d, b) {
    d.rows = 1; d.cols = 1; d.good = 'up';
    var left = cellText(b.sheetId, b.r, b.c - 1), above = cellText(b.sheetId, b.r - 1, b.c);
    d.title = left && isNaN(num(left)) ? left : (above && isNaN(num(above)) ? above : 'KPI');
    if (b.cols >= 2 || b.rows >= 2) {
      d.cmp = b.cols >= 2 ? { r: b.r, c: b.c + 1 } : { r: b.r + 1, c: b.c };
      var head = '';
      for (var r = d.cmp.r - 1; r >= 0 && r >= d.cmp.r - 30 && !head; r--) { var t = cellText(b.sheetId, r, d.cmp.c); if (t && isNaN(num(t))) head = t; }
      d.cmpTitle = head || 'target';
    }
  }

  function insertChart(base, kind, b) {
    var d = Object.assign({}, base, { kind: kind });
    if (kind === 'kpi') prepareKpi(d, b);
    var sheet = b.sheet, at = cellOffset(sheet, b.r, b.c + b.cols + 1);
    var w = kind === 'kpi' ? 260 : 480, h = kind === 'kpi' ? 140 : 290;
    at = freeSpot(sheet, at, w, h);
    var res = null;
    try {
      res = sheet.addFloatDomToPosition({ componentKey: 'AKChart', allowTransform: true, data: d,
        initPosition: { startX: at.x, endX: at.x + w, startY: at.y, endY: at.y + h } });
    } catch (e) { reportError(e); return; }
    if (!res) { H.toast('error', 'The chart could not be placed on this sheet.'); return; }
    setDirty(true);
    H.toast('success', (kind === 'kpi' ? 'KPI card' : 'Chart') + ' inserted. Drag it to move, drag a corner to resize; it follows the cells.');
  }

  /** Like Excel, a new chart goes next to its data; if another chart is already there, place it beside (or below) that one. */
  function freeSpot(sheet, at, w, h) {
    var boxes = Object.keys(charts.mounted).map(function (id) {
      try { var f = sheet.getFloatDomById(id); return f && f.position; } catch (e) { return null; }
    }).filter(Boolean);
    for (var guard = 0; guard < 20; guard++) {
      var hit = boxes.filter(function (q) {
        var L = q.startX != null ? q.startX : q.left, T = q.startY != null ? q.startY : q.top;
        var R = q.endX != null ? q.endX : L + q.width, B = q.endY != null ? q.endY : T + q.height;
        return at.x < R && at.x + w > L && at.y < B && at.y + h > T;
      })[0];
      if (!hit) break;
      var hitR = hit.endX != null ? hit.endX : (hit.startX != null ? hit.startX : hit.left) + hit.width;
      var hitB = hit.endY != null ? hit.endY : (hit.startY != null ? hit.startY : hit.top) + hit.height;
      at = hitR + 12 + w < 2000 ? { x: hitR + 12, y: at.y } : { x: at.x, y: hitB + 12 };
    }
    return at;
  }

  function chartSheet(id) {
    var wb = workbook();
    return wb.getSheets().filter(function (s) { try { return !!s.getFloatDomById(id); } catch (e) { return false; } })[0] || null;
  }

  function updateChartData(id, changes) {
    var sheet = chartSheet(id), m = charts.mounted[id];
    if (!sheet || !m) return;
    var d = Object.assign({}, m.data, changes);
    m.data = d;
    try { sheet.updateFloatDom(id, { data: d }); } catch (e) { /* keep local copy */ }
    drawChartSafe(id);
    setDirty(true);
  }

  function onChartTool(act, id) {
    var m = charts.mounted[id];
    if (!m) return;
    if (act === 'delete') {
      var sheet = chartSheet(id);
      if (sheet) { sheet.removeFloatDom(id); setDirty(true); }
      return;
    }
    if (act === 'labels') updateChartData(id, { labels: !m.data.labels });
    if (act === 'good') updateChartData(id, { good: m.data.good === 'down' ? 'up' : 'down' });
    if (act === 'type') {
      var order = CHART_TYPES.map(function (t) { return t[0]; }).filter(function (k) { return k !== 'kpi'; });
      if (m.data.kind === 'kpi') return;
      updateChartData(id, { kind: order[(order.indexOf(m.data.kind) + 1) % order.length] });
      H.toast('info', 'Chart type: ' + CHART_TYPES.filter(function (t) { return t[0] === charts.mounted[id].data.kind; })[0][1]);
    }
    if (act === 'data') {
      var b = selectionRect();
      if (b.rows * b.cols < 2 && m.data.kind !== 'kpi') { H.toast('info', 'Select the new data block on the sheet first, then click this button.'); return; }
      updateChartData(id, m.data.kind === 'kpi' ? { sheetId: b.sheetId, r: b.r, c: b.c } : { sheetId: b.sheetId, r: b.r, c: b.c, rows: b.rows, cols: b.cols });
      H.toast('success', 'Chart data: ' + sheetName(b.sheetId) + '!' + cellName(b.r, b.c) + ':' + cellName(b.r + b.rows - 1, b.c + b.cols - 1));
    }
    if (act === 'title') {
      var el = m.el;
      var current = m.data.title || '';
      var box = H.el('div', 'ak-xchart__rename');
      box.innerHTML = '<input class="ak-input ak-input--sm" maxlength="80" value="' + esc(current) + '" aria-label="Chart title">';
      el.appendChild(box);
      var input = box.querySelector('input');
      input.focus(); input.select();
      var done = function (save) { if (save) updateChartData(id, { title: input.value.trim() }); if (box.parentNode) box.parentNode.removeChild(box); };
      input.addEventListener('keydown', function (e) { e.stopPropagation(); if (e.key === 'Enter') done(true); if (e.key === 'Escape') done(false); });
      input.addEventListener('blur', function () { done(true); });
    }
  }

  /** Workbooks saved with the 2.0 dashboard strip: put each old card on its sheet as a floating chart. */
  function migrateLegacyDash(dash) {
    if (!dash || !dash.items || !dash.items.length) return;
    dash.items.forEach(function (it, i) {
      var wb = workbook(), sheet = wb.getSheetBySheetId(it.sheetId);
      if (!sheet) return;
      var d = { kind: it.type === 'kpi' ? 'kpi' : (it.kind === 'line' ? 'line' : 'column'), sheetId: it.sheetId, r: it.r, c: it.c, rows: it.rows, cols: it.cols,
        title: it.title, labels: false, legend: true, good: it.good, cmp: it.cmp, cmpTitle: it.cmpTitle };
      var at = cellOffset(sheet, it.r, it.c + it.cols + 1), w = d.kind === 'kpi' ? 260 : 480, h = d.kind === 'kpi' ? 140 : 290;
      try { sheet.addFloatDomToPosition({ componentKey: 'AKChart', allowTransform: true, data: d, initPosition: { startX: at.x, endX: at.x + w, startY: at.y + i * 20, endY: at.y + i * 20 + h } }); }
      catch (e) { /* skip */ }
    });
  }

  // =====================================================================================
  // Combine (Phase 2): join two sheets in the browser
  // =====================================================================================
  /** Reads a sheet as { name, header[], types[], rows[][] } using raw (unformatted) values. */
  function readSheetTable(sheetId) {
    var wb = workbook(), sheet = wb.getSheetBySheetId(sheetId);
    if (!sheet) throw new Error('A sheet used by this combine no longer exists.');
    var name = sheet.getSheetName();
    var region = srcOf(sheetId);
    if (region && region.lastRows) {
      // A sheet with a BC region: Combine and Reconcile use that region's block (without the BC Status column).
      var ra = anchorOf(region), cols = region.dataCols || region.fields.length;
      var block = sheet.getRange(ra.row, ra.col, Math.max(regionEndRow(region) - ra.row + 1, 1), cols).getRawValues();
      var rh = block[0].map(function (h, i) { return h == null || h === '' ? 'Column ' + H.colLetter(ra.col + i) : String(h); });
      var rrows = block.slice(1).filter(function (r) { return r.some(function (v) { return v != null && v !== ''; }); });
      return { name: name, header: rh, types: columnTypes(sheetId, rh, rrows), rows: rrows, headerRow: 0 };
    }
    var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
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
    return !!(srcOf(sheetId) || state.joins[sheetId] || state.recons[sheetId]);
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
    var src = srcOf(sheetId), jn = state.joins[sheetId];
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
      var src = srcOf(id);
      var tag = src ? '  \u00b7  ' + src.tableCaption : state.joins[id] ? '  \u00b7  combined' : state.recons[id] ? '  \u00b7  recon' : '';
      if (src && src.tableCaption === s.getSheetName()) tag = '  \u00b7  BC';
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
      var src = id && srcOf(id);
      return norm(src ? src.tableCaption : id ? sheetName(id) : '');
    };
    var pkOf = function (id) {
      var src = id && srcOf(id);
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
    var bc = sheets.filter(function (s) { return srcOf(s.getSheetId()); });
    var other = sheets.filter(function (s) { var id = s.getSheetId(); return !srcOf(id) && !state.joins[id] && !state.recons[id]; });
    var left = null;
    (preset && preset.leftSheet || []).some(function (pat) {
      left = bc.filter(function (s) { return norm(srcOf(s.getSheetId()).tableCaption).indexOf(pat) >= 0; })[0] || null;
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
      recons: recons,
      slicers: Object.keys(state.slicers).map(function (k) { return state.slicers[k]; })
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
    (env.sources || []).forEach(function (s) {
      // Workbooks saved before regions: one source per sheet, at A1.
      if (!s.id) s.id = 'rg_' + s.sheetId;
      if (!s.anchor) s.anchor = { row: 0, col: 0 };
      state.sources[s.id] = s;
    });
    // Editable sheets: rebuild the row positions and change marks from the saved copy of the loaded rows.
    setTimeout(function () { editableList().forEach(function (s) { s.marks = []; diffSource(s); requestLookups(s); }); updateWriteButton(); updateRegionBar(); }, 600);
    (env.joins || []).forEach(function (j) { state.joins[j.sheetId] = j; });
    state.slicers = {};
    (env.slicers || []).forEach(function (p) { state.slicers[p.id] = p; });
    if (env.dash && env.dash.items && env.dash.items.length) setTimeout(guard(function () { migrateLegacyDash(env.dash); setDirty(true); }), 700);
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
    renderQuickStarts();
    H.invoke('AddInFrameReady', []).then(function () {
      // AL sends the saved workbook during AddInFrameReady; apply the page's region after it is in place.
      if (state.config.autoRegion) setTimeout(guard(function () { applyAutoRegion(state.config.autoRegion); }), 900);
    });
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
    else if (purpose === 'budgetdata') onBudgetData(full);
    else if (purpose === 'budgetresult') onBudgetResult(full);
    else if (purpose === 'formularesult') onFormulaResult(full);
    else if (purpose === 'lookupresult') onLookupResult(full);
  });

  window.RequestSave = guard(function () { requestSave(); });

  window.RequestExport = guard(function () { requestExport(); });

  window.ShowToast = guard(function (level, message) { H.toast(level, message); });

  // Exposed for automated tests only.
  window.AKWorkbookDebug = { replaceWorkbook: replaceWorkbook, budgetDiff: budgetDiff, state: state, diffSource: diffSource, saveToBC: saveToBC, updateRegionBar: updateRegionBar, buildXlsx: buildXlsx, runJoin: runJoin, runRecon: runRecon, importXlsxBuffer: importXlsxBuffer };

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
