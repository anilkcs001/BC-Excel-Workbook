# AK Excel Workbook

## 1. Extension name and purpose

**AK Excel Workbook** (Per-Tenant Extension, runtime 16 / Business Central v27, IN localization compatible)

It puts a full spreadsheet inside Business Central, on a page that is **only the spreadsheet**. Users can:

- place **Business Central regions** anywhere on a sheet (Insert region, or right-click a cell): a region is a live table from **any table they are allowed to read**, with chosen fields and BC-style filters. One sheet can hold many regions next to their own titles, notes and formulas, and a filter can take its value from a cell (`[B1]`), like a slicer;
- work with it like Excel: formulas (400+ functions), formatting, filters, sort, conditional formatting, find and replace, data validation, freeze panes, copy/paste with desktop Excel;
- **refresh** every linked sheet with one click (formulas next to the data survive);
- **save** (or let **AutoSave** do it), **export** to .xlsx, and **import** .xlsx either as new sheets or as a replacement workbook;
- **combine** two sheets on matching keys (Phase 2);
- **reconcile** books against a bank statement, GSTR-2B or a party ledger (Phase 3), with an exceptions-first result sheet and a live summary sheet;
- **work on journal and document lines** (Stage 2): *Edit in Excel Workbook* on the General Journal, Item Journal, Sales Order and Purchase Order pages opens the lines as a region; type or paste many lines and BC fills the template, batch or document, the next Line No. and the journal defaults, like a new line on the page. The same four appear as *Quick start* tiles in the region panel;
- **open BC from any cell** (Stage 5): Ctrl+click a region row to open its record, or a budget / last-year amount to open the G/L budget entries or G/L entries behind it;
- **pick a period once for the whole sheet** (Stage 4): a *period picker* cell (This Month, Year to Date, Last Fiscal Year...) drives every region whose date filter points to it;
- **plan budgets in a grid** (Stage 3): G/L accounts down, months or quarters across, last year's actuals alongside; type amounts (or a yearly total to spread it) and **Save to BC** adds G/L Budget Entries for the differences;
- **edit BC data in the sheet and save it back** (Stage 1 write-back): change values or type new rows below the data, then **Save to BC**. Each field goes through BC's own `Validate`, and a *BC Status* column shows *Saved* or BC's error per row. Only tables an admin allows can be edited.

Every command (AutoSave, Save to BC, Save, Refresh, Import, Export, Combine, Reconcile, Insert region) sits in the add-in's own toolbar, with region commands also on the right-click menu and on a region bar that appears when a cell inside a region is selected. The page has no BC fields or action bar. The look follows Excel: Excel green accent, Calibri 11, and regions styled like an Excel table (blue header, thin border, banded rows).

Delivered: Phase 1 (sheet, data explorer, save, refresh, Excel I/O), Phase 2 (Combine), Phase 3 (Reconcile), Stage 1 of the advanced features (editable regions with write-back) Stage 1b (regions anywhere on a sheet, cell-driven filters, Editing switch, Excel look), Stage 2 (journal and document lines), Stage 3 (budget grid), Stage 4 (period picker) and Stage 5 (drill-down). Next stages: journal and document lines, budget grid, cell-driven filters (slicers), drill-down, BC formulas in cells, KPI cards and charts.

## 2. Object list

| ID | Type | Name |
|---|---|---|
| 89000 | Table | AK Workbook |
| 89001 | Table | AK Workbook Edit Table (allow-list of editable tables) |
| 89000 | Page | AK Workbook List |
| 89001 | Page | AK Workbook Editor (`UserControlHost`: the add-in fills the whole page) |
| 89000 | Codeunit | AK Workbook Data Service |
| 89002 | Page | AK Workbook Edit Tables (admin: which tables can be edited) |
| 89001 | Codeunit | AK Workbook Mgt |
| 89002 | Codeunit | AK Workbook Write Service |
| 89003 | Codeunit | AK Workbook Write Row |
| 89004 | Codeunit | AK Workbook Lines (journal/document line regions, quick starts, open from pages) |
| 89005 | Codeunit | AK Workbook Budget (budget grid load, budget setup for the panel) |
| 89006 | Codeunit | AK Workbook Budget Write (one budget cell per run) |
| 89007 | Codeunit | AK Workbook Drill Down (opens the BC page behind a cell) |
| 89000 | PageExtension | AK General Journal (action *Edit in Excel Workbook*) |
| 89001 | PageExtension | AK Item Journal (action *Edit in Excel Workbook*) |
| 89002 | PageExtension | AK Sales Order (action *Edit in Excel Workbook*) |
| 89003 | PageExtension | AK Purchase Order (action *Edit in Excel Workbook*) |
| 89000 | PermissionSet | AK Workbook |
| 89001 | PermissionSet | AK Workbook Admin (includes AK Workbook + editable-table setup) |
| — | ControlAddIn | AK Excel Workbook |

ID range in app.json: 89000–89049 (room for later phases).

## 3. Dependencies

| Dependency | Use | License |
|---|---|---|
| Base Application / System Application v27 | `Temp Blob`, `Base64 Convert`, `AllObjWithCaption`, `Field` | Microsoft |
| [Univer](https://github.com/dream-num/univer) 1.2.0 (bundled) | Spreadsheet engine and UI | Apache-2.0 |
| [ExcelJS](https://github.com/exceljs/exceljs) 4.4.0 (bundled) | .xlsx read/write in the browser | MIT |

No Azure services, no external API calls, no CDN at runtime. All scripts ship inside the .app.

**Licence.** This extension is under the [MIT licence](LICENSE). The bundled Univer and ExcelJS files keep their own licences; copies are in `ControlAddin/scripts/vendor` (`LICENSE-univer.txt`, `LICENSE-exceljs.txt`).

## 4. Setup steps

1. **Build the vendor bundle (only if you change library versions).** The ready-built files are already in `ControlAddin/scripts/vendor` and `ControlAddin/stylesheets/vendor`. To rebuild: `cd build`, `npm install`, `npm run build`.
2. **Open the folder in VS Code** with the AL Language extension. `app.json` targets runtime `16.0`, platform and application `27.0.0.0`. Set your sandbox in `.vscode/launch.json`, then run *AL: Download Symbols*.
3. **Publish** with F5 (sandbox) or build the .app and upload it in *Extension Management → Upload Extension* (production, PTE).
4. **BC Admin Center:** nothing to configure. The extension uses no outbound HTTP.
5. **Assign permissions:** give users the permission set **AK Workbook** (*Users → User Permission Sets* or a security group). Users still need their normal permissions for each source table (for example *D365 BASIC*, *D365 FIN. ACC*, GST permission sets). The workbook never bypasses them. Give administrators **AK Workbook Admin**.
6. **Allow editing (optional):** an admin opens *Excel Workbook Editable Tables* and adds the tables users may change from a sheet (for example Item, Customer, Vendor, Item Unit of Measure), with *Allow Modify* and, if wanted, *Allow Insert*. Ledger entries, posted documents, registers, archives and system tables cannot be added. Users also need their normal modify/insert permission on the table.
7. **Use it:** search *Excel Workbooks* → add a line with a code and name → **Open**. The spreadsheet fills the page. Click the workbook name at the top left to rename it.

**Typical reconciliations**

| Reconciliation | Load from BC (Get data) | Bring in (Import → Add as new sheets) | Template |
|---|---|---|---|
| GST purchases vs GSTR-2B | *Detailed GST Ledger Entry*, filter **GST Component Code** = `IGST\|CGST` so each line's taxable value counts once | GSTR-2B Excel from the GST portal (title rows and two-row headers are handled) | GST vs GSTR-2B |
| Bank reconciliation | *Bank Account Ledger Entry*, filter on the bank account and period | Bank statement Excel (separate Debit/Credit columns are supported) | Bank vs G/L |
| Vendor/customer balance confirmation | *Vendor Ledger Entry* or *Customer Ledger Entry*, filter on the party | The party's statement | Party ledger |

## 5. Key logic explanation

**Start-up handshake.** The editor page is a `UserControlHost` page, which allows almost no page triggers, so all set-up happens in the add-in's ready event. The add-in builds its UI and raises `ControlAddInReady`. AL checks the workbook exists, sets the page caption and answers with `Init(config)`. JS creates the spreadsheet and raises `AddInFrameReady`. AL then sends the saved workbook, if any.

**Reading data (codeunit 89000 "AK Workbook Data Service").**
- `GetTablesJson` searches `AllObjWithCaption` by caption, number or technical name, and returns at most 60 tables.
- `GetFieldsJson` reads the `Field` table. It skips removed, disabled, BLOB, Media and FlowFilter fields, and flags primary-key fields and FlowFields.
- `BuildDataChunks` opens the table with `RecordRef` as the current user. If the user has no read permission, it stops with a clear message. It uses `AddLoadFields` so SQL only reads the chosen columns, and applies each filter with `FieldRef.SetFilter`, so normal BC filter syntax works. FlowFields are calculated per row. Rows go to the browser in pages of 2,000. The default limit is 10,000 rows and the maximum is 100,000.
- Values keep their type: numbers stay numbers, dates become real Excel dates, Options and Enums go as their caption.

**Writing data into a sheet (main.js).** Each load is a "source" linked to one sheet: table, fields, filters and max rows. On refresh, only the cells the previous load wrote are cleared and rewritten, so formulas, notes and formatting outside that block are kept. The header gets bold styling, a frozen first row, an auto-filter, number formats and auto-fit widths.

**Large payloads.** Saved workbooks and .xlsx files can be several MB. Both directions are split into 400,000-character chunks (`OnPayloadChunk` / `ReceiveChunk`) and reassembled on the other side.

**Saving and AutoSave (codeunit 89001 "AK Workbook Mgt").** JS saves the Univer workbook snapshot together with the definitions of linked sources, combines and reconciliations. AL stores it in the `Workbook Data` BLOB, with size, saved-by and linked tables. With **AutoSave** on (Excel-style switch, remembered per user and browser), the add-in saves quietly about 30 seconds after the first unsaved change; AL stores it without a toast. BC cannot ask "close without saving?" on this page type, so AutoSave is the safety net. Other extensions should use codeunit 89001 as the facade, not the table directly.

**Excel export and import.** Both run in the browser with ExcelJS. Export carries values, formulas, fonts, fills, borders, alignment, number formats, widths, row heights, merges, frozen panes and the auto-filter; AL turns it into a download with `DownloadFromStream`. Import reads the file the user picks directly in the browser and either **adds its sheets** next to the current ones (for statements and GSTR-2B) or **replaces** the workbook.

**Combine sheets (Phase 2, main.js `runJoin`).** Joins two sheets in the browser on one or more key columns: *All main rows*, *Matched only*, *Both sheets* or *Differences*; with several lookup matches it can *Sum numbers*, take the *First match* or repeat for *Every match*. The key is suggested from the headers (Customer `No.` = Cust. Ledger Entry `Customer No.`). The result has a coloured *Match* column in column A and is re-run on **Refresh**.

**Reconcile (Phase 3, recon.js engine + main.js `runRecon`).**
- **Templates** pre-fill everything from the column headers: GST vs GSTR-2B, Bank vs G/L, Party ledger, or Custom. Each choice stays editable.
- **Imported sheets are read smartly.** The header row is detected (title rows above it are skipped), and two-row grouped headers such as GSTR-2B's "Invoice details" over "Invoice number / Invoice Date" are merged so the lower label wins.
- **Values are normalised.** Amounts in text form (`1,23,456.50`, `(500.00)`, `₹ 1,200`) and Indian-order text dates (`05-10-2026`, `05/10/26`, `5-Oct-2026`) are understood. With *Normalise document numbers*, `INV/0012`, `inv-12` and `Inv 12` count as the same key. *Sum lines per key* turns line-level books into one row per invoice. An amount can be *Amount − Minus* (bank Credit − Debit), and the second sheet's sign can be flipped (party ledgers).
- **Matching runs in passes,** each row used at most once: (1) same key and amount within the tolerance → *Matched*; (2) same key, different amount → *Amount difference*; (3) optionally, same amount within the tolerance and dates within N days, closest date first → *Matched*; leftovers → *Only in <first>* / *Only in <second>*.
- **Output.** A result sheet sorted exceptions first (differences, then one-sided rows, then matches), with Status and Match rule in columns A–B, both sides' keys, dates and amounts, the *Difference* and *Days apart*, coloured by status. A **summary sheet** uses live `COUNTIF`/`SUMIF` formulas over the result sheet, so when a reviewer changes a status by hand the totals update instantly. It also shows *Matched %* and *Rows to review*. The definition is saved and re-run on **Refresh** after the BC sheets it depends on.

**Regions (Stage 1b, main.js "Regions" section).**
- **Inserting.** *Insert region* (toolbar) or *Insert Business Central region here* (right-click) opens the side panel for the selected cell. While the panel is open it follows the cell you click; its subtitle always says where the region will go (*New region at Sheet1!A4*) or which region you are changing. Pick the table, fields and filters; *At A5* puts the region there, *New sheet* puts it at A1 of a new sheet. A region cannot start inside another region.
- **What a region is.** A source with an id and an anchor cell (`anchor {row, col}`); several can share a sheet. Loads write only the region's own block, so titles, notes and formulas around it stay. A load never grows into the next region below (it stops one row above it and says so).
- **Region bar and properties.** Selecting a cell inside a region shows a bar with its table, place, rows and filters, the *Editing* switch, *Properties*, *Refresh* and *Unlink*. *Properties* (also on the right-click menu) opens the same panel filled with the region's fields, filters, max rows and editing choice; *Update region* reloads it in place.
- **Moving with the sheet.** Inserting or deleting rows or columns above or left of a region moves its anchor, like an Excel table. Deleting its header row unlinks it (the data stays as plain cells).
- **Cell-driven filters.** A filter may contain `[B1]`, `[Sheet1!B1]` or ranges such as `[B1]..[C1]`. Before each load the reference is replaced by the cell's displayed value (so dates follow the cell's format). When the cell changes, every region that uses it refreshes on its own; an editable region with unsaved BC changes is not refreshed and says why.
- **Excel look.** Region header in Excel blue with white bold text, a thin outside border, and banded rows (one conditional-format rule) for read-only regions. New sheets use Calibri 11 and light gridlines. The single auto-filter Univer allows per sheet goes to the region when it is alone on its sheet.
- **Combine and Reconcile** read the first region of a sheet (its block only, without the BC Status column).
- **Screenshots** (mock data): [panel with a cell filter](docs/images/stage1b-r1-panel.jpg) · [region at A4 with title, filter cell and totals](docs/images/stage1b-r3-two-regions.jpg) · [right-click menu](docs/images/stage1b-r4-context.jpg) · [Editing switched off](docs/images/stage1b-r5-editing-off.jpg).

**Drill-down (Stage 5, codeunit 89007 "AK Workbook Drill Down").**
- **How.** Ctrl+click (Cmd+click on Mac) a cell inside a region, or right-click > *Open in Business Central (Ctrl+click)*, or *Open* on the region bar.
- **Region rows** open the record on its own page: the add-in sends the row's key values; AL filters the table on them (Code/Text values quoted, ISO dates localized) and calls `Page Management.PageRun`, which picks the table's card or list page. The region must contain all key fields (the panel's *Key* button selects them).
- **Budget grid.** A period cell or Total opens *G/L Budget Entries* for that budget, account, period and dimension filters; a last-year cell or LY Total opens *General Ledger Entries* for the same period a year earlier. Heading and total accounts drill into the accounts in their *Totaling*.
- Pages open next to the workbook; if BC cannot find the record (deleted since the load) the add-in says so.

**Period picker (Stage 4, main.js "period picker" section).**
- **Inserting.** *Period picker* in the Quick start tiles or *Insert period picker here* on the right-click menu turns the selected cell into a yellow dropdown: Today, This Week, This/Last Month, This/Last Quarter, Year to Date, This/Last Year, Fiscal Year to Date, This/Last Fiscal Year, Last 7/30/90 Days, Next 30 Days. The cell to its right shows the dates (for example *01-Apr-26 to 31-Mar-27*).
- **Using it.** In a region, put the picker cell as the filter of a date field, for example Posting Date = `[B2]`. The add-in sends the period as `yyyy-mm-dd..yyyy-mm-dd`, and `AK Workbook Data Service.LocalizeIsoDates` rewrites each date in the user's own date format before `SetFilter`. Choosing another period refreshes every region that points to the picker. A BC date filter typed into the picker cell (`01-04-26..30-06-26`) is passed on as typed.
- **Dates.** Periods are counted from the BC work date; fiscal periods use the fiscal year start from the Accounting Periods (as the budget panel does). Pickers move with inserted or deleted rows and columns, and are saved with the workbook.
- **Screenshots** (mock data): [Last 90 Days driving a region](docs/images/stage4-p2-picker.jpg) · [the dropdown](docs/images/stage4-p3-dropdown.jpg).

**Budget grid (Stage 3, codeunits 89005 and 89006).**
- **Inserting.** *Budget grid* in the Quick start tiles or *Insert budget grid here* on the right-click menu opens the budget panel: budget name, starting date, number of periods (months or quarters), G/L account filter, Global Dimension 1 and 2 filters (each can be a cell like `[B1]`) and *Show last year's actuals*.
- **Loading.** AL reads every G/L account in the filter (at most 2,000) and, for each period, `Budgeted Amount` with the budget, date and dimension filters, plus last year's `Net Change` for the same period when asked. The region shows No., Name (indented), one column per period, **Total** (a `SUM` formula), last year's block with its total, and BC Status. Heading and total accounts are bold and grey.
- **Editing.** Only period cells of posting accounts can change; anything else is put back. A number typed into **Total** is spread over the periods by last year's actuals (evenly when there are none) and Total becomes a formula again. Changed cells turn yellow and count on **Save to BC**.
- **Saving.** For each changed cell, `AK Workbook Budget Write` works out the current budget for that account, period, budget and dimension filter and inserts one **G/L Budget Entry for the difference**, dated on the first day of the period, with the dimension values from the filters (they must be empty or one value). Each cell is its own transaction; results show in BC Status. When all cells saved, the grid reloads so totaling accounts are current. Telemetry `AK0003`.
- **Setup.** Add *G/L Budget Entry* (table 96) with *Allow Insert* on the Editable Tables page; otherwise the grid is read-only and says why. Blocked budgets are read-only.
- **Screenshots** (mock data): [budget panel with a [B2] department filter](docs/images/stage3-b1-panel.jpg) · [yearly total spread by last year's actuals](docs/images/stage3-b2-edited.jpg) · [after Save to BC](docs/images/stage3-b4-saved.jpg).

**Journal and document lines (Stage 2, codeunit 89004 "AK Workbook Lines").**
- **Region definitions.** General Journal, Item Journal, Sales Line and Purchase Line regions list their fields in page order (it is also the `Validate` order) using `FieldNo(...)`, filtered on journal template and batch or on document type and no.
- **Open from the page.** *Edit in Excel Workbook* on the General Journal, Item Journal, Sales Order and Purchase Order pages calls `AK Workbook Mgt.OpenWithRegion`: it creates a workbook for that batch or document if needed (code such as `GJ GENERAL CASH` or `S SO-1042`) and passes the region as `autoRegion` in the add-in configuration. The add-in shows that region (refreshing it) or creates it on an empty sheet. The journal page's template and batch come from its filter group 2.
- **Quick starts.** The same four regions are sent as `quickStarts` in the configuration and shown as tiles in the region panel (with the first general/item journal batch and the last open sales/purchase order as starting filters).
- **New lines.** For a new row, a key cell left empty takes the region's filter value when that filter is one plain value (template, batch, document type and no.). In AL (`Write Row.InsertRow`) an empty last Integer key (Line No.) becomes the last line no. of the same batch or document + 10000. Journal lines then run the page's own `SetUpNewLine` (posting date, document no. from the batch's number series, balancing account, source code), before `Insert(true)` and the `Validate` of each typed field. Sales and purchase lines use the order's `OnInsert` checks (status open, header exists).
- **Setup.** Add *Gen. Journal Line*, *Item Journal Line*, *Sales Line* and *Purchase Line* on the Editable Tables page with *Allow Modify* and *Allow Insert*; otherwise the region opens read-only and says why.
- **Screenshots** (mock data): [new journal lines typed below, totals row ignored](docs/images/stage2-j1-new-lines.jpg) · [saved with Line No. and Document No. from BC](docs/images/stage2-j2-saved.jpg) · [quick start tiles](docs/images/stage2-j3-quick.jpg).

**Editable regions and write-back (Stage 1).**
- **Turning it on.** In the region panel, the *Allow editing* box shows for every table. It can be ticked only when the table is on the *Editable Tables* list with *Allow Modify* and the user has write permission (`AK Workbook Write Service.CanModify`); otherwise it is greyed out with the reason (`GetEditBlockReason`). With it ticked, all primary-key fields are loaded (they identify the record). The load adds a *BC Status* column and, per row, the record's `SystemModifiedAt` as a version.
- **Locked and editable columns.** Normal fields are editable; primary-key fields, FlowFields, system fields (`No. >= 2000000000`) and BLOB/Media fields are locked and shown grey. If a user types or pastes into a locked cell of an existing row, the add-in puts the old value back and says why.
- **Finding changes.** The add-in keeps a copy of every loaded row by primary key. After each edit it compares the region with that copy (by key, so sorting and filtering do not matter): an existing key with a different value in an editable column is a *modify*; a row typed directly below the data whose key is not in the copy (key empty when a number series assigns it) is an *insert*. Below the loaded rows the region ends at the first empty row, the first row with a formula (a totals row is never sent to BC) or the next region. A formula inside an editable cell sends its calculated value.
- **Editing switch.** Each editable region has an *Editing* switch on the region bar. Off: nothing is tracked, locked or marked and *Save to BC* ignores the region, so the user can view and use formulas freely. On again: the region is compared with the loaded copy and any differences are marked before anything is saved. Changed cells turn yellow, new rows green, the BC Status says *Changed* or *New*, and the green **Save to BC** button shows the count. Deleting rows in the sheet never deletes in BC.
- **Saving (codeunits 89002 and 89003).** *Save to BC* asks for confirmation and sends the changes (chunked, purpose `write`). AL first checks the request in a `TryFunction` (allow-list, permission, at most 1,000 rows); if it is refused the add-in shows the reason. Then every row runs in its own `WriteRow.Run()`:
  - *Modify*: find the record by its key, stop with "changed in Business Central after you loaded the sheet" if `SystemModifiedAt` differs, `Validate` each changed field in sheet order (text is converted to the field type: Code upper-cased with length check, decimals and dates in XML format, Options/Enums by caption, name or number), then `Modify(true)`.
  - *Insert*: `Init`, `Validate` the key fields that were typed, `Insert(true)` so the number series and `OnInsert` run, then `Validate` the other fields and `Modify(true)`.
  - A failing row rolls back only itself; each successful row is committed on its own (per-row commit on purpose: a row is what the user sees and fixes).
- **Result.** AL sends back each row as BC now has it (Validate often changes other fields, such as *Last Date Modified*), the new version, or BC's own error message. The add-in writes the saved values into the sheet, sets BC Status to *Saved hh:mm* (green) or the error (red), and keeps the failed rows marked so the user can fix and save again. Telemetry `AK0002` logs saved and failed counts per table.
- **Screenshots** (mock data): [Allow editing](docs/images/stage1-e1-get-data.jpg) · [changes marked](docs/images/stage1-e2-edited.jpg) · [confirm](docs/images/stage1-e3-confirm.jpg) · [per-row results](docs/images/stage1-e4-result.jpg).
- **Refresh and save.** Refreshing an editable sheet with unsaved BC changes asks first. The copy of the loaded rows is saved with the workbook, so pending changes survive closing and reopening.

**Safety nets.** The JS error boundary catches script errors and unhandled promises, shows a toast, and logs them to telemetry (`AK0001`) through `OnError`. Resizing is handled with a ResizeObserver. On page close the add-in disposes the spreadsheet and its listeners. Confirmations use an in-add-in dialog, because browser dialogs can be blocked inside the BC iframe.

## 6. Known limitations

- **Write-back is row by row.** It covers tables an admin allows: masters (Item, Customer, Vendor), simple setup tables, and journal, sales and purchase lines (Stage 2). Budgets use the budget grid (Stage 3). Rows are never deleted from BC, and posting still happens on the BC page.
- **Field order matters.** Changed fields are validated left to right, like typing on a page. If one field's validation resets another (for example changing *Unit of Measure* recalculates a price), put the dependent field to the right.
- **Large saves.** At most 1,000 rows per save; click *Save to BC* again for the rest. Each row is its own transaction, so a long save cannot be undone in one step.
- **Reconciliation is one-to-one.** One books row (or one grouped key) matches one statement row. Many-to-one cases, such as one bank deposit covering three receipts, show as exceptions; group them with *Sum lines per key* if they share a reference.
- **Regions.** One auto-filter per sheet (Univer limit), so only a region that is alone on its sheet gets filter buttons. Combine and Reconcile use the first region of a sheet. Cell filters use the cell's displayed text, so format date cells the way BC filters expect (for example `dd-mm-yy`).
- **Combine and Reconcile run in the browser** on the rows already in the sheets (comfortably up to about 100,000 rows per side), and they find columns by header name. If a header is renamed, Refresh reports the missing column; run the template again.
- **No close prompt.** BC cannot ask "close without saving?" on a `UserControlHost` page. Keep AutoSave on, or save before leaving.
- **Excel features not covered:** VBA/macros, pivot tables, charts, and Excel data connections. Conditional formatting and data validation work inside BC but are not written to the exported .xlsx.
- **Excel import** brings values, formulas, styles, widths, merges and freeze panes; it ignores charts, images, comments and theme colours.
- **Size.** The spreadsheet engine is about 12.5 MB (about 3 MB compressed). The first open per browser session takes a moment, then it is cached. Loads above roughly 50,000 rows × 15 columns get slow in the browser; use filters.
- **Not compiled in this hand-off.** The AL targets runtime 16 but was not compiled here, because no AL compiler was available. The JavaScript was tested end to end in Chromium against a mocked BC bridge. Run *AL: Package* once and fix anything the compiler flags.

## 7. Change log

| Version | Date | Change |
|---|---|---|
| 1.8.0.0 | 2026-11-26 | Stage 5 **drill-down**: Ctrl+click, right-click *Open in Business Central* or region bar *Open* opens the record of a region row (Page Management) or, in a budget grid, the G/L Budget Entries / G/L Entries behind an amount (totaling accounts included). New codeunit 89007. |
| 1.7.2.0 | 2026-11-26 | *Update region* always reloads the region where it is (never on a new sheet). After changing a region, clicking an empty cell switches the panel back to *At <cell>* (not *New sheet*). |
| 1.7.1.0 | 2026-11-26 | Fix: with the region or budget panel open, clicking another cell now moves the target to that cell (*New region at A4*); before, the panel kept the first cell or updated the last inserted region. A region or grid can no longer be inserted on a period picker cell. The region bar now follows mouse clicks reliably (selection is read from the `set-selections` operation). |
| 1.7.0.0 | 2026-11-26 | Stage 4 **period picker**: dropdown cell with 16 periods (calendar and fiscal), dates shown beside it; date filters that point to it get the period and refresh on change; AL localizes ISO dates in date filters (`LocalizeIsoDates`); work date in the add-in configuration. |
| 1.6.0.0 | 2026-11-25 | Stage 3 **budget grid**: budget panel (name, start, periods, month/quarter, account and dimension filters with cell references, last year's actuals), grid region with Total formulas, spreading a typed yearly total by last year's actuals, Save to BC adds one G/L Budget Entry per changed cell for the difference, per-cell results and reload after a full save. Formula recalculation is no longer mistaken for user edits. New codeunits 89005 and 89006. |
| 1.5.0.0 | 2026-11-25 | Stage 2 **journal and document lines**: *Edit in Excel Workbook* on General Journal, Item Journal, Sales Order and Purchase Order opens the lines as a region; *Quick start* tiles in the region panel; new rows get template/batch or document keys from the filters, the next Line No. and the journal page's `SetUpNewLine` defaults. Region properties keep the region's column order. New codeunit 89004, page extensions 89000–89003. |
| 1.4.1.0 | 2026-11-25 | Menus and dropdowns of the spreadsheet use the add-in font instead of a serif fallback. Paste menu uses Excel wording (Paste Values, Paste Formatting, Paste Column Widths, Paste All Except Borders); `build.mjs` re-applies these labels after a rebuild. |
| 1.4.0.0 | 2026-11-25 | Stage 1b **regions**: Insert region at any cell (toolbar or right-click), many regions per sheet, region bar with Properties / Refresh / Unlink, regions move with inserted or deleted rows and columns, loads never overwrite the next region. **Cell-driven filters** (`[B1]`, `[Sheet!B1]`) refresh the region when the cell changes. **Editing switch** per region; totals and formula rows below a region are never sent as new records. **Excel look**: Excel green accent, Calibri 11, Excel-table style regions with banding. Saved workbooks from earlier versions open with each linked sheet as a region at A1. |
| 1.3.1.0 | 2026-11-25 | *Allow editing* is always shown after picking a table; when the table cannot be edited it is greyed out with the reason (not on the Editable Tables list, Allow Modify off, or no modify permission). |
| 1.3.0.0 | 2026-11-25 | Stage 1 **editable sheets with write-back**: *Allow editing* option in Get data; locked key/calculated/system columns; change tracking by primary key with yellow/green marks and a *BC Status* column; **Save to BC** with per-row `Validate` + `Modify(true)` / `Insert(true)`, own transaction per row, `SystemModifiedAt` conflict check and per-row results. New table 89001 and page 89002 (allow-list of editable tables), codeunits 89002 Write Service and 89003 Write Row, permission set 89001 AK Workbook Admin. |
| 1.2.1.0 | 2026-11-25 | Prefix renamed from AML to AK (objects, publisher, add-in, CSS classes, bundle global `AKUniver`, telemetry tag). No functional change. |
| 1.2.0.0 | 2026-10-23 | Runtime 16 / BC v27. Editor is a trigger-free `UserControlHost` page (whole page is the spreadsheet). Phase 3 **Reconcile**: GST vs GSTR-2B, Bank vs G/L, Party ledger and Custom templates; multi-pass matching (key + amount, key, amount + date window), tolerance, sign flip, Credit − Debit, document-number normalisation, sum per key; exceptions-first result sheet and live summary sheet. Import can **add sheets** to the current workbook; title rows and two-row headers in imported files are detected. **AutoSave** switch. Visual refresh (Segoe UI Variable, softer surfaces, template tiles). |
| 1.1.0.0 | 2026-10-23 | Editor shows only the spreadsheet; commands moved into the add-in toolbar; inline rename; in-browser import. Phase 2 **Combine** sheets. Page 89001 renamed "AK Workbook Editor". |
| 1.0.0.0 | 2026-10-23 | Phase 1: spreadsheet add-in, data explorer for any table with per-field filters, multi-sheet loads, refresh with formula preservation, save and reopen, .xlsx export and import, permission set. |
