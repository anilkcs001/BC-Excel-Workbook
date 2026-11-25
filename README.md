# AK Excel Workbook

## 1. Extension name and purpose

**AK Excel Workbook** (Per-Tenant Extension, runtime 16 / Business Central v27, IN localization compatible)

It puts a full spreadsheet inside Business Central, on a page that is **only the spreadsheet**. Users can:

- pull data from **any table they are allowed to read**, picking fields and setting BC-style filters per field;
- work with it like Excel: formulas (400+ functions), formatting, filters, sort, conditional formatting, find and replace, data validation, freeze panes, copy/paste with desktop Excel;
- **refresh** every linked sheet with one click (formulas next to the data survive);
- **save** (or let **AutoSave** do it), **export** to .xlsx, and **import** .xlsx either as new sheets or as a replacement workbook;
- **combine** two sheets on matching keys (Phase 2);
- **reconcile** books against a bank statement, GSTR-2B or a party ledger (Phase 3), with an exceptions-first result sheet and a live summary sheet.

Every command (AutoSave, Save, Refresh, Import, Export, Combine, Reconcile, Get data) sits in the add-in's own toolbar. The page has no BC fields or action bar.

Delivered: Phase 1 (sheet, data explorer, save, refresh, Excel I/O), Phase 2 (Combine), Phase 3 (Reconcile). Next: Phase 4 BC formulas and write-back, Phase 5 Copilot.

## 2. Object list

| ID | Type | Name |
|---|---|---|
| 89000 | Table | AK Workbook |
| 89000 | Page | AK Workbook List |
| 89001 | Page | AK Workbook Editor (`UserControlHost`: the add-in fills the whole page) |
| 89000 | Codeunit | AK Workbook Data Service |
| 89001 | Codeunit | AK Workbook Mgt |
| 89000 | PermissionSet | AK Workbook |
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
5. **Assign permissions:** give users the permission set **AK Workbook** (*Users → User Permission Sets* or a security group). Users still need their normal permissions for each source table (for example *D365 BASIC*, *D365 FIN. ACC*, GST permission sets). The workbook never bypasses them.
6. **Use it:** search *Excel Workbooks* → add a line with a code and name → **Open**. The spreadsheet fills the page. Click the workbook name at the top left to rename it.

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

**Safety nets.** The JS error boundary catches script errors and unhandled promises, shows a toast, and logs them to telemetry (`AK0001`) through `OnError`. Resizing is handled with a ResizeObserver. On page close the add-in disposes the spreadsheet and its listeners. Confirmations use an in-add-in dialog, because browser dialogs can be blocked inside the BC iframe.

## 6. Known limitations

- **No write-back yet.** Edits stay in the workbook; they do not change BC records (Phase 4).
- **Reconciliation is one-to-one.** One books row (or one grouped key) matches one statement row. Many-to-one cases, such as one bank deposit covering three receipts, show as exceptions; group them with *Sum lines per key* if they share a reference.
- **Combine and Reconcile run in the browser** on the rows already in the sheets (comfortably up to about 100,000 rows per side), and they find columns by header name. If a header is renamed, Refresh reports the missing column; run the template again.
- **No close prompt.** BC cannot ask "close without saving?" on a `UserControlHost` page. Keep AutoSave on, or save before leaving.
- **Excel features not covered:** VBA/macros, pivot tables, charts, and Excel data connections. Conditional formatting and data validation work inside BC but are not written to the exported .xlsx.
- **Excel import** brings values, formulas, styles, widths, merges and freeze panes; it ignores charts, images, comments and theme colours.
- **Size.** The spreadsheet engine is about 12.5 MB (about 3 MB compressed). The first open per browser session takes a moment, then it is cached. Loads above roughly 50,000 rows × 15 columns get slow in the browser; use filters.
- **Not compiled in this hand-off.** The AL targets runtime 16 but was not compiled here, because no AL compiler was available. The JavaScript was tested end to end in Chromium against a mocked BC bridge. Run *AL: Package* once and fix anything the compiler flags.

## 7. Change log

| Version | Date | Change |
|---|---|---|
| 1.2.1.0 | 2026-11-25 | Prefix renamed from AML to AK (objects, publisher, add-in, CSS classes, bundle global `AKUniver`, telemetry tag). No functional change. |
| 1.2.0.0 | 2026-10-23 | Runtime 16 / BC v27. Editor is a trigger-free `UserControlHost` page (whole page is the spreadsheet). Phase 3 **Reconcile**: GST vs GSTR-2B, Bank vs G/L, Party ledger and Custom templates; multi-pass matching (key + amount, key, amount + date window), tolerance, sign flip, Credit − Debit, document-number normalisation, sum per key; exceptions-first result sheet and live summary sheet. Import can **add sheets** to the current workbook; title rows and two-row headers in imported files are detected. **AutoSave** switch. Visual refresh (Segoe UI Variable, softer surfaces, template tiles). |
| 1.1.0.0 | 2026-10-23 | Editor shows only the spreadsheet; commands moved into the add-in toolbar; inline rename; in-browser import. Phase 2 **Combine** sheets. Page 89001 renamed "AK Workbook Editor". |
| 1.0.0.0 | 2026-10-23 | Phase 1: spreadsheet add-in, data explorer for any table with per-field filters, multi-sheet loads, refresh with formula preservation, save and reopen, .xlsx export and import, permission set. |
