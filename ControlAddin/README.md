# Control add-in: AK Excel Workbook

## Purpose

This add-in is the whole **AK Workbook Editor** page (`UserControlHost`): a full spreadsheet (Univer) with its own toolbar (AutoSave, Save, Refresh, Import, Export, Combine, Reconcile, Get data), an editable workbook name, and three side panels: Get data (load BC tables), Combine (join sheets) and Reconcile (books vs statement / GSTR-2B / party ledger). It moves workbooks and .xlsx files between the browser and AL.

## Files

```
ControlAddin/
├── AKExcelWorkbook.ControlAddin.al   ← AL definition (events + procedures)
├── scripts/
│   ├── main.js                        ← StartupScript: UI shell, data panel, load/refresh, save, Excel I/O
│   ├── helper.js                      ← AL bridge with chunking, date and style converters, dialogs, toasts, icons
│   ├── recon.js                       ← reconciliation engine (pure functions, no DOM): parsing, normalising, matching passes
│   └── vendor/
│       ├── univer.bundle.js           ← Univer 1.2.0 presets bundled as window.AKUniver (built from /build)
│       └── exceljs.min.js             ← ExcelJS 4.4.0
└── stylesheets/
    ├── style.css                      ← Fluent / BC look for header, panel, toasts
    └── vendor/univer.bundle.css       ← Univer styles
```

## Events (JS → AL)

| Event | When |
|---|---|
| `ControlAddInReady()` | Scripts loaded. AL must call `Init`. |
| `AddInFrameReady()` | Spreadsheet created. AL sends a saved workbook with `ReceiveChunk('load', …)`. |
| `OnTablesRequested(SearchText)` | Table search typed. AL answers with `SetTables`. |
| `OnFieldsRequested(TableNo)` | Table picked. AL answers with `SetFields`. |
| `OnDataRequested(RequestJson)` | Load or refresh. AL answers with `AppendData` (paged) or `DataError`. |
| `OnWorkbookSaving(LinkedTables)` | Sent just before a save payload. |
| `OnPayloadChunk(Purpose, Index, Count, Data)` | Chunked `save` / `autosave` (workbook JSON; autosave is stored without a toast) or `export` (xlsx Base64). |
| `OnDescriptionChanged(NewDescription)` | Workbook renamed in the header. AL stores it as Description. |
| `OnDirtyChanged(IsDirty)` | Unsaved-changes state changed. |
| `OnError(Message)` | Error boundary caught something (logged as telemetry `AK0001`). |

## Procedures (AL → JS)

`Init(ConfigJson)`, `Refresh()`, `TogglePanel()`, `SetTables(Json)`, `SetFields(Json)`, `AppendData(ChunkJson)`, `DataError(RequestId, Message)`, `ReceiveChunk(Purpose, Index, Count, Data)` for `load` (or `import`, optional: the toolbar imports .xlsx directly in the browser), `RequestSave()`, `RequestExport()`, `ShowToast(Level, Message)`.

## How to deploy

1. The vendor files are pre-built. Rebuild only to change library versions: `cd build && npm install && npm run build`.
2. Keep the folder paths as they are. The `.al` definition references `ControlAddin/scripts/...` and `ControlAddin/stylesheets/...` relative to the project root.
3. Build and publish the extension. The add-in files are packed into the .app automatically, and no web server or CDN is needed.
4. Hard-refresh the browser (Ctrl+F5) after publishing a new version so the client fetches the new scripts.
