controladdin "AK Excel Workbook"
{
    // Spreadsheet engine: Univer (Apache-2.0), bundled locally - no CDN calls at runtime.
    // Excel file I/O: ExcelJS (MIT). Rebuild vendor files with /build (npm run build).
    Scripts =
        'ControlAddin/scripts/vendor/univer.bundle.js',
        'ControlAddin/scripts/vendor/exceljs.min.js',
        'ControlAddin/scripts/helper.js',
        'ControlAddin/scripts/recon.js';
    StartupScript = 'ControlAddin/scripts/main.js';
    StyleSheets =
        'ControlAddin/stylesheets/vendor/univer.bundle.css',
        'ControlAddin/stylesheets/style.css';

    RequestedHeight = 760;
    MinimumHeight = 460;
    RequestedWidth = 1200;
    MinimumWidth = 640;
    VerticalStretch = true;
    VerticalShrink = true;
    HorizontalStretch = true;
    HorizontalShrink = true;

    // ---------- JS -> AL ----------

    /// <summary>Scripts are loaded and the add-in can receive Init().</summary>
    event ControlAddInReady();

    /// <summary>The spreadsheet is created and sized; safe to load a saved workbook.</summary>
    event AddInFrameReady();

    /// <summary>User typed in the table search box of the data panel.</summary>
    event OnTablesRequested(SearchText: Text);

    /// <summary>User picked a table; AL should answer with SetFields().</summary>
    event OnFieldsRequested(TableNo: Integer);

    /// <summary>User clicked Load or Refresh. JSON: requestId, tableNo, fields[], filters[{fieldNo, filter}], maxRows, editable.</summary>
    event OnDataRequested(RequestJson: Text);

    /// <summary>Sent right before a 'save' payload, with a short summary of linked tables.</summary>
    event OnWorkbookSaving(LinkedTables: Text);

    /// <summary>Chunked transfer of large payloads. Purpose: 'save' or 'autosave' (workbook JSON; autosave is stored without a toast), 'export' (xlsx as Base64) or 'write' (changed sheet rows to write back to BC).</summary>
    event OnPayloadChunk(Purpose: Text; ChunkIndex: Integer; ChunkCount: Integer; Data: Text);

    /// <summary>User renamed the workbook in the add-in header. AL stores it as Description.</summary>
    event OnDescriptionChanged(NewDescription: Text);

    /// <summary>Workbook changed since the last save (true) or was just saved/loaded (false).</summary>
    event OnDirtyChanged(IsDirty: Boolean);

    /// <summary>Error caught by the JS error boundary, for telemetry.</summary>
    event OnError(Message: Text);

    // ---------- AL -> JS ----------

    /// <summary>Starts the spreadsheet. ConfigJson: workbookCode, description, defaultMaxRows, maxRowsLimit, chunkSize, userId, companyName.</summary>
    procedure Init(ConfigJson: Text);

    /// <summary>Re-runs every linked data source, then every combined and reconciled sheet, keeping user formulas outside the data range.</summary>
    procedure Refresh();

    /// <summary>Opens or closes the Get data panel.</summary>
    procedure TogglePanel();

    /// <summary>Answer to OnTablesRequested: JSON array of {id, name, caption}.</summary>
    procedure SetTables(TablesJson: Text);

    /// <summary>Answer to OnFieldsRequested: {tableNo, tableCaption, canRead, fields[{no, name, caption, type, flow, pk}]}.</summary>
    procedure SetFields(FieldsJson: Text);

    /// <summary>One page of rows for a data request. The final chunk has isLast = true.</summary>
    procedure AppendData(ChunkJson: Text);

    /// <summary>A data request failed (bad filter, no permission...).</summary>
    procedure DataError(RequestId: Text; Message: Text);

    /// <summary>Chunked transfer AL -> JS. Purpose: 'load' (saved workbook JSON), 'import' (xlsx as Base64, optional - the toolbar imports directly in the browser) or 'writeresult' (per-row results of a write-back).</summary>
    procedure ReceiveChunk(Purpose: Text; ChunkIndex: Integer; ChunkCount: Integer; Data: Text);

    /// <summary>Asks JS to serialize the workbook and send it back with OnWorkbookSaving + OnPayloadChunk('save').</summary>
    procedure RequestSave();

    /// <summary>Asks JS to build an .xlsx and send it back with OnPayloadChunk('export').</summary>
    procedure RequestExport();

    /// <summary>Shows a toast. Level: info, success, warning, error.</summary>
    procedure ShowToast(Level: Text; Message: Text);
}
