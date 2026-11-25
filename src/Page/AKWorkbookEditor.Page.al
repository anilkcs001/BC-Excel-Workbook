page 89001 "AK Workbook Editor"
{
    // UserControlHost (runtime 15+): the add-in gets the entire page - no fields, no actions.
    // Save, Refresh, Get data, Combine, Reconcile, Import, Export and Save to BC live in the add-in's own toolbar.
    // This page type allows very few triggers, so all set-up happens in ControlAddInReady.
    Caption = 'Excel Workbook';
    PageType = UserControlHost;
    ApplicationArea = All;
    UsageCategory = None;

    layout
    {
        area(Content)
        {
            usercontrol(Workbook; "AK Excel Workbook")
            {
                trigger ControlAddInReady()
                var
                    Workbook: Record "AK Workbook";
                begin
                    if (WorkbookCode = '') or not Workbook.Get(WorkbookCode) then begin
                        CurrPage.Workbook.ShowToast('error', NoWorkbookErr);
                        exit;
                    end;
                    AddInReady := true;
                    UpdateCaption();
                    CurrPage.Workbook.Init(WorkbookMgt.GetEditorConfigJson(WorkbookCode, AutoRegionJson));
                end;

                trigger AddInFrameReady()
                begin
                    SendSavedWorkbook();
                end;

                trigger OnTablesRequested(SearchText: Text)
                begin
                    CurrPage.Workbook.SetTables(DataService.GetTablesJson(SearchText));
                end;

                trigger OnFieldsRequested(TableNo: Integer)
                var
                    FieldsJson: Text;
                begin
                    ClearLastError();
                    if TryGetFields(TableNo, FieldsJson) then
                        CurrPage.Workbook.SetFields(FieldsJson)
                    else
                        CurrPage.Workbook.ShowToast('error', GetLastErrorText());
                end;

                trigger OnDataRequested(RequestJson: Text)
                var
                    Chunks: List of [Text];
                    Chunk: Text;
                begin
                    ClearLastError();
                    if not TryBuildData(RequestJson, Chunks) then begin
                        CurrPage.Workbook.DataError(DataService.GetRequestId(RequestJson), GetLastErrorText());
                        exit;
                    end;
                    foreach Chunk in Chunks do
                        CurrPage.Workbook.AppendData(Chunk);
                end;

                trigger OnWorkbookSaving(LinkedTables: Text)
                begin
                    PendingLinkedTables := LinkedTables;
                end;

                trigger OnPayloadChunk(Purpose: Text; ChunkIndex: Integer; ChunkCount: Integer; Data: Text)
                begin
                    if ChunkIndex = 0 then begin
                        Clear(PayloadBuilder);
                        PayloadPurpose := Purpose;
                    end;
                    if Purpose <> PayloadPurpose then
                        exit;
                    PayloadBuilder.Append(Data);
                    if ChunkIndex < ChunkCount - 1 then
                        exit;
                    HandleIncomingPayload(PayloadPurpose, PayloadBuilder.ToText());
                    Clear(PayloadBuilder);
                    PayloadPurpose := '';
                end;

                trigger OnDescriptionChanged(NewDescription: Text)
                begin
                    WorkbookMgt.SetDescription(WorkbookCode, NewDescription);
                    UpdateCaption();
                end;

                trigger OnDirtyChanged(IsDirty: Boolean)
                begin
                    HasUnsavedChanges := IsDirty;
                end;

                trigger OnError(Message: Text)
                begin
                    Session.LogMessage('AK0001', CopyStr(Message, 1, 2000), Verbosity::Warning, DataClassification::SystemMetadata,
                        TelemetryScope::ExtensionPublisher, 'Category', 'AK Excel Workbook');
                end;
            }
        }
    }

    var
        DataService: Codeunit "AK Workbook Data Service";
        WriteService: Codeunit "AK Workbook Write Service";
        WorkbookMgt: Codeunit "AK Workbook Mgt";
        PayloadBuilder: TextBuilder;
        WorkbookCode: Code[20];
        PayloadPurpose: Text;
        PendingLinkedTables: Text;
        AutoRegionJson: Text;
        AddInReady: Boolean;
        HasUnsavedChanges: Boolean;
        NoWorkbookErr: Label 'Open a workbook from the Excel Workbooks list.';
        SavedMsg: Label 'Workbook saved.';
        CaptionTxt: Label '%1 - %2', Comment = '%1 = workbook code, %2 = workbook description', Locked = true;

    /// <summary>Sets the workbook to edit. Call before Run().</summary>
    procedure SetWorkbook(NewWorkbookCode: Code[20])
    begin
        WorkbookCode := NewWorkbookCode;
    end;

    /// <summary>Optional: a region (JSON from "AK Workbook Lines") to show or create when the editor starts. Call before Run().</summary>
    procedure SetAutoRegion(NewAutoRegionJson: Text)
    begin
        AutoRegionJson := NewAutoRegionJson;
    end;

    local procedure UpdateCaption()
    var
        Workbook: Record "AK Workbook";
    begin
        if not Workbook.Get(WorkbookCode) then
            exit;
        if Workbook.Description = '' then
            CurrPage.Caption := Workbook."Code"
        else
            CurrPage.Caption := StrSubstNo(CaptionTxt, Workbook."Code", Workbook.Description);
    end;

    local procedure SendSavedWorkbook()
    var
        Payload: Text;
    begin
        Payload := WorkbookMgt.GetWorkbookPayload(WorkbookCode);
        if Payload <> '' then
            SendPayload('load', Payload);
    end;

    local procedure SendPayload(Purpose: Text; Payload: Text)
    var
        Chunks: List of [Text];
        Index: Integer;
    begin
        if not AddInReady then
            exit;
        WorkbookMgt.SplitIntoChunks(Payload, DataService.ChunkSize(), Chunks);
        for Index := 1 to Chunks.Count() do
            CurrPage.Workbook.ReceiveChunk(Purpose, Index - 1, Chunks.Count(), Chunks.Get(Index));
    end;

    local procedure HandleIncomingPayload(Purpose: Text; Payload: Text)
    var
        Workbook: Record "AK Workbook";
    begin
        case Purpose of
            'save', 'autosave':
                begin
                    WorkbookMgt.SaveWorkbook(WorkbookCode, Payload, PendingLinkedTables);
                    HasUnsavedChanges := false;
                    if Purpose = 'save' then
                        CurrPage.Workbook.ShowToast('success', SavedMsg);
                end;
            'write':
                WriteToBC(Payload);
            'export':
                begin
                    Workbook.Get(WorkbookCode);
                    WorkbookMgt.DownloadXlsx(Payload, WorkbookMgt.GetExportFileName(Workbook."Code", Workbook.Description));
                end;
        end;
    end;

    /// <summary>Stage 1 write-back: checks the request, applies each row (own transaction per row) and sends the results back.</summary>
    local procedure WriteToBC(Payload: Text)
    begin
        ClearLastError();
        if not TryCheckWrite(Payload) then begin
            SendPayload('writeresult', WriteService.BuildRefusedResponse(Payload, GetLastErrorText()));
            exit;
        end;
        SendPayload('writeresult', WriteService.ApplyChanges(Payload));
    end;

    [TryFunction]
    local procedure TryCheckWrite(Payload: Text)
    begin
        WriteService.CheckRequest(Payload);
    end;

    [TryFunction]
    local procedure TryGetFields(TableNo: Integer; var FieldsJson: Text)
    begin
        FieldsJson := DataService.GetFieldsJson(TableNo);
    end;

    [TryFunction]
    local procedure TryBuildData(RequestJson: Text; var Chunks: List of [Text])
    begin
        DataService.BuildDataChunks(RequestJson, Chunks);
    end;
}
