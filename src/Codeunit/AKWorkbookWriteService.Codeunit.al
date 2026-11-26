codeunit 89002 "AK Workbook Write Service"
{
    // Stage 1 write-back: applies the changes a user made in an editable workbook sheet.
    //
    // Request (JSON, sent by the add-in in chunks with purpose 'write'):
    //   { requestId, sheetId, tableNo, fields: [fieldNo...], keyFields: [fieldNo...],
    //     changes: [ { row, op: 'modify'|'insert', key: [text...], version, values: [ { no, value } ] } ] }
    // Response (sent back with purpose 'writeresult'):
    //   { requestId, sheetId, results: [ { row, ok, error, key: [...], version, values: [...] } ], saved, failed }
    //
    // Every row runs in its own Codeunit.Run, so one failing row (for example a blocked vendor) does not
    // stop the others, and each successful row is committed on its own. That is why this codeunit commits
    // per row instead of every 100 records: a row is the unit the user sees and fixes.

    Access = Public;

    var
        NotEditableErr: Label '%1 (table %2) is not set up for editing. Ask your administrator to add it on the Excel Workbook Editable Tables page.', Comment = '%1 = table caption, %2 = table number';
        NoModifyPermissionErr: Label 'You do not have permission to change records in %1.', Comment = '%1 = table caption';
        NoInsertPermissionErr: Label 'You cannot create new records in %1 from a workbook.', Comment = '%1 = table caption';
        TooManyChangesErr: Label 'You can save at most %1 rows at a time. Save in smaller parts.', Comment = '%1 = maximum number of rows';
        InvalidRequestErr: Label 'The write request is missing the property "%1".', Comment = '%1 = JSON property name';
        UnknownErrorTxt: Label 'The change could not be saved.';
        NotListedTxt: Label '%1 is not on the Excel Workbook Editable Tables list. Ask your administrator to add it.', Comment = '%1 = table caption';
        ModifyOffTxt: Label 'Allow Modify is turned off for %1 on the Excel Workbook Editable Tables page.', Comment = '%1 = table caption';
        SystemTableTxt: Label 'System tables cannot be edited from a workbook.';

    /// <summary>True when the table is in the allow-list with Allow Modify and the user may modify it.</summary>
    procedure CanModify(TableNo: Integer): Boolean
    var
        EditTable: Record "AK Workbook Edit Table";
        RecRef: RecordRef;
        Allowed: Boolean;
    begin
        if not EditTable.Get(TableNo) then
            exit(false);
        if not EditTable."Allow Modify" then
            exit(false);
        RecRef.Open(TableNo);
        Allowed := RecRef.ReadPermission() and RecRef.WritePermission();
        RecRef.Close();
        exit(Allowed);
    end;

    /// <summary>True when the table is in the allow-list with Allow Insert and the user may insert into it.</summary>
    procedure CanInsert(TableNo: Integer): Boolean
    var
        EditTable: Record "AK Workbook Edit Table";
        RecRef: RecordRef;
        Allowed: Boolean;
    begin
        if not EditTable.Get(TableNo) then
            exit(false);
        if not EditTable."Allow Insert" then
            exit(false);
        RecRef.Open(TableNo);
        Allowed := RecRef.ReadPermission() and RecRef.WritePermission();
        RecRef.Close();
        exit(Allowed);
    end;

    /// <summary>Why a table cannot be edited from a workbook, in words for the user. Empty when editing is allowed.</summary>
    procedure GetEditBlockReason(TableNo: Integer): Text
    var
        EditTable: Record "AK Workbook Edit Table";
        RecRef: RecordRef;
        Reason: Text;
    begin
        if TableNo >= 2000000000 then
            exit(SystemTableTxt);
        RecRef.Open(TableNo);
        if not EditTable.Get(TableNo) then
            Reason := StrSubstNo(NotListedTxt, RecRef.Caption())
        else
            if not EditTable."Allow Modify" then
                Reason := StrSubstNo(ModifyOffTxt, RecRef.Caption())
            else
                if not (RecRef.ReadPermission() and RecRef.WritePermission()) then
                    Reason := StrSubstNo(NoModifyPermissionErr, RecRef.Caption());
        RecRef.Close();
        exit(Reason);
    end;

    procedure CheckCanModify(TableNo: Integer)
    var
        EditTable: Record "AK Workbook Edit Table";
        RecRef: RecordRef;
    begin
        RecRef.Open(TableNo);
        if not EditTable.Get(TableNo) then
            Error(NotEditableErr, RecRef.Caption(), TableNo);
        if not EditTable."Allow Modify" then
            Error(NotEditableErr, RecRef.Caption(), TableNo);
        if not (RecRef.ReadPermission() and RecRef.WritePermission()) then
            Error(NoModifyPermissionErr, RecRef.Caption());
    end;

    procedure CheckCanInsert(TableNo: Integer)
    var
        RecRef: RecordRef;
    begin
        RecRef.Open(TableNo);
        if not CanInsert(TableNo) then
            Error(NoInsertPermissionErr, RecRef.Caption());
    end;

    /// <summary>Largest number of rows one save may contain.</summary>
    procedure MaxRowsPerSave(): Integer
    begin
        exit(1000);
    end;

    /// <summary>Checks a write request before anything is written (no database writes, safe inside a TryFunction).</summary>
    procedure CheckRequest(RequestJson: Text)
    var
        Request: JsonObject;
        ChangesToken: JsonToken;
        FieldNos: List of [Integer];
        KeyFieldNos: List of [Integer];
    begin
        Request.ReadFrom(RequestJson);
        ReadIntegerList(Request, 'fields', FieldNos);
        ReadIntegerList(Request, 'keyFields', KeyFieldNos);
        if not Request.Get('changes', ChangesToken) then
            Error(InvalidRequestErr, 'changes');
        if ChangesToken.AsArray().Count() > MaxRowsPerSave() then
            Error(TooManyChangesErr, MaxRowsPerSave());
        CheckCanModify(GetInteger(Request, 'tableNo'));
    end;

    /// <summary>Response used when the whole request was refused, so the add-in can show the reason and stop waiting.</summary>
    procedure BuildRefusedResponse(RequestJson: Text; Message: Text): Text
    var
        Request: JsonObject;
        Response: JsonObject;
        Result: Text;
    begin
        if Request.ReadFrom(RequestJson) then begin
            Response.Add('requestId', GetText(Request, 'requestId'));
            Response.Add('sheetId', GetText(Request, 'sheetId'));
        end;
        Response.Add('refused', Message);
        Response.WriteTo(Result);
        exit(Result);
    end;

    /// <summary>Applies all changes of one sheet and returns the per-row results as JSON.</summary>
    procedure ApplyChanges(RequestJson: Text): Text
    var
        WriteRow: Codeunit "AK Workbook Write Row";
        WriteBatch: Codeunit "AK Workbook Write Batch";
        ResultToken: JsonToken;
        BatchResults: JsonArray;
        Total: Integer;
        BatchStart: Integer;
        BatchEnd: Integer;
        Index: Integer;
        BatchDone: Boolean;
        AccessChecked: Boolean;
        Request: JsonObject;
        Response: JsonObject;
        Results: JsonArray;
        ChangesToken: JsonToken;
        ChangeToken: JsonToken;
        FieldNos: List of [Integer];
        KeyFieldNos: List of [Integer];
        TableNo: Integer;
        Saved: Integer;
        Failed: Integer;
        Result: Text;
        StartedAt: DateTime;
        ElapsedMs: BigInteger;
    begin
        StartedAt := CurrentDateTime();
        Request.ReadFrom(RequestJson);
        TableNo := GetInteger(Request, 'tableNo');
        ReadIntegerList(Request, 'fields', FieldNos);
        ReadIntegerList(Request, 'keyFields', KeyFieldNos);
        if not Request.Get('changes', ChangesToken) then
            Error(InvalidRequestErr, 'changes');
        if ChangesToken.AsArray().Count() > MaxRowsPerSave() then
            Error(TooManyChangesErr, MaxRowsPerSave());

        CheckCanModify(TableNo);
        // Checked once here: rows in a batch skip the per-row check. Without insert rights, inserts are
        // written row by row so each new row reports the reason.
        AccessChecked := (not HasInserts(ChangesToken.AsArray())) or CanInsert(TableNo);

        // Batches of BatchSize() rows: one transaction and one commit per batch. A batch with any error is
        // rolled back and written again row by row, so good rows are saved and each bad row shows its error.
        Total := ChangesToken.AsArray().Count();
        BatchStart := 0;
        while BatchStart < Total do begin
            BatchEnd := BatchStart + BatchSize() - 1;
            if BatchEnd > Total - 1 then
                BatchEnd := Total - 1;
            BatchDone := false;
            if AccessChecked and (BatchEnd > BatchStart) then begin
                Clear(WriteBatch);
                WriteBatch.SetBatch(TableNo, FieldNos, KeyFieldNos, ChangesToken.AsArray(), BatchStart, BatchEnd, true);
                ClearLastError();
                if WriteBatch.Run() then begin
                    Commit();
                    BatchResults := WriteBatch.GetResults();
                    foreach ResultToken in BatchResults do
                        Results.Add(ResultToken);
                    Saved += BatchEnd - BatchStart + 1;
                    BatchDone := true;
                end;
            end;
            if not BatchDone then
                for Index := BatchStart to BatchEnd do begin
                    ChangesToken.AsArray().Get(Index, ChangeToken);
                    Clear(WriteRow);
                    WriteRow.SetChange(TableNo, FieldNos, KeyFieldNos, ChangeToken.AsObject());
                    ClearLastError();
                    if WriteRow.Run() then begin
                        Commit();
                        Results.Add(WriteRow.GetResult());
                        Saved += 1;
                    end else begin
                        Results.Add(BuildErrorResult(ChangeToken.AsObject()));
                        Failed += 1;
                    end;
                end;
            BatchStart := BatchEnd + 1;
        end;

        Response.Add('requestId', GetText(Request, 'requestId'));
        Response.Add('sheetId', GetText(Request, 'sheetId'));
        Response.Add('results', Results);
        Response.Add('saved', Saved);
        Response.Add('failed', Failed);
        // Time spent in BC (validation, triggers, commits), shown to the user next to the total time.
        ElapsedMs := CurrentDateTime() - StartedAt;
        Response.Add('ms', ElapsedMs);
        Response.WriteTo(Result);

        Session.LogMessage('AK0002', StrSubstNo('Workbook write-back: table %1, %2 saved, %3 failed.', TableNo, Saved, Failed),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', 'AK Excel Workbook');
        exit(Result);
    end;

    /// <summary>Rows written in one transaction (and one commit), as in the "commit every 100 records" pattern.</summary>
    procedure BatchSize(): Integer
    begin
        exit(100);
    end;

    local procedure HasInserts(Changes: JsonArray): Boolean
    var
        ChangeToken: JsonToken;
        Token: JsonToken;
    begin
        foreach ChangeToken in Changes do
            if ChangeToken.AsObject().Get('op', Token) then
                if Token.AsValue().AsText() = 'insert' then
                    exit(true);
        exit(false);
    end;

    local procedure BuildErrorResult(Change: JsonObject): JsonObject
    var
        Item: JsonObject;
        Token: JsonToken;
        Message: Text;
    begin
        Message := GetLastErrorText();
        if Message = '' then
            Message := UnknownErrorTxt;
        if Change.Get('row', Token) then
            Item.Add('row', Token.AsValue().AsInteger());
        Item.Add('ok', false);
        Item.Add('error', Message);
        exit(Item);
    end;

    local procedure ReadIntegerList(Request: JsonObject; PropertyName: Text; var Numbers: List of [Integer])
    var
        Token: JsonToken;
        Item: JsonToken;
    begin
        if not Request.Get(PropertyName, Token) then
            Error(InvalidRequestErr, PropertyName);
        foreach Item in Token.AsArray() do
            Numbers.Add(Item.AsValue().AsInteger());
    end;

    local procedure GetInteger(JObject: JsonObject; PropertyName: Text): Integer
    var
        Token: JsonToken;
    begin
        if not JObject.Get(PropertyName, Token) then
            Error(InvalidRequestErr, PropertyName);
        exit(Token.AsValue().AsInteger());
    end;

    local procedure GetText(JObject: JsonObject; PropertyName: Text): Text
    var
        Token: JsonToken;
    begin
        if not JObject.Get(PropertyName, Token) then
            exit('');
        if Token.AsValue().IsNull() then
            exit('');
        exit(Token.AsValue().AsText());
    end;
}
