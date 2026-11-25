codeunit 89007 "AK Workbook Drill Down"
{
    // Stage 5: Ctrl+click (or right-click > Open in Business Central) on a workbook cell opens BC behind it.
    // - Table region row  -> the record's own card/list page (Page Management finds the right page).
    // - Budget grid cell   -> G/L Budget Entries behind the amount (budget, account or its totaling, period, dimensions).
    // - Last year's cell   -> General Ledger Entries behind the actual amount.
    // Request (purpose 'drill'): { kind: 'record' | 'budget' | 'actual', ... }. Pages open non-modal next to the workbook.

    Access = Public;

    var
        UnknownKindErr: Label 'Unknown drill-down "%1".', Comment = '%1 = drill-down kind';
        NotFoundErr: Label 'The record was not found in Business Central. Refresh the region.';
        NoPageErr: Label 'Business Central has no page to show %1 records.', Comment = '%1 = table caption';

    procedure Open(RequestJson: Text)
    var
        Request: JsonObject;
        Kind: Text;
    begin
        Request.ReadFrom(RequestJson);
        Kind := GetText(Request, 'kind');
        case Kind of
            'record':
                OpenRecord(Request);
            'budget':
                OpenBudgetEntries(Request);
            'actual':
                OpenGLEntries(Request);
            else
                Error(UnknownKindErr, Kind);
        end;
    end;

    local procedure OpenRecord(Request: JsonObject)
    var
        PageManagement: Codeunit "Page Management";
        DataService: Codeunit "AK Workbook Data Service";
        RecRef: RecordRef;
        FldRef: FieldRef;
        KeysToken: JsonToken;
        ValuesToken: JsonToken;
        Index: Integer;
        Value: Text;
    begin
        RecRef.Open(GetInteger(Request, 'tableNo'));
        Request.Get('keyFields', KeysToken);
        Request.Get('key', ValuesToken);
        for Index := 0 to KeysToken.AsArray().Count() - 1 do begin
            FldRef := RecRef.Field(GetArrayInteger(KeysToken.AsArray(), Index));
            Value := GetArrayText(ValuesToken.AsArray(), Index);
            case FldRef.Type() of
                FieldType::Code, FieldType::Text:
                    FldRef.SetFilter('''' + Value.Replace('''', '''''') + '''');
                FieldType::Date, FieldType::DateTime:
                    FldRef.SetFilter(DataService.LocalizeIsoDates(Value));
                else
                    FldRef.SetFilter(Value);
            end;
        end;
        if not RecRef.FindFirst() then
            Error(NotFoundErr);
        if not PageManagement.PageRun(RecRef) then
            Error(NoPageErr, RecRef.Caption());
    end;

    local procedure OpenBudgetEntries(Request: JsonObject)
    var
        GLBudgetEntry: Record "G/L Budget Entry";
    begin
        GLBudgetEntry.SetFilter("Budget Name", GetText(Request, 'budgetName'));
        GLBudgetEntry.SetFilter("G/L Account No.", AccountFilter(GetText(Request, 'account')));
        GLBudgetEntry.SetRange(Date, GetDate(Request, 'start'), GetDate(Request, 'end'));
        if GetText(Request, 'dim1Filter') <> '' then
            GLBudgetEntry.SetFilter("Global Dimension 1 Code", GetText(Request, 'dim1Filter'));
        if GetText(Request, 'dim2Filter') <> '' then
            GLBudgetEntry.SetFilter("Global Dimension 2 Code", GetText(Request, 'dim2Filter'));
        Page.Run(Page::"G/L Budget Entries", GLBudgetEntry);
    end;

    local procedure OpenGLEntries(Request: JsonObject)
    var
        GLEntry: Record "G/L Entry";
    begin
        GLEntry.SetFilter("G/L Account No.", AccountFilter(GetText(Request, 'account')));
        GLEntry.SetRange("Posting Date", GetDate(Request, 'start'), GetDate(Request, 'end'));
        if GetText(Request, 'dim1Filter') <> '' then
            GLEntry.SetFilter("Global Dimension 1 Code", GetText(Request, 'dim1Filter'));
        if GetText(Request, 'dim2Filter') <> '' then
            GLEntry.SetFilter("Global Dimension 2 Code", GetText(Request, 'dim2Filter'));
        Page.Run(Page::"General Ledger Entries", GLEntry);
    end;

    /// <summary>Totaling accounts (Total revenue...) drill into the accounts they add up.</summary>
    local procedure AccountFilter(AccountNo: Text): Text
    var
        GLAccount: Record "G/L Account";
    begin
        if GLAccount.Get(CopyStr(AccountNo, 1, MaxStrLen(GLAccount."No."))) then
            if (GLAccount."Account Type" <> GLAccount."Account Type"::Posting) and (GLAccount.Totaling <> '') then
                exit(GLAccount.Totaling);
        exit(AccountNo);
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

    local procedure GetInteger(JObject: JsonObject; PropertyName: Text): Integer
    var
        Token: JsonToken;
    begin
        JObject.Get(PropertyName, Token);
        exit(Token.AsValue().AsInteger());
    end;

    local procedure GetDate(JObject: JsonObject; PropertyName: Text): Date
    var
        Result: Date;
    begin
        Evaluate(Result, GetText(JObject, PropertyName), 9);
        exit(Result);
    end;

    local procedure GetArrayInteger(JArray: JsonArray; Index: Integer): Integer
    var
        Token: JsonToken;
    begin
        JArray.Get(Index, Token);
        exit(Token.AsValue().AsInteger());
    end;

    local procedure GetArrayText(JArray: JsonArray; Index: Integer): Text
    var
        Token: JsonToken;
    begin
        JArray.Get(Index, Token);
        if Token.AsValue().IsNull() then
            exit('');
        exit(Token.AsValue().AsText());
    end;
}
