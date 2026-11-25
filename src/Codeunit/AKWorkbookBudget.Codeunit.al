codeunit 89005 "AK Workbook Budget"
{
    // Stage 3: G/L budget grid in a workbook region.
    //
    // Load (purpose 'budget' -> 'budgetdata'): G/L accounts down, periods across. Each cell is the account's
    // "Budgeted Amount" for that period, budget name and dimension filters; optionally last year's "Net Change"
    // for the same period (actuals, read-only).
    // Save (purpose 'budgetwrite' -> 'budgetresult'): for each changed cell, "AK Workbook Budget Write" adds one
    // G/L Budget Entry for the difference between the typed amount and the current budget, like the standard
    // budget page does, so the change is easy to trace. Each cell runs in its own transaction.

    Access = Public;

    var
        InvalidRequestErr: Label 'The budget request is missing the property "%1".', Comment = '%1 = JSON property name';
        BudgetMissingErr: Label 'Budget %1 does not exist.', Comment = '%1 = budget name';
        TooManyAccountsErr: Label 'The account filter selects more than %1 accounts. Narrow it down.', Comment = '%1 = maximum number of accounts';
        TooManyPeriodsErr: Label 'Use at most %1 periods.', Comment = '%1 = maximum number of periods';
        TooManyCellsErr: Label 'You can save at most %1 budget cells at a time.', Comment = '%1 = maximum number of cells';
        NoBudgetWriteErr: Label 'Budgets cannot be changed from a workbook. Ask your administrator to add G/L Budget Entry (table 96) with Allow Insert on the Excel Workbook Editable Tables page.';
        UnknownErrorTxt: Label 'The amount could not be saved.';
        PeriodLabelTxt: Label '<Month Text,3>-<Year>', Locked = true;
        QuarterLabelTxt: Label 'Q%1-%2', Comment = '%1 = quarter number, %2 = two-digit year';

    /// <summary>Budget names, dimension captions and defaults for the budget panel (part of the add-in configuration).</summary>
    procedure GetSetupJson(): JsonObject
    var
        GLBudgetName: Record "G/L Budget Name";
        GLSetup: Record "General Ledger Setup";
        Setup: JsonObject;
        Names: JsonArray;
        NameJson: JsonObject;
    begin
        if GLBudgetName.FindSet() then
            repeat
                Clear(NameJson);
                NameJson.Add('name', GLBudgetName.Name);
                NameJson.Add('description', GLBudgetName.Description);
                NameJson.Add('blocked', GLBudgetName.Blocked);
                Names.Add(NameJson);
            until GLBudgetName.Next() = 0;
        Setup.Add('names', Names);
        if GLSetup.Get() then begin
            Setup.Add('dim1Caption', DimensionCaption(GLSetup."Global Dimension 1 Code", 1));
            Setup.Add('dim2Caption', DimensionCaption(GLSetup."Global Dimension 2 Code", 2));
        end;
        Setup.Add('startDate', Format(FiscalYearStart(WorkDate()), 0, 9));
        Setup.Add('canWrite', CanWrite());
        if not CanWrite() then
            Setup.Add('writeReason', NoBudgetWriteErr);
        exit(Setup);
    end;

    procedure CanWrite(): Boolean
    var
        WriteService: Codeunit "AK Workbook Write Service";
    begin
        exit(WriteService.CanInsert(Database::"G/L Budget Entry"));
    end;

    procedure MaxAccounts(): Integer
    begin
        exit(2000);
    end;

    procedure MaxPeriods(): Integer
    begin
        exit(24);
    end;

    /// <summary>Builds the grid. Request: requestId, budgetName, startDate (yyyy-mm-dd), periods, periodType (month|quarter), accountFilter, dim1Filter, dim2Filter, showLY.</summary>
    procedure BuildGrid(RequestJson: Text): Text
    var
        GLAccount: Record "G/L Account";
        GLBudgetName: Record "G/L Budget Name";
        Request: JsonObject;
        Response: JsonObject;
        PeriodsJson: JsonArray;
        Rows: JsonArray;
        RowJson: JsonObject;
        BudgetValues: JsonArray;
        LYValues: JsonArray;
        Starts: List of [Date];
        Ends: List of [Date];
        BudgetName: Code[10];
        StartDate: Date;
        PeriodCount: Integer;
        Index: Integer;
        ShowLY: Boolean;
        Quarterly: Boolean;
        Result: Text;
    begin
        Request.ReadFrom(RequestJson);
        BudgetName := CopyStr(GetText(Request, 'budgetName', true), 1, MaxStrLen(BudgetName));
        if not GLBudgetName.Get(BudgetName) then
            Error(BudgetMissingErr, BudgetName);
        Evaluate(StartDate, GetText(Request, 'startDate', true), 9);
        PeriodCount := GetInteger(Request, 'periods');
        if PeriodCount <= 0 then
            PeriodCount := 12;
        if PeriodCount > MaxPeriods() then
            Error(TooManyPeriodsErr, MaxPeriods());
        Quarterly := GetText(Request, 'periodType', false) = 'quarter';
        ShowLY := GetBoolean(Request, 'showLY');

        BuildPeriods(StartDate, PeriodCount, Quarterly, Starts, Ends, PeriodsJson);

        GLAccount.SetFilter("No.", GetText(Request, 'accountFilter', false));
        if GLAccount.Count() > MaxAccounts() then
            Error(TooManyAccountsErr, MaxAccounts());
        GLAccount.SetFilter("Budget Filter", BudgetName);
        if GetText(Request, 'dim1Filter', false) <> '' then
            GLAccount.SetFilter("Global Dimension 1 Filter", GetText(Request, 'dim1Filter', false));
        if GetText(Request, 'dim2Filter', false) <> '' then
            GLAccount.SetFilter("Global Dimension 2 Filter", GetText(Request, 'dim2Filter', false));
        GLAccount.SetLoadFields("No.", Name, "Account Type", Indentation, Blocked);
        if GLAccount.FindSet() then
            repeat
                Clear(RowJson);
                Clear(BudgetValues);
                Clear(LYValues);
                RowJson.Add('no', GLAccount."No.");
                RowJson.Add('name', GLAccount.Name);
                RowJson.Add('type', Format(GLAccount."Account Type"));
                RowJson.Add('posting', GLAccount."Account Type" = GLAccount."Account Type"::Posting);
                RowJson.Add('indent', GLAccount.Indentation);
                for Index := 1 to Starts.Count() do begin
                    GLAccount.SetRange("Date Filter", Starts.Get(Index), Ends.Get(Index));
                    GLAccount.CalcFields("Budgeted Amount");
                    BudgetValues.Add(GLAccount."Budgeted Amount");
                    if ShowLY then begin
                        GLAccount.SetRange("Date Filter", CalcDate('<-1Y>', Starts.Get(Index)), CalcDate('<-1Y>', Ends.Get(Index)));
                        GLAccount.CalcFields("Net Change");
                        LYValues.Add(GLAccount."Net Change");
                    end;
                end;
                RowJson.Add('budget', BudgetValues);
                if ShowLY then
                    RowJson.Add('ly', LYValues);
                Rows.Add(RowJson);
            until GLAccount.Next() = 0;

        Response.Add('requestId', GetText(Request, 'requestId', false));
        Response.Add('budgetName', BudgetName);
        Response.Add('budgetBlocked', GLBudgetName.Blocked);
        Response.Add('periods', PeriodsJson);
        Response.Add('rows', Rows);
        Response.Add('canWrite', CanWrite() and not GLBudgetName.Blocked);
        Response.WriteTo(Result);
        exit(Result);
    end;

    /// <summary>Response when a load fails, so the add-in can show the reason and stop waiting.</summary>
    procedure BuildErrorResponse(RequestJson: Text; Message: Text): Text
    var
        Request: JsonObject;
        Response: JsonObject;
        Result: Text;
    begin
        if Request.ReadFrom(RequestJson) then
            Response.Add('requestId', GetText(Request, 'requestId', false));
        Response.Add('error', Message);
        Response.WriteTo(Result);
        exit(Result);
    end;

    /// <summary>Checks a budget save before anything is written (no database writes, safe inside a TryFunction).</summary>
    procedure CheckWrite(RequestJson: Text)
    var
        GLBudgetName: Record "G/L Budget Name";
        Request: JsonObject;
        Token: JsonToken;
        BudgetName: Code[10];
    begin
        if not CanWrite() then
            Error(NoBudgetWriteErr);
        Request.ReadFrom(RequestJson);
        BudgetName := CopyStr(GetText(Request, 'budgetName', true), 1, MaxStrLen(BudgetName));
        GLBudgetName.Get(BudgetName);
        GLBudgetName.TestField(Blocked, false);
        if not Request.Get('changes', Token) then
            Error(InvalidRequestErr, 'changes');
        if Token.AsArray().Count() > 5000 then
            Error(TooManyCellsErr, 5000);
    end;

    /// <summary>Saves changed cells. Request: requestId, budgetName, dim1Filter, dim2Filter, changes[{ row, col, account, start, end, amount }].</summary>
    procedure WriteChanges(RequestJson: Text): Text
    var
        BudgetWrite: Codeunit "AK Workbook Budget Write";
        Request: JsonObject;
        Response: JsonObject;
        Results: JsonArray;
        Item: JsonObject;
        Change: JsonObject;
        Token: JsonToken;
        ChangeToken: JsonToken;
        Saved: Integer;
        Failed: Integer;
        Result: Text;
    begin
        CheckWrite(RequestJson);
        Request.ReadFrom(RequestJson);
        Request.Get('changes', Token);
        foreach ChangeToken in Token.AsArray() do begin
            Change := ChangeToken.AsObject();
            Clear(BudgetWrite);
            BudgetWrite.SetChange(
                CopyStr(GetText(Request, 'budgetName', true), 1, 10),
                GetText(Request, 'dim1Filter', false), GetText(Request, 'dim2Filter', false), Change);
            Clear(Item);
            Item.Add('row', GetInteger(Change, 'row'));
            Item.Add('col', GetInteger(Change, 'col'));
            ClearLastError();
            if BudgetWrite.Run() then begin
                Commit();
                Item.Add('ok', true);
                Item.Add('amount', BudgetWrite.GetNewAmount());
                Saved += 1;
            end else begin
                Item.Add('ok', false);
                if GetLastErrorText() = '' then
                    Item.Add('error', UnknownErrorTxt)
                else
                    Item.Add('error', GetLastErrorText());
                Failed += 1;
            end;
            Results.Add(Item);
        end;
        Response.Add('requestId', GetText(Request, 'requestId', false));
        Response.Add('results', Results);
        Response.Add('saved', Saved);
        Response.Add('failed', Failed);
        Response.WriteTo(Result);
        Session.LogMessage('AK0003', StrSubstNo('Workbook budget write: %1 saved, %2 failed.', Saved, Failed),
            Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', 'AK Excel Workbook');
        exit(Result);
    end;

    local procedure BuildPeriods(StartDate: Date; PeriodCount: Integer; Quarterly: Boolean; var Starts: List of [Date]; var Ends: List of [Date]; var PeriodsJson: JsonArray)
    var
        PeriodJson: JsonObject;
        PeriodStart: Date;
        PeriodEnd: Date;
        Index: Integer;
        PeriodLabel: Text;
    begin
        PeriodStart := CalcDate('<-CM>', StartDate);
        for Index := 1 to PeriodCount do begin
            if Quarterly then
                PeriodEnd := CalcDate('<3M-1D>', PeriodStart)
            else
                PeriodEnd := CalcDate('<CM>', PeriodStart);
            if Quarterly then
                PeriodLabel := StrSubstNo(QuarterLabelTxt, ((Date2DMY(PeriodStart, 2) - 1) div 3) + 1, Format(PeriodStart, 0, '<Year>'))
            else
                PeriodLabel := Format(PeriodStart, 0, PeriodLabelTxt);
            Starts.Add(PeriodStart);
            Ends.Add(PeriodEnd);
            Clear(PeriodJson);
            PeriodJson.Add('start', Format(PeriodStart, 0, 9));
            PeriodJson.Add('end', Format(PeriodEnd, 0, 9));
            PeriodJson.Add('label', PeriodLabel);
            PeriodsJson.Add(PeriodJson);
            PeriodStart := PeriodEnd + 1;
        end;
    end;

    local procedure FiscalYearStart(ForDate: Date): Date
    var
        AccountingPeriod: Record "Accounting Period";
    begin
        AccountingPeriod.SetRange("New Fiscal Year", true);
        AccountingPeriod.SetRange("Starting Date", 0D, ForDate);
        if AccountingPeriod.FindLast() then
            exit(AccountingPeriod."Starting Date");
        exit(CalcDate('<-CY>', ForDate));
    end;

    local procedure DimensionCaption(DimensionCode: Code[20]; GlobalNo: Integer): Text
    var
        Dimension: Record Dimension;
    begin
        if (DimensionCode <> '') and Dimension.Get(DimensionCode) then
            exit(Dimension.GetMLCodeCaption(GlobalLanguage()));
        exit(StrSubstNo('Global Dimension %1', GlobalNo));
    end;

    local procedure GetText(JObject: JsonObject; PropertyName: Text; Mandatory: Boolean): Text
    var
        Token: JsonToken;
    begin
        if not JObject.Get(PropertyName, Token) then begin
            if Mandatory then
                Error(InvalidRequestErr, PropertyName);
            exit('');
        end;
        if Token.AsValue().IsNull() then
            exit('');
        exit(Token.AsValue().AsText());
    end;

    local procedure GetInteger(JObject: JsonObject; PropertyName: Text): Integer
    var
        Token: JsonToken;
    begin
        if not JObject.Get(PropertyName, Token) then
            exit(0);
        if Token.AsValue().IsNull() then
            exit(0);
        exit(Token.AsValue().AsInteger());
    end;

    local procedure GetBoolean(JObject: JsonObject; PropertyName: Text): Boolean
    var
        Token: JsonToken;
    begin
        if not JObject.Get(PropertyName, Token) then
            exit(false);
        if Token.AsValue().IsNull() then
            exit(false);
        exit(Token.AsValue().AsBoolean());
    end;
}
