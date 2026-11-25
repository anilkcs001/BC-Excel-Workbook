codeunit 89006 "AK Workbook Budget Write"
{
    // Writes ONE budget cell. Called with BudgetWrite.Run() by "AK Workbook Budget", so an error here rolls
    // back only this cell. The cell's new amount minus the current budget for the same account, period,
    // budget and dimension filter is added as one G/L Budget Entry dated on the first day of the period.

    Access = Internal;

    trigger OnRun()
    begin
        SaveCell();
    end;

    var
        Change: JsonObject;
        BudgetName: Code[10];
        Dim1Filter: Text;
        Dim2Filter: Text;
        NewAmount: Decimal;
        NotPostingErr: Label 'G/L account %1 is not a posting account. Budgets are entered on posting accounts.', Comment = '%1 = G/L account no.';
        SingleDimErr: Label 'To change budgets, filter %1 on one value (or leave it empty).', Comment = '%1 = dimension caption';
        BadValueErr: Label '"%1" is not a valid amount.', Comment = '%1 = value typed by the user';
        EntryDescriptionTxt: Label 'Excel Workbook';

    procedure SetChange(NewBudgetName: Code[10]; NewDim1Filter: Text; NewDim2Filter: Text; NewChange: JsonObject)
    begin
        BudgetName := NewBudgetName;
        Dim1Filter := NewDim1Filter;
        Dim2Filter := NewDim2Filter;
        Change := NewChange;
    end;

    procedure GetNewAmount(): Decimal
    begin
        exit(NewAmount);
    end;

    local procedure SaveCell()
    var
        GLAccount: Record "G/L Account";
        GLBudgetEntry: Record "G/L Budget Entry";
        LastEntry: Record "G/L Budget Entry";
        GLSetup: Record "General Ledger Setup";
        AccountNo: Code[20];
        StartDate: Date;
        EndDate: Date;
        Current: Decimal;
        Delta: Decimal;
        AmountText: Text;
    begin
        AccountNo := CopyStr(GetText('account'), 1, MaxStrLen(AccountNo));
        Evaluate(StartDate, GetText('start'), 9);
        Evaluate(EndDate, GetText('end'), 9);
        AmountText := GetText('amount');
        if AmountText = '' then
            AmountText := '0';
        if not Evaluate(NewAmount, AmountText, 9) then
            Error(BadValueErr, AmountText);

        GLAccount.Get(AccountNo);
        if GLAccount."Account Type" <> GLAccount."Account Type"::Posting then
            Error(NotPostingErr, AccountNo);
        GLSetup.Get();
        CheckSingleValue(Dim1Filter, GLSetup."Global Dimension 1 Code");
        CheckSingleValue(Dim2Filter, GLSetup."Global Dimension 2 Code");

        // Current budget exactly as the grid shows it (same budget, period and dimension filters).
        GLAccount.SetFilter("Budget Filter", BudgetName);
        GLAccount.SetRange("Date Filter", StartDate, EndDate);
        if Dim1Filter <> '' then
            GLAccount.SetFilter("Global Dimension 1 Filter", Dim1Filter);
        if Dim2Filter <> '' then
            GLAccount.SetFilter("Global Dimension 2 Filter", Dim2Filter);
        GLAccount.CalcFields("Budgeted Amount");
        Current := GLAccount."Budgeted Amount";
        Delta := NewAmount - Current;
        if Delta = 0 then
            exit;

        LastEntry.LockTable();
        if LastEntry.FindLast() then;
        GLBudgetEntry.Init();
        GLBudgetEntry."Entry No." := LastEntry."Entry No." + 1;
        GLBudgetEntry.Validate("Budget Name", BudgetName);
        GLBudgetEntry.Validate("G/L Account No.", AccountNo);
        GLBudgetEntry.Validate(Date, StartDate);
        if Dim1Filter <> '' then
            GLBudgetEntry.Validate("Global Dimension 1 Code", CopyStr(Dim1Filter, 1, MaxStrLen(GLBudgetEntry."Global Dimension 1 Code")));
        if Dim2Filter <> '' then
            GLBudgetEntry.Validate("Global Dimension 2 Code", CopyStr(Dim2Filter, 1, MaxStrLen(GLBudgetEntry."Global Dimension 2 Code")));
        GLBudgetEntry.Validate(Amount, Delta);
        GLBudgetEntry.Description := EntryDescriptionTxt;
        GLBudgetEntry.Insert(true);
    end;

    /// <summary>A dimension filter must be empty or one plain value, so the new entry gets a clear dimension value.</summary>
    local procedure CheckSingleValue(FilterText: Text; DimensionCode: Code[20])
    begin
        if FilterText = '' then
            exit;
        if StrLen(DelChr(FilterText, '=', '|&<>*@.()=''')) <> StrLen(FilterText) then
            Error(SingleDimErr, DimensionCode);
    end;

    local procedure GetText(PropertyName: Text): Text
    var
        Token: JsonToken;
    begin
        if not Change.Get(PropertyName, Token) then
            exit('');
        if Token.AsValue().IsNull() then
            exit('');
        exit(Token.AsValue().AsText());
    end;
}
