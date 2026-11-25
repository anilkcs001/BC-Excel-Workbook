codeunit 89008 "AK Workbook Formulas"
{
    // Stage 6: BC functions inside spreadsheet formulas. The add-in collects the calls made while the sheet
    // calculates and sends them in one batch (purpose 'formulas' -> 'formularesult'):
    //   BCGLSUM(account; [period]; [dim1]; [dim2])                 net change from G/L entries
    //   BCBUDGET(budget; account; [period]; [dim1]; [dim2])        budgeted amount from G/L budget entries
    //   BCSUM(table; field; [filter field; filter]...)             sum of a normal decimal field
    //   BCCOUNT(table; [filter field; filter]...)                  number of records
    //   BCVALUE(table; field; [filter field; filter]...)           field of the first record (a lookup)
    // Tables and fields can be given by number, name or caption. Every call reads as the current user,
    // so permissions apply. Each call is evaluated on its own; an error is returned as text for that cell.

    Access = Public;

    var
        UnknownFunctionErr: Label 'Unknown function %1.', Comment = '%1 = function name';
        TableNotFoundErr: Label 'Table "%1" was not found.', Comment = '%1 = table name or number';
        FieldNotFoundErr: Label 'Field "%1" was not found in %2.', Comment = '%1 = field name, %2 = table caption';
        NoReadPermissionErr: Label 'No permission to read %1.', Comment = '%1 = table caption';
        NotSummableErr: Label '%1 is not a normal number field, so it cannot be summed.', Comment = '%1 = field caption';
        MissingArgumentErr: Label '%1 needs at least %2 arguments.', Comment = '%1 = function name, %2 = number of arguments';
        PairsErr: Label 'Filters come in pairs: field, then filter.';
        TooManyCallsErr: Label 'Too many BC functions at once (more than %1).', Comment = '%1 = maximum';

    procedure MaxCalls(): Integer
    begin
        exit(2000);
    end;

    /// <summary>Request: { requestId, calls: [{ id, fn, args: [text...] }] }. Response: { requestId, results: [{ id, value | error }] }.</summary>
    procedure EvaluateCalls(RequestJson: Text): Text
    var
        Request: JsonObject;
        Response: JsonObject;
        Results: JsonArray;
        Item: JsonObject;
        CallsToken: JsonToken;
        CallToken: JsonToken;
        Token: JsonToken;
        ResultValue: JsonValue;
        Result: Text;
    begin
        Request.ReadFrom(RequestJson);
        Request.Get('calls', CallsToken);
        if CallsToken.AsArray().Count() > MaxCalls() then
            Error(TooManyCallsErr, MaxCalls());
        foreach CallToken in CallsToken.AsArray() do begin
            Clear(Item);
            Clear(ResultValue);
            CallToken.AsObject().Get('id', Token);
            Item.Add('id', Token.AsValue().AsText());
            ClearLastError();
            if TryCall(CallToken.AsObject(), ResultValue) then
                Item.Add('value', ResultValue)
            else
                Item.Add('error', GetLastErrorText());
            Results.Add(Item);
        end;
        if Request.Get('requestId', Token) then
            Response.Add('requestId', Token.AsValue().AsText());
        Response.Add('results', Results);
        Response.WriteTo(Result);
        exit(Result);
    end;

    [TryFunction]
    local procedure TryCall(Call: JsonObject; var ResultValue: JsonValue)
    var
        Args: List of [Text];
        Token: JsonToken;
        ArgToken: JsonToken;
        Fn: Text;
    begin
        Call.Get('fn', Token);
        Fn := UpperCase(Token.AsValue().AsText());
        if Call.Get('args', Token) then
            foreach ArgToken in Token.AsArray() do
                if ArgToken.AsValue().IsNull() then
                    Args.Add('')
                else
                    Args.Add(ArgToken.AsValue().AsText());
        case Fn of
            'BCGLSUM':
                ResultValue.SetValue(GLSum(Args));
            'BCBUDGET':
                ResultValue.SetValue(BudgetSum(Args));
            'BCSUM':
                ResultValue.SetValue(TableSum(Args));
            'BCCOUNT':
                ResultValue.SetValue(TableCount(Args));
            'BCVALUE':
                TableValue(Args, ResultValue);
            else
                Error(UnknownFunctionErr, Fn);
        end;
    end;

    local procedure GLSum(Args: List of [Text]): Decimal
    var
        GLEntry: Record "G/L Entry";
        DataService: Codeunit "AK Workbook Data Service";
    begin
        NeedArgs('BCGLSUM', Args, 1);
        GLEntry.SetFilter("G/L Account No.", AccountFilter(Arg(Args, 1)));
        if Arg(Args, 2) <> '' then
            GLEntry.SetFilter("Posting Date", DataService.LocalizeIsoDates(Arg(Args, 2)));
        if Arg(Args, 3) <> '' then
            GLEntry.SetFilter("Global Dimension 1 Code", Arg(Args, 3));
        if Arg(Args, 4) <> '' then
            GLEntry.SetFilter("Global Dimension 2 Code", Arg(Args, 4));
        GLEntry.CalcSums(Amount);
        exit(GLEntry.Amount);
    end;

    local procedure BudgetSum(Args: List of [Text]): Decimal
    var
        GLBudgetEntry: Record "G/L Budget Entry";
        DataService: Codeunit "AK Workbook Data Service";
    begin
        NeedArgs('BCBUDGET', Args, 2);
        GLBudgetEntry.SetFilter("Budget Name", Arg(Args, 1));
        GLBudgetEntry.SetFilter("G/L Account No.", AccountFilter(Arg(Args, 2)));
        if Arg(Args, 3) <> '' then
            GLBudgetEntry.SetFilter(Date, DataService.LocalizeIsoDates(Arg(Args, 3)));
        if Arg(Args, 4) <> '' then
            GLBudgetEntry.SetFilter("Global Dimension 1 Code", Arg(Args, 4));
        if Arg(Args, 5) <> '' then
            GLBudgetEntry.SetFilter("Global Dimension 2 Code", Arg(Args, 5));
        GLBudgetEntry.CalcSums(Amount);
        exit(GLBudgetEntry.Amount);
    end;

    local procedure TableSum(Args: List of [Text]): Decimal
    var
        RecRef: RecordRef;
        FldRef: FieldRef;
        Total: Decimal;
    begin
        NeedArgs('BCSUM', Args, 2);
        OpenTable(RecRef, Arg(Args, 1));
        FldRef := RecRef.Field(FindField(RecRef, Arg(Args, 2)));
        if (FldRef.Class() <> FieldClass::Normal) or not (FldRef.Type() in [FieldType::Decimal, FieldType::Integer, FieldType::BigInteger]) then
            Error(NotSummableErr, FldRef.Caption());
        ApplyPairs(RecRef, Args, 3);
        FldRef.CalcSum();
        Total := FldRef.Value();
        exit(Total);
    end;

    local procedure TableCount(Args: List of [Text]): Integer
    var
        RecRef: RecordRef;
    begin
        NeedArgs('BCCOUNT', Args, 1);
        OpenTable(RecRef, Arg(Args, 1));
        ApplyPairs(RecRef, Args, 2);
        exit(RecRef.Count());
    end;

    local procedure TableValue(Args: List of [Text]; var ResultValue: JsonValue)
    var
        RecRef: RecordRef;
        FldRef: FieldRef;
        DecimalValue: Decimal;
    begin
        NeedArgs('BCVALUE', Args, 2);
        OpenTable(RecRef, Arg(Args, 1));
        FldRef := RecRef.Field(FindField(RecRef, Arg(Args, 2)));
        ApplyPairs(RecRef, Args, 3);
        if not RecRef.FindFirst() then begin
            ResultValue.SetValue('');
            exit;
        end;
        if FldRef.Class() = FieldClass::FlowField then
            FldRef.CalcField();
        if FldRef.Type() in [FieldType::Decimal, FieldType::Integer, FieldType::BigInteger] then begin
            DecimalValue := FldRef.Value();
            ResultValue.SetValue(DecimalValue);
        end else
            ResultValue.SetValue(Format(FldRef.Value()));
    end;

    local procedure OpenTable(var RecRef: RecordRef; TableText: Text)
    var
        AllObj: Record AllObjWithCaption;
        TableNo: Integer;
    begin
        AllObj.SetRange("Object Type", AllObj."Object Type"::Table);
        if Evaluate(TableNo, TableText) then
            AllObj.SetRange("Object ID", TableNo)
        else begin
            AllObj.SetFilter("Object Name", '@' + DelChr(TableText, '=', '*?|&<>=()''"'));
            if AllObj.IsEmpty() then begin
                AllObj.SetRange("Object Name");
                AllObj.SetFilter("Object Caption", '@' + DelChr(TableText, '=', '*?|&<>=()''"'));
            end;
        end;
        if not AllObj.FindFirst() then
            Error(TableNotFoundErr, TableText);
        RecRef.Open(AllObj."Object ID");
        if not RecRef.ReadPermission() then
            Error(NoReadPermissionErr, RecRef.Caption());
    end;

    local procedure FindField(var RecRef: RecordRef; FieldText: Text): Integer
    var
        Field: Record Field;
        FieldNo: Integer;
    begin
        if Evaluate(FieldNo, FieldText) then
            if RecRef.FieldExist(FieldNo) then
                exit(FieldNo);
        Field.SetRange(TableNo, RecRef.Number());
        Field.SetFilter(FieldName, '@' + DelChr(FieldText, '=', '*?|&<>=()''"'));
        if Field.FindFirst() then
            exit(Field."No.");
        Field.SetRange(FieldName);
        Field.SetFilter("Field Caption", '@' + DelChr(FieldText, '=', '*?|&<>=()''"'));
        if Field.FindFirst() then
            exit(Field."No.");
        Error(FieldNotFoundErr, FieldText, RecRef.Caption());
    end;

    /// <summary>Arguments from StartIndex on are pairs: field, filter (BC filter syntax; ISO dates are localized).</summary>
    local procedure ApplyPairs(var RecRef: RecordRef; Args: List of [Text]; StartIndex: Integer)
    var
        DataService: Codeunit "AK Workbook Data Service";
        FldRef: FieldRef;
        Index: Integer;
        FilterText: Text;
    begin
        if Args.Count() < StartIndex then
            exit;
        if (Args.Count() - StartIndex + 1) mod 2 <> 0 then
            Error(PairsErr);
        Index := StartIndex;
        while Index < Args.Count() do begin
            FldRef := RecRef.Field(FindField(RecRef, Args.Get(Index)));
            FilterText := Args.Get(Index + 1);
            if FldRef.Type() in [FieldType::Date, FieldType::DateTime] then
                FilterText := DataService.LocalizeIsoDates(FilterText);
            FldRef.SetFilter(FilterText);
            Index += 2;
        end;
    end;

    /// <summary>One totaling account (Total revenue) sums the accounts in its Totaling; anything else is used as a filter.</summary>
    local procedure AccountFilter(AccountText: Text): Text
    var
        GLAccount: Record "G/L Account";
    begin
        if StrLen(AccountText) <= MaxStrLen(GLAccount."No.") then
            if GLAccount.Get(CopyStr(AccountText, 1, MaxStrLen(GLAccount."No."))) then
                if (GLAccount."Account Type" <> GLAccount."Account Type"::Posting) and (GLAccount.Totaling <> '') then
                    exit(GLAccount.Totaling);
        exit(AccountText);
    end;

    local procedure NeedArgs(Fn: Text; Args: List of [Text]; MinCount: Integer)
    begin
        if Args.Count() < MinCount then
            Error(MissingArgumentErr, Fn, MinCount);
    end;

    local procedure Arg(Args: List of [Text]; Index: Integer): Text
    begin
        if Index > Args.Count() then
            exit('');
        exit(Args.Get(Index));
    end;
}
