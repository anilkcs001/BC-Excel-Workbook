codeunit 89003 "AK Workbook Write Row"
{
    // Writes ONE changed or new sheet row to its table. Called with WriteRow.Run() by
    // "AK Workbook Write Service", so an error here rolls back only this row.
    //
    // Modify: find the record by its primary key, stop if someone changed it after the sheet was loaded
    //         (SystemModifiedAt), Validate each changed field in sheet order, then Modify(true).
    // Insert: Init, Validate the key fields that were typed, give line tables the next Line No. and
    //         journal lines the page's SetUpNewLine defaults, Insert(true) (number series and OnInsert run),
    //         then Validate the other fields and Modify(true).
    // Validate runs the same OnValidate code as the standard page, so lookups, blocked checks and
    // dependent fields behave the same way.

    Access = Internal;

    trigger OnRun()
    begin
        case Operation of
            'modify':
                ModifyRow();
            'insert':
                InsertRow();
            else
                Error(UnknownOperationErr, Operation);
        end;
        BuildResult();
    end;

    var
        DataService: Codeunit "AK Workbook Data Service";
        WriteService: Codeunit "AK Workbook Write Service";
        RecRef: RecordRef;
        Change: JsonObject;
        ResultJson: JsonObject;
        FieldNos: List of [Integer];
        KeyFieldNos: List of [Integer];
        TableNo: Integer;
        RowNo: Integer;
        Operation: Text;
        UnknownOperationErr: Label 'Unknown change type "%1".', Comment = '%1 = operation';
        RecordGoneErr: Label 'This record no longer exists in Business Central. Refresh the sheet.';
        ChangedByOtherErr: Label 'This record was changed in Business Central after you loaded the sheet. Refresh the sheet and enter your change again.';
        FieldReadOnlyErr: Label '%1 cannot be changed from a workbook.', Comment = '%1 = field caption';
        TooLongErr: Label '%1 can be at most %2 characters.', Comment = '%1 = field caption, %2 = maximum length';
        BadValueErr: Label '"%1" is not a valid value for %2.', Comment = '%1 = value typed by the user, %2 = field caption';
        UnsupportedTypeErr: Label '%1 (%2) cannot be changed from a workbook.', Comment = '%1 = field caption, %2 = field type';
        KeyMissingErr: Label 'The key value for %1 is missing.', Comment = '%1 = field caption';

    procedure SetChange(NewTableNo: Integer; NewFieldNos: List of [Integer]; NewKeyFieldNos: List of [Integer]; NewChange: JsonObject)
    var
        Token: JsonToken;
    begin
        TableNo := NewTableNo;
        FieldNos := NewFieldNos;
        KeyFieldNos := NewKeyFieldNos;
        Change := NewChange;
        if Change.Get('row', Token) then
            RowNo := Token.AsValue().AsInteger();
        if Change.Get('op', Token) then
            Operation := Token.AsValue().AsText();
    end;

    procedure GetResult(): JsonObject
    begin
        exit(ResultJson);
    end;

    local procedure ModifyRow()
    var
        KeyValues: List of [Text];
        Index: Integer;
        FldRef: FieldRef;
    begin
        WriteService.CheckCanModify(TableNo);
        RecRef.Open(TableNo);
        ReadTextArray('key', KeyValues);
        for Index := 1 to KeyFieldNos.Count() do begin
            if Index > KeyValues.Count() then
                Error(KeyMissingErr, RecRef.Field(KeyFieldNos.Get(Index)).Caption());
            FldRef := RecRef.Field(KeyFieldNos.Get(Index));
            AssignValue(FldRef, KeyValues.Get(Index), false);
            FldRef.SetRange(FldRef.Value());
        end;
        if not RecRef.FindFirst() then
            Error(RecordGoneErr);
        CheckVersion();
        ValidateValues(false);
        RecRef.Modify(true);
    end;

    local procedure InsertRow()
    var
        KeyValues: List of [Text];
        Index: Integer;
        FldRef: FieldRef;
    begin
        WriteService.CheckCanInsert(TableNo);
        RecRef.Open(TableNo);
        RecRef.Init();
        ReadTextArray('key', KeyValues);
        for Index := 1 to KeyFieldNos.Count() do
            if Index <= KeyValues.Count() then
                if KeyValues.Get(Index) <> '' then begin
                    FldRef := RecRef.Field(KeyFieldNos.Get(Index));
                    AssignValue(FldRef, KeyValues.Get(Index), true);
                end;
        SetNextLineNo(KeyValues);
        SetUpNewLine();
        RecRef.Insert(true);
        ValidateValues(true);
        RecRef.Modify(true);
    end;

    /// <summary>
    /// Line tables (journal, sales and purchase lines): when the last key field is an Integer (Line No.) and
    /// the sheet left it empty, use the last line no. with the same other key values + 10000, like the page does.
    /// </summary>
    local procedure SetNextLineNo(KeyValues: List of [Text])
    var
        LastRef: RecordRef;
        FldRef: FieldRef;
        LastIndex: Integer;
        Index: Integer;
        LastLineNo: Integer;
    begin
        LastIndex := KeyFieldNos.Count();
        if LastIndex < 2 then
            exit;
        FldRef := RecRef.Field(KeyFieldNos.Get(LastIndex));
        if FldRef.Type() <> FieldType::Integer then
            exit;
        if LastIndex <= KeyValues.Count() then
            if KeyValues.Get(LastIndex) <> '' then
                exit;
        LastRef.Open(TableNo);
        for Index := 1 to LastIndex - 1 do
            LastRef.Field(KeyFieldNos.Get(Index)).SetRange(RecRef.Field(KeyFieldNos.Get(Index)).Value());
        LastLineNo := 0;
        if LastRef.FindLast() then
            LastLineNo := LastRef.Field(KeyFieldNos.Get(LastIndex)).Value();
        FldRef.Value := LastLineNo + 10000;
    end;

    /// <summary>Journal lines get the same defaults as a new line on the journal page (posting date, document no., source code...).</summary>
    local procedure SetUpNewLine()
    var
        GenJnlLine: Record "Gen. Journal Line";
        LastGenJnlLine: Record "Gen. Journal Line";
        ItemJnlLine: Record "Item Journal Line";
        LastItemJnlLine: Record "Item Journal Line";
    begin
        case TableNo of
            Database::"Gen. Journal Line":
                begin
                    RecRef.SetTable(GenJnlLine);
                    LastGenJnlLine.SetRange("Journal Template Name", GenJnlLine."Journal Template Name");
                    LastGenJnlLine.SetRange("Journal Batch Name", GenJnlLine."Journal Batch Name");
                    if LastGenJnlLine.FindLast() then;
                    GenJnlLine.SetUpNewLine(LastGenJnlLine, 0, true);
                    RecRef.GetTable(GenJnlLine);
                end;
            Database::"Item Journal Line":
                begin
                    RecRef.SetTable(ItemJnlLine);
                    LastItemJnlLine.SetRange("Journal Template Name", ItemJnlLine."Journal Template Name");
                    LastItemJnlLine.SetRange("Journal Batch Name", ItemJnlLine."Journal Batch Name");
                    if LastItemJnlLine.FindLast() then;
                    ItemJnlLine.SetUpNewLine(LastItemJnlLine);
                    RecRef.GetTable(ItemJnlLine);
                end;
        end;
    end;

    local procedure CheckVersion()
    var
        Token: JsonToken;
        Expected: Text;
    begin
        if not Change.Get('version', Token) then
            exit;
        if Token.AsValue().IsNull() then
            exit;
        Expected := Token.AsValue().AsText();
        if Expected = '' then
            exit;
        if DataService.GetRowVersion(RecRef) <> Expected then
            Error(ChangedByOtherErr);
    end;

    local procedure ValidateValues(IsInsert: Boolean)
    var
        ValuesToken: JsonToken;
        ValueToken: JsonToken;
        Item: JsonObject;
        Token: JsonToken;
        FieldNo: Integer;
        ValueText: Text;
    begin
        if not Change.Get('values', ValuesToken) then
            exit;
        foreach ValueToken in ValuesToken.AsArray() do begin
            Item := ValueToken.AsObject();
            Item.Get('no', Token);
            FieldNo := Token.AsValue().AsInteger();
            ValueText := '';
            if Item.Get('value', Token) then
                if not Token.AsValue().IsNull() then
                    ValueText := Token.AsValue().AsText();
            // Key fields: set before Insert(true) for new rows, never changed on existing rows.
            if KeyFieldNos.Contains(FieldNo) then begin
                if not IsInsert then
                    Error(FieldReadOnlyErr, RecRef.Field(FieldNo).Caption());
            end else
                if not DataService.IsFieldEditable(RecRef.Field(FieldNo)) then
                    Error(FieldReadOnlyErr, RecRef.Field(FieldNo).Caption())
                else
                    AssignValue(RecRef.Field(FieldNo), ValueText, true);
        end;
    end;

    /// <summary>Converts the sheet text to the field type and assigns it, with Validate when RunValidate is true.</summary>
    local procedure AssignValue(FldRef: FieldRef; ValueText: Text; RunValidate: Boolean)
    var
        IntegerValue: Integer;
        BigIntegerValue: BigInteger;
        DecimalValue: Decimal;
        BooleanValue: Boolean;
        DateValue: Date;
        DateTimeValue: DateTime;
        TimeValue: Time;
        DurationValue: Duration;
        DateFormulaValue: DateFormula;
        GuidValue: Guid;
        NewValue: Variant;
    begin
        ValueText := DelChr(ValueText, '<>', ' ');
        case FldRef.Type() of
            FieldType::Code:
                begin
                    if StrLen(ValueText) > FldRef.Length() then
                        Error(TooLongErr, FldRef.Caption(), FldRef.Length());
                    NewValue := UpperCase(ValueText);
                end;
            FieldType::Text:
                begin
                    if StrLen(ValueText) > FldRef.Length() then
                        Error(TooLongErr, FldRef.Caption(), FldRef.Length());
                    NewValue := ValueText;
                end;
            FieldType::Integer:
                begin
                    if (ValueText <> '') and not Evaluate(IntegerValue, ValueText, 9) then
                        Error(BadValueErr, ValueText, FldRef.Caption());
                    NewValue := IntegerValue;
                end;
            FieldType::BigInteger:
                begin
                    if (ValueText <> '') and not Evaluate(BigIntegerValue, ValueText, 9) then
                        Error(BadValueErr, ValueText, FldRef.Caption());
                    NewValue := BigIntegerValue;
                end;
            FieldType::Decimal:
                begin
                    if (ValueText <> '') and not Evaluate(DecimalValue, ValueText, 9) then
                        Error(BadValueErr, ValueText, FldRef.Caption());
                    NewValue := DecimalValue;
                end;
            FieldType::Boolean:
                begin
                    if (ValueText <> '') and not Evaluate(BooleanValue, LowerCase(ValueText), 9) then
                        Error(BadValueErr, ValueText, FldRef.Caption());
                    NewValue := BooleanValue;
                end;
            FieldType::Date:
                begin
                    if (ValueText <> '') and not Evaluate(DateValue, ValueText, 9) then
                        Error(BadValueErr, ValueText, FldRef.Caption());
                    NewValue := DateValue;
                end;
            FieldType::DateTime:
                begin
                    if (ValueText <> '') and not Evaluate(DateTimeValue, ValueText, 9) then
                        Error(BadValueErr, ValueText, FldRef.Caption());
                    NewValue := DateTimeValue;
                end;
            FieldType::Time:
                begin
                    if (ValueText <> '') and not Evaluate(TimeValue, ValueText, 9) then
                        Error(BadValueErr, ValueText, FldRef.Caption());
                    NewValue := TimeValue;
                end;
            FieldType::Duration:
                begin
                    if (ValueText <> '') and not Evaluate(DurationValue, ValueText, 9) then
                        Error(BadValueErr, ValueText, FldRef.Caption());
                    NewValue := DurationValue;
                end;
            FieldType::DateFormula:
                begin
                    if (ValueText <> '') and not Evaluate(DateFormulaValue, ValueText) then
                        Error(BadValueErr, ValueText, FldRef.Caption());
                    NewValue := DateFormulaValue;
                end;
            FieldType::Guid:
                begin
                    if (ValueText <> '') and not Evaluate(GuidValue, ValueText) then
                        Error(BadValueErr, ValueText, FldRef.Caption());
                    NewValue := GuidValue;
                end;
            FieldType::Option:
                NewValue := FindOptionOrdinal(FldRef, ValueText);
            else
                Error(UnsupportedTypeErr, FldRef.Caption(), Format(FldRef.Type()));
        end;
        if RunValidate then
            FldRef.Validate(NewValue)
        else
            FldRef.Value := NewValue;
    end;

    /// <summary>Option and Enum fields: accepts the caption shown in the sheet, the AL name or the ordinal number.</summary>
    local procedure FindOptionOrdinal(FldRef: FieldRef; ValueText: Text): Integer
    var
        Index: Integer;
        Ordinal: Integer;
    begin
        for Index := 1 to FldRef.EnumValueCount() do
            if (LowerCase(FldRef.GetEnumValueCaption(Index)) = LowerCase(ValueText)) or
               (LowerCase(FldRef.GetEnumValueName(Index)) = LowerCase(ValueText))
            then
                exit(FldRef.GetEnumValueOrdinal(Index));
        if Evaluate(Ordinal, ValueText) then
            for Index := 1 to FldRef.EnumValueCount() do
                if FldRef.GetEnumValueOrdinal(Index) = Ordinal then
                    exit(Ordinal);
        Error(BadValueErr, ValueText, FldRef.Caption());
    end;

    local procedure ReadTextArray(PropertyName: Text; var Values: List of [Text])
    var
        Token: JsonToken;
        Item: JsonToken;
    begin
        if not Change.Get(PropertyName, Token) then
            exit;
        foreach Item in Token.AsArray() do
            if Item.AsValue().IsNull() then
                Values.Add('')
            else
                Values.Add(Item.AsValue().AsText());
    end;

    /// <summary>Sends back the saved row as BC now has it, so the sheet shows values that Validate changed too.</summary>
    local procedure BuildResult()
    var
        KeyValues: JsonArray;
        FieldNo: Integer;
    begin
        RecRef.Find('=');
        foreach FieldNo in KeyFieldNos do
            KeyValues.Add(Format(RecRef.Field(FieldNo).Value()));
        ResultJson.Add('row', RowNo);
        ResultJson.Add('ok', true);
        ResultJson.Add('op', Operation);
        ResultJson.Add('key', KeyValues);
        ResultJson.Add('version', DataService.GetRowVersion(RecRef));
        ResultJson.Add('values', DataService.RowToJson(RecRef, FieldNos));
    end;
}
