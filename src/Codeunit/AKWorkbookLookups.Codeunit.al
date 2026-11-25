codeunit 89009 "AK Workbook Lookups"
{
    // Stage 8: dropdown lists for the editable columns of a region, like the lookup on a BC page.
    // Request (purpose 'lookups'):  { requestId, tableNo, fields: [fieldNo...] }
    // Response (purpose 'lookupresult'):
    //   { requestId, tableNo,
    //     lookups: [ { field, kind: 'list', values: [caption...] }            Option / Enum fields
    //                { field, kind: 'list', list: '<list key>' }              Code field with a table relation
    //                { field, kind: 'byType', typeField, blankType, map: { '<type caption>': '<list key>' } } ],
    //     lists: { '<list key>': { caption, values: [[code, description]...], truncated } } }
    // "byType" covers fields whose relation depends on another field on the same line:
    // Account No. / Bal. Account No. on journal lines and No. on sales and purchase lines.
    // Only tables the user can read are listed; a list stops at MaxValues() entries. The add-in keeps
    // "allow other values" on, so a code that is not in the list is still checked by BC on save.

    Access = Public;

    procedure MaxValues(): Integer
    begin
        exit(1000);
    end;

    procedure GetLookupsJson(RequestJson: Text): Text
    var
        Request: JsonObject;
        Response: JsonObject;
        Lookup: JsonObject;
        Lookups: JsonArray;
        Lists: JsonObject;
        Token: JsonToken;
        FieldToken: JsonToken;
        RecRef: RecordRef;
        TableNo: Integer;
        FieldNo: Integer;
        Result: Text;
    begin
        Request.ReadFrom(RequestJson);
        Request.Get('tableNo', Token);
        TableNo := Token.AsValue().AsInteger();
        RecRef.Open(TableNo);
        if Request.Get('fields', Token) then
            foreach FieldToken in Token.AsArray() do begin
                FieldNo := FieldToken.AsValue().AsInteger();
                if RecRef.FieldExist(FieldNo) then begin
                    Clear(Lookup);
                    if BuildLookup(RecRef, FieldNo, Lookup, Lists) then
                        Lookups.Add(Lookup);
                end;
            end;
        RecRef.Close();
        if Request.Get('requestId', Token) then
            Response.Add('requestId', Token.AsValue().AsText());
        Response.Add('tableNo', TableNo);
        Response.Add('lookups', Lookups);
        Response.Add('lists', Lists);
        Response.WriteTo(Result);
        exit(Result);
    end;

    local procedure BuildLookup(var RecRef: RecordRef; FieldNo: Integer; var Lookup: JsonObject; var Lists: JsonObject): Boolean
    var
        FldRef: FieldRef;
        Values: JsonArray;
        Map: JsonObject;
        ListKey: Text;
        Caption: Text;
        Index: Integer;
        TypeFieldNo: Integer;
        RelationTableNo: Integer;
        RelationFieldNo: Integer;
    begin
        FldRef := RecRef.Field(FieldNo);
        Lookup.Add('field', FieldNo);
        case FldRef.Type() of
            FieldType::Option:
                begin
                    for Index := 1 to FldRef.EnumValueCount() do begin
                        Caption := FldRef.GetEnumValueCaption(Index);
                        if Caption.Trim() <> '' then
                            Values.Add(Caption);
                    end;
                    Lookup.Add('kind', 'list');
                    Lookup.Add('values', Values);
                    exit(Values.Count() > 0);
                end;
            FieldType::Code:
                begin
                    if TypeDependentRelation(RecRef.Number(), FieldNo, TypeFieldNo) then begin
                        if not BuildTypeMap(RecRef.Field(TypeFieldNo), Map, Lists) then
                            exit(false);
                        Lookup.Add('kind', 'byType');
                        Lookup.Add('typeField', TypeFieldNo);
                        // A new line with an empty type cell gets the type's default value in BC (first enum value).
                        Lookup.Add('blankType', RecRef.Field(TypeFieldNo).GetEnumValueCaption(1));
                        Lookup.Add('map', Map);
                        exit(true);
                    end;
                    if not GetRelation(RecRef.Number(), FieldNo, RelationTableNo, RelationFieldNo) then
                        exit(false);
                    ListKey := AddList(RelationTableNo, RelationFieldNo, DimensionCodeFor(FldRef, RelationTableNo), Lists);
                    if ListKey = '' then
                        exit(false);
                    Lookup.Add('kind', 'list');
                    Lookup.Add('list', ListKey);
                    exit(true);
                end;
        end;
        exit(false);
    end;

    /// <summary>Fields whose table relation follows a type field on the same line (conditional TableRelation).</summary>
    local procedure TypeDependentRelation(TableNo: Integer; FieldNo: Integer; var TypeFieldNo: Integer): Boolean
    var
        GenJournalLine: Record "Gen. Journal Line";
        SalesLine: Record "Sales Line";
        PurchaseLine: Record "Purchase Line";
    begin
        case TableNo of
            Database::"Gen. Journal Line":
                case FieldNo of
                    GenJournalLine.FieldNo("Account No."):
                        TypeFieldNo := GenJournalLine.FieldNo("Account Type");
                    GenJournalLine.FieldNo("Bal. Account No."):
                        TypeFieldNo := GenJournalLine.FieldNo("Bal. Account Type");
                end;
            Database::"Sales Line":
                if FieldNo = SalesLine.FieldNo("No.") then
                    TypeFieldNo := SalesLine.FieldNo(Type);
            Database::"Purchase Line":
                if FieldNo = PurchaseLine.FieldNo("No.") then
                    TypeFieldNo := PurchaseLine.FieldNo(Type);
        end;
        exit(TypeFieldNo <> 0);
    end;

    /// <summary>Type caption (as shown in the sheet) -> list of the table that type points to.</summary>
    local procedure BuildTypeMap(TypeFldRef: FieldRef; var Map: JsonObject; var Lists: JsonObject): Boolean
    var
        Index: Integer;
        TableNo: Integer;
        ListKey: Text;
    begin
        for Index := 1 to TypeFldRef.EnumValueCount() do begin
            TableNo := TableForType(TypeFldRef.GetEnumValueName(Index));
            if TableNo <> 0 then begin
                ListKey := AddList(TableNo, 0, '', Lists);
                if ListKey <> '' then
                    Map.Add(TypeFldRef.GetEnumValueCaption(Index), ListKey);
            end;
        end;
        exit(Map.Keys().Count() > 0);
    end;

    /// <summary>Enum value names are the same in every language, so the match is made on the name.</summary>
    local procedure TableForType(TypeName: Text): Integer
    begin
        case TypeName of
            'G/L Account':
                exit(Database::"G/L Account");
            'Customer':
                exit(Database::Customer);
            'Vendor':
                exit(Database::Vendor);
            'Bank Account':
                exit(Database::"Bank Account");
            'Fixed Asset':
                exit(Database::"Fixed Asset");
            'Employee':
                exit(Database::Employee);
            'IC Partner':
                exit(Database::"IC Partner");
            'Item':
                exit(Database::Item);
            'Resource':
                exit(Database::Resource);
            'Charge (Item)':
                exit(Database::"Item Charge");
        end;
        exit(0);
    end;

    local procedure GetRelation(TableNo: Integer; FieldNo: Integer; var RelationTableNo: Integer; var RelationFieldNo: Integer): Boolean
    var
        Field: Record Field;
    begin
        if not Field.Get(TableNo, FieldNo) then
            exit(false);
        RelationTableNo := Field.RelationTableNo;
        RelationFieldNo := Field.RelationFieldNo;
        exit(RelationTableNo <> 0);
    end;

    /// <summary>Shortcut / global dimension N fields relate to Dimension Value; list only that dimension's values.</summary>
    local procedure DimensionCodeFor(FldRef: FieldRef; RelationTableNo: Integer): Code[20]
    var
        GLSetup: Record "General Ledger Setup";
        FieldName: Text;
        Position: Integer;
        DimensionNo: Integer;
    begin
        if RelationTableNo <> Database::"Dimension Value" then
            exit('');
        FieldName := FldRef.Name();
        Position := StrPos(FieldName, 'Dimension ');
        if Position = 0 then
            exit('');
        if not Evaluate(DimensionNo, CopyStr(FieldName, Position + StrLen('Dimension '), 1)) then
            exit('');
        if not GLSetup.Get() then
            exit('');
        case DimensionNo of
            1:
                exit(GLSetup."Shortcut Dimension 1 Code");
            2:
                exit(GLSetup."Shortcut Dimension 2 Code");
            3:
                exit(GLSetup."Shortcut Dimension 3 Code");
            4:
                exit(GLSetup."Shortcut Dimension 4 Code");
            5:
                exit(GLSetup."Shortcut Dimension 5 Code");
            6:
                exit(GLSetup."Shortcut Dimension 6 Code");
            7:
                exit(GLSetup."Shortcut Dimension 7 Code");
            8:
                exit(GLSetup."Shortcut Dimension 8 Code");
        end;
        exit('');
    end;

    /// <summary>Reads the code (and description or name) of a related table once per response. Returns the list key, or '' when it cannot be listed.</summary>
    local procedure AddList(TableNo: Integer; FieldNo: Integer; DimensionCode: Code[20]; var Lists: JsonObject): Text
    var
        GLAccount: Record "G/L Account";
        DimensionValue: Record "Dimension Value";
        RecRef: RecordRef;
        CodeRef: FieldRef;
        KeyRef: KeyRef;
        ListJson: JsonObject;
        Values: JsonArray;
        Pair: JsonArray;
        ListKey: Text;
        DescriptionFieldNo: Integer;
        BlockedFieldNo: Integer;
        ValueCount: Integer;
        Truncated: Boolean;
    begin
        ListKey := Format(TableNo) + ':' + Format(FieldNo) + ':' + DimensionCode;
        if Lists.Contains(ListKey) then
            exit(ListKey);
        RecRef.Open(TableNo);
        if not RecRef.ReadPermission() then
            exit('');
        if FieldNo = 0 then begin
            // No related field given: the relation is to the primary key, which must be one field.
            KeyRef := RecRef.KeyIndex(1);
            if KeyRef.FieldCount() <> 1 then
                exit('');
            FieldNo := KeyRef.FieldIndex(1).Number();
        end;
        CodeRef := RecRef.Field(FieldNo);
        DescriptionFieldNo := FindFieldByName(TableNo, 'Description');
        if DescriptionFieldNo = 0 then
            DescriptionFieldNo := FindFieldByName(TableNo, 'Name');
        BlockedFieldNo := FindFieldByName(TableNo, 'Blocked');
        if BlockedFieldNo <> 0 then
            if RecRef.Field(BlockedFieldNo).Type() = FieldType::Boolean then
                RecRef.Field(BlockedFieldNo).SetRange(false);
        if TableNo = Database::"G/L Account" then
            RecRef.Field(GLAccount.FieldNo("Account Type")).SetRange(GLAccount."Account Type"::Posting);
        if (TableNo = Database::"Dimension Value") and (DimensionCode <> '') then begin
            RecRef.Field(DimensionValue.FieldNo("Dimension Code")).SetRange(DimensionCode);
            RecRef.Field(DimensionValue.FieldNo("Dimension Value Type")).SetRange(DimensionValue."Dimension Value Type"::Standard);
        end;
        if RecRef.FindSet() then
            repeat
                ValueCount += 1;
                if ValueCount > MaxValues() then
                    Truncated := true
                else begin
                    Clear(Pair);
                    Pair.Add(Format(CodeRef.Value()));
                    if DescriptionFieldNo <> 0 then
                        Pair.Add(Format(RecRef.Field(DescriptionFieldNo).Value()))
                    else
                        Pair.Add('');
                    Values.Add(Pair);
                end;
            until (RecRef.Next() = 0) or Truncated;
        ListJson.Add('caption', RecRef.Caption());
        ListJson.Add('values', Values);
        ListJson.Add('truncated', Truncated);
        Lists.Add(ListKey, ListJson);
        RecRef.Close();
        exit(ListKey);
    end;

    local procedure FindFieldByName(TableNo: Integer; FieldName: Text): Integer
    var
        Field: Record Field;
    begin
        Field.SetRange(TableNo, TableNo);
        Field.SetRange(FieldName, FieldName);
        Field.SetRange(Class, Field.Class::Normal);
        if Field.FindFirst() then
            exit(Field."No.");
        exit(0);
    end;
}
