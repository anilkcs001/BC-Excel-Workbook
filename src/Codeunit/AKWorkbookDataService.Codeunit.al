codeunit 89000 "AK Workbook Data Service"
{
    // Reads metadata and data from any BC table through RecordRef and returns JSON for the
    // AK Excel Workbook control add-in. All reads run as the signed-in user, so standard
    // BC permissions and security filters apply. Nothing in this codeunit writes data.

    Access = Public;

    var
        NoReadPermissionErr: Label 'You do not have permission to read %1 (table %2).', Comment = '%1 = table caption, %2 = table number';
        TableNotFoundErr: Label 'Table %1 does not exist.', Comment = '%1 = table number';
        NoFieldsSelectedErr: Label 'Select at least one field to load.';
        InvalidRequestErr: Label 'The data request is missing the property "%1".', Comment = '%1 = JSON property name';

    procedure GetConfigJson(): Text
    var
        Config: JsonObject;
        Result: Text;
    begin
        Config.Add('defaultMaxRows', DefaultMaxRows());
        Config.Add('maxRowsLimit', MaxRowsLimit());
        Config.Add('chunkSize', ChunkSize());
        Config.Add('userId', UserId());
        Config.Add('companyName', CompanyName());
        Config.WriteTo(Result);
        exit(Result);
    end;

    procedure GetTablesJson(SearchText: Text): Text
    var
        AllObj: Record AllObjWithCaption;
        Tables: JsonArray;
        CleanText: Text;
        TableNo: Integer;
        Counter: Integer;
        Result: Text;
    begin
        CleanText := DelChr(SearchText, '<>', ' ');
        CleanText := DelChr(CleanText, '=', '()|&<>=*@''"');

        AllObj.SetRange("Object Type", AllObj."Object Type"::Table);
        AllObj.SetFilter("Object ID", '<%1', 2000000000);
        if CleanText <> '' then
            if Evaluate(TableNo, CleanText) then
                AllObj.SetRange("Object ID", TableNo)
            else
                AllObj.SetFilter("Object Caption", '@*' + CleanText + '*');
        AddTables(AllObj, Tables, Counter);

        // Fall back to the technical name when no caption matched (e.g. "GSTLedger").
        if (Counter = 0) and (CleanText <> '') and (TableNo = 0) then begin
            AllObj.SetRange("Object Caption");
            AllObj.SetFilter("Object Name", '@*' + CleanText + '*');
            AddTables(AllObj, Tables, Counter);
        end;

        Tables.WriteTo(Result);
        exit(Result);
    end;

    procedure GetFieldsJson(TableNo: Integer): Text
    var
        Field: Record Field;
        RecRef: RecordRef;
        PrimaryKey: KeyRef;
        PkFields: List of [Integer];
        Fields: JsonArray;
        FieldJson: JsonObject;
        Response: JsonObject;
        Index: Integer;
        Result: Text;
    begin
        EnsureTableExists(TableNo);
        RecRef.Open(TableNo);
        PrimaryKey := RecRef.KeyIndex(1);
        for Index := 1 to PrimaryKey.FieldCount() do
            PkFields.Add(PrimaryKey.FieldIndex(Index).Number());

        Field.SetRange(TableNo, TableNo);
        Field.SetRange(Enabled, true);
        Field.SetFilter(ObsoleteState, '<>%1', Field.ObsoleteState::Removed);
        Field.SetFilter(Class, '<>%1', Field.Class::FlowFilter);
        Field.SetFilter(Type, '<>%1&<>%2&<>%3&<>%4', Field.Type::BLOB, Field.Type::Media, Field.Type::MediaSet, Field.Type::TableFilter);
        if Field.FindSet() then
            repeat
                Clear(FieldJson);
                FieldJson.Add('no', Field."No.");
                FieldJson.Add('name', Field.FieldName);
                FieldJson.Add('caption', Field."Field Caption");
                FieldJson.Add('type', Format(Field.Type));
                FieldJson.Add('flow', Field.Class = Field.Class::FlowField);
                FieldJson.Add('pk', PkFields.Contains(Field."No."));
                Fields.Add(FieldJson);
            until Field.Next() = 0;

        Response.Add('tableNo', TableNo);
        Response.Add('tableCaption', RecRef.Caption());
        Response.Add('canRead', RecRef.ReadPermission());
        Response.Add('fields', Fields);
        RecRef.Close();
        Response.WriteTo(Result);
        exit(Result);
    end;

    procedure BuildDataChunks(RequestJson: Text; var Chunks: List of [Text])
    var
        RecRef: RecordRef;
        FldRef: FieldRef;
        Request: JsonObject;
        Columns: JsonArray;
        Rows: JsonArray;
        FieldNos: List of [Integer];
        FieldNo: Integer;
        RequestId: Text;
        TableNo: Integer;
        MaxRows: Integer;
        Total: Integer;
        RowCount: Integer;
        RowsInChunk: Integer;
        ChunkNo: Integer;
    begin
        Request.ReadFrom(RequestJson);
        RequestId := GetJsonText(Request, 'requestId', true);
        TableNo := GetJsonInteger(Request, 'tableNo', true);
        MaxRows := GetJsonInteger(Request, 'maxRows', false);
        if (MaxRows <= 0) or (MaxRows > MaxRowsLimit()) then
            MaxRows := DefaultMaxRows();

        EnsureTableExists(TableNo);
        RecRef.Open(TableNo);
        if not RecRef.ReadPermission() then
            Error(NoReadPermissionErr, RecRef.Caption(), TableNo);

        ReadFieldList(Request, RecRef, FieldNos);
        if FieldNos.Count() = 0 then
            Error(NoFieldsSelectedErr);

        foreach FieldNo in FieldNos do begin
            FldRef := RecRef.Field(FieldNo);
            if FldRef.Class() = FieldClass::Normal then
                RecRef.AddLoadFields(FieldNo);
            Columns.Add(BuildColumnJson(FldRef));
        end;

        ApplyFilters(Request, RecRef);
        Total := RecRef.Count();

        if RecRef.FindSet() then
            repeat
                Rows.Add(BuildRow(RecRef, FieldNos));
                RowCount += 1;
                RowsInChunk += 1;
                if RowsInChunk >= ChunkRows() then begin
                    Chunks.Add(BuildChunk(RequestId, ChunkNo, false, Total, RowCount, MaxRows, RecRef, Columns, Rows));
                    Clear(Rows);
                    RowsInChunk := 0;
                    ChunkNo += 1;
                end;
            until (RecRef.Next() = 0) or (RowCount >= MaxRows);

        // The last chunk is always sent (even when empty) so the add-in knows the load finished.
        Chunks.Add(BuildChunk(RequestId, ChunkNo, true, Total, RowCount, MaxRows, RecRef, Columns, Rows));
        RecRef.Close();
    end;

    procedure GetRequestId(RequestJson: Text): Text
    var
        Request: JsonObject;
    begin
        if not Request.ReadFrom(RequestJson) then
            exit('');
        exit(GetJsonText(Request, 'requestId', false));
    end;

    procedure DefaultMaxRows(): Integer
    begin
        exit(10000);
    end;

    procedure MaxRowsLimit(): Integer
    begin
        exit(100000);
    end;

    procedure ChunkSize(): Integer
    begin
        // Characters per add-in call in either direction. Keeps each call well below client limits.
        exit(400000);
    end;

    local procedure ChunkRows(): Integer
    begin
        exit(2000);
    end;

    local procedure MaxTableResults(): Integer
    begin
        exit(60);
    end;

    local procedure AddTables(var AllObj: Record AllObjWithCaption; var Tables: JsonArray; var Counter: Integer)
    var
        TableJson: JsonObject;
    begin
        if not AllObj.FindSet() then
            exit;
        repeat
            Clear(TableJson);
            TableJson.Add('id', AllObj."Object ID");
            TableJson.Add('name', AllObj."Object Name");
            TableJson.Add('caption', AllObj."Object Caption");
            Tables.Add(TableJson);
            Counter += 1;
        until (AllObj.Next() = 0) or (Counter >= MaxTableResults());
    end;

    local procedure EnsureTableExists(TableNo: Integer)
    var
        AllObj: Record AllObjWithCaption;
    begin
        if not AllObj.Get(AllObj."Object Type"::Table, TableNo) then
            Error(TableNotFoundErr, TableNo);
    end;

    local procedure ReadFieldList(Request: JsonObject; var RecRef: RecordRef; var FieldNos: List of [Integer])
    var
        FieldsToken: JsonToken;
        FieldToken: JsonToken;
        FieldNo: Integer;
    begin
        if not Request.Get('fields', FieldsToken) then
            Error(InvalidRequestErr, 'fields');
        foreach FieldToken in FieldsToken.AsArray() do begin
            FieldNo := FieldToken.AsValue().AsInteger();
            if RecRef.FieldExist(FieldNo) and not FieldNos.Contains(FieldNo) then
                FieldNos.Add(FieldNo);
        end;
    end;

    local procedure ApplyFilters(Request: JsonObject; var RecRef: RecordRef)
    var
        FiltersToken: JsonToken;
        FilterToken: JsonToken;
        FilterObject: JsonObject;
        FldRef: FieldRef;
        FieldNo: Integer;
        FilterText: Text;
    begin
        if not Request.Get('filters', FiltersToken) then
            exit;
        foreach FilterToken in FiltersToken.AsArray() do begin
            FilterObject := FilterToken.AsObject();
            FieldNo := GetJsonInteger(FilterObject, 'fieldNo', true);
            FilterText := GetJsonText(FilterObject, 'filter', false);
            if (FilterText <> '') and RecRef.FieldExist(FieldNo) then begin
                FldRef := RecRef.Field(FieldNo);
                FldRef.SetFilter(FilterText);
            end;
        end;
    end;

    local procedure BuildColumnJson(FldRef: FieldRef): JsonObject
    var
        Column: JsonObject;
    begin
        Column.Add('no', FldRef.Number());
        Column.Add('name', FldRef.Name());
        Column.Add('caption', FldRef.Caption());
        Column.Add('type', Format(FldRef.Type()));
        Column.Add('flow', FldRef.Class() = FieldClass::FlowField);
        exit(Column);
    end;

    local procedure BuildChunk(RequestId: Text; ChunkNo: Integer; IsLast: Boolean; Total: Integer; Loaded: Integer; MaxRows: Integer; var RecRef: RecordRef; Columns: JsonArray; Rows: JsonArray): Text
    var
        Chunk: JsonObject;
        Result: Text;
    begin
        Chunk.Add('requestId', RequestId);
        Chunk.Add('chunk', ChunkNo);
        Chunk.Add('isLast', IsLast);
        Chunk.Add('total', Total);
        Chunk.Add('loaded', Loaded);
        Chunk.Add('truncated', Total > MaxRows);
        if ChunkNo = 0 then begin
            Chunk.Add('tableNo', RecRef.Number());
            Chunk.Add('tableCaption', RecRef.Caption());
            Chunk.Add('filterText', RecRef.GetFilters());
            Chunk.Add('columns', Columns);
        end;
        Chunk.Add('rows', Rows);
        Chunk.WriteTo(Result);
        exit(Result);
    end;

    local procedure BuildRow(var RecRef: RecordRef; FieldNos: List of [Integer]): JsonArray
    var
        FldRef: FieldRef;
        Row: JsonArray;
        FieldNo: Integer;
    begin
        foreach FieldNo in FieldNos do begin
            FldRef := RecRef.Field(FieldNo);
            if FldRef.Class() = FieldClass::FlowField then
                FldRef.CalcField();
            AddFieldValue(Row, FldRef);
        end;
        exit(Row);
    end;

    local procedure AddFieldValue(var Row: JsonArray; FldRef: FieldRef)
    var
        IntegerValue: Integer;
        BigIntegerValue: BigInteger;
        DecimalValue: Decimal;
        BooleanValue: Boolean;
        DateValue: Date;
        DateTimeValue: DateTime;
    begin
        case FldRef.Type() of
            FieldType::Integer:
                begin
                    IntegerValue := FldRef.Value();
                    Row.Add(IntegerValue);
                end;
            FieldType::BigInteger:
                begin
                    BigIntegerValue := FldRef.Value();
                    Row.Add(BigIntegerValue);
                end;
            FieldType::Decimal:
                begin
                    DecimalValue := FldRef.Value();
                    Row.Add(DecimalValue);
                end;
            FieldType::Boolean:
                begin
                    BooleanValue := FldRef.Value();
                    Row.Add(BooleanValue);
                end;
            FieldType::Date:
                begin
                    DateValue := FldRef.Value();
                    if DateValue = 0D then
                        Row.Add('')
                    else
                        Row.Add(Format(DateValue, 0, 9));
                end;
            FieldType::DateTime:
                begin
                    DateTimeValue := FldRef.Value();
                    if DateTimeValue = 0DT then
                        Row.Add('')
                    else
                        Row.Add(Format(DateTimeValue, 0, 9));
                end;
            else
                Row.Add(Format(FldRef.Value()));
        end;
    end;

    local procedure GetJsonText(JObject: JsonObject; PropertyName: Text; Mandatory: Boolean): Text
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

    local procedure GetJsonInteger(JObject: JsonObject; PropertyName: Text; Mandatory: Boolean): Integer
    var
        Token: JsonToken;
    begin
        if not JObject.Get(PropertyName, Token) then begin
            if Mandatory then
                Error(InvalidRequestErr, PropertyName);
            exit(0);
        end;
        if Token.AsValue().IsNull() then
            exit(0);
        exit(Token.AsValue().AsInteger());
    end;
}
