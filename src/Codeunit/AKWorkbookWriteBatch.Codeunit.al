codeunit 89010 "AK Workbook Write Batch"
{
    // Writes a batch of sheet rows (up to 100) in ONE transaction. Called with Batch.Run() by
    // "AK Workbook Write Service": when every row is fine there is one commit for the whole batch, which is
    // much faster than a commit per row. When any row fails, Run() returns false, everything in the batch is
    // rolled back, and the write service writes that batch again row by row to report each row's own error.
    // Every row still goes through Validate and the table triggers, exactly like a single row.

    Access = Internal;

    trigger OnRun()
    var
        WriteRow: Codeunit "AK Workbook Write Row";
        ChangeToken: JsonToken;
        Index: Integer;
    begin
        Clear(Results);
        for Index := FromIndex to ToIndex do begin
            Changes.Get(Index, ChangeToken);
            Clear(WriteRow);
            WriteRow.SetChange(TableNo, FieldNos, KeyFieldNos, ChangeToken.AsObject());
            WriteRow.SetAccessChecked(AccessChecked);
            WriteRow.Execute();
            Results.Add(WriteRow.GetResult());
        end;
    end;

    var
        Changes: JsonArray;
        Results: JsonArray;
        FieldNos: List of [Integer];
        KeyFieldNos: List of [Integer];
        TableNo: Integer;
        FromIndex: Integer;
        ToIndex: Integer;
        AccessChecked: Boolean;

    procedure SetBatch(NewTableNo: Integer; NewFieldNos: List of [Integer]; NewKeyFieldNos: List of [Integer]; NewChanges: JsonArray; NewFromIndex: Integer; NewToIndex: Integer; NewAccessChecked: Boolean)
    begin
        TableNo := NewTableNo;
        FieldNos := NewFieldNos;
        KeyFieldNos := NewKeyFieldNos;
        Changes := NewChanges;
        FromIndex := NewFromIndex;
        ToIndex := NewToIndex;
        AccessChecked := NewAccessChecked;
    end;

    procedure GetResults(): JsonArray
    begin
        exit(Results);
    end;
}
