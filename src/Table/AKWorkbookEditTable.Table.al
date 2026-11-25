table 89001 "AK Workbook Edit Table"
{
    // Allow-list of tables whose data users may change from a workbook sheet (Stage 1 write-back).
    // A table that is not in this list is always read-only in the workbook, whatever the user's permissions.
    // Ledger entries, posted documents, registers and archives cannot be added (see CheckTableAllowed).
    Caption = 'AK Workbook Editable Table';
    DataClassification = CustomerContent;
    LookupPageId = "AK Workbook Edit Tables";
    DrillDownPageId = "AK Workbook Edit Tables";

    fields
    {
        field(1; "Table No."; Integer)
        {
            Caption = 'Table No.';
            NotBlank = true;
            TableRelation = AllObjWithCaption."Object ID" where("Object Type" = const(Table));

            trigger OnValidate()
            begin
                CheckTableAllowed("Table No.");
            end;
        }
        field(2; "Table Caption"; Text[249])
        {
            Caption = 'Table Caption';
            FieldClass = FlowField;
            CalcFormula = lookup(AllObjWithCaption."Object Caption" where("Object Type" = const(Table), "Object ID" = field("Table No.")));
            Editable = false;
        }
        field(10; "Allow Modify"; Boolean)
        {
            Caption = 'Allow Modify';
            InitValue = true;
        }
        field(11; "Allow Insert"; Boolean)
        {
            Caption = 'Allow Insert';
        }
    }

    keys
    {
        key(PK; "Table No.")
        {
            Clustered = true;
        }
    }

    var
        SystemTableErr: Label 'System tables cannot be edited from a workbook.';
        ProtectedTableErr: Label '%1 (table %2) holds posted or ledger data and cannot be edited from a workbook. Use the standard journals and documents instead.', Comment = '%1 = table caption, %2 = table number';
        TableNotFoundErr: Label 'Table %1 does not exist.', Comment = '%1 = table number';

    trigger OnInsert()
    begin
        CheckTableAllowed("Table No.");
    end;

    /// <summary>Errors when a table must never be written from a workbook (system, ledger, posted, register, archive).</summary>
    procedure CheckTableAllowed(TableNo: Integer)
    var
        AllObj: Record AllObjWithCaption;
    begin
        if TableNo >= 2000000000 then
            Error(SystemTableErr);
        if not AllObj.Get(AllObj."Object Type"::Table, TableNo) then
            Error(TableNotFoundErr, TableNo);
        if IsProtectedName(AllObj."Object Name") then
            Error(ProtectedTableErr, AllObj."Object Caption", TableNo);
    end;

    local procedure IsProtectedName(ObjectName: Text): Boolean
    var
        Name: Text;
    begin
        Name := LowerCase(ObjectName);
        exit(
            Name.Contains('ledger entry') or
            Name.Contains('posted') or
            Name.Contains('register') or
            Name.Contains('archive') or
            Name.Contains('value entry') or
            Name.Contains('vat entry') or
            Name.Contains('g/l entry') or
            Name.Contains('change log') or
            Name.StartsWith('sales invoice') or
            Name.StartsWith('sales cr.memo') or
            Name.StartsWith('sales shipment') or
            Name.StartsWith('purch. inv.') or
            Name.StartsWith('purch. cr. memo') or
            Name.StartsWith('purch. rcpt.') or
            Name.StartsWith('return receipt') or
            Name.StartsWith('return shipment'));
    end;
}
