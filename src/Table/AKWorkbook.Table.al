table 89000 "AK Workbook"
{
    Caption = 'AK Workbook';
    DataClassification = CustomerContent;
    LookupPageId = "AK Workbook List";
    DrillDownPageId = "AK Workbook List";

    fields
    {
        field(1; "Code"; Code[20])
        {
            Caption = 'Code';
            NotBlank = true;
        }
        field(2; Description; Text[100])
        {
            Caption = 'Description';
        }
        field(10; "Workbook Data"; Blob)
        {
            Caption = 'Workbook Data';
        }
        field(11; "Has Data"; Boolean)
        {
            Caption = 'Has Data';
            Editable = false;
        }
        field(12; "Data Size (KB)"; Integer)
        {
            Caption = 'Data Size (KB)';
            Editable = false;
        }
        field(20; "Last Saved At"; DateTime)
        {
            Caption = 'Last Saved At';
            Editable = false;
        }
        field(21; "Last Saved By"; Code[50])
        {
            Caption = 'Last Saved By';
            Editable = false;
            DataClassification = EndUserIdentifiableInformation;
        }
        field(30; "Linked Tables"; Text[250])
        {
            Caption = 'Linked Tables';
            Editable = false;
        }
    }

    keys
    {
        key(PK; "Code")
        {
            Clustered = true;
        }
    }

    fieldgroups
    {
        fieldgroup(DropDown; "Code", Description) { }
        fieldgroup(Brick; "Code", Description, "Last Saved At") { }
    }
}
