page 89002 "AK Workbook Edit Tables"
{
    // Admin page: which tables users may edit from an Excel Workbook sheet.
    Caption = 'Excel Workbook Editable Tables';
    PageType = List;
    SourceTable = "AK Workbook Edit Table";
    UsageCategory = Administration;
    ApplicationArea = All;
    AboutTitle = 'Editable tables';
    AboutText = 'Only tables in this list can be changed from a workbook sheet. Every change still runs the field validation of the table, and users still need their normal permission to modify or insert records.';

    layout
    {
        area(Content)
        {
            repeater(Tables)
            {
                field("Table No."; Rec."Table No.")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the table that users may edit from a workbook.';
                }
                field("Table Caption"; Rec."Table Caption")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the name of the table.';
                }
                field("Allow Modify"; Rec."Allow Modify")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies whether users may change existing records of this table from a workbook.';
                }
                field("Allow Insert"; Rec."Allow Insert")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies whether users may create new records of this table by typing new rows below the data.';
                }
            }
        }
    }
}
