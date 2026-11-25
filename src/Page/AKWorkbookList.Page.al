page 89000 "AK Workbook List"
{
    // Create a workbook by entering a code and name on a new line, then choose Open.
    // The editor page shows only the spreadsheet.
    Caption = 'Excel Workbooks';
    PageType = List;
    SourceTable = "AK Workbook";
    UsageCategory = Lists;
    ApplicationArea = All;
    Editable = true;

    layout
    {
        area(Content)
        {
            repeater(Workbooks)
            {
                field("Code"; Rec."Code")
                {
                    ToolTip = 'Specifies the code of the workbook.';
                }
                field(Description; Rec.Description)
                {
                    ToolTip = 'Specifies the name of the workbook. Also used as the export file name.';
                }
                field("Linked Tables"; Rec."Linked Tables")
                {
                    ToolTip = 'Specifies the Business Central tables that feed this workbook.';
                    Editable = false;
                }
                field("Data Size (KB)"; Rec."Data Size (KB)")
                {
                    ToolTip = 'Specifies the size of the saved workbook.';
                    Editable = false;
                }
                field("Last Saved At"; Rec."Last Saved At")
                {
                    ToolTip = 'Specifies when the workbook was last saved.';
                    Editable = false;
                }
                field("Last Saved By"; Rec."Last Saved By")
                {
                    ToolTip = 'Specifies who last saved the workbook.';
                    Editable = false;
                }
            }
        }
    }

    actions
    {
        area(Processing)
        {
            action(OpenWorkbook)
            {
                Caption = 'Open';
                ToolTip = 'Open the workbook in the full-page spreadsheet.';
                Image = Edit;
                Scope = Repeater;

                trigger OnAction()
                begin
                    Rec.TestField("Code");
                    CurrPage.SaveRecord();
                    WorkbookMgt.OpenWorkbook(Rec."Code");
                end;
            }
        }
        area(Promoted)
        {
            actionref(OpenWorkbook_Promoted; OpenWorkbook) { }
        }
    }

    var
        WorkbookMgt: Codeunit "AK Workbook Mgt";
}
