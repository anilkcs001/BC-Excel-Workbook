pageextension 89000 "AK General Journal" extends "General Journal"
{
    // Stage 2: open these lines in an Excel Workbook region (editable when the table is on the Editable Tables list).
    actions
    {
        addlast(Processing)
        {
            action(AKEditInWorkbook)
            {
                ApplicationArea = All;
                Caption = 'Edit in Excel Workbook';
                Image = ExportToExcel;
                ToolTip = 'Edit the lines of this journal batch in an Excel Workbook: type or paste many lines, then Save to BC. Each line is validated like on this page.';

                trigger OnAction()
                var
                    WorkbookLines: Codeunit "AK Workbook Lines";
                begin
                    WorkbookLines.OpenGeneralJournal(Rec);
                end;
            }
        }
    }
}
