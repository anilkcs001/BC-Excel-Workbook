pageextension 89003 "AK Purchase Order" extends "Purchase Order"
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
                ToolTip = 'Edit the lines of this purchase order in an Excel Workbook: type or paste many lines, then Save to BC. Each line is validated like on the order.';

                trigger OnAction()
                var
                    WorkbookLines: Codeunit "AK Workbook Lines";
                begin
                    WorkbookLines.OpenPurchaseDocument(Rec);
                end;
            }
        }
    }
}
