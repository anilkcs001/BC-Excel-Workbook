permissionset 89000 "AK Workbook"
{
    Caption = 'AK Excel Workbook';
    Assignable = true;

    // Grants access to the workbook feature only. Reading BC data inside a workbook still
    // requires the user's normal permissions on each source table (e.g. D365 BASIC, D365 ACC. PAYABLE).
    Permissions =
        tabledata "AK Workbook" = RIMD,
        table "AK Workbook" = X,
        page "AK Workbook List" = X,
        page "AK Workbook Editor" = X,
        codeunit "AK Workbook Data Service" = X,
        codeunit "AK Workbook Mgt" = X;
}
