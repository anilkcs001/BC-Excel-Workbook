permissionset 89000 "AK Workbook"
{
    Caption = 'AK Excel Workbook';
    Assignable = true;

    // Grants access to the workbook feature only. Reading BC data inside a workbook still
    // requires the user's normal permissions on each source table (e.g. D365 BASIC, D365 ACC. PAYABLE).
    // Editing from a sheet also needs the table on the Editable Tables page (admin) and the user's normal
    // modify/insert permission on that table.
    Permissions =
        tabledata "AK Workbook" = RIMD,
        tabledata "AK Workbook Edit Table" = R,
        table "AK Workbook" = X,
        table "AK Workbook Edit Table" = X,
        page "AK Workbook List" = X,
        page "AK Workbook Editor" = X,
        codeunit "AK Workbook Data Service" = X,
        codeunit "AK Workbook Mgt" = X,
        codeunit "AK Workbook Write Service" = X,
        codeunit "AK Workbook Write Row" = X,
        codeunit "AK Workbook Lines" = X,
        codeunit "AK Workbook Budget" = X,
        codeunit "AK Workbook Budget Write" = X,
        codeunit "AK Workbook Drill Down" = X;
}
