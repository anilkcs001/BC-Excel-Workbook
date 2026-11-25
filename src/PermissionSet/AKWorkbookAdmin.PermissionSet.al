permissionset 89001 "AK Workbook Admin"
{
    Caption = 'AK Excel Workbook Admin';
    Assignable = true;
    IncludedPermissionSets = "AK Workbook";

    // For administrators: maintain the list of tables users may edit from a workbook sheet.
    Permissions =
        tabledata "AK Workbook Edit Table" = RIMD,
        page "AK Workbook Edit Tables" = X;
}
