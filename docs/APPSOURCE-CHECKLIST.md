# AppSource checklist – AK Excel Workbook

Code analysis is switched on in `.vscode/settings.json` (CodeCop, UICop, AppSourceCop) with `AppSourceCop.json` and `AK.ruleset.json`.
Open the project in VS Code and look at the **Problems** panel (Ctrl+Shift+M). Fix every AppSourceCop error first.

## Must be done before submission (cannot be done in code alone)

| # | Item | Status |
|---|------|--------|
| 1 | Partner Center account (Microsoft AI Cloud Partner Program) and an MPN/Partner ID | To do |
| 2 | Register a **3-letter affix** with d365val@microsoft.com (send Partner ID, publisher name, 5 suggestions). `AK` is 2 letters and cannot be registered. | To do |
| 3 | Request an **AppSource object ID range** (the current 89000–89049 range is for per-tenant extensions only) | To do |
| 4 | Rename every object and field with the new affix, renumber into the new range, update `AppSourceCop.json` and `app.json` | After 2 and 3 |
| 5 | Final `publisher` name in `app.json` must match the offer in Partner Center | To do |
| 6 | Code-sign the .app file with a code-signing certificate | To do |
| 7 | `privacyStatement` and `EULA` URLs in `app.json` (pages must exist) | To do |
| 8 | `logo` (app icon) and offer listing in Partner Center (description, screenshots, video, contact, support URL) | To do |
| 9 | `applicationInsightsConnectionString` in `app.json` so validation results reach your telemetry | Recommended |
| 10 | Install and Upgrade codeunits (needed from the second version on) | Before v2 on AppSource |
| 11 | Publish / install / uninstall / reinstall test in a BC SaaS sandbox, and run `Run-AlValidation` (BcContainerHelper) | To do |
| 12 | Third-party notices: Univer (Apache-2.0) and ExcelJS (MIT) are bundled; keep their licence texts with the app | Listed in README |

## Code items already in place
- `target: Cloud`, `application` and `platform` set, `NoImplicitWith` and `TranslationFile` features.
- Permission sets `AK Workbook` and `AK Workbook Admin` (users never need SUPER).
- Data classification on all table fields; ApplicationArea on page controls.
- No OnBeforeCompanyOpen / OnAfterCompanyOpen; HTTP is not used.
- Control add-in uses external JavaScript files only.
