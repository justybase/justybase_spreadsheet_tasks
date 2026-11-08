# Microsoft Excel COM conformance tests

These tests use desktop Microsoft Excel through COM as the reference oracle. They
cover both directions:

- **writer** — `XlsxWriter` / `XlsbWriter` writes every case in `cases.json`;
  Excel opens the file and checks sheet structure, visibility, AutoFilter, the
  used range, every cell value and explicit number formats.
- **reader** — Excel writes the same cases; `XlsxReader` / `XlsbReader` reads them
  back and the values are compared with the spec.
- **updater** — an Excel-authored template is updated in place and Excel verifies
  the result.

The suite is self-contained: `cases.json` and the PowerShell scripts live in this
directory.

## Requirements

- Windows
- Desktop Microsoft Excel (COM automation registered)
- PowerShell 7 (`pwsh`)
- Node.js 16+

The suite is opt-in so normal CI (which has no Excel) stays green:

```powershell
$env:EXCEL_COM = '1'
npm run test:excel-com
```

Without `EXCEL_COM=1` the suite is skipped. With `EXCEL_COM=1` but no Excel or
`pwsh`, it fails instead of silently skipping.

## Files

| File | Purpose |
|---|---|
| `cases.json` | Shared workbook spec (sheets, rows, cells, number formats, checks) |
| `ExcelComSpec.psm1` | Spec expansion and shared PowerShell helpers |
| `Invoke-ExcelVerify.ps1` | Opens workbooks in Excel and writes JSON check results |
| `New-ExcelFixture.ps1` | Creates the Excel-authored fixtures used by reader tests |
| `conformance.test.ts` | Node test runner that writes cases, calls Excel and asserts |
| `stage2.test.ts` | Etap 2 checks: styles, formulas, merges, links, validation, CF, comments, widths |
| `Verify-Stage2.ps1` | Opens the Etap 2 workbooks and returns their Excel properties as JSON |
