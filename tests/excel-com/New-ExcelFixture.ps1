<#
.SYNOPSIS
  Writes cases.json workbooks with Microsoft Excel (COM) so the libraries can be tested as readers
  against files produced by Excel itself.

.DESCRIPTION
  Creates <OutDir>\<caseId>.xlsx and <OutDir>\<caseId>.xlsb for each case (or only -CaseIds) and
  writes <OutDir>\targets.json in the format consumed by Invoke-ExcelVerify.ps1.

.PARAMETER SpecPath
  Path to cases.json.
.PARAMETER OutDir
  Output directory (created if missing).
.PARAMETER CaseIds
  Optional subset of case ids.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $SpecPath,
    [Parameter(Mandatory)][string] $OutDir,
    [string[]] $CaseIds
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelComSpec.psm1') -Force

$OutDir = [System.IO.Path]::GetFullPath($OutDir)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$spec = Get-Content -LiteralPath $SpecPath -Raw -Encoding UTF8 | ConvertFrom-Json -DateKind String
$formats = @(
    @{ name = 'xlsx'; code = 51 },   # xlOpenXMLWorkbook
    @{ name = 'xlsb'; code = 50 }    # xlExcel12
)
$missing = [System.Reflection.Missing]::Value

function Fill-Sheet {
    param($Worksheet, $SheetSpec)

    $rows = Expand-SheetRows $SheetSpec
    $nRows = $rows.Count
    $nCols = 0
    foreach ($r in $rows) { if ($r.Count -gt $nCols) { $nCols = $r.Count } }
    if ($nRows -eq 0) { return }

    $arr = New-Object 'object[,]' $nRows, $nCols
    $numberFormats = @{}
    for ($r = 0; $r -lt $nRows; $r++) {
        $row = $rows[$r]
        for ($c = 0; $c -lt $row.Count; $c++) {
            $cell = Get-CellSpecValue $row[$c]
            if ($null -eq $cell) { continue }
            if ($cell -is [hashtable] -and $cell.kind -eq 'date') {
                $arr[$r, $c] = [double](Get-ExcelSerial $cell.iso)
                $numberFormats["$($r + 1),$($c + 1)"] = 'yyyy-mm-dd hh:mm:ss'
                continue
            }
            $v = if ($cell -is [hashtable]) { $cell.value } else { $cell }
            if ($cell -is [hashtable] -and $cell.numberFormat) { $numberFormats["$($r + 1),$($c + 1)"] = [string]$cell.numberFormat }
            $arr[$r, $c] = if ($v -is [string] -or $v -is [bool]) { $v } else { [double]$v }
        }
    }
    $range = $Worksheet.Range($Worksheet.Cells(1, 1), $Worksheet.Cells($nRows, $nCols))
    $range.Value2 = $arr
    foreach ($key in $numberFormats.Keys) {
        $parts = $key.Split(',')
        $Worksheet.Cells([int]$parts[0], [int]$parts[1]).NumberFormat = $numberFormats[$key]
    }
}

$targets = New-Object System.Collections.ArrayList
$excel = New-Object -ComObject Excel.Application
try {
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    # NumberFormat takes en-US codes only when the locale's separators are not used (e.g. Polish Excel rejects "0.00").
    $excel.UseSystemSeparators = $false
    $excel.DecimalSeparator = '.'
    $excel.ThousandsSeparator = ','

    foreach ($case in $spec.cases) {
        if ($CaseIds -and ($CaseIds -notcontains $case.id)) { continue }
        foreach ($fmt in $formats) {
            $wb = $excel.Workbooks.Add()
            try {
                $sheets = @($case.sheets)
                while ($wb.Worksheets.Count -lt $sheets.Count) {
                    [void]$wb.Worksheets.Add($missing, $wb.Worksheets.Item($wb.Worksheets.Count))
                }
                for ($i = 0; $i -lt $sheets.Count; $i++) {
                    $sheetSpec = $sheets[$i]
                    $ws = $wb.Worksheets.Item($i + 1)
                    $ws.Name = [string]$sheetSpec.name
                    Fill-Sheet $ws $sheetSpec
                    if ($sheetSpec.PSObject.Properties['autoFilter'] -and [bool]$sheetSpec.autoFilter) {
                        $rows = Expand-SheetRows $sheetSpec
                        $lastCol = 0; foreach ($r in $rows) { if ($r.Count -gt $lastCol) { $lastCol = $r.Count } }
                        $ws.Range($ws.Cells(1, 1), $ws.Cells($rows.Count, $lastCol)).AutoFilter() | Out-Null
                    }
                    if ($sheetSpec.PSObject.Properties['hidden'] -and [bool]$sheetSpec.hidden) { $ws.Visible = 0 }
                }
                $wb.Worksheets.Item(1).Activate()
                $path = Join-Path $OutDir ("{0}.{1}" -f $case.id, $fmt.name)
                if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
                $wb.SaveAs($path, $fmt.code)
                [void]$targets.Add([pscustomobject]@{ file = $path; caseId = $case.id; format = $fmt.name })
                Write-Host "wrote $path"
            } finally {
                $wb.Close($false)
            }
        }
    }
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

$targetsPath = Join-Path $OutDir 'targets.json'
[System.IO.File]::WriteAllText($targetsPath, (ConvertTo-Json -InputObject @($targets) -Depth 3), [System.Text.UTF8Encoding]::new($false))
Write-Host "targets: $targetsPath ($($targets.Count) files)"
