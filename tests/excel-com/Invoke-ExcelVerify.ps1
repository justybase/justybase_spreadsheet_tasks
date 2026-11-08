<#
.SYNOPSIS
  Opens workbooks in Microsoft Excel (COM) and checks them against cases.json.

.DESCRIPTION
  For every target { file, caseId, format } the script opens the workbook read-only and checks:
    open           - Excel opens the file without error
    sheets         - sheet count and names, in order
    hidden         - sheet visibility matches the spec
    autoFilter     - AutoFilterMode matches the spec
    cells:<sheet>  - every value from the expanded grid equals Range.Value2
    usedRange:<sheet> - Excel's UsedRange does not extend past the expected grid
    numberFormat   - explicit number format checks from the spec
  Results are written as JSON (one object per check). Exit code 1 when any check fails.

.PARAMETER SpecPath
  Path to cases.json.
.PARAMETER TargetsPath
  JSON array: [ { "file": "...", "caseId": "...", "format": "xlsx|xlsb" } ]
.PARAMETER OutPath
  Where to write the JSON result array.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $SpecPath,
    [Parameter(Mandatory)][string] $TargetsPath,
    [Parameter(Mandatory)][string] $OutPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelComSpec.psm1') -Force

$spec = Get-Content -LiteralPath $SpecPath -Raw -Encoding UTF8 | ConvertFrom-Json -DateKind String
$cases = @{}
foreach ($c in $spec.cases) { $cases[$c.id] = $c }
$targets = @(Get-Content -LiteralPath $TargetsPath -Raw -Encoding UTF8 | ConvertFrom-Json -DateKind String)

$results = New-Object System.Collections.ArrayList

function Add-Result {
    param($Target, [string] $Check, [bool] $Pass, $Expected, $Actual, [string] $Message = '', [string] $Sheet = '')
    [void]$results.Add([pscustomobject]@{
        caseId   = $Target.caseId
        format   = $Target.format
        file     = $Target.file
        sheet    = $Sheet
        check    = $Check
        pass     = $Pass
        expected = if ($null -eq $Expected) { $null } else { [string]$Expected }
        actual   = if ($null -eq $Actual) { $null } else { [string]$Actual }
        message  = $Message
    })
}

function Read-Grid {
    param($Worksheet, [int] $Rows, [int] $Cols)
    $rng = $Worksheet.Range($Worksheet.Cells(1, 1), $Worksheet.Cells($Rows, $Cols))
    $raw = $rng.Value2
    if ($Rows * $Cols -eq 1) {
        $grid = New-Object 'object[,]' 1, 1
        $grid[1, 1] = $raw
        return , $grid
    }
    return , $raw
}

function Compare-Sheet {
    param($Target, $Worksheet, $ExpectedSheet, [int] $Index)

    $name = [string]$ExpectedSheet.name
    $expectedVisible = -not ($ExpectedSheet.PSObject.Properties['hidden'] -and [bool]$ExpectedSheet.hidden)
    $actualVisible = ($Worksheet.Visible -eq -1)
    Add-Result $Target 'hidden' ($expectedVisible -eq $actualVisible) (-not $expectedVisible) (-not $actualVisible) "sheet $name" $name

    $expectedAutoFilter = [bool]$ExpectedSheet.autoFilter
    $actualAutoFilter = [bool]$Worksheet.AutoFilterMode
    Add-Result $Target 'autoFilter' ($expectedAutoFilter -eq $actualAutoFilter) $expectedAutoFilter $actualAutoFilter '' $name

    $rows = Expand-SheetRows $ExpectedSheet
    $nRows = $rows.Count
    $nCols = 0
    foreach ($r in $rows) { if ($r.Count -gt $nCols) { $nCols = $r.Count } }

    if ($nRows -eq 0 -or $nCols -eq 0) { return }

    $grid = Read-Grid $Worksheet $nRows $nCols
    $mismatches = New-Object System.Collections.ArrayList
    $mismatchTotal = 0
    $cellsChecked = $nRows * $nCols
    for ($r = 0; $r -lt $nRows; $r++) {
        $row = $rows[$r]
        for ($c = 0; $c -lt $nCols; $c++) {
            $cell = if ($c -lt $row.Count) { $row[$c] } else { $null }
            $expected = Get-ExpectedExcelValue $cell
            $actual = $grid[($r + 1), ($c + 1)]
            if (-not (Test-ExcelValueEqual $expected $actual)) {
                $mismatchTotal++
                if ($mismatches.Count -lt 5) {
                    [void]$mismatches.Add(("{0}{1}: expected [{2}] got [{3}]" -f (Get-ColumnName ($c + 1)), ($r + 1), $expected, $actual))
                }
            }
        }
    }
    Add-Result $Target "cells:$name" ($mismatchTotal -eq 0) "$cellsChecked cells" "$($cellsChecked - $mismatchTotal) matching" ($mismatches -join '; ') $name

    $used = $Worksheet.UsedRange
    $lastRow = [int]$used.Row + [int]$used.Rows.Count - 1
    $lastCol = [int]$used.Column + [int]$used.Columns.Count - 1
    $usedOk = ($lastRow -le $nRows) -and ($lastCol -le $nCols)
    Add-Result $Target "usedRange:$name" $usedOk "A1:$(Get-ColumnName $nCols)$nRows" $used.Address($false, $false) '' $name
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable

    foreach ($target in $targets) {
        $case = $cases[$target.caseId]
        if ($null -eq $case) {
            Add-Result $target 'open' $false 'known case' $target.caseId "unknown caseId"
            continue
        }

        $wb = $null
        try {
            $wb = $excel.Workbooks.Open([string]$target.file, 0, $true)
        } catch {
            Add-Result $target 'open' $false 'opens in Excel' 'error' $_.Exception.Message
            continue
        }
        Add-Result $target 'open' $true 'opens in Excel' 'opened'

        try {
            $expectedSheets = @($case.sheets)
            $actualCount = [int]$wb.Worksheets.Count
            $actualNames = @(); for ($i = 1; $i -le $actualCount; $i++) { $actualNames += [string]$wb.Worksheets.Item($i).Name }
            $expectedNames = @($expectedSheets | ForEach-Object { [string]$_.name })
            $namesOk = ($actualCount -eq $expectedNames.Count) -and (@(Compare-Object $expectedNames $actualNames -CaseSensitive).Count -eq 0)
            Add-Result $target 'sheets' $namesOk ($expectedNames -join ', ') ($actualNames -join ', ')

            for ($i = 0; $i -lt $expectedSheets.Count; $i++) {
                if ($i -ge $actualCount) { break }
                Compare-Sheet $target $wb.Worksheets.Item($i + 1) $expectedSheets[$i] ($i + 1)
            }

            foreach ($sheet in $expectedSheets) {
                if (-not $sheet.PSObject.Properties['checks']) { continue }
                $ws = $wb.Worksheets.Item([string]$sheet.name)
                foreach ($chk in @($sheet.checks)) {
                    if ($chk.type -eq 'numberFormat') {
                        $actualFmt = [string]$ws.Range([string]$chk.address).NumberFormat
                        Add-Result $target "numberFormat:$($chk.address)" ($actualFmt -ceq [string]$chk.expected) $chk.expected $actualFmt '' ([string]$sheet.name)
                    }
                }
            }
        } catch {
            Add-Result $target 'inspect' $false 'inspection completes' 'exception' $_.Exception.Message
        } finally {
            if ($wb) { try { $wb.Close($false) } catch { } }
        }
    }
}
finally {
    if ($excel) {
        try { $excel.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

$json = ConvertTo-Json -InputObject @($results) -Depth 4
[System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($OutPath), $json, [System.Text.UTF8Encoding]::new($false))
$failed = @($results | Where-Object { -not $_.pass }).Count
Write-Host "Excel verify: $($results.Count - $failed) passed, $failed failed"
if ($failed -gt 0) { exit 1 }
