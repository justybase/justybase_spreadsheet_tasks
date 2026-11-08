Set-StrictMode -Version Latest

# Shared helpers for the Excel COM oracle. Keep in sync with the copies in each library's test folder.

$script:Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$script:ExcelEpoch = [datetime]'1899-12-30'

function Get-CellSpecValue {
    <#
      Returns the raw value of one spec cell: $null, string, number, bool,
      or a hashtable with 'value' (number/string/bool) or 'date' (ISO string).
    #>
    param([AllowNull()] $Cell)
    if ($null -eq $Cell) { return $null }
    if ($Cell -is [pscustomobject]) {
        if ($Cell.PSObject.Properties['date']) { return @{ kind = 'date'; iso = [string]$Cell.date } }
        if ($Cell.PSObject.Properties['value']) {
            $fmt = if ($Cell.PSObject.Properties['numberFormat']) { [string]$Cell.numberFormat } else { $null }
            return @{ kind = 'value'; value = $Cell.value; numberFormat = $fmt }
        }
        throw "Unsupported cell object: $($Cell | ConvertTo-Json -Compress)"
    }
    return $Cell
}

function Get-ColumnName {
    param([int] $Index1Based)
    $n = $Index1Based
    $name = ''
    while ($n -gt 0) {
        $rem = [int](($n - 1) % 26)
        $name = [string][char](65 + $rem) + $name
        $n = [int][math]::Floor(($n - 1) / 26)
    }
    return $name
}

function Expand-SheetRows {
    <#
      Expands one sheet of cases.json into a jagged array of raw cell values (see
      cell_encoding in cases.json). Generator output is deterministic and matches
      the other libraries' expansion.
    #>
    param([Parameter(Mandatory)] $Sheet)

    if ($Sheet.PSObject.Properties['generator']) {
        $g = $Sheet.generator
        $rowCount = [int]$g.rows
        if ($g.PSObject.Properties['columns']) {
            $cols = [int]$g.columns
            $header = @(1..$cols | ForEach-Object { 'C{0:D2}' -f $_ })
            $out = New-Object System.Collections.ArrayList
            [void]$out.Add([object[]]$header)
            for ($i = 1; $i -le $rowCount; $i++) {
                $row = New-Object 'object[]' $cols
                for ($j = 1; $j -le $cols; $j++) { $row[$j - 1] = [double]($i * 1000 + $j) }
                [void]$out.Add($row)
            }
            return , $out.ToArray()
        }
        $out = New-Object System.Collections.ArrayList
        [void]$out.Add([object[]]@($g.header))
        for ($i = 1; $i -le $rowCount; $i++) {
            [void]$out.Add([object[]]@([double]$i, ("row-$i"), [double]($i * 0.25)))
        }
        return , $out.ToArray()
    }
    return , @($Sheet.rows | ForEach-Object { , [object[]]@($_) })
}

function Get-ExcelSerial {
    param([Parameter(Mandatory)][string] $Iso)
    $dt = [datetime]::ParseExact($Iso, 'yyyy-MM-ddTHH:mm:ss', $script:Invariant)
    return ($dt - $script:ExcelEpoch).TotalDays
}

function Get-ExpectedExcelValue {
    <#
      Converts a raw spec cell into the value Excel's Range.Value2 returns.
      Dates become serial numbers (1900 date system).
    #>
    param([AllowNull()] $Cell)
    $v = Get-CellSpecValue $Cell
    if ($null -eq $v) { return $null }
    if ($v -is [hashtable]) {
        if ($v.kind -eq 'date') { return [double](Get-ExcelSerial $v.iso) }
        if ($v.kind -eq 'value') { return $v.value }
    }
    return $v
}

function Test-ExcelValueEqual {
    param([AllowNull()] $Expected, [AllowNull()] $Actual)
    if ($null -eq $Expected) { return ($null -eq $Actual) -or ($Actual -is [string] -and $Actual -eq '') }
    if ($Expected -is [bool]) { return ($Actual -is [bool]) -and ($Actual -eq $Expected) }
    if ($Expected -is [string]) { return ($Actual -is [string]) -and ([string]::Equals($Actual, $Expected, [StringComparison]::Ordinal)) }
    if ($Expected -is [double] -or $Expected -is [int] -or $Expected -is [long] -or $Expected -is [decimal]) {
        if ($null -eq $Actual -or $Actual -is [string] -or $Actual -is [bool]) { return $false }
        $e = [double]$Expected
        $a = [double]$Actual
        return [math]::Abs($a - $e) -le (1e-9 * [math]::Max(1.0, [math]::Abs($e)))
    }
    return $false
}

Export-ModuleMember -Function Get-CellSpecValue, Get-ColumnName, Expand-SheetRows, Get-ExcelSerial, Get-ExpectedExcelValue, Test-ExcelValueEqual
