param(
    [Parameter(Mandatory = $true)][string]$Path
)

$ErrorActionPreference = 'Stop'
$result = @{ ok = $false }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($Path, 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $ws.Calculate()

    $a2 = $ws.Range('A2')
    $result.valueA2 = $a2.Value2
    $result.formatA2 = $a2.NumberFormat
    $result.boldA2 = $a2.Font.Bold
    $result.fontColorA2 = $a2.Font.Color
    $result.fillA2 = $a2.Interior.Color
    $result.borderTopA2 = $a2.Borders.Item(8).LineStyle
    $result.alignA2 = $a2.HorizontalAlignment

    $result.formulaB2 = $ws.Range('B2').Formula
    $result.valueB2 = $ws.Range('B2').Value2
    $result.valueB3 = $ws.Range('B3').Value2
    $result.textC2 = $ws.Range('C2').Value2
    $result.merged = $ws.Range('A4:B4').MergeCells
    $result.links = $ws.Hyperlinks.Count
    $result.validationType = $ws.Range('C2').Validation.Type
    $result.validationFormula = $ws.Range('C2').Validation.Formula1
    $result.conditions = $ws.Range('A2:A3').FormatConditions.Count
    $result.widthA = $ws.Columns.Item(1).ColumnWidth
    $result.widthB = $ws.Columns.Item(2).ColumnWidth

    $comment = $ws.Range('A2').Comment
    if ($null -ne $comment) { $result.comment = $comment.Text() }
    else { $result.comment = $null }

    $wb.Close($false)
    $result.ok = $true
}
catch {
    $result.error = $_.Exception.Message
}
finally {
    $excel.Quit()
}

$result | ConvertTo-Json -Compress
