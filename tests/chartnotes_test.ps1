# Chart annotations: live totals series, CAGR arrow and difference arrow on a stacked column chart.
param([string]$OutDir, [string]$Xlam = "")
if (-not $Xlam) { $Xlam = Join-Path $env:APPDATA 'Microsoft/AddIns/Lattice.xlam' }
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
$fails = 0
function Check($name, $ok, $detail) { if ($ok) { "PASS $name" } else { $script:fails++; "FAIL $name  $detail" } }
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
[void]$xl.Workbooks.Open($Xlam)
$wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Data'
$ws.Range('B1:D1').Value2 = @('2022', '2023', '2024')
$ws.Range('A2').Value2 = 'Revenue'; $ws.Range('B2').Value2 = 100; $ws.Range('C2').Value2 = 120; $ws.Range('D2').Value2 = 150
$ws.Range('A3').Value2 = 'Other';   $ws.Range('B3').Value2 = 50;  $ws.Range('C3').Value2 = 55;  $ws.Range('D3').Value2 = 60
$co = $ws.ChartObjects().Add(200, 20, 420, 260); $co.Name = 'Ch1'
$co.Chart.ChartType = 52                      # xlColumnStacked
$co.Chart.SetSourceData($ws.Range('A1:D3'), 1)  # rows
$co.Chart.HasLegend = $true
function Shapes($co) { $names = @(); foreach ($s in $co.Chart.Shapes) { $names += $s.Name }; $names }

$r = $xl.Run('Lattice.xlam!LtxChartNote', 'Data', 'Ch1', 'totals', 0, 0)
Check 'totals: no error' ([string]::IsNullOrEmpty($r)) $r
$tot = $null; foreach ($s in $co.Chart.SeriesCollection()) { if ($s.Name -eq 'Lattice total') { $tot = $s } }
Check 'totals: a "Lattice total" series exists' ($tot -ne $null) "series=$($co.Chart.SeriesCollection().Count)"
if ($tot) {
    $v = @($tot.Values | ForEach-Object { $_ })
    Check 'totals: values 150 / 175 / 210' ($v[0] -eq 150 -and $v[1] -eq 175 -and $v[2] -eq 210) ($v -join ',')
    Check 'totals: live (formula on the helper sheet)' ($tot.Formula -match 'LatticeChartData') $tot.Formula
    Check 'totals: labels above, no line' ($tot.HasDataLabels -and $tot.Format.Line.Visible -eq 0) "labels=$($tot.HasDataLabels) line=$($tot.Format.Line.Visible)"
    $ws.Range('D2').Value2 = 200; $xl.Calculate()
    $v2 = @($tot.Values | ForEach-Object { $_ })
    Check 'totals: update when the data changes (210 -> 260)' ($v2[2] -eq 260) ($v2 -join ',')
    $ws.Range('D2').Value2 = 150; $xl.Calculate()
}
$r = $xl.Run('Lattice.xlam!LtxChartNote', 'Data', 'Ch1', 'cagr', 0, 0)
Check 'cagr: no error' ([string]::IsNullOrEmpty($r)) $r
$lbl = $null; foreach ($s in $co.Chart.Shapes) { if ($s.Name -eq 'LtxCAGR label') { $lbl = $s } }
Check 'cagr: label shape exists' ($lbl -ne $null) ((Shapes $co) -join ',')
if ($lbl) { Check 'cagr: 150 -> 210 over 2 periods = +18.3%' ($lbl.TextFrame2.TextRange.Text -eq 'CAGR +18.3%') $lbl.TextFrame2.TextRange.Text }
$r = $xl.Run('Lattice.xlam!LtxChartNote', 'Data', 'Ch1', 'cagr', 0, 0)
Check 'cagr: re-run replaces instead of duplicating' (((Shapes $co) | Where-Object { $_ -like 'LtxCAGR*' }).Count -eq 3) ((Shapes $co) -join ',')
$r = $xl.Run('Lattice.xlam!LtxChartNote', 'Data', 'Ch1', 'delta', 1, 3)
Check 'delta: no error' ([string]::IsNullOrEmpty($r)) $r
$dl = $null; foreach ($s in $co.Chart.Shapes) { if ($s.Name -eq 'LtxDelta label') { $dl = $s } }
Check 'delta: label shows +40.0% (150 -> 210)' ($dl -ne $null -and $dl.TextFrame2.TextRange.Text -eq '+40.0%') $(if ($dl) { $dl.TextFrame2.TextRange.Text } else { 'missing' })
$wb.Close($false); $xl.Quit()
"chartnotes_test: $fails failure(s)"
