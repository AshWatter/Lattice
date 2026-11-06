# Case runner: runs the cases through the case cell and tabulates outputs on a styled sheet.
param([string]$OutDir, [string]$Xlam = "")
if (-not $Xlam) { $Xlam = Join-Path $env:APPDATA 'Microsoft/AddIns/Lattice.xlam' }
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
$fails = 0
function Check($name, $ok, $detail) { if ($ok) { "PASS $name" } else { $script:fails++; "FAIL $name  $detail" } }
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
[void]$xl.Workbooks.Open($Xlam)
$wb = $xl.Workbooks.Add()
$drv = $wb.Worksheets.Item(1); $drv.Name = 'Drivers'
$drv.Range('A2').Value2 = 'Case'; $drv.Range('B2').Value2 = 1
$drv.Range('C2').Value2 = 'Base'; $drv.Range('C3').Value2 = 'Upside'; $drv.Range('C4').Value2 = 'Downside'
$mdl = $wb.Worksheets.Add(); $mdl.Name = 'Model'
$mdl.Range('A1').Value2 = 'Revenue';  $mdl.Range('B1').Formula = '=CHOOSE(Drivers!B2,100,120,80)'
$mdl.Range('A2').Value2 = 'EBITDA';   $mdl.Range('B2').Formula = '=B1*0.3'
$mdl.Range('A3').Value2 = 'IRR';      $mdl.Range('B3').Formula = '=CHOOSE(Drivers!B2,0.15,0.22,0.08)'; $mdl.Range('B3').NumberFormat = '0.0%'
$mdl.Range('A5').Value2 = 'Net debt'; $mdl.Range('B5').Formula = '=CHOOSE(Drivers!B2,0,50,-30)'
$mdl.Activate()
$res = $xl.Run('Lattice.xlam!LtxRunCasesFor', 'Drivers!$B$2', 'Model!$B$1:$B$3,Model!$B$5', 'Base|Upside|Downside')
Check 'core returned no error' ([string]::IsNullOrEmpty($res)) "res=$res"
$rpt = $null; try { $rpt = $wb.Worksheets.Item('Lattice Cases') } catch {}
Check 'report sheet created' ($rpt -ne $null) ''
if ($rpt) {
    Check 'header row lists the cases' ($rpt.Range('C4').Text -eq 'Base' -and $rpt.Range('D4').Text -eq 'Upside' -and $rpt.Range('E4').Text -eq 'Downside') "$($rpt.Range('C4').Text)|$($rpt.Range('D4').Text)|$($rpt.Range('E4').Text)"
    Check 'label read from the row' ($rpt.Range('A5').Text -eq 'Revenue' -and $rpt.Range('A7').Text -eq 'IRR') "$($rpt.Range('A5').Text) / $($rpt.Range('A7').Text)"
    Check 'revenue by case 100 / 120 / 80' ($rpt.Range('C5').Value2 -eq 100 -and $rpt.Range('D5').Value2 -eq 120 -and $rpt.Range('E5').Value2 -eq 80) "$($rpt.Range('C5').Value2)/$($rpt.Range('D5').Value2)/$($rpt.Range('E5').Value2)"
    Check 'EBITDA follows (30 / 36 / 24)' ([Math]::Round($rpt.Range('D6').Value2, 6) -eq 36) "$($rpt.Range('D6').Value2)"
    Check 'IRR keeps its % format' ($rpt.Range('C7').NumberFormat -eq '0.0%' -and $rpt.Range('C7').Text -eq '15.0%') "$($rpt.Range('C7').NumberFormat) $($rpt.Range('C7').Text)"
    Check 'upside vs base = +20.0%' ($rpt.Range('F5').Text -eq '+20.0%') $rpt.Range('F5').Text
    Check 'downside vs base = -20.0%, red' ($rpt.Range('G5').Text -eq '-20.0%' -and $rpt.Range('G5').Font.Color -eq 192) "$($rpt.Range('G5').Text) colour=$($rpt.Range('G5').Font.Color)"
    Check 'zero base falls back to absolute change (net debt +50)' ($rpt.Range('F8').Value2 -eq 50) "$($rpt.Range('F8').Text)"
    Check 'banner styled' ($rpt.Range('A1').Text -like 'LATTICE*CASE RUNNER*') $rpt.Range('A1').Text
}
Check 'case cell restored to 1' ($drv.Range('B2').Value2 -eq 1) "B2=$($drv.Range('B2').Value2)"
Check 'model back on base case (revenue 100)' ($mdl.Range('B1').Value2 -eq 100) "B1=$($mdl.Range('B1').Value2)"
$wb.Close($false); $xl.Quit()
"cases_test: $fails failure(s)"
