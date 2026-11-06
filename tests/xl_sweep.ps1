# Sweep of the Excel macros not covered by xl_smoke.ps1. Prompts auto-accepted with defaults.
param([string]$OutDir, [string]$Xlam = "")
if (-not $Xlam) { $Xlam = Join-Path $env:APPDATA "Microsoft/AddIns/Lattice.xlam" }
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
function RunM($name, $arg) {
    $m = Invoke-Watched $script:procId { if ($null -ne $arg) { [void]$xl.Run("Lattice.xlam!$name", $arg) } else { [void]$xl.Run("Lattice.xlam!$name") } }.GetNewClosure()
    Start-Sleep -Milliseconds 700; for ($k = 0; $k -lt 20; $k++) { try { $null = $xl.Ready; if ($xl.Ready) { break } } catch {}; Start-Sleep -Milliseconds 300 }
    if ($m -eq $null) { "OK   $name $arg" } elseif ($m -match 'Compile error|Run-time error|host error') { "FAIL $name $arg : $($m -replace '\s+',' ')" } else { "OK   $name $arg  [$(($m -replace '\s+',' ').Substring(0, [Math]::Min(110, ($m -replace '\s+',' ').Length)))]" }
}
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
$script:procId = [LtxWin]::PidOf([IntPtr]$xl.Hwnd)
if (-not (($xl.Workbooks | % Name) -contains 'Lattice.xlam')) { [void]$xl.Workbooks.Open("$env:APPDATA\Microsoft\AddIns\Lattice.xlam") }
$wb = $xl.Workbooks.Add()
$wb.SaveAs((Join-Path $OutDir 'xl_sweep.xlsx'), 51)
$ws = $wb.Worksheets.Item(1); $ws.Name = 'Model'
$ws.Range('A1:E1').Value2 = @('Item','2023','2024','2025','2026')
$ws.Range('A2').Value2 = 'Revenue'; $ws.Range('B2:E2').Value2 = @(100, 110, 121, 133)
$ws.Range('A3').Value2 = 'Costs';   $ws.Range('B3').Formula = '=-B2*0.6'; $ws.Range('C3').Formula = '=-C2*0.6'
$ws.Range('A4').Value2 = 'EBITDA';  $ws.Range('B4').Formula = '=B2+B3'; $ws.Range('C4').Formula = '=C2+C3'
$ws.Range('A6').Value2 = 'Rate';    $ws.Range('B6').Value2 = 0.25
"--- modelling ---"
$ws.Range('B3').Select(); RunM 'LtxCycleAnchors'; "  B3 formula: $($ws.Range('B3').Formula) (expect `$B`$2)"
$ws.Range('B3').Select(); RunM 'LtxCycleAnchors'; RunM 'LtxCycleAnchors'; RunM 'LtxCycleAnchors'; "  back to: $($ws.Range('B3').Formula)"
$ws.Range('B2').Select(); RunM 'LtxFlipSign'; "  B2 = $($ws.Range('B2').Value2) (expect -100)"; RunM 'LtxFlipSign'
$ws.Range('C3:C4').Select(); RunM 'LtxFastFillRight'; "  E4 formula: $($ws.Range('E4').Formula) (expect =E2+E3)"
$ws.Range('B4').Select(); RunM 'LtxCopyExact'; $ws.Range('B8').Select(); RunM 'LtxPasteExact'; "  B8 = $($ws.Range('B8').Formula) (expect =B2+B3)"
$ws.Range('B8').Select(); RunM 'LtxToValues'; "  B8 now constant: $(-not $ws.Range('B8').HasFormula)"
$ws.Range('C4').Select(); RunM 'LtxWrapIfError'; "  C4: $($ws.Range('C4').Formula)"
$ws.Range('B2:E2').Select(); RunM 'LtxInsertCheck'
$ws.Range('B2:E2').Select(); RunM 'LtxFastFillDown'
$ws.Range('B10').Select(); RunM 'LtxSimplify'
$ws.Range('B2:E4').Select(); RunM 'LtxNameRange'; "  names: $($wb.Names.Count)"
RunM 'LtxUnhideAll'
$ws.Range('B2').Select(); RunM 'LtxTagCell'
"--- auditing ---"
$ws.Range('B4').Select(); RunM 'LtxPrecedents'; "  active after precedents: $($xl.ActiveCell.Address($false,$false))"
$ws.Range('B2').Select(); RunM 'LtxDependents'
$ws.Range('B4').Select(); RunM 'LtxFormulaFlow'
$ws.Range('B4').Select(); RunM 'LtxShowArrows'; RunM 'LtxClearArrows'
$ws.Range('B3:E4').Select(); RunM 'LtxInconsistentFormulas'
RunM 'LtxSelectErrors'; RunM 'LtxSelectHardcodes'; RunM 'LtxToggleFormulas'; RunM 'LtxToggleFormulas'
RunM 'LtxCompareSheets'
RunM 'LtxHighlightHardcodes'; RunM 'LtxHighlightHardcodes'
$ws.Range('B3').Select(); RunM 'LtxExtractHardcode'
"--- formatting ---"
$ws.Range('B2:E2').Select()
foreach ($m in 'LtxCyclePercent','LtxCycleMultiple','LtxCycleDate','LtxCycleCurrency','LtxCycleFontColor','LtxCycleFill','LtxCycleBorder','LtxCycleAlign','LtxIndentMore','LtxIndentLess','LtxToggleWrap','LtxCenterAcross','LtxHeaderStyle','LtxTotalStyle','LtxAutofitCols','LtxGridlinesToggle','LtxZoomCycle','LtxColWidth') { RunM $m }
$ws.Range('B2').Select(); RunM 'LtxCopyFormat'; $ws.Range('B3').Select(); RunM 'LtxPasteFormat'
RunM 'LtxAutoColorSheet'
RunM 'LtxEditStyleFile'; Get-Process notepad -ErrorAction SilentlyContinue | Stop-Process -Force; "  style file: $(Test-Path "$env:APPDATA\Lattice\excel_styles.txt")"
"--- library / power tools / export ---"
$ws2 = $wb.Worksheets.Add(); $ws2.Name = 'Sens'; $ws2.Range('A1').Value2 = 0.1; $ws2.Range('A2').Value2 = 0.02; $ws2.Range('A3').Formula = '=1/(A1-A2)'; $ws2.Range('C5').Select()
'SKIP LtxInsertSensitivity (needs a user to pick cells in Excel''s RefEdit prompt)'
$ws.Activate(); $ws.Range('A1:E4').Select(); RunM 'LtxLibSave'
$ws.Range('H1').Select(); RunM 'LtxLibInsert'
RunM 'LtxLibOpen'; foreach ($w in @($xl.Workbooks)) { if ($w.Name -like 'xl_library*') { $w.Close($false) } }
RunM 'LtxLibSetPath'
RunM 'LtxOptimizeWorkbook'
RunM 'LtxHardcodeWorkbook' $true
RunM 'LtxHealthReport'
RunM 'LtxHealthReportRemove'
$ws.Activate(); $ws.Range('A1:E4').Select(); RunM 'LtxCopyForPpt' 'picture'; "  pending.txt: $(Test-Path "$env:APPDATA\Lattice\pending.txt")"
Remove-Item "$env:APPDATA\Lattice\pending.txt" -ErrorAction SilentlyContinue
RunM 'LtxExportPictureFile'; "  png: $(Test-Path (Join-Path $OutDir 'A1-E4.png'))"
RunM 'LtxSendToPpt' 'picture'    # PowerPoint not running -> message
RunM 'LtxRefreshPptFromHere'
"--- charts ---"
$ws3 = $wb.Worksheets.Add(); $ws3.Name = 'FF'
$ws3.Range('A1:C1').Value2 = @('Method','Low','High'); $ws3.Range('A2:C2').Value2 = @('DCF', 80, 120); $ws3.Range('A3:C3').Value2 = @('Comps', 70, 100); $ws3.Range('A4:C4').Value2 = @('Precedents', 90, 140)
$ws3.Range('A1:C4').Select(); RunM 'LtxFootballField'; "  charts on FF: $($ws3.ChartObjects().Count)"
RunM 'LtxFootballFieldLine'
'SKIP LtxInsertMarginRow (RefEdit prompt)'
$ws.Activate(); $ws.Range('B2:E2').Select(); RunM 'LtxQuickChart' 51
"--- keys ---"
RunM 'LtxKeysList'; RunM 'LtxKeysEdit'; Get-Process notepad -ErrorAction SilentlyContinue | Stop-Process -Force; RunM 'LtxKeysLoad'; RunM 'LtxKeysUnload'
RunM 'LtxSaveBackup'; "  backups: $((Get-ChildItem (Join-Path $OutDir 'Backups') -ErrorAction SilentlyContinue | Measure-Object).Count)"
"--- new: PE / HF generators ---"
$ws4 = $wb.Worksheets.Add(); $ws4.Name = 'LBO'; $ws4.Range('B2').Select(); RunM 'LtxInsertLBO'
"  LBO MOIC/IRR: " + (($ws4.UsedRange.Cells | ? { $_.Value2 -eq 'MOIC' } | % { [math]::Round($_.Offset(0,1).Value2,2) }) -join ',') + " / " + (($ws4.UsedRange.Cells | ? { $_.Value2 -eq 'IRR' } | % { [math]::Round($_.Offset(0,1).Value2,3) }) -join ',') + "  errors: " + (($ws4.UsedRange.Cells | ? { $_.Text -like '#*' }) | Measure-Object).Count
$ws5 = $wb.Worksheets.Add(); $ws5.Name = 'Dist'; $ws5.Range('A1').Select(); RunM 'LtxInsertDistributionWaterfall'
"  Dist check cell B20 = " + $ws5.Range('B20').Value2 + "  LP MOIC B17 = " + [math]::Round($ws5.Range('B17').Value2,3) + "  GP% B19 = " + [math]::Round($ws5.Range('B19').Value2,3)
$ws6 = $wb.Worksheets.Add(); $ws6.Name = 'AccDil'; $ws6.Range('A1').Select(); RunM 'LtxInsertAccretionDilution'
"  Acc/dil B24 = " + [math]::Round($ws6.Range('B24').Value2,4)
$ws7 = $wb.Worksheets.Add(); $ws7.Name = 'Credit'; $ws7.Range('A1').Select(); RunM 'LtxInsertCreditStats'
"  net leverage B11 = " + [math]::Round($ws7.Range('B11').Value2,2)
$ws8 = $wb.Worksheets.Add(); $ws8.Name = 'Scen'; $ws8.Range('A1').Select(); RunM 'LtxInsertScenarioBlock'
"  scenario live B4 = " + $ws8.Range('B4').Value2 + "  name Scenario exists: " + (($wb.Names | % Name) -contains 'Scenario')
$ws9 = $wb.Worksheets.Add(); $ws9.Name = 'Time'; $ws9.Range('A1').Select(); RunM 'LtxInsertTimeline'
"  timeline labels: " + (($ws9.Range('B2:E2').Value2 | % { $_ }) -join ' | ')
"--- new: comps / data ---"
$wsC = $wb.Worksheets.Add(); $wsC.Name = 'Comps'
$wsC.Range('A1:E1').Value2 = @('Company','EV/EBITDA','P/E','EBITDA margin','Revenue')
$wsC.Range('A2:E2').Value2 = @('Alpha', 8.5, 15.2, 0.22, 1200); $wsC.Range('A3:E3').Value2 = @('Beta', 10.1, 18.9, 0.25, 800); $wsC.Range('A4:E4').Value2 = @('Gamma', 7.2, 12.1, 0.18, 2100); $wsC.Range('A5:E5').Value2 = @('Delta', 25, 40, 0.31, 300)
$wsC.Range('A1:E5').Select(); RunM 'LtxFormatComps'; "  B2 format: $($wsC.Range('B2').NumberFormat)  D2: $($wsC.Range('D2').NumberFormat)  E2: $($wsC.Range('E2').NumberFormat)"
$wsC.Range('A1:E5').Select(); RunM 'LtxQuickStats'; "  stats: A7='$($wsC.Range('A7').Value2)' B7=$([math]::Round($wsC.Range('B7').Value2,2)) B8=$([math]::Round($wsC.Range('B8').Value2,2))"
$wsC.Range('B2:B5').Select(); RunM 'LtxFlagOutliers'; "  outlier fill B5: $($wsC.Range('B5').Interior.Color)"
$wsC.Range('B2:B5').Select(); RunM 'LtxRankPercentile'; "  rank C2=$($wsC.Range('C2').Value2)  (note: overwrote P/E col on purpose)"
$wsD = $wb.Worksheets.Add(); $wsD.Name = 'Clean'
$wsD.Range('A1').Value2 = "'(1,234.5)"; $wsD.Range('A2').Value2 = "' 12.5% "; $wsD.Range('A3').Value2 = "'8.0x"; $wsD.Range('A4').Value2 = "'  text  "
$wsD.Range('A1:A4').Select(); RunM 'LtxCleanNumbers'; "  A1=$($wsD.Range('A1').Value2) A2=$($wsD.Range('A2').Value2) A3=$($wsD.Range('A3').Value2)"
$wsD.Range('A1:A4').Select(); RunM 'LtxTrimText'; "  A4='$($wsD.Range('A4').Value2)'"
$wsD.Range('C1:D2').Merge(); $wsD.Range('C1').Value2 = 'merged'; $wsD.Range('C1:D2').Select(); RunM 'LtxUnmergeFill'; "  D2='$($wsD.Range('D2').Value2)' merged=$($wsD.Range('C1').MergeCells)"
$wsD.Range('A7').Value2 = 1; $wsD.Range('A9').Value2 = 2; $wsD.Range('A7:A9').Select(); RunM 'LtxRemoveBlankRows'; "  A8 now = $($wsD.Range('A8').Value2) (expect 2)"
$wsD.Range('F1').Value2 = 5000; $wsD.Range('F2').Formula = '=F1*2'; $wsD.Range('F1:F2').Select(); RunM 'LtxToThousands'; "  F1=$($wsD.Range('F1').Value2) F2=$($wsD.Range('F2').Formula)"
$wsD.Range('F1:F2').Select(); RunM 'LtxSelectionStats'
'SKIP LtxTransposeValues (RefEdit prompt)'
$wb.Save(); $wb.Close($false); $xl.Quit()
"done"
