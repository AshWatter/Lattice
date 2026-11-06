# Smoke test of the new Excel modules through a fresh Excel instance with the installed add-in.
param([string]$OutDir)
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot "..\build\lib.ps1")
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $script:procId = [LtxWin]::PidOf([IntPtr]$xl.Hwnd)
function RunM($name) { $m = Invoke-Watched $script:procId { [void]$xl.Run("Lattice.xlam!$name") }.GetNewClosure(); if ($m -eq $null) { "OK   $name" } else { "RAN  $name : $($m -replace '\s+',' ')" } }
try {
    $ai = $xl.AddIns | Where-Object { $_.Name -eq 'Lattice.xlam' }
    "addin registered: " + ($ai -ne $null) + "  installed: " + $ai.Installed
    $loaded = ($xl.Workbooks | ForEach-Object { $_.Name }) -contains 'Lattice.xlam'
    "addin loaded at startup: $loaded"
    if (-not $loaded) { $xl.Workbooks.Open("$env:APPDATA\Microsoft\AddIns\Lattice.xlam") | Out-Null; "opened manually" }
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'DCF'
    $ws.Range('B2').Select()
    RunM 'LtxInsertDCF'
    "  DCF EV cell B28 = " + $ws.Range('B28').Value2 + "   value/share B30 = " + $ws.Range('B30').Value2 + "  (formulas: " + $ws.UsedRange.SpecialCells(-4123).Count + ")"
    $ws2 = $wb.Worksheets.Add(); $ws2.Name = 'BS'; $ws2.Range('A1').Select()
    RunM 'LtxInsertBlackScholes'
    "  BS call B14 = " + [math]::Round($ws2.Range('B14').Value2, 4) + "  put B15 = " + [math]::Round($ws2.Range('B15').Value2, 4) + "  (expected ~10.45 / 7.49)"
    $ws3 = $wb.Worksheets.Add(); $ws3.Name = 'Ret'; $ws3.Range('A1').Select()
    RunM 'LtxInsertReturns'
    "  IRR B8 = " + [math]::Round($ws3.Range('B8').Value2, 4) + "  MOIC B9 = " + [math]::Round($ws3.Range('B9').Value2, 3)
    $ws4 = $wb.Worksheets.Add(); $ws4.Name = 'Debt'; $ws4.Range('A1').Select()
    RunM 'LtxInsertDebtSchedule'
    "  closing yr5 F6 = " + [math]::Round($ws4.Range('F6').Value2, 2)
    # CAGR + growth row
    $ws5 = $wb.Worksheets.Add(); $ws5.Name = 'Growth'
    $ws5.Range('A1').Value2 = 'Year'; $ws5.Range('B1:F1').Value2 = @(2020,2021,2022,2023,2024)
    $ws5.Range('A2').Value2 = 'Revenue'; $ws5.Range('B2').Value2 = 100; $ws5.Range('C2').Value2 = 110; $ws5.Range('D2').Value2 = 125; $ws5.Range('E2').Value2 = 140; $ws5.Range('F2').Value2 = 160
    $ws5.Range('B2:F2').Select()
    RunM 'LtxInsertCAGR'
    "  CAGR G2 = " + [math]::Round($ws5.Range('G2').Value2, 4) + " (expected 0.1247)"
    $ws5.Range('B2:F2').Select()
    RunM 'LtxInsertGrowthRow'
    "  growth row C3 = " + [math]::Round($ws5.Range('C3').Value2, 4)
    # football field
    $ws6 = $wb.Worksheets.Add(); $ws6.Name = 'FF'
    $ws6.Range('A1:C1').Value2 = @('Method','Low','High')
    $ws6.Range('A2:C2').Value2 = @('DCF', 80, 120); $ws6.Range('A3:C3').Value2 = @('Trading comps', 70, 100); $ws6.Range('A4:C4').Value2 = @('Precedents', 90, 140); $ws6.Range('A5:C5').Value2 = @('52w range', 60, 95)
    $ws6.Range('A1:C5').Select()
    # the macro asks Yes/No about sorting: answer via SendKeys is unreliable, so set alerts off and expect a dialog -> use a timer-free approach: call through Application.Run in a job? Instead pre-answer by temporarily wrapping: skip the prompt by calling StackedColumn + Waterfall (no prompts) and report FF separately.
    $ws7 = $wb.Worksheets.Add(); $ws7.Name = 'Stack'
    $ws7.Range('A1:D1').Value2 = @('', 'Seg A', 'Seg B', 'Seg C')
    $ws7.Range('A2:D2').Value2 = @('2023', 10, 20, 5); $ws7.Range('A3:D3').Value2 = @('2024', 12, 22, 6); $ws7.Range('A4:D4').Value2 = @('2025', 15, 25, 8)
    $ws7.Range('A1:D4').Select()
    RunM 'LtxStackedColumn'
    "  charts on Stack: " + $ws7.ChartObjects().Count
    $ws8 = $wb.Worksheets.Add(); $ws8.Name = 'WF'
    $ws8.Range('A1:B1').Value2 = @('Item','Value'); $ws8.Range('A2:B2').Value2 = @('2023 EBITDA', 100); $ws8.Range('A3:B3').Value2 = @('Volume', 15); $ws8.Range('A4:B4').Value2 = @('Price', 8); $ws8.Range('A5:B5').Value2 = @('Costs', -12); $ws8.Range('A6:B6').Value2 = @('2024 EBITDA', 111)
    $ws8.Range('A1:B6').Select()
    RunM 'LtxWaterfall'
    "  charts on WF: " + $ws8.ChartObjects().Count
    # audit on a sheet with deliberate issues
    $ws9 = $wb.Worksheets.Add(); $ws9.Name = 'Audit'
    $ws9.Range('A1').Value2 = 100; $ws9.Range('B1').Formula = '=A1*1.21'; $ws9.Range('C1').Formula = '=IFERROR(1/0,0)'; $ws9.Range('D1').Formula = '=A1/0'; $ws9.Range('E1').Formula = '=A1*12'
    $ws9.Range('A2').Select()
    "  literals of B1: " + $xl.Run('Lattice.xlam!FormulaLiterals', '=A1*1.21') + "   of E1 (12 whitelisted): '" + $xl.Run('Lattice.xlam!FormulaLiterals', '=A1*12') + "'"
    $ws9.Activate(); $ws9.Range('A2').Select()
    RunM 'LtxAuditHardcodes'
    $rpt = $wb.Worksheets.Item('Lattice Audit')
    "  audit rows: " + ($rpt.UsedRange.Rows.Count - 1)
    for ($i = 2; $i -le $rpt.UsedRange.Rows.Count; $i++) { "    " + $rpt.Cells.Item($i,1).Text + " " + $rpt.Cells.Item($i,2).Text + " | " + $rpt.Cells.Item($i,3).Text + " | " + $rpt.Cells.Item($i,4).Text }
    RunM 'LtxHealthReport'
    "  health sheet exists: " + (($wb.Worksheets | % Name) -contains 'Lattice Health')
    # shortcuts
    RunM 'LtxKeysLoad'
    "  shortcuts file: " + (Test-Path "$env:APPDATA\Lattice\excel_shortcuts.txt")
    $wb.SaveAs((Join-Path $OutDir 'xl_smoke.xlsx'), 51)
    $wb.Close($false)
} catch { "SCRIPT ERROR: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)" }
$xl.Quit()

