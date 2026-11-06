# Semantic checks of the Excel formula shortcuts: not "did it run" but "did it produce the right
# formula / value". Runs in its own Excel against the given add-in (default: the installed one).
param([string]$OutDir, [string]$Xlam = "")
if (-not $Xlam) { $Xlam = Join-Path $env:APPDATA 'Microsoft/AddIns/Lattice.xlam' }
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
$fails = 0
function Check($name, $ok, $detail) { if ($ok) { "PASS $name" } else { $script:fails++; "FAIL $name  $detail" } }
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$procId = [LtxWin]::PidOf([IntPtr]$xl.Hwnd)
[void]$xl.Workbooks.Open($Xlam)
$wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'T'
function RunM($name) { $m = Invoke-Watched $procId { [void]$xl.Run("Lattice.xlam!$name") }.GetNewClosure(); if ($m -and $m -match 'error') { "  dialog during $name : $m" } }
function Sel($a) { $ws.Activate(); $ws.Range($a).Select() | Out-Null }

# --- flip sign toggles instead of stacking minus signs ---
$ws.Range('A1').Value2 = 10; $ws.Range('A2').Value2 = 3
$ws.Range('B1').Formula = '=A1+A2'; $ws.Range('B2').Formula = '=A1'; $ws.Range('B3').Value2 = 5
Sel 'B1'; RunM 'LtxFlipSign'; Check 'flip: =A1+A2 -> =-(A1+A2)' ($ws.Range('B1').Formula -eq '=-(A1+A2)') $ws.Range('B1').Formula
Sel 'B1'; RunM 'LtxFlipSign'; Check 'flip again: back to =A1+A2' ($ws.Range('B1').Formula -eq '=A1+A2') $ws.Range('B1').Formula
Sel 'B2'; RunM 'LtxFlipSign'; Check 'flip single ref: =A1 -> =-A1' ($ws.Range('B2').Formula -eq '=-A1') $ws.Range('B2').Formula
Sel 'B2'; RunM 'LtxFlipSign'; Check 'flip single ref back: =-A1 -> =A1' ($ws.Range('B2').Formula -eq '=A1') $ws.Range('B2').Formula
Sel 'B3'; RunM 'LtxFlipSign'; Check 'flip number: 5 -> -5' ($ws.Range('B3').Value2 -eq -5) $ws.Range('B3').Value2
$ws.Range('B4').Formula = '=SUM(A1:A2)'; Sel 'B4'; RunM 'LtxFlipSign'; Check 'flip function call: =SUM(..) -> =-SUM(..)' ($ws.Range('B4').Formula -eq '=-SUM(A1:A2)') $ws.Range('B4').Formula
Sel 'B4'; RunM 'LtxFlipSign'; Check 'flip function back' ($ws.Range('B4').Formula -eq '=SUM(A1:A2)') $ws.Range('B4').Formula

# --- IFERROR wrap / unwrap ---
$ws.Range('C1').Formula = '=A1/A2'
Sel 'C1'; RunM 'LtxWrapIfError'; Check 'iferror wrap' ($ws.Range('C1').Formula -eq '=IFERROR(A1/A2,0)') $ws.Range('C1').Formula
Sel 'C1'; RunM 'LtxWrapIfError'; Check 'iferror unwrap on second press' ($ws.Range('C1').Formula -eq '=A1/A2') $ws.Range('C1').Formula

# --- anchor cycle A1 -> $A$1 -> A$1 -> $A1 -> A1 ---
$ws.Range('D1').Formula = '=A1*2'
$seq = @(); foreach ($k in 1..4) { Sel 'D1'; RunM 'LtxCycleAnchors'; $seq += $ws.Range('D1').Formula }
Check 'anchor cycle sequence' (($seq -join ' ') -eq '=$A$1*2 =A$1*2 =$A1*2 =A1*2') ($seq -join ' ')

# --- to values ---
$ws.Range('E1').Formula = '=A1+A2'; Sel 'E1'; RunM 'LtxToValues'; Check 'to values: 13 and no formula' ($ws.Range('E1').Value2 -eq 13 -and -not $ws.Range('E1').HasFormula) "$($ws.Range('E1').Formula)"

# --- fast fill right: follows the header row above ---
$ws.Range('A10:E10').Value2 = @(2020, 2021, 2022, 2023, 2024)
$ws.Range('A11').Formula = '=A10*2'; Sel 'A11'; RunM 'LtxFastFillRight'
Check 'fast fill right reaches the last header column' ($ws.Range('E11').Formula -eq '=E10*2' -and [string]::IsNullOrEmpty($ws.Range('F11').Formula)) "E11=$($ws.Range('E11').Formula) F11=$($ws.Range('F11').Formula)"

# --- CAGR: 100 -> 121 over 2 periods = 10%, on row 1 too ---
$ws.Range('A13:C13').Value2 = @(100, 110, 121)
Sel 'A13:C13'; RunM 'LtxInsertCAGR'
Check 'CAGR formula in the next column' ($ws.Range('D13').HasFormula) $ws.Range('D13').Formula
Check 'CAGR value = 10.0%' ([Math]::Round($ws.Range('D13').Value2, 4) -eq 0.1) $ws.Range('D13').Value2
$ws2 = $wb.Worksheets.Add(); $ws2.Name = 'Row1'; $ws2.Range('A1:D1').Value2 = @('Revenue', 50, 60, 72)
$ws2.Activate(); $ws2.Range('A1:D1').Select() | Out-Null; RunM 'LtxInsertCAGR'
Check 'CAGR on row 1 with a text label: 20.0%, no crash' ([Math]::Round($ws2.Range('E1').Value2, 4) -eq 0.2) "E1=$($ws2.Range('E1').Text)"

# --- growth row ---
$ws.Range('A15:D15').Value2 = @('Sales', 100, 120, 150)
Sel 'B15:D15'; RunM 'LtxInsertGrowthRow'
Check 'growth row inserted below with 20% then 25%' ([Math]::Round($ws.Range('C16').Value2, 4) -eq 0.2 -and [Math]::Round($ws.Range('D16').Value2, 4) -eq 0.25) "C16=$($ws.Range('C16').Text) D16=$($ws.Range('D16').Text)"

# --- number cycles change the format ---
$ws.Range('G1').Value2 = 1234.5; Sel 'G1'; $f0 = $ws.Range('G1').NumberFormat; RunM 'LtxCycleNumber'; $f1 = $ws.Range('G1').NumberFormat
Check 'number cycle changes the format' ($f1 -ne $f0) "$f0 -> $f1"
Sel 'G1'; RunM 'LtxCyclePercent'; Check 'percent cycle gives a % format' ($ws.Range('G1').NumberFormat -match '%') $ws.Range('G1').NumberFormat
Sel 'G1'; RunM 'LtxCycleMultiple'; Check 'multiple cycle gives an x format' ($ws.Range('G1').NumberFormat -match 'x') $ws.Range('G1').NumberFormat

# --- check row ---
$ws.Range('H1').Value2 = 5; $ws.Range('H2').Value2 = -5; $ws.Range('H3').Value2 = 0; Sel 'H1:H3'; RunM 'LtxInsertCheck'
Check 'check row formula below the block evaluates TRUE' ($ws.Range('H4').Value2 -eq $true) "H4=$($ws.Range('H4').Formula) -> $($ws.Range('H4').Text)"

# --- center across ---
$ws.Range('J1').Value2 = 'Title'; Sel 'J1:L1'; RunM 'LtxCenterAcross'
Check 'center across selection' ($ws.Range('J1').HorizontalAlignment -eq 7) $ws.Range('J1').HorizontalAlignment

$wb.Close($false); $xl.Quit()
"xl_semantics: $fails failure(s)"
