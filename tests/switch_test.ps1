# Case switch: a bare Forms drop-down at the active cell (no helper sheet, no names); with the case
# names selected first, the input range is pre-filled. Copies keep their wiring.
param([string]$Book, [string]$OutDir, [string]$Xlam = "")
if (-not $Xlam) { $Xlam = Join-Path $env:APPDATA 'Microsoft/AddIns/Lattice.xlam' }
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
$fails = 0
function Check($name, $ok, $detail) { if ($ok) { "PASS $name" } else { $script:fails++; "FAIL $name  $detail" } }
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$procId = [LtxWin]::PidOf([IntPtr]$xl.Hwnd)
$wb = $xl.Workbooks.Open($Book)
[void]$xl.Workbooks.Open($Xlam)
$ws = $wb.Worksheets.Item('Model'); $ws.Activate()
$sheetsBefore = $wb.Worksheets.Count; $namesBefore = $wb.Names.Count

# --- mode 1: single cell -> bare control ---
$ws.Range('H2').Select() | Out-Null
$m = Invoke-Watched $procId { [void]$xl.Run('Lattice.xlam!LtxInsertCaseSwitch') }
$dd = $null; foreach ($s in $ws.Shapes) { if ($s.Name -like 'LtxCaseSwitch_*') { $dd = $s } }
Check 'drop-down created at the active cell' ($dd -ne $null -and [Math]::Abs($dd.Top - $ws.Range('H2').Top) -lt 2) "dialog=$m"
if ($dd) {
    Check 'it is a Forms drop-down (type 8, form control 2)' ($dd.Type -eq 8 -and $dd.FormControlType -eq 2) "type=$($dd.Type) fc=$($dd.FormControlType)"
    Check 'no input range and no cell link yet (bare)' ([string]::IsNullOrEmpty($dd.ControlFormat.ListFillRange) -and [string]::IsNullOrEmpty($dd.ControlFormat.LinkedCell)) "list=$($dd.ControlFormat.ListFillRange) link=$($dd.ControlFormat.LinkedCell)"
}
Check 'no sheet was added' ($wb.Worksheets.Count -eq $sheetsBefore) "sheets $sheetsBefore -> $($wb.Worksheets.Count)"
Check 'no workbook name was added' ($wb.Names.Count -eq $namesBefore) "names $namesBefore -> $($wb.Names.Count)"
Check 'no dialog shown' (-not $m) "dialog=$m"
# wire it the way the user would (Format Control) and use it
$ws.Range('J2').Value2 = 'Base'; $ws.Range('J3').Value2 = 'Upside'; $ws.Range('J4').Value2 = 'Downside'
if ($dd) {
    $dd.ControlFormat.ListFillRange = 'Model!$J$2:$J$4'; $dd.ControlFormat.LinkedCell = 'Model!$K$2'
    $dd.ControlFormat.Value = 3; Start-Sleep -Milliseconds 300
    Check 'after manual wiring the 3rd item writes 3 into the link cell' ($ws.Range('K2').Value2 -eq 3) "K2=$($ws.Range('K2').Value2)"
    $other = $wb.Worksheets.Add(); $other.Name = 'Outputs'
    $dd.Copy(); $other.Activate(); $other.Range('B2').Select() | Out-Null; $other.Paste()
    $cp = $other.Shapes.Item($other.Shapes.Count); $cp.ControlFormat.Value = 2; Start-Sleep -Milliseconds 300
    Check 'a copy on another sheet drives the same link cell' ($ws.Range('K2').Value2 -eq 2) "K2=$($ws.Range('K2').Value2) copyLink=$($cp.ControlFormat.LinkedCell)"
}
# --- mode 2: selection holds the names -> input range pre-filled, control to the right ---
$ws.Activate(); $ws.Range('J2:J4').Select() | Out-Null
$m2 = Invoke-Watched $procId { [void]$xl.Run('Lattice.xlam!LtxInsertCaseSwitch') }
$dd2 = $null; foreach ($s in $ws.Shapes) { if ($s.Name -like 'LtxCaseSwitch_*' -and $s.Name -ne $dd.Name) { $dd2 = $s } }
Check 'second control created from the selected names' ($dd2 -ne $null) "dialog=$m2"
if ($dd2) {
    Check 'input range pre-filled with the selection' ($dd2.ControlFormat.ListFillRange -match 'J\$2:\$J\$4') $dd2.ControlFormat.ListFillRange
    Check 'placed to the right of the list' ([Math]::Abs($dd2.Left - $ws.Range('L2').Left) -lt 2) "left=$($dd2.Left) L2=$($ws.Range('L2').Left)"
    Check 'three items listed' ($dd2.ControlFormat.ListCount -eq 3) "count=$($dd2.ControlFormat.ListCount)"
}
$wb.Close($false); $xl.Quit()
"switch_test: $fails failure(s)"
