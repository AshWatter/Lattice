# Jumps must scroll the target into view, with and without frozen panes, far from the visible area.
param([string]$OutDir, [string]$Xlam = "")
if (-not $Xlam) { $Xlam = Join-Path $env:APPDATA 'Microsoft/AddIns/Lattice.xlam' }
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
$fails = 0
function Check($name, $ok, $detail) { if ($ok) { "PASS $name" } else { $script:fails++; "FAIL $name  $detail" } }
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
[void]$xl.Workbooks.Open($Xlam)
$wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Model'
$xl.WindowState = -4137
$ws.Range('A1').Value2 = 'Header'; $ws.Range('A2').Value2 = 'Sub'
$ws.Range('B280').Value2 = 42                       # far down
$ws.Range('AZ300').Value2 = 7                       # far down and right
$ws.Range('C5').Formula = '=B280*2'                 # near the top, precedent far away
$ws.Range('C6').Formula = '=AZ300+C5'               # two precedents, one far right
function Visible($addr) {   # on screen = inside the frozen rows/columns, or within ~35 rows / ~15 columns of the SCROLLABLE pane's scroll position
    $c = $ws.Range($addr); $w = $xl.ActiveWindow; $sr = 0; $sc = 0; if ($w.FreezePanes) { $sr = $w.SplitRow; $sc = $w.SplitColumn }
    $p = $w.Panes.Item($w.Panes.Count); $row = $p.ScrollRow; $col = $p.ScrollColumn
    $rowOk = ($c.Row -le $sr) -or ($c.Row -ge $row -and $c.Row -le $row + 35)
    $colOk = ($c.Column -le $sc) -or ($c.Column -ge $col -and $c.Column -le $col + 15)
    $rowOk -and $colOk }
function Here() { "$($xl.ActiveSheet.Name)!$($xl.ActiveCell.Address($false, $false))" }

# --- no frozen panes ---
$ws.Activate(); $xl.ActiveWindow.FreezePanes = $false; $xl.ActiveWindow.ScrollRow = 1; $xl.ActiveWindow.ScrollColumn = 1
$ws.Range('C5').Select() | Out-Null; $xl.Run('Lattice.xlam!LtxGoPrecedent'); Start-Sleep -Milliseconds 500
Check 'no freeze: Ctrl+[ lands on B280' ((Here) -eq 'Model!B280') (Here)
Check 'no freeze: B280 scrolled into view' (Visible 'B280') "paneRow=$($xl.ActiveWindow.Panes.Item($xl.ActiveWindow.Panes.Count).ScrollRow)"
$xl.Run('Lattice.xlam!LtxGoBack'); Start-Sleep -Milliseconds 400
Check 'no freeze: back to C5, visible' ((Here) -eq 'Model!C5' -and (Visible 'C5')) (Here)

# --- frozen panes at B3 (2 rows, 1 column) ---
$ws.Range('B3').Select() | Out-Null; $xl.ActiveWindow.FreezePanes = $true
$xl.ActiveWindow.ScrollRow = 3; $xl.ActiveWindow.ScrollColumn = 2
$ws.Range('C5').Select() | Out-Null; $xl.Run('Lattice.xlam!LtxGoPrecedent'); Start-Sleep -Milliseconds 500
Check 'frozen: Ctrl+[ lands on B280' ((Here) -eq 'Model!B280') (Here)
Check 'frozen: B280 scrolled into view' (Visible 'B280') "paneRow=$($xl.ActiveWindow.Panes.Item($xl.ActiveWindow.Panes.Count).ScrollRow) winRow=$($xl.ActiveWindow.ScrollRow) split=$($xl.ActiveWindow.SplitRow)/$($xl.ActiveWindow.SplitColumn)"
$xl.Run('Lattice.xlam!LtxGoBack'); Start-Sleep -Milliseconds 400
Check 'frozen: back to C5, visible' ((Here) -eq 'Model!C5' -and (Visible 'C5')) (Here)
# direct jump far right while frozen
$xl.Run('Lattice.xlam!LtxNavJump', "'[$($wb.Name)]Model'!`$AZ`$300"); Start-Sleep -Milliseconds 400
Check 'frozen: jump to AZ300 visible' ((Here) -eq 'Model!AZ300' -and (Visible 'AZ300')) "$(Here) scrollCol=$($xl.ActiveWindow.ScrollColumn)"
# a frozen cell (A1) must not scroll the panes away
$xl.Run('Lattice.xlam!LtxNavJump', "'[$($wb.Name)]Model'!`$A`$1"); Start-Sleep -Milliseconds 400
Check 'frozen: jump to A1 (inside the frozen area) visible' ((Here) -eq 'Model!A1' -and (Visible 'A1')) (Here)
# Trace In window with frozen panes: opening on C6 selects its first precedent and shows it
$ws.Range('C6').Select() | Out-Null; $xl.ActiveWindow.ScrollRow = 3; $xl.ActiveWindow.ScrollColumn = 2
$xl.Run('Lattice.xlam!LtxTraceIn'); Start-Sleep 2
$first = Here
Check 'frozen: Trace In jumps to the first precedent and it is visible' (($first -eq 'Model!AZ300' -or $first -eq 'Model!C5') -and (Visible ($first -replace '^Model!', ''))) "$first row=$($xl.ActiveWindow.ScrollRow) col=$($xl.ActiveWindow.ScrollColumn)"
$xl.Run('Lattice.xlam!LtxCloseSettings')
# unfreeze afterwards: still fine
$xl.ActiveWindow.FreezePanes = $false
$ws.Range('C5').Select() | Out-Null; $xl.Run('Lattice.xlam!LtxGoPrecedent'); Start-Sleep -Milliseconds 500
Check 'after unfreezing: jump still visible' ((Here) -eq 'Model!B280' -and (Visible 'B280')) (Here)
$wb.Close($false); $xl.Quit()
"freeze_test: $fails failure(s)"
