# Renders the Model Check report sheet and the Model Check pane to PNG for a visual review.
param([string]$Book, [string]$OutDir)
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Add-Type -AssemblyName System.Drawing
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force; Start-Sleep 1
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true; $xl.WindowState = -4137   # maximized
$procId = [LtxWin]::PidOf([IntPtr]$xl.Hwnd)
if (-not (($xl.Workbooks | % Name) -contains 'Lattice.xlam')) { [void]$xl.Workbooks.Open("$env:APPDATA\Microsoft\AddIns\Lattice.xlam") }
function Settle() { Start-Sleep -Milliseconds 600; for ($k = 0; $k -lt 30; $k++) { try { if ($xl.Ready) { break } } catch {}; Start-Sleep -Milliseconds 300 } }
function RunM($name, $arg) { $m = Invoke-Watched $procId { if ($null -ne $arg) { [void]$xl.Run("Lattice.xlam!$name", $arg) } else { [void]$xl.Run("Lattice.xlam!$name") } }.GetNewClosure(); Settle; "$name -> $(if ($m) { $m -replace '\s+',' ' } else { 'ok' })" }
function Shot($name) {
    Start-Sleep -Milliseconds 900
    $sx = [int]($xl.Left * 96 / 72); $sy = [int]($xl.Top * 96 / 72); $sw = [int]($xl.Width * 96 / 72); $sh = [int]($xl.Height * 96 / 72)
    $b = New-Object Drawing.Bitmap $sw, $sh; $g = [Drawing.Graphics]::FromImage($b); $g.CopyFromScreen($sx, $sy, 0, 0, $b.Size)
    $p = Join-Path $OutDir "$name.png"; $b.Save($p, [Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $b.Dispose(); "saved $p ($sw x $sh at $sx,$sy)"
}
$wb = $xl.Workbooks.Open($Book)
$ws = $wb.Worksheets.Item('Model')
# plant a few more issues so the report has every category
$ws.Range('F2').Formula = '=B2*1.21'            # hard-code
$ws.Range('F3').Formula = '=IFERROR(B3/0,0)'    # masked error
$ws.Range('F4').Formula = '=B4/0'               # error
$ws.Range('F5').Value2 = "'123"                 # number as text (constant, not flagged) -> make formula
$ws.Range('F5').Formula = '="12"&"3"'           # number stored as text (formula)
$ws.Range('G2').Formula = '=B2+1'; $ws.Range('H2').Formula = '=C2*2'; $ws.Range('I2').Formula = '=D2+1'   # H2 inconsistent
$ws.Range('B7').Value2 = 42                     # unused numeric input
$ws.Activate()
RunM 'LtxAuditHardcodes' $true
$xl.ActiveWindow.Zoom = 110
Shot 'audit_sheet'
$ws.Activate(); $ws.Range('A1').Select()
RunM 'LtxModelCheckPane'
Shot 'model_check_pane'
RunM 'LtxHealthReport'
Shot 'health_report'
$t2 = $wb.Worksheets.Add(); $t2.Name = 'Summary'; $t2.Range('F20').Formula = '=Model!B3-Model!$B$4'; $ws.Range('B3').Formula = '=B2*0.3'
$t2.Activate(); $t2.Range('F20').Select(); RunM 'LtxTraceIn'
Shot 'trace_window'
$wb.Close($false); $xl.DisplayAlerts = $false; $xl.Quit()
"done"
