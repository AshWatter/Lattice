# Isolates the Trace In navigation: direct jump, jump while the window is open, then keyboard steps
# with screenshots of the screen after each key (saved in $OutDir).
param([string]$Book, [string]$OutDir)
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Add-Type -AssemblyName System.Drawing; Add-Type -AssemblyName System.Windows.Forms
function Shot($name) {
    $b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bmp = New-Object Drawing.Bitmap $b.Width, $b.Height; $g = [Drawing.Graphics]::FromImage($bmp); $g.CopyFromScreen($b.Left, $b.Top, 0, 0, $bmp.Size)
    $small = New-Object Drawing.Bitmap $bmp, ([int]($b.Width / 2)), ([int]($b.Height / 2))
    $p = Join-Path $OutDir "$name.png"; $small.Save($p, [Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose(); $small.Dispose(); "saved $p"
}
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force; Start-Sleep 1
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$wb = $xl.Workbooks.Open($Book)
if (-not (($xl.Workbooks | % Name) -contains 'Lattice.xlam')) { [void]$xl.Workbooks.Open("$env:APPDATA\Microsoft\AddIns\Lattice.xlam") }
$calc = $wb.Worksheets.Add(); $calc.Name = 'Calc Sheet'
$calc.Range('A1').Formula = '=Model!B2*2'; $calc.Range('A2').Formula = "='Calc Sheet'!A1+1"
$calc.Activate(); $calc.Range('A2').Select() | Out-Null; $xl.WindowState = -4137
function Pos { "$($xl.ActiveSheet.Name)!$($xl.ActiveCell.Address($false, $false))" }
$target = "'[$($wb.Name)]Model'!`$B`$2"
$xl.Run('Lattice.xlam!LtxNavJump', $target); "1. direct NavJump to $target -> $(Pos) (expect Model!B2)"
$calc.Activate(); $calc.Range('A2').Select() | Out-Null
$xl.Run('Lattice.xlam!LtxTraceIn'); Start-Sleep 2
"2. after TraceIn opened -> $(Pos) (expect Calc Sheet!A1)"
$xl.Run('Lattice.xlam!LtxNavJump', $target); "3. NavJump while window open -> $(Pos) (expect Model!B2)"
$calc.Activate(); $calc.Range('A2').Select() | Out-Null
$shell = New-Object -ComObject WScript.Shell
$ok = $shell.AppActivate('Lattice - Trace In'); "4. AppActivate(Trace In) = $ok"; Start-Sleep -Milliseconds 700
Shot 'trace_before_keys'
$shell.SendKeys('{RIGHT}'); Start-Sleep -Milliseconds 1000; "5. after Right -> $(Pos)"; Shot 'trace_after_right'
$shell.SendKeys('{DOWN}'); Start-Sleep -Milliseconds 1000; "6. after Down -> $(Pos) (expect Model!B2)"; Shot 'trace_after_down'
$shell.SendKeys('{ESC}'); Start-Sleep -Milliseconds 500
$wb.Close($false); $xl.Quit()
"done"
