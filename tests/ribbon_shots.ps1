# Captures the Lattice tab in PowerPoint and Excel (icons check) and times a refresh of a
# larger coloured range (speed check). Output PNGs + timings in $OutDir.
param([string]$Book, [string]$OutDir)
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Add-Type -AssemblyName System.Drawing
$shell = New-Object -ComObject WScript.Shell
function CropWin($left, $top, $width, $height, $name) {
    $sx = [int]($left * 96 / 72); $sy = [int]($top * 96 / 72); $sw = [int]($width * 96 / 72)
    $b = New-Object Drawing.Bitmap $sw, $height; $g = [Drawing.Graphics]::FromImage($b); $g.CopyFromScreen($sx, $sy, 0, 0, $b.Size)
    $p = Join-Path $OutDir "$name.png"; $b.Save($p, [Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $b.Dispose(); "saved $p"
}
Get-Process EXCEL,POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force; Start-Sleep 1
# ---- PowerPoint ribbon ----
$pp = New-Object -ComObject PowerPoint.Application; $pp.Visible = -1
"ppt add-in loaded: " + (($pp.AddIns | ? { $_.Name -eq 'Lattice' }).Loaded)
$pres = $pp.Presentations.Add(-1); [void]$pres.Slides.Add(1, 12); $pp.ActiveWindow.WindowState = 2; Start-Sleep 2
$pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @('Lattice.ppam!LtxShowTab')); Start-Sleep 1
$pp.Activate(); $shell.AppActivate($pp.Caption) | Out-Null; Start-Sleep -Milliseconds 800; $pp.GetType().InvokeMember("Run", [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @("Lattice.ppam!LtxShowTab")); Start-Sleep -Milliseconds 700; $w = $pp.ActiveWindow; CropWin $w.Left $w.Top $w.Width 190 'ribbon_ppt_icons'
$pp.Activate(); $shell.AppActivate($pp.Caption) | Out-Null; Start-Sleep -Milliseconds 600; $shell.SendKeys('%'); Start-Sleep -Milliseconds 900
CropWin $w.Left $w.Top $w.Width 120 'keytips_ppt'; $shell.SendKeys('{ESC}'); Start-Sleep -Milliseconds 400
$pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @('Lattice.ppam!LtxSet', 'Profile', '1'))
# ---- refresh timing on a 40x10 coloured block ----
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$wb = $xl.Workbooks.Open($Book); $ws = $wb.Worksheets.Item('Model')
$big = $wb.Worksheets.Add(); $big.Name = 'Big'
$big.Range('A1:J1').Value2 = @('Line','Y1','Y2','Y3','Y4','Y5','Y6','Y7','Y8','Y9')
$big.Range('A1:J1').Font.Bold = $true; $big.Range('A1:J1').Interior.Color = 8210719; $big.Range('A1:J1').Font.Color = 16777215
for ($r = 2; $r -le 40; $r++) { $big.Cells.Item($r, 1).Value2 = "Item $r"; for ($c = 2; $c -le 10; $c++) { if ($r % 3 -eq 0) { $big.Cells.Item($r, $c).Value2 = $r * $c; $big.Cells.Item($r, $c).Font.Color = 16711680 } elseif ($r % 3 -eq 1) { $big.Cells.Item($r, $c).Formula = "=" + $big.Cells.Item(2, $c).Address($false,$false) + "*1.1"; $big.Cells.Item($r, $c).Font.Color = 0 } else { $big.Cells.Item($r, $c).Formula = "=Model!B2*$r"; $big.Cells.Item($r, $c).Font.Color = 32768 } } }
$big.Activate(); $big.Range('A1:J40').Select()
$procId = [LtxWin]::PidOf([IntPtr]$pp.HWND)
$pp.ActiveWindow.View.GotoSlide(1)
$t0 = Get-Date
$m = Invoke-Watched $procId { $pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @('Lattice.ppam!LtxImport', 'picture')) }
"import 40x10 coloured block: {0:N1}s  {1}" -f ((Get-Date) - $t0).TotalSeconds, $m
$times = @(); foreach ($i in 1..3) { $t0 = Get-Date; $m = Invoke-Watched $procId { $pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @('Lattice.ppam!LtxRefreshAll')) }; $times += ((Get-Date) - $t0).TotalSeconds }
"refresh deck (1 link) x3: " + (($times | % { "{0:N1}s" -f $_ }) -join ', ')
"source colours intact after refresh: B3=$($big.Range('B3').Font.Color) (expect 16711680)  B5=$($big.Range('B5').Font.Color) (expect 32768)"
$pres.Saved = -1; $pres.Close(); $pp.Quit()
# ---- Excel ribbon ----
$xl.WindowState = -4137; $ws.Activate(); Start-Sleep 1
if (-not (($xl.Workbooks | % Name) -contains 'Lattice.xlam')) { [void]$xl.Workbooks.Open("$env:APPDATA\Microsoft\AddIns\Lattice.xlam") }
$xl.Run('Lattice.xlam!LtxShowTab'); Start-Sleep 1
CropWin $xl.Left $xl.Top $xl.Width 200 'ribbon_excel_icons'
$shell.AppActivate($xl.Caption) | Out-Null; Start-Sleep -Milliseconds 600; $shell.SendKeys('%'); Start-Sleep -Milliseconds 900
CropWin $xl.Left $xl.Top $xl.Width 120 'keytips_excel'; $shell.SendKeys('{ESC}'); Start-Sleep -Milliseconds 400
"--- profile log (last refresh) ---"; Get-Content "$env:APPDATA\Lattice\lattice.log" -Tail 14
$wb.Close($false); $xl.Quit()
"done"
