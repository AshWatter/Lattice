# Figures for the designed guide (docs\guide_img): built from scratch in a separate Excel and
# PowerPoint, exported without stealing focus (range-to-picture via a chart, Slide.Export, PrintWindow).
param([string]$Book, [string]$OutDir, [string]$Xlam = "")
if (-not $Xlam) { $Xlam = Join-Path $env:APPDATA 'Microsoft/AddIns/Lattice.xlam' }
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class LtxShot2 {
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc p, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  public static IntPtr found;
  public static IntPtr FindByPrefix(string prefix) { found = IntPtr.Zero; EnumWindows((h, l) => { if (!IsWindowVisible(h)) return true; var sb = new StringBuilder(256); GetWindowText(h, sb, 256); if (sb.ToString().StartsWith(prefix)) { found = h; return false; } return true; }, IntPtr.Zero); return found; }
}
"@
function WinShot([IntPtr]$h, $name) {
    $rc = New-Object LtxShot2+RECT; [void][LtxShot2]::GetWindowRect($h, [ref]$rc)
    $w = $rc.R - $rc.L; $hh = $rc.B - $rc.T; if ($w -le 0) { "skip $name"; return }
    $bmp = New-Object Drawing.Bitmap $w, $hh; $g = [Drawing.Graphics]::FromImage($bmp); $hdc = $g.GetHdc(); [void][LtxShot2]::PrintWindow($h, $hdc, 2); $g.ReleaseHdc($hdc); $g.Dispose()
    $p = Join-Path $OutDir "$name.png"; $bmp.Save($p, [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose(); "saved $name"
}
# range (with the shapes over it) -> PNG, through a temporary chart
function RangeShot($ws, $addr, $name, $scale = 2) {
    $r = $ws.Range($addr); $r.CopyPicture(1, 2) | Out-Null     # xlScreen, xlBitmap
    $co = $ws.ChartObjects().Add(0, 0, $r.Width * $scale, $r.Height * $scale)
    $co.Chart.ChartArea.Format.Line.Visible = 0; $co.Chart.ChartArea.Format.Fill.Visible = 0
    $co.Chart.Paste() | Out-Null
    try { if ($co.Chart.Shapes.Count -ge 1) { $shp = $co.Chart.Shapes.Item(1); $shp.Left = 0; $shp.Top = 0; $shp.Width = $co.Width; $shp.Height = $co.Height } } catch { "note $name : $($_.Exception.Message)" }
    $p = Join-Path $OutDir "$name.png"; [void]$co.Chart.Export($p, 'PNG'); $co.Delete(); "saved $name"
}
New-Item -ItemType Directory -Force $OutDir | Out-Null
# ---------------- Excel ----------------
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
[void]$xl.Workbooks.Open($Xlam)
$wb = $xl.Workbooks.Open($Book); $ws = $wb.Worksheets.Item('Model'); $ws.Activate(); $xl.WindowState = -4137
$xl.ActiveWindow.DisplayGridlines = $false
# AutoColor sample block
$ws.Range('A12').Value2 = 'Revenue'; $ws.Range('B12').Value2 = 1250; $ws.Range('C12').Formula = '=B12*1.08'; $ws.Range('D12').Formula = '=C12*1.08'
$ws.Range('A13').Value2 = 'Costs'; $ws.Range('B13').Value2 = -800; $ws.Range('C13').Formula = "=B13*(1+Inputs!B2)"; $ws.Range('D13').Formula = '=C13*1.05+20'
$ws.Range('A14').Value2 = 'EBITDA'; $ws.Range('B14:D14').Formula = '=B12+B13'
$ws.Range('A11:D11').Value2 = @('EUR m', 'FY24', 'FY25', 'FY26'); $ws.Range('A11:D11').Font.Bold = $true
$ws.Range('B12:D14').NumberFormat = '#,##0;(#,##0)'
$ws.Range('A11:D14').Select() | Out-Null; [void]$xl.Run('Lattice.xlam!LtxAutoColor')
$ws.Range('A14:D14').Font.Bold = $true; $ws.Range('A14:D14').Borders.Item(8).LineStyle = 1
RangeShot $ws 'A11:D14' 'fig_autocolor' 3
# case drop-down wired
$ws.Range('F11').Value2 = 'Case'
$ws.Range('G11').Value2 = 'Base'; $ws.Range('G12').Value2 = 'Upside'; $ws.Range('G13').Value2 = 'Downside'
$ws.Range('G11:G13').Select() | Out-Null; [void]$xl.Run('Lattice.xlam!LtxInsertCaseSwitch')
$dd = $null; foreach ($s in $ws.Shapes) { if ($s.Name -like 'LtxCaseSwitch_*') { $dd = $s } }
if ($dd) { $dd.ControlFormat.LinkedCell = 'Model!$F$12'; $dd.ControlFormat.Value = 3; $dd.Left = $ws.Range('I11').Left; $dd.Top = $ws.Range('I11').Top }
$ws.Range('F12').NumberFormat = '"case #"0'
RangeShot $ws 'F11:K13' 'fig_dropdown' 3
# case runner report
$drv = $wb.Worksheets.Add(); $drv.Name = 'Drivers'; $drv.Range('B2').Value2 = 1
$ws.Range('A17').Value2 = 'Revenue'; $ws.Range('B17').Formula = '=CHOOSE(Drivers!B2,1250,1400,1050)'
$ws.Range('A18').Value2 = 'EBITDA';  $ws.Range('B18').Formula = '=B17*CHOOSE(Drivers!B2,0.28,0.31,0.22)'
$ws.Range('A19').Value2 = 'Net debt / EBITDA'; $ws.Range('B19').Formula = '=CHOOSE(Drivers!B2,3.1,2.6,4.4)'; $ws.Range('B19').NumberFormat = '0.0"x"'
$ws.Range('A20').Value2 = 'IRR'; $ws.Range('B20').Formula = '=CHOOSE(Drivers!B2,0.18,0.26,0.07)'; $ws.Range('B20').NumberFormat = '0.0%'
$ws.Range('A21').Value2 = 'MOIC'; $ws.Range('B21').Formula = '=CHOOSE(Drivers!B2,2.3,3.1,1.4)'; $ws.Range('B21').NumberFormat = '0.0"x"'
$ws.Range('B17:B18').NumberFormat = '#,##0'
[void]$xl.Run('Lattice.xlam!LtxRunCasesFor', 'Drivers!$B$2', 'Model!$B$17:$B$21', 'Base|Upside|Downside')
$rpt = $wb.Worksheets.Item('Lattice Cases'); $rpt.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
RangeShot $rpt 'A1:G9' 'fig_cases' 2
# model check report sheet
$ws.Activate(); $ws.Range('A1:D14').Select() | Out-Null
[void]$xl.Run('Lattice.xlam!LtxAuditHardcodes', $false); Start-Sleep 1
$aud = $null; foreach ($s in $wb.Worksheets) { if ($s.Name -like 'Lattice*' -and $s.Name -ne 'Lattice Cases') { $aud = $s } }
if ($aud) { $aud.Activate(); $xl.ActiveWindow.DisplayGridlines = $false; $last = $aud.UsedRange.Rows.Count; if ($last -gt 12) { $last = 12 }; RangeShot $aud "A1:E$last" 'fig_audit' 2 }
# chart with annotations (bigger)
$ws.Activate()
$ws.Range('I17:L17').Value2 = @('', 'FY24', 'FY25', 'FY26'); $ws.Range('I18').Value2 = 'Core'; $ws.Range('J18').Value2 = 620; $ws.Range('K18').Value2 = 700; $ws.Range('L18').Value2 = 810
$ws.Range('I19').Value2 = 'Services'; $ws.Range('J19').Value2 = 310; $ws.Range('K19').Value2 = 360; $ws.Range('L19').Value2 = 440
$ws.Range('I20').Value2 = 'Other'; $ws.Range('J20').Value2 = 120; $ws.Range('K20').Value2 = 130; $ws.Range('L20').Value2 = 150
$co = $ws.ChartObjects().Add(500, 300, 520, 320); $co.Name = 'Ch2'; $co.Chart.ChartType = 52; $co.Chart.SetSourceData($ws.Range('I17:L20'), 1); $co.Chart.HasLegend = $true
$co.Chart.Legend.Position = -4107; $co.Chart.ChartArea.Format.Line.Visible = 0
$co.Chart.Axes(2).HasMajorGridlines = $true; $co.Chart.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = 15921906; $co.Chart.Axes(2).Format.Line.Visible = 0
$i = 1; foreach ($c in 6567967, 11957550, 15652797) { $co.Chart.SeriesCollection($i).Format.Fill.ForeColor.RGB = $c; $i++ }
[void]$xl.Run('Lattice.xlam!LtxChartNote', 'Model', 'Ch2', 'totals', 0, 0); [void]$xl.Run('Lattice.xlam!LtxChartNote', 'Model', 'Ch2', 'cagr', 0, 0); [void]$xl.Run('Lattice.xlam!LtxChartNote', 'Model', 'Ch2', 'delta', 1, 3)
[void]$co.Chart.Export((Join-Path $OutDir 'fig_chart.png'), 'PNG'); "saved fig_chart"
# windows
$ws.Range('C13').Select() | Out-Null; $xl.Run('Lattice.xlam!LtxTraceIn'); Start-Sleep 2
$h = [LtxShot2]::FindByPrefix('Lattice - Trace In'); WinShot $h 'fig_trace'; $xl.Run('Lattice.xlam!LtxCloseSettings')
$xl.Run('Lattice.xlam!LtxModelCheckPane'); Start-Sleep 2
$h = [LtxShot2]::FindByPrefix('Lattice - Model Check'); WinShot $h 'fig_modelcheck'
$xl.Run('Lattice.xlam!LtxOpenSettings'); Start-Sleep 2
$h = [LtxShot2]::FindByPrefix('Lattice - Settings'); WinShot $h 'fig_settings_excel'; $xl.Run('Lattice.xlam!LtxCloseSettings')
$ws.Activate(); $xl.Run('Lattice.xlam!LtxShowTab'); Start-Sleep 1
$rc = New-Object LtxShot2+RECT; [void][LtxShot2]::GetWindowRect([IntPtr]$xl.Hwnd, [ref]$rc)
$bmp = New-Object Drawing.Bitmap ($rc.R - $rc.L), ($rc.B - $rc.T); $g = [Drawing.Graphics]::FromImage($bmp); $hdc = $g.GetHdc(); [void][LtxShot2]::PrintWindow([IntPtr]$xl.Hwnd, $hdc, 2); $g.ReleaseHdc($hdc); $g.Dispose()
$crop = $bmp.Clone((New-Object Drawing.Rectangle 8, 56, ($bmp.Width - 16), 130), $bmp.PixelFormat); $crop.Save((Join-Path $OutDir 'fig_ribbon_excel.png'), [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose(); $crop.Dispose(); "saved fig_ribbon_excel"
# ---------------- PowerPoint (Excel stays open for the import) ----------------
$pp = New-Object -ComObject PowerPoint.Application; $pp.Visible = -1
$pres = $pp.Presentations.Add(-1); [void]$pres.Slides.Add(1, 12); [void]$pres.Slides.Add(2, 12); $pp.ActiveWindow.WindowState = 2
function RunP($macro, $arg) { if ($null -ne $arg) { $script:pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $script:pp, @($macro, $arg)) } else { $script:pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $script:pp, @($macro)) } }
$sl = $pres.Slides.Item(1); $pp.ActiveWindow.View.GotoSlide(1)
$t = $sl.Shapes.AddTextbox(1, 40, 30, 600, 50); $t.TextFrame.TextRange.Text = 'Workstream status'; $t.TextFrame.TextRange.Font.Size = 28; $t.TextFrame.TextRange.Font.Bold = -1; $t.TextFrame.TextRange.Font.Color.RGB = 6567967
$rows = @('Commercial due diligence', 'Financial model and cases', 'Financing package', 'Legal review')
$y = 110; $i = 0
foreach ($r in $rows) {
    $b = $sl.Shapes.AddTextbox(1, 60, $y, 330, 28); $b.TextFrame.TextRange.Text = $r; $b.TextFrame.TextRange.Font.Size = 16; $b.Name = "Row $i"
    $b.Select(); RunP 'Lattice.ppam!LtxTurbo' (@('harvey100', 'harvey75', 'harvey50', 'harvey25')[$i])
    $hb = $sl.Shapes.Item($sl.Shapes.Count); $hb.Left = 420; $hb.Top = $y + 2
    $b.Select(); RunP 'Lattice.ppam!LtxSticker' (@('light_green', 'light_amber', 'light_amber', 'light_red')[$i])
    $lt = $sl.Shapes.Item($sl.Shapes.Count); $lt.Left = 470; $lt.Top = $y + 5
    $b.Select(); RunP 'Lattice.ppam!LtxSticker' (@('tick', 'tick', 'cross', 'cross')[$i])
    $mk = $sl.Shapes.Item($sl.Shapes.Count); $mk.Left = 510; $mk.Top = $y + 1
    $y += 44; $i++
}
RunP 'Lattice.ppam!LtxSticker' 'stamp_draft'
[void]$sl.Export((Join-Path $OutDir 'fig_slide_stickers.png'), 'PNG', 1600, 900); "saved fig_slide_stickers"
# slide 2: imported Excel picture (fonts black) + chevrons + CONFIDENTIAL stamp
$sl2 = $pres.Slides.Item(2); $pp.ActiveWindow.View.GotoSlide(2)
$t2 = $sl2.Shapes.AddTextbox(1, 40, 30, 600, 50); $t2.TextFrame.TextRange.Text = 'Base case summary'; $t2.TextFrame.TextRange.Font.Size = 28; $t2.TextFrame.TextRange.Font.Bold = -1; $t2.TextFrame.TextRange.Font.Color.RGB = 6567967
$ws.Activate(); $ws.Range('A11:D14').Select() | Out-Null
RunP 'Lattice.ppam!LtxImport' 'picture'
$pic = $sl2.Shapes.Item($sl2.Shapes.Count); $pic.Left = 60; $pic.Top = 110; $pic.Width = 380
$sl2.Shapes.Item(1).Select(); RunP 'Lattice.ppam!LtxTurbo' 'chevrons'
$k = 0; foreach ($s in $sl2.Shapes) { if ($s.Name -like 'Chevron*') { $s.Top = 300; $s.Left = 60 + $k * 115; $k++ } }
RunP 'Lattice.ppam!LtxSticker' 'stamp_confidential'
[void]$sl2.Export((Join-Path $OutDir 'fig_slide_link.png'), 'PNG', 1600, 900); "saved fig_slide_link"
RunP 'Lattice.ppam!LtxOpenSettings'; Start-Sleep 2
$h = [LtxShot2]::FindByPrefix('Lattice - Settings'); WinShot $h 'fig_settings_ppt'; RunP 'Lattice.ppam!LtxCloseSettings'
$pres.Saved = -1; $pres.Close(); $pp.Quit()
$wb.Close($false); $xl.DisplayAlerts = $false; $xl.Quit()
"done"
