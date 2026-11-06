# Captures the Lattice surfaces without stealing focus: the Excel and PowerPoint ribbons with the
# Lattice tab active, the Trace In, Model Check and Settings windows, the case-runner report and a
# chart with annotations. Uses PrintWindow, so other apps can stay in front. PNGs land in $OutDir.
param([string]$Book, [string]$OutDir, [string]$Xlam = "")
if (-not $Xlam) { $Xlam = Join-Path $env:APPDATA 'Microsoft/AddIns/Lattice.xlam' }
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class LtxShot {
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc p, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  public static IntPtr found;
  public static IntPtr FindByPrefix(string prefix) {
    found = IntPtr.Zero;
    EnumWindows((h, l) => { if (!IsWindowVisible(h)) return true; var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
      if (sb.ToString().StartsWith(prefix)) { found = h; return false; } return true; }, IntPtr.Zero);
    return found;
  }
}
"@
function Shot([IntPtr]$h, $name, $cropH = 0) {
    $rc = New-Object LtxShot+RECT; [void][LtxShot]::GetWindowRect($h, [ref]$rc)
    $w = $rc.R - $rc.L; $hh = $rc.B - $rc.T; if ($w -le 0 -or $hh -le 0) { "skip $name (no window)"; return }
    $bmp = New-Object Drawing.Bitmap $w, $hh; $g = [Drawing.Graphics]::FromImage($bmp); $hdc = $g.GetHdc()
    [void][LtxShot]::PrintWindow($h, $hdc, 2); $g.ReleaseHdc($hdc); $g.Dispose()
    if ($cropH -gt 0 -and $cropH -lt $hh) { $c = $bmp.Clone((New-Object Drawing.Rectangle 0, 0, $w, $cropH), $bmp.PixelFormat); $bmp.Dispose(); $bmp = $c }
    $p = Join-Path $OutDir "$name.png"; $bmp.Save($p, [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose(); "saved $p"
}
New-Item -ItemType Directory -Force $OutDir | Out-Null
# ---- Excel ----
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
[void]$xl.Workbooks.Open($Xlam)
$wb = $xl.Workbooks.Open($Book); $ws = $wb.Worksheets.Item('Model'); $ws.Activate(); $xl.WindowState = -4137
$xl.Run('Lattice.xlam!LtxShowTab'); Start-Sleep 1
Shot ([IntPtr]$xl.Hwnd) 'ui_ribbon_excel' 210
# case runner report
$drv = $wb.Worksheets.Add(); $drv.Name = 'Drivers'; $drv.Range('B2').Value2 = 1
$ws.Range('F1').Value2 = 'Revenue'; $ws.Range('G1').Formula = '=CHOOSE(Drivers!B2,100,120,80)'
$ws.Range('F2').Value2 = 'EBITDA';  $ws.Range('G2').Formula = '=G1*0.3'
$ws.Range('F3').Value2 = 'IRR';     $ws.Range('G3').Formula = '=CHOOSE(Drivers!B2,0.15,0.22,0.08)'; $ws.Range('G3').NumberFormat = '0.0%'
[void]$xl.Run('Lattice.xlam!LtxRunCasesFor', 'Drivers!$B$2', 'Model!$G$1:$G$3', 'Base|Upside|Downside'); Start-Sleep 1
Shot ([IntPtr]$xl.Hwnd) 'ui_report_cases' 420
# chart with annotations
$ws.Activate()
$ws.Range('I1:K1').Value2 = @('2022', '2023', '2024'); $ws.Range('H2').Value2 = 'Revenue'; $ws.Range('I2').Value2 = 100; $ws.Range('J2').Value2 = 120; $ws.Range('K2').Value2 = 150
$ws.Range('H3').Value2 = 'Other'; $ws.Range('I3').Value2 = 50; $ws.Range('J3').Value2 = 55; $ws.Range('K3').Value2 = 60
$co = $ws.ChartObjects().Add(420, 60, 420, 260); $co.Name = 'Ch1'; $co.Chart.ChartType = 52; $co.Chart.SetSourceData($ws.Range('H1:K3'), 1); $co.Chart.HasLegend = $true
[void]$xl.Run('Lattice.xlam!LtxChartNote', 'Model', 'Ch1', 'totals', 0, 0); [void]$xl.Run('Lattice.xlam!LtxChartNote', 'Model', 'Ch1', 'cagr', 0, 0); [void]$xl.Run('Lattice.xlam!LtxChartNote', 'Model', 'Ch1', 'delta', 1, 3)
$co.Chart.Export((Join-Path $OutDir 'ui_chart_annotations.png'), 'PNG'); "saved chart"
# Trace In window
$ws.Range('G2').Select() | Out-Null; $xl.Run('Lattice.xlam!LtxTraceIn'); Start-Sleep 2
$h = [LtxShot]::FindByPrefix('Lattice - Trace In'); Shot $h 'ui_window_trace'
if ($h -ne [IntPtr]::Zero) { $xl.Run('Lattice.xlam!LtxCloseSettings'); Start-Sleep -Milliseconds 300 }   # closes any Lattice window by name below
# Model Check pane
$ws.Range('G5').Formula = '=G1*1.05+350'
$xl.Run('Lattice.xlam!LtxModelCheckPane'); Start-Sleep 2
$h = [LtxShot]::FindByPrefix('Lattice - Model Check'); Shot $h 'ui_window_modelcheck'
# Settings dialog
$xl.Run('Lattice.xlam!LtxOpenSettings'); Start-Sleep 2
$h = [LtxShot]::FindByPrefix('Lattice - Settings'); Shot $h 'ui_window_settings_excel'
$xl.Run('Lattice.xlam!LtxCloseSettings')
$wb.Close($false); $xl.DisplayAlerts = $false; $xl.Quit()
# ---- PowerPoint ----
$pp = New-Object -ComObject PowerPoint.Application; $pp.Visible = -1
$pres = $pp.Presentations.Add(-1); [void]$pres.Slides.Add(1, 12); $pp.ActiveWindow.WindowState = 2; Start-Sleep 2
$pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @('Lattice.ppam!LtxShowTab')); Start-Sleep 1
Shot ([IntPtr]$pp.HWND) 'ui_ribbon_ppt' 210
$pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @('Lattice.ppam!LtxOpenSettings')); Start-Sleep 2
$h = [LtxShot]::FindByPrefix('Lattice - Settings'); Shot $h 'ui_window_settings_ppt'
$pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @('Lattice.ppam!LtxCloseSettings'))
$pres.Saved = -1; $pres.Close(); $pp.Quit()
"done"
