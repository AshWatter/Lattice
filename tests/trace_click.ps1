# Trace In window driven like a user: mouse clicks on rows, then arrow keys, across sheets.
# Prints PASS/FAIL lines; screenshots in $OutDir.
param([string]$Book, [string]$OutDir, [string]$Xlam = "")
if (-not $Xlam) { $Xlam = Join-Path $env:APPDATA "Microsoft/AddIns/Lattice.xlam" }
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Add-Type -AssemblyName System.Drawing; Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class LtxMouse {
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
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
  public static void Click(int x, int y) { SetCursorPos(x, y); System.Threading.Thread.Sleep(80); mouse_event(2, 0, 0, 0, UIntPtr.Zero); mouse_event(4, 0, 0, 0, UIntPtr.Zero); }
}
"@
function Shot($name) {
    $b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bmp = New-Object Drawing.Bitmap $b.Width, $b.Height; $g = [Drawing.Graphics]::FromImage($bmp); $g.CopyFromScreen($b.Left, $b.Top, 0, 0, $bmp.Size)
    $small = New-Object Drawing.Bitmap $bmp, ([int]($b.Width / 2)), ([int]($b.Height / 2))
    $p = Join-Path $OutDir "$name.png"; $small.Save($p, [Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose(); $small.Dispose(); "saved $p"
}
$fails = 0
function Check($name, $ok, $detail) { if ($ok) { "PASS $name" } else { $script:fails++; "FAIL $name  $detail" } }
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' -or $_.MainWindowTitle -match '^TestModel' } | Stop-Process -Force; Start-Sleep 1
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$wb = $xl.Workbooks.Open($Book)
if (-not (($xl.Workbooks | % Name) -contains 'Lattice.xlam')) { [void]$xl.Workbooks.Open("$env:APPDATA\Microsoft\AddIns\Lattice.xlam") }
$calc = $wb.Worksheets.Add(); $calc.Name = 'Calc Sheet'
$calc.Range('A1').Formula = '=Model!B2*2'; $calc.Range('A2').Formula = "='Calc Sheet'!A1+1"
$calc.Activate(); $calc.Range('A2').Select() | Out-Null; $xl.WindowState = -4137
function Pos { "$($xl.ActiveSheet.Name)!$($xl.ActiveCell.Address($false, $false))" }
$xl.Run('Lattice.xlam!LtxTraceIn'); Start-Sleep 2
Check 'window opens on the first reference' ((Pos) -eq 'Calc Sheet!A1') (Pos)
$h = [LtxMouse]::FindByPrefix('Lattice - Trace In')
$rc = New-Object LtxMouse+RECT; [void][LtxMouse]::GetWindowRect($h, [ref]$rc)
"form handle: $h  rect: $($rc.L),$($rc.T) - $($rc.R),$($rc.B)"
[void][LtxMouse]::SetForegroundWindow($h); Start-Sleep -Milliseconds 400
if ([LtxMouse]::GetForegroundWindow() -ne $h) { "SKIP: the Trace window could not be brought to the front (another app is active), not clicking"; $shell0 = New-Object -ComObject WScript.Shell; $wb.Close($false); $xl.Quit(); exit 0 }
# list rows measured from the screenshots: first row ~140 px below the form's top edge, rows ~13 px apart
$x = $rc.L + 110; $rowRoot = $rc.T + 154; $rowChild = $rowRoot + 13   # list sits 14 px lower since the column headers were added
[LtxMouse]::Click($x, $rowRoot); Start-Sleep -Milliseconds 900
Check 'click on root row moves Excel to Calc Sheet!A2' ((Pos) -eq 'Calc Sheet!A2') (Pos)
[LtxMouse]::Click($x, $rowChild); Start-Sleep -Milliseconds 900
Check 'click on child row moves Excel to Calc Sheet!A1' ((Pos) -eq 'Calc Sheet!A1') (Pos)
Shot 'click_after'
$shell = New-Object -ComObject WScript.Shell
$shell.SendKeys('{RIGHT}'); Start-Sleep -Milliseconds 1200          # expand Calc Sheet!A1 -> Model!B2
Shot 'after_right'
$shell.SendKeys('{DOWN}'); Start-Sleep -Milliseconds 1000
Check 'Right then Down walks into Model!B2 on another sheet' ((Pos) -eq 'Model!B2') (Pos)
Shot 'keys_after'
# Enter must jump and close the window so F2 works straight away
$shell.SendKeys('{ENTER}'); Start-Sleep -Milliseconds 900
$gone = ([LtxMouse]::FindByPrefix('Lattice - Trace In') -eq [IntPtr]::Zero)
Check 'Enter closes the window and leaves Excel on Model!B2' ($gone -and (Pos) -eq 'Model!B2') "open=$(-not $gone) pos=$(Pos)"
if (-not $gone) { $shell.SendKeys('{ESC}'); Start-Sleep -Milliseconds 500 }
$wb.Close($false); $xl.Quit()
"trace_click: $fails failure(s)"
