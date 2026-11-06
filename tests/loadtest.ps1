# Loads each given .ppam into a fresh PowerPoint and reports whether PowerPoint accepted it.
# A watcher thread closes the "couldn't load add-in" NUIDialog so the test never blocks.
param([string[]]$Files)
$ErrorActionPreference = 'Continue'
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class LW {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc p, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  public static uint PidOf(IntPtr hwnd) { uint p; GetWindowThreadProcessId(hwnd, out p); return p; }
  // closes a visible NUIDialog / #32770 of the process; returns true if one was found
  public static bool CloseDialog(uint pid) { bool found = false;
    EnumWindows((h,l)=>{ uint p; GetWindowThreadProcessId(h, out p); if (p!=pid || !IsWindowVisible(h)) return true;
      var c=new StringBuilder(128); GetClassName(h,c,128); string cn=c.ToString();
      if (cn=="NUIDialog" || cn=="#32770") { found=true; PostMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); }
      return true; }, IntPtr.Zero);
    return found; }
}
"@
foreach ($f in $Files) {
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 700
    $pp = New-Object -ComObject PowerPoint.Application
    $procId = [LW]::PidOf([IntPtr]$pp.HWND)
    $shared = [hashtable]::Synchronized(@{ dlg = $false; stop = $false })
    $rs = [runspacefactory]::CreateRunspace(); $rs.Open()
    $rs.SessionStateProxy.SetVariable('shared', $shared); $rs.SessionStateProxy.SetVariable('procId', $procId)
    $ps = [powershell]::Create(); $ps.Runspace = $rs
    [void]$ps.AddScript({ while (-not $shared.stop) { if ([LW]::CloseDialog($procId)) { $shared.dlg = $true }; Start-Sleep -Milliseconds 150 } })
    $h = $ps.BeginInvoke()
    $res = ''
    try {
        $ai = $pp.AddIns.Add($f)
        $ai.Loaded = -1
        Start-Sleep -Milliseconds 800
        $res = "Loaded=$($ai.Loaded)"
    } catch { $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }; $res = "EXC: " + ($e.Message -replace "`r?`n", ' ') }
    Start-Sleep -Milliseconds 500
    $shared.stop = $true; $ps.EndInvoke($h); $ps.Dispose(); $rs.Close()
    "{0,-34} {1}  dialog={2}" -f (Split-Path $f -Leaf), $res, $shared.dlg
    try { $pp.Quit() } catch {}
    Start-Sleep -Milliseconds 300
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
}
