# =====================================================================
#  Lattice build script
#  Creates dist\Lattice.xlam, dist\Lattice.ppam and dist\Lattice.dotm from
#  the VBA sources in src\ and the ribbon XML in ribbon\.
#  Every project is compiled through the VBE before it is saved; the first
#  compile error (module + line) is printed and the build stops.
#  Run:  powershell -ExecutionPolicy Bypass -File build\build.ps1
# =====================================================================

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Force $dist | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Write-Step($m) { Write-Host "[Lattice] $m" -ForegroundColor Cyan }

# Win32 helper: finds the VBE's modal "Microsoft Visual Basic" dialog of a given
# process, returns its text and presses OK/End. Used to read compile errors.
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class LtxWin {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc p, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc p, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  public static uint PidOf(IntPtr hwnd) { uint p; GetWindowThreadProcessId(hwnd, out p); return p; }
  public static string FindDialog(uint pid, bool click) { string result = null;
    EnumWindows((h,l)=>{ uint p; GetWindowThreadProcessId(h, out p); if (p!=pid) return true;
      var c=new StringBuilder(128); GetClassName(h,c,128);
      if (c.ToString()=="bosa_sdm_XL9") { /* Excel Application.InputBox / RefEdit prompt: cancel it */ SendMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); result = "[RefEdit prompt cancelled]"; return true; }
      if (c.ToString()!="#32770") return true;
      var t=new StringBuilder(512); GetWindowText(h,t,512); if (!IsWindowVisible(h)) return true;
      string txt=""; IntPtr btn=IntPtr.Zero; IntPtr anyBtn=IntPtr.Zero; IntPtr yesBtn=IntPtr.Zero;
      EnumChildWindows(h,(ch,ll)=>{ var tt=new StringBuilder(4096); GetWindowText(ch,tt,4096); var cc=new StringBuilder(128); GetClassName(ch,cc,128);
        if (cc.ToString()=="Static") txt+=tt.ToString()+" ";
        if (cc.ToString()=="Button") { if (anyBtn==IntPtr.Zero) anyBtn=ch; string bt=tt.ToString(); if (bt=="&Yes") yesBtn=ch; if ((bt=="OK" || bt=="&OK" || bt=="&End" || bt=="Cancel") && btn==IntPtr.Zero) btn=ch; } return true; }, IntPtr.Zero);
      if (yesBtn!=IntPtr.Zero) btn=yesBtn;   /* tests answer Yes to Yes/No/Cancel prompts (e.g. version choice) */
      if (btn==IntPtr.Zero) btn=anyBtn;
      bool isVbe = t.ToString().StartsWith("Microsoft Visual Basic");
      result = (isVbe ? "" : "[" + t + "] ") + txt.Trim();
      if (click && btn!=IntPtr.Zero) SendMessage(btn, 0x00F5, IntPtr.Zero, IntPtr.Zero);
      return !isVbe; /* keep looking for a VBE dialog after dismissing an app message box */ }, IntPtr.Zero);
    return result; }
}
"@

# "Trust access to the VBA project object model" is needed to inject code.
function Ensure-VBOM($app) {
    $k = "HKCU:\Software\Microsoft\Office\16.0\$app\Security"
    if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null }
    $v = (Get-ItemProperty $k -ErrorAction SilentlyContinue).AccessVBOM
    if ($v -ne 1) {
        Write-Step "Enabling 'Trust access to the VBA project object model' for $app (HKCU, dev setting)"
        Set-ItemProperty $k -Name AccessVBOM -Value 1 -Type DWord
    }
}

# Static lint: VBA identifiers are case-insensitive, so a variable called nExt or
# name is a reserved word. The VBE only reports these once it has parsed the
# module interactively, so catch them before Office sees the file.
$script:VbaKeywords = @('and','as','boolean','byref','byte','byval','call','case','class','close','const','currency','date','debug','declare','dim','do','double','each','else','elseif','empty','end','endif','enum','eqv','erase','error','event','exit','false','for','friend','function','get','gosub','goto','if','imp','implements','in','input','integer','is','let','lib','like','long','loop','lset','me','mod','new','next','not','nothing','null','on','open','option','optional','or','paramarray','preserve','print','private','property','public','raiseevent','redim','rem','resume','return','rset','seek','select','set','shared','single','spc','static','step','stop','string','sub','tab','then','to','true','type','typeof','until','variant','wend','while','with','withevents','write','xor')
function Test-Keywords([string[]]$files) {
    $bad = @()
    foreach ($f in $files) {
        $n = 0
        foreach ($line in Get-Content $f) {
            $n++
            $code = ($line -split "'")[0]
            $names = @()
            if ($code -match '^\s*(Public |Private |Static )?(Dim|Const)\s+(.*)$') {
                $names += ($matches[3] -split ',') | ForEach-Object { ($_ -split '\s+As\s+|=|\(')[0].Trim() }
            } elseif ($code -match '^\s*(Public |Private |Friend )?(Sub|Function|Property \w+)\s+(\w+)\s*\((.*?)\)') {
                if ($matches[3].Length -eq 1) { $bad += "$(Split-Path $f -Leaf):$n  one-letter procedure name '$($matches[3])' collides with loop variables" }
                $matches = @{ 3 = $matches[4] }
                $names += ($matches[3] -split ',') | ForEach-Object { (($_ -replace '^\s*(Optional\s+|ByVal\s+|ByRef\s+|ParamArray\s+)*', '') -split '\s+As\s+|=')[0].Trim() }
            }
            foreach ($nm in $names) { if ($nm -and ($script:VbaKeywords -contains $nm.ToLower())) { $bad += "$(Split-Path $f -Leaf):$n  '$nm' is a VBA keyword" } }
        }
    }
    if ($bad.Count -gt 0) { throw ("Reserved-word identifiers found:`n  " + ($bad -join "`n  ")) }
}

function Import-Modules($vbproj, [string[]]$files) {
    Test-Keywords $files
    foreach ($f in $files) {
        $name = [IO.Path]::GetFileNameWithoutExtension($f)
        try { $old = $vbproj.VBComponents.Item($name); $vbproj.VBComponents.Remove($old) } catch {}
        [void]$vbproj.VBComponents.Import($f)
        Write-Step "  imported $name"
    }
}

# Builds a list-style UserForm from code (no .frx binaries in the repo):
# one multi-column ListBox, an info label and a row of buttons.
function Add-ListForm($vbproj, [string]$name, [string]$caption, [string]$listName, [int]$cols, [string]$colWidths, $buttons, [string]$codePath) {
    $comp = $vbproj.VBComponents.Add(3)          # vbext_ct_MSForm
    $comp.Name = $name
    $comp.Properties.Item('Caption').Value = $caption
    $comp.Properties.Item('Width').Value = 560
    $comp.Properties.Item('Height').Value = 340
    $d = $comp.Designer
    $lst = $d.Controls.Add('Forms.ListBox.1', $listName, $true)
    $lst.Left = 8; $lst.Top = 8; $lst.Width = 540; $lst.Height = 240
    $lst.ColumnCount = $cols
    $lst.ColumnWidths = $colWidths
    $lbl = $d.Controls.Add('Forms.Label.1', 'lblInfo', $true)
    $lbl.Left = 8; $lbl.Top = 252; $lbl.Width = 540; $lbl.Height = 14
    $x = 8
    foreach ($b in $buttons) {
        $btn = $d.Controls.Add('Forms.CommandButton.1', $b.n, $true)
        $btn.Caption = $b.c; $btn.Left = $x; $btn.Top = 272; $btn.Width = 72; $btn.Height = 24
        $x += 76
    }
    $comp.CodeModule.AddFromString((Get-Content $codePath -Raw))
    Write-Step "  built $name"
}

function Add-ExplorerForm($vbproj, [string]$codePath) {
    Add-ListForm $vbproj 'frmLtxExplorer' 'Lattice - Link Explorer' 'lstLinks' 6 '34;120;120;80;80;100' @(
        @{n='btnRefresh';c='Refresh'}, @{n='btnRefreshAll';c='Refresh All'}, @{n='btnGoTo';c='Go to Source'},
        @{n='btnSource';c='Change Src'}, @{n='btnBreak';c='Break Link'}, @{n='btnReload';c='Reload'}, @{n='btnClose';c='Close'}) $codePath
}

function Add-ModelCheckForm($vbproj, [string]$codePath) {
    Add-ListForm $vbproj 'frmLtxModelCheck' 'Lattice - Model Check' 'lstIssues' 4 '90;60;150;230' @(
        @{n='btnSheet';c='Analyze Sheet'}, @{n='btnBook';c='Analyze Book'}, @{n='btnGo';c='Go to Cell'},
        @{n='btnFix';c='Fix Hard-code'}, @{n='btnReport';c='Report Sheet'}, @{n='btnClose';c='Close'}) $codePath
}

# Runs $action on this thread while a background thread watches the host
# process for the VBE's modal error dialog, reads it and dismisses it.
# Returns the dialog text, or $null when no dialog appeared.
function Invoke-Watched([uint32]$procId, [scriptblock]$action) {
    $shared = [hashtable]::Synchronized(@{ msg = $null; stop = $false })
    $rs = [runspacefactory]::CreateRunspace(); $rs.Open()
    $rs.SessionStateProxy.SetVariable('shared', $shared); $rs.SessionStateProxy.SetVariable('procId', $procId)
    $ps = [powershell]::Create(); $ps.Runspace = $rs
    [void]$ps.AddScript({
        # keep dismissing application message boxes until the action returns;
        # stop only on a VBE (compile / run-time error) dialog
        while (-not $shared.stop) {
            $m = [LtxWin]::FindDialog($procId, $true)
            if ($m -ne $null) {
                if ($shared.msg -eq $null) { $shared.msg = $m } else { $shared.msg = $shared.msg + ' || ' + $m }
                if (-not $m.StartsWith('[')) { break }     # VBE dialog: leave the rest to the caller
                Start-Sleep -Milliseconds 400                # app message box: give the macro time to continue
            }
            Start-Sleep -Milliseconds 150
        }
    })
    $h = $ps.BeginInvoke()
    try { & $action } catch { if ($shared.msg -eq $null) { $shared.msg = "(host error) " + $_.Exception.Message } }
    Start-Sleep -Milliseconds 400
    $shared.stop = $true
    $ps.EndInvoke($h); $ps.Dispose(); $rs.Close()
    return $shared.msg
}

function Get-CompileLocation($vbe) {
    try {
        $cp = $vbe.ActiveCodePane; $sl = 0; $sc = 0; $el = 0; $ec = 0
        $cp.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
        return "`n    module $($cp.CodeModule.Parent.Name), line $sl`:`n    $($cp.CodeModule.Lines($sl, 1).Trim())"
    } catch { return '' }
}

# Compiles the project. Excel and PowerPoint accept the VBE's Debug > Compile
# command (id 578). Word refuses it over COM, so there we run one harmless
# parameterless procedure per module instead: VBA compiles a whole module
# before running anything in it, and a compile error raises the same dialog.
function Test-Compile($app, $vbproj, [uint32]$procId, $hostApp = $null, [string[]]$runProcs = @()) {
    $vbe = $vbproj.VBE
    $vbe.ActiveVBProject = $vbproj
    $vbe.MainWindow.Visible = $true
    $msg = $null
    if ($runProcs.Count -eq 0) {
        $menu = $vbe.CommandBars.Item('Menu Bar').Controls.Item('Debug')
        $ctl = $null
        for ($i = 1; $i -le $menu.Controls.Count; $i++) { $c = $menu.Controls.Item($i); if ($c.Id -eq 578) { $ctl = $c } }
        if (-not $ctl) { throw "Compile command not found in the VBE menu" }
        $msg = Invoke-Watched $procId { $ctl.Execute() }
    } else {
        foreach ($p in $runProcs) {
            $msg = Invoke-Watched $procId { [void]$hostApp.Run($p) }.GetNewClosure()
            if ($msg -ne $null -and $msg -match 'Compile error') { break }
            if ($msg -ne $null -and $msg -notmatch 'Compile error') { Write-Step "  ($p ran: $msg)"; $msg = $null }
        }
    }
    $vbe.MainWindow.Visible = $false
    if ($msg -ne $null) { throw "$app VBA project does not compile: $msg$(Get-CompileLocation $vbe)" }
    Write-Step "  $app project compiles cleanly"
}

# Adds customUI/customUI14.xml + relationship into an OPC package (xlam/ppam/dotm).
# Uses System.IO.Packaging (the OPC API Office itself uses). Rewriting the zip with
# System.IO.Compression produced files PowerPoint refused to load.
Add-Type -AssemblyName WindowsBase
function Add-Ribbon([string]$pkg, [string]$xmlPath) {
    $p = [System.IO.Packaging.Package]::Open($pkg, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite)
    try {
        $uri = New-Object Uri('/customUI/customUI14.xml', [UriKind]::Relative)
        if ($p.PartExists($uri)) { $p.DeletePart($uri) }
        $part = $p.CreatePart($uri, 'application/xml', [System.IO.Packaging.CompressionOption]::Normal)
        $s = $part.GetStream(); $bytes = [IO.File]::ReadAllBytes($xmlPath); $s.Write($bytes, 0, $bytes.Length); $s.Close()
        $relType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
        foreach ($r in @($p.GetRelationshipsByType($relType))) { $p.DeleteRelationship($r.Id) }
        [void]$p.CreateRelationship($uri, [System.IO.Packaging.TargetMode]::Internal, $relType, 'ltxRibbon')
        $p.Flush()
    } finally { $p.Close() }
    Write-Step "  ribbon injected into $(Split-Path $pkg -Leaf)"
}

function Remove-IfExists($p) { if (Test-Path $p) { Remove-Item $p -Force } }

