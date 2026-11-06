# =====================================================================
#  Lattice build script
#  Creates dist\Lattice.xlam, dist\Lattice.ppam and dist\Lattice.dotm from
#  the VBA sources in src\ and the ribbon XML in ribbon\.
#  Every project is compiled through the VBE before it is saved; the first
#  compile error (module + line) is printed and the build stops.
#  Run:  powershell -ExecutionPolicy Bypass -File build\build.ps1
# =====================================================================
param([switch]$NoWord, [switch]$NoExcel, [switch]$NoPpt)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Force $dist | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Write-Step($m) { Write-Host "[Lattice] $m" -ForegroundColor Cyan }

# The VBE writes the path of its MSForms extender cache (%TEMP%\VBE\MSForms.exd) into vbaProject.bin,
# so an add-in built under a user profile would ship that profile path. Office processes started by
# this script (COM activation inherits the environment too, measured) get a neutral TEMP instead.
$script:LtxOrigTemp = @($env:TEMP, $env:TMP)
$script:LtxTemp = 'C:\LatticeBuild\Temp'
New-Item -ItemType Directory -Force $script:LtxTemp | Out-Null
$env:TEMP = $script:LtxTemp; $env:TMP = $script:LtxTemp
function Restore-Temp { $env:TEMP = $script:LtxOrigTemp[0]; $env:TMP = $script:LtxOrigTemp[1] }
trap { Restore-Temp; break }

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
        $inProc = $false; $seenProc = $false
        foreach ($line in Get-Content $f) {
            $n++
            $code = ($line -split "'")[0]
            # module-level declarations (Dim/Const/Declare/Type/Enum/"Private x As") must precede the first procedure
            if ($code -match '^\s*(Public |Private |Friend |Static )?(Sub|Function|Property (Get|Let|Set))\s+\w+' -and $code -notmatch 'End (Sub|Function|Property)\s*$') { $inProc = $true; $seenProc = $true }
            elseif ($code -match '^\s*End (Sub|Function|Property)\b') { $inProc = $false }
            elseif (-not $inProc -and $seenProc -and $code -match '^\s*(Public |Private |Global )?(Dim |Const |Declare |Type |Enum |\w+ As )') { $bad += "$(Split-Path $f -Leaf):$n  module-level declaration after a procedure (VBA needs it at the top of the module)" }
            if ($code -match '^\s*(Public |Private )?(Sub|Function) \w+\(.*\)( As [\w.]+)?\s*:.*\bIf\b.*\bThen\b.*End (Sub|Function)\s*$') { $bad += "$(Split-Path $f -Leaf):$n  one-line procedure with If swallows End Sub" }
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
function Add-ListForm($vbproj, [string]$name, [string]$caption, [string]$listName, [int]$cols, [string]$colWidths, $buttons, [string]$codePath, [switch]$FormulaBox) {
    $comp = $vbproj.VBComponents.Add(3)          # vbext_ct_MSForm
    $comp.Name = $name
    $comp.Properties.Item('Caption').Value = $caption
    $comp.Properties.Item('Width').Value = 560
    $comp.Properties.Item('Height').Value = 340
    try { $comp.Properties.Item('StartUpPosition').Value = 1 } catch {}
    $d = $comp.Designer
    $ttl = $d.Controls.Add('Forms.Label.1', 'lblTitle', $true)
    $ttl.Left = 8; $ttl.Top = 6; $ttl.Width = 540; $ttl.Height = 22; $ttl.Caption = $caption
    $listTop = 32; $listH = 216
    if ($FormulaBox) {
        $fx = $d.Controls.Add('Forms.TextBox.1', 'txtFormula', $true)
        $fx.Left = 8; $fx.Top = 32; $fx.Width = 540; $fx.Height = 40; $fx.MultiLine = $true; $fx.Locked = $true
        $listTop = 78; $listH = 170
    }
    $lst = $d.Controls.Add('Forms.ListBox.1', $listName, $true)
    $lst.Left = 8; $lst.Top = $listTop; $lst.Width = 540; $lst.Height = $listH
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
# Settings dialogs are built from a spec in settings_forms.ps1 (one generic code-behind per app)
. (Join-Path $PSScriptRoot 'settings_forms.ps1')

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
. (Join-Path $PSScriptRoot 'scrub.ps1')     # Invoke-OfficeScrub: strips author and absPath stamps
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
        # custom icons: every image="name" in the XML becomes /customUI/images/name.png,
        # related from the customUI part with Id = name (how Office resolves the attribute)
        $xmlText = [IO.File]::ReadAllText($xmlPath)
        $names = [regex]::Matches($xmlText, ' image="([^"]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
        $imgRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
        $iconDir = Join-Path $root 'ribbon\icons'
        $n = 0
        foreach ($nm in $names) {
            $png = Join-Path $iconDir "$nm.png"
            if (-not (Test-Path $png)) { Write-Warning "icon missing: $png"; continue }
            $iu = New-Object Uri("/customUI/images/$nm.png", [UriKind]::Relative)
            if ($p.PartExists($iu)) { $p.DeletePart($iu) }
            $ip = $p.CreatePart($iu, 'image/png', [System.IO.Packaging.CompressionOption]::Normal)
            $s = $ip.GetStream(); $b = [IO.File]::ReadAllBytes($png); $s.Write($b, 0, $b.Length); $s.Close()
            [void]$part.CreateRelationship((New-Object Uri("images/$nm.png", [UriKind]::Relative)), [System.IO.Packaging.TargetMode]::Internal, $imgRel, $nm)
            $n++
        }
        if ($n -gt 0) { Write-Step "  $n icon(s) embedded" }
        $p.Flush()
    } finally { $p.Close() }
    Write-Step "  ribbon injected into $(Split-Path $pkg -Leaf)"
}

function Remove-IfExists($p) { if (Test-Path $p) { Remove-Item $p -Force } }

# ---------------------------------------------------------------- Excel
if (-not $NoExcel) {
Write-Step "Building Lattice.xlam"
Ensure-VBOM 'Excel'
$out = Join-Path $dist 'Lattice.xlam'; Remove-IfExists $out
$xl = New-Object -ComObject Excel.Application          # always a NEW instance, never the user's
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    Import-Modules $wb.VBProject (Get-ChildItem (Join-Path $root 'src\excel\*.bas') | ForEach-Object FullName)
    Add-ModelCheckForm $wb.VBProject (Join-Path $root 'src\excel\frmLtxModelCheck.vba')
    Add-ListForm $wb.VBProject 'frmLtxNav' 'Lattice - Navigate' 'lstRefs' 3 '150;90;290' @(@{n='btnGo';c='Go'}, @{n='btnClose';c='Close'}) (Join-Path $root 'src\excel\frmLtxNav.vba')
    Add-ListForm $wb.VBProject 'frmLtxTrace' 'Lattice - Trace' 'lstTrace' 3 '200;110;220' @(@{n='btnExpand';c='Expand >'}, @{n='btnCollapse';c='< Collapse'}, @{n='btnGo';c='Go'}, @{n='btnEdit';c='Edit (F2)'}, @{n='btnClose';c='Close'}) (Join-Path $root 'src\excel\frmLtxTrace.vba') -FormulaBox
    Add-SettingsForm $wb.VBProject 'frmLtxSettings' 'Lattice - Settings' $script:LtxExcelSettings (Join-Path $root 'src\excel\frmLtxSettings.vba')
    # Startup hook: bind keyboard shortcuts when the add-in loads, release them on close
    $tw = $wb.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
    $tw.AddFromString(@"
Private Sub Workbook_Open()
    On Error Resume Next
    LtxKeysLoad
End Sub
Private Sub Workbook_AddinInstall()
    On Error Resume Next
    LtxKeysLoad
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    LtxKeysUnload
End Sub
"@)
    Write-Step "  added ThisWorkbook startup hooks"
    $wb.VBProject.Name = 'Lattice'
    Test-Compile 'Excel' $wb.VBProject ([LtxWin]::PidOf([IntPtr]$xl.Hwnd))
    $wb.Title = 'Lattice'
    $wb.IsAddin = $true
    $wb.SaveAs($out, 55)      # xlOpenXMLAddIn
    $wb.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
Add-Ribbon $out (Join-Path $root 'ribbon\excel_customUI14.xml')
Write-Step (Invoke-OfficeScrub $out)
}

# ---------------------------------------------------------------- PowerPoint
if (-not $NoPpt) {
Write-Step "Building Lattice.ppam"
Ensure-VBOM 'PowerPoint'
$out = Join-Path $dist 'Lattice.ppam'; Remove-IfExists $out
$pp = New-Object -ComObject PowerPoint.Application     # PowerPoint is a singleton: attaches if running
$wasOpen = $pp.Presentations.Count
if ($wasOpen -gt 0) { Write-Warning "PowerPoint was already running: its TEMP could not be redirected, so run tests\privacy_scan.ps1 on the ppam" }
try { $ai = $pp.AddIns.Item('Lattice'); $ai.Loaded = 0; $pp.AddIns.Remove('Lattice') } catch {}
$pres = $pp.Presentations.Add(0)                        # msoFalse: no window
try {
    $files = @(Get-ChildItem (Join-Path $root 'src\shared\*.bas') | ForEach-Object FullName) +
             @(Get-ChildItem (Join-Path $root 'src\ppt\*.bas') | ForEach-Object FullName)
    Import-Modules $pres.VBProject $files
    Add-ExplorerForm $pres.VBProject (Join-Path $root 'src\ppt\frmLtxExplorer.vba')
    Add-SettingsForm $pres.VBProject 'frmLtxSettings' 'Lattice - Settings' $script:LtxPptSettings (Join-Path $root 'src\ppt\frmLtxSettings.vba')
    $pres.VBProject.Name = 'Lattice'
    Test-Compile 'PowerPoint' $pres.VBProject ([LtxWin]::PidOf([IntPtr]$pp.HWND))
    $pres.SaveAs($out, 30)    # ppSaveAsOpenXMLAddin
} finally {
    $pres.Saved = -1; $pres.Close()
    if ($wasOpen -eq 0) { $pp.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
}
Add-Ribbon $out (Join-Path $root 'ribbon\ppt_customUI14.xml')
Write-Step (Invoke-OfficeScrub $out)
}

# ---------------------------------------------------------------- Word
if (-not $NoWord) {
Write-Step "Building Lattice.dotm"
Ensure-VBOM 'Word'
$out = Join-Path $dist 'Lattice.dotm'; Remove-IfExists $out
# Word is started with /a (no startup templates). A third-party startup template
# (bloomberguiloader.dot) hooks the save event and hangs SaveAs2 in a hidden instance.
Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force
$wdExe = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe').'(default)'
$wdProc = Start-Process $wdExe -ArgumentList '/a', '/q' -PassThru -WindowStyle Minimized
$wd = $null
for ($i = 0; $i -lt 60 -and -not $wd; $i++) { Start-Sleep -Milliseconds 500; try { $wd = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') } catch {} }
if (-not $wd) { throw "Could not attach to Word started with /a" }
$wd.DisplayAlerts = -1   # alerts on, or the VBE swallows compile errors as E_FAIL
try {
    $doc = $wd.Documents.Add()
    $files = @(Get-ChildItem (Join-Path $root 'src\shared\*.bas') | ForEach-Object FullName) +
             @(Get-ChildItem (Join-Path $root 'src\word\*.bas') | ForEach-Object FullName)
    Import-Modules $doc.VBProject $files
    $doc.VBProject.Name = 'Lattice'
    Test-Compile 'Word' $doc.VBProject ([LtxWin]::PidOf([IntPtr]$wd.ActiveWindow.Hwnd)) $wd @('LtxDataDir', 'WdToggleBlack', 'WdToggleBlack', 'LtxWd_Ping')
    # SaveAs2 called from this process never returns after the VBE compile step, while the
    # same call from a fresh process completes in a second. So the save runs in a child job
    # that attaches to the same Word instance through the running-object table.
    $wd.DisplayAlerts = 0
    $saveJob = Start-Job -ArgumentList $out -ScriptBlock {
        param($path)
        $w = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
        $w.DisplayAlerts = 0
        $w.Documents.Item(1).SaveAs2($path, 15)    # wdFormatXMLTemplateMacroEnabled
        "saved"
    }
    if (-not (Wait-Job $saveJob -Timeout 90)) { Stop-Job $saveJob; throw "Word SaveAs2 timed out" }
    $saveResult = Receive-Job $saveJob; Remove-Job $saveJob -Force
    if (-not (Test-Path $out)) { throw "Word SaveAs2 did not produce $out ($saveResult)" }
    Write-Step "  saved $(Split-Path $out -Leaf) via child job"
    $doc.Close(0)
} finally { $wd.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wd) }
Add-Ribbon $out (Join-Path $root 'ribbon\word_customUI14.xml')
Write-Step (Invoke-OfficeScrub $out)
}

Restore-Temp
Write-Step "Done. Files in $dist"
Get-ChildItem $dist | Select-Object Name, Length, LastWriteTime
