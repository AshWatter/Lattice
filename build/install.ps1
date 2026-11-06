# =====================================================================
#  Lattice install script - registers the built add-ins for the current user
#  Run:  powershell -ExecutionPolicy Bypass -File build\install.ps1
#  Remove with:  -Uninstall
# =====================================================================
param([switch]$Uninstall)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$dist = Join-Path $root 'dist'
$addins = Join-Path $env:APPDATA 'Microsoft\AddIns'        # trusted location for Excel + PowerPoint
$wordStartup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
New-Item -ItemType Directory -Force $addins, $wordStartup | Out-Null
function Write-Step($m) { Write-Host "[Lattice] $m" -ForegroundColor Cyan }

$running = @('EXCEL','POWERPNT','WINWORD') | Where-Object { Get-Process $_ -ErrorAction SilentlyContinue }
if ($running) { Write-Warning "Close $($running -join ', ') first so the add-in files are not locked." }

# ---- Excel: OPEN / OPEN1 / ... values under Excel\Options ----
$xlOpts = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$xlTarget = Join-Path $addins 'Lattice.xlam'
function Get-ExcelOpenValues { (Get-Item $xlOpts).Property | Where-Object { $_ -match '^OPEN\d*$' } }

if ($Uninstall) {
    foreach ($v in Get-ExcelOpenValues) {
        if ((Get-ItemProperty $xlOpts -Name $v).$v -match 'Lattice\.xlam') { Remove-ItemProperty $xlOpts -Name $v; Write-Step "Excel: removed $v" }
    }
    Remove-Item (Join-Path $addins 'Lattice.xlam') -Force -ErrorAction SilentlyContinue
    Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\AddIns\Lattice' -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item (Join-Path $addins 'Lattice.ppam') -Force -ErrorAction SilentlyContinue
    Remove-Item (Join-Path $wordStartup 'Lattice.dotm') -Force -ErrorAction SilentlyContinue
    Write-Step "Uninstalled."
    return
}

# Excel
Copy-Item (Join-Path $dist 'Lattice.xlam') $xlTarget -Force
if (-not (Test-Path $xlOpts)) { New-Item $xlOpts -Force | Out-Null }
$already = Get-ExcelOpenValues | Where-Object { (Get-ItemProperty $xlOpts -Name $_).$_ -match 'Lattice\.xlam' }
if (-not $already) {
    $i = 0; $name = 'OPEN'
    while ((Get-ItemProperty $xlOpts -Name $name -ErrorAction SilentlyContinue)) { $i++; $name = "OPEN$i" }
    Set-ItemProperty $xlOpts -Name $name -Value "`"$xlTarget`""
    Write-Step "Excel: registered as $name"
} else { Write-Step "Excel: already registered" }

# PowerPoint: let PowerPoint register the add-in itself (writes 16.0\PowerPoint\AddIns\Lattice)
$ppTarget = Join-Path $addins 'Lattice.ppam'
$pp = New-Object -ComObject PowerPoint.Application
try { $old = $pp.AddIns.Item('Lattice'); $old.Loaded = 0; $pp.AddIns.Remove('Lattice') } catch {}
Copy-Item (Join-Path $dist 'Lattice.ppam') $ppTarget -Force
$ai = $pp.AddIns.Add($ppTarget)
$ai.AutoLoad = -1
$ai.Loaded = -1
Write-Step "PowerPoint: registered (loaded=$($ai.Loaded), autoload=$($ai.AutoLoad))"
if ($pp.Presentations.Count -eq 0) { $pp.Quit() }
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)

# Word
if (Test-Path (Join-Path $dist 'Lattice.dotm')) {
    Copy-Item (Join-Path $dist 'Lattice.dotm') (Join-Path $wordStartup 'Lattice.dotm') -Force
    Write-Step "Word: copied to STARTUP"
} else { Write-Warning 'Lattice.dotm not built yet - Word skipped' }

Write-Step "Installed. Restart Excel, PowerPoint and Word; look for the 'Lattice' tab."
