# Sequential regression run: builds Word (Excel/PowerPoint are built separately), installs, then runs every
# COM-driven test one after another. Tests must not overlap: each one stops hidden Office instances.
# Usage: tests\run_all.ps1 -Book <TestModel.xlsx> -OutDir <folder>   (logs land in <OutDir>\logs)
param([string]$Book, [string]$OutDir, [switch]$SkipBuild)
$ErrorActionPreference = 'Continue'
$root = Split-Path $PSScriptRoot -Parent
$logs = Join-Path $OutDir 'logs'; New-Item -ItemType Directory -Force $logs | Out-Null
function Quiet { Get-Process EXCEL,POWERPNT,WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' -or $_.MainWindowTitle -match 'Presentation\d|TestModel|Book\d|Document\d|probe' } | Stop-Process -Force -ErrorAction SilentlyContinue; Start-Sleep 2 }
function Step($name, $script, $argList) {
    Quiet
    $t0 = Get-Date
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File $script @argList 2>&1 | ForEach-Object { "$_" }
    $out | Set-Content (Join-Path $logs "$name.log")
    $fails = @($out | Where-Object { $_ -match '\bFAIL\b|\bERROR\b|does not compile|Exception' }).Count
    $passes = @($out | Where-Object { $_ -match '\bPASS\b|\bok\b' }).Count
    "{0,-14} {1,5:N0}s  pass-lines={2,-4} fail-lines={3}" -f $name, ((Get-Date) - $t0).TotalSeconds, $passes, $fails
    if ($fails -gt 0) { $out | Where-Object { $_ -match '\bFAIL\b|\bERROR\b|does not compile|Exception' } | Select-Object -First 6 | ForEach-Object { "    $_" } }
}
if (-not $SkipBuild) {
    Step 'build-word' (Join-Path $root 'build\build.ps1') @('-NoExcel', '-NoPpt')
    Step 'install'    (Join-Path $root 'build\install.ps1') @()
}
Step 'word_test'   (Join-Path $root 'tests\word_test.ps1')   @('-Book', $Book, '-OutDir', $OutDir)
Step 'export_test' (Join-Path $root 'tests\export_test.ps1') @('-Book', $Book, '-OutDir', $OutDir)
Step 'ppt_sweep'   (Join-Path $root 'tests\ppt_sweep.ps1')   @('-Book', $Book, '-OutDir', $OutDir)
Step 'xl_sweep'    (Join-Path $root 'tests\xl_sweep.ps1')    @('-OutDir', $OutDir)
Step 'nav_test'    (Join-Path $root 'tests\nav_test.ps1')    @('-OutDir', $OutDir)
Step 'black_test'  (Join-Path $root 'tests\black_test.ps1')  @('-Book', $Book, '-OutDir', $OutDir)
Quiet
"all done"
