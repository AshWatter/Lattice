# Trace In window: Excel must follow the selection in the window, across sheets (names with spaces too).
param([string]$Book, [string]$OutDir, [string]$Xlam = "")
if (-not $Xlam) { $Xlam = Join-Path $env:APPDATA "Microsoft/AddIns/Lattice.xlam" }
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force; Start-Sleep 1
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$wb = $xl.Workbooks.Open($Book)
if (-not (($xl.Workbooks | % Name) -contains 'Lattice.xlam')) { [void]$xl.Workbooks.Open("$env:APPDATA\Microsoft\AddIns\Lattice.xlam") }
$calc = $wb.Worksheets.Add(); $calc.Name = 'Calc Sheet'          # a space in the name on purpose
$calc.Range('A1').Formula = '=Model!B2*2'
$calc.Range('A2').Formula = "='Calc Sheet'!A1+1"
$calc.Activate(); $calc.Range('A2').Select() | Out-Null
$xl.WindowState = -4137
$xl.Run('Lattice.xlam!LtxTraceIn')                                  # modeless window; its first row is selected on open
Start-Sleep 2
$where = "$($xl.ActiveSheet.Name)!$($xl.ActiveCell.Address($false, $false))"
$ok1 = ($where -eq 'Calc Sheet!A1')
"after opening Trace In on 'Calc Sheet'!A2: Excel shows $where (expect Calc Sheet!A1) -> " + $(if ($ok1) { 'PASS' } else { 'FAIL' })
$shell = New-Object -ComObject WScript.Shell
$null = $shell.AppActivate('Lattice - Trace In'); Start-Sleep -Milliseconds 600
$shell.SendKeys('{RIGHT}'); Start-Sleep -Milliseconds 900          # expand Calc Sheet!A1 -> Model!B2
$shell.SendKeys('{DOWN}'); Start-Sleep -Milliseconds 900           # move onto Model!B2: Excel must follow
$where = "$($xl.ActiveSheet.Name)!$($xl.ActiveCell.Address($false, $false))"
$ok2 = ($where -eq 'Model!B2')
"after Right, Down in the window: Excel shows $where (expect Model!B2) -> " + $(if ($ok2) { 'PASS' } else { 'FAIL' })
$shell.SendKeys('{ESC}'); Start-Sleep -Milliseconds 500
"trace_nav_test: " + $(if ($ok1 -and $ok2) { '0 failure(s)' } else { 'FAILED' })
$wb.Close($false); $xl.Quit()
