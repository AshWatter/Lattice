# Black Fonts switch: sheet toggle, workbook toggle, selection toggle, exact colour restore.
param([string]$Book, [string]$OutDir)
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force; Start-Sleep 1
Copy-Item $Book (Join-Path $OutDir 'BlackModel.xlsx') -Force
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$procId = [LtxWin]::PidOf([IntPtr]$xl.Hwnd)
if (-not (($xl.Workbooks | % Name) -contains 'Lattice.xlam')) { [void]$xl.Workbooks.Open("$env:APPDATA\Microsoft\AddIns\Lattice.xlam") }
function Settle() { Start-Sleep -Milliseconds 500; for ($k = 0; $k -lt 30; $k++) { try { if ($xl.Ready) { break } } catch {}; Start-Sleep -Milliseconds 300 } }
function RunM($name) { $m = Invoke-Watched $procId { [void]$xl.Run("Lattice.xlam!$name") }.GetNewClosure(); Settle; if ($m -eq $null) { "OK   $name" } elseif ($m -match 'Compile error|Run-time error|host error') { "FAIL $name : $($m -replace '\s+',' ')" } else { "OK   $name  [$($m -replace '\s+',' ')]" } }
$wb = $xl.Workbooks.Open((Join-Path $OutDir 'BlackModel.xlsx')); $ws = $wb.Worksheets.Item('Model'); $ws.Activate()
function Colors() { "B2=$($ws.Range('B2').Font.Color) B4=$($ws.Range('B4').Font.Color) B5=$($ws.Range('B5').Font.Color) A1(white)=$($ws.Range('A1').Font.Color) B3(black)=$($ws.Range('B3').Font.Color)" }
"before:      " + (Colors)
RunM 'LtxBlackToggle'
"sheet black: " + (Colors) + "   isOn=" + $xl.Run('Lattice.xlam!LtxBlackIsOn') + "  store sheet hidden=" + ($wb.Worksheets.Item('LatticeColors').Visible)
RunM 'LtxBlackToggle'
"restored:    " + (Colors) + "   isOn=" + $xl.Run('Lattice.xlam!LtxBlackIsOn') + "  store removed=" + (-not (($wb.Worksheets | % Name) -contains 'LatticeColors'))
$inp = $wb.Worksheets.Item('Inputs'); "inputs B1 before: $($inp.Range('B1').Font.Color)"
RunM 'LtxBlackToggleWorkbook'
"workbook on:  model B2=$($ws.Range('B2').Font.Color) inputs B1=$($inp.Range('B1').Font.Color)"
RunM 'LtxBlackToggleWorkbook'
"workbook off: model B2=$($ws.Range('B2').Font.Color) inputs B1=$($inp.Range('B1').Font.Color)"
$ws.Activate(); $ws.Range('B2:C2').Select(); RunM 'LtxBlackSelection'
"selection:   B2=$($ws.Range('B2').Font.Color) C2=$($ws.Range('C2').Font.Color) D2(untouched)=$($ws.Range('D2').Font.Color)"
$ws.Range('B2:C2').Select(); RunM 'LtxBlackSelection'
"sel restore: B2=$($ws.Range('B2').Font.Color) C2=$($ws.Range('C2').Font.Color)"
RunM 'LtxKeysLoad'; "  Ctrl+Shift+Z bound: " + ((Get-Content "$env:APPDATA\Lattice\excel_shortcuts.txt" -Raw) -match '\^\+z=LtxBlackToggle')
$wb.Close($false); $xl.Quit()
"done"
