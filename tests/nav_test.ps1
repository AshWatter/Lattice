# Keyboard navigation: Ctrl+[ / Ctrl+] / back, across sheets, plus the picker for multiple refs.
param([string]$OutDir)
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force; Start-Sleep 1
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$procId = [LtxWin]::PidOf([IntPtr]$xl.Hwnd)
if (-not (($xl.Workbooks | % Name) -contains 'Lattice.xlam')) { [void]$xl.Workbooks.Open("$env:APPDATA\Microsoft\AddIns\Lattice.xlam") }
function Settle() { Start-Sleep -Milliseconds 500; for ($k = 0; $k -lt 30; $k++) { try { if ($xl.Ready) { break } } catch {}; Start-Sleep -Milliseconds 300 } }
function RunM($name) { $m = Invoke-Watched $procId { [void]$xl.Run("Lattice.xlam!$name") }.GetNewClosure(); Settle; if ($m -eq $null) { "OK   $name" } elseif ($m -match 'Compile error|Run-time error|host error') { "FAIL $name : $($m -replace '\s+',' ')" } else { "OK   $name  [$($m -replace '\s+',' ')]" } }
function Here() { "$($xl.ActiveSheet.Name)!$($xl.ActiveCell.Address($false,$false))" }
$wb = $xl.Workbooks.Add(); $wb.SaveAs((Join-Path $OutDir 'nav_test.xlsx'), 51)
$a = $wb.Worksheets.Item(1); $a.Name = 'Inputs'; $b = $wb.Worksheets.Add(); $b.Name = 'Calc'
$a.Range('B2').Value2 = 100; $a.Range('B3').Value2 = 0.25
$b.Range('C5').Formula = '=Inputs!B2*(1-Inputs!B3)'     # two precedents, other sheet
$b.Range('C6').Formula = '=C5*2'                         # one precedent
$b.Range('C7').Formula = '=C6+1'
"--- single precedent: C6 -> C5 ---"
$b.Activate(); $b.Range('C6').Select(); RunM 'LtxGoPrecedent'; "  now at: $(Here) (expect Calc!C5)"
"--- back ---"
RunM 'LtxGoBack'; "  now at: $(Here) (expect Calc!C6)"
"--- single dependent: C6 -> C7 ---"
RunM 'LtxGoDependent'; "  now at: $(Here) (expect Calc!C7)"
"--- cross-sheet single precedent: Inputs!B2 -> dependents on Calc ---"
$a.Activate(); $a.Range('B2').Select(); RunM 'LtxGoDependent'; "  now at: $(Here) (expect Calc!C5)"
"--- multiple precedents: Calc!C5 -> picker ---"
RunM 'LtxGoPrecedent'; "  now at: $(Here) (unchanged, picker shown)   items: " + $(try { $v = $xl.Run('Lattice.xlam!LtxCollectRefs', $b.Range('C5'), $true); $v.Count } catch { "n/a ($($_.Exception.Message))" })
"--- shortcut file has nav keys ---"
RunM 'LtxKeysLoad'
$keys = Get-Content "$env:APPDATA\Lattice\excel_shortcuts.txt" -Raw
"  ^{[}=" + ($keys -match '\^\{\[\}=LtxGoPrecedent') + "  ^{]}=" + ($keys -match '\^\{\]\}=LtxGoDependent') + "  back=" + ($keys -match '\^\+~=LtxGoBack') + "  old broken '^+[=' gone=" + (-not ($keys -match '\^\+\[='))
$wb.Save(); $xl.DisplayAlerts = $false; $xl.Quit()
"done"
