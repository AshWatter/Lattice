# Reproduces the "wrong Excel" failure: a hidden automation Excel (registered first) with another
# workbook selected, plus the user's visible Excel with the model. PowerPoint import and the
# Excel-side export must both use the visible instance. Output: PASS/FAIL lines.
param([string]$Book, [string]$OutDir)
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Get-Process EXCEL,POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' -or $_.MainWindowTitle -match '^(Book\d+|TestModel|Presentation\d+)' } | Stop-Process -Force; Start-Sleep 2
$fails = 0
function Check($name, $ok, $detail) { if ($ok) { "PASS $name" } else { $script:fails++; "FAIL $name  $detail" } }

# 1. stray hidden instance, registered first, with its own selection
$hid = New-Object -ComObject Excel.Application; $hid.Visible = $false
$hb = $hid.Workbooks.Add(); $hb.Worksheets.Item(1).Range('A1').Value2 = 'STRAY'; $hb.Worksheets.Item(1).Range('A1:B2').Select() | Out-Null
Start-Sleep 1
# 2. the user's visible instance with the model
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$wb = $xl.Workbooks.Open($Book); $ws = $wb.Worksheets.Item('Model'); $ws.Activate(); $ws.Range('A1:D5').Select() | Out-Null
Start-Sleep 1
$pp = New-Object -ComObject PowerPoint.Application; $pp.Visible = -1
$pres = $pp.Presentations.Add(-1); [void]$pres.Slides.Add(1, 12); $sld = $pres.Slides.Item(1)
$procId = [LtxWin]::PidOf([IntPtr]$pp.HWND)
function RunPpt($macro, $arg) { if ($null -ne $arg) { $script:pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $script:pp, @($macro, $arg)) } else { $script:pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $script:pp, @($macro)) } }

# PowerPoint-side import of the Excel selection (picture)
$m = Invoke-Watched $procId { RunPpt 'Lattice.ppam!LtxImport' 'picture' }
Start-Sleep 1
$shp = $null; if ($sld.Shapes.Count -ge 1) { $shp = $sld.Shapes.Item($sld.Shapes.Count) }
Check 'import picture created a shape' ($sld.Shapes.Count -eq 1) "shapes=$($sld.Shapes.Count) dialog=$m"
if ($shp) {
    Check 'picture link points at the visible workbook' ($shp.Tags.Item('LTX_WB') -ieq $Book) "LTX_WB=$($shp.Tags.Item('LTX_WB'))"
    Check 'picture link range is the visible selection' ($shp.Tags.Item('LTX_RANGE') -eq 'A1:D5') "range=$($shp.Tags.Item('LTX_RANGE'))"
}
# PowerPoint-side import as editable table
$ws.Range('A1:D5').Select() | Out-Null
$m = Invoke-Watched $procId { RunPpt 'Lattice.ppam!LtxImport' 'table' }
Start-Sleep 1
$tbl = $null; foreach ($s in $sld.Shapes) { if ($s.HasTable -eq -1) { $tbl = $s } }
Check 'import table created a table' ($tbl -ne $null) "shapes=$($sld.Shapes.Count) dialog=$m"
if ($tbl) { Check 'table text comes from the visible workbook' ($tbl.Table.Cell(1,1).Shape.TextFrame.TextRange.Text -eq $ws.Range('A1').Text) "got '$($tbl.Table.Cell(1,1).Shape.TextFrame.TextRange.Text)' want '$($ws.Range('A1').Text)'" }
# refresh everything with the stray still alive
$m = Invoke-Watched $procId { RunPpt 'Lattice.ppam!LtxRefreshAll' }
$log = Get-Content "$env:APPDATA\Lattice\lattice.log" -Tail 12
Check 'refresh all did not fail' (-not ($log -match 'Refresh failed')) ($log -join ' | ')
# Excel-side export (Ctrl+Alt+R path) from the visible instance
if (-not (($xl.Workbooks | % Name) -contains 'Lattice.xlam')) { [void]$xl.Workbooks.Open("$env:APPDATA\Microsoft\AddIns\Lattice.xlam") }
$ws.Activate(); $ws.Range('A1:C3').Select() | Out-Null
$before = $sld.Shapes.Count
$pp.ActiveWindow.View.GotoSlide(1)
$m = Invoke-Watched ([LtxWin]::PidOf([IntPtr]$xl.Hwnd)) { $script:xl.Run('Lattice.xlam!LtxSendToPpt', 'picture') }
Start-Sleep 2
$new = $null; if ($sld.Shapes.Count -gt $before) { $new = $sld.Shapes.Item($sld.Shapes.Count) }
Check 'export from Excel added a shape' ($new -ne $null) "before=$before after=$($sld.Shapes.Count) dialog=$m"
if ($new) { Check 'exported link points at the visible workbook' ($new.Tags.Item('LTX_WB') -ieq $Book -and $new.Tags.Item('LTX_RANGE') -eq 'A1:C3') "LTX_WB=$($new.Tags.Item('LTX_WB')) range=$($new.Tags.Item('LTX_RANGE'))" }
Check 'stray workbook untouched' ($hb.Worksheets.Item(1).Range('A1').Value2 -eq 'STRAY') ''
Check 'model fonts restored' ($ws.Range('B2').Font.Color -eq 16711680 -or $ws.Range('B2').Font.Color -eq 0) "B2 colour=$($ws.Range('B2').Font.Color)"

$pres.Saved = -1; $pres.Close(); $pp.Quit()
$wb.Close($false); $xl.Quit(); $hb.Close($false); $hid.Quit()
"instance_test: $fails failure(s)"
