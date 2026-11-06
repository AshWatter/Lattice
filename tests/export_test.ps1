# Export from Excel with Match Size over a selected shape, V1/V2 version-aware refresh,
# keytip cards, shortcut bindings, Model Check pane.
param([string]$Book, [string]$OutDir)
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Get-Process EXCEL,POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force; Start-Sleep 1
$v1 = Join-Path $OutDir 'Model_v1.xlsx'; $v2 = Join-Path $OutDir 'Model_v2.xlsx'
Copy-Item $Book $v1 -Force
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$xlPid = [LtxWin]::PidOf([IntPtr]$xl.Hwnd)
if (-not (($xl.Workbooks | % Name) -contains 'Lattice.xlam')) { [void]$xl.Workbooks.Open("$env:APPDATA\Microsoft\AddIns\Lattice.xlam") }
function Settle() { Start-Sleep -Milliseconds 700; for ($k = 0; $k -lt 30; $k++) { try { if ($xl.Ready) { break } } catch {}; Start-Sleep -Milliseconds 300 } }
function RunX($name, $arg) { $m = Invoke-Watched $xlPid { if ($null -ne $arg) { [void]$xl.Run("Lattice.xlam!$name", $arg) } else { [void]$xl.Run("Lattice.xlam!$name") } }.GetNewClosure(); Settle; if ($m -eq $null) { "OK   $name $arg" } elseif ($m -match 'Compile error|Run-time error|host error') { "FAIL $name : $($m -replace '\s+',' ')" } else { "OK   $name $arg  [$($m -replace '\s+',' ')]" } }
$wb = $xl.Workbooks.Open($v1); $ws = $wb.Worksheets.Item('Model')
$pp = New-Object -ComObject PowerPoint.Application; $pp.Visible = -1
$ppPid = [LtxWin]::PidOf([IntPtr]$pp.HWND)
function RunP($name, $arg) { $m = Invoke-Watched $ppPid { if ($null -ne $arg) { $pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @("Lattice.ppam!$name", $arg)) } else { $pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @("Lattice.ppam!$name")) } }.GetNewClosure(); Settle; if ($m -eq $null) { "OK   $name $arg" } elseif ($m -match 'Compile error|Run-time error|host error') { "FAIL $name : $($m -replace '\s+',' ')" } else { "OK   $name $arg  [$($m -replace '\s+',' ')]" } }
$pres = $pp.Presentations.Add(-1); [void]$pres.Slides.Add(1, 12); $pp.ActiveWindow.View.GotoSlide(1)
$s = $pres.Slides.Item(1)
"--- export to PowerPoint, match width of selected placeholder shape ---"
$ph = $s.Shapes.AddShape(1, 100, 100, 300, 150); $ph.Name = 'Placeholder'; $ph.Select()
RunX 'LtxSetExportSizeMode' 'width'
$ws.Activate(); $ws.Range('A1:D5').Select()
RunX 'LtxSendToPpt' 'picture'
$sh = $s.Shapes.Item($s.Shapes.Count)
"  shapes on slide: $($s.Shapes.Count) (expect 1)  name=$($sh.Name) type=$($sh.Tags.Item('LTX_TYPE')) left=$([int]$sh.Left) top=$([int]$sh.Top) width=$([int]$sh.Width) (expect 100,100,300)  sizemode=$($sh.Tags.Item('LTX_SIZEMODE'))"
"--- export as table over the picture, match size ---"
$sh.Select(); RunX 'LtxSetExportSizeMode' 'size'
$ws.Range('A1:D5').Select(); RunX 'LtxSendToPpt' 'table'
$sh = $s.Shapes.Item($s.Shapes.Count)
"  shapes: $($s.Shapes.Count)  hasTable=$($sh.HasTable) width=$([int]$sh.Width) height=$([int]$sh.Height) (expect 300x150)"
"--- V1 -> V2 version prompt on refresh ---"
$ws.Range('B2').Value2 = 100
$wb.SaveCopyAs($v2)
$wb.Close($false)
$wb2 = $xl.Workbooks.Open($v2); $wb2.Worksheets.Item('Model').Range('B2').Value2 = 777; $wb2.Save()
RunP 'LtxRefreshAll' $null
$sh = $s.Shapes.Item($s.Shapes.Count)
"  link now points to: $(Split-Path $sh.Tags.Item('LTX_WB') -Leaf) (expect Model_v2.xlsx)   cell(2,2)='$($sh.Table.Cell(2,2).Shape.TextFrame.TextRange.Text)' (expect 777.0)"
"--- highlight links toggle ---"
RunP 'LtxHighlightLinks' $null; "  highlight shapes: " + (($s.Shapes | ? { $_.Name -like 'Lattice highlight*' }) | Measure-Object).Count + " view=" + $pp.ActiveWindow.ViewType
RunP 'LtxHighlightLinks' $null; "  after toggle off: " + (($s.Shapes | ? { $_.Name -like 'Lattice highlight*' }) | Measure-Object).Count
"--- cards and shortcuts ---"
RunP 'LtxKeytipCard' $null; "  ppt card: " + (Test-Path "$env:APPDATA\Lattice\lattice_powerpoint_keytips.html")
RunX 'LtxShortcutCard'; "  excel card: " + (Test-Path "$env:APPDATA\Lattice\lattice_excel_shortcuts.html")
RunX 'LtxKeysLoad'; "  Ctrl+Alt+R bound in file: " + ((Get-Content "$env:APPDATA\Lattice\excel_shortcuts.txt" -Raw) -match '\^%r=LtxSendToPptPicture')
Get-Process msedge,chrome -ErrorAction SilentlyContinue | Out-Null
"--- model check pane (collect issues) ---"
$wb2.Worksheets.Item('Model').Activate()
$issues = $xl.Run('Lattice.xlam!LtxCollectIssues', $false)
"  issues on Model sheet: " + $(if ($issues -is [array]) { ($issues.GetLength(0)) } else { 'none' })
RunX 'LtxModelCheckPane'; "  pane shown (modeless)"
"--- native chart and embedded workbook exports ---"
$ws2 = $wb2.Worksheets.Item('Model')
$ch = $ws2.ChartObjects().Add(300, 10, 200, 120); $ch.Chart.SetSourceData($ws2.Range('A1:D4')); $ch.Chart.ChartType = 51; $ch.Name = 'ExpChart'; $ch.Activate()
[void]$pres.Slides.Add(2, 12); $pp.ActiveWindow.View.GotoSlide(2); $s2 = $pres.Slides.Item(2)
RunX 'LtxSetExportSizeMode' 'none'
RunX 'LtxSendToPpt' 'nativechart'
$nc = $s2.Shapes.Item($s2.Shapes.Count); "  native chart: hasChart=$($nc.HasChart) type=$($nc.Tags.Item('LTX_TYPE'))"
$ws2.Range('A1:D5').Select(); RunX 'LtxSendToPpt' 'embed'
$em = $s2.Shapes.Item($s2.Shapes.Count); "  embedded: shapeType=$($em.Type) (7=OLE) tag=$($em.Tags.Item('LTX_TYPE'))"
$ws2.Range('B3').Value2 = 321; RunP 'LtxRefreshSlide' $null
"  after refresh shapes on slide 2: $($s2.Shapes.Count) (expect 2)  native chart still chart: $($s2.Shapes.Item(1).HasChart)"
"--- formula error warning on export ---"
$ws2.Range('D5').Formula = '=1/0'; $ws2.Range('A1:D5').Select(); RunX 'LtxSendToPpt' 'picture'
Start-Sleep 2; for ($k = 0; $k -lt 20; $k++) { try { $null = $pp.Visible; break } catch { Start-Sleep -Milliseconds 500 } }
$pres.SaveAs((Join-Path $OutDir 'export_test.pptx')); $pres.Close(); $pp.Quit(); $wb2.Close($false); $xl.Quit()
"done"
