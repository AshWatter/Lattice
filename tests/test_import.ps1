# End-to-end test of Lattice import + refresh, driven through PowerPoint COM.
# Steps: open test workbook in a visible Excel, write the pending descriptor the way the Excel
# add-in does, start PowerPoint (loads Lattice.ppam), run LtxImport, inspect tags, export slide
# PNG, change a value in Excel, run LtxRefreshSelected, export again, compare.
param([string]$Book, [string]$OutDir)
$ErrorActionPreference = 'Stop'
function Run($pp, $name) { $pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @("Lattice.ppam!$name")) }
function RunArg($pp, $name, $arg) { $pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @("Lattice.ppam!$name", $arg)) }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$wb = $xl.Workbooks.Open($Book)
$ws = $wb.Worksheets.Item('Model')
$ws.Activate(); $ws.Range('A1:D5').Select()

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = -1
"PowerPoint add-ins: " + (($pp.AddIns | ForEach-Object { "$($_.Name) loaded=$($_.Loaded)" }) -join '; ')
$pres = $pp.Presentations.Add(-1)
[void]$pres.Slides.Add(1, 12)    # blank
$pp.ActiveWindow.View.GotoSlide(1)

# 1) Import via the running-Excel selection path
"--- import (picture, black fonts) ---"
RunArg $pp 'LtxImport' 'picture'
$sld = $pres.Slides.Item(1)
"shapes on slide: " + $sld.Shapes.Count
$shp = $sld.Shapes.Item($sld.Shapes.Count)
"shape: $($shp.Name) type=$($shp.Type) w=$([math]::Round($shp.Width,1)) h=$([math]::Round($shp.Height,1)) left=$([math]::Round($shp.Left,1)) top=$([math]::Round($shp.Top,1))"
foreach ($k in 'LTX_TYPE','LTX_WB','LTX_SHEET','LTX_RANGE','LTX_NAME','LTX_BLACK','LTX_SIZEMODE','LTX_UPDATED') { "  $k = " + $shp.Tags.Item($k) }
$sld.Export((Join-Path $OutDir 'after_import.png'), 'PNG', 1600, 900)

# source colours must be untouched
"source font colours after import: B2=$($ws.Range('B2').Font.Color) B3=$($ws.Range('B3').Font.Color) B4=$($ws.Range('B4').Font.Color) B5=$($ws.Range('B5').Font.Color) A1=$($ws.Range('A1').Font.Color)  saved=$($wb.Saved)"

# 2) Change a value and refresh
"--- refresh ---"
$ws.Range('B2').Value2 = 999
$shp.Select()
Run $pp 'LtxRefreshSelected'
$shp2 = $sld.Shapes.Item($sld.Shapes.Count)
"after refresh: shapes=$($sld.Shapes.Count) name=$($shp2.Name) w=$([math]::Round($shp2.Width,1)) h=$([math]::Round($shp2.Height,1)) left=$([math]::Round($shp2.Left,1)) top=$([math]::Round($shp2.Top,1)) updated=$($shp2.Tags.Item('LTX_UPDATED'))"
$sld.Export((Join-Path $OutDir 'after_refresh.png'), 'PNG', 1600, 900)

# 3) Table import
"--- import (table) ---"
$ws.Range('A1:D5').Select()
RunArg $pp 'LtxImport' 'table'
$t = $sld.Shapes.Item($sld.Shapes.Count)
"table shape: $($t.Name) hasTable=$($t.HasTable) rows=$($t.Table.Rows.Count) cols=$($t.Table.Columns.Count) fontColor(2,2)=$($t.Table.Cell(2,2).Shape.TextFrame.TextRange.Font.Color.RGB) text(2,2)='$($t.Table.Cell(2,2).Shape.TextFrame.TextRange.Text)'"
$t.Top = 320
$sld.Export((Join-Path $OutDir 'after_table.png'), 'PNG', 1600, 900)

# 4) Refresh whole deck with workbook CLOSED (tests reopen path)
"--- refresh all with workbook closed ---"
$wb.Close($false)
Run $pp 'LtxRefreshAll'
"workbooks open in Excel after refresh: " + $xl.Workbooks.Count
"link table rows: " + ((Run $pp 'LtxLinkTable') | Measure-Object).Count

$pres.SaveAs((Join-Path $OutDir 'test_deck.pptx'))
$pres.Close(); $pp.Quit(); $xl.Quit()
"done"
