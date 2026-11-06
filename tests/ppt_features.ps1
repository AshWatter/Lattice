# PowerPoint feature test: text links, styles, pagination, agenda, deck check, turbo shapes,
# arrange grid, convert to picture. Dialogs (InputBox / MsgBox) are auto-accepted by the watcher.
param([string]$Book, [string]$OutDir, [string]$Lib)
$ErrorActionPreference = 'Continue'
. $Lib
function RunW($pp, $procId, $name, $arg) {
    $m = Invoke-Watched $procId { if ($null -ne $arg) { $pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @("Lattice.ppam!$name", $arg)) } else { $pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @("Lattice.ppam!$name")) } }.GetNewClosure()
    if ($m -eq $null) { "OK   $name" } else { "RAN  $name : $($m -replace '\s+',' ')" }
}
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$wb = $xl.Workbooks.Open($Book); $ws = $wb.Worksheets.Item('Model'); $ws.Activate()
$pp = New-Object -ComObject PowerPoint.Application; $pp.Visible = -1
$procId = [LtxWin]::PidOf([IntPtr]$pp.HWND)
$pres = $pp.Presentations.Add(-1)
$lay = $pres.SlideMaster.CustomLayouts.Item(2)
1..4 | ForEach-Object { [void]$pres.Slides.AddSlide($_, $lay) }
# sections
$pres.SectionProperties.AddBeforeSlide(1, 'Overview') | Out-Null
$pres.SectionProperties.AddBeforeSlide(3, 'Financials') | Out-Null
for ($i = 1; $i -le 4; $i++) { $pres.Slides.Item($i).Shapes.Title.TextFrame.TextRange.Text = "Slide $i title" }

# ---- text link: whole shape ----
$pp.ActiveWindow.View.GotoSlide(1)
$s1 = $pres.Slides.Item(1)
$tb = $s1.Shapes.AddTextbox(1, 50, 150, 300, 30); $tb.TextFrame.TextRange.Text = 'placeholder'
$ws.Range('B2').Select()           # value 999 (from previous test) or 100
$tb.Select()
RunW $pp $procId 'LtxLinkTextWhole' $null
"  whole-text link now reads: '" + $tb.TextFrame.TextRange.Text + "'  tags type=" + $tb.Tags.Item('LTX_TYPE') + " template=" + $tb.Tags.Item('LTX_TEMPLATE')
# ---- text link: substring inside a sentence ----
$tb2 = $s1.Shapes.AddTextbox(1, 50, 200, 400, 30); $tb2.TextFrame.TextRange.Text = 'Revenue of XX in 2024A'
$ws.Range('B2').Select()
$tb2.TextFrame.TextRange.Characters(12, 2).Select()
RunW $pp $procId 'LtxLinkTextSelection' $null
"  substring link now reads: '" + $tb2.TextFrame.TextRange.Text + "'  template=" + $tb2.Tags.Item('LTX_TEMPLATE')
$ws.Range('B1').Select()
$tb2.TextFrame.TextRange.Characters(($tb2.TextFrame.TextRange.Text.Length - 4), 5).Select()
RunW $pp $procId 'LtxLinkTextSelection' $null
"  second placeholder: '" + $tb2.TextFrame.TextRange.Text + "'  template=" + $tb2.Tags.Item('LTX_TEMPLATE')
# change Excel and refresh all
$ws.Range('B2').Value2 = 123.4
RunW $pp $procId 'LtxRefreshAll' $null
"  after refresh: whole='" + $tb.TextFrame.TextRange.Text + "'  sentence='" + $tb2.TextFrame.TextRange.Text + "'"

# ---- styles ----
$box = $s1.Shapes.AddShape(5, 50, 260, 150, 50); $box.Fill.ForeColor.RGB = 0x5A1F00; $box.Line.Visible = 0
$box.TextFrame.TextRange.Text = 'Styled'; $box.TextFrame.TextRange.Font.Size = 14; $box.TextFrame.TextRange.Font.Bold = -1
$box.Select()
RunW $pp $procId 'LtxStyleSave' $null          # InputBox default name accepted -> "Body box"
$box2 = $pres.Slides.Item(2).Shapes.AddShape(5, 50, 260, 150, 50); $box2.TextFrame.TextRange.Text = 'Plain'
$pp.ActiveWindow.View.GotoSlide(2); $box2.Select()
RunW $pp $procId 'LtxStyleApply' $null         # InputBox default = last style
"  box2 fill after style: " + $box2.Fill.ForeColor.RGB + " (expected " + $box.Fill.ForeColor.RGB + ")  bold=" + $box2.TextFrame.TextRange.Font.Bold + " tag=" + $box2.Tags.Item('LTX_STYLE')

# ---- turbo shapes + arrange grid + convert to picture ----
$pp.ActiveWindow.View.GotoSlide(3)
$s3 = $pres.Slides.Item(3)
$before = $s3.Shapes.Count
foreach ($k in 'harvey50','chevrons','sticky','timeline') { RunW $pp $procId 'LtxTurbo' $k | Out-Null }
"  turbo shapes added on slide 3: " + ($s3.Shapes.Count - $before)
$names = @(); foreach ($i in 1..4) { $sh = $s3.Shapes.AddShape(1, 10 * $i, 400, 40, 40); $sh.Name = "grid$i"; $names += $sh.Name }
$s3.Shapes.Range([string[]]$names).Select()
RunW $pp $procId 'LtxArrangeGrid' $null        # InputBox default columns accepted
"  grid positions: " + (($names | % { $sh = $s3.Shapes.Item($_); "(" + [int]$sh.Left + "," + [int]$sh.Top + ")" }) -join ' ')
$s3.Shapes.Item('grid1').Select()
RunW $pp $procId 'LtxConvertToPicture' 'emf'
"  grid1 now type: " + $s3.Shapes.Item($s3.Shapes.Count).Type + " (13 = picture)"

# ---- pagination + agenda + deck check ----
RunW $pp $procId 'LtxPageNumbers' $null        # InputBoxes accept defaults {n} and 1
"  page number boxes: " + (($pres.Slides | % { if ($_.Shapes | ? { $_.Name -eq 'Lattice page number' }) { 1 } }) | Measure-Object).Count
RunW $pp $procId 'LtxBlankPages' 'sections'
"  slides after blank pages: " + $pres.Slides.Count + "  blank-tagged: " + (($pres.Slides | ? { $_.Tags.Item('LTX_BLANK') -ne '' }) | Measure-Object).Count
RunW $pp $procId 'LtxBlankPagesRemove' $null
RunW $pp $procId 'LtxAgenda' $null
"  slides after agenda: " + $pres.Slides.Count + "  agenda-tagged: " + (($pres.Slides | ? { $_.Tags.Item('LTX_AGENDA') -ne '' }) | Measure-Object).Count
RunW $pp $procId 'LtxDeckCheck' $null
$pres.SaveAs((Join-Path $OutDir 'features_deck.pptx'))
$pres.Slides.Item(2).Export((Join-Path $OutDir 'features_slide_links.png'), 'PNG', 1600, 900)
$pres.Close(); $pp.Quit(); $wb.Close($false); $xl.Quit()
"done"
