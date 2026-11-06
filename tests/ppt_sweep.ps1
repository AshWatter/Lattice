# Sweep of the PowerPoint macros not covered by test_import / ppt_features.
# Each macro runs under the dialog watcher; prompts are auto-accepted with defaults.
param([string]$Book, [string]$OutDir)
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
function RunW($name, $arg) {
    $m = Invoke-Watched $script:procId { if ($null -ne $arg) { $pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @("Lattice.ppam!$name", $arg)) } else { $pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @("Lattice.ppam!$name")) } }.GetNewClosure()
    if ($m -eq $null) { "OK   $name $arg" } elseif ($m -match 'Compile error|Run-time error|host error') { "FAIL $name $arg : $($m -replace '\s+',' ')" } else { "OK   $name $arg  [$($m -replace '\s+',' ')]" }
}
function RunW2($name, $a1, $a2) { $m = Invoke-Watched $script:procId { $pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $pp, @("Lattice.ppam!$name", $a1, $a2)) }.GetNewClosure(); if ($m -eq $null) { "OK   $name $a1 $a2" } elseif ($m -match 'Compile error|Run-time error|host error') { "FAIL $name : $($m -replace '\s+',' ')" } else { "OK   $name $a1 $a2  [$($m -replace '\s+',' ')]" } }
Get-Process EXCEL,POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq "" } | Stop-Process -Force; Start-Sleep 1
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$wb = $xl.Workbooks.Open($Book); $ws = $wb.Worksheets.Item('Model'); $ws.Activate()
$pp = New-Object -ComObject PowerPoint.Application; $pp.Visible = -1
$script:procId = [LtxWin]::PidOf([IntPtr]$pp.HWND)
$pres = $pp.Presentations.Add(-1)
$lay = $pres.SlideMaster.CustomLayouts.Item(2)
1..3 | ForEach-Object { [void]$pres.Slides.AddSlide($_, $lay) }
$pres.SectionProperties.AddBeforeSlide(1, 'Part A') | Out-Null
$pres.SectionProperties.AddBeforeSlide(3, 'Part B') | Out-Null
$pres.SaveAs((Join-Path $OutDir 'sweep_deck.pptx'))
$pp.ActiveWindow.View.GotoSlide(1)
$s = $pres.Slides.Item(1)
$a = $s.Shapes.AddShape(1, 50, 120, 120, 60); $a.Name = 'A'; $a.TextFrame.TextRange.Text = 'Alpha'
$b = $s.Shapes.AddShape(1, 220, 170, 80, 40);  $b.Name = 'B'; $b.TextFrame.TextRange.Text = 'Beta'
$c = $s.Shapes.AddShape(1, 400, 100, 100, 100); $c.Name = 'C'; $c.TextFrame.TextRange.Text = 'Gamma'
function Sel([string[]]$names) { $s.Shapes.Range($names).Select() }

"--- shapes ---"
Sel @('A','B','C'); RunW 'LtxAlign' 'left';    "  B.Left=$([int]$b.Left) C.Left=$([int]$c.Left) (expect 50)"
Sel @('A','B','C'); RunW 'LtxAlign' 'top';     "  B.Top=$([int]$b.Top) (expect 120)"
Sel @('A','B','C'); RunW 'LtxAlign' 'center'
Sel @('A','B','C'); RunW 'LtxAlign' 'middle'
Sel @('A','B','C'); RunW 'LtxAlign' 'right'
Sel @('A','B','C'); RunW 'LtxAlign' 'bottom'
$a.Left = 50; $b.Left = 300; $c.Left = 600; $a.Top = 120; $b.Top = 120; $c.Top = 120
Sel @('A','B','C'); RunW 'LtxDistribute' 'h';  "  gaps: $([int]($b.Left-$a.Left-$a.Width)) / $([int]($c.Left-$b.Left-$b.Width)) (expect equal)"
Sel @('A','B','C'); RunW 'LtxMatchSize' 'width';  "  B.Width=$([int]$b.Width) (expect 120)"
Sel @('A','B','C'); RunW 'LtxMatchSize' 'height'; "  C.Height=$([int]$c.Height) (expect 60)"
Sel @('A','B');     RunW 'LtxSwap' $null;       "  A.Left=$([int]$a.Left) B.Left=$([int]$b.Left)"
Sel @('A','B','C'); RunW 'LtxSetGap' 'h'
Sel @('A','B','C'); RunW 'LtxCenterOnSlide' 'both'
Sel @('A','B');     RunW 'LtxStretch' 'right'
Sel @('C');         RunW 'LtxMakeSquare' $null
Sel @('A','B','C'); RunW 'LtxCopyGeometry' $null
Sel @('A');         RunW 'LtxRememberPosition' $null
Sel @('B');         RunW 'LtxApplyPosition' $true; "  B at A? $([int]$b.Left),$([int]$b.Top) vs $([int]$a.Left),$([int]$a.Top)"
Sel @('A');         RunW2 'LtxNudge' 5 0
Sel @('A');         RunW 'LtxZ' 0
Sel @('A','B');     RunW 'LtxGroupToggle' $null; "  shapes now: $($s.Shapes.Count) (group made)"
$s.Shapes.Item($s.Shapes.Count).Select(); RunW 'LtxGroupToggle' $null; "  shapes now: $($s.Shapes.Count) (ungrouped)"
$s.Shapes.Item('A').Select(); RunW 'LtxSelectMatch' 'type'; "  selected after SelectMatch type: $($pp.ActiveWindow.Selection.ShapeRange.Count) (expect 3)"
$s.Shapes.Item('A').Select(); RunW 'LtxSelectMatch' 'fill'
$s.Shapes.Item('A').Select(); RunW 'LtxSelectMatch' 'text'
Sel @('C'); RunW 'LtxFitToSlide' $null; "  C size: $([int]$c.Width)x$([int]$c.Height)"
Sel @('C'); RunW 'LtxCropToSlide' $null
"--- text / colour ---"
Sel @('A'); foreach ($cmd in 'grow','shrink','bullets','upper','lower','title','sentence','super','sub','clear','spacemore','spaceless','alignleft','aligncenter','justify','stripspaces') { RunW 'LtxText' $cmd }
Sel @('A'); RunW 'LtxCycleFontColor' $null; "  A font colour: $($a.TextFrame.TextRange.Font.Color.RGB)"
Sel @('A'); RunW 'LtxCycleFill' $null; "  A fill: $($a.Fill.ForeColor.RGB)"
Sel @('A'); RunW 'LtxCycleLine' $null; "  A line weight: $($a.Line.Weight)"
Sel @('A'); RunW2 'LtxThemeColor' 'fill' 5
Sel @('A'); RunW 'LtxCopyFormat' $null
Sel @('B'); RunW 'LtxPasteFormat' $null; "  B fill after paste format: $($b.Fill.ForeColor.RGB) (A=$($a.Fill.ForeColor.RGB))"
Sel @('A'); RunW 'LtxAutofit' 1
Sel @('A'); RunW 'LtxMargins' 0
Sel @('A'); RunW 'LtxWordWrapToggle' $null
RunW2 'LtxStamp' 'REVIEW' 255
RunW 'LtxTurbo' 'stampdraft'
"--- master shapes / library ---"
Sel @('A'); RunW 'LtxMasterSave' $null
RunW 'LtxMasterInsert' $null
RunW 'LtxMasterManage' $null
$pp.ActiveWindow.View.GotoSlide(1); $pres.Slides.Item(1).Select()
RunW 'LtxLibrarySave' $null
RunW 'LtxLibraryInsert' $null
"  presentations open: $($pp.Presentations.Count)"
foreach ($p in @($pp.Presentations)) { if ($p.Name -ne $pres.Name) { $p.Saved = -1; $p.Close() } }; $pres.Windows.Item(1).Activate()
"--- import chart + explorer data + link maintenance ---"
$ch = $ws.ChartObjects().Add(300, 10, 200, 120); $ch.Chart.SetSourceData($ws.Range('A1:D4')); $ch.Chart.ChartType = 51; $ch.Name = 'SweepChart'
$ch.Activate()
$pp.ActiveWindow.View.GotoSlide(2)
RunW 'LtxImport' 'chart'; $sh = $pres.Slides.Item(2).Shapes.Item($pres.Slides.Item(2).Shapes.Count); "  chart link type=$($sh.Tags.Item('LTX_TYPE')) range=$($sh.Tags.Item('LTX_RANGE'))"
$sh.Select(); RunW 'LtxSetSizeMode' 'height'; "  sizemode tag: $($sh.Tags.Item('LTX_SIZEMODE'))"
$sh.Select(); RunW 'LtxRefreshSelected' $null; $sh = $pres.Slides.Item(2).Shapes.Item($pres.Slides.Item(2).Shapes.Count); "  after refresh type=$($sh.Tags.Item('LTX_TYPE')) range=$($sh.Tags.Item('LTX_RANGE'))"
$sh.Select(); RunW 'LtxGoToSource' $null
$sh.Select(); RunW 'LtxChangeSource' $null
RunW 'LtxOpenAllSources' $null
$sh.Select(); RunW 'LtxBreakLink' $null; "  after break, LTX_WB='$($sh.Tags.Item('LTX_WB'))'"
RunW 'LtxToggleAirplane' $null; RunW 'LtxToggleAirplane' $null
RunW 'LtxToggleBlackFonts' $null; RunW 'LtxToggleBlackFonts' $null
"--- deck / file ---"
RunW 'LtxDuplicateSlides' $null; "  slides: $($pres.Slides.Count)"
RunW 'LtxSlideNumbers' $true
RunW 'LtxSlideNumbers' $false
RunW 'LtxRemoveNotes' $null
RunW 'LtxRemoveStamps' $null
RunW 'LtxDeleteHiddenSlides' $null
RunW 'LtxSaveBackup' $null; "  backup exists: $((Get-ChildItem (Join-Path $OutDir 'Backups') -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0)"
RunW 'LtxExportPdf' $false; "  pdf exists: $(Test-Path (Join-Path $OutDir 'sweep_deck.pdf'))"
RunW 'LtxSyncSectionTitles' $null
RunW 'LtxSectionFromTitle' $null
RunW 'LtxTableOfContents' $null
RunW 'LtxPageSkipToggle' $null
RunW 'LtxFirstSlideNumber' $null
RunW 'LtxStyleApplyToSimilar' $null
RunW 'LtxStyleUpdateAll' $null
RunW 'LtxStyleClearTag' $null
RunW 'LtxTextLinkInfo' $null
RunW 'LtxLinkTable' $null
RunW 'LtxRibbon_Ping' $null
"--- new: deck tools ---"
RunW 'LtxHarmonizeFonts' $null
RunW 'LtxFontReport' $null
RunW 'LtxReplaceFontDeck' $null
RunW 'LtxAlignTitles' $null
$pp.ActiveWindow.View.GotoSlide(2); $s2 = $pres.Slides.Item(2)
$tb = $s2.Shapes.AddTable(3, 3, 50, 300, 300, 80); $tb.Name = 'T1'
$tb.Table.Cell(1,1).Shape.TextFrame.TextRange.Text = 'Item'; $tb.Table.Cell(1,2).Shape.TextFrame.TextRange.Text = '2024'; $tb.Table.Cell(1,3).Shape.TextFrame.TextRange.Text = '2025'
$tb.Table.Cell(2,1).Shape.TextFrame.TextRange.Text = 'Revenue'; $tb.Table.Cell(2,2).Shape.TextFrame.TextRange.Text = '1234.567'; $tb.Table.Cell(2,3).Shape.TextFrame.TextRange.Text = '-89.1'
$tb.Table.Cell(3,1).Shape.TextFrame.TextRange.Text = 'Margin'; $tb.Table.Cell(3,2).Shape.TextFrame.TextRange.Text = '12.34%'; $tb.Table.Cell(3,3).Shape.TextFrame.TextRange.Text = '8.5x'
$tb.Select(); RunW 'LtxTableNumberFormat' $null; "  cells: '$($tb.Table.Cell(2,2).Shape.TextFrame.TextRange.Text)' '$($tb.Table.Cell(2,3).Shape.TextFrame.TextRange.Text)' '$($tb.Table.Cell(3,2).Shape.TextFrame.TextRange.Text)' '$($tb.Table.Cell(3,3).Shape.TextFrame.TextRange.Text)'"
$tb.Select(); RunW 'LtxTableHeaderStyle' $null; "  header fill: $($tb.Table.Cell(1,1).Shape.Fill.ForeColor.RGB)"
$tb.Select(); RunW 'LtxTableDistribute' $null
$tb.Select(); RunW 'LtxTableFontSize' $null
$chs = $s2.Shapes.AddChart2(-1, 51, 400, 300, 300, 150); RunW 'LtxChartPalette' $null
RunW 'LtxSetChartPalette' $null
$s2.Comments.Add(10, 10, 'Reviewer', 'RV', 'Please check this number') | Out-Null
RunW 'LtxTurbo' 'sticky'
RunW 'LtxCollectNotes' $null; "  review slide present: " + ((($pres.Slides | ? { $_.Tags.Item('LTX_REVIEW') -ne '' }) | Measure-Object).Count)
RunW 'LtxDeleteComments' $null; "  comments left on slide 2: $($s2.Comments.Count)"
RunW 'LtxExportSectionsPdf' $null; "  section pdfs: " + ((Get-ChildItem $OutDir -Filter 'sweep_deck - *.pdf' | Measure-Object).Count)
"--- new: arrange / auto align / body / slide links ---"
$pp.ActiveWindow.View.GotoSlide(1); $s = $pres.Slides.Item(1)
$q = @(); foreach ($i in 1..4) { $sh = $s.Shapes.AddShape(1, 100 + $i*3, 100 + 50*$i + ($i % 2), 60 + ($i % 2) * 2, 30); $sh.Name = "aa$i"; $q += $sh.Name }
$s.Shapes.Range([string[]]$q).Select(); RunW 'LtxAutoAlign' $null; "  lefts after auto align: " + (($q | % { [int]$s.Shapes.Item($_).Left }) -join ',') + "  widths: " + (($q | % { [int]$s.Shapes.Item($_).Width }) -join ',')
$cont = $s.Shapes.AddShape(1, 400, 100, 300, 200); $cont.Name = 'cont'
$s.Shapes.Range([string[]](@('cont') + $q)).Select(); RunW 'LtxArrangeInContainer' $null; "  aa1 inside container: left=$([int]$s.Shapes.Item('aa1').Left) top=$([int]$s.Shapes.Item('aa1').Top) w=$([int]$s.Shapes.Item('aa1').Width)"
$s.Shapes.Range([string[]]$q).Select(); RunW 'LtxStack' 'v'
$s.Shapes.Item('aa1').Select(); foreach ($e in 'left','center','right','top','middle','bottom','width','fill') { RunW 'LtxAlignToBody' $e | Out-Null }; "  body align ran; aa1 now w=$([int]$s.Shapes.Item('aa1').Width)"
RunW 'LtxFontConsistencyCheck' $null
$pres.Slides.Item(1).Shapes.Title.Select(); RunW 'LtxConformRole' 'title'; RunW 'LtxConformRole' 'subtitle'; RunW 'LtxConformRole' 'body'
$s.Shapes.Item('aa1').Select(); RunW 'LtxConformLikeSelected' $null
RunW 'LtxRefreshSlideLinks' $null
RunW 'LtxUpdateSlideLink' $null
$pres.Save(); $pres.Close(); $pp.Quit(); $wb.Close($false); $xl.Quit()
"done"
