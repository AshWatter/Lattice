# Fallback icon generator (32x32 PNG) -> ribbon\icons_generated. The shipped set in ribbon\icons is the designed one
# (sources in ribbon\icons_svg, 16 px variants in ribbon\icons_16); this script never overwrites it.
param([string]$Out = (Join-Path (Split-Path $PSScriptRoot -Parent) 'ribbon\icons_generated'))
Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force $Out | Out-Null
$navy = [Drawing.Color]::FromArgb(31, 56, 100); $blue = [Drawing.Color]::FromArgb(46, 117, 182); $light = [Drawing.Color]::FromArgb(189, 215, 238)
$green = [Drawing.Color]::FromArgb(33, 115, 70); $white = [Drawing.Color]::White

function New-Icon([string]$name, [scriptblock]$draw) {
    $bmp = New-Object Drawing.Bitmap 32, 32
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'; $g.Clear([Drawing.Color]::Transparent)
    $tile = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(236, 241, 249)); $tilePen = New-Object Drawing.Pen ([Drawing.Color]::FromArgb(214, 223, 236)), 1
    RR $g $tilePen 0.5 0.5 31 31 9 $tile          # one rounded tile behind every glyph: the set reads as a family
    $p = New-Object Drawing.Pen $navy, 2.4; $p.StartCap = 'Round'; $p.EndCap = 'Round'; $p.LineJoin = 'Round'
    $pa = New-Object Drawing.Pen $blue, 2.4; $pa.StartCap = 'Round'; $pa.EndCap = 'Round'
    $bn = New-Object Drawing.SolidBrush $navy; $bb = New-Object Drawing.SolidBrush $blue; $bl = New-Object Drawing.SolidBrush $light; $bw = New-Object Drawing.SolidBrush $white
    & $draw $g $p $pa $bn $bb $bl $bw
    $g.Dispose()
    $bmp.Save((Join-Path $Out "$name.png"), [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
}
function Txt($g, $s, $x, $y, $size, $brush, $bold = $true) {
    $f = New-Object Drawing.Font 'Segoe UI', $size, $(if ($bold) { [Drawing.FontStyle]::Bold } else { [Drawing.FontStyle]::Regular }), 'Pixel'
    $g.DrawString($s, $f, $brush, [single]$x, [single]$y)
}
function RR($g, $pen, $x, $y, $w, $h, $r, $fill = $null) {   # rounded rect
    $path = New-Object Drawing.Drawing2D.GraphicsPath
    $path.AddArc($x, $y, $r, $r, 180, 90); $path.AddArc($x + $w - $r, $y, $r, $r, 270, 90); $path.AddArc($x + $w - $r, $y + $h - $r, $r, $r, 0, 90); $path.AddArc($x, $y + $h - $r, $r, $r, 90, 90); $path.CloseFigure()
    if ($fill) { $g.FillPath($fill, $path) }
    if ($pen) { $g.DrawPath($pen, $path) }
}
function Page($g, $p, $bl) { RR $g $p 7 3 18 26 3 $bl; $g.DrawLines($p, [Drawing.PointF[]]@([Drawing.PointF]::new(18,3), [Drawing.PointF]::new(18,10), [Drawing.PointF]::new(25,10))) }
function Grid($g, $p, $bl, $x, $y, $w, $h) { $g.FillRectangle($bl, $x, $y, $w, $h); $g.DrawRectangle($p, $x, $y, $w, $h); $g.DrawLine($p, $x, $y + $h/3, $x + $w, $y + $h/3); $g.DrawLine($p, $x, $y + 2*$h/3, $x + $w, $y + 2*$h/3); $g.DrawLine($p, $x + $w/3, $y, $x + $w/3, $y + $h); $g.DrawLine($p, $x + 2*$w/3, $y, $x + 2*$w/3, $y + $h) }
function Arrow($g, $pa, $x1, $y1, $x2, $y2) { $pa.CustomEndCap = New-Object Drawing.Drawing2D.AdjustableArrowCap 3, 3; $g.DrawLine($pa, [single]$x1, [single]$y1, [single]$x2, [single]$y2); $pa.EndCap = 'Round' }
function Check($g, $pen, $x, $y) { $g.DrawLines($pen, [Drawing.PointF[]]@([Drawing.PointF]::new($x, $y+4), [Drawing.PointF]::new($x+4, $y+8), [Drawing.PointF]::new($x+11, $y))) }

New-Icon 'new'       { param($g,$p,$pa,$bn,$bb,$bl,$bw) Page $g $p $bl; $g.DrawLine($pa, 16, 15, 16, 25); $g.DrawLine($pa, 11, 20, 21, 20) }
New-Icon 'file'      { param($g,$p,$pa,$bn,$bb,$bl,$bw) Page $g $p $bl; $g.DrawLine($p, 11, 16, 21, 16); $g.DrawLine($p, 11, 21, 21, 21) }
New-Icon 'send'      { param($g,$p,$pa,$bn,$bb,$bl,$bw) $pts = [Drawing.PointF[]]@([Drawing.PointF]::new(3,16), [Drawing.PointF]::new(29,4), [Drawing.PointF]::new(20,29), [Drawing.PointF]::new(15,18)); $g.FillPolygon($bl, $pts); $g.DrawPolygon($p, $pts); $g.DrawLine($p, 15, 18, 29, 4) }
New-Icon 'slides'    { param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.FillRectangle($bl, 9, 9, 20, 15); $g.DrawRectangle($p, 9, 9, 20, 15); $g.FillRectangle($bw, 3, 4, 20, 15); $g.DrawRectangle($p, 3, 4, 20, 15); $g.DrawLine($pa, 7, 9, 17, 9) }
New-Icon 'agenda'    { param($g,$p,$pa,$bn,$bb,$bl,$bw) foreach ($i in 0..2) { $y = 7 + $i*9; $g.FillEllipse($bb, 4, $y-2, 5, 5); $g.DrawLine($p, 13, $y, 28, $y) } }
New-Icon 'pagination'{ param($g,$p,$pa,$bn,$bb,$bl,$bw) Page $g $p $bl; RR $g $null 11 19 11 8 2 $bb; Txt $g '1' 13 17 9 $bw }
New-Icon 'deckcheck' { param($g,$p,$pa,$bn,$bb,$bl,$bw) $pts = [Drawing.PointF[]]@([Drawing.PointF]::new(16,3), [Drawing.PointF]::new(28,7), [Drawing.PointF]::new(26,20), [Drawing.PointF]::new(16,29), [Drawing.PointF]::new(6,20), [Drawing.PointF]::new(4,7)); $g.FillPolygon($bl, $pts); $g.DrawPolygon($p, $pts); $g2 = New-Object Drawing.Pen $green, 3; $g2.StartCap='Round'; $g2.EndCap='Round'; Check $g $g2 10 12 }
New-Icon 'decktools' { param($g,$p,$pa,$bn,$bb,$bl,$bw) foreach ($i in 0..2) { $y = 7 + $i*9; $g.DrawLine($p, 4, $y, 28, $y); $x = @(20, 9, 15)[$i]; $g.FillEllipse($bw, $x-4, $y-4, 8, 8); $g.DrawEllipse($pa, $x-4, $y-4, 8, 8) } }
New-Icon 'library'   { param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.FillRectangle($bl, 4, 5, 7, 24); $g.DrawRectangle($p, 4, 5, 7, 24); $g.FillRectangle($bw, 12, 5, 7, 24); $g.DrawRectangle($p, 12, 5, 7, 24); $g.FillRectangle($bb, 20, 7, 8, 22); $g.DrawRectangle($p, 20, 7, 8, 22); $g.DrawLine($p, 6, 10, 9, 10); $g.DrawLine($p, 14, 10, 17, 10) }
New-Icon 'turbo'     { param($g,$p,$pa,$bn,$bb,$bl,$bw) RR $g $p 3 3 26 26 6 $bl; $pts = [Drawing.PointF[]]@([Drawing.PointF]::new(18,6), [Drawing.PointF]::new(10,18), [Drawing.PointF]::new(16,18), [Drawing.PointF]::new(14,26), [Drawing.PointF]::new(22,14), [Drawing.PointF]::new(16,14)); $g.FillPolygon($bb, $pts) }
New-Icon 'master'    { param($g,$p,$pa,$bn,$bb,$bl,$bw) $pts = @(); for ($i = 0; $i -lt 10; $i++) { $a = [Math]::PI/2 + $i * [Math]::PI/5; $r = $(if ($i % 2 -eq 0) { 13 } else { 5.5 }); $pts += [Drawing.PointF]::new(16 + $r*[Math]::Cos($a), 16 - $r*[Math]::Sin($a)) }; $g.FillPolygon($bl, [Drawing.PointF[]]$pts); $g.DrawPolygon($p, [Drawing.PointF[]]$pts) }
New-Icon 'colortools'{ param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.FillEllipse($bl, 3, 3, 26, 26); $g.DrawEllipse($p, 3, 3, 26, 26); $g.FillEllipse($bb, 8, 9, 6, 6); $g.FillEllipse((New-Object Drawing.SolidBrush $green), 17, 7, 6, 6); $g.FillEllipse((New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(192,0,0))), 20, 16, 6, 6); $g.FillEllipse($bn, 10, 19, 6, 6) }
New-Icon 'styles'    { param($g,$p,$pa,$bn,$bb,$bl,$bw) RR $g $p 4 4 18 9 2 $bb; $g.DrawLine($p, 22, 8, 27, 8); $g.DrawLine($p, 27, 8, 27, 15); $g.DrawLine($p, 27, 15, 14, 15); $g.DrawLine($p, 14, 15, 14, 19); RR $g $p 11 19 6 9 2 $bl }
New-Icon 'selectmatch'{ param($g,$p,$pa,$bn,$bb,$bl,$bw) $d = New-Object Drawing.Pen $navy, 2; $d.DashStyle = 'Dash'; $g.DrawRectangle($d, 4, 4, 18, 18); $pts = [Drawing.PointF[]]@([Drawing.PointF]::new(17,15), [Drawing.PointF]::new(29,21), [Drawing.PointF]::new(23,23), [Drawing.PointF]::new(21,29)); $g.FillPolygon($bb, $pts); $g.DrawPolygon($p, $pts) }
New-Icon 'arrange'   { param($g,$p,$pa,$bn,$bb,$bl,$bw) foreach ($x in 4, 17) { foreach ($y in 4, 17) { $g.FillRectangle($(if ($x -eq 4 -and $y -eq 4) { $bb } else { $bl }), $x, $y, 11, 11); $g.DrawRectangle($p, $x, $y, 11, 11) } } }
New-Icon 'excel'     { param($g,$p,$pa,$bn,$bb,$bl,$bw) $gp = New-Object Drawing.Pen $green, 2.2; $gb = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(226,243,232)); Grid $g $gp $gb 3 5 17 17; Arrow $g $pa 21 14 30 14 }
New-Icon 'linktext'  { param($g,$p,$pa,$bn,$bb,$bl,$bw) Txt $g 'T' 3 1 20 $bn; $g.DrawEllipse($pa, 17, 18, 8, 6); $g.DrawEllipse($pa, 22, 22, 8, 6) }
New-Icon 'refresh'   { param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.DrawArc($p, 5, 5, 22, 22, 30, 280); $pts = [Drawing.PointF[]]@([Drawing.PointF]::new(26,4), [Drawing.PointF]::new(28,13), [Drawing.PointF]::new(19,11)); $g.FillPolygon($bn, $pts) }
New-Icon 'explorer'  { param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.FillRectangle($bw, 3, 4, 26, 24); $g.DrawRectangle($p, 3, 4, 26, 24); $g.FillRectangle($bl, 3, 4, 8, 24); $g.DrawLine($p, 11, 4, 11, 28); foreach ($y in 10, 16, 22) { $g.DrawLine($pa, 14, $y, 25, $y) } }
New-Icon 'highlight' { param($g,$p,$pa,$bn,$bb,$bl,$bw) $yb = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(255,230,0)); $g.FillRectangle($yb, 4, 20, 24, 8); $g.TranslateTransform(16,16); $g.RotateTransform(-45); RR $g $p -4 -14 8 20 2 $bb; $g.ResetTransform() }
New-Icon 'airplane'  { param($g,$p,$pa,$bn,$bb,$bl,$bw) $pts = [Drawing.PointF[]]@([Drawing.PointF]::new(29,16), [Drawing.PointF]::new(4,7), [Drawing.PointF]::new(9,16), [Drawing.PointF]::new(4,25)); $g.FillPolygon($bl, $pts); $g.DrawPolygon($p, $pts); $g.DrawLine($p, 9, 16, 29, 16) }
New-Icon 'link'      { param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.DrawEllipse($p, 3, 11, 14, 10); $g.DrawEllipse($pa, 15, 11, 14, 10) }
New-Icon 'settings'  { param($g,$p,$pa,$bn,$bb,$bl,$bw) for ($i = 0; $i -lt 8; $i++) { $a = $i * [Math]::PI/4; $g.DrawLine($p, 16 + 9*[Math]::Cos($a), 16 + 9*[Math]::Sin($a), 16 + 13*[Math]::Cos($a), 16 + 13*[Math]::Sin($a)) }; $g.FillEllipse($bl, 7, 7, 18, 18); $g.DrawEllipse($p, 7, 7, 18, 18); $g.FillEllipse($bw, 12, 12, 8, 8); $g.DrawEllipse($p, 12, 12, 8, 8) }
New-Icon 'shortcuts' { param($g,$p,$pa,$bn,$bb,$bl,$bw) RR $g $p 3 6 26 20 4 $bl; foreach ($x in 7, 13, 19) { $g.FillRectangle($bn, $x, 10, 4, 4) }; $g.FillRectangle($bn, 25, 10, 2, 4); $g.FillRectangle($bn, 9, 17, 14, 4) }
# Excel-specific
New-Icon 'backup'    { param($g,$p,$pa,$bn,$bb,$bl,$bw) RR $g $p 4 4 24 24 3 $bl; $g.FillRectangle($bw, 9, 4, 14, 8); $g.DrawRectangle($p, 9, 4, 14, 8); $g.FillRectangle($bb, 9, 18, 14, 10); $g.FillRectangle($bn, 18, 6, 3, 4) }
New-Icon 'optimize'  { param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.DrawArc($p, 4, 6, 24, 24, 180, 180); $g.DrawLine($pa, 16, 20, 24, 10); $g.FillEllipse($bn, 13, 17, 6, 6) }
New-Icon 'share'     { param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.DrawLine($p, 9, 16, 23, 8); $g.DrawLine($p, 9, 16, 23, 24); $g.FillEllipse($bb, 4, 11, 10, 10); $g.FillEllipse($bl, 20, 3, 10, 10); $g.DrawEllipse($p, 20, 3, 10, 10); $g.FillEllipse($bl, 20, 19, 10, 10); $g.DrawEllipse($p, 20, 19, 10, 10) }
New-Icon 'autocolor' { param($g,$p,$pa,$bn,$bb,$bl,$bw) Txt $g 'A' 6 -2 20 $bn; $g.FillRectangle($bb, 4, 25, 8, 5); $g.FillRectangle((New-Object Drawing.SolidBrush $green), 12, 25, 8, 5); $g.FillRectangle((New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(192,0,0))), 20, 25, 8, 5) }
New-Icon 'blackfonts'{ param($g,$p,$pa,$bn,$bb,$bl,$bw) Txt $g 'A' 6 -2 20 $bn; $g.FillRectangle($bn, 4, 25, 24, 5); $g.FillEllipse($bb, 22, 2, 8, 8); $g.FillEllipse((New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(192,0,0))), 25, 9, 6, 6) }
New-Icon 'format'    { param($g,$p,$pa,$bn,$bb,$bl,$bw) RR $g $p 13 3 6 12 2 $bl; $g.FillRectangle($bn, 10, 15, 12, 3); RR $g $null 10 18 12 9 3 $bb; $g.DrawLine($bw, 13, 21, 13, 25); $g.DrawLine($bw, 19, 21, 19, 25) }
New-Icon 'formulas'  { param($g,$p,$pa,$bn,$bb,$bl,$bw) $f = New-Object Drawing.Font 'Cambria', 20, ([Drawing.FontStyle]::Italic), 'Pixel'; $g.DrawString('fx', $f, $bn, 3, 2) }
New-Icon 'charts'    { param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.DrawLine($p, 4, 28, 29, 28); $g.FillRectangle($bl, 6, 16, 5, 12); $g.FillRectangle($bb, 13, 9, 5, 19); $g.FillRectangle($bn, 20, 4, 5, 24) }
New-Icon 'modelcheck'{ param($g,$p,$pa,$bn,$bb,$bl,$bw) Grid $g $p $bl 3 3 18 18; $g.FillEllipse($bw, 16, 16, 14, 14); $gp = New-Object Drawing.Pen $green, 2.6; $gp.StartCap='Round'; $gp.EndCap='Round'; $g.DrawEllipse($gp, 16, 16, 14, 14); Check $g $gp 18 20 }
New-Icon 'tracein'   { param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.FillRectangle($bl, 17, 9, 12, 14); $g.DrawRectangle($p, 17, 9, 12, 14); Arrow $g $pa 3 16 15 16; $g.DrawLine($pa, 3, 8, 8, 16); $g.DrawLine($pa, 3, 24, 8, 16) }
New-Icon 'visualize' { param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.FillEllipse($bl, 3, 3, 18, 18); $g.DrawEllipse($p, 3, 3, 18, 18); $g.DrawLine($p, 19, 19, 29, 29); $g.FillEllipse($bb, 8, 8, 8, 8) }
New-Icon 'stats'     { param($g,$p,$pa,$bn,$bb,$bl,$bw) $pts = [Drawing.PointF[]]@([Drawing.PointF]::new(25,6), [Drawing.PointF]::new(7,6), [Drawing.PointF]::new(17,16), [Drawing.PointF]::new(7,26), [Drawing.PointF]::new(25,26)); $g.DrawLines($p, $pts) }
New-Icon 'exportppt' { param($g,$p,$pa,$bn,$bb,$bl,$bw) $gp = New-Object Drawing.Pen $green, 2; $gb = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(226,243,232)); Grid $g $gp $gb 2 4 13 13; Arrow $g $pa 15 11 20 11; $ob = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(183,71,42)); $g.FillRectangle($ob, 19, 15, 11, 13); $g.FillRectangle($bw, 21, 18, 7, 2); $g.FillRectangle($bw, 21, 22, 7, 2) }
New-Icon 'exportword'{ param($g,$p,$pa,$bn,$bb,$bl,$bw) $gp = New-Object Drawing.Pen $green, 2; $gb = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(226,243,232)); Grid $g $gp $gb 2 4 13 13; Arrow $g $pa 15 11 20 11; $wb2 = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(43,87,154)); $g.FillRectangle($wb2, 19, 15, 11, 13); Txt $g 'W' 20 14 9 $bw }

New-Icon 'stickers'  { param($g,$p,$pa,$bn,$bb,$bl,$bw) $g.FillEllipse($bw, 4, 6, 20, 20); $g.DrawEllipse($p, 4, 6, 20, 20); $g.FillPie($bb, 4, 6, 20, 20, -90, 180); $gp = New-Object Drawing.Pen ([Drawing.Color]::FromArgb(0,140,60)), 3; Check $g $gp 18 18 }
New-Icon 'cases'     { param($g,$p,$pa,$bn,$bb,$bl,$bw) RR $g $p 3 4 26 10 2 $bl; $g.FillRectangle($bb, 5, 6, 10, 6); $g.DrawLine($p, 6, 20, 26, 20); $g.DrawLine($p, 6, 25, 26, 25); $g.DrawLine($p, 6, 30, 20, 30) }
"icons written to $Out : " + ((Get-ChildItem $Out -Filter *.png | Measure-Object).Count)
