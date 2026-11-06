# PowerPoint stickers: traffic lights, tick / cross, stamps (one per slide, removable), Harvey cycling.
param([string]$OutDir)
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
$fails = 0
function Check($name, $ok, $detail) { if ($ok) { "PASS $name" } else { $script:fails++; "FAIL $name  $detail" } }
$pp = New-Object -ComObject PowerPoint.Application; $pp.Visible = -1
$pres = $pp.Presentations.Add(-1); [void]$pres.Slides.Add(1, 12); [void]$pres.Slides.Add(2, 12); $sl = $pres.Slides.Item(1)
$pp.ActiveWindow.View.GotoSlide(1)
$procId = [LtxWin]::PidOf([IntPtr]$pp.HWND)
function RunP($macro, $arg) { $script:pp.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $script:pp, @($macro, $arg)) }
function Named($sl, $prefix) { $out = @(); foreach ($s in $sl.Shapes) { if ($s.Name -like "$prefix*") { $out += $s } }; return ,$out }

$n0 = $sl.Shapes.Count
RunP 'Lattice.ppam!LtxSticker' 'light_green'
$g = Named $sl 'LtxSticker light green'
Check 'green light inserted' ($g.Count -eq 1 -and $sl.Shapes.Count -eq $n0 + 1) "count=$($sl.Shapes.Count)"
if ($g.Count) { Check 'green light fill colour' ($g[0].Fill.ForeColor.RGB -eq (0 + 176*256 + 80*65536)) $g[0].Fill.ForeColor.RGB }
RunP 'Lattice.ppam!LtxSticker' 'lights'
Check 'three lights in a row' ((Named $sl 'LtxSticker light').Count -eq 4) "lights=$((Named $sl 'LtxSticker light').Count)"
RunP 'Lattice.ppam!LtxSticker' 'tick'
$t = Named $sl 'LtxSticker tick'
Check 'tick inserted with a check mark' ($t.Count -eq 1 -and $t[0].TextFrame.TextRange.Text -eq [string][char]10004) $(if ($t.Count) { [int][char]$t[0].TextFrame.TextRange.Text } else { 'missing' })
RunP 'Lattice.ppam!LtxSticker' 'stamp_draft'
$st = Named $sl 'LtxStamp'
Check 'DRAFT stamp top-right' ($st.Count -eq 1 -and $st[0].Left -gt $pres.PageSetup.SlideWidth / 2 -and $st[0].Top -lt 40) "n=$($st.Count)"
RunP 'Lattice.ppam!LtxSticker' 'stamp_confidential'
$st = Named $sl 'LtxStamp'
Check 'second stamp replaces the first (one per slide)' ($st.Count -eq 1 -and $st[0].TextFrame.TextRange.Text -eq 'CONFIDENTIAL') "n=$($st.Count) text=$(if ($st.Count) { $st[0].TextFrame.TextRange.Text })"
RunP 'Lattice.ppam!LtxSticker' 'stamp_draft_all'
Check 'DRAFT on every slide' ((Named $pres.Slides.Item(2) 'LtxStamp').Count -eq 1) ''
$m = Invoke-Watched $procId { RunP 'Lattice.ppam!LtxSticker' 'stamps_remove' }
Check 'remove all stamps' ((Named $sl 'LtxStamp').Count -eq 0 -and (Named $pres.Slides.Item(2) 'LtxStamp').Count -eq 0) "dialog=$m"
# Harvey cycle: insert a 25% ball via Turbo, select it, cycle twice -> 75
RunP 'Lattice.ppam!LtxTurbo' 'harvey25'
$h = Named $sl 'Harvey 25'
Check 'Harvey 25 inserted by Turbo' ($h.Count -eq 1) ''
if ($h.Count) {
    $h[0].Select(); RunP 'Lattice.ppam!LtxSticker' 'harvey_cycle'
    $h2 = Named $sl 'Harvey 50'; Check 'cycle 25 -> 50' ($h2.Count -eq 1 -and (Named $sl 'Harvey 25').Count -eq 0) ''
    if ($h2.Count) { $h2[0].Select(); RunP 'Lattice.ppam!LtxSticker' 'harvey_cycle'; Check 'cycle 50 -> 75' ((Named $sl 'Harvey 75').Count -eq 1) '' }
}
$pres.Saved = -1; $pres.Close(); $pp.Quit()
"stickers_test: $fails failure(s)"
