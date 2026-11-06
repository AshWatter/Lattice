# Builds docs\Lattice_Manual.html and docs\Lattice_Manual.pdf: a cover, the Excel key bindings read
# from the live shortcut file, keytips for the three apps, and the feature map. The PDF is printed
# headlessly with Chrome or Edge (HTML only if neither works).
param([string]$OutDir = (Join-Path (Split-Path $PSScriptRoot -Parent) 'docs'))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutDir | Out-Null
$keysFile = Join-Path $env:APPDATA 'Lattice\excel_shortcuts.txt'

function Key($k) {   # OnKey syntax -> key caps
    $s = $k; $h = ''
    $h += '<span class="k' + $(if ($s -like '*^*') { ' on' } else { '' }) + '">Ctrl</span>'
    $h += '<span class="k' + $(if ($s -like '*%*') { ' on alt' } else { '' }) + '">Alt</span>'
    $h += '<span class="k' + $(if ($s -like '*+*') { ' on sh' } else { '' }) + '">Shift</span>'
    $s = $s -replace '[\^%+{}]', ''
    if ($s -eq '~') { $s = 'Enter' }
    $h += '<span class="k key">' + $s.ToUpper() + '</span>'
    $h
}
function Chips([string]$seq) { ($seq -split ' ' | % { if ($_ -eq '/') { '<span class="or">or</span>' } else { '<span class="k key' + $(if ($_ -eq 'Alt') { ' alt' } else { '' }) + '">' + $_ + '</span>' } }) -join '' }
function Row($label, $chips) { "<tr><td>$label</td><td>$chips</td></tr>" }
function Grp($t) { "<tr class='g'><td colspan='2'>$t</td></tr>" }
function Describe($m) { $s = $m -replace '^Ltx', '' -replace '^Xl', ''; $s = [regex]::Replace($s, '([a-z])([A-Z])', '$1 $2'); $s = $s -replace 'Cycle ', 'Cycle: ' -replace 'Ppt', 'PowerPoint' -replace ' To ', ' to ' -replace 'Go Back', 'Go back (after a jump)' -replace 'Flip Sign', 'Flip sign (toggle)' -replace 'Wrap If Error', 'IFERROR wrap (toggle)' -replace 'Insert CAGR', 'CAGR cell'; $s }

# ---- Excel keyboard shortcuts (from the live file, else defaults) ----
$xlRows = ''
$lines = if (Test-Path $keysFile) { Get-Content $keysFile } else { @('# --- formatting ---','^+1=LtxCycleNumber','^+5=LtxCyclePercent','^+6=LtxCycleMultiple','^+3=LtxCycleDate','^+4=LtxCycleCurrency','^+f=LtxCycleFontColor','^+h=LtxCycleFill','^+b=LtxCycleBorder','^+a=LtxAutoColor','^+z=LtxBlackToggle','# --- modelling ---','^+q=LtxCycleAnchors','^+n=LtxFlipSign','^+v=LtxToValues','^+r=LtxFastFillRight','^+d=LtxFastFillDown','^+e=LtxWrapIfError','# --- auditing ---','^{[}=LtxGoPrecedent','^{]}=LtxGoDependent','^+~=LtxGoBack','^+{[}=LtxTraceIn','^+{]}=LtxTraceOut','^+l=LtxFormulaFlow','^+k=LtxHighlightHardcodes','# --- export ---','^%r=LtxSendToPptPicture','^%t=LtxSendToPptTable','^%c=LtxSendToPptChart','^%w=LtxSendToWordPicture','^+u=LtxRefreshPptFromHere','# --- charts ---','^+m=LtxInsertCAGR','# --- file ---','^+s=LtxSaveBackup') }
foreach ($ln in $lines) {
    $ln = $ln.Trim()
    if ($ln -like '# ---*') { $xlRows += Grp(($ln -replace '# ---','' -replace '---','').Trim()) }
    elseif ($ln -match '^([^#=]+)=(\w+)') { $xlRows += Row (Describe $matches[2]) (Key $matches[1]) }
}

$excelKeytips = (Grp 'Excel ribbon keytips (Alt, K, then the letter)') +
  (Row 'Save + Backup' (Chips 'Alt K B')) + (Row 'Optimize workbook' (Chips 'Alt K O')) + (Row 'Share menu' (Chips 'Alt K H')) +
  (Row 'AutoColor' (Chips 'Alt K A')) + (Row 'Black Fonts switch' (Chips 'Alt K Z')) + (Row 'Format menu (Color / Number / Border / Alignment / Font / Case)' (Chips 'Alt K F')) +
  (Row 'Formulas menu (anchors, flip, IFERROR, fast fill, case drop-down)' (Chips 'Alt K U')) + (Row 'Charts menu (football field, waterfall, CAGR, chart annotations)' (Chips 'Alt K C')) + (Row 'Library (LBO, DCF, waterfall, run cases, ...)' (Chips 'Alt K L')) +
  (Row 'Model Check pane' (Chips 'Alt K M')) + (Row 'Trace In window' (Chips 'Alt K P')) + (Row 'Visualize menu' (Chips 'Alt K V')) +
  (Row 'Comps stats rows' (Chips 'Alt K S')) + (Row 'Export to PowerPoint' (Chips 'Alt K E')) + (Row 'Export to Word' (Chips 'Alt K W')) +
  (Row 'Refresh deck / document' (Chips 'Alt K R')) + (Row 'Shortcut card' (Chips 'Alt K K')) + (Row 'Settings' (Chips 'Alt K G'))

$pptRows = (Grp 'Presentation') +
  (Row 'New deck from template' (Chips 'Alt L N')) + (Row 'Agenda (sections, dividers, TOC)' (Chips 'Alt L G')) + (Row 'Pagination' (Chips 'Alt L P')) +
  (Row 'Deck Check' (Chips 'Alt L K')) + (Row 'Deck Tools (fonts, tables, charts, notes, conform titles)' (Chips 'Alt L D')) +
  (Grp 'Insert') + (Row 'Slide Library / Slide Links' (Chips 'Alt L B')) + (Row 'Stickers (lights, tick / cross, Harvey cycle, stamps)' (Chips 'Alt L J')) + (Row 'Turbo Shapes' (Chips 'Alt L U')) + (Row 'Master Shapes' (Chips 'Alt L M')) +
  (Grp 'Color and text') + (Row 'Color Tools' (Chips 'Alt L C')) + (Row 'Styles' (Chips 'Alt L Y')) + (Row 'Select Match' (Chips 'Alt L S')) + (Row 'Arrange Shapes (grid, container, stack, Auto Align, body area)' (Chips 'Alt L W')) +
  (Grp 'Shapes') + (Row 'Align left / centre / right' (Chips 'Alt L 1 / 2 / 3')) + (Row 'Align top / middle / bottom' (Chips 'Alt L 4 / 5 / 6')) +
  (Row 'Distribute horizontally / vertically' (Chips 'Alt L 7 / 8')) + (Row 'Same width / height / size' (Chips 'Alt L 9 / 0 / Q')) +
  (Grp 'Import and links') + (Row 'Import Excel selection (fonts forced black)' (Chips 'Alt L I')) + (Row 'Link Text' (Chips 'Alt L T')) +
  (Row 'Refresh only the selected item (also right-click > Lattice)' (Chips 'Alt L R')) + (Row 'Explorer Pane' (Chips 'Alt L X')) + (Row 'Highlight Links' (Chips 'Alt L H')) +
  (Row 'Airplane Mode' (Chips 'Alt L A')) + (Row 'Settings' (Chips 'Alt L O'))

$wordRows = (Grp 'Word') + (Row 'Import Excel selection' (Chips 'Alt L E')) + (Row 'Refresh selection' (Chips 'Alt L R')) + (Row 'Links list' (Chips 'Alt L L')) + (Row 'Airplane Mode' (Chips 'Alt L A')) + (Row 'Shortcut card' (Chips 'Alt L K'))

$crossRows = (Grp 'From Excel into PowerPoint / Word') +
  (Row 'Export selection as picture to PowerPoint (replaces the selected shape with Match Size)' (Key '^%r')) + (Row 'Export as editable table' (Key '^%t')) +
  (Row 'Export chart' (Key '^%c')) + (Row 'Export picture to Word' (Key '^%w')) + (Row 'Refresh the open deck from Excel' (Key '^+u'))

$traceRows = (Grp 'Inside the Trace In / Trace Out window') +
  (Row 'Move: Excel follows the highlighted reference' (Chips 'Up / Down')) + (Row 'Expand a reference into its own inputs' (Chips 'Right')) +
  (Row 'Collapse, or go to the parent' (Chips 'Left')) + (Row 'Go to the cell and close the window (then F2)' (Chips 'Enter')) +
  (Row 'Go, close and start editing' (Chips 'F2')) + (Row 'Toggle value / formula columns' (Chips 'Ctrl E')) + (Row 'Close' (Chips 'Esc'))

$features = @'
<div class="cols">
<div class="app xl"><h3>Excel</h3><ul>
<li><b>Workbook</b>: Save + Backup, Optimize (used ranges, broken names, unused styles), Share (values, PNG, health report).</li>
<li><b>Color</b>: AutoColor (blue inputs, black formulas, green links, purple external, dark red embedded hard-codes), font / fill / border cycles, Black Fonts switch.</li>
<li><b>Model</b>: format cascades; formulas ($ anchors, flip and IFERROR toggles, values, fast fill, exact paste, bare case drop-down); charts (football field, stacked + totals, waterfall, CAGR, growth rows, <b>chart annotations</b>: live totals, CAGR arrow, difference arrow); library (LBO, DCF, distribution waterfall, accretion / dilution, returns, debt, credit stats, Black-Scholes, sensitivity, scenario switch, timeline, <b>run cases</b> scenario table); cells, sheets, view.</li>
<li><b>Audit</b>: Model Check pane and report, Trace In / Out windows, Ctrl+[ / ] jumps, Visualize (hard-codes, inconsistent formulas, errors).</li>
<li><b>Comps</b>: stats rows, comps formatter, rank / percentile, outliers.</li>
<li><b>Export</b>: to PowerPoint (picture, table, chart, native chart, embedded workbook) and Word; Match Width / Height / Size; Refresh deck.</li>
</ul></div>
<div class="app pp"><h3>PowerPoint</h3><ul>
<li><b>Presentation</b>: template, backup, PDF, email, Agenda and section sync, Pagination, Deck Check, Deck Tools (harmonize / conform fonts, table formats, chart palette, notes, per-section PDFs).</li>
<li><b>Insert</b>: Slide Library with Slide Links, <b>Stickers</b> (traffic lights, tick / cross, Harvey cycling, DRAFT-style stamps), Turbo Shapes, Master Shapes.</li>
<li><b>Color / Text</b>: cycles, hex and theme colours, format painter, Styles, text tools, Select Match, Arrange Shapes.</li>
<li><b>Shapes</b>: align to first shape or slide, distribute, gaps, same size, stretch, swap, remember / apply position, z-order, convert to picture.</li>
<li><b>Import / Links</b>: Excel import with black fonts, Link Text, Match Size, Refresh (item / slide / deck), Explorer Pane, Highlight Links, Airplane Mode, V1 / V2 refresh prompt, go to source, change source, break link.</li>
</ul></div>
<div class="app wd"><h3>Word</h3><ul>
<li>Import from Excel as picture (black fonts) or table, refresh selection / document, links list, go to source, break link.</li>
<li>Document tools: strip double spaces, autofit tables, fit pictures to width, black fonts toggle.</li>
</ul></div>
</div>
'@

$css = @'
:root{--navy:#1f3864;--azure:#2e75b6;--sky:#dee5f0;--mist:#f3f5f9;--ink:#1f2937;--slate:#6b7280;--xl:#217346;--pp:#b7472a;--wd:#2b579a}
*{box-sizing:border-box}body{font-family:Segoe UI,Arial,sans-serif;color:var(--ink);margin:0;font-size:12px;background:#fff}
.page{padding:28px 36px}
.cover{height:100vh;background:linear-gradient(135deg,var(--navy) 0%,#17315a 55%,var(--azure) 100%);color:#fff;display:flex;flex-direction:column;justify-content:center;padding:70px 64px;page-break-after:always;position:relative;overflow:hidden}
.cover .mark{font-size:84px;font-weight:700;letter-spacing:-2px;line-height:1;margin:0}
.cover .tag{font-size:20px;opacity:.9;margin:14px 0 40px;max-width:720px;line-height:1.4}
.cover .pills span{display:inline-block;padding:8px 16px;border-radius:999px;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.35);margin-right:10px;font-weight:600;font-size:14px}
.cover .hero{margin-top:48px;display:flex;gap:40px}
.cover .hero div{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.25);border-radius:12px;padding:16px 20px;min-width:220px}
.cover .hero b{display:block;font-size:11px;letter-spacing:1.5px;text-transform:uppercase;opacity:.8;margin-bottom:8px}
.cover .k{background:rgba(255,255,255,.95);color:var(--navy);font-size:13px;min-width:30px;padding:5px 9px;margin-left:0;margin-right:6px}
.cover .k.alt{background:#f97316;color:#fff}.cover .foot{position:absolute;bottom:40px;left:64px;font-size:12px;opacity:.75}
.cover .grid{position:absolute;right:-60px;top:-60px;width:420px;height:420px;border-radius:50%;background:radial-gradient(circle,rgba(255,255,255,.12),rgba(255,255,255,0) 70%)}
h1{margin:0 0 2px;font-size:24px;color:var(--navy)}h1 span{background:linear-gradient(90deg,var(--navy),var(--azure));color:#fff;padding:3px 12px;border-radius:6px;margin-left:8px;font-size:14px;font-weight:600;vertical-align:middle}
p.sub{color:var(--slate);margin:0 0 16px}
h2{font-size:15px;color:var(--navy);margin:22px 0 8px;padding-bottom:4px;border-bottom:2px solid var(--sky);position:relative}
h2:after{content:"";position:absolute;left:0;bottom:-2px;width:64px;height:2px;background:var(--azure)}
h3{font-size:13px;margin:0 0 6px;color:#fff;background:var(--navy);display:inline-block;padding:3px 10px;border-radius:4px}
.app.xl h3{background:var(--xl)}.app.pp h3{background:var(--pp)}.app.wd h3{background:var(--wd)}
.app{border-left:3px solid var(--sky);padding-left:12px}.app.xl{border-color:var(--xl)}.app.pp{border-color:var(--pp)}.app.wd{border-color:var(--wd)}
table{width:100%;border-collapse:collapse;page-break-inside:auto}tr{page-break-inside:avoid}td{padding:4px 6px;border-bottom:1px solid #eef0f3;font-size:11px}
td:last-child{text-align:right;white-space:nowrap}tr.g td{background:var(--mist);font-weight:700;color:var(--navy);padding:8px 6px 4px;border:0;font-size:10.5px;text-transform:uppercase;letter-spacing:.8px}
.k{display:inline-block;min-width:22px;text-align:center;padding:2px 6px;margin-left:3px;border-radius:5px;background:#e5e7eb;color:#9ca3af;font-size:10px;font-weight:600;box-shadow:inset 0 -1px 0 rgba(0,0,0,.08)}
.k.on{background:var(--azure);color:#fff}.k.on.alt{background:#f97316}.k.on.sh{background:#16a34a}.k.key{background:var(--navy);color:#fff;min-width:14px}.k.key.alt{background:#f97316}.or{color:#9ca3af;font-size:10px;margin:0 3px}
.cols{display:flex;gap:22px}.cols>div{flex:1}ul{margin:4px 0;padding-left:16px}li{margin:3px 0;font-size:11px}
.two{display:flex;gap:28px}.two>div{flex:1}
.note{background:var(--mist);border-left:3px solid var(--azure);padding:8px 12px;border-radius:0 6px 6px 0;color:var(--ink);font-size:11px;margin:10px 0}
@page{size:A4;margin:0}@media print{.page{padding:14mm 16mm}}
'@
$date = Get-Date -Format 'dd MMM yyyy'
$html = "<!doctype html><html><head><meta charset='utf-8'><title>Lattice manual</title><style>$css</style></head><body>" +
  "<section class='cover'><div class='grid'></div><p class='mark'>Lattice</p>" +
  "<p class='tag'>The banker's add-in for Excel, PowerPoint and Word: linked exports that refresh, keyboard-first formula auditing, model and chart tools, slide hygiene. No licence server, no telemetry.</p>" +
  "<div class='pills'><span>Excel</span><span>PowerPoint</span><span>Word</span></div>" +
  "<div class='hero'><div><b>Open the tab</b><span class='k alt'>Alt</span><span class='k'>K</span> in Excel &nbsp;&nbsp; <span class='k alt'>Alt</span><span class='k'>L</span> in PowerPoint and Word</div>" +
  "<div><b>Audit</b><span class='k'>Ctrl</span><span class='k'>[</span> precedent &nbsp; <span class='k'>Ctrl</span><span class='k'>]</span> dependent</div>" +
  "<div><b>Export</b><span class='k'>Ctrl</span><span class='k'>Alt</span><span class='k'>R</span> picture to the deck</div></div>" +
  "<div class='foot'>Shortcuts and keytips &nbsp;|&nbsp; generated $date from your shortcut file</div></section>" +
  "<div class='page'><h1>Lattice <span>SHORTCUTS AND KEYTIPS</span></h1><p class='sub'>Keytips: press Alt, then K in Excel or L in PowerPoint and Word for the Lattice tab, then the letter. Office reserves T, E, I, O, V and D for its legacy menus, and L is Excel's Developer tab. Edit the Excel bindings under Lattice &gt; Shortcuts &gt; Edit; this manual regenerates from that file.</p>" +
  "<div class='two'><div><h2>Excel keyboard shortcuts</h2><table>$xlRows</table></div><div><h2>Excel keytips and cross-app</h2><table>$excelKeytips$crossRows$traceRows</table></div></div>" +
  "<div class='two'><div><h2>PowerPoint keytips</h2><table>$pptRows</table></div><div><h2>Word keytips</h2><table>$wordRows</table><div class='note'>Everything in the tab also works with the mouse; keytips are the fast path. Right-click a linked picture or table in PowerPoint for Refresh this item, Go to source, Conform and Break link.</div></div></div>" +
  "<h2>What each group does</h2>$features</div></body></html>"
$htmlPath = Join-Path $OutDir 'Lattice_Manual.html'
[IO.File]::WriteAllText($htmlPath, $html, (New-Object Text.UTF8Encoding $false))
$pdfPath = Join-Path $OutDir 'Lattice_Manual.pdf'
$browsers = @("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
              "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
              "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ }
if (Test-Path $pdfPath) { Remove-Item $pdfPath -Force }
$fileUrl = (New-Object System.Uri($htmlPath)).AbsoluteUri
$profileDir = Join-Path $env:TEMP 'lattice_pdf_profile'
foreach ($exe in $browsers) {
    foreach ($mode in @('--headless=new', '--headless')) {
        $pdfArgs = @($mode, '--disable-gpu', '--no-first-run', '--no-default-browser-check', "--user-data-dir=`"$profileDir`"", "--print-to-pdf=`"$pdfPath`"", '--no-pdf-header-footer', $fileUrl)
        $proc = Start-Process -FilePath $exe -ArgumentList $pdfArgs -WindowStyle Hidden -PassThru
        if (-not $proc.WaitForExit(60000)) { try { $proc.Kill() } catch {} }
        if ((Test-Path $pdfPath) -and (Get-Item $pdfPath).Length -gt 1000) { "printed with: $exe $mode"; break }
    }
    if (Test-Path $pdfPath) { break }
}
"html: $htmlPath"; if (Test-Path $pdfPath) { "pdf:  $pdfPath ($([int]((Get-Item $pdfPath).Length/1024)) KB)" } else { "pdf:  not produced (no headless browser could print it)" }
