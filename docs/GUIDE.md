# Lattice — what it does, what it doesn't, and how it compares

Lattice is a clean-room set of Office add-ins (Excel, PowerPoint, Word) that reproduces the
mechanics bankers, private-equity associates and hedge-fund analysts pay Macabacus and think-cell
for: model formatting and auditing shortcuts, linked Excel-to-PowerPoint exports with one-click
refresh, slide and shape tools, chart annotations, agendas and pagination. It is written in plain
VBA with a PowerShell build, carries no telemetry and no licence server, and installs per user
without admin rights.

It was built by one person directing an AI coding agent over a handful of sessions; the repository
history and `README.md` record what was built when. Nothing was decompiled or copied from the
commercial products: feature lists and screenshots of their public documentation were the spec,
the implementation is original.

> Status: personal-use quality, tested through COM-driven scripts on Microsoft 365 for Windows.
> Not a drop-in replacement for every corner of the commercial tools (see the gaps below).

---

## 1. The tab, the keys, the manual

- One **Lattice** tab per app. Keytips: `Alt`, then **K** in Excel, **L** in PowerPoint and Word.
  (Office reserves T, E, I, O, V and D for its 2003 menus and reassigns a custom tab that uses one
  of them, which is why the tab is not Alt+T.)
- Excel keyboard shortcuts live in a plain text file (`%APPDATA%\Lattice\excel_shortcuts.txt`,
  Lattice > Shortcuts > Edit) and can be rebound; the PDF manual (`docs\Lattice_Manual.pdf`)
  regenerates from that file with `build\manual.ps1`.
- Settings are key/value pairs under `%APPDATA%\Lattice` (hide gridlines on export, chart
  border, formula-error check, quiet refresh, picture format, airplane mode, match size mode).
- A Settings dialog per app (Lattice > Settings > Settings dialog...) edits all of them in one window.
- Everything Lattice draws follows one design system: navy / azure palette, tiled ribbon icons, a
  shared window theme (title band, column headers, flat buttons, key hints), styled report sheets,
  shortcut cards and the PDF manual with its cover.

## 2. Excel

### Formatting
- Number cycles (number, percent, multiple, date, currency, negative style), font / fill / border
  cycles, alignment, indent, wrap, case (UPPER / lower / Proper), header and total styles, sum row,
  copy / paste format, centre across selection.
- **AutoColor**: blue inputs, black formulas, green sheet links, purple external links, dark red for
  hard-codes embedded in formulas; selection, sheet or whole workbook, in bulk.
- **Black Fonts switch**: turn every font colour to black for a print or a screenshot and back again
  (the original colours are stored on a very-hidden sheet).

### Modelling
- `$` anchor cycle, flip sign (toggle), IFERROR wrap (toggle), to values, fast fill right / down
  following the header row, exact copy / paste, name range, tag cell, insert check row.
- **Case drop-down**: a bare Forms combo (input range = case names, cell link = selected case),
  the control everyone hand-builds for Base / Upside / Downside; copies work on any tab.
- **Case runner**: writes 1..n into the case cell, recalculates, and tabulates the output cells you
  pick (one row per output, one column per case, change versus the first case), on a styled sheet.
- **Library** of generators: LBO, DCF, distribution waterfall, accretion / dilution, returns (IRR /
  MOIC), debt schedule, credit statistics, Black-Scholes, two-way sensitivity (real data table),
  scenario switch block, timeline header, plus a personal snippet library.
- Cells / sheets / view helpers: clean numbers, trim, unmerge-and-fill, remove blank rows, scale to
  thousands / millions, transpose values, selection stats, unhide all, hide inactive sheets, compare
  sheets, gridlines / headings / formulas / freeze / zoom toggles, column width, autofit.

### Auditing (the part analysts use most)
- **Ctrl+[** jumps to the precedent, **Ctrl+]** to the dependent, **Ctrl+Shift+Enter** goes back;
  several candidates open a picker that follows your arrow keys live.
- **Trace In / Trace Out** window (Ctrl+Shift+[ / ]): a tree of precedents or dependents with values
  and formulas; Up / Down move Excel live, Right / Left expand / collapse, Enter jumps and closes
  (so F2 edits immediately), Ctrl+E toggles value / formula.
- **Model Check** pane: hard-codes inside formulas, inconsistent formulas across a row, errors, masked
  errors; click to jump, "Fix" pulls a hard-code out into its own input cell; styled audit report.
- Highlight hard-codes (Ctrl+Shift+K), formula flow (Ctrl+Shift+L), select errors / hard-codes /
  inconsistent, arrows on / off.

### Charts
- Football field (with reference line), stacked column with totals, waterfall / bridge, quick column
  / line / bar, CAGR (Ctrl+Shift+M), growth rows, margin row.
- **Chart annotations** on any native Excel chart: live total labels above stacked columns (a real
  series driven by formulas on a hidden helper sheet), a CAGR arrow across the categories, and a
  difference arrow between two categories with a dotted level line. They are shapes inside the chart,
  so they move with it and travel into PowerPoint through the picture link.

### Comps and data
- Stats rows (mean / median / quartiles), comps formatter, rank / percentile, outlier flags.

### Export to PowerPoint and Word
- As picture with fonts forced black (Ctrl+Alt+R), as editable table (Ctrl+Alt+T), chart as vector
  picture (Ctrl+Alt+C), native PowerPoint chart, embedded workbook; to Word as picture or table
  (Ctrl+Alt+W). Match Width / Height / Size replaces the selected shape in the deck in place.
- Gridlines and chart borders are removed for the export; formula errors in the range are flagged.

## 3. PowerPoint

- **Import from Excel** (picture with black fonts, table, chart, native chart, embedded workbook)
  and **linked text** (a number or a year inside a title tied to a cell).
- **Refresh** the selected item, the slide or the whole deck. Links are resilient (defined names,
  version-aware prompt when a V2 of the model is open instead of V1, file picker when the file moved,
  airplane mode). Refreshing a 40x10 block takes about a second: the black-font pass runs inside
  Excel through Application.Run instead of cell-by-cell COM calls.
- Explorer pane of every link, highlight links in slide sorter, go to source, change source, break link.
- **Shapes**: align to the first shape or the slide, distribute, gaps, same width / height / size,
  stretch, swap, nudge, arrange grid, Auto Align, align to the body area, arrange in a container,
  stack, z-order, group toggle, Select Match (same fill / font / size), convert to picture, crop to
  slide, remember / apply position.
- **Colour and text**: cycles, hex and theme colours, format painter, Styles (save a look, apply
  everywhere), text tools (case, bullets, superscript, strip double spaces), harmonise and conform
  fonts, font consistency check across titles and subtitles.
- **Presentation**: template, backup, PDF, email, **Agenda** (sections, dividers, table of contents
  kept in sync), **Pagination** (page numbers, blank pages), Deck Check, Deck Tools (table formats,
  chart palette, notes collector, per-section PDFs).
- **Insert**: Slide Library with Slide Links (push / pull), Turbo Shapes (Harvey balls, chevrons,
  timeline, numbered circles, sticky notes, callouts), Master Shapes (personal shape library),
  **Stickers** (traffic lights, tick / cross, Harvey-ball cycling, DRAFT / CONFIDENTIAL /
  PRELIMINARY / FOR DISCUSSION / ILLUSTRATIVE stamps, one per slide, removable in one go).

## 4. Word

- Import from Excel (picture with black fonts, or table), refresh selection or document, links list,
  go to source, break link, document tools (strip double spaces, autofit tables, fit pictures to
  width, black fonts toggle).

## 5. Compared with Macabacus and think-cell

Legend: **Yes** = Lattice has it; **Partial** = a smaller version; **No** = not built.

| Capability | Macabacus | think-cell | Lattice |
|---|---|---|---|
| Excel formatting cycles and AutoColor | Yes | — | **Yes** |
| Formula auditing: precedent / dependent jumps, trace window | Yes | — | **Yes** |
| Hard-code and consistency audit | Yes | — | **Yes** |
| Model generators (LBO, DCF, returns, debt, sensitivity) | Partial | — | **Yes** |
| Case drop-down and case runner (scenario table) | No | — | **Yes** |
| Excel to PowerPoint / Word links with refresh | Yes | Partial (data links) | **Yes** |
| Match size / replace in place on export | Yes | — | **Yes** |
| Version-aware refresh (V1 vs V2 open) | Partial | — | **Yes** |
| Shape alignment, distribution, Auto Align | Yes | — | **Yes** |
| Agenda / table of contents, pagination | Yes | Yes | **Yes** |
| Slide library and slide links | Yes | — | **Yes** |
| Styles (saved looks applied across a deck) | Yes | — | **Yes** |
| Harvey balls, stickers, stamps | Partial | Yes | **Yes** |
| Chart totals, CAGR and difference arrows | Partial | Yes | **Yes** (native Excel charts) |
| Waterfall, football field, stacked with totals | Yes | Yes | **Yes** |
| think-cell native chart engine (Mekko, Gantt, editable datasheet in PowerPoint) | — | Yes | **No** |
| Mekko / marimekko and Gantt charts | No | Yes | **No** |
| Automatic label placement and collision avoidance on charts | — | Yes | **No** |
| Workbook diff (compare two versions) | Yes | — | Partial (compare sheets) |
| Settings dialog (GUI) | Yes | Yes | Partial (prompts and a text file) |
| Mac support | Yes | Yes | **No** (Windows only) |
| Telemetry, licence server, subscription | Yes | Yes | **No** |

### What Lattice deliberately does not do
- It does not replace PowerPoint's chart engine. Annotations go on native Excel charts and arrive in
  slides as linked pictures (or native PowerPoint charts without annotations). think-cell's in-slide
  charts with automatic label layout, Mekko and Gantt are its real moat and are not in scope.
- No keyboard hook in PowerPoint (VBA cannot do it safely), so PowerPoint shortcuts are keytips.
- No cloud, no collaboration features, no Mac.

## 6. How it is built

- `src\excel`, `src\ppt`, `src\word`: VBA modules and form code; `src\shared` is compiled into
  both PowerPoint and Word. `ribbon\*.xml` are the customUI definitions, `ribbon\icons` the
  generated flat icons (`build\icons.ps1`).
- `build\build.ps1` drives Excel, PowerPoint and Word through COM: imports the modules, builds the
  UserForms from code (no binary .frx in the repo), compiles the project through the VBE with a
  watcher thread that dismisses the compile dialog, saves `.xlam` / `.ppam` / `.dotm` into `dist`,
  and injects the ribbon XML and icons with System.IO.Packaging. Static lints catch the VBA traps
  that only surface at compile time (reserved identifiers, one-letter procedures, one-line `If`
  procedures, declarations after a procedure).
- `build\install.ps1` installs per user: `%APPDATA%\Microsoft\AddIns` plus the Excel OPEN registry
  key, PowerPoint's own AddIns collection, and Word's STARTUP folder.
- `tests\` drives real Office instances through COM with a dialog watcher: sweeps of every command,
  semantic checks of the formula tools, the Excel-instance and refresh scenarios, the Trace window
  driven by mouse and keyboard, the case runner, the chart annotations and the stickers.
  `tests\run_all.ps1` chains them.

## 7. Install

1. Close Excel, PowerPoint and Word.
2. `powershell -File build\install.ps1` (or build first with `build\build.ps1`).
3. Open the apps; the Lattice tab appears. Shortcuts: `docs\Lattice_Manual.pdf`.

## 8. Licence

To be decided before publication (MIT suggested). Macabacus and think-cell are trademarks of
their owners; Lattice is not affiliated with either.
