# Builds docs\examples\Lattice_frozen_panes_example.xlsx: two sheets with frozen panes and formulas whose
# precedents sit far out of view, to try the Trace In window and Ctrl+[ / Ctrl+] on. The file is
# scrubbed afterwards (build\scrub.ps1) so it carries no author name or local path.
param([string]$Out = (Join-Path (Split-Path $PSScriptRoot -Parent) 'docs\examples\Lattice_frozen_panes_example.xlsx'))
. (Join-Path $PSScriptRoot '..\build\scrub.ps1')
New-Item -ItemType Directory -Force (Split-Path $Out -Parent) | Out-Null
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
$calc = $wb.Worksheets.Item(1); $calc.Name = 'Calc'
$inp = $wb.Worksheets.Add([Type]::Missing, $calc); $inp.Name = 'Inputs'
# --- Inputs: frozen at A5, values far down ---
$inp.Range('A1').Value2 = 'INPUTS (frozen at row 5)'; $inp.Range('A1').Font.Bold = $true
$inp.Range('A2').Value2 = 'Growth'; $inp.Range('B2').Value2 = 0.05; $inp.Range('B2').NumberFormat = '0.0%'; $inp.Range('B2').Font.Color = 16711680
$inp.Range('A200').Value2 = 'Tax rate (far down)'; $inp.Range('B200').Value2 = 0.25; $inp.Range('B200').NumberFormat = '0.0%'; $inp.Range('B200').Font.Color = 16711680
# --- Calc: frozen at B3, formulas with precedents far away ---
$calc.Range('A1').Value2 = 'FROZEN PANES TEST   (frozen at B3: rows 1-2 and column A stay put)'; $calc.Range('A1').Font.Bold = $true; $calc.Range('A1').Font.Size = 12
$calc.Range('A2').Value2 = 'Try: select L170 then Ctrl+Shift+[  (Trace In);  Ctrl+[ / Ctrl+]  jump;  Ctrl+Shift+Enter  back'
$calc.Range('A2').Font.Color = 6567967
$calc.Range('K140').Value2 = 'Revenue base (row 140)'; $calc.Range('L140').Value2 = 1000; $calc.Range('L140').Font.Color = 16711680
$calc.Range('K160').Value2 = 'Link to row 140'; $calc.Range('L160').Formula = '=+L140'
$calc.Range('K170').Value2 = 'Grown with Inputs!B2'; $calc.Range('L170').Formula = '=L160*(1+Inputs!B2)'
$calc.Range('K171').Value2 = 'After tax (Inputs!B200, far down on Inputs)'; $calc.Range('L171').Formula = '=L170*(1-Inputs!B200)'
$calc.Range('AZ300').Value2 = 55; $calc.Range('AZ300').Font.Color = 16711680; $calc.Range('AY300').Value2 = 'Far right and far down'
$calc.Range('K172').Value2 = 'Plus AZ300 (far right)'; $calc.Range('L172').Formula = '=L171+AZ300'
$calc.Range('K175').Value2 = 'Check: L172 should be 1000*1.05*0.75+55 = 842.5'; $calc.Range('L175').Formula = '=L172'
$calc.Range('L140:L175').NumberFormat = '#,##0.0'; $calc.Columns('K').ColumnWidth = 44; $calc.Columns('L').ColumnWidth = 12
$calc.Range('A3').Value2 = 'Row labels in column A stay visible while you scroll'; $calc.Range('A150').Value2 = 'row 150 (column A frozen)'; $calc.Range('A170').Value2 = 'row 170'
$calc.Range('K5').Value2 = 'Also try the Capital IQ way: select L170, press Ctrl+[ twice (L160, then L140), then Ctrl+Shift+Enter twice to come back.'
$xl.Visible = $true
$inp.Activate(); $inp.Range('A5').Select() | Out-Null; $xl.ActiveWindow.FreezePanes = $true; $inp.Range('B2').Select() | Out-Null
$calc.Activate(); $calc.Range('B3').Select() | Out-Null; $xl.ActiveWindow.FreezePanes = $true
$xl.ActiveWindow.ScrollRow = 155; $xl.ActiveWindow.ScrollColumn = 10; $calc.Range('L170').Select() | Out-Null
if (Test-Path $Out) { Remove-Item $Out -Force }
$wb.SaveAs($Out, 51); $wb.Close($false); $xl.Quit()
Start-Sleep 1
Invoke-OfficeScrub $Out
"saved $Out"
