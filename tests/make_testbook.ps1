# Builds a small "model" workbook with blue inputs, black formulas, green links and red text
param([string]$Out)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1); $ws.Name = 'Model'
$ws2 = $wb.Worksheets.Add([Type]::Missing, $ws); $ws2.Name = 'Inputs'; $ws.Activate()
$ws2.Range('A1').Value2 = 'Tax rate'; $ws2.Range('B1').Value2 = 0.25; $ws2.Range('B1').Font.Color = 16711680  # blue (BGR)
$ws.Range('A1').Value2 = 'EUR m';     $ws.Range('B1').Value2 = '2024A'; $ws.Range('C1').Value2 = '2025E'; $ws.Range('D1').Value2 = '2026E'
$ws.Range('A1:D1').Font.Bold = $true; $ws.Range('A1:D1').Interior.Color = 8210719; $ws.Range('A1:D1').Font.Color = 16777215   # dark blue fill, white text
$ws.Range('A2').Value2 = 'Revenue';   $ws.Range('B2').Value2 = 100; $ws.Range('C2').Value2 = 112; $ws.Range('D2').Value2 = 125
$ws.Range('B2:D2').Font.Color = 16711680                                  # blue hard-codes
$ws.Range('A3').Value2 = 'EBITDA';    $ws.Range('B3').Formula = '=B2*0.3'; $ws.Range('C3').Formula = '=C2*0.32'; $ws.Range('D3').Formula = '=D2*0.33'
$ws.Range('B3:D3').Font.Color = 0                                         # black formulas
$ws.Range('A4').Value2 = 'Tax';       $ws.Range('B4').Formula = '=-B3*Inputs!$B$1'; $ws.Range('C4').Formula = '=-C3*Inputs!$B$1'; $ws.Range('D4').Formula = '=-D3*Inputs!$B$1'
$ws.Range('B4:D4').Font.Color = 32768                                     # green links
$ws.Range('A5').Value2 = 'Check';     $ws.Range('B5').Value2 = 'TBU';   $ws.Range('B5').Font.Color = 255   # red
$ws.Range('B2:D4').NumberFormat = '#,##0.0_);(#,##0.0)'
$ws.Range('A2:A5').Font.Color = 0
$ws.Columns.Item(1).ColumnWidth = 12; $ws.Range('B:D').ColumnWidth = 9
$ws.Range('A5:D5').Borders.Item(9).LineStyle = 1   # bottom border
$wb.SaveAs($Out, 51)
$wb.Close($false); $xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
"created $Out"
