# Word: import a range as picture and as table from the running Excel, refresh, list, break.
param([string]$Book, [string]$OutDir)
$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '..\build\lib.ps1')
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq "" } | Stop-Process -Force; Start-Sleep 1
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $true
$wb = $xl.Workbooks.Open($Book); $ws = $wb.Worksheets.Item('Model'); $ws.Activate(); $ws.Range('A1:D5').Select()
Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force
$exe = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe').'(default)'
$null = Start-Process $exe -ArgumentList '/q' -PassThru -WindowStyle Minimized
$wd = $null; for ($i = 0; $i -lt 40 -and -not $wd; $i++) { Start-Sleep -Milliseconds 500; try { $wd = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') } catch {} }
Start-Sleep 2
$doc = $wd.Documents.Add()
$procId = [LtxWin]::PidOf([IntPtr]$wd.ActiveWindow.Hwnd)
function RunW($name, $arg) {
    $m = Invoke-Watched $procId { if ($null -ne $arg) { [void]$wd.Run($name, $arg) } else { [void]$wd.Run($name) } }.GetNewClosure()
    if ($m -eq $null) { "OK   $name $arg" } elseif ($m -match 'Compile error|Run-time error|host error') { "FAIL $name $arg : $($m -replace '\s+',' ')" } else { "OK   $name $arg  [$($m -replace '\s+',' ')]" }
}
$doc.Content.Text = "Lattice Word test`r`r"
$doc.Paragraphs.Item(2).Range.Select()
RunW 'WdImport' 'picture'
"  inline shapes: $($doc.InlineShapes.Count)  alt: $($doc.InlineShapes.Item(1).AlternativeText.Substring(0, [Math]::Min(60, $doc.InlineShapes.Item(1).AlternativeText.Length)))"
$doc.Content.InsertParagraphAfter(); $doc.Paragraphs.Item($doc.Paragraphs.Count).Range.Select()
$ws.Range('A1:D5').Select()
RunW 'WdImport' 'table'
"  tables: $($doc.Tables.Count)  title: $($doc.Tables.Item(1).Title.Substring(0, [Math]::Min(40, $doc.Tables.Item(1).Title.Length)))  cell(2,2)='$($doc.Tables.Item(1).Cell(2,2).Range.Text.Trim([char]13, [char]7))'"
$ws.Range('B2').Value2 = 555
RunW 'WdRefreshAll'
"  inline shapes after refresh: $($doc.InlineShapes.Count)  tables: $($doc.Tables.Count)"
"  after refresh cell(2,2)='$($doc.Tables.Item(1).Cell(2,2).Range.Text.Trim([char]13, [char]7))' (expect 555.0)"
RunW 'WdListLinks'
$doc.InlineShapes.Item(1).Select(); RunW 'WdGoToSource'
$doc.InlineShapes.Item(1).Select(); RunW 'WdBreakLink'
RunW 'WdStripDoubleSpaces'; RunW 'WdTablesAutofit'; RunW 'WdPicturesToWidth'; RunW 'WdToggleBlack'; RunW 'WdToggleBlack'
$doc.SaveAs2((Join-Path $OutDir 'word_test.docx'), 16)
$doc.Close(0); $wd.Quit(0); $wb.Close($false); $xl.Quit()
"done"
