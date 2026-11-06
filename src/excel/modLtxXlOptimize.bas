Attribute VB_Name = "modLtxXlOptimize"
' =====================================================================
'  Lattice for Excel - workbook optimizer and health report
'  Shrinks bloated files (used-range reset, orphan names, excess styles,
'  dead formatting) and reports what makes a model slow or fragile.
' =====================================================================
Option Explicit

' ---------- Optimizer ----------
Public Sub LtxOptimizeWorkbook()
    Dim wb As Workbook, ws As Worksheet, before As Double, after As Double, log As String
    Dim nNames As Long, nStyles As Long, nSheets As Long, nCF As Long, nLinks As Long
    Set wb = ActiveWorkbook
    If wb.Path = "" Then MsgBox "Save the workbook first (the optimizer compares file size before and after).", vbInformation, LTX_APP_XL: Exit Sub
    If MsgBox("Lattice will: reset used ranges, delete #REF! names, remove unused custom cell styles, " & _
              "clear formats beyond the data on every sheet, and save a backup copy first." & vbLf & vbLf & "Continue?", vbOKCancel + vbQuestion, LTX_APP_XL) <> vbOK Then Exit Sub
    before = FileLen(wb.FullName)
    LtxSaveBackup
    Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
    On Error Resume Next
    ' 1) used ranges
    For Each ws In wb.Worksheets
        If ResetUsedRange(ws) Then nSheets = nSheets + 1
    Next
    ' 2) broken names
    nNames = DeleteBrokenNames(wb)
    ' 3) styles
    nStyles = DeleteUnusedStyles(wb)
    ' 4) external links report only (never auto-break)
    nLinks = CountExternalLinks(wb)
    Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
    wb.Save
    after = FileLen(wb.FullName)
    log = "Used range reset on " & nSheets & " sheet(s)" & vbLf & _
          nNames & " broken name(s) deleted" & vbLf & _
          nStyles & " unused style(s) deleted" & vbLf & _
          nLinks & " external link(s) found (not changed)" & vbLf & vbLf & _
          "File size: " & Format$(before / 1024, "#,##0") & " KB -> " & Format$(after / 1024, "#,##0") & " KB"
    MsgBox log, vbInformation, LTX_APP_XL & " - Optimizer"
End Sub

' Deletes rows/columns beyond the last cell that has a value, formula or comment.
Private Function ResetUsedRange(ws As Worksheet) As Boolean
    Dim lastR As Long, lastC As Long, ur As Range, c As Range
    On Error Resume Next
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function
    Set c = ws.Cells.Find("*", , xlFormulas, , xlByRows, xlPrevious)
    If c Is Nothing Then lastR = 1 Else lastR = c.Row
    Set c = ws.Cells.Find("*", , xlFormulas, , xlByColumns, xlPrevious)
    If c Is Nothing Then lastC = 1 Else lastC = c.Column
    ' keep room for comments and shapes
    Dim s As Shape
    For Each s In ws.Shapes
        If s.BottomRightCell.Row > lastR Then lastR = s.BottomRightCell.Row
        If s.BottomRightCell.Column > lastC Then lastC = s.BottomRightCell.Column
    Next
    If ur.Row + ur.Rows.Count - 1 > lastR Or ur.Column + ur.Columns.Count - 1 > lastC Then
        If ws.ProtectContents Then Exit Function
        ws.Range(ws.Cells(lastR + 1, 1), ws.Cells(ws.Rows.Count, 1)).EntireRow.Delete
        ws.Range(ws.Cells(1, lastC + 1), ws.Cells(1, ws.Columns.Count)).EntireColumn.Delete
        ResetUsedRange = True
    End If
    ur = ws.UsedRange   ' forces Excel to recompute
End Function

Private Function DeleteBrokenNames(wb As Workbook) As Long
    Dim n As Name, i As Long
    On Error Resume Next
    For i = wb.Names.Count To 1 Step -1
        Set n = wb.Names(i)
        If InStr(n.RefersTo, "#REF!") > 0 Then n.Delete: DeleteBrokenNames = DeleteBrokenNames + 1
    Next
End Function

Private Function DeleteUnusedStyles(wb As Workbook) As Long
    Dim st As Style, i As Long, used As Object, ws As Worksheet, c As Range, keep As Long
    On Error Resume Next
    If wb.Styles.Count < 60 Then Exit Function          ' nothing to gain
    Set used = CreateObject("Scripting.Dictionary")
    For Each ws In wb.Worksheets
        For Each c In ws.UsedRange.Cells
            used(c.Style.Name) = 1
        Next
    Next
    For i = wb.Styles.Count To 1 Step -1
        Set st = wb.Styles(i)
        If Not st.BuiltIn And Not used.Exists(st.Name) Then
            st.Delete
            If Err.Number = 0 Then DeleteUnusedStyles = DeleteUnusedStyles + 1
            Err.Clear
        End If
    Next
End Function

Private Function CountExternalLinks(wb As Workbook) As Long
    Dim l As Variant
    On Error Resume Next
    l = wb.LinkSources(xlExcelLinks)
    If Not IsEmpty(l) Then CountExternalLinks = UBound(l)
End Function

' Convert every formula in the workbook (or sheet) to values - for sending out
Public Sub LtxHardcodeWorkbook(Optional sheetOnly As Boolean = False)
    Dim ws As Worksheet, n As Long
    If MsgBox(IIf(sheetOnly, "Convert all formulas on this sheet to values?", "Convert ALL formulas in the workbook to values? (a backup copy is saved first)"), vbOKCancel + vbExclamation, LTX_APP_XL) <> vbOK Then Exit Sub
    If Not sheetOnly Then LtxSaveBackup
    Application.ScreenUpdating = False
    On Error Resume Next
    If sheetOnly Then
        ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value
    Else
        For Each ws In ActiveWorkbook.Worksheets
            If Not ws.ProtectContents Then ws.UsedRange.Value = ws.UsedRange.Value: n = n + 1
        Next
    End If
    Application.ScreenUpdating = True
End Sub

' ---------- Health report ----------
Public Sub LtxHealthReport()
    Dim wb As Workbook, ws As Worksheet, rpt As Worksheet, r As Long, c As Range, f As String
    Dim nFormulas As Long, nConst As Long, nVol As Long, nErr As Long, nExtRef As Long, nArr As Long, nCF As Long, nHidden As Long, nIferr As Long
    Dim vol As Variant, v As Variant, tot As Long
    Set wb = ActiveWorkbook
    vol = Array("OFFSET(", "INDIRECT(", "NOW(", "TODAY(", "RAND(", "RANDBETWEEN(", "CELL(", "INFO(")
    Application.ScreenUpdating = False
    On Error Resume Next
    Set rpt = wb.Worksheets("Lattice Health")
    On Error GoTo 0
    If rpt Is Nothing Then
        Set rpt = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        rpt.Name = "Lattice Health"
    Else
        rpt.Cells.Clear
    End If
    rpt.Range("A4:I4").Value = Array("Sheet", "Formulas", "Constants", "Volatile", "Errors", "External refs", "Array formulas", "Cond. formats", "IFERROR wraps")
    r = 5
    For Each ws In wb.Worksheets
        If ws.Name <> rpt.Name Then
            nFormulas = 0: nConst = 0: nVol = 0: nErr = 0: nExtRef = 0: nArr = 0: nIferr = 0
            On Error Resume Next
            nConst = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers).Count
            nErr = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors).Count
            nCF = ws.Cells.FormatConditions.Count
            For Each c In ws.UsedRange.SpecialCells(xlCellTypeFormulas).Cells
                nFormulas = nFormulas + 1
                f = UCase$(c.Formula)
                For Each v In vol
                    If InStr(f, v) > 0 Then nVol = nVol + 1: Exit For
                Next
                If InStr(f, "[") > 0 Then nExtRef = nExtRef + 1
                If c.HasArray Then nArr = nArr + 1
                If InStr(f, "IFERROR(") > 0 Then nIferr = nIferr + 1
            Next
            On Error GoTo 0
            If ws.Visible <> xlSheetVisible Then nHidden = nHidden + 1
            rpt.Cells(r, 1).Value = ws.Name & IIf(ws.Visible <> xlSheetVisible, " (hidden)", "")
            rpt.Cells(r, 2).Resize(1, 8).Value = Array(nFormulas, nConst, nVol, nErr, nExtRef, nArr, nCF, nIferr)
            If nVol > 0 Then rpt.Cells(r, 4).Font.Color = RGB(192, 0, 0)
            If nErr > 0 Then rpt.Cells(r, 5).Font.Color = RGB(192, 0, 0)
            r = r + 1
        End If
    Next
    rpt.Cells(r, 1).Value = "Total": rpt.Cells(r, 1).Font.Bold = True
    For tot = 2 To 9
        rpt.Cells(r, tot).Formula = "=SUM(" & rpt.Range(rpt.Cells(5, tot), rpt.Cells(r - 1, tot)).Address & ")"
        rpt.Cells(r, tot).Font.Bold = True
    Next
    rpt.Range(rpt.Cells(r, 1), rpt.Cells(r, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
    rpt.Range(rpt.Cells(5, 2), rpt.Cells(r, 9)).NumberFormat = "#,##0;-#,##0;""-"""
    rpt.Range(rpt.Cells(5, 2), rpt.Cells(r, 9)).HorizontalAlignment = xlRight
    Dim sizeTxt As String
    If wb.Path = "" Then sizeTxt = "unsaved" Else sizeTxt = Format$(FileLen(wb.FullName) / 1024, "#,##0") & " KB"   ' IIf would evaluate FileLen eagerly
    LtxStyleReport rpt, "Health report", wb.Name & "   |   " & sizeTxt & "   |   calc " & IIf(Application.Calculation = xlCalculationAutomatic, "automatic", "manual") & _
        "   |   names " & wb.Names.Count & " (" & CountBrokenNames(wb) & " broken)   |   styles " & wb.Styles.Count & "   |   external links " & CountExternalLinks(wb) & "   |   hidden sheets " & nHidden, 4, r, 9
    Application.ScreenUpdating = True
End Sub

Private Function CountBrokenNames(wb As Workbook) As Long
    Dim n As Name
    On Error Resume Next
    For Each n In wb.Names
        If InStr(n.RefersTo, "#REF!") > 0 Then CountBrokenNames = CountBrokenNames + 1
    Next
End Function

' Remove the report sheet again
Public Sub LtxHealthReportRemove()
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets("Lattice Health").Delete
    ActiveWorkbook.Worksheets("Lattice Audit").Delete
    Application.DisplayAlerts = True
End Sub
