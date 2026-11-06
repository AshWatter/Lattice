Attribute VB_Name = "modLtxXlCases"
' =====================================================================
'  Lattice for Excel - case runner
'  Runs every case of the model (Base / Upside / Downside ...) by writing 1..n into the case
'  cell, recalculating and reading the chosen output cells, then writes a styled table
'  "Lattice Cases": one row per output, one column per case, plus change vs the first case.
'  The case cell is restored afterwards. LtxRunCasesFor is the non-interactive core (tests).
' =====================================================================
Option Explicit

Public Sub LtxRunCases()
    Dim caseCell As Range, outputs As Range, caseText As String, dd As Shape, def As String, res As String
    On Error GoTo fail
    Set dd = FindCaseDropDown(ActiveWorkbook)
    If Not dd Is Nothing Then def = dd.ControlFormat.LinkedCell
    If Len(def) = 0 Then def = NameRef(ActiveWorkbook, "Case")
    On Error Resume Next
    Set caseCell = Application.InputBox("Case cell (the number the model reads, 1 = first case):", "Lattice - run cases", def, Type:=8)
    On Error GoTo fail
    If caseCell Is Nothing Then Exit Sub
    Set caseCell = caseCell.Cells(1)
    If Not dd Is Nothing Then
        If Len(dd.ControlFormat.ListFillRange) > 0 Then caseText = JoinRange(Application.Range(dd.ControlFormat.ListFillRange))
    End If
    If Len(caseText) = 0 Then caseText = XlGet("CaseNames", "Base case|Upside|Downside")
    caseText = InputBox("Case names in order (separated by |):", "Lattice - run cases", caseText)
    If Len(Trim$(caseText)) = 0 Then Exit Sub
    XlSet "CaseNames", caseText
    On Error Resume Next
    Set outputs = Application.InputBox("Output cells to tabulate (Ctrl+click to add several; labels are read from the row):", "Lattice - run cases", Type:=8)
    On Error GoTo fail
    If outputs Is Nothing Then Exit Sub
    res = LtxRunCasesFor(caseCell.Address(True, True, xlA1, True), outputs.Address(True, True, xlA1, True), caseText)
    If Len(res) > 0 Then MsgBox "Run cases: " & res, vbExclamation, "Lattice"
    Exit Sub
fail:
    MsgBox "Run cases: " & Err.Description, vbExclamation, "Lattice"
End Sub

' Core: caseAddr and outputsAddr are A1 addresses (sheet-qualified; several areas allowed),
' namesText the case names separated by |. Returns "" or an error message.
Public Function LtxRunCasesFor(caseAddr As Variant, outputsAddr As Variant, namesText As Variant) As String
    Dim caseCell As Range, outs As Range, arr As Variant, n As Long, i As Long, k As Long, c As Range, ar As Range
    Dim wb As Workbook, rpt As Worksheet, prev As Variant, outCells As Collection, m As Long
    Dim vals() As Variant, fmts() As String, labels() As String, addrs() As String
    Dim hdr As Long, r As Long, baseV As Variant, lastCol As Long, keep As Object, wasScreen As Boolean
    On Error GoTo fail
    Set caseCell = Application.Range(CStr(caseAddr)).Cells(1)
    Set outs = Application.Range(CStr(outputsAddr))
    arr = Split(CStr(namesText), "|"): n = UBound(arr) + 1
    If n < 2 Then LtxRunCasesFor = "At least two case names are needed.": Exit Function
    Set wb = caseCell.Parent.Parent
    Set outCells = New Collection
    For Each ar In outs.Areas
        For Each c In ar.Cells
            outCells.Add c
        Next
    Next
    m = outCells.Count
    If m = 0 Then LtxRunCasesFor = "No output cells.": Exit Function
    ReDim vals(1 To m, 1 To n): ReDim fmts(1 To m): ReDim labels(1 To m): ReDim addrs(1 To m)
    prev = caseCell.Value
    wasScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set keep = ActiveSheet
    For k = 1 To n
        caseCell.Value = k
        Application.Calculate
        For i = 1 To m
            Set c = outCells(i)
            vals(i, k) = c.Value
            If k = 1 Then
                fmts(i) = c.NumberFormat
                labels(i) = RowLabel(c)
                addrs(i) = c.Parent.Name & "!" & c.Address(False, False)
            End If
        Next
    Next
    caseCell.Value = prev
    Application.Calculate

    Set rpt = ReportSheet(wb, "Lattice Cases")
    rpt.Cells.Clear
    hdr = 4
    rpt.Cells(hdr, 1).Value = "Output": rpt.Cells(hdr, 2).Value = "Cell"
    For k = 1 To n
        rpt.Cells(hdr, 2 + k).Value = Trim$(arr(k - 1))
    Next
    For k = 2 To n
        rpt.Cells(hdr, 2 + n + k - 1).Value = Trim$(arr(k - 1)) & " vs " & Trim$(arr(0))
    Next
    lastCol = 2 + n + (n - 1)
    For i = 1 To m
        r = hdr + i
        rpt.Cells(r, 1).Value = labels(i)
        rpt.Cells(r, 2).Value = addrs(i): rpt.Cells(r, 2).Font.Color = RGB(128, 128, 128): rpt.Cells(r, 2).Font.Size = 8
        For k = 1 To n
            rpt.Cells(r, 2 + k).Value = vals(i, k)
            rpt.Cells(r, 2 + k).NumberFormat = fmts(i)
            rpt.Cells(r, 2 + k).HorizontalAlignment = xlRight
        Next
        baseV = vals(i, 1)
        For k = 2 To n
            With rpt.Cells(r, 2 + n + k - 1)
                If IsNumeric(baseV) And IsNumeric(vals(i, k)) And Not IsEmpty(baseV) And Not IsEmpty(vals(i, k)) Then
                    If baseV <> 0 Then
                        .Value = vals(i, k) / baseV - 1
                        .NumberFormat = "+0.0%;-0.0%;0.0%"
                    Else
                        .Value = vals(i, k) - baseV
                        .NumberFormat = fmts(i)
                    End If
                    .HorizontalAlignment = xlRight
                    If .Value < 0 Then .Font.Color = RGB(192, 0, 0) Else .Font.Color = RGB(0, 112, 60)
                End If
            End With
        Next
    Next
    If n >= 2 Then
        With rpt.Range(rpt.Cells(hdr + 1, 2 + n + 1), rpt.Cells(hdr + m, lastCol)).FormatConditions.AddDatabar
            .BarColor.Color = RGB(46, 117, 182)
            .BarFillType = xlDataBarFillSolid
            .NegativeBarFormat.ColorType = xlDataBarColor
            .NegativeBarFormat.Color.Color = RGB(192, 0, 0)
        End With
    End If
    LtxStyleReport rpt, "Case runner", "Case cell " & caseCell.Parent.Name & "!" & caseCell.Address(False, False) & "   |   " & n & " cases   |   " & m & " outputs   |   values as of the run; re-run after changing the model", hdr, hdr + m, lastCol
    rpt.Columns(1).ColumnWidth = 34: rpt.Columns(2).ColumnWidth = 16
    For k = 3 To lastCol
        rpt.Columns(k).ColumnWidth = 14
    Next
    rpt.Range(rpt.Cells(hdr, 3), rpt.Cells(hdr, lastCol)).HorizontalAlignment = xlRight
    rpt.Activate
    Application.ScreenUpdating = wasScreen
    Exit Function
fail:
    LtxRunCasesFor = Err.Description
    On Error Resume Next
    If Not caseCell Is Nothing Then caseCell.Value = prev
    Application.ScreenUpdating = True
End Function

' Label for an output: the nearest text cell to its left on the same row, else its address.
Private Function RowLabel(c As Range) As String
    Dim k As Long, t As Range
    For k = c.Column - 1 To 1 Step -1
        Set t = c.Parent.Cells(c.Row, k)
        If VarType(t.Value) = vbString Then
            If Len(Trim$(t.Value)) > 0 Then RowLabel = Trim$(t.Value): Exit Function
        End If
    Next
    RowLabel = c.Parent.Name & "!" & c.Address(False, False)
End Function

' First Forms drop-down in the workbook that has a cell link (the case switch).
Private Function FindCaseDropDown(wb As Workbook) As Shape
    Dim ws As Worksheet, shp As Shape
    On Error Resume Next
    For Each ws In wb.Worksheets
        For Each shp In ws.Shapes
            If shp.Type = 8 Then
                If shp.FormControlType = xlDropDown Then
                    If Len(shp.ControlFormat.LinkedCell) > 0 Then Set FindCaseDropDown = shp: Exit Function
                End If
            End If
        Next
    Next
End Function

Private Function NameRef(wb As Workbook, label As String) As String
    On Error Resume Next
    NameRef = wb.Names(label).RefersToRange.Address(True, True, xlA1, True)
End Function

Private Function JoinRange(rng As Range) As String
    Dim c As Range, s As String
    For Each c In rng.Cells
        If Len(Trim$(c.Text)) > 0 Then s = s & IIf(Len(s) > 0, "|", "") & Trim$(c.Text)
    Next
    JoinRange = s
End Function

' The report sheet (reused when it exists), added after the last sheet, active sheet kept.
Private Function ReportSheet(wb As Workbook, nm As String) As Worksheet
    Dim ws As Worksheet, keep As Object
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set keep = ActiveSheet
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nm
        ws.Tab.Color = RGB(31, 56, 100)
        keep.Activate
    End If
    Set ReportSheet = ws
End Function
