Attribute VB_Name = "modLtxXlChartNotes"
' =====================================================================
'  Lattice for Excel - chart annotations (think-cell style)
'  On a native Excel chart: total labels above stacked columns, a CAGR arrow across the
'  categories and a difference arrow between two categories. The arrows are shapes inside the
'  chart (Chart.Shapes), so they move with it and travel into PowerPoint through the picture
'  link. Totals are a real series (line without a line, labels above): when every series reads
'  from a range, the totals are formulas on a hidden helper sheet and stay live; otherwise they
'  are static values. Re-run a command to recompute; earlier annotations of that kind are replaced.
' =====================================================================
Option Explicit

Private Const HELPER_SHEET As String = "LatticeChartData"
Private Const NAVY_R As Long = 31, NAVY_G As Long = 56, NAVY_B As Long = 100

' ---------- ribbon entry points (work on the selected chart) ----------
Public Sub LtxChartTotals()
    Dim ch As Chart
    Set ch = TargetChart()
    If ch Is Nothing Then Exit Sub
    LtxChartTotalsFor ch
End Sub

Public Sub LtxChartCAGR()
    Dim ch As Chart
    Set ch = TargetChart()
    If ch Is Nothing Then Exit Sub
    LtxChartCAGRFor ch
End Sub

Public Sub LtxChartDelta()
    Dim ch As Chart, s As String, p As Long, n As Long
    Set ch = TargetChart()
    If ch Is Nothing Then Exit Sub
    n = ch.SeriesCollection(1).Points.Count
    s = InputBox("Difference arrow between categories (from,to):", "Lattice - difference arrow", "1," & n)
    If Len(s) = 0 Then Exit Sub
    p = InStr(s, ",")
    If p = 0 Then Exit Sub
    LtxChartDeltaFor ch, CLng(Val(Left$(s, p - 1))), CLng(Val(Mid$(s, p + 1)))
End Sub

Public Sub LtxChartNotesRemove()
    Dim ch As Chart, i As Long
    Set ch = TargetChart()
    If ch Is Nothing Then Exit Sub
    RemoveNotes ch, "LtxCAGR"
    RemoveNotes ch, "LtxDelta"
    For i = ch.SeriesCollection.Count To 1 Step -1
        If ch.SeriesCollection(i).Name = "Lattice total" Then ch.SeriesCollection(i).Delete
    Next
End Sub

' Non-interactive entry (tests): kind = totals | cagr | delta, a/b = category indexes for delta.
Public Function LtxChartNote(sheetName As Variant, chartName As Variant, kind As Variant, a As Variant, b As Variant) As String
    Dim ch As Chart
    On Error GoTo fail
    Set ch = Worksheets(CStr(sheetName)).ChartObjects(CStr(chartName)).Chart
    Select Case LCase$(CStr(kind))
        Case "totals": LtxChartTotalsFor ch
        Case "cagr":   LtxChartCAGRFor ch
        Case "delta":  LtxChartDeltaFor ch, CLng(a), CLng(b)
    End Select
    Exit Function
fail:
    LtxChartNote = Err.Description
End Function

Private Function TargetChart() As Chart
    On Error Resume Next
    Set TargetChart = ActiveChart
    If TargetChart Is Nothing Then
        If TypeName(Selection) = "ChartObject" Then Set TargetChart = Selection.Chart
    End If
    If TargetChart Is Nothing Then
        If ActiveSheet.ChartObjects.Count = 1 Then Set TargetChart = ActiveSheet.ChartObjects(1).Chart
    End If
    If TargetChart Is Nothing Then MsgBox "Select a chart first.", vbInformation, "Lattice"
End Function

' ---------- totals ----------
Public Sub LtxChartTotalsFor(ch As Chart)
    Dim nSer As Long, nPt As Long, i As Long, k As Long, tot As Series, helper As Worksheet, rowIdx As Long
    Dim f As String, rng As Range, srcRanges As Collection, dyn As Boolean, vals() As Double, v As Variant, fmt As String
    On Error GoTo fail
    For i = ch.SeriesCollection.Count To 1 Step -1
        If ch.SeriesCollection(i).Name = "Lattice total" Then ch.SeriesCollection(i).Delete
    Next
    nSer = ch.SeriesCollection.Count
    If nSer = 0 Then Exit Sub
    nPt = ch.SeriesCollection(1).Points.Count
    Set srcRanges = New Collection: dyn = True
    For i = 1 To nSer
        Set rng = SeriesValuesRange(ch.SeriesCollection(i))
        If rng Is Nothing Then dyn = False: Exit For
        If rng.Cells.Count <> nPt Then dyn = False: Exit For
        srcRanges.Add rng
    Next
    fmt = "#,##0"
    If dyn Then fmt = srcRanges(1).Cells(1).NumberFormat
    If fmt = "General" Then fmt = "#,##0.0"
    Set tot = ch.SeriesCollection.NewSeries
    tot.Name = "Lattice total"
    If dyn Then
        Set helper = HelperSheet(ChartWorkbook(ch))
        rowIdx = HelperRow(helper, ChartTag(ch))
        For k = 1 To nPt
            f = "="
            For i = 1 To srcRanges.Count
                f = f & IIf(i > 1, "+", "") & "'" & srcRanges(i).Parent.Name & "'!" & srcRanges(i).Cells(k).Address(True, True)
            Next
            helper.Cells(rowIdx, 1 + k).Formula = f
        Next
        tot.Values = helper.Range(helper.Cells(rowIdx, 2), helper.Cells(rowIdx, 1 + nPt))
    Else
        ReDim vals(1 To nPt)
        For i = 1 To nSer
            v = ch.SeriesCollection(i).Values
            For k = 1 To nPt
                If IsNumeric(v(k)) Then vals(k) = vals(k) + v(k)
            Next
        Next
        tot.Values = vals
    End If
    tot.ChartType = xlLine
    tot.Format.Line.Visible = msoFalse
    tot.MarkerStyle = xlMarkerStyleNone
    tot.HasDataLabels = True
    tot.DataLabels.Position = xlLabelPositionAbove
    tot.DataLabels.Font.Bold = True: tot.DataLabels.Font.Size = 9
    tot.DataLabels.Font.Color = RGB(NAVY_R, NAVY_G, NAVY_B)
    tot.DataLabels.NumberFormat = fmt
    On Error Resume Next
    If ch.HasLegend Then ch.Legend.LegendEntries(ch.Legend.LegendEntries.Count).Delete
    On Error GoTo 0
    Exit Sub
fail:
    MsgBox "Chart totals: " & Err.Description, vbExclamation, "Lattice"
End Sub

' ---------- CAGR arrow across the categories ----------
Public Sub LtxChartCAGRFor(ch As Chart)
    Dim n As Long, first As Double, last As Double, cagr As Double, x1 As Single, x2 As Single, y As Single
    Dim ln As Shape, lbl As Shape, tick As Shape, k As Long, topMin As Single, t As Single
    On Error GoTo fail
    RemoveNotes ch, "LtxCAGR"
    n = ch.SeriesCollection(1).Points.Count
    If n < 2 Then MsgBox "CAGR needs at least two categories.", vbInformation, "Lattice": Exit Sub
    first = CategoryTotal(ch, 1): last = CategoryTotal(ch, n)
    If first <= 0 Or last <= 0 Then MsgBox "CAGR needs positive first and last values.", vbInformation, "Lattice": Exit Sub
    cagr = (last / first) ^ (1 / (n - 1)) - 1
    topMin = 1E+9
    For k = 1 To n
        t = CategoryTop(ch, k)
        If t < topMin Then topMin = t
    Next
    x1 = CategoryCenter(ch, 1): x2 = CategoryCenter(ch, n)
    y = topMin - 24
    If y < 6 Then y = 6
    Set ln = ch.Shapes.AddLine(x1, y, x2, y)
    StyleLine ln, True
    ln.Name = "LtxCAGR line"
    Set tick = ch.Shapes.AddLine(x1, y - 4, x1, y + 4): StyleLine tick, False: tick.Name = "LtxCAGR tick"
    Set lbl = ch.Shapes.AddTextbox(msoTextOrientationHorizontal, (x1 + x2) / 2 - 42, y - 19, 84, 16)
    lbl.Name = "LtxCAGR label"
    StyleLabel lbl, "CAGR " & Format$(cagr, "+0.0%;-0.0%")
    Exit Sub
fail:
    MsgBox "CAGR arrow: " & Err.Description, vbExclamation, "Lattice"
End Sub

' ---------- difference arrow between two categories ----------
Public Sub LtxChartDeltaFor(ch As Chart, fromK As Long, toK As Long)
    Dim n As Long, vFrom As Double, vTo As Double, x As Single, yFrom As Single, yTo As Single
    Dim ln As Shape, guide As Shape, lbl As Shape, txt As String, lblLeft As Single
    On Error GoTo fail
    RemoveNotes ch, "LtxDelta"
    n = ch.SeriesCollection(1).Points.Count
    If fromK < 1 Or toK < 1 Or fromK > n Or toK > n Or fromK = toK Then MsgBox "Pick two different categories between 1 and " & n & ".", vbInformation, "Lattice": Exit Sub
    vFrom = CategoryTotal(ch, fromK): vTo = CategoryTotal(ch, toK)
    yFrom = CategoryTop(ch, fromK): yTo = CategoryTop(ch, toK)
    x = CategoryRight(ch, toK) + 8
    If fromK > toK Then x = CategoryRight(ch, fromK) + 8
    ' dotted level line from the source column to the arrow, then the vertical arrow itself
    Set guide = ch.Shapes.AddLine(CategoryCenter(ch, fromK), yFrom, x, yFrom)
    StyleLine guide, False: guide.Line.DashStyle = msoLineRoundDot: guide.Name = "LtxDelta guide"
    Set ln = ch.Shapes.AddLine(x, yFrom, x, yTo)
    StyleLine ln, True
    ln.Name = "LtxDelta arrow"
    If vFrom <> 0 Then txt = Format$(vTo / vFrom - 1, "+0.0%;-0.0%") Else txt = Format$(vTo - vFrom, "+#,##0;-#,##0")
    lblLeft = x + 4
    If lblLeft + 60 > ch.ChartArea.Width Then lblLeft = x - 64
    Set lbl = ch.Shapes.AddTextbox(msoTextOrientationHorizontal, lblLeft, (yFrom + yTo) / 2 - 8, 60, 16)
    lbl.Name = "LtxDelta label"
    StyleLabel lbl, txt
    Exit Sub
fail:
    MsgBox "Difference arrow: " & Err.Description, vbExclamation, "Lattice"
End Sub

' ---------- geometry helpers (chart-area coordinates) ----------
Private Function CategoryTotal(ch As Chart, k As Long) As Double
    Dim i As Long, v As Variant
    For i = 1 To ch.SeriesCollection.Count
        If ch.SeriesCollection(i).Name <> "Lattice total" Then
            v = ch.SeriesCollection(i).Values
            If IsNumeric(v(k)) Then CategoryTotal = CategoryTotal + v(k)
        End If
    Next
End Function

Private Function CategoryTop(ch As Chart, k As Long) As Single
    Dim i As Long, t As Single
    CategoryTop = 1E+9
    For i = 1 To ch.SeriesCollection.Count
        If ch.SeriesCollection(i).Name <> "Lattice total" Then
            t = ch.SeriesCollection(i).Points(k).Top
            If t < CategoryTop Then CategoryTop = t
        End If
    Next
End Function

Private Function CategoryCenter(ch As Chart, k As Long) As Single
    Dim i As Long, l As Single, r As Single, p As Point, n As Long
    l = 1E+9: r = -1E+9
    For i = 1 To ch.SeriesCollection.Count
        If ch.SeriesCollection(i).Name <> "Lattice total" Then
            Set p = ch.SeriesCollection(i).Points(k)
            If p.Left < l Then l = p.Left
            If p.Left + p.Width > r Then r = p.Left + p.Width
        End If
    Next
    CategoryCenter = (l + r) / 2
End Function

Private Function CategoryRight(ch As Chart, k As Long) As Single
    Dim i As Long, r As Single, p As Point
    r = -1E+9
    For i = 1 To ch.SeriesCollection.Count
        If ch.SeriesCollection(i).Name <> "Lattice total" Then
            Set p = ch.SeriesCollection(i).Points(k)
            If p.Left + p.Width > r Then r = p.Left + p.Width
        End If
    Next
    CategoryRight = r
End Function

Private Sub StyleLine(s As Shape, arrow As Boolean)
    s.Line.ForeColor.RGB = RGB(NAVY_R, NAVY_G, NAVY_B)
    s.Line.Weight = 1.25
    If arrow Then s.Line.EndArrowheadStyle = msoArrowheadTriangle: s.Line.EndArrowheadLength = msoArrowheadShort
End Sub

Private Sub StyleLabel(s As Shape, txt As String)
    With s.TextFrame2                                 ' Excel chart text boxes: TextFrame2 carries the text
        .TextRange.Text = txt
        .TextRange.Font.Size = 9: .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(NAVY_R, NAVY_G, NAVY_B)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 1: .MarginBottom = 1
        .WordWrap = msoFalse
    End With
    s.Fill.ForeColor.RGB = RGB(255, 255, 255): s.Fill.Visible = msoTrue
    s.Line.Visible = msoFalse
End Sub

Private Sub RemoveNotes(ch As Chart, prefix As String)
    Dim i As Long
    On Error Resume Next
    For i = ch.Shapes.Count To 1 Step -1
        If Left$(ch.Shapes(i).Name, Len(prefix)) = prefix Then ch.Shapes(i).Delete
    Next
End Sub

' ---------- live totals helper sheet ----------
Private Function ChartWorkbook(ch As Chart) As Workbook
    On Error Resume Next
    Set ChartWorkbook = ch.Parent.Parent.Parent          ' ChartObject -> Worksheet -> Workbook
    If ChartWorkbook Is Nothing Then Set ChartWorkbook = ActiveWorkbook
End Function

Private Function ChartTag(ch As Chart) As String
    On Error Resume Next
    ChartTag = ch.Parent.Parent.Name & "|" & ch.Parent.Name
End Function

Private Function HelperSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet, keep As Object
    On Error Resume Next
    Set ws = wb.Worksheets(HELPER_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set keep = ActiveSheet
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = HELPER_SHEET
        ws.Range("A1").Value = "Lattice chart totals (one row per chart: sheet|chart, then the totals). Hidden on purpose."
        ws.Visible = xlSheetHidden
        keep.Activate
    End If
    Set HelperSheet = ws
End Function

Private Function HelperRow(ws As Worksheet, tag As String) As Long
    Dim r As Long
    For r = 2 To 500
        If ws.Cells(r, 1).Value = tag Then HelperRow = r: Exit Function
        If IsEmpty(ws.Cells(r, 1).Value) Then ws.Cells(r, 1).Value = tag: HelperRow = r: Exit Function
    Next
    HelperRow = 501
End Function

' Values range of a series from its =SERIES(name, categories, values, order) formula; Nothing
' when the values are a literal array or a multi-area reference.
Private Function SeriesValuesRange(s As Series) As Range
    Dim parts As Collection, f As String, arg As String
    On Error Resume Next
    f = s.Formula
    If Left$(f, 8) <> "=SERIES(" Then Exit Function
    Set parts = SplitTopLevel(Mid$(f, 9, Len(f) - 9))
    If parts.Count < 3 Then Exit Function
    arg = parts(3)
    If Left$(arg, 1) = "{" Or Left$(arg, 1) = "(" Then Exit Function
    Set SeriesValuesRange = Application.Range(arg)
End Function

Private Function SplitTopLevel(s As String) As Collection
    Dim parts As Collection, i As Long, depth As Long, inText As Boolean, ch As String, cur As String
    Set parts = New Collection
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then inText = Not inText
        If Not inText Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then depth = depth - 1
            If ch = "," And depth = 0 Then parts.Add cur: cur = "": GoTo nextCh
        End If
        cur = cur & ch
nextCh:
    Next
    parts.Add cur
    Set SplitTopLevel = parts
End Function
