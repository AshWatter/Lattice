Attribute VB_Name = "modLtxXlCharts"
' =====================================================================
'  Lattice for Excel - advanced charts and growth metrics
'  Football field, stacked column with totals, waterfall, CAGR and YoY rows.
' =====================================================================
Option Explicit

' ---------- Football field ----------
' Selection: 3 columns  [label | low | high]  (header row optional).
' Smart ordering: bars are plotted top-down in selection order (Excel
' reverses category axes, so we flip it back), optional sort by low value.
Public Sub LtxFootballField()
    Dim r As Range, hasHdr As Boolean, n As Long, ws As Worksheet, helper As Range, ch As ChartObject
    Dim i As Long, first As Long, sortIt As VbMsgBoxResult, med As Double
    Set r = Selection
    If r.Columns.Count < 3 Then MsgBox "Select three columns: label, low, high.", vbExclamation, LTX_APP_XL: Exit Sub
    Set ws = r.Parent
    hasHdr = Not IsNumeric(r.Cells(1, 2).Value)
    first = IIf(hasHdr, 2, 1)
    n = r.Rows.Count - first + 1
    sortIt = MsgBox("Sort bars by low value (ascending)?", vbYesNo + vbQuestion, LTX_APP_XL)
    ' helper block to the right: label | low | spread
    Set helper = ws.Cells(r.Row, r.Column + r.Columns.Count + 1).Resize(n + 1, 3)
    helper.Cells(1, 1).Value = "Lattice FF": helper.Cells(1, 2).Value = "Low": helper.Cells(1, 3).Value = "Spread"
    For i = 1 To n
        helper.Cells(i + 1, 1).Value = r.Cells(first + i - 1, 1).Value
        helper.Cells(i + 1, 2).Formula = "=" & r.Cells(first + i - 1, 2).Address
        helper.Cells(i + 1, 3).Formula = "=" & r.Cells(first + i - 1, 3).Address & "-" & r.Cells(first + i - 1, 2).Address
    Next
    If sortIt = vbYes Then
        helper.Offset(1, 0).Resize(n, 3).Sort Key1:=helper.Cells(2, 2), Order1:=xlDescending, Header:=xlNo
    End If
    helper.Font.Color = RGB(150, 150, 150): helper.Font.Size = 8
    Set ch = ws.ChartObjects.Add(r.Left + r.Width + 160, r.Top, 480, 40 + 28 * n)
    With ch.Chart
        .ChartType = xlBarStacked
        .SetSourceData helper.Offset(0, 0).Resize(n + 1, 3), xlColumns
        .HasLegend = False
        .HasTitle = True: .ChartTitle.Text = "Valuation summary": .ChartTitle.Font.Size = 11: .ChartTitle.Font.Bold = True
        .SeriesCollection(1).Format.Fill.Visible = msoFalse
        .SeriesCollection(1).Format.Line.Visible = msoFalse
        With .SeriesCollection(2)
            .Format.Fill.ForeColor.RGB = RGB(31, 56, 100)
            .Format.Line.Visible = msoFalse
            .HasDataLabels = False
        End With
        .ChartGroups(1).GapWidth = 60
        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlCategory).TickLabels.Font.Size = 9
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
        .Axes(xlValue).TickLabels.Font.Size = 9
        .Axes(xlValue).Format.Line.Visible = msoFalse
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
    End With
    ' low / high labels at the ends of each bar
    AddEndLabels ch.Chart, helper, n
    ch.Name = "Lattice football field"
End Sub

Private Sub AddEndLabels(cht As Chart, helper As Range, n As Long)
    Dim i As Long, s As Series
    On Error Resume Next
    Set s = cht.SeriesCollection(2)
    s.HasDataLabels = True
    For i = 1 To n
        s.Points(i).DataLabel.Text = Format$(helper.Cells(i + 1, 2).Value, "#,##0") & " - " & Format$(helper.Cells(i + 1, 2).Value + helper.Cells(i + 1, 3).Value, "#,##0")
        s.Points(i).DataLabel.Position = xlLabelPositionInsideBase
        s.Points(i).DataLabel.Font.Size = 8
        s.Points(i).DataLabel.Font.Color = RGB(255, 255, 255)
    Next
End Sub

' Vertical reference line (e.g. current share price) on the football field
Public Sub LtxFootballFieldLine()
    Dim ch As ChartObject, v As Double, s As Series, ws As Worksheet
    On Error Resume Next
    Set ch = ActiveSheet.ChartObjects("Lattice football field")
    If ch Is Nothing Then MsgBox "Build a football field on this sheet first.", vbInformation, LTX_APP_XL: Exit Sub
    v = Val(InputBox("Value for the reference line (e.g. current price):", LTX_APP_XL))
    If v = 0 Then Exit Sub
    With ch.Chart
        Set s = .SeriesCollection.NewSeries
        s.Name = "Reference"
        s.ChartType = xlXYScatterLines
        s.XValues = Array(v, v)
        s.Values = Array(0, 1)
        s.MarkerStyle = xlMarkerStyleNone
        s.Format.Line.ForeColor.RGB = RGB(192, 0, 0)
        s.Format.Line.DashStyle = msoLineDash
        s.Format.Line.Weight = 1.5
        .Axes(xlValue, xlSecondary).MinimumScale = 0
        .Axes(xlValue, xlSecondary).MaximumScale = 1
        .Axes(xlValue, xlSecondary).TickLabelPosition = xlTickLabelPositionNone
        .Axes(xlValue, xlSecondary).Format.Line.Visible = msoFalse
        .Axes(xlCategory, xlSecondary).Delete
    End With
End Sub

' ---------- Stacked column with totals ----------
' Selection: first row = series headers (first cell = corner), first column = categories.
Public Sub LtxStackedColumn()
    Dim r As Range, ws As Worksheet, ch As ChartObject, tot As Range, i As Long, nCat As Long, nSer As Long, s As Series
    Set r = Selection
    If r.Rows.Count < 2 Or r.Columns.Count < 2 Then MsgBox "Select a block with headers across the top and categories down the side.", vbExclamation, LTX_APP_XL: Exit Sub
    Set ws = r.Parent
    nCat = r.Rows.Count - 1: nSer = r.Columns.Count - 1
    ' total helper column to the right
    Set tot = ws.Cells(r.Row, r.Column + r.Columns.Count).Resize(r.Rows.Count, 1)
    tot.Cells(1, 1).Value = "Total"
    For i = 2 To r.Rows.Count
        tot.Cells(i, 1).Formula = "=SUM(" & r.Cells(i, 2).Resize(1, nSer).Address(False, False) & ")"
    Next
    tot.Font.Color = RGB(150, 150, 150): tot.Font.Size = 8
    Set ch = ws.ChartObjects.Add(r.Left + r.Width + 120, r.Top, 480, 300)
    With ch.Chart
        .ChartType = xlColumnStacked
        .SetSourceData ws.Range(r.Cells(1, 1), tot.Cells(r.Rows.Count, 1)), xlColumns
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom: .Legend.Font.Size = 9
        .ChartGroups(1).GapWidth = 60
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
        .Axes(xlValue).Format.Line.Visible = msoFalse
        .Axes(xlCategory).TickLabels.Font.Size = 9
        .Axes(xlValue).TickLabels.Font.Size = 9
        .ChartArea.Format.Line.Visible = msoFalse
        ' total series: invisible, labels at inside base show the total above the stack
        Set s = .SeriesCollection(nSer + 1)
        s.Format.Fill.Visible = msoFalse
        s.Format.Line.Visible = msoFalse
        s.HasDataLabels = True
        s.DataLabels.Position = xlLabelPositionInsideBase
        s.DataLabels.Font.Bold = True: s.DataLabels.Font.Size = 9
        s.DataLabels.NumberFormat = "#,##0"
        .Legend.LegendEntries(nSer + 1).Delete
        ApplyPalette ch.Chart, nSer
    End With
    ch.Name = "Lattice stacked column"
End Sub

Private Sub ApplyPalette(cht As Chart, n As Long)
    Dim pal As Variant, i As Long
    pal = Split(XlGet("ChartPalette", "1F3864|2E75B6|9DC3E6|BDD7EE|7F7F7F|BFBFBF|C00000|ED7D31"), "|")
    For i = 1 To n
        cht.SeriesCollection(i).Format.Fill.ForeColor.RGB = HexBgr(CStr(pal((i - 1) Mod (UBound(pal) + 1))))
        cht.SeriesCollection(i).Format.Line.Visible = msoFalse
    Next
End Sub

Private Function HexBgr(h As String) As Long
    HexBgr = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
End Function

' ---------- Waterfall (bridge) ----------
' Selection: 2 columns [label | delta]; first and last rows are treated as totals.
Public Sub LtxWaterfall()
    Dim r As Range, ws As Worksheet, n As Long, h As Range, i As Long, ch As ChartObject, first As Long, hasHdr As Boolean
    Set r = Selection
    If r.Columns.Count < 2 Then MsgBox "Select two columns: label, value (first and last rows are totals).", vbExclamation, LTX_APP_XL: Exit Sub
    Set ws = r.Parent
    hasHdr = Not IsNumeric(r.Cells(1, 2).Value)
    first = IIf(hasHdr, 2, 1)
    n = r.Rows.Count - first + 1
    ' helper: label | base (invisible) | up | down | total
    Set h = ws.Cells(r.Row, r.Column + r.Columns.Count + 1).Resize(n + 1, 5)
    h.Rows(1).Value = Array("Lattice WF", "Base", "Increase", "Decrease", "Total")
    For i = 1 To n
        h.Cells(i + 1, 1).Value = r.Cells(first + i - 1, 1).Value
        If i = 1 Or i = n Then
            h.Cells(i + 1, 2).Value = 0
            h.Cells(i + 1, 3).Value = 0: h.Cells(i + 1, 4).Value = 0
            h.Cells(i + 1, 5).Formula = IIf(i = 1, "=" & r.Cells(first, 2).Address, "=SUM(" & r.Cells(first, 2).Resize(n - 1, 1).Address & ")")
        Else
            h.Cells(i + 1, 5).Value = 0
            h.Cells(i + 1, 3).Formula = "=MAX(" & r.Cells(first + i - 1, 2).Address & ",0)"
            h.Cells(i + 1, 4).Formula = "=MAX(-" & r.Cells(first + i - 1, 2).Address & ",0)"
            h.Cells(i + 1, 2).Formula = "=SUM(" & r.Cells(first, 2).Resize(i - 1, 1).Address & ")-" & h.Cells(i + 1, 4).Address
        End If
    Next
    h.Font.Color = RGB(150, 150, 150): h.Font.Size = 8
    Set ch = ws.ChartObjects.Add(r.Left + r.Width + 200, r.Top, 520, 300)
    With ch.Chart
        .ChartType = xlColumnStacked
        .SetSourceData h, xlColumns
        .HasLegend = False
        .ChartGroups(1).GapWidth = 40
        .SeriesCollection(1).Format.Fill.Visible = msoFalse
        .SeriesCollection(1).Format.Line.Visible = msoFalse
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(84, 130, 53)
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
        .SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(31, 56, 100)
        For i = 2 To 4
            .SeriesCollection(i).Format.Line.Visible = msoFalse
            .SeriesCollection(i).HasDataLabels = True
            .SeriesCollection(i).DataLabels.NumberFormat = "#,##0;(#,##0);"
            .SeriesCollection(i).DataLabels.Font.Size = 8
        Next
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
        .Axes(xlValue).Format.Line.Visible = msoFalse
        .Axes(xlCategory).TickLabels.Font.Size = 9
        .ChartArea.Format.Line.Visible = msoFalse
    End With
    ch.Name = "Lattice waterfall"
End Sub

' ---------- CAGR and growth rows ----------
' Selection: a row (or several rows) of values across years. Inserts CAGR in the
' column right of the selection: =(last/first)^(1/(n-1))-1
' Ctrl+Shift+M: a CAGR cell to the right of each selected row (first to last numeric value,
' periods = numeric cells - 1). Works on row 1 and with text labels inside the selection.
Public Sub LtxInsertCAGR()
    Dim r As Range, rowRng As Range, n As Long, c As Range, firstC As Range, lastC As Range, done As Long, tgt As Range
    On Error GoTo fail
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set r = Selection
    For Each rowRng In r.Rows
        Set firstC = Nothing: Set lastC = Nothing: n = 0
        For Each c In rowRng.Cells
            If Not IsEmpty(c.Value) And Not IsError(c.Value) Then
                If VarType(c.Value) = vbDouble Or VarType(c.Value) = vbCurrency Or VarType(c.Value) = vbInteger Or VarType(c.Value) = vbLong Then
                    If firstC Is Nothing Then Set firstC = c
                    Set lastC = c: n = n + 1
                End If
            End If
        Next
        If n >= 2 Then
            Set tgt = rowRng.Cells(1, rowRng.Cells.Count).Offset(0, 1)
            tgt.Formula = "=IFERROR((" & lastC.Address(False, False) & "/" & firstC.Address(False, False) & ")^(1/" & (n - 1) & ")-1,""n.m."")"
            tgt.NumberFormat = "0.0%": tgt.Font.Italic = True: tgt.Font.Color = 0
            done = done + 1
        End If
    Next
    If done = 0 Then
        Application.StatusBar = "Lattice: CAGR needs a row with at least two numbers (first to last value, periods = count - 1)"
        Exit Sub
    End If
    If r.Row > 1 Then
        Set tgt = r.Cells(1, 1).Offset(-1, r.Columns.Count)
        If IsEmpty(tgt.Value) Then tgt.Value = "CAGR": tgt.Font.Bold = True: tgt.HorizontalAlignment = xlRight
    End If
    Application.StatusBar = "Lattice: CAGR added for " & done & " row(s)"
    Exit Sub
fail:
    MsgBox "CAGR: " & Err.Description, vbExclamation, "Lattice"
End Sub

' Inserts a "% growth" row under each selected row
Public Sub LtxInsertGrowthRow()
    Dim r As Range, i As Long, row As Range, c As Long
    Set r = Selection
    For i = r.Rows.Count To 1 Step -1
        Set row = r.Rows(i)
        row.Offset(1, 0).EntireRow.Insert
        row.Offset(1, 0).Cells(1, 1).Offset(0, -IIf(row.Column > 1, 1, 0)).Value = "  % growth"
        For c = 2 To row.Cells.Count
            With row.Offset(1, 0).Cells(1, c)
                .Formula = "=IFERROR(" & row.Cells(1, c).Address(False, False) & "/" & row.Cells(1, c - 1).Address(False, False) & "-1,""n.m."")"
                .NumberFormat = "0.0%"
                .Font.Italic = True: .Font.Color = 0: .Font.Size = row.Cells(1, c).Font.Size - 1
            End With
        Next
    Next
End Sub

' Margin row: selected row as % of a reference row (prompted)
Public Sub LtxInsertMarginRow()
    Dim r As Range, ref As Range, c As Long
    Set r = Selection.Rows(1)
    On Error Resume Next
    Set ref = Application.InputBox("Reference row (e.g. revenue):", LTX_APP_XL, Type:=8)
    On Error GoTo 0
    If ref Is Nothing Then Exit Sub
    r.Offset(1, 0).EntireRow.Insert
    r.Offset(1, 0).Cells(1, 1).Offset(0, -IIf(r.Column > 1, 1, 0)).Value = "  % margin"
    For c = 1 To r.Cells.Count
        With r.Offset(1, 0).Cells(1, c)
            .Formula = "=IFERROR(" & r.Cells(1, c).Address(False, False) & "/" & ref.Cells(1, c).Address(False, False) & ",""n.m."")"
            .NumberFormat = "0.0%": .Font.Italic = True: .Font.Color = 0
        End With
    Next
End Sub
