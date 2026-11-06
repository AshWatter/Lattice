Attribute VB_Name = "modLtxXlData"
' =====================================================================
'  Lattice for Excel - comps and data tools
'  Quick statistics rows, rank / percentile, comps table formatter,
'  outlier flags, data cleanup (text numbers, trim, unmerge-and-fill,
'  blank rows) and unit scaling that keeps formulas live.
' =====================================================================
Option Explicit

' Adds Mean / Median / High / Low / 25th / 75th rows under a comps block.
' Selection = the numeric block (header row optional, label column optional).
Public Sub LtxQuickStats()
    Dim r As Range, firstRow As Long, firstCol As Long, nR As Long, nC As Long, i As Long, j As Long, out As Range
    Dim labels As Variant, fx As Variant, colRng As String
    Set r = Selection
    If r.Rows.Count < 2 Then MsgBox "Select the comps block (at least two rows).", vbExclamation, LTX_APP_XL: Exit Sub
    firstRow = IIf(IsNumeric(r.Cells(1, r.Columns.Count).Value) And Not IsEmpty(r.Cells(1, r.Columns.Count).Value), 1, 2)
    firstCol = IIf(IsNumeric(r.Cells(r.Rows.Count, 1).Value) And Not IsEmpty(r.Cells(r.Rows.Count, 1).Value), 1, 2)
    nR = r.Rows.Count - firstRow + 1: nC = r.Columns.Count - firstCol + 1
    labels = Array("Mean", "Median", "High", "Low", "25th percentile", "75th percentile")
    fx = Array("AVERAGE", "MEDIAN", "MAX", "MIN", "QUARTILE.INC", "QUARTILE.INC")
    Set out = r.Cells(r.Rows.Count, 1).Offset(2, 0)
    For i = 0 To 5
        out.Offset(i, 0).Value = labels(i): out.Offset(i, 0).Font.Bold = (i < 2): out.Offset(i, 0).Font.Italic = (i >= 2)
        For j = 0 To nC - 1
            colRng = r.Cells(firstRow, firstCol + j).Resize(nR, 1).Address(False, False)
            With out.Offset(i, firstCol - 1 + j)
                If i = 4 Then
                    .Formula = "=IFERROR(QUARTILE.INC(" & colRng & ",1),""n.a."")"
                ElseIf i = 5 Then
                    .Formula = "=IFERROR(QUARTILE.INC(" & colRng & ",3),""n.a."")"
                Else
                    .Formula = "=IFERROR(" & fx(i) & "(" & colRng & "),""n.a."")"
                End If
                .NumberFormat = r.Cells(r.Rows.Count, firstCol + j).NumberFormat
                .Font.Bold = (i < 2): .Font.Italic = (i >= 2): .Font.Color = 0
            End With
        Next
    Next
    out.Resize(1, r.Columns.Count).Borders(xlEdgeTop).LineStyle = xlContinuous
End Sub

' Rank (1 = highest) and percentile columns to the right of a single column
Public Sub LtxRankPercentile()
    Dim r As Range, c As Range, addr As String
    Set r = Selection.Columns(1)
    addr = r.Address(True, True)
    r.Cells(1, 1).Offset(-1, 1).Value = "Rank": r.Cells(1, 1).Offset(-1, 2).Value = "Pctile"
    For Each c In r.Cells
        If IsNumeric(c.Value) And Not IsEmpty(c.Value) Then
            c.Offset(0, 1).Formula = "=RANK.EQ(" & c.Address(False, False) & "," & addr & ",0)"
            c.Offset(0, 2).Formula = "=PERCENTRANK.INC(" & addr & "," & c.Address(False, False) & ")"
            c.Offset(0, 2).NumberFormat = "0%"
            c.Offset(0, 1).Font.Color = 0: c.Offset(0, 2).Font.Color = 0
        End If
    Next
End Sub

' Flags values outside mean +/- k standard deviations (per column) with a light red fill
Public Sub LtxFlagOutliers()
    Dim r As Range, col As Range, c As Range, m As Double, sd As Double, k As Double, n As Long
    Set r = Selection
    k = Val(InputBox("Flag values further than how many standard deviations from the column mean?", LTX_APP_XL, "2"))
    If k <= 0 Then Exit Sub
    For Each col In r.Columns
        If Application.WorksheetFunction.Count(col) >= 3 Then
            m = Application.WorksheetFunction.Average(col)
            sd = Application.WorksheetFunction.StDev(col)
            For Each c In col.Cells
                If IsNumeric(c.Value) And Not IsEmpty(c.Value) And sd > 0 Then
                    If Abs(c.Value - m) > k * sd Then c.Interior.Color = RGB(255, 199, 206): n = n + 1
                End If
            Next
        End If
    Next
    MsgBox n & " outlier(s) flagged.", vbInformation, LTX_APP_XL
End Sub

' Comps table formatter: header row style, banded rows, number formats
' guessed per column (x multiples, %, plain), right-aligned numbers, thin rules.
Public Sub LtxFormatComps()
    Dim r As Range, hdr As Range, col As Range, i As Long, sample As Variant, txt As String, body As Range
    Set r = Selection
    If r.Rows.Count < 2 Then Exit Sub
    Application.ScreenUpdating = False
    Set hdr = r.Rows(1)
    hdr.Font.Bold = True: hdr.WrapText = True: hdr.VerticalAlignment = xlBottom
    hdr.Interior.Color = RGB(31, 56, 100): hdr.Font.Color = RGB(255, 255, 255)
    hdr.HorizontalAlignment = xlCenter: hdr.Cells(1, 1).HorizontalAlignment = xlLeft
    Set body = r.Offset(1, 0).Resize(r.Rows.Count - 1, r.Columns.Count)
    body.Interior.ColorIndex = xlNone
    For i = 2 To r.Rows.Count Step 2: r.Rows(i).Interior.Color = RGB(242, 242, 242): Next
    For Each col In body.Columns
        txt = LCase$(hdr.Cells(1, col.Column - r.Column + 1).Value)
        sample = col.Cells(1, 1).Value
        If IsNumeric(sample) And Not IsEmpty(sample) Then
            col.HorizontalAlignment = xlRight
            If InStr(txt, "%") > 0 Or InStr(txt, "margin") > 0 Or InStr(txt, "growth") > 0 Or InStr(txt, "yield") > 0 Then
                col.NumberFormat = "0.0%;(0.0%);""-"""
            ElseIf InStr(txt, "/") > 0 Or InStr(txt, "x") = Len(txt) Or InStr(txt, "multiple") > 0 Or InStr(txt, "ev/") > 0 Or InStr(txt, "p/e") > 0 Then
                col.NumberFormat = "0.0x;(0.0x);""-"""
            ElseIf Abs(sample) < 10 And Abs(sample) <> Int(Abs(sample)) Then
                col.NumberFormat = "0.0x;(0.0x);""-"""
            Else
                col.NumberFormat = "#,##0;(#,##0);""-"""
            End If
        Else
            col.HorizontalAlignment = xlLeft
        End If
    Next
    r.Borders.LineStyle = xlNone
    hdr.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r.Font.Size = IIf(r.Font.Size > 10, 9, r.Font.Size)
    r.Columns.AutoFit
    Application.ScreenUpdating = True
End Sub

' ---------- cleanup ----------
Public Sub LtxCleanNumbers()
    Dim c As Range, s As String, n As Long, neg As Boolean
    For Each c In Selection.Cells
        If VarType(c.Value) = vbString And Len(c.Value) > 0 And Not c.HasFormula Then
            s = Trim$(Replace(Replace(Replace(c.Value, Chr$(160), ""), ",", ""), " ", ""))
            neg = (Left$(s, 1) = "(" And Right$(s, 1) = ")")
            s = Replace(Replace(Replace(Replace(s, "(", ""), ")", ""), "%", ""), "x", "")
            s = Replace(Replace(Replace(s, "$", ""), ChrW(8364), ""), ChrW(163), "")
            If IsNumeric(s) Then
                c.NumberFormat = IIf(InStr(c.Value, "%") > 0, "0.0%", "General")
                c.Value = IIf(neg, -1, 1) * Val(s) / IIf(InStr(c.Value, "%") > 0, 100, 1)
                n = n + 1
            End If
        End If
    Next
    MsgBox n & " text value(s) converted to numbers.", vbInformation, LTX_APP_XL
End Sub

Public Sub LtxTrimText()
    Dim c As Range, n As Long
    For Each c In Selection.Cells
        If VarType(c.Value) = vbString And Not c.HasFormula Then
            If c.Value <> Application.WorksheetFunction.Clean(Application.WorksheetFunction.Trim(c.Value)) Then
                c.Value = Application.WorksheetFunction.Clean(Application.WorksheetFunction.Trim(c.Value)): n = n + 1
            End If
        End If
    Next
    MsgBox n & " cell(s) trimmed.", vbInformation, LTX_APP_XL
End Sub

' Unmerge every merged area in the selection and fill the value down/right
Public Sub LtxUnmergeFill()
    Dim c As Range, ma As Range, v As Variant, n As Long
    Application.ScreenUpdating = False
    For Each c In Selection.Cells
        If c.MergeCells Then
            Set ma = c.MergeArea
            v = ma.Cells(1, 1).Value
            ma.UnMerge
            ma.Value = v
            n = n + 1
        End If
    Next
    Application.ScreenUpdating = True
    MsgBox n & " merged area(s) unmerged and filled.", vbInformation, LTX_APP_XL
End Sub

Public Sub LtxRemoveBlankRows()
    Dim r As Range, i As Long, n As Long
    Set r = Selection
    For i = r.Rows.Count To 1 Step -1
        If Application.WorksheetFunction.CountA(r.Rows(i)) = 0 Then r.Rows(i).EntireRow.Delete: n = n + 1
    Next
    MsgBox n & " blank row(s) removed.", vbInformation, LTX_APP_XL
End Sub

' Divide (or multiply) the selection by a factor, keeping formulas live
Public Sub LtxScaleSelection(factor As Double)
    Dim c As Range
    For Each c In Selection.Cells
        If c.HasFormula Then
            c.Formula = "=(" & Mid$(c.Formula, 2) & ")/" & factor
        ElseIf IsNumeric(c.Value) And Not IsEmpty(c.Value) Then
            c.Value = c.Value / factor
        End If
    Next
End Sub

Public Sub LtxToThousands(): LtxScaleSelection 1000: End Sub
Public Sub LtxToMillions(): LtxScaleSelection 1000000: End Sub
Public Sub LtxFromThousands(): LtxScaleSelection 0.001: End Sub

' Transpose-paste values of the selection at the active cell (prompt)
Public Sub LtxTransposeValues()
    Dim src As Range, dst As Range
    Set src = Selection
    On Error Resume Next
    Set dst = Application.InputBox("Top-left cell for the transposed values:", LTX_APP_XL, Type:=8)
    On Error GoTo 0
    If dst Is Nothing Then Exit Sub
    src.Copy
    dst.PasteSpecial xlPasteValuesAndNumberFormats, , , True
    Application.CutCopyMode = False
End Sub

' Sum / average / count of the selection shown in a message (and copied)
Public Sub LtxSelectionStats()
    Dim r As Range, s As String
    Set r = Selection
    With Application.WorksheetFunction
        s = "Sum: " & Format$(.Sum(r), "#,##0.00") & vbLf & "Average: " & Format$(.Average(r), "#,##0.00") & vbLf & _
            "Median: " & Format$(.Median(r), "#,##0.00") & vbLf & "Min: " & Format$(.Min(r), "#,##0.00") & vbLf & "Max: " & Format$(.Max(r), "#,##0.00") & vbLf & "Count: " & .Count(r)
    End With
    MsgBox s, vbInformation, LTX_APP_XL & " - Selection"
End Sub
