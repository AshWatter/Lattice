Attribute VB_Name = "modLtxXlModel"
' =====================================================================
'  Lattice for Excel - modelling helpers
' =====================================================================
Option Explicit

' Cycle $ anchoring of every reference in the selected formulas:
' A1 -> $A$1 -> A$1 -> $A1 -> A1
Public Sub LtxCycleAnchors()
    Dim c As Range, f As String, re As Object, m As Object, out As String, pos As Long, ref As String, nxt As String
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(\$?)([A-Za-z]{1,3})(\$?)(\d+)"
    For Each c In Selection.Cells
        If c.HasFormula Then
            f = c.Formula: out = "": pos = 1
            For Each m In re.Execute(f)
                out = out & Mid$(f, pos, m.FirstIndex + 1 - pos)
                If m.SubMatches(0) = "" And m.SubMatches(2) = "" Then
                    nxt = "$" & m.SubMatches(1) & "$" & m.SubMatches(3)
                ElseIf m.SubMatches(0) = "$" And m.SubMatches(2) = "$" Then
                    nxt = m.SubMatches(1) & "$" & m.SubMatches(3)
                ElseIf m.SubMatches(0) = "" And m.SubMatches(2) = "$" Then
                    nxt = "$" & m.SubMatches(1) & m.SubMatches(3)
                Else
                    nxt = m.SubMatches(1) & m.SubMatches(3)
                End If
                out = out & nxt
                pos = m.FirstIndex + m.Length + 1
            Next
            out = out & Mid$(f, pos)
            On Error Resume Next
            c.Formula = out
            On Error GoTo 0
        End If
    Next
End Sub

' Ctrl+Shift+N: toggles the sign. A formula already wrapped as =-( ... ) is unwrapped instead of
' getting another minus; a plain =-X becomes =X; numbers are negated.
Public Sub LtxFlipSign()
    Dim c As Range, f As String
    For Each c In Selection.Cells
        If c.HasFormula Then
            f = Mid$(c.Formula, 2)
            If WrapsWhole(f, "-(") Then
                c.Formula = "=" & Mid$(f, 3, Len(f) - 3)
            ElseIf Left$(f, 1) = "-" And IsSingleTerm(Mid$(f, 2)) Then
                c.Formula = "=" & Mid$(f, 2)
            ElseIf IsSingleTerm(f) Then
                c.Formula = "=-" & f
            Else
                c.Formula = "=-(" & f & ")"
            End If
        ElseIf IsNumeric(c.Value) And Not IsEmpty(c.Value) Then
            c.Value = -c.Value
        End If
    Next
End Sub

' True when the expression (without "=") starts with the given opener and the parenthesis opened
' there closes at the very end, i.e. the opener wraps the whole formula.
Private Function WrapsWhole(f As String, opener As String) As Boolean
    Dim i As Long, depth As Long, inText As Boolean, ch As String
    If StrComp(Left$(f, Len(opener)), opener, vbTextCompare) <> 0 Then Exit Function
    If Right$(f, 1) <> ")" Then Exit Function
    For i = Len(opener) To Len(f)
        ch = Mid$(f, i, 1)
        If ch = """" Then inText = Not inText
        If Not inText Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then
                depth = depth - 1
                If depth = 0 Then WrapsWhole = (i = Len(f)): Exit Function
            End If
        End If
    Next
End Function

' A reference, name, number or single function call with no top-level operator (no brackets needed).
Private Function IsSingleTerm(f As String) As Boolean
    Dim i As Long, depth As Long, inText As Boolean, ch As String
    If Len(f) = 0 Then Exit Function
    For i = 1 To Len(f)
        ch = Mid$(f, i, 1)
        If ch = """" Then inText = Not inText
        If Not inText Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then depth = depth - 1
            If depth = 0 And InStr("+-*/^&<>=", ch) > 0 Then Exit Function
        End If
    Next
    IsSingleTerm = True
End Function

Public Sub LtxToValues()
    Selection.Value = Selection.Value
End Sub

' Paste formulas exactly (no reference shifting)
Public Sub LtxCopyExact()
    Dim c As Range, s As String
    For Each c In Selection.Cells
        s = s & IIf(c.HasFormula, c.Formula, c.Value) & vbTab
    Next
    SaveSetting "Lattice", "Excel", "ExactClip", s
    SaveSetting "Lattice", "Excel", "ExactRows", CStr(Selection.Rows.Count)
    SaveSetting "Lattice", "Excel", "ExactCols", CStr(Selection.Columns.Count)
End Sub

Public Sub LtxPasteExact()
    Dim parts As Variant, r As Long, cc As Long, i As Long, rows As Long, cols As Long
    parts = Split(GetSetting("Lattice", "Excel", "ExactClip"), vbTab)
    rows = Val(GetSetting("Lattice", "Excel", "ExactRows", "0"))
    cols = Val(GetSetting("Lattice", "Excel", "ExactCols", "0"))
    If rows = 0 Then MsgBox "Use 'Copy Exact' first.", vbInformation, "Lattice": Exit Sub
    For r = 0 To rows - 1
        For cc = 0 To cols - 1
            If i <= UBound(parts) Then ActiveCell.Offset(r, cc).Formula = parts(i)
            i = i + 1
        Next
    Next
End Sub

' Fast fill right: fill the selection's first column across to the right
' edge of the data block above it.
Public Sub LtxFastFillRight()
    Dim lastCol As Long, r As Range
    Set r = Selection
    lastCol = r.Cells(1).Offset(-1, 0).End(xlToRight).Column
    If lastCol <= r.Column Or lastCol > 16000 Then
        lastCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.Count).Column
    End If
    If lastCol > r.Column Then r.Columns(1).AutoFill Range(r.Columns(1), r.Columns(1).Offset(0, lastCol - r.Column)), xlFillDefault
End Sub

Public Sub LtxFastFillDown()
    Dim lastRow As Long, r As Range
    Set r = Selection
    lastRow = r.Cells(1).Offset(0, -1).End(xlDown).Row
    If lastRow <= r.Row Or lastRow > 1000000 Then
        lastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    End If
    If lastRow > r.Row Then r.Rows(1).AutoFill Range(r.Rows(1), r.Rows(1).Offset(lastRow - r.Row, 0)), xlFillDefault
End Sub

' Simplify: evaluate the formula and replace it with its value, but keep the
' original in a note so it can be reviewed.
Public Sub LtxSimplify()
    Dim c As Range
    For Each c In Selection.Cells
        If c.HasFormula Then
            On Error Resume Next
            c.AddComment "Was: " & c.Formula
            On Error GoTo 0
            c.Value = c.Value
        End If
    Next
End Sub

' Wrap each selected formula in IFERROR(...,0)
' Ctrl+Shift+E: wraps the formula in IFERROR( ... ,0); run again on a wrapped formula to unwrap it.
Public Sub LtxWrapIfError()
    Dim c As Range, f As String
    For Each c In Selection.Cells
        If c.HasFormula Then
            f = Mid$(c.Formula, 2)
            If WrapsWhole(f, "IFERROR(") And Right$(f, 3) = ",0)" Then
                c.Formula = "=" & Mid$(f, 9, Len(f) - 11)
            ElseIf InStr(1, f, "IFERROR(", vbTextCompare) = 0 Then
                c.Formula = "=IFERROR(" & f & ",0)"
            End If
        End If
    Next
End Sub

' Insert a row of "check" cells: sum of selection should equal zero
Public Sub LtxInsertCheck()
    Dim r As Range
    Set r = Selection
    r.Offset(r.Rows.Count, 0).Rows(1).Cells(1).Formula = "=ROUND(SUM(" & r.Address & "),6)=0"
    r.Offset(r.Rows.Count, 0).Rows(1).Cells(1).Font.Color = RGB(128, 128, 128)
End Sub

' Quick chart from the selection
Public Sub LtxQuickChart(kind As Long)
    Dim ch As Object
    Set ch = ActiveSheet.Shapes.AddChart2(-1, kind, Selection.Left + Selection.Width + 20, Selection.Top, 360, 220)
    ch.Chart.SetSourceData Selection
    ch.Chart.HasLegend = (Selection.Rows.Count > 2)
End Sub

' Convert selection to a defined name (prompt)
Public Sub LtxNameRange()
    Dim n As String
    n = InputBox("Defined name for the selection:", "Lattice", "rng_" & Replace(Selection.Address(False, False), ":", "_"))
    If Len(n) = 0 Then Exit Sub
    ActiveWorkbook.Names.Add n, "=" & Selection.Address(External:=True)
End Sub

' Hide / show all sheets at once
Public Sub LtxUnhideAll()
    Dim ws As Object
    For Each ws In ActiveWorkbook.Sheets: ws.Visible = -1: Next
End Sub

' Quick save with backup copy in \Backups
Public Sub LtxSaveBackup()
    Dim bdir As String, p As String, nm As String
    If Len(ActiveWorkbook.Path) = 0 Then MsgBox "Save the workbook first.", vbInformation, "Lattice": Exit Sub
    bdir = ActiveWorkbook.Path & "\Backups"
    If Len(Dir(bdir, vbDirectory)) = 0 Then MkDir bdir
    nm = ActiveWorkbook.Name
    nm = Left$(nm, InStrRev(nm, ".") - 1) & "_" & Format$(Now, "yyyymmdd_hhnn") & Mid$(nm, InStrRev(nm, "."))
    p = bdir & "\" & nm
    ActiveWorkbook.SaveCopyAs p
    ActiveWorkbook.Save
    Application.StatusBar = "Lattice: backup saved to " & p
End Sub

' Add a dated tag comment (who / when) to the active cell
Public Sub LtxTagCell()
    Dim t As String
    t = InputBox("Note:", "Lattice")
    If Len(t) = 0 Then Exit Sub
    On Error Resume Next
    ActiveCell.Comment.Delete
    ActiveCell.AddComment Format$(Now, "yyyy-mm-dd") & " " & Application.UserName & ": " & t
End Sub
