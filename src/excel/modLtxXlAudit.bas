Attribute VB_Name = "modLtxXlAudit"
' =====================================================================
'  Lattice for Excel - auditing
' =====================================================================
Option Explicit

' Lists direct precedents of the active cell (any sheet / workbook) and
' jumps to the chosen one. Uses NavigateArrow which works across sheets.
Public Sub LtxPrecedents()
    ListAndGo True
End Sub

Public Sub LtxDependents()
    ListAndGo False
End Sub

Private Sub ListAndGo(prec As Boolean)
    Dim c As Range, list As String, items As Collection, i As Long, a As Long, link As Long, r As Range, choice As String
    Set c = ActiveCell
    Set items = New Collection
    Application.ScreenUpdating = False
    On Error Resume Next
    If prec Then c.ShowPrecedents Else c.ShowDependents
    a = 1
    Do
        link = 1
        Do
            Set r = Nothing
            Set r = c.NavigateArrow(prec, a, link)
            If r Is Nothing Then Exit Do
            If r.Address(External:=True) = c.Address(External:=True) Then Exit Do
            items.Add r.Address(External:=True)
            link = link + 1
            If link > 200 Then Exit Do
        Loop
        If link = 1 Then Exit Do
        a = a + 1
        If a > 50 Then Exit Do
    Loop
    c.Parent.ClearArrows
    c.Select
    On Error GoTo 0
    Application.ScreenUpdating = True
    If items.Count = 0 Then
        MsgBox IIf(prec, "No precedents.", "No dependents."), vbInformation, "Lattice"
        Exit Sub
    End If
    For i = 1 To items.Count: list = list & i & ": " & items(i) & vbLf: Next
    choice = InputBox(IIf(prec, "Precedents", "Dependents") & " of " & c.Address(False, False) & ":" & vbLf & list & vbLf & "Number to jump to:", "Lattice", "1")
    If Val(choice) >= 1 And Val(choice) <= items.Count Then
        Application.Goto Range(items(Val(choice))), False
    End If
End Sub

Public Sub LtxShowArrows()
    ActiveCell.ShowPrecedents
    ActiveCell.ShowDependents
End Sub

Public Sub LtxClearArrows()
    ActiveSheet.ClearArrows
End Sub

' Highlight formulas whose pattern differs from their left neighbour
Public Sub LtxInconsistentFormulas()
    Dim c As Range, n As Long
    For Each c In Selection.Cells
        If c.HasFormula And c.Column > 1 Then
            If c.Offset(0, -1).HasFormula Then
                If c.FormulaR1C1 <> c.Offset(0, -1).FormulaR1C1 Then
                    c.Interior.Color = RGB(255, 199, 206): n = n + 1
                End If
            End If
        End If
    Next
    MsgBox n & " inconsistent formula(s) highlighted.", vbInformation, "Lattice"
End Sub

' Select all cells with errors in the sheet
Public Sub LtxSelectErrors()
    On Error Resume Next
    ActiveSheet.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors).Select
    If Err.Number <> 0 Then MsgBox "No error cells on this sheet.", vbInformation, "Lattice"
End Sub

Public Sub LtxSelectHardcodes()
    On Error Resume Next
    ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers).Select
    If Err.Number <> 0 Then MsgBox "No numeric constants on this sheet.", vbInformation, "Lattice"
End Sub

' Show the formula of the active cell with each reference's value
Public Sub LtxFormulaFlow()
    Dim c As Range, f As String, re As Object, m As Object, s As String, v As Variant, seen As Object
    Set c = ActiveCell
    If Not c.HasFormula Then MsgBox "Active cell has no formula.", vbInformation, "Lattice": Exit Sub
    f = c.Formula
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(?:'[^']+'!|[A-Za-z0-9_]+!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?"
    Set seen = CreateObject("Scripting.Dictionary")
    s = f & vbLf & vbLf
    On Error Resume Next
    For Each m In re.Execute(f)
        If Not seen.Exists(m.Value) Then
            seen(m.Value) = 1
            v = Empty
            v = c.Parent.Evaluate(m.Value)
            If IsArray(v) Then
                s = s & m.Value & " = [range]" & vbLf
            Else
                s = s & m.Value & " = " & v & vbLf
            End If
        End If
    Next
    s = s & vbLf & "Result = " & c.Value
    MsgBox s, vbInformation, "Lattice - Formula Flow"
End Sub

' Toggle showing formulas
Public Sub LtxToggleFormulas()
    ActiveWindow.DisplayFormulas = Not ActiveWindow.DisplayFormulas
End Sub

' Compare two selected sheets' used ranges and colour differences on the active one
Public Sub LtxCompareSheets()
    Dim other As String, ws As Object, c As Range, n As Long
    other = InputBox("Name of the sheet to compare the active sheet against:", "Lattice")
    If Len(other) = 0 Then Exit Sub
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(other)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Sheet not found.", vbExclamation, "Lattice": Exit Sub
    For Each c In ActiveSheet.UsedRange.Cells
        If c.Formula <> ws.Range(c.Address).Formula Then c.Interior.Color = RGB(255, 235, 156): n = n + 1
    Next
    MsgBox n & " differing cell(s) highlighted.", vbInformation, "Lattice"
End Sub
