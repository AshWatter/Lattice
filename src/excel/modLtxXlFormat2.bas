Attribute VB_Name = "modLtxXlFormat2"
' =====================================================================
'  Lattice for Excel - additional formatting cycles for the Format menu
'  (case, font size, underline, borders, vertical alignment, freeze panes)
' =====================================================================
Option Explicit

Public Sub LtxCaseUpper(): CaseApply 1: End Sub
Public Sub LtxCaseLower(): CaseApply 2: End Sub
Public Sub LtxCaseProper(): CaseApply 3: End Sub

Private Sub CaseApply(mode As Long)
    Dim c As Range
    For Each c In Selection.Cells
        If VarType(c.Value) = vbString And Not c.HasFormula Then
            Select Case mode
                Case 1: c.Value = UCase$(c.Value)
                Case 2: c.Value = LCase$(c.Value)
                Case 3: c.Value = Application.WorksheetFunction.Proper(c.Value)
            End Select
        End If
    Next
End Sub

Public Sub LtxFontGrow(): Selection.Font.Size = Selection.Font.Size + 1: End Sub
Public Sub LtxFontShrink()
    If Selection.Font.Size > 6 Then Selection.Font.Size = Selection.Font.Size - 1
End Sub

' none -> single -> double accounting -> none
Public Sub LtxUnderlineCycle()
    Select Case Selection.Cells(1).Font.Underline
        Case xlUnderlineStyleNone: Selection.Font.Underline = xlUnderlineStyleSingle
        Case xlUnderlineStyleSingle: Selection.Font.Underline = xlUnderlineStyleDoubleAccounting
        Case Else: Selection.Font.Underline = xlUnderlineStyleNone
    End Select
End Sub

Public Sub LtxBorderOutline(): Selection.BorderAround xlContinuous, xlThin: End Sub
Public Sub LtxBorderRemove(): Selection.Borders.LineStyle = xlNone: End Sub
Public Sub LtxBorderBottomCycle()
    With Selection.Borders(xlEdgeBottom)
        If .LineStyle = xlNone Then
            .LineStyle = xlContinuous: .Weight = xlThin
        ElseIf .Weight = xlThin Then
            .Weight = xlMedium
        ElseIf .LineStyle = xlContinuous Then
            .LineStyle = xlDouble
        Else
            .LineStyle = xlNone
        End If
    End With
End Sub
Public Sub LtxBorderTopCycle()
    With Selection.Borders(xlEdgeTop)
        If .LineStyle = xlNone Then .LineStyle = xlContinuous: .Weight = xlThin Else .LineStyle = xlNone
    End With
End Sub

Public Sub LtxVAlignCycle()
    Select Case Selection.VerticalAlignment
        Case xlTop: Selection.VerticalAlignment = xlCenter
        Case xlCenter: Selection.VerticalAlignment = xlBottom
        Case Else: Selection.VerticalAlignment = xlTop
    End Select
End Sub

Public Sub LtxCenterHCycle()
    Select Case Selection.HorizontalAlignment
        Case xlCenter: Selection.HorizontalAlignment = xlCenterAcrossSelection
        Case xlCenterAcrossSelection: Selection.HorizontalAlignment = xlGeneral
        Case Else: Selection.HorizontalAlignment = xlCenter
    End Select
End Sub

Public Sub LtxFreezeToggle()
    With ActiveWindow
        If .FreezePanes Then
            .FreezePanes = False
        Else
            .FreezePanes = True
        End If
    End With
End Sub

Public Sub LtxHideInactiveSheets()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If Not ws Is ActiveSheet Then ws.Visible = xlSheetHidden
    Next
End Sub

Public Sub LtxToggleHeadings(): ActiveWindow.DisplayHeadings = Not ActiveWindow.DisplayHeadings: End Sub

' Negative numbers in red / parentheses toggle on the current number format
Public Sub LtxNegativeStyleCycle()
    Dim f As String
    f = Selection.Cells(1).NumberFormat
    If InStr(f, ";") = 0 Then
        Selection.NumberFormat = f & ";(" & f & ")"
    ElseIf InStr(f, "[Red]") = 0 Then
        Selection.NumberFormat = Replace(f, ";", ";[Red]", 1, 1)
    Else
        Selection.NumberFormat = Split(f, ";")(0)
    End If
End Sub

' Insert a thin check row of "x = y" style formulas is in modLtxXlModel; here: quick sum row
Public Sub LtxSumRow()
    Dim r As Range, c As Long
    Set r = Selection
    For c = 1 To r.Columns.Count
        With r.Cells(r.Rows.Count, c).Offset(1, 0)
            .Formula = "=SUM(" & r.Columns(c).Address(False, False) & ")"
            .Font.Bold = True: .Font.Color = 0
            .Borders(xlEdgeTop).LineStyle = xlContinuous
        End With
    Next
End Sub
