Attribute VB_Name = "modLtxXlSwitch"
' =====================================================================
'  Lattice for Excel - case switch
'  Inserts a bare Forms drop-down (the classic "Downside Case" selector) at the active cell and
'  nothing else: no helper sheet, no names. Wire it with right-click > Format Control
'  (Input range = the cells holding the case names, Cell link = the cell that receives the
'  selected number, 1 = first case). If the selection already holds the case names, the input
'  range is pre-filled and the control is placed to their right. No macro is involved, so copies
'  on any sheet keep working.
' =====================================================================
Option Explicit

Public Sub LtxInsertCaseSwitch()
    Dim ws As Worksheet, anchor As Range, dd As Shape, listRng As Range
    On Error GoTo fail
    If TypeName(Selection) <> "Range" Then MsgBox "Select a cell first.", vbInformation, "Lattice": Exit Sub
    Set ws = ActiveSheet
    If Selection.Cells.Count >= 2 And Application.WorksheetFunction.CountA(Selection) >= 2 Then
        Set listRng = Selection
        Set anchor = listRng.Cells(1).Offset(0, listRng.Columns.Count + 1)
    Else
        Set anchor = ActiveCell
    End If
    Set dd = ws.Shapes.AddFormControl(xlDropDown, anchor.Left, anchor.Top, 150, 20)
    dd.Name = "LtxCaseSwitch_" & Format$(Now, "hhnnss")
    dd.ControlFormat.DropDownLines = 8
    dd.Placement = xlFreeFloating
    If Not listRng Is Nothing Then dd.ControlFormat.ListFillRange = "'" & listRng.Parent.Name & "'!" & listRng.Address(True, True)
    dd.Select
    Application.StatusBar = "Lattice: case drop-down inserted. Right-click > Format Control to set Input range (case names) and Cell link (selected number)."
    Exit Sub
fail:
    MsgBox "Case switch: " & Err.Description, vbExclamation, "Lattice"
End Sub
