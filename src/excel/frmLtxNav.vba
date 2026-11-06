' Code-behind for frmLtxNav (controls created by build.ps1)
' Controls: lstRefs (ListBox, 3 columns), btnGo, btnClose, lblInfo
Option Explicit

Private Sub UserForm_Activate()
    Dim i As Long
    LtxTheme "Navigate"
    LtxAddListHeaders Me, lstRefs, Array("ADDRESS", "VALUE", "FORMULA"), Array(0, 150, 240)
    lstRefs.Clear
    If IsEmpty(gNavItems) Then Exit Sub
    For i = 0 To UBound(gNavItems, 1)
        lstRefs.AddItem gNavItems(i, 0)
        lstRefs.List(lstRefs.ListCount - 1, 1) = gNavItems(i, 1)
        lstRefs.List(lstRefs.ListCount - 1, 2) = gNavItems(i, 2)
    Next
    lblInfo.Caption = lstRefs.ListCount & " reference(s)     Up/Down: Excel follows     Enter or double-click: go and close     Esc: close"
    If lstRefs.ListCount > 0 Then lstRefs.ListIndex = 0
    lstRefs.SetFocus
End Sub

Private Sub Jump()
    If lstRefs.ListIndex < 0 Then Exit Sub
    LtxNavJump CStr(lstRefs.List(lstRefs.ListIndex, 0))
End Sub

' A single click (or arrow key) moves Excel to the reference; the window stays open and focused
Private Sub lstRefs_Change()
    If lstRefs.ListIndex < 0 Then Exit Sub
    Jump
    LtxFocusForm Me.Caption
End Sub

Private Sub lstRefs_DblClick(ByVal Cancel As MSForms.ReturnBoolean): Jump: Unload Me: End Sub
Private Sub btnGo_Click(): Jump: Unload Me: End Sub
Private Sub btnClose_Click(): Unload Me: End Sub

Private Sub lstRefs_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0: Jump: Unload Me
    If KeyCode = 27 Then KeyCode = 0: Unload Me
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

' ---- Lattice form theme (white card, navy title, Segoe UI, navy buttons) ----
Private Sub LtxTheme(titleText As String)
    LtxThemeForm Me, titleText
End Sub