' Code-behind for frmLtxExplorer (controls are created by build.ps1)
' Controls: lstLinks (ListBox, 6 columns), btnRefresh, btnRefreshAll,
'           btnGoTo, btnSource, btnBreak, btnReload, btnClose, lblInfo
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Lattice - Link Explorer"
    LtxTheme "Link Explorer"
    Reload
End Sub

Private Sub Reload()
    Dim t As Variant, i As Long
    lstLinks.Clear
    t = LtxLinkTable()
    If IsEmpty(t) Then
        lblInfo.Caption = "No Lattice links in this deck."
        Exit Sub
    End If
    For i = 0 To UBound(t, 1)
        lstLinks.AddItem t(i, 0)
        lstLinks.List(lstLinks.ListCount - 1, 1) = t(i, 1)
        lstLinks.List(lstLinks.ListCount - 1, 2) = t(i, 2)
        lstLinks.List(lstLinks.ListCount - 1, 3) = t(i, 3)
        lstLinks.List(lstLinks.ListCount - 1, 4) = t(i, 4)
        lstLinks.List(lstLinks.ListCount - 1, 5) = t(i, 5)
    Next
    lblInfo.Caption = (UBound(t, 1) + 1) & " link(s). Columns: slide | object | workbook | sheet | range | updated"
End Sub

Private Sub SelectCurrent()
    If lstLinks.ListIndex < 0 Then Exit Sub
    LtxSelectLink CLng(lstLinks.List(lstLinks.ListIndex, 0)), CStr(lstLinks.List(lstLinks.ListIndex, 1))
End Sub

Private Sub lstLinks_Click()
    SelectCurrent
End Sub

Private Sub lstLinks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SelectCurrent
    LtxGoToSource
End Sub

Private Sub btnRefresh_Click()
    SelectCurrent
    LtxRefreshSelected
    Reload
End Sub

Private Sub btnRefreshAll_Click()
    LtxRefreshAll
    Reload
End Sub

Private Sub btnGoTo_Click()
    SelectCurrent
    LtxGoToSource
End Sub

Private Sub btnSource_Click()
    SelectCurrent
    LtxChangeSource
    Reload
End Sub

Private Sub btnBreak_Click()
    SelectCurrent
    LtxBreakLink
    Reload
End Sub

Private Sub btnReload_Click()
    Reload
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' ---- Lattice form theme (white card, navy title, Segoe UI, navy buttons) ----
Private Sub LtxTheme(titleText As String)
    LtxThemeForm Me, titleText
End Sub