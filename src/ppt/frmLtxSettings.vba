' Code-behind for frmLtxSettings (controls created by build.ps1 from a spec):
' chk_<Key> checkboxes (Tag = default "1"/"0"), txt_<Key> text fields (Tag = default),
' cbo_<Key> drop-downs (Tag = "option|option;default"). Save writes every control back.
Option Explicit

Private Sub UserForm_Initialize()
    Dim ctl As Object, k As String, tg As String, opts As Variant, i As Long, cur As String
    Me.Caption = "Lattice - Settings"
    LtxThemeForm Me, "Settings"
    On Error Resume Next
    For Each ctl In Me.Controls
        k = Mid$(ctl.Name, 5)
        Select Case Left$(ctl.Name, 4)
            Case "chk_"
                ctl.Value = (LtxGet(k, ctl.Tag) = "1")
            Case "txt_"
                ctl.Text = LtxGet(k, ctl.Tag)
            Case "cbo_"
                tg = ctl.Tag
                opts = Split(Left$(tg, InStr(tg, ";") - 1), "|")
                ctl.Clear
                For i = 0 To UBound(opts)
                    ctl.AddItem opts(i)
                Next
                cur = LtxGet(k, Mid$(tg, InStr(tg, ";") + 1))
                For i = 0 To ctl.ListCount - 1
                    If ctl.List(i) = cur Then ctl.ListIndex = i
                Next
        End Select
    Next
End Sub

Private Sub btnSave_Click()
    Dim ctl As Object, k As String
    On Error Resume Next
    For Each ctl In Me.Controls
        k = Mid$(ctl.Name, 5)
        Select Case Left$(ctl.Name, 4)
            Case "chk_": LtxSet k, IIf(ctl.Value, "1", "0")
            Case "txt_": LtxSet k, ctl.Text
            Case "cbo_": If ctl.ListIndex >= 0 Then LtxSet k, ctl.List(ctl.ListIndex)
        End Select
    Next
    gRibbon.Invalidate
    Unload Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
    If KeyCode = 13 Then btnSave_Click
End Sub
