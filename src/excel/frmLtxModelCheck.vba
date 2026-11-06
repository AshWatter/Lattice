' Code-behind for frmLtxModelCheck (controls created by build.ps1)
' Controls: lstIssues (ListBox, 4 columns), btnSheet, btnBook, btnGo, btnFix, btnReport, btnClose, lblInfo
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Lattice - Model Check"
    LtxTheme "Model Check"
    lblInfo.Caption = "Analyze the active sheet or the whole workbook. Click an issue to jump to it; arrow keys walk the list while Excel follows."
    LtxAddListHeaders Me, lstIssues, Array("SHEET", "CELL", "ISSUE", "DETAIL"), Array(0, 90, 150, 300)
    On Error Resume Next
    Fill False                     ' start with the active sheet analysed
    On Error GoTo 0
End Sub

Private Sub Fill(wholeBook As Boolean)
    Dim t As Variant, i As Long
    lstIssues.Clear
    t = LtxCollectIssues(wholeBook)
    If IsEmpty(t) Then
        lblInfo.Caption = "No issues found."
        Exit Sub
    End If
    For i = 0 To UBound(t, 1)
        lstIssues.AddItem t(i, 0)
        lstIssues.List(lstIssues.ListCount - 1, 1) = t(i, 1)
        lstIssues.List(lstIssues.ListCount - 1, 2) = t(i, 2)
        lstIssues.List(lstIssues.ListCount - 1, 3) = t(i, 3)
    Next
    lblInfo.Caption = (UBound(t, 1) + 1) & " issue(s). Columns: sheet | cell | issue | detail"
End Sub

Private Sub GoCurrent()
    Dim r As Range
    If lstIssues.ListIndex < 0 Then Exit Sub
    On Error Resume Next
    Set r = ActiveWorkbook.Worksheets(CStr(lstIssues.List(lstIssues.ListIndex, 0))).Range(CStr(lstIssues.List(lstIssues.ListIndex, 1)))
    If r Is Nothing Then lblInfo.Caption = "Cannot find " & lstIssues.List(lstIssues.ListIndex, 0) & "!" & lstIssues.List(lstIssues.ListIndex, 1): Exit Sub
    If Not LtxJumpTo(r) Then lblInfo.Caption = "Cannot show " & r.Address(False, False) & " (hidden sheet?)"
End Sub

' A single click (or arrow key) jumps; the pane keeps the keyboard so you can walk the list
Private Sub lstIssues_Change()
    If lstIssues.ListIndex < 0 Then Exit Sub
    GoCurrent
    LtxFocusForm Me.Caption
End Sub

Private Sub btnSheet_Click(): Fill False: End Sub
Private Sub btnBook_Click(): Fill True: End Sub
Private Sub btnGo_Click(): GoCurrent: End Sub
Private Sub lstIssues_DblClick(ByVal Cancel As MSForms.ReturnBoolean): GoCurrent: End Sub
Private Sub btnReport_Click(): LtxAuditHardcodes True: End Sub
Private Sub btnFix_Click()
    GoCurrent
    If lstIssues.ListIndex >= 0 Then
        If lstIssues.List(lstIssues.ListIndex, 2) = "Hard-code in formula" Then LtxExtractHardcode
    End If
End Sub
Private Sub btnClose_Click(): Unload Me: End Sub

' ---- Lattice form theme (white card, navy title, Segoe UI, navy buttons) ----
Private Sub LtxTheme(titleText As String)
    LtxThemeForm Me, titleText
End Sub