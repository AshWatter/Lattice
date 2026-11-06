' Code-behind for frmLtxTrace (controls created by build.ps1)
' Controls: lstTrace (ListBox, 3 columns: reference | value | formula), btnExpand, btnCollapse, btnGo, btnEdit, btnClose, lblInfo, lblTitle
' Keys inside the window: Up/Down navigate Excel live, Right expands, Left collapses / goes to parent,
' F2 edits the formula in Excel, Ctrl+E toggles value / formula view, Esc closes.
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
#End If

' node arrays: 0 addr (external), 1 depth, 2 expanded (0/1), 3 hasChildren (-1 unknown / 0 / 1)
Private mNodes As Collection
Private mPrec As Boolean
Private mShowFormula As Boolean
Private mBusy As Boolean

Public Sub Start(root As Range, prec As Boolean)
    mPrec = prec
    Set mNodes = New Collection
    mNodes.Add Array(root.Address(External:=True), 0, 0, -1)
    Me.Caption = "Lattice - " & IIf(prec, "Trace In (precedents)", "Trace Out (dependents)") & "   Up/Down go  Right expand  Left back  F2 edit  Ctrl+E value/formula  Esc close"
    LtxTheme IIf(prec, "Trace In", "Trace Out") & "  -  " & root.Address(False, False)
    Expand 1
    Render
    If lstTrace.ListCount > 1 Then lstTrace.ListIndex = 1 Else lstTrace.ListIndex = 0
End Sub

Private Sub UserForm_Initialize()
    lblInfo.Caption = "Up/Down: Excel follows   Right / Left: expand / collapse   Enter: go there and close (then F2)   F2: go and edit   Ctrl+E: value / formula   Esc: close"
    AddHeaders
End Sub

Private Function NodeRange(i As Long) As Range
    On Error Resume Next
    Set NodeRange = LtxRangeFromExternal(CStr(mNodes(i)(0)))
End Function

Private Sub Expand(i As Long)
    Dim r As Range, refs As Collection, k As Long, node As Variant, child As Variant
    node = mNodes(i)
    If node(2) = 1 Then Exit Sub
    Set r = NodeRange(i)
    If r Is Nothing Then Exit Sub
    Set refs = LtxCollectRefs(r.Cells(1), mPrec)
    node(2) = 1: node(3) = IIf(refs.Count > 0, 1, 0)
    mNodes.Remove i
    If i > mNodes.Count Then mNodes.Add node Else mNodes.Add node, , i
    For k = refs.Count To 1 Step -1
        child = Array(refs(k), node(1) + 1, 0, -1)
        If i + 1 > mNodes.Count Then mNodes.Add child Else mNodes.Add child, , i + 1
    Next
End Sub

Private Sub Collapse(i As Long)
    Dim node As Variant, d As Long
    node = mNodes(i)
    d = node(1)
    Do While i < mNodes.Count
        If mNodes(i + 1)(1) > d Then mNodes.Remove i + 1 Else Exit Do
    Loop
    node(2) = 0
    mNodes.Remove i
    If i > mNodes.Count Then mNodes.Add node Else mNodes.Add node, , i
End Sub

Private Function ParentIndex(i As Long) As Long
    Dim d As Long, j As Long
    d = mNodes(i)(1)
    For j = i - 1 To 1 Step -1
        If mNodes(j)(1) < d Then ParentIndex = j: Exit Function
    Next
    ParentIndex = 1
End Function

Private Sub Render()
    Dim i As Long, node As Variant, r As Range, disp As String, v As String, f As String, keep As Long
    keep = lstTrace.ListIndex
    mBusy = True
    lstTrace.Clear
    For i = 1 To mNodes.Count
        node = mNodes(i)
        Set r = NodeRange(i)
        disp = String$(node(1) * 3, " ") & IIf(node(2) = 1, ChrW(9662) & " ", IIf(node(3) = 0, "  ", ChrW(9656) & " ")) & ShortAddr(CStr(node(0)))
        v = "": f = ""
        If Not r Is Nothing Then
            If r.Cells.Count > 1 Then v = r.Cells(1).Text & "  (" & r.Cells.Count & " cells)" Else v = r.Text
            If r.Cells(1).HasFormula Then f = r.Cells(1).Formula
        Else
            v = "(not available)"
        End If
        lstTrace.AddItem disp
        lstTrace.List(lstTrace.ListCount - 1, 1) = IIf(mShowFormula And Len(f) > 0, f, v)
        lstTrace.List(lstTrace.ListCount - 1, 2) = IIf(mShowFormula, v, f)
    Next
    lblInfo.Caption = (mNodes.Count - 1) & " reference(s)   |   columns: reference | " & IIf(mShowFormula, "formula | value", "value | formula")
    mBusy = False
    If keep >= 0 And keep < lstTrace.ListCount Then lstTrace.ListIndex = keep
End Sub

Private Function ShortAddr(a As String) As String
    ' drop the [Workbook] part when it is the active workbook
    Dim s As String
    s = a
    If InStr(s, "[" & ActiveWorkbook.Name & "]") > 0 Then s = Replace(s, "[" & ActiveWorkbook.Name & "]", "")
    If Left$(s, 1) = "'" And InStr(s, "'!") > 0 Then
        If InStr(s, " ") = 0 Then s = Replace(Replace(s, "'", ""), "!", "!")
    End If
    ShortAddr = s
End Function

Private Sub GoTo_(i As Long)
    Dim r As Range
    If i < 1 Or i > mNodes.Count Then Exit Sub
    Set r = NodeRange(i)
    If r Is Nothing Then Exit Sub
    On Error Resume Next
    If Not LtxJumpTo(r) Then lblInfo.Caption = "Cannot show " & ShortAddr(CStr(mNodes(i)(0))) & " (hidden sheet?)"
    KeepFocus
End Sub

Private Sub KeepFocus()
    #If VBA7 Then
        Dim h As LongPtr
    #Else
        Dim h As Long
    #End If
    On Error Resume Next
    h = FindWindowA("ThunderDFrame", Me.Caption)
    If h <> 0 Then SetForegroundWindow h
    lstTrace.SetFocus
End Sub

Private Sub ShowFormula(i As Long)
    Dim r As Range
    On Error Resume Next
    Set r = NodeRange(i)
    If r Is Nothing Then txtFormula.Text = "": Exit Sub
    txtFormula.Text = ShortAddr(CStr(mNodes(i)(0))) & ":   " & IIf(r.Cells(1).HasFormula, r.Cells(1).Formula, "(input)  " & r.Cells(1).Text)
End Sub

Private Sub lstTrace_Change()
    If mBusy Then Exit Sub
    If lstTrace.ListIndex >= 0 Then ShowFormula lstTrace.ListIndex + 1: GoTo_ lstTrace.ListIndex + 1
End Sub

Private Sub lstTrace_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i As Long
    i = lstTrace.ListIndex + 1
    Select Case KeyCode
        Case 39   ' Right: expand
            KeyCode = 0
            If i >= 1 Then Expand i: Render: GoTo_ i
        Case 37   ' Left: collapse, or go to parent
            KeyCode = 0
            If i >= 1 Then
                If mNodes(i)(2) = 1 Then
                    Collapse i: Render
                Else
                    lstTrace.ListIndex = ParentIndex(i) - 1
                End If
            End If
        Case 113  ' F2: go to the cell, close the window, start editing
            KeyCode = 0
            GoTo_ i
            CloseToExcel
            Application.SendKeys "{F2}", False
        Case 69   ' Ctrl+E: toggle value / formula
            If Shift = 2 Then KeyCode = 0: mShowFormula = Not mShowFormula: Render
        Case 13   ' Enter: go to the cell and close, so F2 / typing works at once
            KeyCode = 0
            GoTo_ i
            CloseToExcel
        Case 27
            KeyCode = 0: Unload Me
    End Select
End Sub

Private Sub btnExpand_Click()
    If lstTrace.ListIndex >= 0 Then Expand lstTrace.ListIndex + 1: Render
End Sub
Private Sub btnCollapse_Click()
    If lstTrace.ListIndex >= 0 Then Collapse lstTrace.ListIndex + 1: Render
End Sub
Private Sub btnGo_Click()
    If lstTrace.ListIndex >= 0 Then GoTo_ lstTrace.ListIndex + 1: CloseToExcel
End Sub

' Hands the keyboard back to Excel on the cell just selected and closes the window
Private Sub CloseToExcel()
    On Error Resume Next
    Unload Me
    AppActivate Application.Caption
End Sub
Private Sub btnEdit_Click()
    If lstTrace.ListIndex < 0 Then Exit Sub
    GoTo_ lstTrace.ListIndex + 1
    CloseToExcel
    Application.SendKeys "{F2}", False
End Sub
Private Sub btnClose_Click(): Unload Me: End Sub
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

' ---- Lattice form theme (white card, navy title, Segoe UI, navy buttons) ----
Private Sub LtxTheme(titleText As String)
    LtxThemeForm Me, titleText
End Sub

' Column headers above the list (the ListBox itself cannot show them), aligned with the column widths
Private Sub AddHeaders()
    Dim cap As Variant, lefts As Variant, i As Long, lbl As Object
    On Error Resume Next
    cap = Array("REFERENCE", "VALUE", "FORMULA")
    lefts = Array(0, 200, 310)
    lstTrace.Top = lstTrace.Top + 14: lstTrace.Height = lstTrace.Height - 14
    For i = 0 To 2
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblHead" & i, True)
        lbl.Left = lstTrace.Left + lefts(i) + 4: lbl.Top = lstTrace.Top - 13: lbl.Width = 150: lbl.Height = 12
        lbl.Caption = cap(i): lbl.Font.Name = "Segoe UI": lbl.Font.Size = 7: lbl.Font.Bold = True
        lbl.ForeColor = RGB(107, 114, 128)
    Next
End Sub
