Attribute VB_Name = "modLtxXlNav"
' =====================================================================
'  Lattice for Excel - keyboard formula navigation
'   Ctrl+[            go to precedent  (one press; picker if several)
'   Ctrl+]            go to dependent
'   Ctrl+Shift+Enter  back along the trail
'  Works across sheets and open workbooks via NavigateArrow.
' =====================================================================
Option Explicit

' Used by LtxFocusForm: give the keyboard back to a modeless window after Excel navigated
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long

Private mTrail As Collection      ' external addresses, most recent last
Public gNavItems As Variant       ' 2-D array (address, value, formula) for the picker

Public Function LtxCollectRefs(c As Range, prec As Boolean) As Collection
    Dim items As Collection, a As Long, link As Long, r As Range, seen As Object, k As String
    Set items = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    If prec Then
        If Not c.HasFormula Then Set LtxCollectRefs = items: Exit Function
    End If
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
            k = r.Address(External:=True)
            If k = c.Address(External:=True) Then Exit Do
            If Not seen.Exists(k) Then seen(k) = 1: items.Add k
            link = link + 1
            If link > 300 Then Exit Do
        Loop
        If link = 1 Then Exit Do
        a = a + 1
        If a > 60 Then Exit Do
    Loop
    c.Parent.ClearArrows
    c.Parent.Activate
    c.Select
    On Error GoTo 0
    Application.ScreenUpdating = True
    Set LtxCollectRefs = items
End Function

Private Sub Push(c As Range)
    If mTrail Is Nothing Then Set mTrail = New Collection
    mTrail.Add c.Address(External:=True) & "#" & LtxScrollPos()      ' address plus the scroll position to restore
    If mTrail.Count > 50 Then mTrail.Remove 1
End Sub

Public Sub LtxGoBack()
    Dim addr As String
    If mTrail Is Nothing Then Exit Sub
    If mTrail.Count = 0 Then Application.StatusBar = "Lattice: no earlier cell to return to": Exit Sub
    addr = mTrail(mTrail.Count)
    mTrail.Remove mTrail.Count
    On Error Resume Next
    If InStr(addr, "#") > 0 Then
        LtxJumpTo LtxRangeFromExternal(Left$(addr, InStr(addr, "#") - 1))
        LtxScrollTo Mid$(addr, InStr(addr, "#") + 1)             ' back = the exact view you left
    Else
        LtxJumpTo LtxRangeFromExternal(addr)
    End If
End Sub

Public Sub LtxGoPrecedent(): GoRef True: End Sub
Public Sub LtxGoDependent(): GoRef False: End Sub

Private Sub GoRef(prec As Boolean)
    Dim c As Range, items As Collection, i As Long, arr() As String, r As Range
    Set c = ActiveCell
    Set items = LtxCollectRefs(c, prec)
    If items.Count = 0 Then
        Application.StatusBar = "Lattice: " & IIf(prec, "no precedents", "no dependents") & " for " & c.Address(False, False)
        Exit Sub
    End If
    Push c
    If items.Count = 1 Then
        LtxJumpTo LtxRangeFromExternal(CStr(items(1)))
        Application.StatusBar = "Lattice: " & IIf(prec, "precedent", "dependent") & " of " & c.Address(False, False) & " (Ctrl+Shift+Enter to go back)"
        Exit Sub
    End If
    ReDim arr(0 To items.Count - 1, 0 To 2)
    For i = 1 To items.Count
        Set r = Nothing
        On Error Resume Next
        Set r = LtxRangeFromExternal(CStr(items(i)))
        On Error GoTo 0
        arr(i - 1, 0) = items(i)
        If Not r Is Nothing Then
            arr(i - 1, 1) = CStr(r.Cells(1).Text)
            arr(i - 1, 2) = IIf(r.Cells(1).HasFormula, r.Cells(1).Formula, "")
        End If
    Next
    gNavItems = arr
    frmLtxNav.Caption = "Lattice - " & IIf(prec, "Precedents", "Dependents") & " of " & c.Address(False, False) & "  (Enter = go, Esc = close)"
    frmLtxNav.Show vbModeless
End Sub

' ---------- Trace In / Trace Out windows (Ctrl+Shift+[ / Ctrl+Shift+]) ----------
Public Sub LtxTraceIn(): OpenTrace True: End Sub
Public Sub LtxTraceOut(): OpenTrace False: End Sub

Private Sub OpenTrace(prec As Boolean)
    Dim c As Range
    Set c = ActiveCell
    If prec And Not c.HasFormula Then
        Application.StatusBar = "Lattice: " & c.Address(False, False) & " has no formula to trace"
        Exit Sub
    End If
    On Error Resume Next
    Unload frmLtxTrace
    On Error GoTo 0
    frmLtxTrace.Show vbModeless
    frmLtxTrace.Start c, prec
End Sub

' Called by the picker
Public Sub LtxNavJump(addr As String)
    On Error Resume Next
    LtxJumpTo LtxRangeFromExternal(addr)
End Sub

' ---------- navigation helpers shared by the Trace / Navigate / Model Check windows ----------
' Range from an address in Address(External:=True) form ('[Book.xlsx]Sheet name'!$A$1, [Book]Sheet!A1
' or Sheet!A1), resolved through Workbooks / Worksheets explicitly. Application.Range does not
' accept the [Book] prefix reliably from a UserForm, which left the windows unable to jump.
Public Function LtxRangeFromExternal(addr As String) As Range
    Dim s As String, bookName As String, sheetName As String, cellPart As String, p As Long, wb As Workbook
    On Error Resume Next
    s = addr
    p = InStrRev(s, "!")
    If p = 0 Then Set LtxRangeFromExternal = ActiveSheet.Range(s): Exit Function
    cellPart = Mid$(s, p + 1)
    s = Left$(s, p - 1)
    If Left$(s, 1) = "'" And Right$(s, 1) = "'" Then s = Mid$(s, 2, Len(s) - 2)
    s = Replace(s, "''", "'")
    If Left$(s, 1) = "[" Then
        p = InStr(s, "]")
        bookName = Mid$(s, 2, p - 2): sheetName = Mid$(s, p + 1)
    Else
        sheetName = s
    End If
    If Len(bookName) > 0 Then Set wb = Workbooks(bookName) Else Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function
    Set LtxRangeFromExternal = wb.Worksheets(sheetName).Range(cellPart)
End Function

' Activates the range's workbook and sheet, selects it and scrolls it into view (Capital IQ style),
' respecting frozen panes: the scrollable pane moves, the frozen rows / columns stay. Works from a
' modeless window. False when the sheet is hidden or protected view blocks it.
Public Function LtxJumpTo(r As Range) As Boolean
    On Error Resume Next
    If r Is Nothing Then Exit Function
    If Not r.Parent.Parent Is ActiveWorkbook Then r.Parent.Parent.Activate
    If Not r.Parent Is ActiveSheet Then r.Parent.Activate
    Err.Clear
    r.Cells(1).Select
    If Err.Number <> 0 Then Err.Clear: Application.Goto r, False
    LtxJumpTo = (Err.Number = 0)
    LtxEnsureVisible r.Cells(1)
End Function

' Scrolls the active window so that the cell is on screen with a few rows / columns of context.
' With frozen or split panes only the last pane scrolls (Window.ScrollRow addresses the frozen
' top-left pane, which cannot move), so the pane is scrolled explicitly. Cells inside the frozen
' rows or columns are always visible, so nothing moves for them.
Public Sub LtxEnsureVisible(c As Range)
    Dim win As Object, p As Object, splitR As Long, splitC As Long, topRow As Long, leftCol As Long, vr As Range
    On Error Resume Next
    Application.ScreenUpdating = True
    Set win = ActiveWindow
    If win Is Nothing Then Exit Sub
    For Each p In win.Panes
        Set vr = p.VisibleRange
        If Not vr Is Nothing Then
            ' the last visible row / column is often cut off: require one row / column of slack
            If c.Row >= vr.Row And c.Row < vr.Row + vr.Rows.Count - 1 And c.Column >= vr.Column And c.Column < vr.Column + vr.Columns.Count - 1 Then Exit Sub
        End If
    Next
    If win.FreezePanes Then splitR = win.SplitRow: splitC = win.SplitColumn
    Set p = win.Panes(win.Panes.Count)                 ' the scrollable pane (the window itself when unsplit)
    If c.Row > splitR Then
        topRow = c.Row - 4
        If topRow < splitR + 1 Then topRow = splitR + 1
        p.ScrollRow = topRow
        If p.ScrollRow <> topRow Then win.ScrollRow = topRow
    End If
    If c.Column > splitC Then
        leftCol = c.Column - 2
        If leftCol < splitC + 1 Then leftCol = splitC + 1
        p.ScrollColumn = leftCol
        If p.ScrollColumn <> leftCol Then win.ScrollColumn = leftCol
    End If
End Sub

' Scroll position of the scrollable pane, as "row|column" (used to restore the view on Back)
Public Function LtxScrollPos() As String
    Dim win As Object, p As Object
    On Error Resume Next
    Set win = ActiveWindow
    Set p = win.Panes(win.Panes.Count)
    LtxScrollPos = p.ScrollRow & "|" & p.ScrollColumn
End Function

Public Sub LtxScrollTo(pos As String)
    Dim win As Object, p As Object, parts As Variant
    On Error Resume Next
    If InStr(pos, "|") = 0 Then Exit Sub
    parts = Split(pos, "|")
    Set win = ActiveWindow
    Set p = win.Panes(win.Panes.Count)
    p.ScrollRow = CLng(parts(0)): p.ScrollColumn = CLng(parts(1))
End Sub

' Brings a modeless UserForm (by caption) back to the foreground after a jump, so the user can
' keep moving through the list while Excel follows. Returns quietly when the window is not found.
Public Sub LtxFocusForm(caption As String)
    Dim h As LongPtr
    On Error Resume Next
    h = FindWindowA("ThunderDFrame", caption)
    If h <> 0 Then SetForegroundWindow h
End Sub
' Column headers above a ListBox (ListBoxes cannot show their own), aligned with its column widths.
' Shrinks the list a little to make room. Safe to call once per form.
Public Sub LtxAddListHeaders(frm As Object, lst As Object, caps As Variant, lefts As Variant)
    Dim i As Long, lbl As Object
    On Error Resume Next
    Set lbl = frm.Controls("lblHead0")
    If Not lbl Is Nothing Then Exit Sub
    lst.Top = lst.Top + 14: lst.Height = lst.Height - 14
    For i = LBound(caps) To UBound(caps)
        Set lbl = frm.Controls.Add("Forms.Label.1", "lblHead" & i, True)
        lbl.Left = lst.Left + lefts(i) + 4: lbl.Top = lst.Top - 13: lbl.Width = 150: lbl.Height = 12
        lbl.Caption = caps(i): lbl.Font.Name = "Segoe UI": lbl.Font.Size = 7: lbl.Font.Bold = True
        lbl.ForeColor = RGB(107, 114, 128)
    Next
End Sub
