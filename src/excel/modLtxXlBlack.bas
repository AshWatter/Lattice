Attribute VB_Name = "modLtxXlBlack"
' =====================================================================
'  Lattice for Excel - Black Fonts switch ("presentation mode")
'  Turns every non-light font on the active sheet (or workbook) black and
'  remembers the original colours in a very hidden sheet, so the next
'  click restores blue inputs / green links / red flags exactly.
'  Uses block splitting: uniform blocks are recoloured in one call, so big
'  sheets take seconds, not minutes.
' =====================================================================
Option Explicit

Private Const STORE_SHEET As String = "LatticeColors"

Private Function Store(wb As Workbook, createIt As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(STORE_SHEET)
    On Error GoTo 0
    If ws Is Nothing And createIt Then
        Dim keep As Object, wasUpd As Boolean
        Set keep = wb.ActiveSheet                      ' Worksheets.Add activates the new sheet; put the user back
        wasUpd = Application.ScreenUpdating: Application.ScreenUpdating = False
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = STORE_SHEET
        ws.Range("A1:C1").Value = Array("Sheet", "Range", "Color")
        ws.Visible = xlSheetVeryHidden
        keep.Activate
        Application.ScreenUpdating = wasUpd
    End If
    Set Store = ws
End Function

Private Function IsLight(bgr As Long) As Boolean
    Dim r As Long, g As Long, b As Long
    r = bgr And &HFF: g = (bgr \ &H100) And &HFF: b = (bgr \ &H10000) And &HFF
    IsLight = ((0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.75)
End Function

' Is the sheet currently in black mode? (store has rows for it)
Public Function LtxBlackIsOn(Optional ws As Worksheet) As Boolean
    Dim st As Worksheet
    If ws Is Nothing Then Set ws = ActiveSheet
    Set st = Store(ws.Parent, False)
    If st Is Nothing Then Exit Function
    LtxBlackIsOn = Application.WorksheetFunction.CountIf(st.Columns(1), ws.Name) > 0
End Function

' Block splitting: uniform blocks in one call, mixed blocks halved
Private Sub Blacken(blk As Range, addrs As Collection, cols As Collection, depth As Long)
    Dim col As Variant, h As Long, w As Long
    col = blk.Font.Color
    If Not IsNull(col) Then
        If CLng(col) <> 0 And Not IsLight(CLng(col)) Then
            addrs.Add blk.Address(False, False): cols.Add CLng(col)
            blk.Font.Color = 0
        End If
        Exit Sub
    End If
    h = blk.Rows.Count: w = blk.Columns.Count
    If (h = 1 And w = 1) Or depth > 40 Then Exit Sub
    If h >= w Then
        Blacken blk.Rows(1).Resize(h \ 2), addrs, cols, depth + 1
        Blacken blk.Rows(h \ 2 + 1).Resize(h - h \ 2), addrs, cols, depth + 1
    Else
        Blacken blk.Columns(1).Resize(, w \ 2), addrs, cols, depth + 1
        Blacken blk.Columns(w \ 2 + 1).Resize(, w - w \ 2), addrs, cols, depth + 1
    End If
End Sub

Private Sub ApplyBlackRange(ws As Worksheet, target As Range)
    Dim st As Worksheet, r As Long, addrs As Collection, cols As Collection, k As Long
    If ws.Name = STORE_SHEET Or ws.ProtectContents Or target Is Nothing Then Exit Sub
    Set addrs = New Collection: Set cols = New Collection
    Blacken target, addrs, cols, 0
    If addrs.Count = 0 Then Exit Sub
    Set st = Store(ws.Parent, True)
    r = st.Cells(st.Rows.Count, 1).End(xlUp).Row + 1
    For k = 1 To addrs.Count
        st.Cells(r, 1).Value = ws.Name: st.Cells(r, 2).Value = addrs(k): st.Cells(r, 3).Value = cols(k)
        r = r + 1
    Next
End Sub

Private Sub ApplyBlack(ws As Worksheet)
    ApplyBlackRange ws, ws.UsedRange
End Sub

Private Sub RestoreColors(ws As Worksheet, Optional cleanup As Boolean = True)
    Dim st As Worksheet, i As Long, last As Long
    Set st = Store(ws.Parent, False)
    If st Is Nothing Then Exit Sub
    last = st.Cells(st.Rows.Count, 1).End(xlUp).Row
    On Error Resume Next
    For i = last To 2 Step -1
        If st.Cells(i, 1).Value = ws.Name Then
            ws.Range(st.Cells(i, 2).Value).Font.Color = CLng(st.Cells(i, 3).Value)
            st.Rows(i).Delete
        End If
    Next
    On Error GoTo 0
    If cleanup Then DropStoreIfEmpty ws.Parent
End Sub

Private Sub DropStoreIfEmpty(wb As Workbook)
    Dim st As Worksheet
    Set st = Store(wb, False)
    If st Is Nothing Then Exit Sub
    If st.Cells(st.Rows.Count, 1).End(xlUp).Row <= 1 Then
        Application.DisplayAlerts = False
        st.Visible = xlSheetVisible
        st.Delete
        Application.DisplayAlerts = True
    End If
End Sub

Private Sub Fast(onOff As Boolean)
    Static calc As Long
    If onOff Then
        calc = Application.Calculation
        Application.ScreenUpdating = False: Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = IIf(calc = 0, xlCalculationAutomatic, calc)
        Application.EnableEvents = True: Application.ScreenUpdating = True
    End If
End Sub

' Toggle for the active sheet (Ctrl+Shift+Z)
Public Sub LtxBlackToggle()
    Dim saved As Boolean
    saved = ActiveWorkbook.Saved
    Fast True
    If LtxBlackIsOn(ActiveSheet) Then
        RestoreColors ActiveSheet
        Application.StatusBar = "Lattice: colours restored on " & ActiveSheet.Name
    Else
        ApplyBlack ActiveSheet
        Application.StatusBar = "Lattice: black fonts on " & ActiveSheet.Name & " (click again to restore colours)"
    End If
    Fast False
    On Error Resume Next
    gRibbon.Invalidate
End Sub

' Toggle for every sheet in the workbook
Public Sub LtxBlackToggleWorkbook()
    Dim ws As Worksheet, turnOn As Boolean
    turnOn = Not LtxBlackIsOn(ActiveSheet)
    Fast True
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> STORE_SHEET Then
            Application.StatusBar = "Lattice: " & IIf(turnOn, "blackening ", "restoring ") & ws.Name
            If turnOn Then
                If Not LtxBlackIsOn(ws) Then ApplyBlack ws
            Else
                If LtxBlackIsOn(ws) Then RestoreColors ws, False
            End If
        End If
    Next
    If Not turnOn Then DropStoreIfEmpty ActiveWorkbook
    Fast False
    Application.StatusBar = "Lattice: black fonts " & IIf(turnOn, "ON", "OFF") & " for the whole workbook"
    On Error Resume Next
    gRibbon.Invalidate
End Sub

' Selection only: black now, restore with the same button on the same selection
Public Sub LtxBlackSelection()
    Dim st As Worksheet, anyOn As Boolean, i As Long, last As Long
    Set st = Store(ActiveWorkbook, False)
    If Not st Is Nothing Then
        last = st.Cells(st.Rows.Count, 1).End(xlUp).Row
        For i = last To 2 Step -1
            If st.Cells(i, 1).Value = ActiveSheet.Name Then
                If Not Application.Intersect(Selection, ActiveSheet.Range(st.Cells(i, 2).Value)) Is Nothing Then
                    ActiveSheet.Range(st.Cells(i, 2).Value).Font.Color = CLng(st.Cells(i, 3).Value)
                    st.Rows(i).Delete: anyOn = True
                End If
            End If
        Next
        If anyOn Then DropStoreIfEmpty ActiveWorkbook: Exit Sub
    End If
    Fast True
    ApplyBlackRange ActiveSheet, Selection
    Fast False
End Sub
