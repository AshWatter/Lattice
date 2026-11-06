Attribute VB_Name = "modLtxShapes"
' =====================================================================
'  Lattice for PowerPoint - shape alignment, sizing, arranging, selection
'  Convention (same as the big add-ins): the FIRST selected shape is the
'  anchor; everything else aligns / resizes to it. With one shape
'  selected, alignment is relative to the slide.
' =====================================================================
Option Explicit

Private Const msoTrue = -1
Private Const msoFalse = 0

Public Function LtxSel() As Object
    On Error Resume Next
    If ActiveWindow.Selection.Type = 2 Or ActiveWindow.Selection.Type = 3 Then
        Set LtxSel = ActiveWindow.Selection.ShapeRange
    End If
End Function

Private Function NeedSel(minCount As Long) As Object
    Dim sr As Object
    Set sr = LtxSel()
    If sr Is Nothing Then
        MsgBox "Select at least " & minCount & " shape(s).", vbInformation, LTX_APP
    ElseIf sr.Count < minCount Then
        MsgBox "Select at least " & minCount & " shape(s).", vbInformation, LTX_APP
        Set sr = Nothing
    End If
    Set NeedSel = sr
End Function

Private Function SlideW() As Single: SlideW = ActivePresentation.PageSetup.SlideWidth: End Function
Private Function SlideH() As Single: SlideH = ActivePresentation.PageSetup.SlideHeight: End Function

' ---------------- Align (to first shape, or slide if single) ----------------
Public Sub LtxAlign(edge As String)
    Dim sr As Object, a As Object, s As Object, i As Long
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    If sr.Count = 1 Then
        Set s = sr(1)
        Select Case edge
            Case "left":    s.Left = 0
            Case "right":   s.Left = SlideW() - s.Width
            Case "top":     s.Top = 0
            Case "bottom":  s.Top = SlideH() - s.Height
            Case "center":  s.Left = (SlideW() - s.Width) / 2
            Case "middle":  s.Top = (SlideH() - s.Height) / 2
        End Select
        Exit Sub
    End If
    Set a = sr(1)
    For i = 2 To sr.Count
        Set s = sr(i)
        Select Case edge
            Case "left":    s.Left = a.Left
            Case "right":   s.Left = a.Left + a.Width - s.Width
            Case "top":     s.Top = a.Top
            Case "bottom":  s.Top = a.Top + a.Height - s.Height
            Case "center":  s.Left = a.Left + (a.Width - s.Width) / 2
            Case "middle":  s.Top = a.Top + (a.Height - s.Height) / 2
        End Select
    Next
End Sub

' Center the whole selection on the slide (keeps relative layout)
Public Sub LtxCenterOnSlide(axis As String)
    Dim sr As Object, L As Single, T As Single, R As Single, B As Single, s As Object, dx As Single, dy As Single
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    Bounds sr, L, T, R, B
    dx = (SlideW() - (R - L)) / 2 - L
    dy = (SlideH() - (B - T)) / 2 - T
    For Each s In sr
        If axis <> "v" Then s.Left = s.Left + dx
        If axis <> "h" Then s.Top = s.Top + dy
    Next
End Sub

Private Sub Bounds(sr As Object, ByRef L As Single, ByRef T As Single, ByRef R As Single, ByRef B As Single)
    Dim s As Object
    L = 1E+9: T = 1E+9: R = -1E+9: B = -1E+9
    For Each s In sr
        If s.Left < L Then L = s.Left
        If s.Top < T Then T = s.Top
        If s.Left + s.Width > R Then R = s.Left + s.Width
        If s.Top + s.Height > B Then B = s.Top + s.Height
    Next
End Sub

' ---------------- Distribute ----------------
' Equal gaps between shapes, keeping the outermost ones in place.
Public Sub LtxDistribute(axis As String)
    Dim sr As Object, n As Long, i As Long, j As Long, arr() As Object, tmp As Object
    Dim total As Single, L As Single, T As Single, R As Single, B As Single, gap As Single, pos As Single
    Set sr = NeedSel(3)
    If sr Is Nothing Then Exit Sub
    n = sr.Count
    ReDim arr(1 To n)
    For i = 1 To n: Set arr(i) = sr(i): Next
    ' sort by position
    For i = 1 To n - 1
        For j = i + 1 To n
            If IIf(axis = "h", arr(j).Left < arr(i).Left, arr(j).Top < arr(i).Top) Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next
    Next
    Bounds sr, L, T, R, B
    For i = 1 To n: total = total + IIf(axis = "h", arr(i).Width, arr(i).Height): Next
    gap = (IIf(axis = "h", R - L, B - T) - total) / (n - 1)
    pos = IIf(axis = "h", L, T)
    For i = 1 To n
        If axis = "h" Then arr(i).Left = pos: pos = pos + arr(i).Width + gap _
        Else arr(i).Top = pos: pos = pos + arr(i).Height + gap
    Next
End Sub

' Set a fixed gap between shapes (in points) in reading order
Public Sub LtxSetGap(axis As String)
    Dim sr As Object, g As String, gap As Single, n As Long, i As Long, j As Long, arr() As Object, tmp As Object, pos As Single
    Set sr = NeedSel(2)
    If sr Is Nothing Then Exit Sub
    g = InputBox("Gap between shapes (points, 72 = 1 inch):", LTX_APP, LtxGet("Gap", "8"))
    If Len(g) = 0 Then Exit Sub
    gap = Val(g): LtxSet "Gap", g
    n = sr.Count: ReDim arr(1 To n)
    For i = 1 To n: Set arr(i) = sr(i): Next
    For i = 1 To n - 1
        For j = i + 1 To n
            If IIf(axis = "h", arr(j).Left < arr(i).Left, arr(j).Top < arr(i).Top) Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next
    Next
    pos = IIf(axis = "h", arr(1).Left, arr(1).Top)
    For i = 1 To n
        If axis = "h" Then arr(i).Left = pos: pos = pos + arr(i).Width + gap _
        Else arr(i).Top = pos: pos = pos + arr(i).Height + gap
    Next
End Sub

' ---------------- Size ----------------
Public Sub LtxMatchSize(what As String)
    Dim sr As Object, a As Object, i As Long
    Set sr = NeedSel(2)
    If sr Is Nothing Then Exit Sub
    Set a = sr(1)
    For i = 2 To sr.Count
        With sr(i)
            Select Case what
                Case "width":  .LockAspectRatio = msoFalse: .Width = a.Width
                Case "height": .LockAspectRatio = msoFalse: .Height = a.Height
                Case "both":   .LockAspectRatio = msoFalse: .Width = a.Width: .Height = a.Height
                Case "widthlock": .LockAspectRatio = msoTrue: .Width = a.Width
                Case "heightlock": .LockAspectRatio = msoTrue: .Height = a.Height
            End Select
        End With
    Next
End Sub

' Stretch the other shapes so their edges reach the anchor's edge
Public Sub LtxStretch(edge As String)
    Dim sr As Object, a As Object, s As Object, i As Long, r As Single
    Set sr = NeedSel(2)
    If sr Is Nothing Then Exit Sub
    Set a = sr(1)
    For i = 2 To sr.Count
        Set s = sr(i): s.LockAspectRatio = msoFalse
        Select Case edge
            Case "left":   r = s.Left + s.Width: s.Left = a.Left: s.Width = r - a.Left
            Case "right":  s.Width = (a.Left + a.Width) - s.Left
            Case "top":    r = s.Top + s.Height: s.Top = a.Top: s.Height = r - a.Top
            Case "bottom": s.Height = (a.Top + a.Height) - s.Top
        End Select
    Next
End Sub

Public Sub LtxMakeSquare()
    Dim sr As Object, s As Object, m As Single
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    For Each s In sr
        m = IIf(s.Width > s.Height, s.Width, s.Height)
        s.LockAspectRatio = msoFalse: s.Width = m: s.Height = m
    Next
End Sub

Public Sub LtxFitToSlide()
    Dim sr As Object, s As Object, m As Single
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    m = Val(LtxGet("MarginPt", "36"))
    For Each s In sr
        s.LockAspectRatio = msoFalse
        s.Left = m: s.Top = m: s.Width = SlideW() - 2 * m: s.Height = SlideH() - 2 * m
    Next
End Sub

Public Sub LtxSwap()
    Dim sr As Object, L As Single, T As Single
    Set sr = NeedSel(2)
    If sr Is Nothing Then Exit Sub
    L = sr(1).Left: T = sr(1).Top
    sr(1).Left = sr(2).Left: sr(1).Top = sr(2).Top
    sr(2).Left = L: sr(2).Top = T
End Sub

Public Sub LtxNudge(dx As Single, dy As Single)
    Dim sr As Object, s As Object
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    For Each s In sr: s.Left = s.Left + dx: s.Top = s.Top + dy: Next
End Sub

' ---------------- Arrange shapes in a grid ----------------
Public Sub LtxArrangeGrid()
    Dim sr As Object, n As Long, cols As Long, gap As Single, i As Long, gr As Long, gc As Long
    Dim L As Single, T As Single, R As Single, B As Single, cw As Single, ch As Single, inp As String
    Set sr = NeedSel(2)
    If sr Is Nothing Then Exit Sub
    n = sr.Count
    inp = InputBox("Number of columns:", LTX_APP, CStr(Int(Sqr(n) + 0.999)))
    If Len(inp) = 0 Then Exit Sub
    cols = Val(inp): If cols < 1 Then cols = 1
    gap = Val(LtxGet("Gap", "8"))
    Bounds sr, L, T, R, B
    cw = 0: ch = 0
    For i = 1 To n
        If sr(i).Width > cw Then cw = sr(i).Width
        If sr(i).Height > ch Then ch = sr(i).Height
    Next
    For i = 1 To n
        gr = (i - 1) \ cols: gc = (i - 1) Mod cols
        sr(i).Left = L + gc * (cw + gap)
        sr(i).Top = T + gr * (ch + gap)
    Next
End Sub

' ---------------- Auto Align ----------------
' Snaps near-equal lefts, tops, widths and heights (within a tolerance) to a
' common value, so shapes that are "almost" aligned become exactly aligned.
Public Sub LtxAutoAlign()
    Dim sr As Object, tol As Single, n As Long, i As Long, j As Long, k As Long, s As Object
    Dim props As Variant, vals() As Single, p As Long, grp As Long
    Set sr = NeedSel(2)
    If sr Is Nothing Then Exit Sub
    tol = Val(LtxGet("AutoAlignTol", "6"))
    n = sr.Count
    props = Array("Left", "Top", "Width", "Height")
    For p = 0 To 3
        ReDim vals(1 To n)
        For i = 1 To n: vals(i) = CallByName(sr(i), CStr(props(p)), 2): Next
        For i = 1 To n
            For j = i + 1 To n
                If Abs(vals(j) - vals(i)) <= tol And vals(j) <> vals(i) Then
                    If p >= 2 Then sr(j).LockAspectRatio = msoFalse
                    CallByName sr(j), CStr(props(p)), 4, vals(i)
                    vals(j) = vals(i)
                End If
            Next
        Next
    Next
    ' right edges and bottom edges too
    For i = 1 To n
        For j = i + 1 To n
            If Abs((sr(j).Left + sr(j).Width) - (sr(i).Left + sr(i).Width)) <= tol And (sr(j).Left + sr(j).Width) <> (sr(i).Left + sr(i).Width) Then
                sr(j).LockAspectRatio = msoFalse: sr(j).Width = sr(i).Left + sr(i).Width - sr(j).Left
            End If
            If Abs((sr(j).Top + sr(j).Height) - (sr(i).Top + sr(i).Height)) <= tol And (sr(j).Top + sr(j).Height) <> (sr(i).Top + sr(i).Height) Then
                sr(j).LockAspectRatio = msoFalse: sr(j).Height = sr(i).Top + sr(i).Height - sr(j).Top
            End If
        Next
    Next
End Sub

' ---------------- Align to the template body area ----------------
' Uses the body placeholder of the slide's layout as the reference box.
Private Function BodyBox(sld As Object, ByRef L As Single, ByRef T As Single, ByRef W As Single, ByRef H As Single) As Boolean
    Dim shp As Object, m As Single
    On Error Resume Next
    For Each shp In sld.CustomLayout.Shapes
        If shp.Type = 14 Then
            If shp.PlaceholderFormat.Type = 2 Or shp.PlaceholderFormat.Type = 7 Then   ' body / object
                L = shp.Left: T = shp.Top: W = shp.Width: H = shp.Height: BodyBox = True: Exit Function
            End If
        End If
    Next
    m = Val(LtxGet("MarginPt", "36"))
    L = m: T = m * 2.5: W = SlideW() - 2 * m: H = SlideH() - T - m
End Function

Public Sub LtxAlignToBody(edge As String)
    Dim sr As Object, s As Object, L As Single, T As Single, W As Single, H As Single
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    BodyBox LtxActiveSlide(), L, T, W, H
    For Each s In sr
        Select Case edge
            Case "left":   s.Left = L
            Case "right":  s.Left = L + W - s.Width
            Case "top":    s.Top = T
            Case "bottom": s.Top = T + H - s.Height
            Case "center": s.Left = L + (W - s.Width) / 2
            Case "middle": s.Top = T + (H - s.Height) / 2
            Case "fill":   s.LockAspectRatio = msoFalse: s.Left = L: s.Top = T: s.Width = W: s.Height = H
            Case "width":  s.Left = L: s.LockAspectRatio = msoFalse: s.Width = W
        End Select
    Next
End Sub

' ---------------- Arrange inside a container ----------------
' First selected shape = container; the others are laid out in a grid that
' fills it (rows x columns chosen from the count), with the standard gap.
Public Sub LtxArrangeInContainer()
    Dim sr As Object, cont As Object, n As Long, cols As Long, rows As Long, gap As Single, i As Long, r As Long, c As Long
    Dim cw As Single, ch As Single, pad As Single, inp As String
    Set sr = NeedSel(3)
    If sr Is Nothing Then Exit Sub
    Set cont = sr(1)
    n = sr.Count - 1
    inp = InputBox("Columns inside the container:", LTX_APP, CStr(Int(Sqr(n) + 0.999)))
    If Len(inp) = 0 Then Exit Sub
    cols = Val(inp): If cols < 1 Then cols = 1
    rows = Int((n + cols - 1) / cols)
    gap = Val(LtxGet("Gap", "8")): pad = gap
    cw = (cont.Width - 2 * pad - (cols - 1) * gap) / cols
    ch = (cont.Height - 2 * pad - (rows - 1) * gap) / rows
    For i = 2 To sr.Count
        r = (i - 2) \ cols: c = (i - 2) Mod cols
        sr(i).LockAspectRatio = msoFalse
        sr(i).Left = cont.Left + pad + c * (cw + gap)
        sr(i).Top = cont.Top + pad + r * (ch + gap)
        sr(i).Width = cw: sr(i).Height = ch
        sr(i).ZOrder 0
    Next
End Sub

' Stack selected shapes vertically (or horizontally) flush with the first, standard gap
Public Sub LtxStack(axis As String)
    Dim sr As Object, i As Long, pos As Single, gap As Single
    Set sr = NeedSel(2)
    If sr Is Nothing Then Exit Sub
    gap = Val(LtxGet("Gap", "8"))
    If axis = "v" Then pos = sr(1).Top + sr(1).Height + gap Else pos = sr(1).Left + sr(1).Width + gap
    For i = 2 To sr.Count
        If axis = "v" Then
            sr(i).Left = sr(1).Left: sr(i).Top = pos: pos = pos + sr(i).Height + gap
        Else
            sr(i).Top = sr(1).Top: sr(i).Left = pos: pos = pos + sr(i).Width + gap
        End If
    Next
End Sub

' ---------------- Z-order / grouping ----------------
Public Sub LtxZ(cmd As Long)
    Dim sr As Object
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    sr.ZOrder cmd     ' 0 front, 1 back, 2 forward, 3 backward
End Sub

Public Sub LtxGroupToggle()
    Dim sr As Object, s As Object
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    If sr.Count = 1 Then
        If sr(1).Type = 6 Then sr(1).Ungroup.Select
    Else
        sr.Group.Select
    End If
End Sub

' ---------------- Select Match ----------------
' Select every shape on the slide sharing a property with the current one.
Public Sub LtxSelectMatch(prop As String)
    Dim sr As Object, a As Object, s As Object, first As Boolean, hit As Boolean, sld As Object
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    Set a = sr(1)
    Set sld = LtxActiveSlide()
    first = True
    For Each s In sld.Shapes
        hit = False
        On Error Resume Next
        Select Case prop
            Case "type":  hit = (s.Type = a.Type) And (s.AutoShapeType = a.AutoShapeType)
            Case "fill":  hit = (s.Fill.ForeColor.RGB = a.Fill.ForeColor.RGB) And (s.Fill.Visible = a.Fill.Visible)
            Case "line":  hit = (s.Line.ForeColor.RGB = a.Line.ForeColor.RGB) And (s.Line.Weight = a.Line.Weight)
            Case "font":  hit = (s.TextFrame.TextRange.Font.Size = a.TextFrame.TextRange.Font.Size) And _
                                (s.TextFrame.TextRange.Font.Name = a.TextFrame.TextRange.Font.Name)
            Case "fontcolor": hit = (s.TextFrame.TextRange.Font.Color.RGB = a.TextFrame.TextRange.Font.Color.RGB)
            Case "size":  hit = (Abs(s.Width - a.Width) < 0.5) And (Abs(s.Height - a.Height) < 0.5)
            Case "width": hit = (Abs(s.Width - a.Width) < 0.5)
            Case "height": hit = (Abs(s.Height - a.Height) < 0.5)
            Case "linked": hit = LtxIsLinked(s)
            Case "text": hit = s.HasTextFrame And s.TextFrame.HasText
            Case "pictures": hit = (s.Type = 13 Or s.Type = 11)
        End Select
        On Error GoTo 0
        If hit Then
            s.Select Replace:=first
            first = False
        End If
    Next
End Sub

' ---------------- Convert to picture ----------------
' Replaces the selection with a picture of itself (EMF), keeping position.
Public Sub LtxConvertToPicture(Optional fmt As String = "emf")
    Dim sr As Object, L As Single, T As Single, W As Single, H As Single, sld As Object, shp As Object, z As Long
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    Set sld = LtxActiveSlide()
    L = sr(1).Left: T = sr(1).Top
    If sr.Count > 1 Then
        Dim R As Single, B As Single
        Bounds sr, L, T, R, B
    End If
    z = sr(1).ZOrderPosition
    sr.Copy
    If fmt = "png" Then
        Set shp = sld.Shapes.PasteSpecial(6)(1)
    Else
        Set shp = sld.Shapes.PasteSpecial(2)(1)
    End If
    sr.Delete
    shp.Left = L: shp.Top = T
    Do While shp.ZOrderPosition > z
        shp.ZOrder 3
        If shp.ZOrderPosition = 1 Then Exit Do
    Loop
    shp.Select
End Sub

' Crop a picture to the slide bounds (removes overhang)
Public Sub LtxCropToSlide()
    Dim sr As Object, s As Object
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    On Error Resume Next
    For Each s In sr
        If s.Left < 0 Then s.PictureFormat.CropLeft = s.PictureFormat.CropLeft - s.Left: s.Left = 0
        If s.Top < 0 Then s.PictureFormat.CropTop = s.PictureFormat.CropTop - s.Top: s.Top = 0
        If s.Left + s.Width > SlideW() Then s.PictureFormat.CropRight = s.PictureFormat.CropRight + (s.Left + s.Width - SlideW())
        If s.Top + s.Height > SlideH() Then s.PictureFormat.CropBottom = s.PictureFormat.CropBottom + (s.Top + s.Height - SlideH())
    Next
End Sub

' Copy exact size & position from the anchor to the others
Public Sub LtxCopyGeometry()
    Dim sr As Object, i As Long
    Set sr = NeedSel(2)
    If sr Is Nothing Then Exit Sub
    For i = 2 To sr.Count
        sr(i).LockAspectRatio = msoFalse
        sr(i).Left = sr(1).Left: sr(i).Top = sr(1).Top
        sr(i).Width = sr(1).Width: sr(i).Height = sr(1).Height
    Next
End Sub

' Paste geometry across slides: remember anchor, apply to selection later
Public Sub LtxRememberPosition()
    Dim sr As Object
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    LtxSet "PosL", CStr(sr(1).Left): LtxSet "PosT", CStr(sr(1).Top)
    LtxSet "PosW", CStr(sr(1).Width): LtxSet "PosH", CStr(sr(1).Height)
End Sub

Public Sub LtxApplyPosition(Optional sizeToo As Boolean = True)
    Dim sr As Object, s As Object
    Set sr = NeedSel(1)
    If sr Is Nothing Then Exit Sub
    If Len(LtxGet("PosL")) = 0 Then MsgBox "Nothing remembered yet. Use 'Remember Position' first.", vbInformation, LTX_APP: Exit Sub
    For Each s In sr
        s.Left = Val(LtxGet("PosL")): s.Top = Val(LtxGet("PosT"))
        If sizeToo Then s.LockAspectRatio = msoFalse: s.Width = Val(LtxGet("PosW")): s.Height = Val(LtxGet("PosH"))
    Next
End Sub
