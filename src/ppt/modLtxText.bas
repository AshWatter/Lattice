Attribute VB_Name = "modLtxText"
' =====================================================================
'  Lattice for PowerPoint - text, font and colour tools
' =====================================================================
Option Explicit

Private Const msoTrue = -1
Private Const msoFalse = 0

' Apply an action to every text range in the selection (shapes, table cells,
' or just the selected text when editing).
Private Sub ForEachText(cb As String, Optional arg As Variant)
    Dim sr As Object, s As Object, r As Long, c As Long
    On Error Resume Next
    If ActiveWindow.Selection.Type = 3 Then          ' text selection
        ApplyText ActiveWindow.Selection.TextRange, cb, arg
        Exit Sub
    End If
    Set sr = LtxSel()
    If sr Is Nothing Then Exit Sub
    For Each s In sr
        If s.HasTable Then
            For r = 1 To s.Table.Rows.Count
                For c = 1 To s.Table.Columns.Count
                    ApplyText s.Table.Cell(r, c).Shape.TextFrame.TextRange, cb, arg
                Next
            Next
        ElseIf s.Type = 6 Then      ' group
            Dim g As Object
            For Each g In s.GroupItems
                If g.HasTextFrame Then ApplyText g.TextFrame.TextRange, cb, arg
            Next
        ElseIf s.HasTextFrame Then
            ApplyText s.TextFrame.TextRange, cb, arg
        End If
    Next
End Sub

Private Sub ApplyText(tr As Object, cb As String, arg As Variant)
    On Error Resume Next
    Select Case cb
        Case "grow":      tr.Font.Size = tr.Font.Size + 1
        Case "shrink":    If tr.Font.Size > 2 Then tr.Font.Size = tr.Font.Size - 1
        Case "size":      tr.Font.Size = CSng(arg)
        Case "font":      tr.Font.Name = CStr(arg)
        Case "color":     tr.Font.Color.RGB = CLng(arg)
        Case "bold":      tr.Font.Bold = Not tr.Font.Bold
        Case "italic":    tr.Font.Italic = Not tr.Font.Italic
        Case "underline": tr.Font.Underline = Not tr.Font.Underline
        Case "super":     tr.Font.BaselineOffset = IIf(tr.Font.BaselineOffset > 0, 0, 0.3)
        Case "sub":       tr.Font.BaselineOffset = IIf(tr.Font.BaselineOffset < 0, 0, -0.25)
        Case "upper":     tr.ChangeCase 2
        Case "lower":     tr.ChangeCase 1
        Case "title":     tr.ChangeCase 4
        Case "sentence":  tr.ChangeCase 3
        Case "toggle":    tr.ChangeCase 5
        Case "bullets":   tr.ParagraphFormat.Bullet.Visible = Not tr.ParagraphFormat.Bullet.Visible
        Case "spacemore": tr.ParagraphFormat.SpaceWithin = tr.ParagraphFormat.SpaceWithin + 0.1
        Case "spaceless": If tr.ParagraphFormat.SpaceWithin > 0.5 Then tr.ParagraphFormat.SpaceWithin = tr.ParagraphFormat.SpaceWithin - 0.1
        Case "parabefore": tr.ParagraphFormat.SpaceBefore = CSng(arg)
        Case "paraafter":  tr.ParagraphFormat.SpaceAfter = CSng(arg)
        Case "alignleft":   tr.ParagraphFormat.Alignment = 1
        Case "aligncenter": tr.ParagraphFormat.Alignment = 2
        Case "alignright":  tr.ParagraphFormat.Alignment = 3
        Case "justify":     tr.ParagraphFormat.Alignment = 4
        Case "clear":
            tr.Font.Bold = msoFalse: tr.Font.Italic = msoFalse: tr.Font.Underline = msoFalse
            tr.Font.BaselineOffset = 0
            tr.Font.Color.ObjectThemeColor = 13       ' msoThemeColorText1
        Case "stripspaces":
            tr.Text = Trim$(Replace(Replace(tr.Text, "  ", " "), vbCr & vbCr, vbCr))
    End Select
End Sub

Public Sub LtxText(cmd As String, Optional arg As Variant)
    ForEachText cmd, arg
End Sub

' ---------------- Font colour cycle ----------------
' Each click moves the selected text to the next colour in the palette.
Public Sub LtxCycleFontColor()
    Dim pal As Variant, i As Long, cur As Long, nxt As Long, tr As Object
    pal = Palette("FontCycle", "0,8421504,16777215,12611584,32768,255")   ' black, grey, white, blue, green, red (BGR longs)
    On Error Resume Next
    If ActiveWindow.Selection.Type = 3 Then
        cur = ActiveWindow.Selection.TextRange.Font.Color.RGB
    Else
        cur = LtxSel()(1).TextFrame.TextRange.Font.Color.RGB
    End If
    nxt = CLng(pal(0))
    For i = 0 To UBound(pal)
        If CLng(pal(i)) = cur Then nxt = CLng(pal((i + 1) Mod (UBound(pal) + 1))): Exit For
    Next
    ForEachText "color", nxt
End Sub

Public Sub LtxCycleFill()
    Dim pal As Variant, i As Long, cur As Long, nxt As Long, sr As Object, s As Object
    pal = Palette("FillCycle", "16777215,15921906,13421772,12611584,0")
    Set sr = LtxSel(): If sr Is Nothing Then Exit Sub
    On Error Resume Next
    cur = sr(1).Fill.ForeColor.RGB
    If sr(1).Fill.Visible = msoFalse Then cur = -1
    nxt = CLng(pal(0))
    For i = 0 To UBound(pal)
        If CLng(pal(i)) = cur Then nxt = CLng(pal((i + 1) Mod (UBound(pal) + 1))): Exit For
    Next
    For Each s In sr
        s.Fill.Visible = msoTrue: s.Fill.Solid: s.Fill.ForeColor.RGB = nxt
    Next
End Sub

Public Sub LtxCycleLine()
    Dim sr As Object, s As Object, w As Single
    Set sr = LtxSel(): If sr Is Nothing Then Exit Sub
    On Error Resume Next
    w = sr(1).Line.Weight
    If sr(1).Line.Visible = msoFalse Then w = 0
    Select Case True
        Case w = 0:   w = 0.75
        Case w < 1:   w = 1.5
        Case w < 2:   w = 2.25
        Case Else:    w = 0
    End Select
    For Each s In sr
        If w = 0 Then
            s.Line.Visible = msoFalse
        Else
            s.Line.Visible = msoTrue: s.Line.Weight = w
        End If
    Next
End Sub

Private Function Palette(key As String, def As String) As Variant
    Palette = Split(LtxGet(key, def), ",")
End Function

' Apply a theme colour (1..12) to fill / line / font of the selection
Public Sub LtxThemeColor(target As String, idx As Long)
    Dim sr As Object, s As Object
    Set sr = LtxSel(): If sr Is Nothing Then Exit Sub
    On Error Resume Next
    For Each s In sr
        Select Case target
            Case "fill": s.Fill.Visible = msoTrue: s.Fill.Solid: s.Fill.ForeColor.ObjectThemeColor = idx
            Case "line": s.Line.Visible = msoTrue: s.Line.ForeColor.ObjectThemeColor = idx
            Case "font": s.TextFrame.TextRange.Font.Color.ObjectThemeColor = idx
        End Select
    Next
End Sub

' Pick any RGB via hex input e.g. 1F3864
Public Sub LtxHexColor(target As String)
    Dim h As String, v As Long, sr As Object, s As Object
    h = InputBox("Hex colour (RRGGBB):", LTX_APP, LtxGet("LastHex", "1F3864"))
    If Len(h) <> 6 Then Exit Sub
    LtxSet "LastHex", h
    v = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
    Set sr = LtxSel(): If sr Is Nothing Then Exit Sub
    On Error Resume Next
    For Each s In sr
        Select Case target
            Case "fill": s.Fill.Visible = msoTrue: s.Fill.Solid: s.Fill.ForeColor.RGB = v
            Case "line": s.Line.Visible = msoTrue: s.Line.ForeColor.RGB = v
            Case "font": s.TextFrame.TextRange.Font.Color.RGB = v
        End Select
    Next
End Sub

' Format painter that survives across slides: copy, then paste as often as needed
Public Sub LtxCopyFormat()
    Dim sr As Object
    Set sr = LtxSel(): If sr Is Nothing Then Exit Sub
    sr(1).PickUp
    LtxSet "FormatCopied", "1"
End Sub

Public Sub LtxPasteFormat()
    Dim sr As Object, s As Object
    Set sr = LtxSel(): If sr Is Nothing Then Exit Sub
    On Error Resume Next
    For Each s In sr: s.Apply: Next
End Sub

' Autofit controls
Public Sub LtxAutofit(mode As Long)   ' 0 none, 1 shape to text, 2 shrink text
    Dim sr As Object, s As Object
    Set sr = LtxSel(): If sr Is Nothing Then Exit Sub
    On Error Resume Next
    For Each s In sr
        If s.HasTextFrame Then s.TextFrame2.AutoSize = mode
    Next
End Sub

Public Sub LtxMargins(pt As Single)
    Dim sr As Object, s As Object
    Set sr = LtxSel(): If sr Is Nothing Then Exit Sub
    On Error Resume Next
    For Each s In sr
        With s.TextFrame
            .MarginLeft = pt: .MarginRight = pt: .MarginTop = pt: .MarginBottom = pt
        End With
    Next
End Sub

Public Sub LtxWordWrapToggle()
    Dim sr As Object, s As Object
    Set sr = LtxSel(): If sr Is Nothing Then Exit Sub
    On Error Resume Next
    For Each s In sr
        s.TextFrame.WordWrap = Not s.TextFrame.WordWrap
    Next
End Sub

' Insert a sticky note / stamp text box
Public Sub LtxStamp(txt As String, colorRgb As Long)
    Dim sld As Object, s As Object
    Set sld = LtxActiveSlide(): If sld Is Nothing Then Exit Sub
    Set s = sld.Shapes.AddTextbox(1, 20, 20, 120, 30)
    With s
        .Name = "Lattice stamp"
        .TextFrame.TextRange.Text = txt
        .TextFrame.TextRange.Font.Size = 14
        .TextFrame.TextRange.Font.Bold = msoTrue
        .TextFrame.TextRange.Font.Color.RGB = colorRgb
        .TextFrame.TextRange.ParagraphFormat.Alignment = 2
        .Line.Visible = msoTrue: .Line.ForeColor.RGB = colorRgb: .Line.Weight = 1.5
        .Fill.Visible = msoTrue: .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Rotation = -10
    End With
    s.Select
End Sub
