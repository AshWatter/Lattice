Attribute VB_Name = "modLtxStickers"
' =====================================================================
'  Lattice for PowerPoint - stickers
'  Status marks bankers put on slides: traffic lights, tick / cross, Harvey-ball cycling, and
'  DRAFT / CONFIDENTIAL style stamps in the top-right corner (one per slide, replaced on re-run,
'  removable in one go). Shapes are named "LtxSticker ..." / "LtxStamp ..." so Select Match and
'  the remove command find them.
' =====================================================================
Option Explicit

Private Const msoTrue = -1
Private Const msoFalse = 0

Public Sub LtxSticker(kind As String)
    Dim sl As Object, x As Single, y As Single, s As Object
    Set sl = LtxActiveSlide()
    If sl Is Nothing Then Exit Sub
    On Error GoTo fail
    Select Case kind
        Case "light_green", "light_amber", "light_red"
            StickerAnchor x, y
            Set s = Light(sl, x, y, Mid$(kind, 7))
        Case "lights"
            StickerAnchor x, y
            Set s = Light(sl, x, y, "red"): Set s = Light(sl, x + 22, y, "amber"): Set s = Light(sl, x + 44, y, "green")
        Case "tick", "cross"
            StickerAnchor x, y
            Set s = Mark(sl, x, y, kind)
        Case "harvey_cycle"
            Set s = HarveyCycle()
        Case "stamp_draft":        Set s = Stamp(sl, "DRAFT")
        Case "stamp_confidential": Set s = Stamp(sl, "CONFIDENTIAL")
        Case "stamp_preliminary":  Set s = Stamp(sl, "PRELIMINARY")
        Case "stamp_discussion":   Set s = Stamp(sl, "FOR DISCUSSION")
        Case "stamp_illustrative": Set s = Stamp(sl, "ILLUSTRATIVE")
        Case "stamp_draft_all"
            For Each sl In ActivePresentation.Slides
                Stamp sl, "DRAFT"
            Next
            Set s = Nothing
        Case "stamps_remove"
            RemoveStamps
    End Select
    If Not s Is Nothing Then s.Select
    Exit Sub
fail:
    MsgBox "Sticker: " & Err.Description, vbExclamation, LTX_APP
End Sub

' Next to the selected shape, else a default spot
Private Sub StickerAnchor(ByRef x As Single, ByRef y As Single)
    Dim sr As Object
    Set sr = LtxSel()
    If sr Is Nothing Then
        x = 60: y = 100
    Else
        x = sr(1).Left + sr(1).Width + 8: y = sr(1).Top + (sr(1).Height - 18) / 2
        If x > ActivePresentation.PageSetup.SlideWidth - 40 Then x = sr(1).Left: y = sr(1).Top + sr(1).Height + 6
    End If
End Sub

Private Function Light(sl As Object, x As Single, y As Single, colour As String) As Object
    Dim s As Object, c As Long
    Select Case colour
        Case "green": c = RGB(0, 176, 80)
        Case "amber": c = RGB(255, 192, 0)
        Case Else:    c = RGB(192, 0, 0)
    End Select
    Set s = sl.Shapes.AddShape(9, x, y, 18, 18)        ' msoShapeOval
    s.Fill.ForeColor.RGB = c
    s.Line.ForeColor.RGB = RGB(255, 255, 255): s.Line.Weight = 1
    s.Shadow.Visible = msoFalse
    s.Name = "LtxSticker light " & colour & " " & Format$(Now, "hhnnss")
    Set Light = s
End Function

Private Function Mark(sl As Object, x As Single, y As Single, kind As String) As Object
    Dim s As Object
    Set s = sl.Shapes.AddTextbox(1, x, y - 2, 22, 22)
    With s.TextFrame
        .WordWrap = msoFalse
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .TextRange.Text = IIf(kind = "tick", ChrW(10004), ChrW(10006))
        .TextRange.Font.Name = "Segoe UI Symbol"
        .TextRange.Font.Size = 16: .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Color.RGB = IIf(kind = "tick", RGB(0, 140, 60), RGB(192, 0, 0))
        .TextRange.ParagraphFormat.Alignment = 2
    End With
    s.Fill.Visible = msoFalse: s.Line.Visible = msoFalse
    s.Name = "LtxSticker " & kind & " " & Format$(Now, "hhnnss")
    Set Mark = s
End Function

' A rotated red outline stamp in the top-right corner; one per slide
Private Function Stamp(sl As Object, txt As String) As Object
    Dim s As Object, w As Single, i As Long
    For i = sl.Shapes.Count To 1 Step -1
        If Left$(sl.Shapes(i).Name, 8) = "LtxStamp" Then sl.Shapes(i).Delete
    Next
    w = 34 + Len(txt) * 11
    Set s = sl.Shapes.AddShape(5, ActivePresentation.PageSetup.SlideWidth - w - 18, 12, w, 30)   ' rounded rectangle
    s.Fill.Visible = msoFalse
    s.Line.ForeColor.RGB = RGB(192, 0, 0): s.Line.Weight = 2
    s.Shadow.Visible = msoFalse
    With s.TextFrame
        .WordWrap = msoFalse
        .TextRange.Text = txt
        .TextRange.Font.Name = "Segoe UI": .TextRange.Font.Size = 14: .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Color.RGB = RGB(192, 0, 0)
        .TextRange.ParagraphFormat.Alignment = 2
        .MarginTop = 2: .MarginBottom = 2
    End With
    s.Rotation = -6
    s.Name = "LtxStamp " & txt
    Set Stamp = s
End Function

Private Sub RemoveStamps()
    Dim sl As Object, i As Long, n As Long
    For Each sl In ActivePresentation.Slides
        For i = sl.Shapes.Count To 1 Step -1
            If Left$(sl.Shapes(i).Name, 8) = "LtxStamp" Then sl.Shapes(i).Delete: n = n + 1
        Next
    Next
    MsgBox n & " stamp(s) removed.", vbInformation, LTX_APP
End Sub

' Selected Harvey ball (from Turbo Shapes) -> next quarter: 0, 25, 50, 75, 100, 0 ...
Private Function HarveyCycle() As Object
    Dim sr As Object, s As Object, pct As Long, x As Single, y As Single, sl As Object
    Set sr = LtxSel()
    If sr Is Nothing Then MsgBox "Select a Harvey ball first (Turbo Shapes > Harvey balls).", vbInformation, LTX_APP: Exit Function
    Set s = sr(1)
    If Left$(s.Name, 7) <> "Harvey " Then MsgBox "The selected shape is not a Lattice Harvey ball.", vbInformation, LTX_APP: Exit Function
    pct = Val(Mid$(s.Name, 8)) + 25
    If pct > 100 Then pct = 0
    x = s.Left: y = s.Top
    Set sl = s.Parent
    s.Delete
    Set HarveyCycle = LtxHarveyBall(sl, x, y, pct)
End Function
