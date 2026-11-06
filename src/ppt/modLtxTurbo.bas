Attribute VB_Name = "modLtxTurbo"
' =====================================================================
'  Lattice for PowerPoint - Turbo Shapes, Master Shapes, Slide Library
'  Turbo shapes  : one-click consulting shapes (harvey balls, chevrons,
'                  sticky notes, callouts, stamps, brackets, timelines)
'  Master shapes : your own saved shapes, stored in %APPDATA%\Lattice\shapes.pptx
'  Library       : your own saved slides, stored in %APPDATA%\Lattice\library.pptx
' =====================================================================
Option Explicit

Private Const msoTrue = -1
Private Const msoFalse = 0

Private Function Sld() As Object
    Set Sld = LtxActiveSlide()
    If Sld Is Nothing Then MsgBox "Select a slide first.", vbInformation, LTX_APP
End Function

Private Function Anchor(ByRef x As Single, ByRef y As Single) As Boolean
    Dim sr As Object
    Set sr = LtxSel()
    If sr Is Nothing Then
        x = 60: y = 100
    Else
        x = sr(1).Left + sr(1).Width + 10: y = sr(1).Top
    End If
    ' keep new shapes on the slide
    If x > ActivePresentation.PageSetup.SlideWidth - 200 Then x = 60: y = y + sr(1).Height + 10
    If y > ActivePresentation.PageSetup.SlideHeight - 60 Then y = 100
    Anchor = True
End Function

' ---------------- Turbo shapes ----------------
Public Sub LtxTurbo(kind As String)
    Dim s As Object, x As Single, y As Single, sl As Object, i As Long, g As Object
    Set sl = Sld(): If sl Is Nothing Then Exit Sub
    Anchor x, y
    Select Case kind
        Case "harvey0", "harvey25", "harvey50", "harvey75", "harvey100"
            Set s = LtxHarveyBall(sl, x, y, Val(Mid$(kind, 7)))
        Case "chevron"
            Set s = sl.Shapes.AddShape(52, x, y, 110, 40)   ' msoShapeChevron
            s.TextFrame.TextRange.Text = "Step"
        Case "chevrons"
            For i = 0 To 3
                Set s = sl.Shapes.AddShape(52, x + i * 115, y, 110, 40)
                s.TextFrame.TextRange.Text = "Step " & (i + 1)
                s.Name = "Chevron " & (i + 1)
            Next
        Case "sticky"
            Set s = sl.Shapes.AddShape(1, x, y, 160, 90)
            With s
                .Fill.ForeColor.RGB = RGB(255, 242, 157): .Line.Visible = msoFalse
                .Shadow.Visible = msoTrue
                .TextFrame.TextRange.Text = "Note"
                .TextFrame.TextRange.Font.Color.RGB = 0
                .TextFrame.TextRange.Font.Size = 11
                .TextFrame.TextRange.ParagraphFormat.Alignment = 1
                .TextFrame.VerticalAnchor = 1
            End With
        Case "callout"
            Set s = sl.Shapes.AddShape(106, x, y, 150, 60)   ' msoShapeRectangularCallout
            s.TextFrame.TextRange.Text = "Comment"
        Case "rounded"
            Set s = sl.Shapes.AddShape(5, x, y, 150, 60)
            s.Adjustments(1) = 0.12
        Case "oval"
            Set s = sl.Shapes.AddShape(9, x, y, 80, 80)
        Case "arrowr"
            Set s = sl.Shapes.AddShape(33, x, y, 90, 36)    ' right arrow
        Case "arrowd"
            Set s = sl.Shapes.AddShape(36, x, y, 36, 90)    ' down arrow
        Case "bracketr"
            Set s = sl.Shapes.AddShape(30, x, y, 20, 120)   ' msoShapeRightBracket
        Case "bracel"
            Set s = sl.Shapes.AddShape(31, x, y, 20, 120)   ' msoShapeLeftBrace
        Case "line"
            Set s = sl.Shapes.AddLine(x, y, x + 200, y)
            s.Line.Weight = 1
        Case "dashed"
            Set s = sl.Shapes.AddLine(x, y, x + 200, y)
            s.Line.Weight = 1: s.Line.DashStyle = 4
        Case "check"
            Set s = sl.Shapes.AddTextbox(1, x, y, 30, 30)
            s.TextFrame.TextRange.Text = ChrW(10004)
            s.TextFrame.TextRange.Font.Color.RGB = RGB(0, 128, 0): s.TextFrame.TextRange.Font.Size = 20
        Case "cross"
            Set s = sl.Shapes.AddTextbox(1, x, y, 30, 30)
            s.TextFrame.TextRange.Text = ChrW(10008)
            s.TextFrame.TextRange.Font.Color.RGB = RGB(192, 0, 0): s.TextFrame.TextRange.Font.Size = 20
        Case "timeline"
            Set s = sl.Shapes.AddLine(x, y + 20, x + 400, y + 20)
            s.Line.Weight = 2
            For i = 0 To 4
                Set g = sl.Shapes.AddShape(9, x + i * 100 - 6, y + 14, 12, 12)
                g.Fill.ForeColor.ObjectThemeColor = 5: g.Line.Visible = msoFalse
                Set g = sl.Shapes.AddTextbox(1, x + i * 100 - 40, y + 30, 80, 20)
                g.TextFrame.TextRange.Text = "Q" & (i + 1)
                g.TextFrame.TextRange.Font.Size = 10
                g.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            Next
        Case "numbered"
            Set s = sl.Shapes.AddShape(9, x, y, 24, 24)
            s.Fill.ForeColor.ObjectThemeColor = 5: s.Line.Visible = msoFalse
            s.TextFrame.TextRange.Text = "1"
            s.TextFrame.TextRange.Font.Size = 11: s.TextFrame.TextRange.Font.Bold = msoTrue
            s.TextFrame.MarginLeft = 0: s.TextFrame.MarginRight = 0: s.TextFrame.MarginTop = 0: s.TextFrame.MarginBottom = 0
        Case "stampdraft"
            LtxStamp "DRAFT", RGB(192, 0, 0): Exit Sub
        Case "stampconf"
            LtxStamp "CONFIDENTIAL", RGB(192, 0, 0): Exit Sub
        Case "stampprelim"
            LtxStamp "PRELIMINARY", RGB(90, 90, 90): Exit Sub
        Case "source"
            Set s = sl.Shapes.AddTextbox(1, 30, ActivePresentation.PageSetup.SlideHeight - 40, 400, 18)
            s.TextFrame.TextRange.Text = "Source: "
            s.TextFrame.TextRange.Font.Size = 8
            s.Name = "Source note"
        Case "tablehdr"
            Set s = sl.Shapes.AddTable(4, 4, x, y, 320, 100)
            s.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Item"
    End Select
    If Not s Is Nothing Then
        If s.HasTextFrame Then
            If s.TextFrame.HasText And kind <> "sticky" And kind <> "check" And kind <> "cross" And kind <> "source" Then
                s.TextFrame.TextRange.Font.Size = 11
            End If
        End If
        s.Select
    End If
End Sub

' Harvey ball = circle outline + pie wedge
Public Function LtxHarveyBall(sl As Object, x As Single, y As Single, pct As Long) As Object
    Dim ring As Object, pie As Object, g As Object, d As Single
    d = 24
    Set ring = sl.Shapes.AddShape(9, x, y, d, d)
    ring.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ring.Line.ForeColor.RGB = 0: ring.Line.Weight = 1
    If pct >= 100 Then
        ring.Fill.ForeColor.RGB = 0
        Set LtxHarveyBall = ring: Exit Function
    End If
    If pct <= 0 Then Set LtxHarveyBall = ring: Exit Function
    Set pie = sl.Shapes.AddShape(142, x, y, d, d)     ' msoShapePie
    pie.Fill.ForeColor.RGB = 0: pie.Line.Visible = msoFalse
    pie.Adjustments(1) = -90                           ' start at 12 o'clock
    pie.Adjustments(2) = -90 + 360 * pct / 100
    Set g = sl.Shapes.Range(Array(ring.Name, pie.Name)).Group
    g.Name = "Harvey " & pct
    Set LtxHarveyBall = g
End Function

' ---------------- Master shapes (personal shape library) ----------------
Private Function ShapesFile() As String: ShapesFile = LtxDataDir() & "\shapes.pptx": End Function
Private Function LibraryFile() As String: LibraryFile = LtxDataDir() & "\library.pptx": End Function

Private Function OpenStore(path As String) As Object
    Dim p As Object
    On Error Resume Next
    Set p = Presentations(LtxBaseName(path))
    On Error GoTo 0
    If p Is Nothing Then
        If LtxFileExists(path) Then
            Set p = Presentations.Open(path, msoFalse, msoFalse, msoFalse)
        Else
            Set p = Presentations.Add(msoFalse)
            p.PageSetup.SlideWidth = ActivePresentation.PageSetup.SlideWidth
            p.PageSetup.SlideHeight = ActivePresentation.PageSetup.SlideHeight
            p.SaveAs path
        End If
    End If
    Set OpenStore = p
End Function

' Save the selected shape(s) as a named master shape
Public Sub LtxMasterSave()
    Dim sr As Object, store As Object, sld As Object, nm As String, shp As Object
    Set sr = LtxSel(): If sr Is Nothing Then MsgBox "Select the shape(s) to save.", vbInformation, LTX_APP: Exit Sub
    nm = InputBox("Name for this master shape:", LTX_APP, sr(1).Name)
    If Len(nm) = 0 Then Exit Sub
    Set store = OpenStore(ShapesFile())
    If store.Slides.Count = 0 Then store.Slides.Add 1, 12   ' ppLayoutBlank
    Set sld = store.Slides(1)
    sr.Copy
    Set shp = sld.Shapes.Paste
    If shp.Count > 1 Then Set shp = sld.Shapes.Range(ShapeNames(shp)).Group Else Set shp = shp(1)
    shp.Name = nm
    shp.Tags.Add "LTX_MASTER", nm
    store.Save: store.Close
    MsgBox "Saved master shape '" & nm & "'.", vbInformation, LTX_APP
End Sub

Private Function ShapeNames(sr As Object) As Variant
    Dim a() As String, i As Long
    ReDim a(1 To sr.Count)
    For i = 1 To sr.Count: a(i) = sr(i).Name: Next
    ShapeNames = a
End Function

' Insert a master shape by name (prompted with the list)
Public Sub LtxMasterInsert()
    Dim store As Object, shp As Object, list As String, nm As String, sl As Object, x As Single, y As Single, pasted As Object
    If Not LtxFileExists(ShapesFile()) Then MsgBox "No master shapes saved yet. Select a shape and use 'Save as Master'.", vbInformation, LTX_APP: Exit Sub
    Set sl = Sld(): If sl Is Nothing Then Exit Sub
    Set store = OpenStore(ShapesFile())
    For Each shp In store.Slides(1).Shapes: list = list & shp.Name & vbLf: Next
    nm = InputBox("Master shapes:" & vbLf & list & vbLf & "Type the name to insert:", LTX_APP, store.Slides(1).Shapes(1).Name)
    If Len(nm) = 0 Then store.Close: Exit Sub
    On Error Resume Next
    Set shp = store.Slides(1).Shapes(nm)
    On Error GoTo 0
    If shp Is Nothing Then MsgBox "No master shape called '" & nm & "'.", vbExclamation, LTX_APP: store.Close: Exit Sub
    shp.Copy
    Set pasted = sl.Shapes.Paste
    Anchor x, y
    pasted.Left = x: pasted.Top = y
    pasted.Select
    store.Close
End Sub

Public Sub LtxMasterManage()
    If Not LtxFileExists(ShapesFile()) Then MsgBox "No master shapes saved yet.", vbInformation, LTX_APP: Exit Sub
    Presentations.Open ShapesFile()
End Sub

' ---------------- Slide library ----------------
Public Sub LtxLibrarySave()
    Dim store As Object, src As Object, i As Long
    Set src = ActiveWindow.Selection.SlideRange
    If src Is Nothing Then Exit Sub
    Set store = OpenStore(LibraryFile())
    For i = 1 To src.Count
        src(i).Copy
        store.Slides.Paste store.Slides.Count + 1
    Next
    store.Save: store.Close
    MsgBox src.Count & " slide(s) saved to the library.", vbInformation, LTX_APP
End Sub

Public Sub LtxLibraryInsert()
    Dim store As Object, n As String, idx As Long, pos As Long, i As Long, list As String, t As String
    If Not LtxFileExists(LibraryFile()) Then MsgBox "The library is empty. Select slides and use 'Save to Library'.", vbInformation, LTX_APP: Exit Sub
    Set store = OpenStore(LibraryFile())
    For i = 1 To store.Slides.Count
        t = ""
        On Error Resume Next
        t = store.Slides(i).Shapes.Title.TextFrame.TextRange.Text
        On Error GoTo 0
        list = list & i & ": " & Left$(t, 50) & vbLf
    Next
    n = InputBox("Library slides:" & vbLf & list & vbLf & "Number to insert (0 = open the library file):", LTX_APP)
    If Len(n) = 0 Then store.Close: Exit Sub
    idx = Val(n)
    If idx = 0 Then store.Close: Presentations.Open LibraryFile(): Exit Sub
    If idx < 1 Or idx > store.Slides.Count Then store.Close: Exit Sub
    pos = LtxActiveSlide().SlideIndex + 1
    store.Slides(idx).Copy
    ActivePresentation.Slides.Paste pos
    ' Slide Link: remember where it came from so it can be refreshed later
    With ActivePresentation.Slides(pos).Tags
        .Add "LTX_SLIDELINK", LibraryFile()
        .Add "LTX_SLIDELINK_ID", CStr(store.Slides(idx).SlideID)
        .Add "LTX_SLIDELINK_TITLE", SlideTitle(store.Slides(idx))
    End With
    store.Close
    ActiveWindow.View.GotoSlide pos
End Sub

Private Function SlideTitle(sld As Object) As String
    On Error Resume Next
    If sld.Shapes.HasTitle Then SlideTitle = sld.Shapes.Title.TextFrame.TextRange.Text
End Function

' Refresh every slide that was inserted from a library (matched by slide id, then title)
Public Sub LtxRefreshSlideLinks()
    Dim i As Long, sld As Object, store As Object, src As Object, s As Object, n As Long, pos As Long, lay As Object
    Dim path As String, id As String, ttl As String, stores As Object, k As Variant
    Set stores = CreateObject("Scripting.Dictionary"): stores.CompareMode = 1
    For i = ActivePresentation.Slides.Count To 1 Step -1
        Set sld = ActivePresentation.Slides(i)
        path = sld.Tags("LTX_SLIDELINK")
        If Len(path) > 0 Then
            If Not LtxFileExists(path) Then GoTo nextSlide
            If Not stores.Exists(path) Then stores.Add path, Presentations.Open(path, msoFalse, msoFalse, msoFalse)
            Set store = stores(path)
            id = sld.Tags("LTX_SLIDELINK_ID"): ttl = sld.Tags("LTX_SLIDELINK_TITLE")
            Set src = Nothing
            On Error Resume Next
            Set src = store.Slides.FindBySlideID(CLng(id))
            On Error GoTo 0
            If src Is Nothing And Len(ttl) > 0 Then
                For Each s In store.Slides
                    If SlideTitle(s) = ttl Then Set src = s: Exit For
                Next
            End If
            If Not src Is Nothing Then
                pos = sld.SlideIndex
                src.Copy
                ActivePresentation.Slides.Paste pos + 1
                With ActivePresentation.Slides(pos + 1).Tags
                    .Add "LTX_SLIDELINK", path: .Add "LTX_SLIDELINK_ID", CStr(src.SlideID): .Add "LTX_SLIDELINK_TITLE", SlideTitle(src)
                End With
                sld.Delete
                n = n + 1
            End If
        End If
nextSlide:
    Next
    For Each k In stores.Keys: stores(k).Close: Next
    MsgBox n & " linked slide(s) refreshed from the library.", vbInformation, LTX_APP
End Sub

' Replace the library copy of the current slide with this version (push the change back)
Public Sub LtxUpdateSlideLink()
    Dim sld As Object, store As Object, src As Object, path As String, pos As Long
    Set sld = LtxActiveSlide()
    path = sld.Tags("LTX_SLIDELINK")
    If Len(path) = 0 Then MsgBox "This slide was not inserted from the library.", vbInformation, LTX_APP: Exit Sub
    Set store = OpenStore(path)
    On Error Resume Next
    Set src = store.Slides.FindBySlideID(CLng(sld.Tags("LTX_SLIDELINK_ID")))
    On Error GoTo 0
    If src Is Nothing Then store.Close: MsgBox "Library slide not found.", vbExclamation, LTX_APP: Exit Sub
    pos = src.SlideIndex
    sld.Copy
    store.Slides.Paste pos + 1
    src.Delete
    sld.Tags.Add "LTX_SLIDELINK_ID", CStr(store.Slides(pos).SlideID)
    store.Save: store.Close
    MsgBox "Library updated. Other decks pick it up with 'Refresh slide links'.", vbInformation, LTX_APP
End Sub
