Attribute VB_Name = "modLtxPages"
' =====================================================================
'  Lattice for PowerPoint - pagination
'  Page numbers "X" or "X of Y" with an offset, blank pages for facing-page
'  printing, section title sync, and cleanup of everything Lattice inserted.
' =====================================================================
Option Explicit

Private Const msoTrue = -1
Private Const msoFalse = 0

' ---------- page numbers ----------
' Draws a small text box on every slide (except those tagged LTX_NOPAGE) with
' the page number. Re-running updates in place.
Public Sub LtxPageNumbers()
    Dim fmt As String, offset As Long, sld As Object, shp As Object, n As Long, total As Long, i As Long
    Dim x As Single, y As Single, w As Single, h As Single, sw As Single, sh As Single, txt As String
    fmt = InputBox("Format ({n} = number, {t} = total):", LTX_APP, LtxGet("PageFmt", "{n}"))
    If Len(fmt) = 0 Then Exit Sub
    LtxSet "PageFmt", fmt
    offset = Val(InputBox("Start numbering at:", LTX_APP, LtxGet("PageStart", "1"))) - 1
    LtxSet "PageStart", CStr(offset + 1)
    sw = ActivePresentation.PageSetup.SlideWidth: sh = ActivePresentation.PageSetup.SlideHeight
    w = 60: h = 16: x = sw - w - 14: y = sh - h - 8
    For Each sld In ActivePresentation.Slides
        If sld.Tags("LTX_NOPAGE") = "" And sld.Tags("LTX_BLANK") = "" Then total = total + 1
    Next
    i = 0
    For Each sld In ActivePresentation.Slides
        Set shp = Nothing
        On Error Resume Next
        Set shp = sld.Shapes("Lattice page number")
        On Error GoTo 0
        If sld.Tags("LTX_NOPAGE") <> "" Or sld.Tags("LTX_BLANK") <> "" Then
            If Not shp Is Nothing Then shp.Delete
        Else
            i = i + 1
            txt = Replace(Replace(fmt, "{n}", CStr(i + offset)), "{t}", CStr(total + offset))
            If shp Is Nothing Then
                Set shp = sld.Shapes.AddTextbox(1, x, y, w, h)
                shp.Name = "Lattice page number"
                shp.TextFrame.TextRange.Font.Size = 8
                shp.TextFrame.TextRange.ParagraphFormat.Alignment = 3
                shp.TextFrame.MarginLeft = 0: shp.TextFrame.MarginRight = 0
                shp.TextFrame.WordWrap = msoFalse
                shp.TextFrame2.AutoSize = 0
            End If
            shp.TextFrame.TextRange.Text = txt
            n = n + 1
        End If
    Next
    MsgBox n & " slide(s) numbered.", vbInformation, LTX_APP
End Sub

Public Sub LtxPageNumbersRemove()
    Dim sld As Object
    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        sld.Shapes("Lattice page number").Delete
    Next
End Sub

' Exclude / include the current slide(s) from numbering
Public Sub LtxPageSkipToggle()
    Dim sld As Object
    For Each sld In ActiveWindow.Selection.SlideRange
        If sld.Tags("LTX_NOPAGE") = "" Then sld.Tags.Add "LTX_NOPAGE", "1" Else sld.Tags.Delete "LTX_NOPAGE"
    Next
    LtxPageNumbersSilent
End Sub

Private Sub LtxPageNumbersSilent()
    ' re-run numbering with the stored settings, no prompts
    Dim fmt As String, offset As Long, sld As Object, shp As Object, total As Long, i As Long, txt As String
    fmt = LtxGet("PageFmt", "{n}"): offset = Val(LtxGet("PageStart", "1")) - 1
    For Each sld In ActivePresentation.Slides
        If sld.Tags("LTX_NOPAGE") = "" And sld.Tags("LTX_BLANK") = "" Then total = total + 1
    Next
    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        Set shp = Nothing
        Set shp = sld.Shapes("Lattice page number")
        If sld.Tags("LTX_NOPAGE") <> "" Or sld.Tags("LTX_BLANK") <> "" Then
            If Not shp Is Nothing Then shp.Delete
        ElseIf Not shp Is Nothing Then
            i = i + 1
            shp.TextFrame.TextRange.Text = Replace(Replace(fmt, "{n}", CStr(i + offset)), "{t}", CStr(total + offset))
        End If
    Next
End Sub

' Native slide-number placeholder start value (File > Page Setup equivalent)
Public Sub LtxFirstSlideNumber()
    Dim v As String
    v = InputBox("First slide number (native numbering):", LTX_APP, CStr(ActivePresentation.PageSetup.FirstSlideNumber))
    If Len(v) > 0 Then ActivePresentation.PageSetup.FirstSlideNumber = Val(v)
End Sub

' ---------- blank pages (facing-page printing) ----------
' Inserts a blank slide after every slide (or after each section) so that a
' double-sided print shows each content page on the right-hand side.
Public Sub LtxBlankPages(mode As String)
    Dim i As Long, sld As Object, lay As Object, n As Long, secs As Object, s As Long, pos As Long
    LtxBlankPagesRemove True
    Set lay = ActivePresentation.SlideMaster.CustomLayouts(ActivePresentation.SlideMaster.CustomLayouts.Count)   ' usually Blank
    Select Case mode
        Case "every"
            For i = ActivePresentation.Slides.Count To 1 Step -1
                Set sld = ActivePresentation.Slides.AddSlide(i + 1, lay)
                MarkBlank sld: n = n + 1
            Next
        Case "sections"
            Set secs = ActivePresentation.SectionProperties
            For s = secs.Count To 1 Step -1
                If secs.SlidesCount(s) > 0 Then
                    pos = secs.FirstSlide(s) + secs.SlidesCount(s)
                    Set sld = ActivePresentation.Slides.AddSlide(pos, lay)
                    sld.MoveToSectionStart s   ' keep it inside the section
                    sld.MoveTo pos
                    MarkBlank sld: n = n + 1
                End If
            Next
        Case "odd"   ' pad so every section starts on an odd page
            Set secs = ActivePresentation.SectionProperties
            For s = secs.Count To 2 Step -1
                If (secs.FirstSlide(s) Mod 2) = 0 Then
                    Set sld = ActivePresentation.Slides.AddSlide(secs.FirstSlide(s), lay)
                    MarkBlank sld: n = n + 1
                End If
            Next
    End Select
    MsgBox n & " blank page(s) inserted.", vbInformation, LTX_APP
End Sub

Private Sub MarkBlank(sld As Object)
    Dim shp As Object, i As Long
    sld.Tags.Add "LTX_BLANK", "1"
    For i = sld.Shapes.Count To 1 Step -1: sld.Shapes(i).Delete: Next
    If LtxBool("BlankPageText", True) Then
        Set shp = sld.Shapes.AddTextbox(1, 0, ActivePresentation.PageSetup.SlideHeight / 2 - 10, ActivePresentation.PageSetup.SlideWidth, 20)
        shp.TextFrame.TextRange.Text = LtxGet("BlankText", "This page is intentionally left blank")
        shp.TextFrame.TextRange.Font.Size = 9
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(128, 128, 128)
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        shp.Name = "Lattice blank note"
    End If
End Sub

Public Sub LtxBlankPagesRemove(Optional quiet As Boolean = False)
    Dim i As Long, n As Long
    For i = ActivePresentation.Slides.Count To 1 Step -1
        If ActivePresentation.Slides(i).Tags("LTX_BLANK") <> "" Then ActivePresentation.Slides(i).Delete: n = n + 1
    Next
    If Not quiet Then MsgBox n & " blank page(s) removed.", vbInformation, LTX_APP
End Sub

' ---------- section titles ----------
' Writes each section's name into the title of its first slide (or of the
' Lattice divider slide if there is one) so renaming a section updates the deck.
Public Sub LtxSyncSectionTitles()
    Dim secs As Object, s As Long, sld As Object, n As Long
    Set secs = ActivePresentation.SectionProperties
    If secs.Count = 0 Then MsgBox "No sections in this deck.", vbInformation, LTX_APP: Exit Sub
    On Error Resume Next
    For s = 1 To secs.Count
        If secs.SlidesCount(s) > 0 Then
            Set sld = ActivePresentation.Slides(secs.FirstSlide(s))
            If sld.Tags("LTX_AGENDA") <> "" And sld.Tags("LTX_AGENDA") <> "master" Then
                ' divider built by Agenda: highlight already matches; refresh the agenda instead
            ElseIf sld.Shapes.HasTitle Then
                sld.Shapes.Title.TextFrame.TextRange.Text = secs.Name(s): n = n + 1
            End If
        End If
    Next
    MsgBox n & " section title(s) synced. Run Agenda to rebuild dividers.", vbInformation, LTX_APP
End Sub

' Rename the section containing the current slide from the slide's title
Public Sub LtxSectionFromTitle()
    Dim sld As Object, secs As Object, idx As Long
    Set sld = LtxActiveSlide()
    If Not sld.Shapes.HasTitle Then Exit Sub
    Set secs = ActivePresentation.SectionProperties
    idx = sld.sectionIndex
    If idx > 0 Then secs.Rename idx, sld.Shapes.Title.TextFrame.TextRange.Text
End Sub

' Table of contents slide listing every slide title with its page number
Public Sub LtxTableOfContents()
    Dim sld As Object, t As String, i As Long, toc As Object, shp As Object, lay As Object
    For Each sld In ActivePresentation.Slides
        If sld.Tags("LTX_TOC") <> "" Then sld.Delete
    Next
    For Each sld In ActivePresentation.Slides
        If sld.Tags("LTX_BLANK") = "" And sld.Tags("LTX_AGENDA") = "" Then
            If sld.Shapes.HasTitle Then
                If sld.Shapes.Title.TextFrame.HasText Then
                    i = i + 1
                    t = t & sld.SlideIndex & vbTab & Left$(sld.Shapes.Title.TextFrame.TextRange.Text, 70) & vbCr
                End If
            End If
        End If
    Next
    Set lay = ActivePresentation.SlideMaster.CustomLayouts(IIf(ActivePresentation.SlideMaster.CustomLayouts.Count >= 2, 2, 1))
    Set toc = ActivePresentation.Slides.AddSlide(1, lay)
    toc.Tags.Add "LTX_TOC", "1"
    On Error Resume Next
    toc.Shapes.Title.TextFrame.TextRange.Text = "Contents"
    For i = toc.Shapes.Count To 1 Step -1
        If toc.Shapes(i).Type = 14 Then If toc.Shapes(i).PlaceholderFormat.Type <> 1 And toc.Shapes(i).PlaceholderFormat.Type <> 3 Then toc.Shapes(i).Delete
    Next
    Set shp = toc.Shapes.AddTextbox(1, ActivePresentation.PageSetup.SlideWidth * 0.1, ActivePresentation.PageSetup.SlideHeight * 0.25, ActivePresentation.PageSetup.SlideWidth * 0.8, ActivePresentation.PageSetup.SlideHeight * 0.65)
    shp.TextFrame.TextRange.Text = t
    shp.TextFrame.TextRange.Font.Size = 11
    shp.Name = "Lattice TOC"
End Sub
