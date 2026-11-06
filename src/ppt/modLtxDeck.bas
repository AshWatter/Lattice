Attribute VB_Name = "modLtxDeck"
' =====================================================================
'  Lattice for PowerPoint - deck-level tools
'  Agenda / section tracker, Deck Check, slide utilities, share & send
' =====================================================================
Option Explicit

Private Const msoTrue = -1
Private Const msoFalse = 0

' ---------------- Agenda ----------------
' Builds (or rebuilds) an agenda slide from the deck's section names.
' Each section also gets a divider slide with the current item highlighted,
' unless dividers are switched off in settings.
Public Sub LtxAgenda()
    Dim pres As Object, secs As Object, i As Long, names As Collection, s As Object
    Set pres = ActivePresentation
    Set secs = pres.SectionProperties
    If secs.Count = 0 Then
        MsgBox "This deck has no sections. Add sections (Home > Section) and run Agenda again.", vbInformation, LTX_APP
        Exit Sub
    End If
    Set names = New Collection
    For i = 1 To secs.Count
        If secs.SlidesCount(i) > 0 Then names.Add secs.Name(i)
    Next
    ' Remove previously generated agenda slides
    For i = pres.Slides.Count To 1 Step -1
        If pres.Slides(i).Tags("LTX_AGENDA") <> "" Then pres.Slides(i).Delete
    Next
    ' Master agenda at the front
    BuildAgendaSlide pres, 1, names, 0
    If LtxBool("AgendaDividers", True) Then
        For i = 1 To secs.Count
            If secs.SlidesCount(i) > 0 Then
                BuildAgendaSlide pres, secs.FirstSlide(i), names, IndexOf(names, secs.Name(i))
            End If
        Next
    End If
    MsgBox "Agenda built for " & names.Count & " section(s).", vbInformation, LTX_APP
End Sub

Private Function IndexOf(col As Collection, v As String) As Long
    Dim i As Long
    For i = 1 To col.Count
        If col(i) = v Then IndexOf = i: Exit Function
    Next
End Function

Private Sub BuildAgendaSlide(pres As Object, pos As Long, names As Collection, current As Long)
    Dim sld As Object, lay As Object, i As Long, tb As Object, title As Object, y As Single, w As Single
    Set lay = pres.SlideMaster.CustomLayouts(IIf(pres.SlideMaster.CustomLayouts.Count >= 2, 2, 1))
    Set sld = pres.Slides.AddSlide(pos, lay)
    sld.Tags.Add "LTX_AGENDA", IIf(current = 0, "master", CStr(current))
    On Error Resume Next
    Set title = sld.Shapes.Title
    If Not title Is Nothing Then title.TextFrame.TextRange.Text = LtxGet("AgendaTitle", "Agenda")
    ' Remove body placeholders, we draw our own list
    For i = sld.Shapes.Count To 1 Step -1
        If sld.Shapes(i).Type = 14 Then
            If sld.Shapes(i).PlaceholderFormat.Type <> 1 And sld.Shapes(i).PlaceholderFormat.Type <> 3 Then sld.Shapes(i).Delete
        End If
    Next
    On Error GoTo 0
    w = pres.PageSetup.SlideWidth * 0.7
    y = pres.PageSetup.SlideHeight * 0.3
    For i = 1 To names.Count
        Set tb = sld.Shapes.AddTextbox(1, pres.PageSetup.SlideWidth * 0.15, y, w, 28)
        With tb.TextFrame.TextRange
            .Text = i & ". " & names(i)
            .Font.Size = 18
            If current = 0 Or current = i Then
                .Font.Bold = msoTrue
                .Font.Color.ObjectThemeColor = 13      ' text 1
            Else
                .Font.Color.RGB = RGB(150, 150, 150)
            End If
        End With
        tb.Name = "Agenda item " & i
        y = y + 34
    Next
End Sub

' ---------------- Deck Check ----------------
' Scans the deck for the usual pre-send problems and writes a report slide
' (or shows it in a message box if short).
Public Sub LtxDeckCheck()
    Dim pres As Object, sld As Object, shp As Object, issues As Collection, t As String, fonts As Object
    Dim sw As Single, sh As Single, k As Variant, txt As String, i As Long, pat As Variant
    Set pres = ActivePresentation
    Set issues = New Collection
    Set fonts = CreateObject("Scripting.Dictionary")
    sw = pres.PageSetup.SlideWidth: sh = pres.PageSetup.SlideHeight
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            ' off-slide
            If shp.Left + shp.Width < 0 Or shp.Top + shp.Height < 0 Or shp.Left > sw Or shp.Top > sh Then
                issues.Add "Slide " & sld.SlideIndex & ": '" & shp.Name & "' is entirely off the slide"
            ElseIf shp.Left < -1 Or shp.Top < -1 Or shp.Left + shp.Width > sw + 1 Or shp.Top + shp.Height > sh + 1 Then
                issues.Add "Slide " & sld.SlideIndex & ": '" & shp.Name & "' overhangs the slide edge"
            End If
            ' empty placeholders
            If shp.Type = 14 Then
                If shp.HasTextFrame Then
                    If Not shp.TextFrame.HasText Then issues.Add "Slide " & sld.SlideIndex & ": empty placeholder '" & shp.Name & "'"
                End If
            End If
            ' leftover markers and fonts
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    txt = shp.TextFrame.TextRange.Text
                    For Each pat In Split(LtxGet("CheckWords", "TBD,TBU,xx,XXX,[placeholder],lorem,DRAFT,TODO"), ",")
                        If InStr(1, txt, CStr(pat), vbTextCompare) > 0 Then
                            issues.Add "Slide " & sld.SlideIndex & ": '" & shp.Name & "' contains '" & pat & "'"
                            Exit For
                        End If
                    Next
                    If InStr(txt, "  ") > 0 Then issues.Add "Slide " & sld.SlideIndex & ": double spaces in '" & shp.Name & "'"
                    On Error Resume Next
                    For i = 1 To shp.TextFrame.TextRange.Runs.Count
                        t = shp.TextFrame.TextRange.Runs(i).Font.Name
                        If Len(t) > 0 Then fonts(t) = fonts(t) + 1
                    Next
                    On Error GoTo 0
                End If
            End If
            ' stale links
            If LtxIsLinked(shp) Then
                If Not LtxFileExists(shp.Tags(TAG_WB)) Then issues.Add "Slide " & sld.SlideIndex & ": link source missing - " & shp.Tags(TAG_WB)
            End If
            If shp.Name Like "Lattice stamp*" Then issues.Add "Slide " & sld.SlideIndex & ": review stamp still present"
        Next
        If sld.Shapes.Count = 0 Then issues.Add "Slide " & sld.SlideIndex & ": empty slide"
    Next
    If fonts.Count > 2 Then
        t = ""
        For Each k In fonts.Keys: t = t & k & " (" & fonts(k) & "), ": Next
        issues.Add "Deck uses " & fonts.Count & " fonts: " & Left$(t, Len(t) - 2)
    End If
    If issues.Count = 0 Then
        MsgBox "Deck Check: no issues found.", vbInformation, LTX_APP
        Exit Sub
    End If
    t = ""
    For i = 1 To issues.Count: t = t & issues(i) & vbLf: Next
    If issues.Count <= 25 Then
        MsgBox "Deck Check found " & issues.Count & " issue(s):" & vbLf & vbLf & t, vbExclamation, LTX_APP
    Else
        Set sld = pres.Slides.AddSlide(pres.Slides.Count + 1, pres.SlideMaster.CustomLayouts(pres.SlideMaster.CustomLayouts.Count))
        Set shp = sld.Shapes.AddTextbox(1, 20, 20, sw - 40, sh - 40)
        shp.TextFrame.TextRange.Text = "DECK CHECK (" & issues.Count & " issues) - delete this slide" & vbCr & t
        shp.TextFrame.TextRange.Font.Size = 9
        shp.Name = "Lattice stamp deck check"
        ActiveWindow.View.GotoSlide sld.SlideIndex
    End If
End Sub

' ---------------- Slide utilities ----------------
Public Sub LtxDeleteHiddenSlides()
    Dim i As Long, n As Long
    For i = ActivePresentation.Slides.Count To 1 Step -1
        If ActivePresentation.Slides(i).SlideShowTransition.Hidden = msoTrue Then
            ActivePresentation.Slides(i).Delete: n = n + 1
        End If
    Next
    MsgBox n & " hidden slide(s) deleted.", vbInformation, LTX_APP
End Sub

Public Sub LtxRemoveNotes()
    Dim sld As Object, shp As Object
    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.NotesPage.Shapes
            If shp.PlaceholderFormat.Type = 2 Then shp.TextFrame.TextRange.Text = ""
        Next
    Next
    MsgBox "Speaker notes cleared.", vbInformation, LTX_APP
End Sub

Public Sub LtxRemoveStamps()
    Dim sld As Object, i As Long, n As Long
    For Each sld In ActivePresentation.Slides
        For i = sld.Shapes.Count To 1 Step -1
            If sld.Shapes(i).Name Like "Lattice stamp*" Then sld.Shapes(i).Delete: n = n + 1
        Next
    Next
    MsgBox n & " stamp(s) removed.", vbInformation, LTX_APP
End Sub

' Duplicate the current slide(s) right after themselves
Public Sub LtxDuplicateSlides()
    On Error Resume Next
    ActiveWindow.Selection.SlideRange.Duplicate.Select
End Sub

' Put slide numbers on every slide via the footer placeholder
Public Sub LtxSlideNumbers(onOff As Boolean)
    Dim sld As Object
    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        sld.HeadersFooters.SlideNumber.Visible = IIf(onOff, msoTrue, msoFalse)
    Next
End Sub

' ---------------- File tools ----------------
Public Sub LtxSaveBackup()
    Dim p As String, bdir As String, nm As String
    If Len(ActivePresentation.Path) = 0 Then MsgBox "Save the deck first.", vbInformation, LTX_APP: Exit Sub
    bdir = ActivePresentation.Path & "\Backups"
    If Len(Dir(bdir, vbDirectory)) = 0 Then MkDir bdir
    nm = ActivePresentation.Name
    nm = Left$(nm, InStrRev(nm, ".") - 1) & "_" & Format$(Now, "yyyymmdd_hhnn") & Mid$(nm, InStrRev(nm, "."))
    p = bdir & "\" & nm
    ActivePresentation.SaveCopyAs p
    MsgBox "Backup saved:" & vbLf & p, vbInformation, LTX_APP
End Sub

' Share = PDF copy next to the deck (optionally of the selected slides only)
Public Sub LtxExportPdf(Optional selectedOnly As Boolean = False)
    Dim p As String, rng As Object
    If Len(ActivePresentation.Path) = 0 Then MsgBox "Save the deck first.", vbInformation, LTX_APP: Exit Sub
    p = ActivePresentation.Path & "\" & Left$(ActivePresentation.Name, InStrRev(ActivePresentation.Name, ".") - 1) & ".pdf"
    If selectedOnly Then
        ActivePresentation.ExportAsFixedFormat p, 2, 1, msoFalse, 1, 1, msoFalse, Nothing, 4   ' ppPrintSelection
    Else
        ActivePresentation.ExportAsFixedFormat p, 2
    End If
    MsgBox "PDF saved:" & vbLf & p, vbInformation, LTX_APP
End Sub

' Send = new Outlook mail with the deck attached (nothing is sent automatically)
Public Sub LtxSendMail(asPdf As Boolean)
    Dim ol As Object, m As Object, p As String
    On Error GoTo fail
    If Len(ActivePresentation.Path) = 0 Then MsgBox "Save the deck first.", vbInformation, LTX_APP: Exit Sub
    If ActivePresentation.Saved = msoFalse Then ActivePresentation.Save
    p = ActivePresentation.FullName
    If asPdf Then
        p = ActivePresentation.Path & "\" & Left$(ActivePresentation.Name, InStrRev(ActivePresentation.Name, ".") - 1) & ".pdf"
        ActivePresentation.ExportAsFixedFormat p, 2
    End If
    Set ol = CreateObject("Outlook.Application")
    Set m = ol.CreateItem(0)
    m.Subject = Left$(ActivePresentation.Name, InStrRev(ActivePresentation.Name, ".") - 1)
    m.Attachments.Add p
    m.Display
    Exit Sub
fail:
    MsgBox "Could not create the email: " & Err.Description, vbExclamation, LTX_APP
End Sub

' New deck from the configured template
Public Sub LtxNewFromTemplate()
    Dim t As String
    t = LtxGet("Template")
    If Not LtxFileExists(t) Then
        t = InputBox("Full path of your .potx / .pptx template:", LTX_APP, t)
        If Not LtxFileExists(t) Then Exit Sub
        LtxSet "Template", t
    End If
    Presentations.Open t, msoFalse, msoTrue, msoTrue
End Sub

' Collect every linked workbook and open them (handy before a big refresh)
Public Sub LtxOpenAllSources()
    Dim sld As Object, shp As Object, d As Object, k As Variant, xl As Object, st As Boolean
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If LtxIsLinked(shp) Then d(shp.Tags(TAG_WB)) = 1
        Next
    Next
    If d.Count = 0 Then MsgBox "No links in this deck.", vbInformation, LTX_APP: Exit Sub
    Set xl = LtxExcel(st): xl.Visible = True
    On Error Resume Next
    For Each k In d.Keys
        If LtxFindWorkbook(xl, CStr(k)) Is Nothing Then xl.Workbooks.Open CStr(k)
    Next
End Sub
