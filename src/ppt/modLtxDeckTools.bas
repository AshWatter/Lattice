Attribute VB_Name = "modLtxDeckTools"
' =====================================================================
'  Lattice for PowerPoint - deck-wide tools for analysts
'  Font harmonizer, table number formats and layout, chart palette,
'  comment cleanup, sticky-note collector, per-section PDF export.
' =====================================================================
Option Explicit

Private Const msoTrue = -1
Private Const msoFalse = 0

' ---------- fonts ----------
' Set every text run in the deck (shapes, tables, groups) to one font family;
' optionally raise anything below a minimum size.
Public Sub LtxHarmonizeFonts()
    Dim fam As String, minSize As Single, sld As Object, shp As Object, n As Long
    fam = InputBox("Font family for the whole deck (blank = keep):", LTX_APP, LtxGet("DeckFont", ActivePresentation.SlideMaster.Theme.ThemeFontScheme.MinorFont(1).Name))
    minSize = Val(InputBox("Minimum font size (0 = no change):", LTX_APP, LtxGet("MinFont", "8")))
    If Len(fam) > 0 Then LtxSet "DeckFont", fam
    LtxSet "MinFont", CStr(minSize)
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            n = n + FixFonts(shp, fam, minSize)
        Next
    Next
    MsgBox n & " text container(s) harmonized.", vbInformation, LTX_APP
End Sub

Private Function FixFonts(shp As Object, fam As String, minSize As Single) As Long
    Dim g As Object, r As Long, c As Long, n As Long
    On Error Resume Next
    If shp.Type = 6 Then
        For Each g In shp.GroupItems: n = n + FixFonts(g, fam, minSize): Next
    ElseIf shp.HasTable Then
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                n = n + FixRange(shp.Table.Cell(r, c).Shape.TextFrame.TextRange, fam, minSize)
            Next
        Next
    ElseIf shp.HasTextFrame Then
        If shp.TextFrame.HasText Then n = n + FixRange(shp.TextFrame.TextRange, fam, minSize)
    End If
    FixFonts = n
End Function

Private Function FixRange(tr As Object, fam As String, minSize As Single) As Long
    Dim i As Long
    On Error Resume Next
    If Len(fam) > 0 Then tr.Font.Name = fam
    If minSize > 0 Then
        For i = 1 To tr.Runs.Count
            If tr.Runs(i).Font.Size < minSize Then tr.Runs(i).Font.Size = minSize
        Next
    End If
    FixRange = 1
End Function

' Report fonts used, with counts (what Deck Check summarises)
Public Sub LtxFontReport()
    Dim sld As Object, shp As Object, d As Object, k As Variant, s As String
    Set d = CreateObject("Scripting.Dictionary")
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            CountFonts shp, d
        Next
    Next
    For Each k In d.Keys: s = s & k & ": " & d(k) & " run(s)" & vbLf: Next
    MsgBox IIf(Len(s) = 0, "No text found.", s), vbInformation, LTX_APP & " - Fonts"
End Sub

Private Sub CountFonts(shp As Object, d As Object)
    Dim g As Object, i As Long, r As Long, c As Long, tr As Object
    On Error Resume Next
    If shp.Type = 6 Then
        For Each g In shp.GroupItems: CountFonts g, d: Next
    ElseIf shp.HasTable Then
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                Set tr = shp.Table.Cell(r, c).Shape.TextFrame.TextRange
                If Len(tr.Text) > 0 Then d(tr.Font.Name) = d(tr.Font.Name) + 1
            Next
        Next
    ElseIf shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            For i = 1 To shp.TextFrame.TextRange.Runs.Count
                d(shp.TextFrame.TextRange.Runs(i).Font.Name) = d(shp.TextFrame.TextRange.Runs(i).Font.Name) + 1
            Next
        End If
    End If
End Sub

' ---------- tables ----------
Private Function SelTables() As Collection
    Dim sr As Object, s As Object, col As Collection
    Set col = New Collection
    Set sr = LtxSel()
    If sr Is Nothing Then Set SelTables = col: Exit Function
    For Each s In sr
        If s.HasTable Then col.Add s
    Next
    Set SelTables = col
End Function

' Re-format every numeric cell in the selected tables: thousands separators,
' parentheses for negatives, fixed decimals, right aligned.
Public Sub LtxTableNumberFormat()
    Dim t As Object, r As Long, c As Long, tr As Object, txt As String, v As Double, dec As Long, neg As Boolean, pct As Boolean, mult As Boolean
    Dim tbls As Collection, n As Long
    Set tbls = SelTables()
    If tbls.Count = 0 Then MsgBox "Select one or more tables.", vbInformation, LTX_APP: Exit Sub
    dec = Val(InputBox("Decimals for plain numbers:", LTX_APP, LtxGet("TblDecimals", "1")))
    LtxSet "TblDecimals", CStr(dec)
    For Each t In tbls
        For r = 1 To t.Table.Rows.Count
            For c = 1 To t.Table.Columns.Count
                Set tr = t.Table.Cell(r, c).Shape.TextFrame.TextRange
                txt = Trim$(tr.Text)
                If Len(txt) > 0 Then
                    pct = (Right$(txt, 1) = "%"): mult = (LCase$(Right$(txt, 1)) = "x")
                    neg = (Left$(txt, 1) = "(" And Right$(txt, 1) = ")") Or Left$(txt, 1) = "-"
                    txt = Replace(Replace(Replace(Replace(Replace(txt, "(", ""), ")", ""), ",", ""), "%", ""), "x", "")
                    txt = Replace(Replace(txt, "-", ""), " ", "")
                    If IsNumeric(txt) And Len(txt) > 0 Then
                        v = Val(txt) * IIf(neg, -1, 1)
                        If pct Then
                            tr.Text = Format$(Abs(v), "0.0") & "%": If v < 0 Then tr.Text = "(" & tr.Text & ")"
                        ElseIf mult Then
                            tr.Text = Format$(Abs(v), "0.0") & "x": If v < 0 Then tr.Text = "(" & tr.Text & ")"
                        Else
                            tr.Text = Format$(Abs(v), "#,##0" & IIf(dec > 0, "." & String$(dec, "0"), "")): If v < 0 Then tr.Text = "(" & tr.Text & ")"
                        End If
                        tr.ParagraphFormat.Alignment = 3
                        n = n + 1
                    End If
                End If
            Next
        Next
    Next
    MsgBox n & " numeric cell(s) formatted.", vbInformation, LTX_APP
End Sub

' Equalise column widths (except the first) and row heights in selected tables
Public Sub LtxTableDistribute()
    Dim t As Object, c As Long, r As Long, w As Single, h As Single, tbls As Collection
    Set tbls = SelTables()
    If tbls.Count = 0 Then Exit Sub
    For Each t In tbls
        If t.Table.Columns.Count > 1 Then
            w = 0
            For c = 2 To t.Table.Columns.Count: w = w + t.Table.Columns(c).Width: Next
            For c = 2 To t.Table.Columns.Count: t.Table.Columns(c).Width = w / (t.Table.Columns.Count - 1): Next
        End If
        h = 0
        For r = 1 To t.Table.Rows.Count: h = h + t.Table.Rows(r).Height: Next
        For r = 1 To t.Table.Rows.Count: t.Table.Rows(r).Height = h / t.Table.Rows.Count: Next
    Next
End Sub

Public Sub LtxTableFontSize()
    Dim t As Object, r As Long, c As Long, sz As Single, tbls As Collection
    Set tbls = SelTables()
    If tbls.Count = 0 Then Exit Sub
    sz = Val(InputBox("Font size for the selected table(s):", LTX_APP, LtxGet("TableFontSize", "10")))
    If sz <= 0 Then Exit Sub
    For Each t In tbls
        For r = 1 To t.Table.Rows.Count
            For c = 1 To t.Table.Columns.Count
                t.Table.Cell(r, c).Shape.TextFrame.TextRange.Font.Size = sz
                t.Table.Cell(r, c).Shape.TextFrame.MarginTop = 1: t.Table.Cell(r, c).Shape.TextFrame.MarginBottom = 1
            Next
        Next
    Next
End Sub

' House header: dark fill, white bold text, bottom rule; body banding off
Public Sub LtxTableHeaderStyle()
    Dim t As Object, c As Long, tbls As Collection, hexc As String
    Set tbls = SelTables()
    If tbls.Count = 0 Then Exit Sub
    hexc = LtxGet("TableHeaderHex", "1F3864")
    For Each t In tbls
        t.Table.FirstRow = False: t.Table.HorizBanding = False
        For c = 1 To t.Table.Columns.Count
            With t.Table.Cell(1, c).Shape
                .Fill.Visible = msoTrue: .Fill.Solid
                .Fill.ForeColor.RGB = RGB(CLng("&H" & Mid$(hexc, 1, 2)), CLng("&H" & Mid$(hexc, 3, 2)), CLng("&H" & Mid$(hexc, 5, 2)))
                .TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
                .TextFrame.TextRange.Font.Bold = msoTrue
            End With
        Next
        Dim r As Long
        For r = 2 To t.Table.Rows.Count
            For c = 1 To t.Table.Columns.Count
                t.Table.Cell(r, c).Shape.Fill.Visible = msoFalse
                t.Table.Cell(r, c).Borders(1).Visible = msoFalse
                t.Table.Cell(r, c).Borders(3).Visible = msoTrue
                t.Table.Cell(r, c).Borders(3).Weight = 0.5
                t.Table.Cell(r, c).Borders(3).ForeColor.RGB = RGB(217, 217, 217)
            Next
        Next
    Next
End Sub

' ---------- charts ----------
' Apply the house palette (Settings > chart palette, hex list) to every chart in the deck
Public Sub LtxChartPalette()
    Dim sld As Object, shp As Object, pal As Variant, n As Long
    pal = Split(LtxGet("ChartPalette", "1F3864|2E75B6|9DC3E6|BDD7EE|7F7F7F|BFBFBF|C00000|ED7D31"), "|")
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasChart Then
                If PaintChart(shp.Chart, pal) Then n = n + 1
            End If
        Next
    Next
    MsgBox n & " chart(s) recoloured.", vbInformation, LTX_APP
End Sub

Private Function PaintChart(ch As Object, pal As Variant) As Boolean
    Dim i As Long, s As Object, h As String
    On Error Resume Next
    For i = 1 To ch.SeriesCollection.Count
        h = pal((i - 1) Mod (UBound(pal) + 1))
        Set s = ch.SeriesCollection(i)
        s.Format.Fill.ForeColor.RGB = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
        s.Format.Line.ForeColor.RGB = s.Format.Fill.ForeColor.RGB
        s.MarkerBackgroundColor = s.Format.Fill.ForeColor.RGB
    Next
    ch.ChartArea.Format.Line.Visible = msoFalse
    If ch.HasLegend Then ch.Legend.Font.Size = 9
    PaintChart = True
End Function

Public Sub LtxSetChartPalette()
    LtxSet "ChartPalette", InputBox("Chart palette, hex colours separated by |:", LTX_APP, LtxGet("ChartPalette", "1F3864|2E75B6|9DC3E6|BDD7EE|7F7F7F|BFBFBF|C00000|ED7D31"))
End Sub

' ---------- cleanup before sending ----------
Public Sub LtxDeleteComments()
    Dim sld As Object, n As Long, i As Long
    For Each sld In ActivePresentation.Slides
        For i = sld.Comments.Count To 1 Step -1
            sld.Comments(i).Delete: n = n + 1
        Next
    Next
    MsgBox n & " comment(s) deleted.", vbInformation, LTX_APP
End Sub

' Gather every sticky note / stamp / comment into one review slide at the end
Public Sub LtxCollectNotes()
    Dim sld As Object, shp As Object, t As String, i As Long, rev As Object, box As Object, cm As Object
    For Each sld In ActivePresentation.Slides
        If sld.Tags("LTX_REVIEW") <> "" Then sld.Delete
    Next
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Name Like "Lattice stamp*" Or (shp.HasTextFrame And shp.Fill.ForeColor.RGB = RGB(255, 242, 157)) Then
                If shp.HasTextFrame Then If shp.TextFrame.HasText Then t = t & "Slide " & sld.SlideIndex & ": " & Replace(shp.TextFrame.TextRange.Text, vbCr, " ") & vbCr: i = i + 1
            End If
        Next
        For Each cm In sld.Comments
            t = t & "Slide " & sld.SlideIndex & " (" & cm.Author & "): " & cm.Text & vbCr: i = i + 1
        Next
    Next
    If i = 0 Then MsgBox "No sticky notes, stamps or comments found.", vbInformation, LTX_APP: Exit Sub
    Set rev = ActivePresentation.Slides.AddSlide(ActivePresentation.Slides.Count + 1, ActivePresentation.SlideMaster.CustomLayouts(ActivePresentation.SlideMaster.CustomLayouts.Count))
    rev.Tags.Add "LTX_REVIEW", "1"
    Set box = rev.Shapes.AddTextbox(1, 30, 30, ActivePresentation.PageSetup.SlideWidth - 60, ActivePresentation.PageSetup.SlideHeight - 60)
    box.TextFrame.TextRange.Text = "REVIEW ITEMS (" & i & ") - delete this slide before sending" & vbCr & t
    box.TextFrame.TextRange.Font.Size = 10
    box.Name = "Lattice stamp review"
    ActiveWindow.View.GotoSlide rev.SlideIndex
End Sub

' One PDF per section, saved next to the deck
Public Sub LtxExportSectionsPdf()
    Dim secs As Object, s As Long, p As String, base As String, pr As Object, n As Long, first As Long, last As Long, i As Long, one As Object
    If Len(ActivePresentation.Path) = 0 Then MsgBox "Save the deck first.", vbInformation, LTX_APP: Exit Sub
    Set secs = ActivePresentation.SectionProperties
    If secs.Count = 0 Then MsgBox "No sections in this deck.", vbInformation, LTX_APP: Exit Sub
    base = ActivePresentation.Path & "\" & Left$(ActivePresentation.Name, InStrRev(ActivePresentation.Name, ".") - 1)
    For s = 1 To secs.Count
        If secs.SlidesCount(s) > 0 Then
            first = secs.FirstSlide(s): last = first + secs.SlidesCount(s) - 1
            Set pr = ActivePresentation.PrintOptions.Ranges
            pr.ClearAll
            Set one = pr.Add(first, last)
            p = base & " - " & s & " " & CleanFile(secs.Name(s)) & ".pdf"
            ActivePresentation.ExportAsFixedFormat p, 2, 1, msoFalse, 1, 1, msoFalse, one, 4   ' ppPrintSlideRange with a PrintRange
            n = n + 1
        End If
    Next
    ActivePresentation.PrintOptions.Ranges.ClearAll
    MsgBox n & " PDF(s) written next to the deck.", vbInformation, LTX_APP
End Sub

Private Function CleanFile(s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr("\/:*?""<>|", ch) = 0 Then CleanFile = CleanFile & ch
    Next
End Function

' Replace one font family with another everywhere (e.g. Calibri -> Arial)
Public Sub LtxReplaceFontDeck()
    Dim a As String, b As String
    a = InputBox("Replace font:", LTX_APP, "Calibri")
    If Len(a) = 0 Then Exit Sub
    b = InputBox("With font:", LTX_APP, "Arial")
    If Len(b) = 0 Then Exit Sub
    On Error Resume Next
    ActivePresentation.Fonts.Replace a, b
    MsgBox "Replaced " & a & " with " & b & " throughout the deck.", vbInformation, LTX_APP
End Sub

' ---------- Conform titles / subtitles / like-selected across the deck ----------
' Reference = the selected shape if there is one, else the first slide's placeholder
' of that role. Applies font name, size, bold, italic, colour, alignment, and the
' box position / size to the matching element on every slide.
Public Sub LtxConformRole(role As String)
    Dim ref As Object, sld As Object, shp As Object, n As Long, sr As Object
    Set sr = LtxSel()
    If Not sr Is Nothing Then Set ref = sr(1)
    If ref Is Nothing Then
        For Each sld In ActivePresentation.Slides
            Set ref = RoleShape(sld, role)
            If Not ref Is Nothing Then Exit For
        Next
    End If
    If ref Is Nothing Then MsgBox "No " & role & " found to use as the reference.", vbInformation, LTX_APP: Exit Sub
    For Each sld In ActivePresentation.Slides
        Set shp = RoleShape(sld, role)
        If Not shp Is Nothing Then
            If Not (shp.Parent.SlideIndex = ref.Parent.SlideIndex And shp.Name = ref.Name) Then
                ApplyTextLook shp, ref, True
                n = n + 1
            End If
        End If
    Next
    MsgBox n & " " & role & "(s) conformed to slide " & ref.Parent.SlideIndex & "'s " & role & ".", vbInformation, LTX_APP
End Sub

' Placeholder for a role on a slide: title (1 / 3), subtitle (4), body (2 / 7)
Private Function RoleShape(sld As Object, role As String) As Object
    Dim shp As Object, t As Long
    On Error Resume Next
    If role = "title" Then
        If sld.Shapes.HasTitle Then Set RoleShape = sld.Shapes.Title
        Exit Function
    End If
    For Each shp In sld.Shapes
        If shp.Type = 14 Then
            t = shp.PlaceholderFormat.Type
            If role = "subtitle" And t = 4 Then Set RoleShape = shp: Exit Function
            If role = "body" And (t = 2 Or t = 7) Then Set RoleShape = shp: Exit Function
        End If
    Next
    If role = "subtitle" Then
        ' no subtitle placeholder: use the first non-title text box just under the title
        Dim title As Object, best As Object, gap As Single, bestGap As Single
        bestGap = 1E+9
        If sld.Shapes.HasTitle Then Set title = sld.Shapes.Title Else Exit Function
        For Each shp In sld.Shapes
            If shp.HasTextFrame And shp.Name <> title.Name And shp.Type <> 14 Then
                If shp.TextFrame.HasText And shp.Top >= title.Top And shp.Top < title.Top + title.Height + 60 Then
                    gap = shp.Top - (title.Top + title.Height)
                    If Abs(gap) < bestGap Then bestGap = Abs(gap): Set best = shp
                End If
            End If
        Next
        Set RoleShape = best
    End If
End Function

Private Sub ApplyTextLook(shp As Object, ref As Object, geometry As Boolean)
    On Error Resume Next
    With shp.TextFrame.TextRange.Font
        .Name = ref.TextFrame.TextRange.Font.Name
        .Size = ref.TextFrame.TextRange.Font.Size
        .Bold = ref.TextFrame.TextRange.Font.Bold
        .Italic = ref.TextFrame.TextRange.Font.Italic
        If ref.TextFrame.TextRange.Font.Color.Type = 2 Then
            .Color.ObjectThemeColor = ref.TextFrame.TextRange.Font.Color.ObjectThemeColor
        Else
            .Color.RGB = ref.TextFrame.TextRange.Font.Color.RGB
        End If
    End With
    shp.TextFrame.TextRange.ParagraphFormat.Alignment = ref.TextFrame.TextRange.ParagraphFormat.Alignment
    shp.TextFrame.VerticalAnchor = ref.TextFrame.VerticalAnchor
    shp.TextFrame.MarginLeft = ref.TextFrame.MarginLeft: shp.TextFrame.MarginTop = ref.TextFrame.MarginTop
    If geometry Then
        shp.Left = ref.Left: shp.Top = ref.Top: shp.Width = ref.Width: shp.Height = ref.Height
    End If
End Sub

' Conform the shape on every slide that sits where the selected one sits
' (same slot by overlap), for text boxes that are not placeholders.
Public Sub LtxConformLikeSelected()
    Dim sr As Object, ref As Object, sld As Object, shp As Object, best As Object, ov As Single, bestOv As Single, n As Long
    Set sr = LtxSel()
    If sr Is Nothing Then MsgBox "Select the reference text box first.", vbInformation, LTX_APP: Exit Sub
    Set ref = sr(1)
    For Each sld In ActivePresentation.Slides
        If sld.SlideIndex <> ref.Parent.SlideIndex Then
            Set best = Nothing: bestOv = 0.4
            For Each shp In sld.Shapes
                If shp.HasTextFrame And shp.Type = ref.Type Then
                    ov = Overlap(shp, ref)
                    If ov > bestOv Then bestOv = ov: Set best = shp
                End If
            Next
            If Not best Is Nothing Then ApplyTextLook best, ref, True: n = n + 1
        End If
    Next
    MsgBox n & " matching text box(es) conformed to the selected one.", vbInformation, LTX_APP
End Sub

Private Function Overlap(a As Object, b As Object) As Single
    Dim ix As Single, iy As Single, inter As Single, uni As Single
    ix = Min2(a.Left + a.Width, b.Left + b.Width) - Max2(a.Left, b.Left)
    iy = Min2(a.Top + a.Height, b.Top + b.Height) - Max2(a.Top, b.Top)
    If ix <= 0 Or iy <= 0 Then Exit Function
    inter = ix * iy
    uni = a.Width * a.Height + b.Width * b.Height - inter
    If uni > 0 Then Overlap = inter / uni
End Function
Private Function Min2(x As Single, y As Single) As Single
    If x < y Then Min2 = x Else Min2 = y
End Function
Private Function Max2(x As Single, y As Single) As Single
    If x > y Then Max2 = x Else Max2 = y
End Function

' Font consistency check: most common font / size per role, list the slides that deviate
Public Sub LtxFontConsistencyCheck()
    Dim roles As Variant, role As Variant, sld As Object, shp As Object, d As Object, k As String, modeKey As String, modeN As Long
    Dim rep As String, dev As String, kk As Variant, total As Long
    roles = Array("title", "subtitle", "body")
    For Each role In roles
        Set d = CreateObject("Scripting.Dictionary")
        total = 0
        For Each sld In ActivePresentation.Slides
            Set shp = RoleShape(sld, CStr(role))
            If Not shp Is Nothing Then
                If shp.TextFrame.HasText Then
                    k = shp.TextFrame.TextRange.Font.Name & " " & shp.TextFrame.TextRange.Font.Size & "pt" & IIf(shp.TextFrame.TextRange.Font.Bold, " bold", "")
                    d(k) = d(k) + 1: total = total + 1
                End If
            End If
        Next
        If total > 0 Then
            modeN = 0
            For Each kk In d.Keys
                If d(kk) > modeN Then modeN = d(kk): modeKey = kk
            Next
            rep = rep & UCase$(CStr(role)) & ": " & modeKey & " on " & modeN & " of " & total & " slide(s)" & vbLf
            dev = ""
            For Each sld In ActivePresentation.Slides
                Set shp = RoleShape(sld, CStr(role))
                If Not shp Is Nothing Then
                    If shp.TextFrame.HasText Then
                        k = shp.TextFrame.TextRange.Font.Name & " " & shp.TextFrame.TextRange.Font.Size & "pt" & IIf(shp.TextFrame.TextRange.Font.Bold, " bold", "")
                        If k <> modeKey Then dev = dev & "   slide " & sld.SlideIndex & ": " & k & vbLf
                    End If
                End If
            Next
            rep = rep & IIf(Len(dev) = 0, "   all consistent" & vbLf, dev) & vbLf
        End If
    Next
    If Len(rep) = 0 Then rep = "No titles, subtitles or body placeholders with text found."
    MsgBox rep & "Use Deck Tools > Conform titles / subtitles to fix.", vbInformation, LTX_APP & " - Font consistency"
End Sub

' Set all slide titles to a consistent size and position (from the first slide's title)
Public Sub LtxAlignTitles()
    Dim ref As Object, sld As Object, n As Long
    For Each sld In ActivePresentation.Slides
        If sld.Shapes.HasTitle Then
            If ref Is Nothing Then
                Set ref = sld.Shapes.Title
            Else
                With sld.Shapes.Title
                    .Left = ref.Left: .Top = ref.Top: .Width = ref.Width: .Height = ref.Height
                    .TextFrame.TextRange.Font.Size = ref.TextFrame.TextRange.Font.Size
                End With
                n = n + 1
            End If
        End If
    Next
    MsgBox n & " title(s) aligned to slide " & IIf(ref Is Nothing, "-", ref.Parent.SlideIndex) & ".", vbInformation, LTX_APP
End Sub
