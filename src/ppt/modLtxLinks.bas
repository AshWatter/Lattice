Attribute VB_Name = "modLtxLinks"
' =====================================================================
'  Lattice for PowerPoint - Excel import and link refresh engine
' =====================================================================
Option Explicit

Private Const ppPasteEnhancedMetafile = 2
Private Const ppPastePNG = 6
Private Const ppPasteHTML = 8
Private Const ppPasteDefault = 0
Private Const ppPasteOLEObject = 10
Private mWarnings As String       ' formula-error warnings collected during a session
Private mCalc As Long             ' Excel calculation mode to restore after a session
Private mT0 As Double
Private Const msoTrue = -1
Private Const msoFalse = 0

' Cache of workbooks opened by Lattice during a refresh batch (closed at the end)
Private mOpened As Object      ' Scripting.Dictionary  path -> workbook
Private mRemap As Object       ' Scripting.Dictionary  old path -> chosen path ("" = keep old)
Private mXl As Object
Private mStartedXl As Boolean

' ---------------------------------------------------------------------
'  Session helpers (so "Refresh All" opens each workbook only once)
' ---------------------------------------------------------------------
' Optional per-stage timing (Settings value Profile=1 writes it to lattice.log)
Private Sub Tick(stage As String)
    If LtxBool("Profile") Then LtxLog "  [" & Format$((Timer - mT0) * 1000, "0") & " ms] " & stage: mT0 = Timer
End Sub

Private Sub BeginSession()
    Set mOpened = CreateObject("Scripting.Dictionary")
    mOpened.CompareMode = 1
    Set mRemap = CreateObject("Scripting.Dictionary")
    mRemap.CompareMode = 1
    mWarnings = ""
    Set mXl = LtxExcel(mStartedXl)
    mXl.ScreenUpdating = False
    On Error Resume Next
    mCalc = mXl.Calculation
    mXl.Calculation = -4135      ' xlCalculationManual while rendering; restored in EndSession
    mXl.EnableEvents = False
    On Error GoTo 0
End Sub

' Formula-error warnings gathered since BeginSession (shown once by the caller)
Private Sub ReportWarnings()
    If Len(mWarnings) > 0 Then MsgBox "Exported ranges contain formula errors:" & mWarnings, vbExclamation, LTX_APP & " - check formulas"
    mWarnings = ""
End Sub

' Version-aware source resolution. When the linked file is not open but another
' open workbook has the same sheet (V2 saved from V1), ask once per source which
' to use; the answer is remembered for the rest of the refresh and the links are
' repointed. When the file is missing on disk, offer a file picker.
Private Function ResolveSource(ByRef fullPath As String, sheetName As String, definedName As String) As Object
    Dim wb As Object, cand As Object, c As Object, msg As String, ans As VbMsgBoxResult, fd As Object, names As String
    If mRemap.Exists(fullPath) Then
        If Len(mRemap(fullPath)) > 0 Then fullPath = mRemap(fullPath)
    Else
        Set wb = LtxFindWorkbook(mXl, fullPath)
        If wb Is Nothing Then
            ' candidates: open workbooks with the same sheet (and defined name when we have one)
            For Each c In mXl.Workbooks
                If StrComp(c.FullName, fullPath, vbTextCompare) <> 0 And HasSheet(c, sheetName) Then
                    If Len(definedName) = 0 Or HasName(c, definedName) Then
                        If cand Is Nothing Then Set cand = c
                        names = names & vbLf & "  - " & c.Name
                    End If
                End If
            Next
            If Not cand Is Nothing Then
                msg = "This link points to:" & vbLf & "  " & LtxBaseName(fullPath) & IIf(LtxFileExists(fullPath), " (not open)", " (file not found)") & vbLf & vbLf & _
                      "Open workbook(s) with the same sheet '" & sheetName & "':" & names & vbLf & vbLf & _
                      "Yes = use " & cand.Name & " and repoint every link to " & LtxBaseName(fullPath) & vbLf & _
                      "No = keep " & LtxBaseName(fullPath) & IIf(LtxFileExists(fullPath), " (open it from disk)", "") & vbLf & "Cancel = skip these links"
                ans = MsgBox(msg, vbYesNoCancel + vbQuestion, LTX_APP & " - which version?")
                If ans = vbYes Then
                    mRemap(fullPath) = cand.FullName
                    fullPath = cand.FullName
                ElseIf ans = vbNo Then
                    mRemap(fullPath) = ""
                Else
                    mRemap(fullPath) = "SKIP"
                End If
            ElseIf Not LtxFileExists(fullPath) Then
                ans = MsgBox("Source not found:" & vbLf & fullPath & vbLf & vbLf & "Locate the workbook now? (links are repointed)", vbYesNo + vbQuestion, LTX_APP)
                If ans = vbYes Then
                    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker
                    fd.Title = "Locate " & LtxBaseName(fullPath)
                    fd.AllowMultiSelect = False
                    fd.Filters.Clear: fd.Filters.Add "Excel", "*.xls*"
                    If fd.Show = -1 Then
                        mRemap(fullPath) = fd.SelectedItems(1)
                        fullPath = fd.SelectedItems(1)
                    Else
                        mRemap(fullPath) = "SKIP"
                    End If
                Else
                    mRemap(fullPath) = "SKIP"
                End If
            Else
                mRemap(fullPath) = ""
            End If
        Else
            mRemap(fullPath) = ""
        End If
    End If
    If mRemap(fullPath) = "SKIP" Or (mRemap.Exists(fullPath) And mRemap(fullPath) = "SKIP") Then Exit Function
    Set ResolveSource = SessionWorkbook(fullPath)
End Function

Private Function HasSheet(wb As Object, sheetName As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = wb.Worksheets(sheetName)
    HasSheet = Not s Is Nothing
End Function

Private Function HasName(wb As Object, nm As String) As Boolean
    Dim n As Object
    On Error Resume Next
    Set n = wb.Names(nm)
    HasName = Not n Is Nothing
End Function

Private Sub EndSession()
    Dim k As Variant
    On Error Resume Next
    If Not mOpened Is Nothing Then
        For Each k In mOpened.Keys
            mOpened(k).Close SaveChanges:=False
        Next
    End If
    If Not mXl Is Nothing Then
        mXl.ScreenUpdating = True
        mXl.EnableEvents = True
        If mCalc <> 0 Then mXl.Calculation = mCalc
        mXl.CutCopyMode = False
        If mStartedXl Then mXl.Quit
    End If
    Set mXl = Nothing: Set mOpened = Nothing
End Sub

Private Function SessionWorkbook(fullPath As String) As Object
    Dim wb As Object, openedHere As Boolean
    If mOpened.Exists(fullPath) Then
        Set SessionWorkbook = mOpened(fullPath): Exit Function
    End If
    Set wb = LtxGetWorkbook(mXl, fullPath, openedHere)
    If openedHere Then mOpened.Add fullPath, wb
    Set SessionWorkbook = wb
End Function

' ---------------------------------------------------------------------
'  Current slide / selection helpers
' ---------------------------------------------------------------------
Public Function LtxActiveSlide() As Object
    On Error Resume Next
    Set LtxActiveSlide = ActiveWindow.Selection.SlideRange(1)
    If LtxActiveSlide Is Nothing Then Set LtxActiveSlide = ActiveWindow.View.Slide
End Function

Public Function LtxIsLinked(shp As Object) As Boolean
    On Error Resume Next
    LtxIsLinked = (Len(shp.Tags(TAG_WB)) > 0)
End Function

' ---------------------------------------------------------------------
'  IMPORT  (entry points called from the ribbon)
'  kind: "picture" | "table" | "chart"
' ---------------------------------------------------------------------
Public Sub LtxImport(kind As String)
    Dim sld As Object, src As Object, shp As Object, pend As Object
    Dim wbPath As String, sheetName As String, addr As String, dname As String
    Dim black As Boolean
    Set sld = LtxActiveSlide()
    If sld Is Nothing Then MsgBox "Select a slide first.", vbExclamation, LTX_APP: Exit Sub
    black = LtxBool("BlackFonts", True)

    On Error GoTo fail
    mT0 = Timer
    BeginSession
    Tick "BeginSession"
    ' 1) A range copied with the Excel add-in ("Copy for PowerPoint") wins if recent
    Dim exportSize As String, target As Object
    Set pend = LtxReadPending()
    If pend.Exists("wb") And pend.Exists("time") Then
        If Abs(DateDiff("s", CDate(pend("time")), Now)) < 180 Then
            wbPath = pend("wb"): sheetName = pend("sheet"): addr = pend("range")
            If pend.Exists("name") Then dname = pend("name")
            If pend.Exists("kind") Then kind = pend("kind")
            If pend.Exists("sizemode") Then exportSize = pend("sizemode")
            Kill LtxPendingPath()
        End If
    End If
    ' Paste over: if a shape is selected on the slide, the import takes its place and
    ' matches its width / height / size (Excel's Match Size setting); a selected
    ' Lattice link is replaced like a refresh.
    On Error Resume Next
    Dim selType As Long
    selType = -1
    selType = ActiveWindow.Selection.Type
    If selType = 2 Then Set target = ActiveWindow.Selection.ShapeRange(1)
    If target Is Nothing And selType = 3 Then Set target = ActiveWindow.Selection.ShapeRange(1)   ' text cursor inside a shape
    LtxLog "Import: kind=" & kind & " selType=" & selType & " target=" & IIf(target Is Nothing, "(none)", target.Name) & " exportSize=" & exportSize & " err=" & Err.Description
    On Error GoTo fail
    If Len(exportSize) = 0 Then exportSize = LtxGet("SizeMode", "width")
    If exportSize = "none" Or exportSize = "source" Then Set target = Nothing
    ' 2) Otherwise take whatever is selected in the running Excel
    If Len(wbPath) = 0 Then
        If mStartedXl Then
            EndSession
            MsgBox "Excel is not running. Select a range in Excel (or use 'Copy for PowerPoint' there) and try again.", vbExclamation, LTX_APP
            Exit Sub
        End If
        Set src = mXl.Selection
        If TypeName(src) = "ChartArea" Or TypeName(src) = "Chart" Then
            Set src = mXl.ActiveChart
            wbPath = src.Parent.Parent.Parent.FullName
            sheetName = src.Parent.Parent.Name
            addr = src.Parent.Name
            kind = "chart"
        ElseIf TypeName(src) = "ChartObject" Then
            wbPath = src.Parent.Parent.FullName: sheetName = src.Parent.Name: addr = src.Name: kind = "chart"
        ElseIf TypeName(src) = "Range" And kind = "chart" Then
            ' chart requested but a range is selected: use the active chart, or the only chart on the sheet
            If Not mXl.ActiveChart Is Nothing Then
                wbPath = mXl.ActiveChart.Parent.Parent.Parent.FullName: sheetName = mXl.ActiveChart.Parent.Parent.Name: addr = mXl.ActiveChart.Parent.Name
            ElseIf src.Parent.ChartObjects.Count >= 1 Then
                ' use the chart nearest the selected cell
                wbPath = src.Parent.Parent.FullName: sheetName = src.Parent.Name: addr = NearestChart(src).Name
            Else
                EndSession
                MsgBox "Select the chart in Excel first (click its border).", vbExclamation, LTX_APP
                Exit Sub
            End If
        ElseIf TypeName(src) = "Range" Then
            wbPath = src.Parent.Parent.FullName: sheetName = src.Parent.Name: addr = src.Address(False, False)
        Else
            EndSession
            MsgBox "Select a range or a chart in Excel first.", vbExclamation, LTX_APP
            Exit Sub
        End If
    End If

    If target Is Nothing Then
        Set shp = RenderLink(sld, kind, wbPath, sheetName, addr, dname, black, Nothing)
    Else
        ' use the selected shape as the geometry template, then remove it
        If Len(target.Tags(TAG_SIZEMODE)) = 0 Then target.Tags.Add TAG_SIZEMODE, IIf(exportSize = "size", "keep", exportSize)
        If exportSize = "size" Then target.Tags.Add TAG_SIZEMODE, "keep" Else target.Tags.Add TAG_SIZEMODE, exportSize
        If Len(target.Tags(TAG_ID)) = 0 Then target.Tags.Add TAG_ID, LtxGuid()
        Set shp = RenderLink(sld, kind, wbPath, sheetName, addr, dname, black, target)
    End If
    EndSession
    ReportWarnings
    If shp Is Nothing Then Exit Sub
    shp.Select
    Exit Sub
fail:
    EndSession
    MsgBox "Import failed: " & Err.Description, vbCritical, LTX_APP
End Sub

' Chart object on the selection's sheet whose top-left is closest to the selection
Private Function NearestChart(src As Object) As Object
    Dim co As Object, best As Object, d As Double, bd As Double
    bd = 1E+12
    For Each co In src.Parent.ChartObjects
        d = (co.Left - src.Left) ^ 2 + (co.Top - src.Top) ^ 2
        If d < bd Then bd = d: Set best = co
    Next
    Set NearestChart = best
End Function

' Renders one link onto a slide. When oldShape is given, the new shape replaces
' it in place (same position, size mode, z-order, name and tags).
Private Function RenderLink(sld As Object, kind As String, wbPath As String, sheetName As String, _
                            addr As String, dname As String, black As Boolean, oldShape As Object) As Object
    Dim wb As Object, rng As Object, shp As Object, sizeMode As String
    Dim L As Single, T As Single, W As Single, H As Single, z As Long, nm As String, id As String
    Tick "RenderLink start"
    Set wb = ResolveSource(wbPath, sheetName, dname)     ' may repoint wbPath to another version
    Tick "ResolveSource"
    If wb Is Nothing Then
        If mRemap.Exists(wbPath) Then If mRemap(wbPath) = "SKIP" Then Err.Raise vbObjectError + 5, , "skipped by user"
        Err.Raise vbObjectError + 1, , "Workbook not available: " & wbPath & IIf(LtxBool("AirplaneMode"), " (Airplane Mode is on)", "")
    End If
    wbPath = wb.FullName

    ' Copy the source to the clipboard
    Dim errTxt As String
    Select Case kind
        Case "chart"
            LtxCopyChartPicture wb, sheetName, addr
        Case "nativechart"
            LtxCopyChartObject wb, sheetName, addr
        Case "table", "embed"
            Set rng = LtxResolveRange(wb, sheetName, addr, dname)
            rng.Copy
        Case Else
            Set rng = LtxResolveRange(wb, sheetName, addr, dname)
            If Not LtxServedCopy(mXl, wb, sheetName, rng, black, (LtxGet("PictureFormat", "emf") = "png")) Then
                LtxCopyRangePicture rng, black, (LtxGet("PictureFormat", "emf") = "png")
            End If
    End Select
    Tick "copy to clipboard"
    If Not rng Is Nothing Then
        addr = rng.Address(False, False)
        If LtxBool("CheckErrors", True) Then
            errTxt = LtxRangeHasErrors(rng)
            If Len(errTxt) > 0 Then LtxLog "Export warning " & sheetName & "!" & addr & ": " & errTxt: mWarnings = mWarnings & vbLf & sheetName & "!" & addr & ": " & errTxt
        End If
    End If

    ' Paste onto the slide
    Select Case kind
        Case "table"
            Set shp = sld.Shapes.PasteSpecial(ppPasteHTML)(1)
            StyleTableFromRange shp, rng
            If black Then BlackenTable shp
        Case "nativechart"
            Set shp = sld.Shapes.PasteSpecial(ppPasteDefault)(1)      ' native, editable PowerPoint chart
        Case "embed"
            Set shp = sld.Shapes.PasteSpecial(ppPasteOLEObject)(1)    ' embedded workbook object
        Case Else
            If LtxGet("PictureFormat", "emf") = "png" Then
                Set shp = sld.Shapes.PasteSpecial(ppPastePNG)(1)
            Else
                Set shp = sld.Shapes.PasteSpecial(ppPasteEnhancedMetafile)(1)
            End If
    End Select
    mXl.CutCopyMode = False
    Tick "paste"

    ' Geometry
    If oldShape Is Nothing Then
        PlaceNewShape sld, shp
        sizeMode = LtxGet("SizeMode", "width")
        nm = "Lattice " & kind & " " & Format$(Now, "hhnnss")
        id = LtxGuid()
    Else
        L = oldShape.Left: T = oldShape.Top: W = oldShape.Width: H = oldShape.Height
        z = oldShape.ZOrderPosition: nm = oldShape.Name: id = oldShape.Tags(TAG_ID)
        sizeMode = oldShape.Tags(TAG_SIZEMODE)
        If Len(sizeMode) = 0 Then sizeMode = "width"
        shp.LockAspectRatio = msoTrue
        Select Case sizeMode
            Case "width":  shp.Width = W
            Case "height": shp.Height = H
            Case "keep":   shp.LockAspectRatio = msoFalse: shp.Width = W: shp.Height = H
            Case Else      ' source size: leave as pasted
        End Select
        shp.Left = L: shp.Top = T
        oldShape.Delete
        Do While shp.ZOrderPosition > z
            shp.ZOrder 3   ' msoSendBackward
            If shp.ZOrderPosition = 1 Then Exit Do
        Loop
    End If

    ' Tags
    On Error Resume Next
    shp.Name = nm
    On Error GoTo 0
    With shp.Tags
        .Add TAG_ID, id
        .Add TAG_TYPE, kind
        .Add TAG_WB, wbPath
        .Add TAG_SHEET, sheetName
        .Add TAG_RANGE, addr
        .Add TAG_NAME, dname
        .Add TAG_BLACK, IIf(black, "1", "0")
        .Add TAG_SIZEMODE, sizeMode
        .Add TAG_UPDATED, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    End With
    Tick "geometry+tags"
    Set RenderLink = shp
End Function

' New imports: centre on the slide and shrink to fit the content area.
Private Sub PlaceNewShape(sld As Object, shp As Object)
    Dim sw As Single, sh As Single, maxW As Single, maxH As Single, m As Single
    sw = ActivePresentation.PageSetup.SlideWidth
    sh = ActivePresentation.PageSetup.SlideHeight
    m = Val(LtxGet("MarginPt", "36"))
    maxW = sw - 2 * m: maxH = sh - 2 * m
    shp.LockAspectRatio = msoTrue
    If shp.Width > maxW Then shp.Width = maxW
    If shp.Height > maxH Then shp.Height = maxH
    shp.Left = (sw - shp.Width) / 2
    shp.Top = (sh - shp.Height) / 2
End Sub

' Native PowerPoint table: drop PowerPoint's default banding and carry over the
' Excel fills, bold, italic, alignment, font size and borders cell by cell.
Private Sub StyleTableFromRange(shp As Object, rng As Object)
    Dim r As Long, c As Long, xc As Object, cell As Object, tr As Object, fs As Single
    On Error Resume Next
    If rng Is Nothing Then Exit Sub
    If Not shp.HasTable Then Exit Sub
    shp.Table.ApplyStyle "{2D5ABB26-0587-4C30-8999-92F81FD0307C}", False   ' No Style, No Grid
    shp.Table.FirstRow = False: shp.Table.HorizBanding = False
    fs = Val(LtxGet("TableFontSize", "10"))
    For r = 1 To shp.Table.Rows.Count
        For c = 1 To shp.Table.Columns.Count
            If r <= rng.Rows.Count And c <= rng.Columns.Count Then
                Set xc = rng.Cells(r, c)
                Set cell = shp.Table.Cell(r, c)
                Set tr = cell.Shape.TextFrame.TextRange
                If xc.Interior.ColorIndex <> -4142 Then
                    cell.Shape.Fill.Visible = msoTrue: cell.Shape.Fill.Solid
                    cell.Shape.Fill.ForeColor.RGB = xc.Interior.Color
                Else
                    cell.Shape.Fill.Visible = msoFalse
                End If
                tr.Font.Bold = IIf(xc.Font.Bold, msoTrue, msoFalse)
                tr.Font.Italic = IIf(xc.Font.Italic, msoTrue, msoFalse)
                tr.Font.Size = IIf(fs > 0, fs, xc.Font.Size)
                tr.Font.Color.RGB = xc.Font.Color
                Select Case xc.HorizontalAlignment
                    Case -4131: tr.ParagraphFormat.Alignment = 1
                    Case -4108, 7: tr.ParagraphFormat.Alignment = 2
                    Case -4152: tr.ParagraphFormat.Alignment = 3
                    Case Else: tr.ParagraphFormat.Alignment = IIf(IsNumeric(xc.Value), 3, 1)
                End Select
                cell.Shape.TextFrame.MarginLeft = 3: cell.Shape.TextFrame.MarginRight = 3
                cell.Shape.TextFrame.MarginTop = 1: cell.Shape.TextFrame.MarginBottom = 1
                CopyBorder cell.Borders(1), xc.Borders(8)    ' top
                CopyBorder cell.Borders(3), xc.Borders(9)    ' bottom
            End If
        Next
    Next
End Sub

Private Sub CopyBorder(pb As Object, xb As Object)
    On Error Resume Next
    If xb.LineStyle = -4142 Then
        pb.Visible = msoFalse
    Else
        pb.Visible = msoTrue
        pb.Weight = IIf(xb.Weight >= 3, 1.5, 0.75)
        pb.ForeColor.RGB = IIf(xb.ColorIndex = -4105, 0, xb.Color)
        If xb.LineStyle = -4119 Then pb.Weight = 2.25   ' double -> heavy
    End If
End Sub

' Native PowerPoint table: force dark fonts to black, keep light ones.
Private Sub BlackenTable(shp As Object)
    Dim r As Long, c As Long, tr As Object
    On Error Resume Next
    If Not shp.HasTable Then Exit Sub
    For r = 1 To shp.Table.Rows.Count
        For c = 1 To shp.Table.Columns.Count
            Set tr = shp.Table.Cell(r, c).Shape.TextFrame.TextRange
            If Not LtxIsLightColor(tr.Font.Color.RGB) Then tr.Font.Color.RGB = 0
        Next
    Next
End Sub

' ---------------------------------------------------------------------
'  REFRESH
' ---------------------------------------------------------------------
Public Sub LtxRefreshSelected()
    Dim shp As Object, n As Long, bad As Long, msg As String, items As Collection
    On Error GoTo fail
    Set items = New Collection
    If ActiveWindow.Selection.Type = 2 Or ActiveWindow.Selection.Type = 3 Then   ' shapes / text
        For Each shp In ActiveWindow.Selection.ShapeRange
            If LtxIsLinked(shp) Then items.Add shp
        Next
    End If
    If items.Count = 0 Then
        MsgBox "Select one or more Lattice-linked objects first.", vbInformation, LTX_APP
        Exit Sub
    End If
    BeginSession
    For Each shp In items
        If RefreshOne(shp, msg) Then n = n + 1 Else bad = bad + 1
    Next
    EndSession
    ReportWarnings
    Report n, bad, msg
    Exit Sub
fail:
    EndSession
    MsgBox "Refresh failed: " & Err.Description, vbCritical, LTX_APP
End Sub

Public Sub LtxRefreshSlide()
    Dim sld As Object
    Set sld = LtxActiveSlide()
    If sld Is Nothing Then Exit Sub
    RefreshSlides Array(sld.SlideIndex)
End Sub

Public Sub LtxRefreshAll()
    Dim idx() As Long, i As Long
    ReDim idx(1 To ActivePresentation.Slides.Count)
    For i = 1 To ActivePresentation.Slides.Count: idx(i) = i: Next
    RefreshSlides idx
End Sub

Private Sub RefreshSlides(slideIndexes As Variant)
    Dim i As Variant, shp As Object, items As Collection, n As Long, bad As Long, msg As String
    On Error GoTo fail
    Set items = New Collection
    For Each i In slideIndexes
        For Each shp In ActivePresentation.Slides(i).Shapes
            If LtxIsLinked(shp) Then items.Add shp
        Next
    Next
    If items.Count = 0 Then
        MsgBox "No Lattice links found.", vbInformation, LTX_APP: Exit Sub
    End If
    BeginSession
    For Each shp In items
        If RefreshOne(shp, msg) Then n = n + 1 Else bad = bad + 1
    Next
    EndSession
    ReportWarnings
    Report n, bad, msg
    Exit Sub
fail:
    EndSession
    MsgBox "Refresh failed: " & Err.Description, vbCritical, LTX_APP
End Sub

Private Function RefreshOne(shp As Object, ByRef msg As String) As Boolean
    Dim sld As Object
    On Error GoTo bad
    Set sld = shp.Parent
    If shp.Tags(TAG_TYPE) = "text" Then
        LtxRenderText shp, mXl
    Else
        RenderLink sld, shp.Tags(TAG_TYPE), shp.Tags(TAG_WB), shp.Tags(TAG_SHEET), shp.Tags(TAG_RANGE), _
                   shp.Tags(TAG_NAME), (shp.Tags(TAG_BLACK) = "1"), shp
    End If
    RefreshOne = True
    Exit Function
bad:
    msg = msg & vbLf & "Slide " & sld.SlideIndex & " / " & shp.Name & ": " & Err.Description
    LtxLog "Refresh failed: " & shp.Name & " - " & Err.Description
End Function

Private Sub Report(n As Long, bad As Long, msg As String)
    If bad = 0 Then
        If LtxBool("QuietRefresh", True) Then Exit Sub
        MsgBox n & " link(s) refreshed.", vbInformation, LTX_APP
    Else
        MsgBox n & " refreshed, " & bad & " failed:" & msg, vbExclamation, LTX_APP
    End If
End Sub

' ---------------------------------------------------------------------
'  Link maintenance
' ---------------------------------------------------------------------
Public Sub LtxBreakLink()
    Dim shp As Object, n As Long, k As Variant
    On Error Resume Next
    For Each shp In ActiveWindow.Selection.ShapeRange
        If LtxIsLinked(shp) Then
            For Each k In Array(TAG_ID, TAG_TYPE, TAG_WB, TAG_SHEET, TAG_RANGE, TAG_NAME, TAG_BLACK, TAG_SIZEMODE, TAG_UPDATED)
                shp.Tags.Delete CStr(k)
            Next
            n = n + 1
        End If
    Next
    MsgBox n & " link(s) removed. The objects stay as static pictures.", vbInformation, LTX_APP
End Sub

Public Sub LtxSetSizeMode(mode As String)
    Dim shp As Object
    On Error Resume Next
    For Each shp In ActiveWindow.Selection.ShapeRange
        If LtxIsLinked(shp) Then shp.Tags.Add TAG_SIZEMODE, mode
    Next
    LtxSet "SizeMode", mode
End Sub

' Repoint the selected links at another workbook (same sheet/range names).
Public Sub LtxChangeSource()
    Dim shp As Object, p As String, n As Long
    p = InputBox("New workbook full path for the selected link(s):", LTX_APP)
    If Len(p) = 0 Then Exit Sub
    If Not LtxFileExists(p) Then MsgBox "File not found.", vbExclamation, LTX_APP: Exit Sub
    On Error Resume Next
    For Each shp In ActiveWindow.Selection.ShapeRange
        If LtxIsLinked(shp) Then shp.Tags.Add TAG_WB, p: n = n + 1
    Next
    MsgBox n & " link(s) repointed. Refresh to render.", vbInformation, LTX_APP
End Sub

' Open the source workbook and select the linked range.
Public Sub LtxGoToSource()
    Dim shp As Object, xl As Object, wb As Object, rng As Object, openedHere As Boolean, st As Boolean
    On Error GoTo fail
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If Not LtxIsLinked(shp) Then MsgBox "Not a Lattice link.", vbInformation, LTX_APP: Exit Sub
    Set xl = LtxExcel(st)
    Set wb = LtxFindWorkbook(xl, shp.Tags(TAG_WB))
    If wb Is Nothing Then Set wb = xl.Workbooks.Open(shp.Tags(TAG_WB))
    xl.Visible = True
    wb.Activate
    If shp.Tags(TAG_TYPE) = "chart" Then
        wb.Worksheets(shp.Tags(TAG_SHEET)).Activate
        wb.Worksheets(shp.Tags(TAG_SHEET)).ChartObjects(shp.Tags(TAG_RANGE)).Select
    Else
        Set rng = LtxResolveRange(wb, shp.Tags(TAG_SHEET), shp.Tags(TAG_RANGE), shp.Tags(TAG_NAME))
        rng.Parent.Activate
        rng.Select
    End If
    AppActivate xl.Caption
    Exit Sub
fail:
    MsgBox "Could not open source: " & Err.Description, vbExclamation, LTX_APP
End Sub

' Returns a 2-D array (rows x 6): slide, shape name, workbook, sheet, range, updated
Public Function LtxLinkTable() As Variant
    Dim sld As Object, shp As Object, out() As String, n As Long, i As Long
    n = 0
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If LtxIsLinked(shp) Then n = n + 1
        Next
    Next
    If n = 0 Then LtxLinkTable = Empty: Exit Function
    ReDim out(0 To n - 1, 0 To 5)
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If LtxIsLinked(shp) Then
                out(i, 0) = CStr(sld.SlideIndex)
                out(i, 1) = shp.Name
                out(i, 2) = LtxBaseName(shp.Tags(TAG_WB))
                out(i, 3) = shp.Tags(TAG_SHEET)
                out(i, 4) = shp.Tags(TAG_RANGE)
                out(i, 5) = shp.Tags(TAG_UPDATED)
                i = i + 1
            End If
        Next
    Next
    LtxLinkTable = out
End Function

Public Sub LtxSelectLink(slideIndex As Long, shapeName As String)
    On Error Resume Next
    ActiveWindow.View.GotoSlide slideIndex
    ActivePresentation.Slides(slideIndex).Shapes(shapeName).Select
End Sub

' Highlight every linked object with a translucent yellow overlay; run again to clear.
Public Sub LtxHighlightLinks()
    Dim sld As Object, shp As Object, i As Long, n As Long, hl As Object, on_ As Boolean
    on_ = True
    For Each sld In ActivePresentation.Slides
        For i = sld.Shapes.Count To 1 Step -1
            If sld.Shapes(i).Name Like "Lattice highlight*" Then sld.Shapes(i).Delete: on_ = False
        Next
    Next
    If Not on_ Then Exit Sub
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If LtxIsLinked(shp) And Not shp.Name Like "Lattice highlight*" Then
                Set hl = sld.Shapes.AddShape(1, shp.Left - 2, shp.Top - 2, shp.Width + 4, shp.Height + 4)
                hl.Name = "Lattice highlight " & n
                hl.Fill.ForeColor.RGB = RGB(255, 230, 0): hl.Fill.Transparency = 0.6
                hl.Line.ForeColor.RGB = RGB(255, 192, 0): hl.Line.Weight = 2
                hl.TextFrame.TextRange.Text = IIf(shp.Tags(TAG_TYPE) = "text", "T", "")
                n = n + 1
            End If
        Next
    Next
    If n = 0 Then MsgBox "No Lattice links in this deck.", vbInformation, LTX_APP Else ActiveWindow.ViewType = 7   ' slide sorter
End Sub

Public Sub LtxToggleAirplane()
    LtxSet "AirplaneMode", IIf(LtxBool("AirplaneMode"), "0", "1")
    LtxRibbonInvalidate
End Sub

Public Sub LtxToggleBlackFonts()
    LtxSet "BlackFonts", IIf(LtxBool("BlackFonts", True), "0", "1")
    LtxRibbonInvalidate
End Sub
