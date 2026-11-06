Attribute VB_Name = "modLtxWdLinks"
' =====================================================================
'  Lattice for Word - Excel import and refresh
'  The link descriptor lives in the picture's AlternativeText:
'  LTX|kind|workbook|sheet|range|name|black|updated
' =====================================================================
Option Explicit

Private Const wdPasteEnhancedMetafile = 9
Private Const wdPasteHTML = 10
Private Const wdPastePNG = 8   ' not an official constant; PNG falls back to metafile

Private Function Desc(kind As String, wb As String, sh As String, rg As String, nm As String, black As Boolean) As String
    Desc = "LTX|" & kind & "|" & wb & "|" & sh & "|" & rg & "|" & nm & "|" & IIf(black, "1", "0") & "|" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Public Function WdIsLinked(alt As String) As Boolean
    WdIsLinked = (Left$(alt, 4) = "LTX|")
End Function

' ---------------- Import ----------------
Public Sub WdImport(kind As String)
    Dim xl As Object, st As Boolean, src As Object, pend As Object
    Dim wbPath As String, sheetName As String, addr As String, dname As String, black As Boolean
    black = LtxBool("BlackFonts", True)
    On Error GoTo fail
    Set pend = LtxReadPending()
    If pend.Exists("wb") And pend.Exists("time") Then
        If Abs(DateDiff("s", CDate(pend("time")), Now)) < 180 Then
            wbPath = pend("wb"): sheetName = pend("sheet"): addr = pend("range")
            If pend.Exists("name") Then dname = pend("name")
            If pend.Exists("kind") Then kind = pend("kind")
            Kill LtxPendingPath()
        End If
    End If
    Set xl = LtxExcel(st)
    If Len(wbPath) = 0 Then
        If st Then
            xl.Quit
            MsgBox "Excel is not running. Select a range in Excel and try again.", vbExclamation, LTX_APP
            Exit Sub
        End If
        Set src = xl.Selection
        If TypeName(src) = "Range" Then
            wbPath = src.Parent.Parent.FullName: sheetName = src.Parent.Name: addr = src.Address(False, False)
        ElseIf Not xl.ActiveChart Is Nothing Then
            wbPath = xl.ActiveChart.Parent.Parent.Parent.FullName: sheetName = xl.ActiveChart.Parent.Parent.Name
            addr = xl.ActiveChart.Parent.Name: kind = "chart"
        Else
            MsgBox "Select a range or chart in Excel first.", vbExclamation, LTX_APP: Exit Sub
        End If
    End If
    Render Selection.Range, kind, wbPath, sheetName, addr, dname, black, xl
    If st Then xl.Quit
    Exit Sub
fail:
    MsgBox "Import failed: " & Err.Description, vbCritical, LTX_APP
End Sub

' Renders into the given range (replacing its contents)
Private Sub Render(target As Object, kind As String, wbPath As String, sheetName As String, addr As String, _
                   dname As String, black As Boolean, xl As Object, Optional oldWidth As Single = 0)
    Dim wb As Object, openedHere As Boolean, rng As Object, ils As Object, t As Object, r As Long, c As Long
    Set wb = LtxGetWorkbook(xl, wbPath, openedHere)
    If wb Is Nothing Then Err.Raise vbObjectError + 1, , "Workbook not available: " & wbPath
    If kind = "chart" Then
        LtxCopyChartPicture wb, sheetName, addr
    Else
        Set rng = LtxResolveRange(wb, sheetName, addr, dname)
        addr = rng.Address(False, False)
        If kind = "table" Then
            rng.Copy
        ElseIf Not LtxServedCopy(xl, wb, sheetName, rng, black, False) Then
            LtxCopyRangePicture rng, black
        End If
    End If
    Dim startPos As Long, doc As Object, x As Object
    startPos = target.Start
    Set doc = target.Document
    If kind = "table" Then
        target.PasteSpecial DataType:=wdPasteHTML
        Set t = Nothing
        For Each x In doc.Tables
            If x.Range.Start >= startPos Then
                If t Is Nothing Then Set t = x Else If x.Range.Start < t.Range.Start Then Set t = x
            End If
        Next
        If t Is Nothing Then Err.Raise vbObjectError + 3, , "Pasted table not found"
        t.Title = Desc(kind, wbPath, sheetName, addr, dname, black)   ' descriptor on the table
        If black Then
            For r = 1 To t.Rows.Count
                For c = 1 To t.Columns.Count
                    On Error Resume Next
                    If Not LtxIsLightColor(t.Cell(r, c).Range.Font.Color) Then t.Cell(r, c).Range.Font.Color = 0
                    On Error GoTo 0
                Next
            Next
        End If
    Else
        Dim nInline As Long, nFloat As Long
        nInline = doc.InlineShapes.Count: nFloat = doc.Shapes.Count
        On Error Resume Next
        target.PasteSpecial DataType:=wdPasteEnhancedMetafile, Placement:=0   ' wdInLine
        If Err.Number <> 0 Then Err.Clear: target.Paste
        On Error GoTo 0
        Set ils = Nothing
        For Each x In doc.InlineShapes
            If x.Range.Start >= startPos Then
                If ils Is Nothing Then Set ils = x Else If x.Range.Start < ils.Range.Start Then Set ils = x
            End If
        Next
        If ils Is Nothing And doc.Shapes.Count > nFloat Then
            ' Word pasted a floating picture: convert the newest one to inline
            Set ils = doc.Shapes(doc.Shapes.Count).ConvertToInlineShape
        End If
        If ils Is Nothing And doc.InlineShapes.Count > nInline Then Set ils = doc.InlineShapes(doc.InlineShapes.Count)
        If ils Is Nothing Then Err.Raise vbObjectError + 4, , "Pasted picture not found (inline " & doc.InlineShapes.Count & ", floating " & doc.Shapes.Count & ")"
        ils.AlternativeText = Desc(kind, wbPath, sheetName, addr, dname, black)
        ils.LockAspectRatio = -1
        If oldWidth > 0 Then
            ils.Width = oldWidth
        ElseIf ils.Width > target.Document.PageSetup.PageWidth - target.Document.PageSetup.LeftMargin - target.Document.PageSetup.RightMargin Then
            ils.Width = target.Document.PageSetup.PageWidth - target.Document.PageSetup.LeftMargin - target.Document.PageSetup.RightMargin
        End If
    End If
    xl.CutCopyMode = False
    If openedHere Then wb.Close SaveChanges:=False
End Sub

' ---------------- Refresh ----------------
Public Sub WdRefreshAll()
    RefreshSet ActiveDocument.InlineShapes, ActiveDocument.Tables
End Sub

Public Sub WdRefreshSelected()
    RefreshSet Selection.Range.InlineShapes, Selection.Range.Tables
End Sub

Private Sub RefreshSet(ilsCol As Object, tblCol As Object)
    Dim xl As Object, st As Boolean, i As Long, ils As Object, p As Variant, w As Single, rg As Object, n As Long, bad As Long
    Dim list As Collection, t As Object, msg As String
    Set list = New Collection
    For Each ils In ilsCol
        If WdIsLinked(ils.AlternativeText) Then list.Add ils
    Next
    For Each t In tblCol
        If WdIsLinked(t.Title) Then list.Add t
    Next
    If list.Count = 0 Then MsgBox "No Lattice links found.", vbInformation, LTX_APP: Exit Sub
    Set xl = LtxExcel(st)
    xl.ScreenUpdating = False
    Application.ScreenUpdating = False
    On Error Resume Next
    For i = 1 To list.Count
        Err.Clear
        If TypeName(list(i)) = "Table" Then
            p = Split(list(i).Title, "|")
            Set rg = list(i).Range
            list(i).Delete
            Render rg, "table", CStr(p(2)), CStr(p(3)), CStr(p(4)), CStr(p(5)), (p(6) = "1"), xl
        Else
            p = Split(list(i).AlternativeText, "|")
            w = list(i).Width
            Set rg = list(i).Range
            list(i).Delete                 ' PasteSpecial over an inline shape inserts next to it, so remove it first
            Render rg, CStr(p(1)), CStr(p(2)), CStr(p(3)), CStr(p(4)), CStr(p(5)), (p(6) = "1"), xl, w
        End If
        If Err.Number <> 0 Then bad = bad + 1: msg = msg & vbLf & Err.Description Else n = n + 1
    Next
    On Error GoTo 0
    Application.ScreenUpdating = True
    xl.ScreenUpdating = True
    If st Then xl.Quit
    If bad > 0 Then MsgBox n & " refreshed, " & bad & " failed:" & msg, vbExclamation, LTX_APP _
    Else If Not LtxBool("QuietRefresh", True) Then MsgBox n & " link(s) refreshed.", vbInformation, LTX_APP
End Sub

Public Sub WdBreakLink()
    Dim ils As Object, t As Object, n As Long
    For Each ils In Selection.Range.InlineShapes
        If WdIsLinked(ils.AlternativeText) Then ils.AlternativeText = "": n = n + 1
    Next
    For Each t In Selection.Range.Tables
        If WdIsLinked(t.Title) Then t.Title = "": n = n + 1
    Next
    MsgBox n & " link(s) removed.", vbInformation, LTX_APP
End Sub

Public Sub WdGoToSource()
    Dim ils As Object, p As Variant, xl As Object, st As Boolean, wb As Object, rng As Object
    If Selection.Range.InlineShapes.Count = 0 Then MsgBox "Select a linked picture.", vbInformation, LTX_APP: Exit Sub
    Set ils = Selection.Range.InlineShapes(1)
    If Not WdIsLinked(ils.AlternativeText) Then MsgBox "Not a Lattice link.", vbInformation, LTX_APP: Exit Sub
    p = Split(ils.AlternativeText, "|")
    Set xl = LtxExcel(st)
    Set wb = LtxFindWorkbook(xl, CStr(p(2)))
    If wb Is Nothing Then Set wb = xl.Workbooks.Open(CStr(p(2)))
    xl.Visible = True: wb.Activate
    If p(1) = "chart" Then
        wb.Worksheets(CStr(p(3))).Activate
    Else
        Set rng = LtxResolveRange(wb, CStr(p(3)), CStr(p(4)), CStr(p(5)))
        rng.Parent.Activate: rng.Select
    End If
    AppActivate xl.Caption
End Sub

Public Sub WdListLinks()
    Dim ils As Object, t As Object, s As String, p As Variant, n As Long
    For Each ils In ActiveDocument.InlineShapes
        If WdIsLinked(ils.AlternativeText) Then
            p = Split(ils.AlternativeText, "|"): n = n + 1
            s = s & n & ". " & LtxBaseName(CStr(p(2))) & " / " & p(3) & "!" & p(4) & "  (" & p(7) & ")" & vbLf
        End If
    Next
    For Each t In ActiveDocument.Tables
        If WdIsLinked(t.Title) Then
            p = Split(t.Title, "|"): n = n + 1
            s = s & n & ". [table] " & LtxBaseName(CStr(p(2))) & " / " & p(3) & "!" & p(4) & "  (" & p(7) & ")" & vbLf
        End If
    Next
    If n = 0 Then s = "No Lattice links in this document."
    MsgBox s, vbInformation, LTX_APP & " - Links"
End Sub

' ---------------- Small document tools ----------------
Public Sub WdStripDoubleSpaces()
    With ActiveDocument.Content.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = "  ": .Replacement.Text = " "
        .Execute Replace:=2
    End With
End Sub

Public Sub WdTablesAutofit()
    Dim t As Object
    For Each t In ActiveDocument.Tables: t.AutoFitBehavior 2: Next
End Sub

Public Sub WdPicturesToWidth()
    Dim ils As Object, w As Single
    w = ActiveDocument.PageSetup.PageWidth - ActiveDocument.PageSetup.LeftMargin - ActiveDocument.PageSetup.RightMargin
    For Each ils In ActiveDocument.InlineShapes
        ils.LockAspectRatio = -1
        If ils.Width > w Then ils.Width = w
    Next
End Sub

Public Sub WdKeytipCard()
    Dim r As String
    r = LtxCardGroup("Import and links") & _
        LtxCardRow("Import Excel selection (black fonts)", "Alt L E") & LtxCardRow("Refresh selection", "Alt L R") & _
        LtxCardRow("Links list", "Alt L L") & LtxCardRow("Airplane Mode", "Alt L A") & _
        LtxCardGroup("Document") & _
        LtxCardRow("Strip double spaces", "Alt L S") & LtxCardRow("Autofit tables", "Alt L F") & LtxCardRow("Fit pictures to margins", "Alt L P") & _
        LtxCardGroup("From Excel") & _
        LtxCardRow("Export selection to Word (picture)", "Ctrl Alt W")
    LtxOpenHtml LtxDataDir() & "\lattice_word_keytips.html", LtxCardHtml("Word", r)
End Sub

Public Sub WdToggleBlack()
    LtxSet "BlackFonts", IIf(LtxBool("BlackFonts", True), "0", "1")
End Sub
