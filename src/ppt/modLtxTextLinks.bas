Attribute VB_Name = "modLtxTextLinks"
' =====================================================================
'  Lattice for PowerPoint - dynamic text links
'  A text shape can be linked to one or more Excel cells. The shape keeps a
'  template in its tags, e.g.  "EV/EBITDA of {0}x in {1}"  plus the cell
'  behind each placeholder. Refresh re-renders the template with the cells'
'  displayed text (so Excel's number formats carry over).
'
'  Tags: LTX_TYPE = "text", LTX_TEMPLATE, LTX_CELLS ("wb|sheet|addr|name" per
'        placeholder, separated by ;;), plus LTX_WB (first workbook) so the
'        shape shows up in the Explorer pane like any other link.
' =====================================================================
Option Explicit

Public Const TAG_TEMPLATE As String = "LTX_TEMPLATE"
Public Const TAG_CELLS As String = "LTX_CELLS"

' Link the whole text of the selected shape to the active Excel cell
Public Sub LtxLinkTextWhole()
    LinkText False
End Sub

' Link only the currently selected text (inside a longer sentence)
Public Sub LtxLinkTextSelection()
    LinkText True
End Sub

Private Sub LinkText(selectionOnly As Boolean)
    Dim shp As Object, xl As Object, st As Boolean, cell As Object, tr As Object
    Dim template As String, cells As String, idx As Long, selStart As Long, selLen As Long, cur As String
    On Error GoTo fail
    If ActiveWindow.Selection.Type = 3 Then
        Set shp = ActiveWindow.Selection.ShapeRange(1)
        Set tr = ActiveWindow.Selection.TextRange
    ElseIf ActiveWindow.Selection.Type = 2 Then
        Set shp = ActiveWindow.Selection.ShapeRange(1)
        If Not shp.HasTextFrame Then MsgBox "Select a text shape.", vbInformation, LTX_APP: Exit Sub
        Set tr = shp.TextFrame.TextRange
        selectionOnly = False
    Else
        MsgBox "Select a text shape (or some text inside one) first.", vbInformation, LTX_APP: Exit Sub
    End If
    Set xl = LtxExcel(st)
    If st Then xl.Quit: MsgBox "Excel is not running. Select the source cell in Excel first.", vbExclamation, LTX_APP: Exit Sub
    Set cell = xl.ActiveCell
    If cell Is Nothing Then MsgBox "Select the source cell in Excel first.", vbExclamation, LTX_APP: Exit Sub
    If cell.Parent.Parent.Path = "" Then MsgBox "Save the workbook first.", vbExclamation, LTX_APP: Exit Sub

    template = shp.Tags(TAG_TEMPLATE)
    cells = shp.Tags(TAG_CELLS)
    cur = shp.TextFrame.TextRange.Text
    If Len(template) = 0 Then template = cur
    idx = CountPlaceholders(template)

    If selectionOnly Then
        selStart = tr.Start: selLen = tr.Length
        ' map the selection into the template: the template equals the current text with
        ' placeholders expanded, so we rebuild it from the current text instead.
        template = Left$(cur, selStart - 1) & "{" & idx & "}" & Mid$(cur, selStart + selLen)
        template = ReinsertPlaceholders(template, cur, shp)
    Else
        template = "{" & idx & "}"
        cells = ""
        idx = 0
    End If
    cells = cells & IIf(Len(cells) > 0, ";;", "") & cell.Parent.Parent.FullName & "|" & cell.Parent.Name & "|" & cell.Address(False, False) & "|" & EnsureCellName(cell)

    With shp.Tags
        .Add TAG_TYPE, "text"
        .Add TAG_TEMPLATE, template
        .Add TAG_CELLS, cells
        .Add TAG_WB, cell.Parent.Parent.FullName
        .Add TAG_SHEET, cell.Parent.Name
        .Add TAG_RANGE, cell.Address(False, False)
        .Add TAG_ID, IIf(Len(.Item(TAG_ID)) > 0, .Item(TAG_ID), LtxGuid())
    End With
    LtxRenderText shp, xl
    Exit Sub
fail:
    MsgBox "Text link failed: " & Err.Description, vbCritical, LTX_APP
End Sub

' When the user links a second substring, the current text already contains the
' rendered value of earlier placeholders. Put those placeholders back so the
' template stays generic. Earlier rendered values are stored in LTX_LASTVALS.
Private Function ReinsertPlaceholders(template As String, cur As String, shp As Object) As String
    Dim vals As Variant, i As Long, t As String
    t = template
    If Len(shp.Tags("LTX_LASTVALS")) > 0 Then
        vals = Split(shp.Tags("LTX_LASTVALS"), ";;")
        For i = 0 To UBound(vals)
            If Len(vals(i)) > 0 Then t = Replace(t, vals(i), "{" & i & "}", 1, 1)
        Next
    End If
    ReinsertPlaceholders = t
End Function

Private Function CountPlaceholders(t As String) As Long
    Dim i As Long
    Do While InStr(t, "{" & i & "}") > 0
        i = i + 1
    Loop
    CountPlaceholders = i
End Function

Private Function EnsureCellName(cell As Object) As String
    Dim nm As String
    On Error Resume Next
    nm = "ltx_" & Format$(Now, "yymmddhhnnss") & "_" & Int(Rnd * 1000)
    cell.Parent.Parent.Names.Add nm, "=" & cell.Address(External:=True), False
    If Err.Number = 0 Then EnsureCellName = nm
End Function

' Render a text link. xl may be Nothing (a session is opened then).
Public Sub LtxRenderText(shp As Object, Optional xl As Object)
    Dim parts As Variant, i As Long, p As Variant, wb As Object, rng As Object, txt As String, v As String
    Dim lastVals As String, openedHere As Boolean, st As Boolean, ownXl As Boolean
    If xl Is Nothing Then Set xl = LtxExcel(st): ownXl = True
    parts = Split(shp.Tags(TAG_CELLS), ";;")
    txt = shp.Tags(TAG_TEMPLATE)
    For i = 0 To UBound(parts)
        p = Split(parts(i), "|")
        If UBound(p) >= 2 Then
            Set wb = LtxGetWorkbook(xl, CStr(p(0)), openedHere)
            If wb Is Nothing Then Err.Raise vbObjectError + 2, , "Workbook not available: " & p(0)
            Set rng = LtxResolveRange(wb, CStr(p(1)), CStr(p(2)), IIf(UBound(p) >= 3, CStr(p(3)), ""))
            v = CStr(rng.Cells(1).Text)
            txt = Replace(txt, "{" & i & "}", v)
            lastVals = lastVals & IIf(i > 0, ";;", "") & v
            If openedHere Then wb.Close SaveChanges:=False
        End If
    Next
    ' Replace the text but keep the formatting of the first run
    With shp.TextFrame.TextRange
        If .Text <> txt Then .Text = txt
    End With
    shp.Tags.Add "LTX_LASTVALS", lastVals
    shp.Tags.Add TAG_UPDATED, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    If ownXl And st Then xl.Quit
End Sub

' Show the template and cells behind the selected text link
Public Sub LtxTextLinkInfo()
    Dim shp As Object
    On Error Resume Next
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If shp Is Nothing Then Exit Sub
    If shp.Tags(TAG_TYPE) <> "text" Then MsgBox "Not a text link.", vbInformation, LTX_APP: Exit Sub
    MsgBox "Template: " & shp.Tags(TAG_TEMPLATE) & vbLf & vbLf & "Cells:" & vbLf & Replace(shp.Tags(TAG_CELLS), ";;", vbLf) & vbLf & vbLf & "Updated: " & shp.Tags(TAG_UPDATED), vbInformation, LTX_APP & " - Text link"
End Sub

' Make the current text the new template (after manual edits), keeping the cells
Public Sub LtxTextLinkRetemplate()
    Dim shp As Object, t As String
    On Error Resume Next
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If shp.Tags(TAG_TYPE) <> "text" Then Exit Sub
    t = InputBox("Edit the template. {0}, {1}... are the linked cells in order:", LTX_APP, shp.Tags(TAG_TEMPLATE))
    If Len(t) = 0 Then Exit Sub
    shp.Tags.Add TAG_TEMPLATE, t
    LtxRenderText shp
End Sub
