Attribute VB_Name = "modLtxStyles"
' =====================================================================
'  Lattice for PowerPoint - global shape styles
'  A style is a named bundle of shape formatting (fill, line, font, margins,
'  shadow, autofit) saved to %APPDATA%\Lattice\styles.txt. Applying a style
'  tags the shape (LTX_STYLE = name); "Update all" re-applies every style to
'  every tagged shape in the deck, so a brand change is one click.
' =====================================================================
Option Explicit

Private Const msoTrue = -1
Private Const msoFalse = 0
Public Const TAG_STYLE As String = "LTX_STYLE"

Private Function StylesFile() As String: StylesFile = LtxDataDir() & "\styles.txt": End Function

' ---------- storage: one line per style, name=key:value;key:value ----------
Private Function LoadStyles() As Object
    Dim d As Object, ln As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    For Each ln In Split(Replace(LtxReadFile(StylesFile()), vbCr, ""), vbLf)
        i = InStr(ln, "=")
        If i > 1 Then d(Trim$(Left$(ln, i - 1))) = Mid$(ln, i + 1)
    Next
    Set LoadStyles = d
End Function

Private Sub SaveStyles(d As Object)
    Dim k As Variant, s As String
    For Each k In d.Keys: s = s & k & "=" & d(k) & vbCrLf: Next
    LtxWriteFile StylesFile(), s
End Sub

Private Function Serialize(shp As Object) As String
    Dim s As String, tr As Object
    On Error Resume Next
    s = "fillvis:" & shp.Fill.Visible & ";fill:" & shp.Fill.ForeColor.RGB & ";filltheme:" & shp.Fill.ForeColor.ObjectThemeColor & ";filltrans:" & shp.Fill.Transparency
    s = s & ";linevis:" & shp.Line.Visible & ";line:" & shp.Line.ForeColor.RGB & ";linetheme:" & shp.Line.ForeColor.ObjectThemeColor & ";lineweight:" & shp.Line.Weight & ";linedash:" & shp.Line.DashStyle
    s = s & ";shadow:" & shp.Shadow.Visible
    If shp.HasTextFrame Then
        Set tr = shp.TextFrame.TextRange
        s = s & ";font:" & tr.Font.Name & ";size:" & tr.Font.Size & ";bold:" & tr.Font.Bold & ";italic:" & tr.Font.Italic
        s = s & ";fontcolor:" & tr.Font.Color.RGB & ";fonttheme:" & tr.Font.Color.ObjectThemeColor
        s = s & ";align:" & tr.ParagraphFormat.Alignment & ";vanchor:" & shp.TextFrame.VerticalAnchor
        s = s & ";ml:" & shp.TextFrame.MarginLeft & ";mr:" & shp.TextFrame.MarginRight & ";mt:" & shp.TextFrame.MarginTop & ";mb:" & shp.TextFrame.MarginBottom
        s = s & ";autosize:" & shp.TextFrame2.AutoSize & ";wrap:" & shp.TextFrame.WordWrap & ";bullet:" & tr.ParagraphFormat.Bullet.Visible
        s = s & ";spacewithin:" & tr.ParagraphFormat.SpaceWithin & ";spacebefore:" & tr.ParagraphFormat.SpaceBefore
    End If
    Serialize = s
End Function

Private Sub ApplyTo(shp As Object, spec As String, Optional textToo As Boolean = True)
    Dim kv As Variant, p As Variant, k As String, v As String, tr As Object
    On Error Resume Next
    If shp.HasTextFrame Then Set tr = shp.TextFrame.TextRange
    For Each kv In Split(spec, ";")
        p = Split(kv, ":")
        If UBound(p) >= 1 Then
            k = p(0): v = p(1)
            Select Case k
                Case "fillvis": shp.Fill.Visible = CLng(v)
                Case "filltheme": If CLng(v) > 0 Then shp.Fill.ForeColor.ObjectThemeColor = CLng(v)
                Case "fill": If Not HasTheme(spec, "filltheme") Then shp.Fill.Solid: shp.Fill.ForeColor.RGB = CLng(v)
                Case "filltrans": shp.Fill.Transparency = CSng(v)
                Case "linevis": shp.Line.Visible = CLng(v)
                Case "linetheme": If CLng(v) > 0 Then shp.Line.ForeColor.ObjectThemeColor = CLng(v)
                Case "line": If Not HasTheme(spec, "linetheme") Then shp.Line.ForeColor.RGB = CLng(v)
                Case "lineweight": shp.Line.Weight = CSng(v)
                Case "linedash": shp.Line.DashStyle = CLng(v)
                Case "shadow": shp.Shadow.Visible = CLng(v)
            End Select
            If textToo And Not tr Is Nothing Then
                Select Case k
                    Case "font": tr.Font.Name = v
                    Case "size": tr.Font.Size = CSng(v)
                    Case "bold": tr.Font.Bold = CLng(v)
                    Case "italic": tr.Font.Italic = CLng(v)
                    Case "fonttheme": If CLng(v) > 0 Then tr.Font.Color.ObjectThemeColor = CLng(v)
                    Case "fontcolor": If Not HasTheme(spec, "fonttheme") Then tr.Font.Color.RGB = CLng(v)
                    Case "align": tr.ParagraphFormat.Alignment = CLng(v)
                    Case "vanchor": shp.TextFrame.VerticalAnchor = CLng(v)
                    Case "ml": shp.TextFrame.MarginLeft = CSng(v)
                    Case "mr": shp.TextFrame.MarginRight = CSng(v)
                    Case "mt": shp.TextFrame.MarginTop = CSng(v)
                    Case "mb": shp.TextFrame.MarginBottom = CSng(v)
                    Case "autosize": shp.TextFrame2.AutoSize = CLng(v)
                    Case "wrap": shp.TextFrame.WordWrap = CLng(v)
                    Case "bullet": tr.ParagraphFormat.Bullet.Visible = CLng(v)
                    Case "spacewithin": tr.ParagraphFormat.SpaceWithin = CSng(v)
                    Case "spacebefore": tr.ParagraphFormat.SpaceBefore = CSng(v)
                End Select
            End If
        End If
    Next
End Sub

Private Function HasTheme(spec As String, key As String) As Boolean
    Dim i As Long, v As String
    i = InStr(spec, key & ":")
    If i = 0 Then Exit Function
    v = Mid$(spec, i + Len(key) + 1)
    If InStr(v, ";") > 0 Then v = Left$(v, InStr(v, ";") - 1)
    HasTheme = (Val(v) > 0)
End Function

' ---------- commands ----------
Public Sub LtxStyleSave()
    Dim sr As Object, nm As String, d As Object
    Set sr = LtxSel(): If sr Is Nothing Then MsgBox "Select a shape to save as a style.", vbInformation, LTX_APP: Exit Sub
    nm = InputBox("Style name:", LTX_APP, IIf(Len(sr(1).Tags(TAG_STYLE)) > 0, sr(1).Tags(TAG_STYLE), "Body box"))
    If Len(nm) = 0 Then Exit Sub
    Set d = LoadStyles()
    d(nm) = Serialize(sr(1))
    SaveStyles d
    sr(1).Tags.Add TAG_STYLE, nm
    MsgBox "Style '" & nm & "' saved.", vbInformation, LTX_APP
End Sub

Public Sub LtxStyleApply()
    Dim sr As Object, d As Object, k As Variant, list As String, nm As String, s As Object
    Set sr = LtxSel(): If sr Is Nothing Then MsgBox "Select the shapes first.", vbInformation, LTX_APP: Exit Sub
    Set d = LoadStyles()
    If d.Count = 0 Then MsgBox "No styles saved yet. Use 'Save style' on a formatted shape.", vbInformation, LTX_APP: Exit Sub
    For Each k In d.Keys: list = list & k & vbLf: Next
    nm = LtxGet("LastStyle")
    If Not d.Exists(nm) Then nm = d.Keys()(0)
    nm = InputBox("Styles:" & vbLf & list & vbLf & "Style to apply:", LTX_APP, nm)
    If Len(nm) = 0 Or Not d.Exists(nm) Then Exit Sub
    LtxSet "LastStyle", nm
    For Each s In sr
        ApplyTo s, d(nm)
        s.Tags.Add TAG_STYLE, nm
    Next
End Sub

' Re-apply every style to every tagged shape in the deck
Public Sub LtxStyleUpdateAll()
    Dim d As Object, sld As Object, shp As Object, n As Long, g As Object
    Set d = LoadStyles()
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = 6 Then
                For Each g In shp.GroupItems
                    If d.Exists(g.Tags(TAG_STYLE)) Then ApplyTo g, d(g.Tags(TAG_STYLE)): n = n + 1
                Next
            ElseIf d.Exists(shp.Tags(TAG_STYLE)) Then
                ApplyTo shp, d(shp.Tags(TAG_STYLE)): n = n + 1
            End If
        Next
    Next
    MsgBox n & " shape(s) updated from " & d.Count & " style(s).", vbInformation, LTX_APP
End Sub

' Apply a style to every shape in the deck that looks like the selected one
' (same type and roughly the same size), tagging them on the way.
Public Sub LtxStyleApplyToSimilar()
    Dim sr As Object, a As Object, d As Object, nm As String, sld As Object, shp As Object, n As Long
    Set sr = LtxSel(): If sr Is Nothing Then Exit Sub
    Set a = sr(1)
    nm = a.Tags(TAG_STYLE)
    If Len(nm) = 0 Then MsgBox "Apply or save a style on this shape first.", vbInformation, LTX_APP: Exit Sub
    Set d = LoadStyles()
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = a.Type And shp.AutoShapeType = a.AutoShapeType And shp.HasTextFrame = a.HasTextFrame Then
                If Abs(shp.Height - a.Height) < a.Height * 0.5 Then
                    ApplyTo shp, d(nm): shp.Tags.Add TAG_STYLE, nm: n = n + 1
                End If
            End If
        Next
    Next
    MsgBox n & " similar shape(s) styled as '" & nm & "'.", vbInformation, LTX_APP
End Sub

Public Sub LtxStyleManage()
    If Not LtxFileExists(StylesFile()) Then LtxWriteFile StylesFile(), ""
    Shell "notepad.exe """ & StylesFile() & """", vbNormalFocus
End Sub

Public Sub LtxStyleClearTag()
    Dim sr As Object, s As Object
    Set sr = LtxSel(): If sr Is Nothing Then Exit Sub
    For Each s In sr: s.Tags.Delete TAG_STYLE: Next
End Sub
