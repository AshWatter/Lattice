Attribute VB_Name = "modLtxXlExport"
' =====================================================================
'  Lattice for Excel - export to PowerPoint / Word
'   Copy for PowerPoint : writes a descriptor the PowerPoint/Word add-ins
'                         pick up on their Import button (3 minutes).
'   Export to PowerPoint: pushes the selection straight onto the active
'                         slide; if a shape is selected there, the new
'                         object takes its place and matches its width /
'                         height / size (Export size mode), like a refresh.
'   Export to Word      : same into the running Word at the cursor.
'  Shortcut: Ctrl+Alt+R exports the selection as a picture to PowerPoint.
' =====================================================================
Option Explicit

Private Function DataDir() As String
    DataDir = Environ$("APPDATA") & "\Lattice"
    If Dir(DataDir, vbDirectory) = "" Then MkDir DataDir
End Function

' Make sure the range has a defined name so links survive row/column inserts
Private Function EnsureName(rng As Range) As String
    Dim n As Object, nm As String
    If Not GetSetting("Lattice", "Excel", "AutoName", "1") = "1" Then Exit Function
    For Each n In rng.Parent.Parent.Names
        On Error Resume Next
        If n.RefersToRange.Address(External:=True) = rng.Address(External:=True) Then
            EnsureName = n.Name: Exit Function
        End If
        On Error GoTo 0
    Next
    nm = "ltx_" & Format$(Now, "yymmddhhnnss") & "_" & Int(Rnd * 1000)
    On Error Resume Next
    rng.Parent.Parent.Names.Add nm, "=" & rng.Address(External:=True), False   ' hidden name
    If Err.Number = 0 Then EnsureName = nm
End Function

' Export size mode: width | height | size | none  (Match Size buttons on the ribbon)
Public Function LtxExportSizeMode() As String
    LtxExportSizeMode = XlGet("ExportSizeMode", "width")
End Function
Public Sub LtxSetExportSizeMode(m As String)
    XlSet "ExportSizeMode", m
    Application.StatusBar = "Lattice: export size mode = " & m
End Sub

' Writes the descriptor file. Returns False when nothing exportable is selected.
Private Function WriteDescriptor(kind As String) As Boolean
    Dim rng As Range, wb As Workbook, txt As String, nm As String, sel As Object
    If ActiveWorkbook.Path = "" Then MsgBox "Save the workbook first so the link has a file to point at.", vbExclamation, LTX_APP_XL: Exit Function
    Set sel = Selection
    Set wb = ActiveWorkbook
    If TypeName(sel) = "Range" And kind <> "chart" And kind <> "nativechart" Then
        Set rng = sel
        nm = EnsureName(rng)
        txt = "wb=" & wb.FullName & vbLf & "sheet=" & rng.Parent.Name & vbLf & "range=" & rng.Address(False, False) & vbLf & _
              "name=" & nm & vbLf & "kind=" & kind
        rng.Copy    ' ordinary copy too, so a plain paste still works
    ElseIf Not ActiveChart Is Nothing Then
        txt = "wb=" & wb.FullName & vbLf & "sheet=" & ActiveChart.Parent.Parent.Name & vbLf & "range=" & ActiveChart.Parent.Name & vbLf & "name=" & vbLf & "kind=" & IIf(kind = "nativechart", "nativechart", "chart")
    ElseIf TypeName(sel) = "Range" And ActiveSheet.ChartObjects.Count >= 1 Then
        txt = "wb=" & wb.FullName & vbLf & "sheet=" & ActiveSheet.Name & vbLf & "range=" & NearestChartName(sel) & vbLf & "name=" & vbLf & "kind=" & IIf(kind = "nativechart", "nativechart", "chart")
    Else
        MsgBox "Select a range or a chart.", vbExclamation, LTX_APP_XL: Exit Function
    End If
    txt = txt & vbLf & "sizemode=" & LtxExportSizeMode() & vbLf & "time=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    WriteText DataDir() & "\pending.txt", txt
    WriteDescriptor = True
End Function

Private Function NearestChartName(sel As Range) As String
    Dim co As ChartObject, best As ChartObject, d As Double, bd As Double
    bd = 1E+12
    For Each co In ActiveSheet.ChartObjects
        d = (co.Left - sel.Left) ^ 2 + (co.Top - sel.Top) ^ 2
        If d < bd Then bd = d: Set best = co
    Next
    NearestChartName = best.Name
End Function

Public Sub LtxCopyForPpt(kind As String)
    If WriteDescriptor(kind) Then Application.StatusBar = "Lattice: ready - press Import in PowerPoint or Word (" & kind & ")"
End Sub

Private Sub WriteText(p As String, t As String)
    Dim f As Integer
    f = FreeFile
    Open p For Output As #f
    Print #f, t;
    Close #f
End Sub

' ---------- Export to PowerPoint ----------
Public Sub LtxSendToPpt(kind As String)
    Dim pp As Object
    On Error Resume Next
    Set pp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If pp Is Nothing Then MsgBox "PowerPoint is not running. Open the deck first.", vbExclamation, LTX_APP_XL: Exit Sub
    If Not WriteDescriptor(kind) Then Exit Sub
    Dim try As Long
    On Error Resume Next
    For try = 1 To 4
        Err.Clear
        pp.Run "Lattice.ppam!LtxImport", kind
        If Err.Number = 0 Then Exit For
        If InStr(Err.Description, "message filter") = 0 And InStr(Err.Description, "busy") = 0 Then Exit For
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next
    If Err.Number <> 0 Then GoTo fail
    On Error GoTo fail
    If XlGet("ExportActivate", "1") = "1" Then AppActivate pp.Caption
    Exit Sub
fail:
    MsgBox "PowerPoint add-in not responding: " & Err.Description & vbLf & "Use the Import button in PowerPoint instead.", vbExclamation, LTX_APP_XL
End Sub

Public Sub LtxSendToPptPicture(): LtxSendToPpt "picture": End Sub
Public Sub LtxSendToPptTable(): LtxSendToPpt "table": End Sub
Public Sub LtxSendToPptChart(): LtxSendToPpt "chart": End Sub

' ---------- Export to Word ----------
Public Sub LtxSendToWord(kind As String)
    Dim wd As Object
    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo 0
    If wd Is Nothing Then MsgBox "Word is not running. Open the document first.", vbExclamation, LTX_APP_XL: Exit Sub
    If Not WriteDescriptor(kind) Then Exit Sub
    On Error GoTo fail
    wd.Run "Lattice.modLtxWdLinks.WdImport", kind
    If XlGet("ExportActivate", "1") = "1" Then AppActivate wd.Caption
    Exit Sub
fail:
    MsgBox "Word add-in not responding: " & Err.Description & vbLf & "Use the Import button in Word instead.", vbExclamation, LTX_APP_XL
End Sub

Public Sub LtxSendToWordPicture(): LtxSendToWord "picture": End Sub
Public Sub LtxSendToWordTable(): LtxSendToWord "table": End Sub

' Ask the running PowerPoint to refresh every link (saves this workbook first)
Public Sub LtxRefreshPptFromHere()
    Dim pp As Object
    On Error Resume Next
    Set pp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If pp Is Nothing Then MsgBox "PowerPoint is not running.", vbExclamation, LTX_APP_XL: Exit Sub
    If ActiveWorkbook.Saved = False Then ActiveWorkbook.Save
    On Error GoTo fail
    pp.Run "Lattice.ppam!LtxRefreshAll"
    Exit Sub
fail:
    MsgBox "PowerPoint add-in not responding: " & Err.Description, vbExclamation, LTX_APP_XL
End Sub

Public Sub LtxRefreshWordFromHere()
    Dim wd As Object
    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo 0
    If wd Is Nothing Then MsgBox "Word is not running.", vbExclamation, LTX_APP_XL: Exit Sub
    If ActiveWorkbook.Saved = False Then ActiveWorkbook.Save
    On Error GoTo fail
    wd.Run "Lattice.modLtxWdLinks.WdRefreshAll"
    Exit Sub
fail:
    MsgBox "Word add-in not responding: " & Err.Description, vbExclamation, LTX_APP_XL
End Sub

' Export the selection as a picture file (PNG) next to the workbook
Public Sub LtxExportPictureFile()
    Dim p As String, ch As Object, ws As Object
    If ActiveWorkbook.Path = "" Then MsgBox "Save the workbook first.", vbExclamation, LTX_APP_XL: Exit Sub
    p = ActiveWorkbook.Path & "\" & Replace(Selection.Address(False, False), ":", "-") & ".png"
    Selection.CopyPicture xlScreen, xlBitmap
    Set ws = ActiveSheet
    Set ch = ws.ChartObjects.Add(0, 0, Selection.Width + 4, Selection.Height + 4)
    ch.Chart.Paste
    ch.Chart.Export p, "PNG"
    ch.Delete
    MsgBox "Saved " & p, vbInformation, LTX_APP_XL
End Sub

Public Sub LtxPreviewBlack()
    MsgBox "On import, Lattice turns every non-light font in the range black and restores the original colours afterwards." & vbLf & _
           "Toggle this per deck with the 'Black fonts' button on the PowerPoint tab.", vbInformation, LTX_APP_XL
End Sub

' ---------- Shortcut card (HTML, opens in the browser) ----------
Public Sub LtxShortcutCard()
    Dim p As String, f As Integer, ln As String, i As Long, rows As String, k As String, m As String, grp As String
    p = DataDir() & "\lattice_excel_shortcuts.html"
    If Dir(KeysFilePath()) = "" Then LtxKeysLoad
    f = FreeFile
    Open KeysFilePath() For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        ln = Trim$(ln)
        If Left$(ln, 5) = "# ---" Then
            grp = Trim$(Replace(Replace(ln, "# ---", ""), "---", ""))
            rows = rows & "<tr class='g'><td colspan='2'>" & grp & "</td></tr>"
        ElseIf InStr(ln, "=") > 1 And Left$(ln, 1) <> "#" Then
            k = Trim$(Left$(ln, InStr(ln, "=") - 1)): m = Trim$(Mid$(ln, InStr(ln, "=") + 1))
            rows = rows & "<tr><td>" & Describe(m) & "</td><td>" & KeyHtml(k) & "</td></tr>"
        End If
    Loop
    Close #f
    WriteText p, CardHtml("Excel", rows, _
        "<tr class='g'><td colspan='2'>Ribbon keytips (press Alt, then T, then the letter)</td></tr>" & _
        KeyRow("AutoColor", "T A") & KeyRow("Black Fonts switch", "T Z") & KeyRow("Format menu", "T F") & KeyRow("Model Check pane", "T M") & KeyRow("Library (DCF, LBO, ...)", "T L") & _
        KeyRow("Trace menu (go to precedent / dependent / back)", "T P") & KeyRow("Charts", "T C") & KeyRow("Export to PowerPoint", "T E") & KeyRow("Export to Word", "T W") & _
        KeyRow("Refresh deck", "T R") & KeyRow("Comps stats rows", "T S") & KeyRow("Optimize", "T O") & KeyRow("Shortcuts", "T K"))
    Shell "cmd /c start """" """ & p & """", vbHide
End Sub

Private Function Describe(macro As String) As String
    Dim s As String
    s = Replace(macro, "Ltx", "")
    s = Replace(s, "Cycle", "Cycle ")
    Dim i As Long, out As String, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If i > 1 Then                         ' VBA's And does not short-circuit: Mid$(s, 0, 1) would raise error 5
            If ch Like "[A-Z]" And Mid$(s, i - 1, 1) Like "[a-z]" Then out = out & " "
        End If
        out = out & ch
    Next
    Describe = out
End Function

Public Function KeyHtml(k As String) As String
    Dim s As String, body As String
    s = k
    If InStr(s, "^") > 0 Then body = body & "<span class='k on'>Ctrl</span>" Else body = body & "<span class='k'>Ctrl</span>"
    If InStr(s, "%") > 0 Then body = body & "<span class='k on alt'>Alt</span>" Else body = body & "<span class='k'>Alt</span>"
    If InStr(s, "+") > 0 Then body = body & "<span class='k on sh'>Shift</span>" Else body = body & "<span class='k'>Shift</span>"
    s = Replace(Replace(Replace(s, "^", ""), "%", ""), "+", "")
    s = Replace(Replace(s, "{", ""), "}", "")
    body = body & "<span class='k key'>" & UCase$(s) & "</span>"
    KeyHtml = body
End Function

Public Function KeyRow(label As String, keys As String) As String
    Dim parts As Variant, p As Variant, h As String
    parts = Split(keys, " ")
    For Each p In parts
        If p = "/" Then h = h & " <span class='or'>or</span> " Else h = h & "<span class='k key'>" & p & "</span>"
    Next
    KeyRow = "<tr><td>" & label & "</td><td>" & h & "</td></tr>"
End Function

Public Function CardHtml(app As String, rows As String, extra As String) As String
    CardHtml = "<!doctype html><html><head><meta charset='utf-8'><title>Lattice " & app & " shortcuts</title><style>" & _
        "body{font-family:Segoe UI,Arial,sans-serif;background:#f5f6f8;margin:0;padding:24px;color:#1f2937}" & _
        ".card{background:#fff;max-width:900px;margin:auto;border-radius:14px;box-shadow:0 8px 30px rgba(31,56,100,.12);padding:28px 36px;border-top:4px solid #2e75b6}" & _
        "h1{margin:0 0 4px;font-size:22px;letter-spacing:.5px}h1 span{background:linear-gradient(90deg,#1f3864,#2e75b6);color:#fff;padding:4px 12px;border-radius:6px;margin-left:8px;font-weight:600;font-size:16px;vertical-align:middle}" & _
        "p.sub{color:#6b7280;margin:0 0 18px;font-size:13px}table{width:100%;border-collapse:collapse}td{padding:6px 8px;border-bottom:1px solid #eef0f3;font-size:13px}" & _
        "td:last-child{text-align:right;white-space:nowrap}tr.g td{background:#eef2f7;font-weight:700;color:#1f3864;padding-top:10px;border:0;font-size:12px;text-transform:uppercase;letter-spacing:.6px}" & _
        ".k{display:inline-block;min-width:26px;text-align:center;padding:3px 7px;margin-left:4px;border-radius:5px;background:#e5e7eb;color:#9ca3af;font-size:11px;font-weight:600}" & _
        ".k.on{background:#2563eb;color:#fff}.k.on.alt{background:#f97316}.k.on.sh{background:#16a34a}.k.key{background:#1f3864;color:#fff;min-width:18px}.or{color:#9ca3af;font-size:11px}" & _
        "</style></head><body><div class='card'><h1>Lattice <span>" & UCase$(app) & " SHORTCUTS</span></h1>" & _
        "<p class='sub'>Edit the key bindings under Lattice &gt; Shortcuts &gt; Edit; this card regenerates from that file.</p>" & _
        "<table>" & rows & extra & "</table></div></body></html>"
End Function

' Called by PowerPoint / Word through Application.Run ("Lattice.xlam!LtxServeCopyPicture").
' Runs the black-font picture copy inside Excel, in-process: hides gridlines, paints dark
' fonts black block by block, copies the picture and restores everything (colours, gridlines,
' active sheet, Saved flag). Hundreds of cross-process COM calls become one. Returns "" or the error.
Public Function LtxServeCopyPicture(wbName As Variant, sheetName As Variant, addr As Variant, blackFonts As Variant, asBitmap As Variant, Optional hideGrid As Variant = True) As String
    Dim rng As Range, wb As Workbook, ws As Worksheet, win As Window, prev As Object
    Dim hadGrid As Boolean, saved As Boolean, wasEvents As Boolean, wasScreen As Boolean, k As Long
    Dim addrs As Collection, cols As Collection
    Set addrs = New Collection: Set cols = New Collection
    On Error GoTo fail
    Set wb = Workbooks(CStr(wbName)): Set ws = wb.Worksheets(CStr(sheetName)): Set rng = ws.Range(CStr(addr))
    saved = wb.Saved
    wasEvents = Application.EnableEvents: wasScreen = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    If CBool(hideGrid) Then
        On Error Resume Next
        Set win = wb.Windows(1)
        If Not win Is Nothing Then
            If win.ActiveSheet.Name <> ws.Name Then Set prev = win.ActiveSheet: ws.Activate
            hadGrid = win.DisplayGridlines
            If hadGrid Then win.DisplayGridlines = False
        End If
        On Error GoTo fail
    End If
    If CBool(blackFonts) And Not ws.ProtectContents Then ServeBlacken rng, addrs, cols, 0
    rng.CopyPicture xlScreen, IIf(CBool(asBitmap), xlBitmap, xlPicture)
cleanup:
    On Error Resume Next
    For k = 1 To addrs.Count
        ws.Range(addrs(k)).Font.Color = cols(k)
    Next
    If hadGrid Then win.DisplayGridlines = True
    If Not prev Is Nothing Then prev.Activate
    Application.EnableEvents = wasEvents: Application.ScreenUpdating = wasScreen
    wb.Saved = saved
    Exit Function
fail:
    LtxServeCopyPicture = Err.Description
    Resume cleanup
End Function

' Block-splitting blacken: a block with a uniform Font.Color is handled in one call; only
' mixed blocks are split. Light fonts (white on dark fills) are left alone.
Private Sub ServeBlacken(blk As Range, addrs As Collection, cols As Collection, depth As Long)
    Dim col As Variant, h As Long, w As Long
    col = blk.Font.Color
    If Not IsNull(col) Then
        If CLng(col) <> 0 And Not ServeIsLight(CLng(col)) Then
            addrs.Add blk.Address(False, False): cols.Add CLng(col)
            blk.Font.Color = 0
        End If
        Exit Sub
    End If
    h = blk.Rows.Count: w = blk.Columns.Count
    If h = 1 And w = 1 Then Exit Sub
    If depth > 40 Then Exit Sub
    If h >= w Then
        ServeBlacken blk.Rows(1).Resize(h \ 2), addrs, cols, depth + 1
        ServeBlacken blk.Rows(h \ 2 + 1).Resize(h - h \ 2), addrs, cols, depth + 1
    Else
        ServeBlacken blk.Columns(1).Resize(, w \ 2), addrs, cols, depth + 1
        ServeBlacken blk.Columns(w \ 2 + 1).Resize(, w - w \ 2), addrs, cols, depth + 1
    End If
End Sub

Private Function ServeIsLight(bgr As Long) As Boolean
    Dim r As Long, g As Long, b As Long
    r = bgr And &HFF: g = (bgr \ &H100) And &HFF: b = (bgr \ &H10000) And &HFF
    ServeIsLight = ((0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.75)
End Function
