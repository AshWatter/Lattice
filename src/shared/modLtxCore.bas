Attribute VB_Name = "modLtxCore"
' =====================================================================
'  Lattice - shared core (used by the PowerPoint and Word add-ins)
'  Settings, Excel-instance acquisition, link descriptors, helpers.
'  Excel is always late-bound (Object) so no reference is required.
' =====================================================================
Option Explicit
' Served-copy state (see LtxServedCopy): the Excel instance checked last and whether its add-in answered
Private mServeApp As Object
Private mServeOk As Boolean

' Windows API used to enumerate every running Excel instance. GetObject only returns the first
' instance registered with Windows, which may be a hidden automation instance (a test, a stalled
' export) rather than the user's window; that made PowerPoint import from the wrong Excel.
Private Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hWnd As LongPtr, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long

Public Const LTX_APP As String = "Lattice"
Public Const LTX_VERSION As String = "0.1.0"

' Tag keys stored on linked shapes
Public Const TAG_ID As String = "LTX_ID"
Public Const TAG_TYPE As String = "LTX_TYPE"       ' picture | table | chart
Public Const TAG_WB As String = "LTX_WB"           ' full path of workbook
Public Const TAG_SHEET As String = "LTX_SHEET"
Public Const TAG_RANGE As String = "LTX_RANGE"     ' A1 address or chart name
Public Const TAG_NAME As String = "LTX_NAME"       ' defined name (optional)
Public Const TAG_BLACK As String = "LTX_BLACK"     ' "1" = force black fonts
Public Const TAG_UPDATED As String = "LTX_UPDATED"
Public Const TAG_SIZEMODE As String = "LTX_SIZEMODE" ' keep | width | height | source

' ---------- Settings (HKCU\Software\VB and VBA Program Settings\Lattice) ----------
Public Function LtxGet(key As String, Optional def As String = "") As String
    LtxGet = GetSetting(LTX_APP, "Settings", key, def)
End Function

Public Sub LtxSet(key As String, val As String)
    SaveSetting LTX_APP, "Settings", key, val
End Sub

Public Function LtxBool(key As String, Optional def As Boolean = False) As Boolean
    LtxBool = (LtxGet(key, IIf(def, "1", "0")) = "1")
End Function

Public Function LtxDataDir() As String
    Dim p As String
    p = Environ$("APPDATA") & "\Lattice"
    If Dir(p, vbDirectory) = "" Then MkDir p
    LtxDataDir = p
End Function

' ---------- Small helpers ----------
Public Function LtxGuid() As String
    Dim s As String, i As Long
    Randomize
    For i = 1 To 16
        s = s & Hex$(Int(Rnd * 16))
    Next
    LtxGuid = Format$(Now, "yyyymmddhhnnss") & "-" & s
End Function

Public Function LtxFileExists(p As String) As Boolean
    On Error Resume Next
    LtxFileExists = (Len(p) > 0) And (Dir(p, vbNormal) <> "")
End Function

Public Function LtxBaseName(p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    LtxBaseName = Mid$(p, i + 1)
End Function

Public Function LtxReadFile(p As String) As String
    Dim f As Integer
    If Not LtxFileExists(p) Then Exit Function
    f = FreeFile
    Open p For Input As #f
    LtxReadFile = Input$(LOF(f), f)
    Close #f
End Function

Public Sub LtxWriteFile(p As String, txt As String)
    Dim f As Integer
    f = FreeFile
    Open p For Output As #f
    Print #f, txt;
    Close #f
End Sub

' Parse "key=value" lines into a Dictionary-like Collection (keys are lower case)
Public Function LtxParseKV(txt As String) As Object
    Dim d As Object, ln As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For Each ln In Split(Replace(txt, vbCr, ""), vbLf)
        i = InStr(ln, "=")
        If i > 1 Then d(Trim$(Left$(ln, i - 1))) = Mid$(ln, i + 1)
    Next
    Set LtxParseKV = d
End Function

' ---------- Excel instance handling ----------
' Returns the running Excel instance, or starts one (hidden) if none exists.
' startedNew is set when Lattice created the instance so the caller can quit it.
Public Function LtxExcel(Optional ByRef startedNew As Boolean, Optional preferPath As String = "") As Object
    Dim xl As Object
    startedNew = False
    On Error Resume Next
    Set xl = LtxPickExcel(preferPath)            ' the user's visible Excel, not merely the first one registered
    If xl Is Nothing Then Set xl = GetObject(, "Excel.Application")
    On Error GoTo 0
    If xl Is Nothing Then
        Set xl = CreateObject("Excel.Application")
        xl.Visible = False
        startedNew = True
    End If
    Set LtxExcel = xl
End Function

' Finds an open workbook by full path (or by file name as fallback). Looks in the given Excel
' instance first, then in every other running instance (the file may be open in the user's
' window while xl is a hidden one). Returns Nothing when it is not open anywhere.
Public Function LtxFindWorkbook(xl As Object, fullPath As String) As Object
    Dim other As Object
    Set LtxFindWorkbook = FindWorkbookIn(xl, fullPath)
    If LtxFindWorkbook Is Nothing Then
        Set other = LtxPickExcel(fullPath)
        If Not other Is Nothing Then Set LtxFindWorkbook = FindWorkbookIn(other, fullPath)
    End If
End Function

Private Function FindWorkbookIn(xl As Object, fullPath As String) As Object
    Dim wb As Object
    On Error Resume Next
    For Each wb In xl.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set FindWorkbookIn = wb: Exit Function
        End If
    Next
    For Each wb In xl.Workbooks
        If StrComp(wb.Name, LtxBaseName(fullPath), vbTextCompare) = 0 Then
            Set FindWorkbookIn = wb: Exit Function
        End If
    Next
End Function

' Opens the workbook if needed. openedHere is True when Lattice opened it, so the
' caller can close it again without saving.
Public Function LtxGetWorkbook(xl As Object, fullPath As String, ByRef openedHere As Boolean) As Object
    Dim wb As Object
    openedHere = False
    Set wb = LtxFindWorkbook(xl, fullPath)
    If wb Is Nothing Then
        If Not LtxFileExists(fullPath) Then Exit Function
        If LtxBool("AirplaneMode") Then Exit Function
        xl.DisplayAlerts = False
        Set wb = xl.Workbooks.Open(fullPath, UpdateLinks:=0, ReadOnly:=True)
        xl.DisplayAlerts = True
        openedHere = True
    End If
    Set LtxGetWorkbook = wb
End Function

' Resolve the range for a link: prefer the defined name (survives row inserts),
' fall back to the stored address.
Public Function LtxResolveRange(wb As Object, sheetName As String, addr As String, definedName As String) As Object
    Dim r As Object
    On Error Resume Next
    If Len(definedName) > 0 Then
        Set r = wb.Names(definedName).RefersToRange
    End If
    If r Is Nothing Then Set r = wb.Worksheets(sheetName).Range(addr)
    On Error GoTo 0
    Set LtxResolveRange = r
End Function

' ---------- Black-font rendering ----------
' Temporarily paints the fonts in rng black (light fonts stay as they are),
' copies the range as a picture to the clipboard, then restores the colours.
' Nothing is left modified and the workbook's Saved flag is preserved.
Public Sub LtxCopyRangePicture(rng As Object, blackFonts As Boolean, Optional asBitmap As Boolean = False)
    Dim xl As Object, wb As Object, saved As Boolean, wasEvents As Boolean, wasScreen As Boolean
    Dim colors() As Variant, n As Long, c As Object, i As Long, protectedSheet As Boolean
    Dim win As Object, hadGrid As Boolean, hideGrid As Boolean, w As Object
    Const xlScreen = 1, xlPicture = -4147, xlBitmap = 2
    Set xl = rng.Application
    Set wb = rng.Parent.Parent
    saved = wb.Saved
    wasEvents = xl.EnableEvents: wasScreen = xl.ScreenUpdating
    xl.EnableEvents = False: xl.ScreenUpdating = False
    protectedSheet = rng.Parent.ProtectContents
    ' gridlines are part of an "as shown on screen" picture: hide them during the copy
    hideGrid = LtxBool("HideGridlines", True)
    If hideGrid Then
        On Error Resume Next
        For Each w In wb.Windows
            If w.Index = 1 Then Set win = w
        Next
        If Not win Is Nothing Then
            If win.ActiveSheet.Name <> rng.Parent.Name Then rng.Parent.Activate
            hadGrid = win.DisplayGridlines
            If hadGrid Then win.DisplayGridlines = False
        End If
        On Error GoTo 0
    End If
    On Error GoTo cleanup
    ' Block-splitting black pass: a block whose Font.Color is uniform is handled in one
    ' call; only mixed blocks are split. Typical tables need a few dozen calls instead of
    ' two per cell, which is what made refresh slow.
    Dim saveAddr As Collection, saveCol As Collection, k As Long
    Set saveAddr = New Collection: Set saveCol = New Collection
    If blackFonts And Not protectedSheet Then
        BlackenBlock rng, saveAddr, saveCol, 0
        n = saveAddr.Count
    End If
    rng.CopyPicture xlScreen, IIf(asBitmap, xlBitmap, xlPicture)
cleanup:
    If blackFonts And Not protectedSheet And n > 0 Then
        For k = 1 To saveAddr.Count
            rng.Worksheet.Range(saveAddr(k)).Font.Color = saveCol(k)
        Next
    End If
    On Error Resume Next
    If hideGrid And hadGrid Then win.DisplayGridlines = True
    xl.EnableEvents = wasEvents: xl.ScreenUpdating = wasScreen
    wb.Saved = saved
    If Err.Number <> 0 Then Err.Raise Err.Number, "LtxCopyRangePicture", Err.Description
End Sub

' True when any cell in the range shows an error value (#DIV/0!, #REF!, ...)
' Reads the whole block with one Value2 call (hundreds of per-cell COM calls otherwise).
Public Function LtxRangeHasErrors(rng As Object) As String
    Dim v As Variant, i As Long, j As Long, n As Long, first As String
    On Error Resume Next
    v = rng.Value2
    If Not IsArray(v) Then
        If IsError(v) Then LtxRangeHasErrors = "1 error cell: " & rng.Address(False, False) & " " & rng.Text
        Exit Function
    End If
    For i = LBound(v, 1) To UBound(v, 1)
        For j = LBound(v, 2) To UBound(v, 2)
            If IsError(v(i, j)) Then
                n = n + 1
                If Len(first) = 0 Then first = rng.Cells(i, j).Address(False, False) & " " & rng.Cells(i, j).Text
            End If
        Next
    Next
    If n > 0 Then LtxRangeHasErrors = n & " error cell(s), e.g. " & first
End Function

' ---------- Served copy (fast path for PowerPoint / Word) ----------
' Every COM call from PowerPoint or Word into Excel crosses a process boundary (~1 ms each),
' so the black-font pass over a large table dominated refresh time. When the Lattice Excel
' add-in is loaded in the source Excel instance, the whole copy runs inside Excel through
' Application.Run instead. Falls back to the cross-process path when it is not available.

Public Function LtxServedCopy(xl As Object, wb As Object, sheetName As String, rng As Object, blackFonts As Boolean, asBitmap As Boolean) As Boolean
    Dim res As Variant, sameApp As Boolean, app As Object
    On Error Resume Next
    Set app = wb.Application                  ' the instance that owns the workbook (may differ from xl)
    If app Is Nothing Then Set app = xl
    If Not mServeApp Is Nothing Then sameApp = (mServeApp Is app)
    If Not sameApp Then
        Set mServeApp = app
        mServeOk = ServeAvailable(app)
    End If
    If Not mServeOk Then Exit Function
    Err.Clear
    res = app.Run("Lattice.xlam!LtxServeCopyPicture", wb.Name, sheetName, rng.Address(False, False), blackFonts, asBitmap, LtxBool("HideGridlines", True))
    If Err.Number <> 0 Then
        LtxLog "Served copy unavailable: " & Err.Description
        mServeOk = False
        Exit Function
    End If
    If Len(CStr(res)) > 0 Then LtxLog "Served copy failed: " & res: Exit Function
    LtxServedCopy = True
End Function

' The Excel add-in is loaded in the user's Excel; an Excel that Lattice started itself
' (or an automation instance) does not load add-ins, so load it from the AddIns folder.
Private Function ServeAvailable(xl As Object) As Boolean
    Dim p As String, w As Object
    On Error Resume Next
    Set w = xl.Workbooks("Lattice.xlam")
    If Not w Is Nothing Then ServeAvailable = True: Exit Function
    p = Environ$("APPDATA") & "\Microsoft\AddIns\Lattice.xlam"
    If Len(Dir$(p)) = 0 Then Exit Function
    Set w = xl.Workbooks.Open(p)
    ServeAvailable = Not w Is Nothing
End Function

' Recursive block blackening (see LtxCopyRangePicture). Records what it changed.
Public Sub BlackenBlock(blk As Object, saveAddr As Collection, saveCol As Collection, depth As Long)
    Dim col As Variant, h As Long, w As Long
    col = blk.Font.Color
    If Not IsNull(col) Then
        If CLng(col) <> 0 And Not LtxIsLightColor(CLng(col)) Then
            saveAddr.Add blk.Address(False, False): saveCol.Add CLng(col)
            blk.Font.Color = 0
        End If
        Exit Sub
    End If
    h = blk.Rows.Count: w = blk.Columns.Count
    If h = 1 And w = 1 Then Exit Sub           ' mixed colours inside one cell: leave as is
    If depth > 40 Then Exit Sub
    If h >= w Then
        BlackenBlock blk.Rows(1).Resize(h \ 2), saveAddr, saveCol, depth + 1
        BlackenBlock blk.Rows(h \ 2 + 1).Resize(h - h \ 2), saveAddr, saveCol, depth + 1
    Else
        BlackenBlock blk.Columns(1).Resize(, w \ 2), saveAddr, saveCol, depth + 1
        BlackenBlock blk.Columns(w \ 2 + 1).Resize(, w - w \ 2), saveAddr, saveCol, depth + 1
    End If
End Sub

' BGR long -> perceived luminance; "light" fonts (white on dark fills) are kept.
Public Function LtxIsLightColor(bgr As Long) As Boolean
    Dim r As Long, g As Long, b As Long, lum As Double
    r = bgr And &HFF: g = (bgr \ &H100) And &HFF: b = (bgr \ &H10000) And &HFF
    lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    LtxIsLightColor = (lum > 0.75)
End Function

' Copies a chart object as a picture to the clipboard (chart border hidden for the copy).
Public Sub LtxCopyChartPicture(wb As Object, sheetName As String, chartName As String)
    Const xlScreen = 1, xlPicture = -4147
    Dim ch As Object, hadLine As Boolean, saved As Boolean
    Set ch = wb.Worksheets(sheetName).ChartObjects(chartName).Chart
    saved = wb.Saved
    On Error Resume Next
    If LtxBool("ChartNoBorder", True) Then
        hadLine = ch.ChartArea.Format.Line.Visible
        If hadLine Then ch.ChartArea.Format.Line.Visible = 0
    End If
    ch.CopyPicture xlScreen, xlPicture, xlScreen
    If hadLine Then ch.ChartArea.Format.Line.Visible = -1
    wb.Saved = saved
End Sub

' Copies the chart object itself (for native chart / embedded paste)
Public Sub LtxCopyChartObject(wb As Object, sheetName As String, chartName As String)
    wb.Worksheets(sheetName).ChartObjects(chartName).Copy
End Sub

' ---------- Pending clipboard descriptor written by the Excel add-in ----------
Public Function LtxPendingPath() As String
    LtxPendingPath = LtxDataDir() & "\pending.txt"
End Function

Public Function LtxReadPending() As Object
    Set LtxReadPending = LtxParseKV(LtxReadFile(LtxPendingPath()))
End Function

' ---------- Shortcut / keytip card (HTML opened in the browser) ----------
Public Function LtxCardHtml(app As String, rows As String) As String
    LtxCardHtml = "<!doctype html><html><head><meta charset='utf-8'><title>Lattice " & app & " shortcuts</title><style>" & _
        "body{font-family:Segoe UI,Arial,sans-serif;background:#f5f6f8;margin:0;padding:24px;color:#1f2937}" & _
        ".card{background:#fff;max-width:900px;margin:auto;border-radius:14px;box-shadow:0 8px 30px rgba(31,56,100,.12);padding:28px 36px;border-top:4px solid #2e75b6}" & _
        "h1{margin:0 0 4px;font-size:22px}h1 span{background:linear-gradient(90deg,#1f3864,#2e75b6);color:#fff;padding:4px 12px;border-radius:6px;margin-left:8px;font-weight:600;font-size:16px;vertical-align:middle}" & _
        "p.sub{color:#6b7280;margin:0 0 18px;font-size:13px}table{width:100%;border-collapse:collapse}td{padding:6px 8px;border-bottom:1px solid #eef0f3;font-size:13px}" & _
        "td:last-child{text-align:right;white-space:nowrap}tr.g td{background:#eef2f7;font-weight:700;color:#1f3864;padding-top:10px;border:0;font-size:12px;text-transform:uppercase;letter-spacing:.6px}" & _
        ".k{display:inline-block;min-width:18px;text-align:center;padding:3px 7px;margin-left:4px;border-radius:5px;background:#1f3864;color:#fff;font-size:11px;font-weight:600}.alt{background:#f97316}.or{color:#9ca3af;font-size:11px}" & _
        "</style></head><body><div class='card'><h1>Lattice <span>" & UCase$(app) & " KEYTIPS</span></h1>" & _
        "<p class='sub'>Press <b>Alt</b>, then <b>T</b> to open the Lattice tab, then the letter shown. Buttons without a letter are reachable with the arrow keys.</p>" & _
        "<table>" & rows & "</table></div></body></html>"
End Function

Public Function LtxCardGroup(title As String) As String
    LtxCardGroup = "<tr class='g'><td colspan='2'>" & title & "</td></tr>"
End Function

Public Function LtxCardRow(label As String, keys As String) As String
    Dim p As Variant, h As String
    For Each p In Split(keys, " ")
        If p = "/" Then h = h & " <span class='or'>or</span> " Else h = h & "<span class='k" & IIf(p = "Alt", " alt", "") & "'>" & p & "</span>"
    Next
    LtxCardRow = "<tr><td>" & label & "</td><td>" & h & "</td></tr>"
End Function

Public Sub LtxOpenHtml(path As String, html As String)
    LtxWriteFile path, html
    Shell "cmd /c start """" """ & path & """", vbHide
End Sub

Public Sub LtxLog(msg As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open LtxDataDir() & "\lattice.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & msg
    Close #f
End Sub

' ---------- Excel instance discovery ----------
' Application object of the Excel instance behind an XLMAIN window (Nothing when it has no workbook window).
Private Function ExcelFromHwnd(ByVal hMain As LongPtr) As Object
    Dim hDesk As LongPtr, hBook As LongPtr, obj As Object, iid(0 To 3) As Long
    iid(0) = &H20400: iid(1) = 0: iid(2) = &HC0: iid(3) = &H46000000       ' IID_IDispatch
    hDesk = FindWindowExA(hMain, 0, "XLDESK", vbNullString)
    If hDesk = 0 Then Exit Function
    hBook = FindWindowExA(hDesk, 0, "EXCEL7", vbNullString)
    If hBook = 0 Then Exit Function
    On Error Resume Next
    If AccessibleObjectFromWindow(hBook, &HFFFFFFF0, iid(0), obj) = 0 Then Set ExcelFromHwnd = obj.Application
End Function

' The best running Excel instance: the one holding preferPath (if given) first, then visible
' windows before hidden automation instances. Nothing when no instance has a workbook open.
Public Function LtxPickExcel(Optional preferPath As String = "") As Object
    Dim h As LongPtr, app As Object, best As Object, score As Long, bestScore As Long
    bestScore = -1
    Do
        h = FindWindowExA(0, h, "XLMAIN", vbNullString)
        If h = 0 Then Exit Do
        Set app = ExcelFromHwnd(h)
        If Not app Is Nothing Then
            score = 0
            If IsWindowVisible(h) <> 0 Then score = 2
            If Len(preferPath) > 0 Then
                If Not FindWorkbookIn(app, preferPath) Is Nothing Then score = score + 10
            End If
            If score > bestScore Then bestScore = score: Set best = app
        End If
    Loop
    Set LtxPickExcel = best
End Function
