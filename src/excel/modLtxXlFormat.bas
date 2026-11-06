Attribute VB_Name = "modLtxXlFormat"
' =====================================================================
'  Lattice for Excel - formatting cycles and AutoColor
'  AutoColor convention (configurable): hard-codes blue, formulas black,
'  links to other sheets green, links to other workbooks purple/red.
' =====================================================================
Option Explicit

Public Const LTX_APP_XL As String = "Lattice"

' Palettes and cycles come from %APPDATA%\Lattice\excel_styles.txt when that
' file defines the key (one "key=a|b|c" line each), else from the registry,
' else from the built-in default. Edit the file with Lattice > Settings > Style file.
Private Function Pal(key As String, def As String) As Variant
    Dim v As String
    v = StyleFileValue(key)
    If Len(v) = 0 Then v = GetSetting("Lattice", "Excel", key, def)
    Pal = Split(v, "|")
End Function

Public Function StyleFilePath() As String
    StyleFilePath = Environ$("APPDATA") & "\Lattice\excel_styles.txt"
End Function

Public Function StyleFileValue(key As String) As String
    Static cache As Object, stamp As Date
    Dim f As Integer, ln As String, i As Long
    If Dir(StyleFilePath()) = "" Then Exit Function
    If cache Is Nothing Or FileDateTime(StyleFilePath()) <> stamp Then
        Set cache = CreateObject("Scripting.Dictionary"): cache.CompareMode = 1
        stamp = FileDateTime(StyleFilePath())
        f = FreeFile
        Open StyleFilePath() For Input As #f
        Do While Not EOF(f)
            Line Input #f, ln
            i = InStr(ln, "=")
            If i > 1 And Left$(Trim$(ln), 1) <> "#" Then cache(Trim$(Left$(ln, i - 1))) = Trim$(Mid$(ln, i + 1))
        Loop
        Close #f
    End If
    If cache.Exists(key) Then StyleFileValue = cache(key)
End Function

' Creates the style file with documented defaults (if missing) and opens it
Public Sub LtxEditStyleFile()
    Dim p As String, f As Integer
    p = StyleFilePath()
    If Dir(Environ$("APPDATA") & "\Lattice", vbDirectory) = "" Then MkDir Environ$("APPDATA") & "\Lattice"
    If Dir(p) = "" Then
        f = FreeFile
        Open p For Output As #f
        Print #f, "# Lattice for Excel - style guide. One key per line, options separated by |"
        Print #f, "# Colours are BGR longs (Excel's Font.Color). Blue 16711680, green 32768, red 255, purple 8388736, grey 8421504"
        Print #f, "NumCycle=#,##0_);(#,##0)|#,##0.0_);(#,##0.0)|#,##0.00_);(#,##0.00)|General"
        Print #f, "PctCycle=0%|0.0%|0.00%"
        Print #f, "MultCycle=0.0x|0.00x|0x"
        Print #f, "DateCycle=mmm-yy|dd-mmm-yy|yyyy|mmm yyyy"
        Print #f, "CurCycle=$#,##0_);($#,##0)|" & ChrW(8364) & "#,##0_);(" & ChrW(8364) & "#,##0)|" & ChrW(163) & "#,##0_);(" & ChrW(163) & "#,##0)"
        Print #f, "FontCycle=0|16711680|32768|8388736|255|8421504"
        Print #f, "FillCycle=none|65535|15921906|14277081|13434879"
        Print #f, "# AutoColor"
        Print #f, "ColInput=16711680"
        Print #f, "ColFormula=0"
        Print #f, "ColLink=32768"
        Print #f, "ColExternal=8388736"
        Print #f, "ColPartial=128"
        Print #f, "# Chart palette (hex RRGGBB)"
        Print #f, "ChartPalette=1F3864|2E75B6|9DC3E6|BDD7EE|7F7F7F|BFBFBF|C00000|ED7D31"
        Close #f
    End If
    Shell "notepad.exe """ & p & """", vbNormalFocus
End Sub

Public Sub XlSet(key As String, v As String)
    SaveSetting "Lattice", "Excel", key, v
End Sub

Public Function XlGet(key As String, Optional def As String = "") As String
    XlGet = GetSetting("Lattice", "Excel", key, def)
End Function

' ---------------- AutoColor ----------------
' Bulk version: numeric constants in one call, formulas in one call, then only the
' exceptions (sheet links, external links, embedded hard-codes) cell by cell, read
' from the formula array in memory. Workbook-scale runs take seconds.
Public Sub LtxAutoColor(Optional rng As Range)
    Dim f As String, area As Range, arr As Variant, r As Long, c As Long, fr As Range, nums As Range
    Dim cInput As Long, cFormula As Long, cLink As Long, cExt As Long, cPartial As Long, calc As Long
    If rng Is Nothing Then Set rng = Selection
    cInput = Val(XlGet("ColInput", CStr(RGB(0, 0, 255))))
    cFormula = Val(XlGet("ColFormula", "0"))
    cLink = Val(XlGet("ColLink", CStr(RGB(0, 128, 0))))
    cExt = Val(XlGet("ColExternal", CStr(RGB(128, 0, 128))))
    cPartial = Val(XlGet("ColPartial", CStr(RGB(128, 0, 0))))   ' formula containing hard-coded numbers
    calc = Application.Calculation
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
    On Error Resume Next
    Set nums = rng.SpecialCells(xlCellTypeConstants, xlNumbers)
    Set fr = rng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not nums Is Nothing Then nums.Font.Color = cInput
    If Not fr Is Nothing Then
        fr.Font.Color = cFormula
        For Each area In fr.Areas
            If area.Cells.Count = 1 Then
                ReDim arr(1 To 1, 1 To 1): arr(1, 1) = area.Formula
            Else
                arr = area.Formula
            End If
            For r = 1 To UBound(arr, 1)
                For c = 1 To UBound(arr, 2)
                    f = CStr(arr(r, c))
                    If Len(f) > 1 Then
                        If InStr(f, "[") > 0 Then
                            area.Cells(r, c).Font.Color = cExt
                        ElseIf InStr(f, "!") > 0 Then
                            area.Cells(r, c).Font.Color = cLink
                        ElseIf HasHardcode(f) Then
                            area.Cells(r, c).Font.Color = cPartial
                        End If
                    End If
                Next
            Next
        Next
    End If
    Application.Calculation = calc: Application.EnableEvents = True: Application.ScreenUpdating = True
End Sub

' AutoColor every sheet in the workbook
Public Sub LtxAutoColorWorkbook()
    Dim ws As Worksheet, cur As Worksheet
    Set cur = ActiveSheet
    Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And Not ws.ProtectContents Then
            Application.StatusBar = "Lattice: AutoColor " & ws.Name
            LtxAutoColor ws.UsedRange
        End If
    Next
    cur.Activate
    Application.StatusBar = "Lattice: AutoColor done for the workbook"
    Application.ScreenUpdating = True
End Sub

' True when a formula embeds a literal number other than 0/1/100/12 etc.
Private Function HasHardcode(f As String) As Boolean
    Dim re As Object, m As Object, v As Double
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(?:^|[^A-Za-z0-9_$.!""])(\d+\.?\d*)(?![A-Za-z0-9_:(])"
    For Each m In re.Execute(Replace(f, """", "'"))
        v = Val(m.SubMatches(0))
        If v <> 0 And v <> 1 And v <> 100 And v <> 12 And v <> 365 And v <> 1000 And v <> 4 Then
            HasHardcode = True: Exit Function
        End If
    Next
End Function

Public Sub LtxAutoColorSheet()
    LtxAutoColor ActiveSheet.UsedRange
End Sub

' ---------------- Cycles ----------------
Private Sub CycleProp(getter As String, pal As Variant)
    Dim cur As String, i As Long, nxt As String
    Select Case getter
        Case "numfmt": cur = Selection.Cells(1).NumberFormat
        Case "font":   cur = CStr(Selection.Cells(1).Font.Color)
        Case "fill":   cur = IIf(Selection.Cells(1).Interior.ColorIndex = xlNone, "none", CStr(Selection.Cells(1).Interior.Color))
    End Select
    nxt = pal(0)
    For i = 0 To UBound(pal)
        If pal(i) = cur Then nxt = pal((i + 1) Mod (UBound(pal) + 1)): Exit For
    Next
    Select Case getter
        Case "numfmt": Selection.NumberFormat = nxt
        Case "font":   Selection.Font.Color = CLng(nxt)
        Case "fill":   If nxt = "none" Then Selection.Interior.ColorIndex = xlNone Else Selection.Interior.Color = CLng(nxt)
    End Select
End Sub

Public Sub LtxCycleNumber()
    CycleProp "numfmt", Pal("NumCycle", "#,##0_);(#,##0)|#,##0.0_);(#,##0.0)|#,##0.00_);(#,##0.00)|General")
End Sub
Public Sub LtxCyclePercent()
    CycleProp "numfmt", Pal("PctCycle", "0%|0.0%|0.00%")
End Sub
Public Sub LtxCycleMultiple()
    CycleProp "numfmt", Pal("MultCycle", "0.0x|0.00x|0x")
End Sub
Public Sub LtxCycleDate()
    CycleProp "numfmt", Pal("DateCycle", "mmm-yy|dd-mmm-yy|yyyy|mmm yyyy")
End Sub
Public Sub LtxCycleCurrency()
    CycleProp "numfmt", Pal("CurCycle", "$#,##0_);($#,##0)|???#,##0_);(???#,##0)|??#,##0_);(??#,##0)")
End Sub
Public Sub LtxCycleFontColor()
    CycleProp "font", Pal("FontCycle", "0|16711680|32768|8388736|255|8421504")
End Sub
Public Sub LtxCycleFill()
    CycleProp "fill", Pal("FillCycle", "none|65535|15921906|14277081|13434879")
End Sub

Public Sub LtxCycleBorder()
    Dim r As Range, state As Long
    Set r = Selection
    ' 0 none -> 1 bottom thin -> 2 bottom medium -> 3 top thin + bottom double -> 4 outline
    state = Val(r.Cells(1).ID)
    state = (state + 1) Mod 5
    r.Borders.LineStyle = xlNone
    Select Case state
        Case 1: r.Borders(xlEdgeBottom).LineStyle = xlContinuous: r.Borders(xlEdgeBottom).Weight = xlThin
        Case 2: r.Borders(xlEdgeBottom).LineStyle = xlContinuous: r.Borders(xlEdgeBottom).Weight = xlMedium
        Case 3: r.Borders(xlEdgeTop).LineStyle = xlContinuous: r.Borders(xlEdgeBottom).LineStyle = xlDouble
        Case 4: r.BorderAround xlContinuous, xlThin
    End Select
    r.Cells(1).ID = CStr(state)
End Sub

Public Sub LtxCycleAlign()
    Select Case Selection.HorizontalAlignment
        Case xlLeft: Selection.HorizontalAlignment = xlCenter
        Case xlCenter: Selection.HorizontalAlignment = xlRight
        Case xlRight: Selection.HorizontalAlignment = xlGeneral
        Case Else: Selection.HorizontalAlignment = xlLeft
    End Select
End Sub

Public Sub LtxIndentMore(): Selection.IndentLevel = Selection.IndentLevel + 1: End Sub
Public Sub LtxIndentLess()
    If Selection.IndentLevel > 0 Then Selection.IndentLevel = Selection.IndentLevel - 1
End Sub

Public Sub LtxToggleWrap(): Selection.WrapText = Not Selection.Cells(1).WrapText: End Sub

Public Sub LtxCenterAcross()
    Selection.HorizontalAlignment = xlCenterAcrossSelection
End Sub

' Title / header row styling
Public Sub LtxHeaderStyle()
    With Selection
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .HorizontalAlignment = xlCenter
    End With
End Sub

Public Sub LtxTotalStyle()
    With Selection
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
    End With
End Sub

' Sticky format painter
Public Sub LtxCopyFormat()
    Selection.Copy
    XlSet "FmtAddr", Selection.Address(External:=True)
End Sub

Public Sub LtxPasteFormat()
    Dim src As Range
    On Error Resume Next
    Set src = Range(XlGet("FmtAddr"))
    If src Is Nothing Then MsgBox "Copy a format first.", vbInformation, "Lattice": Exit Sub
    src.Copy
    Selection.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

' Column width / row height helpers
Public Sub LtxColWidth()
    Dim w As String
    w = InputBox("Column width:", "Lattice", CStr(Selection.Columns(1).ColumnWidth))
    If Len(w) > 0 Then Selection.ColumnWidth = Val(w)
End Sub

Public Sub LtxAutofitCols(): Selection.Columns.AutoFit: End Sub

Public Sub LtxGridlinesToggle()
    ActiveWindow.DisplayGridlines = Not ActiveWindow.DisplayGridlines
End Sub

Public Sub LtxZoomCycle()
    Select Case ActiveWindow.Zoom
        Case 100: ActiveWindow.Zoom = 85
        Case 85: ActiveWindow.Zoom = 70
        Case Else: ActiveWindow.Zoom = 100
    End Select
End Sub
