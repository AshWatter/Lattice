Attribute VB_Name = "modLtxXlHardcodes"
' =====================================================================
'  Lattice for Excel - hard-code and hidden-error audit
'  Finds numbers typed inside formulas (which should be inputs), formulas
'  whose errors are masked by IFERROR, broken formulas, and writes a
'  clickable audit sheet. Highlighting is optional and reversible.
' =====================================================================
Option Explicit

Private Const HDR_ROW As Long = 4     ' report layout: banner rows 1-3, header row 4, data from row 5

Private Function Whitelist() As Variant
    Whitelist = Split(XlGet("HardcodeWhitelist", "0|1|2|100|12|365|360|1000|1000000|4|52|30|31|0.5"), "|")
End Function

' Returns the numeric literals embedded in a formula (comma separated), "" if none
Public Function FormulaLiterals(f As String) As String
    Dim re As Object, m As Object, v As Double, w As Variant, ok As Boolean, out As String, clean As String
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    ' strip strings and sheet names in quotes
    re.Pattern = """[^""]*""|'[^']*'"
    clean = re.Replace(f, "x")
    ' a number not glued to a cell ref / function / name, and not an exponent
    re.Pattern = "(?:^|[^A-Za-z0-9_$.!:])(\d+\.?\d*|\.\d+)(?![A-Za-z0-9_:(!])"
    For Each m In re.Execute(clean)
        v = Val(m.SubMatches(0))
        ok = True
        For Each w In Whitelist()
            If v = Val(w) Then ok = False: Exit For
        Next
        If ok Then out = out & IIf(Len(out) > 0, ", ", "") & m.SubMatches(0)
    Next
    FormulaLiterals = out
End Function

' Collect issues as a 2-D array (sheet, cell, issue, detail) for the Model Check pane.
' Same checks as the audit sheet, without writing anything to the workbook.
Public Function LtxCollectIssues(Optional wholeWorkbook As Boolean = False) As Variant
    Dim ws As Worksheet, rng As Range, c As Range, lits As String, inner As String, masked As Variant
    Dim rows As Collection, target As Worksheet, out() As String, i As Long, it As Variant
    Set rows = New Collection
    Set target = ActiveSheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "Lattice Audit" And ws.Name <> "Lattice Health" Then
            If wholeWorkbook Or ws Is target Then
                On Error Resume Next
                Set rng = Nothing
                Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
                On Error GoTo 0
                If Not rng Is Nothing Then
                    For Each c In rng.Cells
                        lits = FormulaLiterals(c.Formula)
                        If Len(lits) > 0 Then rows.Add Array(ws.Name, c.Address(False, False), "Hard-code in formula", lits)
                        If IsError(c.Value) Then rows.Add Array(ws.Name, c.Address(False, False), "Error", CStr(c.Text))
                        If Left$(UCase$(c.Formula), 9) = "=IFERROR(" Or Left$(UCase$(c.Formula), 6) = "=IFNA(" Then
                            inner = FirstArgument(c.Formula)
                            If Len(inner) > 0 Then
                                On Error Resume Next
                                masked = Empty: masked = ws.Evaluate(inner)
                                On Error GoTo 0
                                If IsError(masked) Then rows.Add Array(ws.Name, c.Address(False, False), "Error masked by IFERROR", "inner result: " & CStr(masked))
                            End If
                        End If
                        If VarType(c.Value) = vbString Then
                            If IsNumeric(c.Value) And Len(c.Value) > 0 Then rows.Add Array(ws.Name, c.Address(False, False), "Number stored as text", CStr(c.Value))
                        End If
                        If c.Column > 1 Then
                            If c.Offset(0, -1).HasFormula And c.Offset(0, 1).HasFormula Then
                                If c.FormulaR1C1 <> c.Offset(0, -1).FormulaR1C1 And c.Offset(0, -1).FormulaR1C1 = c.Offset(0, 1).FormulaR1C1 Then
                                    rows.Add Array(ws.Name, c.Address(False, False), "Inconsistent formula", "neighbours: " & c.Offset(0, -1).Formula)
                                End If
                            End If
                        End If
                    Next
                End If
                ' unused numeric inputs: constants no formula depends on
                On Error Resume Next
                Set rng = Nothing
                Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers)
                On Error GoTo 0
                If Not rng Is Nothing Then
                    If rng.Cells.Count <= 2000 Then
                        For Each c In rng.Cells
                            If Not HasDependents(c) Then rows.Add Array(ws.Name, c.Address(False, False), "Unused numeric input", CStr(c.Text))
                        Next
                    End If
                End If
            End If
        End If
    Next
    If rows.Count = 0 Then LtxCollectIssues = Empty: Exit Function
    ReDim out(0 To rows.Count - 1, 0 To 3)
    i = 0
    For Each it In rows
        out(i, 0) = it(0): out(i, 1) = it(1): out(i, 2) = it(2): out(i, 3) = it(3)
        i = i + 1
    Next
    LtxCollectIssues = out
End Function

Private Function HasDependents(c As Range) As Boolean
    Dim d As Range
    On Error Resume Next
    Set d = c.DirectDependents
    HasDependents = Not d Is Nothing
End Function

Public Sub LtxModelCheckPane()
    frmLtxModelCheck.Show vbModeless
End Sub

' Scan selection (or whole sheet) and build the audit sheet
Public Sub LtxAuditHardcodes(Optional wholeWorkbook As Boolean = False)
    Dim ws As Worksheet, rng As Range, c As Range, rpt As Worksheet, r As Long, lits As String, scanned As Long
    Dim wb As Workbook, inner As String, masked As Variant, kind As String, target As Worksheet, selRange As Range
    Set wb = ActiveWorkbook
    Set target = ActiveSheet                      ' remember before the report sheet is created (Add activates it)
    If TypeName(Selection) = "Range" Then Set selRange = Selection
    Application.ScreenUpdating = False
    Set rpt = AuditSheet(wb)
    rpt.Range("A4:E4").Value = Array("Sheet", "Cell", "Issue", "Detail", "Formula")
    r = HDR_ROW + 1
    For Each ws In wb.Worksheets
        If ws.Name <> rpt.Name And ws.Name <> "Lattice Health" Then
            If wholeWorkbook Or ws Is target Then
                On Error Resume Next
                Set rng = Nothing
                If wholeWorkbook Or selRange Is Nothing Then
                    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
                ElseIf selRange.Cells.Count = 1 Then
                    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
                Else
                    Set rng = selRange.SpecialCells(xlCellTypeFormulas)
                End If
                On Error GoTo 0
                If Not rng Is Nothing Then
                    For Each c In rng.Cells
                        scanned = scanned + 1
                        ' 1) hard-codes inside formulas
                        lits = FormulaLiterals(c.Formula)
                        If Len(lits) > 0 Then AddRow rpt, r, ws, c, "Hard-code in formula", lits
                        ' 2) formulas that error
                        If IsError(c.Value) Then AddRow rpt, r, ws, c, "Error", CStr(c.Text)
                        ' 3) hidden errors behind IFERROR / IFNA
                        If Left$(UCase$(c.Formula), 9) = "=IFERROR(" Or Left$(UCase$(c.Formula), 6) = "=IFNA(" Then
                            inner = FirstArgument(c.Formula)
                            If Len(inner) > 0 Then
                                On Error Resume Next
                                masked = Empty
                                masked = ws.Evaluate(inner)
                                On Error GoTo 0
                                If IsError(masked) Then AddRow rpt, r, ws, c, "Error masked by IFERROR", "inner result: " & CStr(masked)
                            End If
                        End If
                        ' 4) formula returning text that looks like a number
                        If VarType(c.Value) = vbString Then
                            If IsNumeric(c.Value) And Len(c.Value) > 0 Then AddRow rpt, r, ws, c, "Number stored as text", CStr(c.Value)
                        End If
                        ' 5) inconsistent with left neighbour
                        If c.Column > 1 Then
                            If c.Offset(0, -1).HasFormula And c.Offset(0, 1).HasFormula Then
                                If c.FormulaR1C1 <> c.Offset(0, -1).FormulaR1C1 And c.Offset(0, -1).FormulaR1C1 = c.Offset(0, 1).FormulaR1C1 Then
                                    AddRow rpt, r, ws, c, "Inconsistent formula", "neighbours: " & c.Offset(0, -1).Formula
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End If
    Next
    LtxStyleReport rpt, "Model Check", IIf(wholeWorkbook, "Workbook: " & wb.Name, "Sheet: " & target.Name) & "   |   " & scanned & " formulas scanned   |   " & (r - HDR_ROW - 1) & " issues", HDR_ROW, r - 1, 5
    rpt.Activate
    rpt.Range("A" & (HDR_ROW + 1)).Select
    Application.ScreenUpdating = True
    If r - HDR_ROW - 1 = 0 Then MsgBox "No issues found in " & scanned & " formula(s).", vbInformation, LTX_APP_XL & " - Model Check"
End Sub


Private Sub AddRow(rpt As Worksheet, ByRef r As Long, ws As Worksheet, c As Range, issue As String, detail As String)
    rpt.Cells(r, 1).Value = ws.Name
    rpt.Hyperlinks.Add rpt.Cells(r, 2), "", "'" & ws.Name & "'!" & c.Address(False, False), "Jump to " & ws.Name & "!" & c.Address(False, False), c.Address(False, False)
    rpt.Cells(r, 3).Value = issue
    rpt.Cells(r, 4).Value = detail
    rpt.Cells(r, 5).Value = "'" & c.Formula
    r = r + 1
End Sub

' ---------- report styling shared by Model Check and Health report ----------
Public Sub LtxStyleReport(rpt As Worksheet, title As String, subtitle As String, headerRow As Long, lastRow As Long, lastCol As Long)
    Dim navy As Long, i As Long, c As Long, issue As String, col As Long, hdr As Range, body As Range, counts As Object, k As Variant, x As Long
    navy = RGB(31, 56, 100)
    Application.ScreenUpdating = False
    With rpt
        .Cells.Font.Name = "Segoe UI": .Cells.Font.Size = 9
        ' banner
        .Range(.Cells(1, 1), .Cells(2, lastCol)).Interior.Color = navy
        .Cells(1, 1).Value = "LATTICE  |  " & UCase$(title)
        .Cells(1, 1).Font.Color = RGB(255, 255, 255): .Cells(1, 1).Font.Bold = True: .Cells(1, 1).Font.Size = 14
        .Cells(2, 1).Value = subtitle & "   |   " & Format$(Now, "dd mmm yyyy hh:nn")
        .Cells(2, 1).Font.Color = RGB(189, 215, 238): .Cells(2, 1).Font.Size = 9
        .Rows(1).RowHeight = 26: .Rows(2).RowHeight = 16
        .Range(.Cells(3, 1), .Cells(3, lastCol)).Interior.Color = RGB(46, 117, 182): .Rows(3).RowHeight = 3   ' azure accent under the banner
        ' header
        Set hdr = .Range(.Cells(headerRow, 1), .Cells(headerRow, lastCol))
        hdr.Font.Bold = True: hdr.Font.Color = navy: hdr.Interior.Color = RGB(222, 229, 240)
        hdr.Borders(xlEdgeBottom).LineStyle = xlContinuous: hdr.Borders(xlEdgeBottom).Color = navy
        .Rows(headerRow).RowHeight = 20: hdr.VerticalAlignment = xlCenter
        If lastRow > headerRow Then
            Set body = .Range(.Cells(headerRow + 1, 1), .Cells(lastRow, lastCol))
            body.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            body.Borders(xlInsideHorizontal).Color = RGB(235, 238, 243)
            body.VerticalAlignment = xlCenter
            For i = headerRow + 1 To lastRow
                .Rows(i).RowHeight = 17
                If (i - headerRow) Mod 2 = 0 Then .Range(.Cells(i, 1), .Cells(i, lastCol)).Interior.Color = RGB(247, 248, 250)
            Next
            ' issue chips (column 3 when present)
            col = 0
            For c = 1 To lastCol
                If LCase$(.Cells(headerRow, c).Value) = "issue" Then col = c
            Next
            If col > 0 Then
                Set counts = CreateObject("Scripting.Dictionary")
                For i = headerRow + 1 To lastRow
                    issue = .Cells(i, col).Value
                    counts(issue) = counts(issue) + 1
                    With .Cells(i, col)
                        .Font.Bold = True
                        .IndentLevel = 1
                        Select Case issue
                            Case "Error":                   .Font.Color = RGB(192, 0, 0): .Interior.Color = RGB(253, 230, 230)
                            Case "Error masked by IFERROR": .Font.Color = RGB(197, 90, 17): .Interior.Color = RGB(255, 236, 217)
                            Case "Hard-code in formula":    .Font.Color = RGB(31, 78, 184): .Interior.Color = RGB(226, 236, 253)
                            Case "Inconsistent formula":    .Font.Color = RGB(158, 118, 0): .Interior.Color = RGB(255, 246, 204)
                            Case "Number stored as text":   .Font.Color = RGB(112, 48, 160): .Interior.Color = RGB(239, 229, 247)
                            Case "Unused numeric input":    .Font.Color = RGB(89, 89, 89): .Interior.Color = RGB(237, 237, 237)
                        End Select
                    End With
                Next
                ' counts in the banner, right side
                x = lastCol
                For Each k In counts.Keys
                    .Cells(2, x).Value = k & ": " & counts(k)
                    .Cells(2, x).Font.Color = RGB(255, 255, 255): .Cells(2, x).Font.Size = 9: .Cells(2, x).HorizontalAlignment = xlRight
                    x = x - 1
                    If x <= 1 Then Exit For
                Next
            End If
            ' formula column in monospace, grey
            If LCase$(.Cells(headerRow, lastCol).Value) = "formula" Then
                .Range(.Cells(headerRow + 1, lastCol), .Cells(lastRow, lastCol)).Font.Name = "Consolas"
                .Range(.Cells(headerRow + 1, lastCol), .Cells(lastRow, lastCol)).Font.Color = RGB(89, 89, 89)
            End If
            hdr.AutoFilter
        End If
        .Range(.Cells(headerRow, 1), .Cells(IIf(lastRow > headerRow, lastRow, headerRow), lastCol)).Columns.AutoFit
        For c = 1 To lastCol
            If .Columns(c).ColumnWidth < 10 Then .Columns(c).ColumnWidth = 10
            If .Columns(c).ColumnWidth > 70 Then .Columns(c).ColumnWidth = 70
        Next
        .Columns(1).ColumnWidth = IIf(.Columns(1).ColumnWidth < 14, 14, .Columns(1).ColumnWidth)
    End With
    rpt.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
    rpt.Range("A" & (headerRow + 1)).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    rpt.Tab.Color = navy
    Application.ScreenUpdating = True
    ' report sheets read like a card: no gridlines, header row frozen
    Dim keepSheet As Object
    On Error Resume Next
    Set keepSheet = ActiveSheet
    rpt.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    rpt.Cells(headerRow + 1, 1).Select
    ActiveWindow.FreezePanes = True
    rpt.Cells(headerRow + 1, 1).Select
    keepSheet.Activate
End Sub

' First argument of =FUNC(arg1, arg2) honouring nested parentheses and strings
Private Function FirstArgument(f As String) As String
    Dim i As Long, depth As Long, ch As String, inStr_ As Boolean, start As Long
    start = InStr(f, "(") + 1
    For i = start To Len(f)
        ch = Mid$(f, i, 1)
        If ch = """" Then inStr_ = Not inStr_
        If Not inStr_ Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then
                If depth = 0 Then FirstArgument = "=" & Mid$(f, start, i - start): Exit Function
                depth = depth - 1
            End If
            If (ch = "," Or ch = ";") And depth = 0 Then FirstArgument = "=" & Mid$(f, start, i - start): Exit Function
        End If
    Next
End Function

Private Function AuditSheet(wb As Workbook) As Worksheet
    Dim rpt As Worksheet
    On Error Resume Next
    Set rpt = wb.Worksheets("Lattice Audit")
    On Error GoTo 0
    If rpt Is Nothing Then
        Set rpt = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        rpt.Name = "Lattice Audit"
    Else
        rpt.Cells.Clear
        rpt.Hyperlinks.Delete
        On Error Resume Next
        rpt.AutoFilterMode = False
        On Error GoTo 0
    End If
    Set AuditSheet = rpt
End Function

' Highlight (on the sheet itself) formulas that embed hard-codes; run again to clear
Public Sub LtxHighlightHardcodes()
    Dim c As Range, rng As Range, n As Long, on_ As Boolean
    On Error Resume Next
    Set rng = ActiveSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    on_ = Not NameExists("ltx_hl_" & CleanName(ActiveSheet.Name))   ' a hidden name remembers the toggle
    Application.ScreenUpdating = False
    For Each c In rng.Cells
        If Len(FormulaLiterals(c.Formula)) > 0 Then
            If on_ Then
                c.ID = CStr(c.Interior.Color) & "|" & c.Interior.ColorIndex
                c.Interior.Color = RGB(255, 230, 153): n = n + 1
            Else
                If InStr(c.ID, "|") > 0 Then
                    If Val(Split(c.ID, "|")(1)) = -4142 Then c.Interior.ColorIndex = xlNone Else c.Interior.Color = CLng(Split(c.ID, "|")(0))
                    c.ID = ""
                End If
            End If
        End If
    Next
    If on_ Then
        ActiveWorkbook.Names.Add "ltx_hl_" & CleanName(ActiveSheet.Name), "=1", False
        MsgBox n & " formula(s) with embedded hard-codes highlighted. Run again to clear.", vbInformation, LTX_APP_XL
    Else
        ActiveWorkbook.Names("ltx_hl_" & CleanName(ActiveSheet.Name)).Delete
    End If
    Application.ScreenUpdating = True
End Sub

Private Function NameExists(n As String) As Boolean
    Dim x As Name
    On Error Resume Next
    Set x = ActiveWorkbook.Names(n)
    NameExists = Not x Is Nothing
End Function

Private Function CleanName(s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then CleanName = CleanName & ch Else CleanName = CleanName & "_"
    Next
End Function

' Pull a hard-code out of the active formula into an input cell next to it
Public Sub LtxExtractHardcode()
    Dim c As Range, lits As String, v As String, tgt As Range, f As String, i As Long
    Set c = ActiveCell
    If Not c.HasFormula Then MsgBox "Active cell has no formula.", vbInformation, LTX_APP_XL: Exit Sub
    lits = FormulaLiterals(c.Formula)
    If Len(lits) = 0 Then MsgBox "No embedded hard-codes in this formula.", vbInformation, LTX_APP_XL: Exit Sub
    v = InputBox("Hard-codes found: " & lits & vbLf & vbLf & "Which one to move to an input cell?", LTX_APP_XL, Split(lits, ", ")(0))
    If Len(v) = 0 Then Exit Sub
    On Error Resume Next
    Set tgt = Application.InputBox("Input cell to hold " & v & ":", LTX_APP_XL, Type:=8)
    On Error GoTo 0
    If tgt Is Nothing Then Exit Sub
    tgt.Value = Val(v): tgt.Font.Color = 16711680
    f = c.Formula
    i = InStr(f, v)
    If i > 0 Then
        f = Left$(f, i - 1) & tgt.Address(True, True) & Mid$(f, i + Len(v))
        c.Formula = f
    End If
End Sub
