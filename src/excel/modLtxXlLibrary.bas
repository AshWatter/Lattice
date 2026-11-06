Attribute VB_Name = "modLtxXlLibrary"
' =====================================================================
'  Lattice for Excel - model library
'  Two sources of pre-built analyses:
'   1. Built-in generators (DCF, Black-Scholes, returns table, debt
'      schedule, sensitivity grid) written as live formulas at the active
'      cell, formatted with the Lattice colour convention.
'   2. Your own library workbook (%APPDATA%\Lattice\xl_library.xlsx, path
'      configurable): any range you save is stored as a sheet and can be
'      pasted back with formulas, formats and column widths.
' =====================================================================
Option Explicit

Private Const BLUE As Long = 16711680     ' inputs (BGR)
Private Const BLACK As Long = 0
Private Const GREEN As Long = 32768

Private Function LibPath() As String
    LibPath = XlGet("LibraryPath", Environ$("APPDATA") & "\Lattice\xl_library.xlsx")
End Function

' ---------- helpers ----------
Private Sub Lbl(c As Range, txt As String, Optional bold As Boolean = False)
    c.Value = txt: c.Font.Bold = bold: c.Font.Color = BLACK
End Sub
Private Sub Inp(c As Range, v As Variant, Optional fmt As String = "")
    c.Value = v: c.Font.Color = BLUE
    If Len(fmt) > 0 Then c.NumberFormat = fmt
End Sub
Private Sub Fml(c As Range, f As String, Optional fmt As String = "")
    c.Formula = f: c.Font.Color = BLACK
    If Len(fmt) > 0 Then c.NumberFormat = fmt
End Sub
Private Sub Header(r As Range)
    r.Font.Bold = True
    r.Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub
Private Function Adr(r As Range) As String: Adr = r.Address(False, False): End Function

' ---------- DCF ----------
Public Sub LtxInsertDCF()
    Dim o As Range, i As Long, yrs As Long
    Set o = ActiveCell
    yrs = 5
    Application.ScreenUpdating = False
    Lbl o, "Discounted cash flow", True
    Lbl o.Offset(1, 0), "EUR m"
    For i = 1 To yrs: Lbl o.Offset(1, i), "Year " & i, True: Next
    Header o.Offset(1, 0).Resize(1, yrs + 1)
    Lbl o.Offset(2, 0), "Revenue":            Inp o.Offset(2, 1), 100, "#,##0.0"
    For i = 2 To yrs: Fml o.Offset(2, i), "=" & Adr(o.Offset(2, i - 1)) & "*(1+" & Adr(o.Offset(10, 1)) & ")", "#,##0.0": Next
    Lbl o.Offset(3, 0), "EBITDA"
    For i = 1 To yrs: Fml o.Offset(3, i), "=" & Adr(o.Offset(2, i)) & "*" & Adr(o.Offset(11, 1)), "#,##0.0": Next
    Lbl o.Offset(4, 0), "D&A"
    For i = 1 To yrs: Fml o.Offset(4, i), "=-" & Adr(o.Offset(2, i)) & "*" & Adr(o.Offset(12, 1)), "#,##0.0;(#,##0.0)": Next
    Lbl o.Offset(5, 0), "Tax on EBIT"
    For i = 1 To yrs: Fml o.Offset(5, i), "=-(" & Adr(o.Offset(3, i)) & "+" & Adr(o.Offset(4, i)) & ")*" & Adr(o.Offset(13, 1)), "#,##0.0;(#,##0.0)": Next
    Lbl o.Offset(6, 0), "Capex"
    For i = 1 To yrs: Fml o.Offset(6, i), "=-" & Adr(o.Offset(2, i)) & "*" & Adr(o.Offset(14, 1)), "#,##0.0;(#,##0.0)": Next
    Lbl o.Offset(7, 0), "Change in NWC"
    Fml o.Offset(7, 1), "=-" & Adr(o.Offset(2, 1)) & "*" & Adr(o.Offset(15, 1)), "#,##0.0;(#,##0.0)"
    For i = 2 To yrs: Fml o.Offset(7, i), "=-(" & Adr(o.Offset(2, i)) & "-" & Adr(o.Offset(2, i - 1)) & ")*" & Adr(o.Offset(15, 1)), "#,##0.0;(#,##0.0)": Next
    Lbl o.Offset(8, 0), "Unlevered FCF", True
    For i = 1 To yrs: Fml o.Offset(8, i), "=" & Adr(o.Offset(3, i)) & "+" & Adr(o.Offset(5, i)) & "+" & Adr(o.Offset(6, i)) & "+" & Adr(o.Offset(7, i)), "#,##0.0;(#,##0.0)": Next
    Header o.Offset(8, 0).Resize(1, yrs + 1)
    Lbl o.Offset(10, 0), "Revenue growth":    Inp o.Offset(10, 1), 0.05, "0.0%"
    Lbl o.Offset(11, 0), "EBITDA margin":     Inp o.Offset(11, 1), 0.25, "0.0%"
    Lbl o.Offset(12, 0), "D&A % revenue":     Inp o.Offset(12, 1), 0.04, "0.0%"
    Lbl o.Offset(13, 0), "Tax rate":          Inp o.Offset(13, 1), 0.25, "0.0%"
    Lbl o.Offset(14, 0), "Capex % revenue":   Inp o.Offset(14, 1), 0.05, "0.0%"
    Lbl o.Offset(15, 0), "NWC % revenue":     Inp o.Offset(15, 1), 0.1, "0.0%"
    Lbl o.Offset(16, 0), "WACC":              Inp o.Offset(16, 1), 0.09, "0.0%"
    Lbl o.Offset(17, 0), "Terminal growth":   Inp o.Offset(17, 1), 0.02, "0.0%"
    Lbl o.Offset(18, 0), "Net debt":          Inp o.Offset(18, 1), 50, "#,##0.0"
    Lbl o.Offset(19, 0), "Shares (m)":        Inp o.Offset(19, 1), 10, "#,##0.0"
    Lbl o.Offset(21, 0), "Discount factor"
    For i = 1 To yrs: Fml o.Offset(21, i), "=1/(1+" & Adr(o.Offset(16, 1)) & ")^" & i, "0.000": Next
    Lbl o.Offset(22, 0), "PV of FCF"
    For i = 1 To yrs: Fml o.Offset(22, i), "=" & Adr(o.Offset(8, i)) & "*" & Adr(o.Offset(21, i)), "#,##0.0;(#,##0.0)": Next
    Lbl o.Offset(23, 0), "Sum of PV":         Fml o.Offset(23, 1), "=SUM(" & o.Offset(22, 1).Resize(1, yrs).Address(False, False) & ")", "#,##0.0"
    Lbl o.Offset(24, 0), "Terminal value":    Fml o.Offset(24, 1), "=" & Adr(o.Offset(8, yrs)) & "*(1+" & Adr(o.Offset(17, 1)) & ")/(" & Adr(o.Offset(16, 1)) & "-" & Adr(o.Offset(17, 1)) & ")", "#,##0.0"
    Lbl o.Offset(25, 0), "PV of TV":          Fml o.Offset(25, 1), "=" & Adr(o.Offset(24, 1)) & "*" & Adr(o.Offset(21, yrs)), "#,##0.0"
    Lbl o.Offset(26, 0), "Enterprise value", True: Fml o.Offset(26, 1), "=" & Adr(o.Offset(23, 1)) & "+" & Adr(o.Offset(25, 1)), "#,##0.0"
    Lbl o.Offset(27, 0), "Equity value", True:     Fml o.Offset(27, 1), "=" & Adr(o.Offset(26, 1)) & "-" & Adr(o.Offset(18, 1)), "#,##0.0"
    Lbl o.Offset(28, 0), "Value per share", True:  Fml o.Offset(28, 1), "=" & Adr(o.Offset(27, 1)) & "/" & Adr(o.Offset(19, 1)), "#,##0.00"
    Lbl o.Offset(29, 0), "TV % of EV":             Fml o.Offset(29, 1), "=" & Adr(o.Offset(25, 1)) & "/" & Adr(o.Offset(26, 1)), "0.0%"
    Header o.Offset(26, 0).Resize(1, 2)
    o.Resize(30, yrs + 1).Columns.AutoFit
    Application.ScreenUpdating = True
End Sub

' ---------- Black-Scholes ----------
Public Sub LtxInsertBlackScholes()
    Dim o As Range
    Set o = ActiveCell
    Lbl o, "Black-Scholes option pricing", True
    Lbl o.Offset(1, 0), "Spot (S)":            Inp o.Offset(1, 1), 100, "#,##0.00"
    Lbl o.Offset(2, 0), "Strike (K)":          Inp o.Offset(2, 1), 100, "#,##0.00"
    Lbl o.Offset(3, 0), "Risk-free rate (r)":  Inp o.Offset(3, 1), 0.03, "0.00%"
    Lbl o.Offset(4, 0), "Dividend yield (q)":  Inp o.Offset(4, 1), 0, "0.00%"
    Lbl o.Offset(5, 0), "Volatility (sigma)":  Inp o.Offset(5, 1), 0.25, "0.00%"
    Lbl o.Offset(6, 0), "Time to expiry (yrs)": Inp o.Offset(6, 1), 1, "0.00"
    Dim S As String, K As String, r As String, q As String, v As String, T As String
    S = Adr(o.Offset(1, 1)): K = Adr(o.Offset(2, 1)): r = Adr(o.Offset(3, 1)): q = Adr(o.Offset(4, 1)): v = Adr(o.Offset(5, 1)): T = Adr(o.Offset(6, 1))
    Lbl o.Offset(8, 0), "d1":   Fml o.Offset(8, 1), "=(LN(" & S & "/" & K & ")+(" & r & "-" & q & "+" & v & "^2/2)*" & T & ")/(" & v & "*SQRT(" & T & "))", "0.0000"
    Lbl o.Offset(9, 0), "d2":   Fml o.Offset(9, 1), "=" & Adr(o.Offset(8, 1)) & "-" & v & "*SQRT(" & T & ")", "0.0000"
    Dim d1 As String, d2 As String
    d1 = Adr(o.Offset(8, 1)): d2 = Adr(o.Offset(9, 1))
    Lbl o.Offset(10, 0), "N(d1)": Fml o.Offset(10, 1), "=NORM.S.DIST(" & d1 & ",TRUE)", "0.0000"
    Lbl o.Offset(11, 0), "N(d2)": Fml o.Offset(11, 1), "=NORM.S.DIST(" & d2 & ",TRUE)", "0.0000"
    Lbl o.Offset(13, 0), "Call value", True: Fml o.Offset(13, 1), "=" & S & "*EXP(-" & q & "*" & T & ")*" & Adr(o.Offset(10, 1)) & "-" & K & "*EXP(-" & r & "*" & T & ")*" & Adr(o.Offset(11, 1)), "#,##0.00"
    Lbl o.Offset(14, 0), "Put value", True:  Fml o.Offset(14, 1), "=" & K & "*EXP(-" & r & "*" & T & ")*NORM.S.DIST(-" & d2 & ",TRUE)-" & S & "*EXP(-" & q & "*" & T & ")*NORM.S.DIST(-" & d1 & ",TRUE)", "#,##0.00"
    Lbl o.Offset(16, 0), "Greeks", True
    Lbl o.Offset(17, 0), "Delta (call)": Fml o.Offset(17, 1), "=EXP(-" & q & "*" & T & ")*" & Adr(o.Offset(10, 1)), "0.0000"
    Lbl o.Offset(18, 0), "Gamma":        Fml o.Offset(18, 1), "=EXP(-" & q & "*" & T & ")*NORM.S.DIST(" & d1 & ",FALSE)/(" & S & "*" & v & "*SQRT(" & T & "))", "0.0000"
    Lbl o.Offset(19, 0), "Vega (per 1%)": Fml o.Offset(19, 1), "=" & S & "*EXP(-" & q & "*" & T & ")*NORM.S.DIST(" & d1 & ",FALSE)*SQRT(" & T & ")/100", "0.0000"
    Lbl o.Offset(20, 0), "Theta (call, per day)": Fml o.Offset(20, 1), "=(-" & S & "*EXP(-" & q & "*" & T & ")*NORM.S.DIST(" & d1 & ",FALSE)*" & v & "/(2*SQRT(" & T & "))-" & r & "*" & K & "*EXP(-" & r & "*" & T & ")*" & Adr(o.Offset(11, 1)) & "+" & q & "*" & S & "*EXP(-" & q & "*" & T & ")*" & Adr(o.Offset(10, 1)) & ")/365", "0.0000"
    Lbl o.Offset(21, 0), "Rho (call, per 1%)": Fml o.Offset(21, 1), "=" & K & "*" & T & "*EXP(-" & r & "*" & T & ")*" & Adr(o.Offset(11, 1)) & "/100", "0.0000"
    Header o.Offset(13, 0).Resize(1, 2)
    o.Resize(22, 2).Columns.AutoFit
End Sub

' ---------- Returns (IRR / MOIC) ----------
Public Sub LtxInsertReturns()
    Dim o As Range, i As Long
    Set o = ActiveCell
    Lbl o, "Returns analysis", True
    For i = 0 To 5: Lbl o.Offset(1, i + 1), "Year " & i, True: Next
    Header o.Offset(1, 0).Resize(1, 7)
    Lbl o.Offset(2, 0), "Equity invested":  Inp o.Offset(2, 1), -100, "#,##0.0;(#,##0.0)"
    For i = 2 To 5: Inp o.Offset(2, i), 0, "#,##0.0;(#,##0.0)": Next
    Lbl o.Offset(3, 0), "Dividends"
    For i = 1 To 5: Inp o.Offset(3, i + 1), 5, "#,##0.0;(#,##0.0)": Next
    Lbl o.Offset(4, 0), "Exit proceeds"
    For i = 1 To 4: Inp o.Offset(4, i + 1), 0, "#,##0.0;(#,##0.0)": Next
    Inp o.Offset(4, 6), 220, "#,##0.0;(#,##0.0)"
    Lbl o.Offset(5, 0), "Net cash flow", True
    For i = 0 To 5: Fml o.Offset(5, i + 1), "=SUM(" & o.Offset(2, i + 1).Resize(3, 1).Address(False, False) & ")", "#,##0.0;(#,##0.0)": Next
    Header o.Offset(5, 0).Resize(1, 7)
    Lbl o.Offset(7, 0), "IRR", True:  Fml o.Offset(7, 1), "=IRR(" & o.Offset(5, 1).Resize(1, 6).Address(False, False) & ")", "0.0%"
    Lbl o.Offset(8, 0), "MOIC", True: Fml o.Offset(8, 1), "=SUMIF(" & o.Offset(5, 1).Resize(1, 6).Address(False, False) & ","">0"")/-SUMIF(" & o.Offset(5, 1).Resize(1, 6).Address(False, False) & ",""<0"")", "0.00x"
    o.Resize(9, 7).Columns.AutoFit
End Sub

' ---------- Debt schedule ----------
Public Sub LtxInsertDebtSchedule()
    Dim o As Range, i As Long
    Set o = ActiveCell
    Lbl o, "Debt schedule", True
    For i = 1 To 5: Lbl o.Offset(1, i), "Year " & i, True: Next
    Header o.Offset(1, 0).Resize(1, 6)
    Lbl o.Offset(2, 0), "Opening balance"
    Lbl o.Offset(3, 0), "Mandatory amortisation"
    Lbl o.Offset(4, 0), "Cash sweep"
    Lbl o.Offset(5, 0), "Closing balance", True
    Lbl o.Offset(6, 0), "Interest expense"
    Lbl o.Offset(8, 0), "Initial debt":      Inp o.Offset(8, 1), 300, "#,##0.0"
    Lbl o.Offset(9, 0), "Amortisation % p.a.": Inp o.Offset(9, 1), 0.05, "0.0%"
    Lbl o.Offset(10, 0), "Interest rate":    Inp o.Offset(10, 1), 0.06, "0.00%"
    Lbl o.Offset(11, 0), "Cash available for sweep"
    For i = 1 To 5: Inp o.Offset(11, i), 20, "#,##0.0": Next
    Lbl o.Offset(12, 0), "Sweep %":          Inp o.Offset(12, 1), 0.5, "0%"
    Fml o.Offset(2, 1), "=" & Adr(o.Offset(8, 1)), "#,##0.0"
    For i = 2 To 5: Fml o.Offset(2, i), "=" & Adr(o.Offset(5, i - 1)), "#,##0.0": Next
    For i = 1 To 5
        Fml o.Offset(3, i), "=-MIN(" & Adr(o.Offset(2, i)) & "," & Adr(o.Offset(8, 1)) & "*" & Adr(o.Offset(9, 1)) & ")", "#,##0.0;(#,##0.0)"
        Fml o.Offset(4, i), "=-MIN(" & Adr(o.Offset(2, i)) & "+" & Adr(o.Offset(3, i)) & "," & Adr(o.Offset(11, i)) & "*" & Adr(o.Offset(12, 1)) & ")", "#,##0.0;(#,##0.0)"
        Fml o.Offset(5, i), "=" & Adr(o.Offset(2, i)) & "+" & Adr(o.Offset(3, i)) & "+" & Adr(o.Offset(4, i)), "#,##0.0"
        Fml o.Offset(6, i), "=AVERAGE(" & Adr(o.Offset(2, i)) & "," & Adr(o.Offset(5, i)) & ")*" & Adr(o.Offset(10, 1)), "#,##0.0"
    Next
    Header o.Offset(5, 0).Resize(1, 6)
    o.Resize(13, 6).Columns.AutoFit
End Sub

' ---------- Sensitivity grid (data table) ----------
' Builds a two-way table around a formula cell: prompts for the output cell,
' row input and column input, then uses Excel's TABLE() via Range.Table.
Public Sub LtxInsertSensitivity()
    Dim outC As Range, rowC As Range, colC As Range, o As Range, i As Long, rv As Double, cv As Double, stp As Double
    On Error GoTo fail
    Set outC = Application.InputBox("Output cell (e.g. value per share):", LTX_APP_XL, Type:=8)
    Set rowC = Application.InputBox("Input cell that varies ACROSS the top (e.g. WACC):", LTX_APP_XL, Type:=8)
    Set colC = Application.InputBox("Input cell that varies DOWN the side (e.g. terminal growth):", LTX_APP_XL, Type:=8)
    Set o = ActiveCell
    rv = rowC.Value: cv = colC.Value
    o.Formula = "=" & outC.Address
    o.NumberFormat = outC.NumberFormat
    For i = 1 To 5
        stp = IIf(Abs(rv) < 1, 0.005, Abs(rv) * 0.05)
        o.Offset(0, i).Value = rv + (i - 3) * stp: o.Offset(0, i).NumberFormat = rowC.NumberFormat: o.Offset(0, i).Font.Color = BLUE: o.Offset(0, i).Font.Bold = True
        stp = IIf(Abs(cv) < 1, 0.005, Abs(cv) * 0.05)
        o.Offset(i, 0).Value = cv + (i - 3) * stp: o.Offset(i, 0).NumberFormat = colC.NumberFormat: o.Offset(i, 0).Font.Color = BLUE: o.Offset(i, 0).Font.Bold = True
    Next
    o.Resize(6, 6).Table rowC, colC
    o.Offset(1, 1).Resize(5, 5).NumberFormat = outC.NumberFormat
    o.Resize(6, 6).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    o.Resize(6, 6).Borders(xlInsideHorizontal).Color = RGB(200, 200, 200)
    o.Resize(1, 6).Borders(xlEdgeBottom).LineStyle = xlContinuous
    o.Resize(6, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
    Exit Sub
fail:
    If Err.Number <> 424 Then MsgBox "Sensitivity table: " & Err.Description, vbExclamation, LTX_APP_XL
End Sub

' ---------- Personal library workbook ----------
Private Function OpenLib(createIfMissing As Boolean) As Workbook
    Dim p As String, wb As Workbook
    p = LibPath()
    On Error Resume Next
    Set wb = Workbooks(Mid$(p, InStrRev(p, "\") + 1))
    On Error GoTo 0
    If wb Is Nothing Then
        If Dir(p) <> "" Then
            Set wb = Workbooks.Open(p)
        ElseIf createIfMissing Then
            Set wb = Workbooks.Add
            wb.SaveAs p, 51
        End If
    End If
    Set OpenLib = wb
End Function

Public Sub LtxLibSave()
    Dim src As Range, nm As String, libWb As Workbook, ws As Worksheet, cur As Workbook
    Set src = Selection
    Set cur = ActiveWorkbook
    nm = InputBox("Name for this analysis in your library:", LTX_APP_XL)
    If Len(nm) = 0 Then Exit Sub
    Application.ScreenUpdating = False
    Set libWb = OpenLib(True)
    On Error Resume Next
    Set ws = libWb.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = libWb.Worksheets.Add(After:=libWb.Worksheets(libWb.Worksheets.Count))
        ws.Name = Left$(nm, 31)
    Else
        ws.Cells.Clear
    End If
    src.Copy
    ws.Range("A1").PasteSpecial xlPasteAll
    ws.Range("A1").PasteSpecial xlPasteColumnWidths
    Application.CutCopyMode = False
    ws.Range("A1").Select
    libWb.Save
    libWb.Close
    cur.Activate
    Application.ScreenUpdating = True
    MsgBox "Saved '" & nm & "' to " & LibPath(), vbInformation, LTX_APP_XL
End Sub

Public Sub LtxLibInsert()
    Dim libWb As Workbook, ws As Worksheet, list As String, nm As String, cur As Workbook, dest As Range
    Set cur = ActiveWorkbook: Set dest = ActiveCell
    If Dir(LibPath()) = "" Then MsgBox "Your library is empty. Select a range and use 'Save to library'.", vbInformation, LTX_APP_XL: Exit Sub
    Application.ScreenUpdating = False
    Set libWb = OpenLib(False)
    For Each ws In libWb.Worksheets: list = list & ws.Name & vbLf: Next
    nm = InputBox("Library analyses:" & vbLf & list & vbLf & "Name to insert at the active cell:", LTX_APP_XL)
    If Len(nm) = 0 Then libWb.Close False: cur.Activate: Exit Sub
    On Error Resume Next
    Set ws = libWb.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then libWb.Close False: cur.Activate: MsgBox "Not found.", vbExclamation, LTX_APP_XL: Exit Sub
    ws.UsedRange.Copy
    dest.PasteSpecial xlPasteAll
    dest.PasteSpecial xlPasteColumnWidths
    Application.CutCopyMode = False
    libWb.Close False
    cur.Activate
    dest.Select
    Application.ScreenUpdating = True
End Sub

Public Sub LtxLibOpen()
    If Dir(LibPath()) = "" Then MsgBox "No library yet.", vbInformation, LTX_APP_XL: Exit Sub
    Workbooks.Open LibPath()
End Sub

Public Sub LtxLibSetPath()
    Dim p As String
    p = InputBox("Full path of the library workbook (.xlsx). A shared drive works too:", LTX_APP_XL, LibPath())
    If Len(p) > 0 Then XlSet "LibraryPath", p
End Sub
