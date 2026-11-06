Attribute VB_Name = "modLtxXlPE"
' =====================================================================
'  Lattice for Excel - private equity / hedge fund generators
'  LBO quick model, equity distribution waterfall (pref / catch-up / carry),
'  scenario switch block, timeline header, credit statistics block.
'  All written as live formulas at the active cell, Lattice colour code.
' =====================================================================
Option Explicit

Private Const BLUE As Long = 16711680
Private Const BLACK As Long = 0

Private Sub Lb(c As Range, txt As String, Optional bold As Boolean = False)
    c.Value = txt: c.Font.Bold = bold: c.Font.Color = BLACK
End Sub
Private Sub Ip(c As Range, v As Variant, Optional fmt As String = "")
    c.Value = v: c.Font.Color = BLUE
    If Len(fmt) > 0 Then c.NumberFormat = fmt
End Sub
Private Sub Fm(c As Range, f As String, Optional fmt As String = "")
    c.Formula = f: c.Font.Color = BLACK
    If Len(fmt) > 0 Then c.NumberFormat = fmt
End Sub
Private Sub Hd(r As Range)
    r.Font.Bold = True: r.Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub
Private Function Ad(r As Range) As String: Ad = r.Address(False, False): End Function
Private Function AdAbs(r As Range) As String: AdAbs = r.Address(True, True): End Function

' ---------------------------------------------------------------
'  LBO quick model: sources & uses, two debt tranches with cash sweep,
'  5-year operating case, exit at multiple, IRR / MOIC, exit sensitivity.
' ---------------------------------------------------------------
Public Sub LtxInsertLBO()
    Dim o As Range, y As Long, c As Long
    Const N As Long = 5
    Set o = ActiveCell
    Application.ScreenUpdating = False
    ' ---- assumptions ----
    Lb o, "LBO quick model", True
    Lb o.Offset(1, 0), "Entry assumptions", True
    Lb o.Offset(2, 0), "LTM EBITDA":            Ip o.Offset(2, 1), 100, "#,##0.0"
    Lb o.Offset(3, 0), "Entry EV / EBITDA":     Ip o.Offset(3, 1), 10, "0.0x"
    Lb o.Offset(4, 0), "Enterprise value":      Fm o.Offset(4, 1), "=" & Ad(o.Offset(2, 1)) & "*" & Ad(o.Offset(3, 1)), "#,##0.0"
    Lb o.Offset(5, 0), "Senior debt (x EBITDA)": Ip o.Offset(5, 1), 4, "0.0x"
    Lb o.Offset(6, 0), "Mezz / 2nd lien (x EBITDA)": Ip o.Offset(6, 1), 1.5, "0.0x"
    Lb o.Offset(7, 0), "Fees & expenses (% EV)": Ip o.Offset(7, 1), 0.02, "0.0%"
    Lb o.Offset(8, 0), "Minimum cash":          Ip o.Offset(8, 1), 10, "#,##0.0"
    Lb o.Offset(9, 0), "Senior rate":           Ip o.Offset(9, 1), 0.065, "0.00%"
    Lb o.Offset(10, 0), "Mezz rate":            Ip o.Offset(10, 1), 0.11, "0.00%"
    Lb o.Offset(11, 0), "Senior amort. % p.a.": Ip o.Offset(11, 1), 0.05, "0.0%"
    Lb o.Offset(12, 0), "Cash sweep %":         Ip o.Offset(12, 1), 0.75, "0%"
    Lb o.Offset(13, 0), "Tax rate":             Ip o.Offset(13, 1), 0.25, "0.0%"
    Lb o.Offset(14, 0), "Exit EV / EBITDA":     Ip o.Offset(14, 1), 10, "0.0x"
    Lb o.Offset(15, 0), "Exit year":            Ip o.Offset(15, 1), 5, "0"
    ' ---- sources & uses ----
    Lb o.Offset(1, 4), "Sources", True:            Lb o.Offset(1, 7), "Uses", True
    Lb o.Offset(2, 4), "Senior debt":              Fm o.Offset(2, 5), "=" & Ad(o.Offset(2, 1)) & "*" & Ad(o.Offset(5, 1)), "#,##0.0"
    Lb o.Offset(3, 4), "Mezzanine":                Fm o.Offset(3, 5), "=" & Ad(o.Offset(2, 1)) & "*" & Ad(o.Offset(6, 1)), "#,##0.0"
    Lb o.Offset(4, 4), "Sponsor equity":           Fm o.Offset(4, 5), "=" & Ad(o.Offset(5, 8)) & "-" & Ad(o.Offset(2, 5)) & "-" & Ad(o.Offset(3, 5)), "#,##0.0"
    Lb o.Offset(5, 4), "Total sources", True:      Fm o.Offset(5, 5), "=SUM(" & Ad(o.Offset(2, 5)) & ":" & Ad(o.Offset(4, 5)) & ")", "#,##0.0"
    Lb o.Offset(2, 7), "Purchase EV":              Fm o.Offset(2, 8), "=" & Ad(o.Offset(4, 1)), "#,##0.0"
    Lb o.Offset(3, 7), "Fees":                     Fm o.Offset(3, 8), "=" & Ad(o.Offset(4, 1)) & "*" & Ad(o.Offset(7, 1)), "#,##0.0"
    Lb o.Offset(4, 7), "Cash to balance sheet":    Fm o.Offset(4, 8), "=" & Ad(o.Offset(8, 1)), "#,##0.0"
    Lb o.Offset(5, 7), "Total uses", True:         Fm o.Offset(5, 8), "=SUM(" & Ad(o.Offset(2, 8)) & ":" & Ad(o.Offset(4, 8)) & ")", "#,##0.0"
    Lb o.Offset(6, 4), "Equity % of sources":      Fm o.Offset(6, 5), "=" & Ad(o.Offset(4, 5)) & "/" & Ad(o.Offset(5, 5)), "0.0%"
    Hd o.Offset(5, 4).Resize(1, 2): Hd o.Offset(5, 7).Resize(1, 2)
    ' ---- operating case ----
    Dim r0 As Long: r0 = 18
    Lb o.Offset(r0, 0), "Operating case", True
    Lb o.Offset(r0 + 1, 0), "Year"
    For y = 0 To N: Lb o.Offset(r0 + 1, 1 + y), IIf(y = 0, "Entry", "Yr " & y), True: Next
    Hd o.Offset(r0 + 1, 0).Resize(1, N + 2)
    Lb o.Offset(r0 + 2, 0), "EBITDA growth":   For y = 1 To N: Ip o.Offset(r0 + 2, 1 + y), 0.06, "0.0%": Next
    Lb o.Offset(r0 + 3, 0), "EBITDA":          Fm o.Offset(r0 + 3, 1), "=" & Ad(o.Offset(2, 1)), "#,##0.0"
    For y = 1 To N: Fm o.Offset(r0 + 3, 1 + y), "=" & Ad(o.Offset(r0 + 3, y)) & "*(1+" & Ad(o.Offset(r0 + 2, 1 + y)) & ")", "#,##0.0": Next
    Lb o.Offset(r0 + 4, 0), "Capex % EBITDA":  For y = 1 To N: Ip o.Offset(r0 + 4, 1 + y), 0.15, "0.0%": Next
    Lb o.Offset(r0 + 5, 0), "NWC change % EBITDA": For y = 1 To N: Ip o.Offset(r0 + 5, 1 + y), 0.03, "0.0%": Next
    Lb o.Offset(r0 + 6, 0), "D&A % EBITDA":    For y = 1 To N: Ip o.Offset(r0 + 6, 1 + y), 0.2, "0.0%": Next
    ' ---- debt schedule ----
    Dim d0 As Long: d0 = r0 + 8
    Lb o.Offset(d0, 0), "Cash flow and debt", True
    Lb o.Offset(d0 + 1, 0), "Senior opening":   Fm o.Offset(d0 + 1, 2), "=" & Ad(o.Offset(2, 5)), "#,##0.0"
    Lb o.Offset(d0 + 2, 0), "Mezz opening":     Fm o.Offset(d0 + 2, 2), "=" & Ad(o.Offset(3, 5)), "#,##0.0"
    Lb o.Offset(d0 + 3, 0), "Interest (senior + mezz)"
    Lb o.Offset(d0 + 4, 0), "EBIT":
    Lb o.Offset(d0 + 5, 0), "Taxes"
    Lb o.Offset(d0 + 6, 0), "Capex"
    Lb o.Offset(d0 + 7, 0), "NWC change"
    Lb o.Offset(d0 + 8, 0), "Cash available for debt service", True
    Lb o.Offset(d0 + 9, 0), "Senior mandatory amort."
    Lb o.Offset(d0 + 10, 0), "Cash sweep (senior then mezz)"
    Lb o.Offset(d0 + 11, 0), "Senior closing", True
    Lb o.Offset(d0 + 12, 0), "Mezz closing", True
    Lb o.Offset(d0 + 13, 0), "Cash balance"
    Lb o.Offset(d0 + 14, 0), "Net debt", True
    Lb o.Offset(d0 + 15, 0), "Net debt / EBITDA"
    Fm o.Offset(d0 + 13, 1), "=" & Ad(o.Offset(8, 1)), "#,##0.0"
    Fm o.Offset(d0 + 11, 1), "=" & Ad(o.Offset(2, 5)), "#,##0.0"
    Fm o.Offset(d0 + 12, 1), "=" & Ad(o.Offset(3, 5)), "#,##0.0"
    Fm o.Offset(d0 + 14, 1), "=" & Ad(o.Offset(d0 + 11, 1)) & "+" & Ad(o.Offset(d0 + 12, 1)) & "-" & Ad(o.Offset(d0 + 13, 1)), "#,##0.0"
    Fm o.Offset(d0 + 15, 1), "=" & Ad(o.Offset(d0 + 14, 1)) & "/" & Ad(o.Offset(r0 + 3, 1)), "0.0x"
    For y = 1 To N
        c = 1 + y
        Fm o.Offset(d0 + 1, c), "=" & Ad(o.Offset(d0 + 11, c - 1)), "#,##0.0"
        Fm o.Offset(d0 + 2, c), "=" & Ad(o.Offset(d0 + 12, c - 1)), "#,##0.0"
        Fm o.Offset(d0 + 3, c), "=-(" & Ad(o.Offset(d0 + 1, c)) & "*" & AdAbs(o.Offset(9, 1)) & "+" & Ad(o.Offset(d0 + 2, c)) & "*" & AdAbs(o.Offset(10, 1)) & ")", "#,##0.0;(#,##0.0)"
        Fm o.Offset(d0 + 4, c), "=" & Ad(o.Offset(r0 + 3, c)) & "*(1-" & Ad(o.Offset(r0 + 6, c)) & ")", "#,##0.0"
        Fm o.Offset(d0 + 5, c), "=-MAX(0,(" & Ad(o.Offset(d0 + 4, c)) & "+" & Ad(o.Offset(d0 + 3, c)) & ")*" & AdAbs(o.Offset(13, 1)) & ")", "#,##0.0;(#,##0.0)"
        Fm o.Offset(d0 + 6, c), "=-" & Ad(o.Offset(r0 + 3, c)) & "*" & Ad(o.Offset(r0 + 4, c)), "#,##0.0;(#,##0.0)"
        Fm o.Offset(d0 + 7, c), "=-" & Ad(o.Offset(r0 + 3, c)) & "*" & Ad(o.Offset(r0 + 5, c)), "#,##0.0;(#,##0.0)"
        Fm o.Offset(d0 + 8, c), "=" & Ad(o.Offset(r0 + 3, c)) & "+" & Ad(o.Offset(d0 + 3, c)) & "+" & Ad(o.Offset(d0 + 5, c)) & "+" & Ad(o.Offset(d0 + 6, c)) & "+" & Ad(o.Offset(d0 + 7, c)), "#,##0.0;(#,##0.0)"
        Fm o.Offset(d0 + 9, c), "=-MIN(" & Ad(o.Offset(d0 + 1, c)) & "," & AdAbs(o.Offset(2, 5)) & "*" & AdAbs(o.Offset(11, 1)) & ")", "#,##0.0;(#,##0.0)"
        Fm o.Offset(d0 + 10, c), "=-MIN(" & Ad(o.Offset(d0 + 1, c)) & "+" & Ad(o.Offset(d0 + 9, c)) & "+" & Ad(o.Offset(d0 + 2, c)) & ",MAX(0,(" & Ad(o.Offset(d0 + 8, c)) & "+" & Ad(o.Offset(d0 + 9, c)) & ")*" & AdAbs(o.Offset(12, 1)) & "))", "#,##0.0;(#,##0.0)"
        Fm o.Offset(d0 + 11, c), "=MAX(0," & Ad(o.Offset(d0 + 1, c)) & "+" & Ad(o.Offset(d0 + 9, c)) & "+MAX(" & Ad(o.Offset(d0 + 10, c)) & ",-(" & Ad(o.Offset(d0 + 1, c)) & "+" & Ad(o.Offset(d0 + 9, c)) & ")))", "#,##0.0"
        Fm o.Offset(d0 + 12, c), "=MAX(0," & Ad(o.Offset(d0 + 2, c)) & "+MIN(0," & Ad(o.Offset(d0 + 10, c)) & "+" & Ad(o.Offset(d0 + 1, c)) & "+" & Ad(o.Offset(d0 + 9, c)) & "))", "#,##0.0"
        Fm o.Offset(d0 + 13, c), "=" & Ad(o.Offset(d0 + 13, c - 1)) & "+" & Ad(o.Offset(d0 + 8, c)) & "+" & Ad(o.Offset(d0 + 9, c)) & "+" & Ad(o.Offset(d0 + 10, c)), "#,##0.0"
        Fm o.Offset(d0 + 14, c), "=" & Ad(o.Offset(d0 + 11, c)) & "+" & Ad(o.Offset(d0 + 12, c)) & "-" & Ad(o.Offset(d0 + 13, c)), "#,##0.0"
        Fm o.Offset(d0 + 15, c), "=" & Ad(o.Offset(d0 + 14, c)) & "/" & Ad(o.Offset(r0 + 3, c)), "0.0x"
    Next
    Hd o.Offset(d0 + 8, 0).Resize(1, N + 2): Hd o.Offset(d0 + 14, 0).Resize(1, N + 2)
    ' ---- returns ----
    Dim x0 As Long: x0 = d0 + 17
    Lb o.Offset(x0, 0), "Exit and returns", True
    Lb o.Offset(x0 + 1, 0), "Exit EBITDA":       Fm o.Offset(x0 + 1, 1), "=INDEX(" & o.Offset(r0 + 3, 2).Resize(1, N).Address & "," & Ad(o.Offset(15, 1)) & ")", "#,##0.0"
    Lb o.Offset(x0 + 2, 0), "Exit EV":           Fm o.Offset(x0 + 2, 1), "=" & Ad(o.Offset(x0 + 1, 1)) & "*" & Ad(o.Offset(14, 1)), "#,##0.0"
    Lb o.Offset(x0 + 3, 0), "Less: net debt at exit": Fm o.Offset(x0 + 3, 1), "=-INDEX(" & o.Offset(d0 + 14, 2).Resize(1, N).Address & "," & Ad(o.Offset(15, 1)) & ")", "#,##0.0;(#,##0.0)"
    Lb o.Offset(x0 + 4, 0), "Equity at exit", True: Fm o.Offset(x0 + 4, 1), "=" & Ad(o.Offset(x0 + 2, 1)) & "+" & Ad(o.Offset(x0 + 3, 1)), "#,##0.0"
    Lb o.Offset(x0 + 5, 0), "Equity invested":   Fm o.Offset(x0 + 5, 1), "=" & Ad(o.Offset(4, 5)), "#,##0.0"
    Lb o.Offset(x0 + 6, 0), "MOIC", True:        Fm o.Offset(x0 + 6, 1), "=" & Ad(o.Offset(x0 + 4, 1)) & "/" & Ad(o.Offset(x0 + 5, 1)), "0.00x"
    Lb o.Offset(x0 + 7, 0), "IRR", True:         Fm o.Offset(x0 + 7, 1), "=" & Ad(o.Offset(x0 + 6, 1)) & "^(1/" & Ad(o.Offset(15, 1)) & ")-1", "0.0%"
    Hd o.Offset(x0 + 4, 0).Resize(1, 2)
    ' ---- exit multiple x exit year sensitivity (IRR) ----
    Lb o.Offset(x0 + 9, 0), "IRR sensitivity: exit multiple (down) vs exit year (across)", True
    For y = 3 To 5: Ip o.Offset(x0 + 10, y - 2), y, "0": o.Offset(x0 + 10, y - 2).Font.Bold = True: Next
    Dim m As Long
    For m = 0 To 4
        Ip o.Offset(x0 + 11 + m, 0), 8 + m, "0.0x": o.Offset(x0 + 11 + m, 0).Font.Bold = True
        For y = 3 To 5
            ' equity at exit = EBITDA(y)*mult - netdebt(y)
            Fm o.Offset(x0 + 11 + m, y - 2), "=IFERROR(((INDEX(" & o.Offset(r0 + 3, 2).Resize(1, N).Address & "," & Ad(o.Offset(x0 + 10, y - 2)) & ")*" & Ad(o.Offset(x0 + 11 + m, 0)) & _
                "-INDEX(" & o.Offset(d0 + 14, 2).Resize(1, N).Address & "," & Ad(o.Offset(x0 + 10, y - 2)) & "))/" & AdAbs(o.Offset(4, 5)) & ")^(1/" & Ad(o.Offset(x0 + 10, y - 2)) & ")-1,""n.m."")", "0.0%"
        Next
    Next
    o.Resize(x0 + 16, 10).Columns.AutoFit
    o.Offset(0, 0).Select
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------
'  Equity distribution waterfall: return of capital, preferred return,
'  GP catch-up, then carry split. European (whole-fund) style.
' ---------------------------------------------------------------
Public Sub LtxInsertDistributionWaterfall()
    Dim o As Range
    Set o = ActiveCell
    Lb o, "Distribution waterfall (European)", True
    Lb o.Offset(1, 0), "Total proceeds to distribute": Ip o.Offset(1, 1), 250, "#,##0.0"
    Lb o.Offset(2, 0), "LP + GP contributed capital":  Ip o.Offset(2, 1), 100, "#,##0.0"
    Lb o.Offset(3, 0), "Preferred return (hurdle)":    Ip o.Offset(3, 1), 0.08, "0.0%"
    Lb o.Offset(4, 0), "Years held":                   Ip o.Offset(4, 1), 4, "0.0"
    Lb o.Offset(5, 0), "Carry":                        Ip o.Offset(5, 1), 0.2, "0%"
    Lb o.Offset(6, 0), "GP catch-up share":            Ip o.Offset(6, 1), 1, "0%"
    Lb o.Offset(7, 0), "GP commitment % of capital":   Ip o.Offset(7, 1), 0.02, "0.0%"
    Lb o.Offset(9, 0), "Tier", True: Lb o.Offset(9, 1), "Amount", True: Lb o.Offset(9, 2), "To LPs", True: Lb o.Offset(9, 3), "To GP", True
    Hd o.Offset(9, 0).Resize(1, 4)
    Lb o.Offset(10, 0), "1. Return of capital"
    Fm o.Offset(10, 1), "=MIN(" & Ad(o.Offset(1, 1)) & "," & Ad(o.Offset(2, 1)) & ")", "#,##0.0"
    Fm o.Offset(10, 2), "=" & Ad(o.Offset(10, 1)) & "*(1-" & Ad(o.Offset(7, 1)) & ")", "#,##0.0"
    Fm o.Offset(10, 3), "=" & Ad(o.Offset(10, 1)) & "*" & Ad(o.Offset(7, 1)), "#,##0.0"
    Lb o.Offset(11, 0), "2. Preferred return"
    Fm o.Offset(11, 1), "=MIN(" & Ad(o.Offset(1, 1)) & "-" & Ad(o.Offset(10, 1)) & "," & Ad(o.Offset(2, 1)) & "*((1+" & Ad(o.Offset(3, 1)) & ")^" & Ad(o.Offset(4, 1)) & "-1))", "#,##0.0"
    Fm o.Offset(11, 2), "=" & Ad(o.Offset(11, 1)) & "*(1-" & Ad(o.Offset(7, 1)) & ")", "#,##0.0"
    Fm o.Offset(11, 3), "=" & Ad(o.Offset(11, 1)) & "*" & Ad(o.Offset(7, 1)), "#,##0.0"
    Lb o.Offset(12, 0), "3. GP catch-up"
    ' catch-up until GP has carry% of (pref + catch-up): catchup = pref * carry/(1-carry) when 100% catch-up
    Fm o.Offset(12, 1), "=MIN(" & Ad(o.Offset(1, 1)) & "-" & Ad(o.Offset(10, 1)) & "-" & Ad(o.Offset(11, 1)) & ",IF(" & Ad(o.Offset(6, 1)) & ">" & Ad(o.Offset(5, 1)) & "," & Ad(o.Offset(11, 2)) & "*" & Ad(o.Offset(5, 1)) & "/(" & Ad(o.Offset(6, 1)) & "-" & Ad(o.Offset(5, 1)) & "),0))", "#,##0.0"
    Fm o.Offset(12, 2), "=" & Ad(o.Offset(12, 1)) & "*(1-" & Ad(o.Offset(6, 1)) & ")", "#,##0.0"
    Fm o.Offset(12, 3), "=" & Ad(o.Offset(12, 1)) & "*" & Ad(o.Offset(6, 1)), "#,##0.0"
    Lb o.Offset(13, 0), "4. Carried interest split"
    Fm o.Offset(13, 1), "=MAX(0," & Ad(o.Offset(1, 1)) & "-SUM(" & Ad(o.Offset(10, 1)) & ":" & Ad(o.Offset(12, 1)) & "))", "#,##0.0"
    Fm o.Offset(13, 2), "=" & Ad(o.Offset(13, 1)) & "*(1-" & Ad(o.Offset(5, 1)) & ")", "#,##0.0"
    Fm o.Offset(13, 3), "=" & Ad(o.Offset(13, 1)) & "*" & Ad(o.Offset(5, 1)), "#,##0.0"
    Lb o.Offset(14, 0), "Total", True
    Fm o.Offset(14, 1), "=SUM(" & Ad(o.Offset(10, 1)) & ":" & Ad(o.Offset(13, 1)) & ")", "#,##0.0"
    Fm o.Offset(14, 2), "=SUM(" & Ad(o.Offset(10, 2)) & ":" & Ad(o.Offset(13, 2)) & ")", "#,##0.0"
    Fm o.Offset(14, 3), "=SUM(" & Ad(o.Offset(10, 3)) & ":" & Ad(o.Offset(13, 3)) & ")", "#,##0.0"
    Hd o.Offset(14, 0).Resize(1, 4)
    Lb o.Offset(16, 0), "LP net MOIC":   Fm o.Offset(16, 1), "=" & Ad(o.Offset(14, 2)) & "/(" & Ad(o.Offset(2, 1)) & "*(1-" & Ad(o.Offset(7, 1)) & "))", "0.00x"
    Lb o.Offset(17, 0), "LP net IRR (approx.)": Fm o.Offset(17, 1), "=" & Ad(o.Offset(16, 1)) & "^(1/" & Ad(o.Offset(4, 1)) & ")-1", "0.0%"
    Lb o.Offset(18, 0), "GP carry as % of profit": Fm o.Offset(18, 1), "=IFERROR((" & Ad(o.Offset(14, 3)) & "-" & Ad(o.Offset(10, 3)) & ")/(" & Ad(o.Offset(1, 1)) & "-" & Ad(o.Offset(2, 1)) & "),0)", "0.0%"
    Lb o.Offset(19, 0), "Check (distributed = proceeds)": Fm o.Offset(19, 1), "=ROUND(" & Ad(o.Offset(14, 1)) & "-" & Ad(o.Offset(1, 1)) & ",6)=0"
    o.Resize(20, 4).Columns.AutoFit
End Sub

' ---------------------------------------------------------------
'  Scenario switch: a selector cell plus Base / Upside / Downside
'  columns for N drivers and a live column that follows the selector.
' ---------------------------------------------------------------
Public Sub LtxInsertScenarioBlock()
    Dim o As Range, n As Long, i As Long, s As String
    Set o = ActiveCell
    s = InputBox("Driver names (comma separated):", LTX_APP_XL, "Revenue growth,EBITDA margin,Capex % sales,Exit multiple")
    If Len(s) = 0 Then Exit Sub
    Dim drv As Variant: drv = Split(s, ",")
    n = UBound(drv) + 1
    Lb o, "Scenario", True: Ip o.Offset(0, 1), 1, "0"
    o.Offset(0, 2).Value = "1 = Base, 2 = Upside, 3 = Downside": o.Offset(0, 2).Font.Italic = True: o.Offset(0, 2).Font.Color = RGB(128, 128, 128)
    On Error Resume Next
    ActiveWorkbook.Names.Add "Scenario", "=" & o.Offset(0, 1).Address(True, True, xlA1, True)
    On Error GoTo 0
    Lb o.Offset(2, 0), "Driver", True: Lb o.Offset(2, 1), "Live", True: Lb o.Offset(2, 2), "Base", True: Lb o.Offset(2, 3), "Upside", True: Lb o.Offset(2, 4), "Downside", True
    Hd o.Offset(2, 0).Resize(1, 5)
    For i = 0 To n - 1
        Lb o.Offset(3 + i, 0), Trim$(drv(i))
        Ip o.Offset(3 + i, 2), 0.05, "0.0%": Ip o.Offset(3 + i, 3), 0.07, "0.0%": Ip o.Offset(3 + i, 4), 0.03, "0.0%"
        Fm o.Offset(3 + i, 1), "=CHOOSE(" & AdAbs(o.Offset(0, 1)) & "," & Ad(o.Offset(3 + i, 2)) & "," & Ad(o.Offset(3 + i, 3)) & "," & Ad(o.Offset(3 + i, 4)) & ")", "0.0%"
        o.Offset(3 + i, 1).Interior.Color = RGB(242, 242, 242)
    Next
    o.Offset(0, 1).Interior.Color = RGB(255, 242, 204)
    o.Resize(4 + n, 5).Columns.AutoFit
    MsgBox "Link your model to the 'Live' column. Change the selector cell (named 'Scenario') to switch cases.", vbInformation, LTX_APP_XL
End Sub

' ---------------------------------------------------------------
'  Timeline header: period end dates, labels, year/quarter numbers and
'  flags (actual vs forecast, fiscal year) across the row.
' ---------------------------------------------------------------
Public Sub LtxInsertTimeline()
    Dim o As Range, n As Long, i As Long, per As String, startD As Date, lastA As Date, s As String
    Set o = ActiveCell
    per = InputBox("Period: Y (annual), Q (quarterly) or M (monthly):", LTX_APP_XL, "Y")
    If Len(per) = 0 Then Exit Sub
    s = InputBox("First period end date (dd/mm/yyyy):", LTX_APP_XL, Format$(DateSerial(Year(Date), 12, 31), "dd/mm/yyyy"))
    If Len(s) = 0 Then Exit Sub
    startD = CDate(s)
    n = Val(InputBox("Number of periods:", LTX_APP_XL, IIf(UCase$(per) = "Y", "7", IIf(UCase$(per) = "Q", "12", "24"))))
    If n < 1 Then Exit Sub
    s = InputBox("Last ACTUAL period end (blank = all forecast):", LTX_APP_XL, "")
    If Len(s) > 0 Then lastA = CDate(s)
    Lb o, "Period end", True
    Lb o.Offset(1, 0), "Label", True
    Lb o.Offset(2, 0), "Fiscal year", True
    Lb o.Offset(3, 0), "Period #", True
    Lb o.Offset(4, 0), "A / E", True
    Lb o.Offset(5, 0), "Days in period", True
    Ip o.Offset(0, 1), startD, "dd-mmm-yy"
    For i = 1 To n
        If i > 1 Then
            Select Case UCase$(per)
                Case "Y": Fm o.Offset(0, i), "=EDATE(" & Ad(o.Offset(0, i - 1)) & ",12)", "dd-mmm-yy"
                Case "Q": Fm o.Offset(0, i), "=EOMONTH(" & Ad(o.Offset(0, i - 1)) & ",3)", "dd-mmm-yy"
                Case Else: Fm o.Offset(0, i), "=EOMONTH(" & Ad(o.Offset(0, i - 1)) & ",1)", "dd-mmm-yy"
            End Select
        End If
        Select Case UCase$(per)
            Case "Y": Fm o.Offset(1, i), "=TEXT(" & Ad(o.Offset(0, i)) & ",""yyyy"")&" & Ad(o.Offset(4, i))
            Case "Q": Fm o.Offset(1, i), "=""Q""&ROUNDUP(MONTH(" & Ad(o.Offset(0, i)) & ")/3,0)&"" ""&TEXT(" & Ad(o.Offset(0, i)) & ",""yy"")&" & Ad(o.Offset(4, i))
            Case Else: Fm o.Offset(1, i), "=TEXT(" & Ad(o.Offset(0, i)) & ",""mmm-yy"")&" & Ad(o.Offset(4, i))
        End Select
        Fm o.Offset(2, i), "=YEAR(" & Ad(o.Offset(0, i)) & ")", "0"
        Fm o.Offset(3, i), "=" & i, "0"
        If lastA > 0 Then
            Fm o.Offset(4, i), "=IF(" & Ad(o.Offset(0, i)) & "<=DATE(" & Year(lastA) & "," & Month(lastA) & "," & Day(lastA) & "),""A"",""E"")"
        Else
            Fm o.Offset(4, i), "=""E"""
        End If
        If i = 1 Then
            Fm o.Offset(5, i), "=" & Ad(o.Offset(0, i)) & "-EDATE(" & Ad(o.Offset(0, i)) & ",-" & IIf(UCase$(per) = "Y", 12, IIf(UCase$(per) = "Q", 3, 1)) & ")", "0"
        Else
            Fm o.Offset(5, i), "=" & Ad(o.Offset(0, i)) & "-" & Ad(o.Offset(0, i - 1)), "0"
        End If
        o.Offset(1, i).HorizontalAlignment = xlRight: o.Offset(1, i).Font.Bold = True
    Next
    o.Offset(1, 0).Resize(1, n + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    o.Resize(6, n + 1).Columns.AutoFit
End Sub

' ---------------------------------------------------------------
'  Credit statistics block from four cells: EBITDA, total debt, cash, interest
' ---------------------------------------------------------------
Public Sub LtxInsertCreditStats()
    Dim o As Range
    Set o = ActiveCell
    Lb o, "Credit statistics", True
    Lb o.Offset(1, 0), "EBITDA":          Ip o.Offset(1, 1), 100, "#,##0.0"
    Lb o.Offset(2, 0), "Total debt":      Ip o.Offset(2, 1), 450, "#,##0.0"
    Lb o.Offset(3, 0), "Cash":            Ip o.Offset(3, 1), 30, "#,##0.0"
    Lb o.Offset(4, 0), "Interest expense": Ip o.Offset(4, 1), 32, "#,##0.0"
    Lb o.Offset(5, 0), "Capex":           Ip o.Offset(5, 1), 15, "#,##0.0"
    Lb o.Offset(6, 0), "Cash taxes":      Ip o.Offset(6, 1), 12, "#,##0.0"
    Lb o.Offset(7, 0), "Enterprise value": Ip o.Offset(7, 1), 1000, "#,##0.0"
    Lb o.Offset(9, 0), "Gross leverage":  Fm o.Offset(9, 1), "=" & Ad(o.Offset(2, 1)) & "/" & Ad(o.Offset(1, 1)), "0.0x"
    Lb o.Offset(10, 0), "Net leverage":   Fm o.Offset(10, 1), "=(" & Ad(o.Offset(2, 1)) & "-" & Ad(o.Offset(3, 1)) & ")/" & Ad(o.Offset(1, 1)), "0.0x"
    Lb o.Offset(11, 0), "Interest cover (EBITDA / interest)": Fm o.Offset(11, 1), "=" & Ad(o.Offset(1, 1)) & "/" & Ad(o.Offset(4, 1)), "0.0x"
    Lb o.Offset(12, 0), "(EBITDA - capex) / interest": Fm o.Offset(12, 1), "=(" & Ad(o.Offset(1, 1)) & "-" & Ad(o.Offset(5, 1)) & ")/" & Ad(o.Offset(4, 1)), "0.0x"
    Lb o.Offset(13, 0), "FCF (EBITDA - capex - tax - interest)": Fm o.Offset(13, 1), "=" & Ad(o.Offset(1, 1)) & "-" & Ad(o.Offset(5, 1)) & "-" & Ad(o.Offset(6, 1)) & "-" & Ad(o.Offset(4, 1)), "#,##0.0"
    Lb o.Offset(14, 0), "FCF / debt":     Fm o.Offset(14, 1), "=" & Ad(o.Offset(13, 1)) & "/" & Ad(o.Offset(2, 1)), "0.0%"
    Lb o.Offset(15, 0), "Debt / EV (LTV)": Fm o.Offset(15, 1), "=" & Ad(o.Offset(2, 1)) & "/" & Ad(o.Offset(7, 1)), "0.0%"
    Lb o.Offset(16, 0), "Equity cushion": Fm o.Offset(16, 1), "=1-" & Ad(o.Offset(15, 1)), "0.0%"
    Lb o.Offset(17, 0), "Years to repay (debt / FCF)": Fm o.Offset(17, 1), "=IFERROR(" & Ad(o.Offset(2, 1)) & "/" & Ad(o.Offset(13, 1)) & ",""n.m."")", "0.0"
    Hd o.Offset(8, 0).Resize(1, 2)
    o.Resize(18, 2).Columns.AutoFit
End Sub

' ---------------------------------------------------------------
'  Accretion / dilution quick check (all-stock or cash mix)
' ---------------------------------------------------------------
Public Sub LtxInsertAccretionDilution()
    Dim o As Range
    Set o = ActiveCell
    Lb o, "Accretion / dilution", True
    Lb o.Offset(1, 0), "", False: Lb o.Offset(1, 1), "Acquirer", True: Lb o.Offset(1, 2), "Target", True
    Hd o.Offset(1, 0).Resize(1, 3)
    Lb o.Offset(2, 0), "Net income":      Ip o.Offset(2, 1), 500, "#,##0.0": Ip o.Offset(2, 2), 120, "#,##0.0"
    Lb o.Offset(3, 0), "Shares out (m)":  Ip o.Offset(3, 1), 200, "#,##0.0": Ip o.Offset(3, 2), 60, "#,##0.0"
    Lb o.Offset(4, 0), "Share price":     Ip o.Offset(4, 1), 50, "#,##0.00": Ip o.Offset(4, 2), 30, "#,##0.00"
    Lb o.Offset(5, 0), "EPS":             Fm o.Offset(5, 1), "=" & Ad(o.Offset(2, 1)) & "/" & Ad(o.Offset(3, 1)), "#,##0.00": Fm o.Offset(5, 2), "=" & Ad(o.Offset(2, 2)) & "/" & Ad(o.Offset(3, 2)), "#,##0.00"
    Lb o.Offset(6, 0), "P / E":           Fm o.Offset(6, 1), "=" & Ad(o.Offset(4, 1)) & "/" & Ad(o.Offset(5, 1)), "0.0x": Fm o.Offset(6, 2), "=" & Ad(o.Offset(4, 2)) & "/" & Ad(o.Offset(5, 2)), "0.0x"
    Lb o.Offset(8, 0), "Offer premium":   Ip o.Offset(8, 1), 0.3, "0.0%"
    Lb o.Offset(9, 0), "% stock consideration": Ip o.Offset(9, 1), 0.5, "0%"
    Lb o.Offset(10, 0), "Cost of debt (pre-tax)": Ip o.Offset(10, 1), 0.05, "0.0%"
    Lb o.Offset(11, 0), "Tax rate":       Ip o.Offset(11, 1), 0.25, "0.0%"
    Lb o.Offset(12, 0), "Pre-tax synergies": Ip o.Offset(12, 1), 40, "#,##0.0"
    Lb o.Offset(14, 0), "Offer price":    Fm o.Offset(14, 1), "=" & Ad(o.Offset(4, 2)) & "*(1+" & Ad(o.Offset(8, 1)) & ")", "#,##0.00"
    Lb o.Offset(15, 0), "Equity purchase price": Fm o.Offset(15, 1), "=" & Ad(o.Offset(14, 1)) & "*" & Ad(o.Offset(3, 2)), "#,##0.0"
    Lb o.Offset(16, 0), "New shares issued (m)": Fm o.Offset(16, 1), "=" & Ad(o.Offset(15, 1)) & "*" & Ad(o.Offset(9, 1)) & "/" & Ad(o.Offset(4, 1)), "#,##0.0"
    Lb o.Offset(17, 0), "New debt":       Fm o.Offset(17, 1), "=" & Ad(o.Offset(15, 1)) & "*(1-" & Ad(o.Offset(9, 1)) & ")", "#,##0.0"
    Lb o.Offset(18, 0), "After-tax interest": Fm o.Offset(18, 1), "=" & Ad(o.Offset(17, 1)) & "*" & Ad(o.Offset(10, 1)) & "*(1-" & Ad(o.Offset(11, 1)) & ")", "#,##0.0"
    Lb o.Offset(19, 0), "After-tax synergies": Fm o.Offset(19, 1), "=" & Ad(o.Offset(12, 1)) & "*(1-" & Ad(o.Offset(11, 1)) & ")", "#,##0.0"
    Lb o.Offset(20, 0), "Pro forma net income": Fm o.Offset(20, 1), "=" & Ad(o.Offset(2, 1)) & "+" & Ad(o.Offset(2, 2)) & "-" & Ad(o.Offset(18, 1)) & "+" & Ad(o.Offset(19, 1)), "#,##0.0"
    Lb o.Offset(21, 0), "Pro forma shares (m)": Fm o.Offset(21, 1), "=" & Ad(o.Offset(3, 1)) & "+" & Ad(o.Offset(16, 1)), "#,##0.0"
    Lb o.Offset(22, 0), "Pro forma EPS", True: Fm o.Offset(22, 1), "=" & Ad(o.Offset(20, 1)) & "/" & Ad(o.Offset(21, 1)), "#,##0.00"
    Lb o.Offset(23, 0), "Accretion / (dilution)", True: Fm o.Offset(23, 1), "=" & Ad(o.Offset(22, 1)) & "/" & Ad(o.Offset(5, 1)) & "-1", "0.0%;(0.0%)"
    Lb o.Offset(24, 0), "Breakeven synergies (pre-tax)": Fm o.Offset(24, 1), "=MAX(0,(" & Ad(o.Offset(5, 1)) & "*" & Ad(o.Offset(21, 1)) & "-" & Ad(o.Offset(2, 1)) & "-" & Ad(o.Offset(2, 2)) & "+" & Ad(o.Offset(18, 1)) & ")/(1-" & Ad(o.Offset(11, 1)) & "))", "#,##0.0"
    Hd o.Offset(23, 0).Resize(1, 2)
    o.Resize(25, 3).Columns.AutoFit
End Sub
