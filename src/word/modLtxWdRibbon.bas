Attribute VB_Name = "modLtxWdRibbon"
' =====================================================================
'  Lattice for Word - ribbon callbacks
' =====================================================================
Option Explicit

Public gRibbon As IRibbonUI

' No-op used by the build to force this module to compile
Public Sub LtxWd_Ping()
End Sub

Public Sub LtxWd_OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Public Sub LtxWd_GetPressed(control As IRibbonControl, ByRef pressed)
    Select Case control.Id
        Case "wdBlackFonts": pressed = LtxBool("BlackFonts", True)
        Case "wdAirplane":   pressed = LtxBool("AirplaneMode")
        Case Else:           pressed = False
    End Select
End Sub

Public Sub LtxWd_OnToggle(control As IRibbonControl, pressed As Boolean)
    Select Case control.Id
        Case "wdBlackFonts": LtxSet "BlackFonts", IIf(pressed, "1", "0")
        Case "wdAirplane":   LtxSet "AirplaneMode", IIf(pressed, "1", "0")
    End Select
End Sub

Public Sub LtxWd_OnAction(control As IRibbonControl)
    On Error GoTo fail
    Select Case control.Id
        Case "wdImportPic":    WdImport "picture"
        Case "wdImportTable":  WdImport "table"
        Case "wdImportChart":  WdImport "chart"
        Case "wdRefreshSel":   WdRefreshSelected
        Case "wdRefreshAll":   WdRefreshAll
        Case "wdGoSource":     WdGoToSource
        Case "wdBreak":        WdBreakLink
        Case "wdList":         WdListLinks
        Case "wdStrip":        WdStripDoubleSpaces
        Case "wdTablesFit":    WdTablesAutofit
        Case "wdPicsWidth":    WdPicturesToWidth
        Case "wdKeytips":      WdKeytipCard
        Case "wdAbout":        MsgBox "Lattice for Word " & LTX_VERSION, vbInformation, LTX_APP
        Case Else:             MsgBox "No handler for " & control.Id, vbExclamation, LTX_APP
    End Select
    Exit Sub
fail:
    MsgBox "Lattice: " & Err.Description, vbExclamation, LTX_APP
End Sub
