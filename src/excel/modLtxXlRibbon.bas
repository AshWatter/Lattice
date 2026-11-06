Attribute VB_Name = "modLtxXlRibbon"
' =====================================================================
'  Lattice for Excel - ribbon callbacks (one OnAction routed by control id)
' =====================================================================
Option Explicit

Public gRibbon As IRibbonUI

Public Sub LtxXl_OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Public Sub LtxShowTab()
    On Error Resume Next
    gRibbon.ActivateTab "xlTab"
End Sub

Public Sub LtxXl_GetLabel(control As IRibbonControl, ByRef label)
    label = "About Lattice"
End Sub

Public Sub LtxXl_GetPressed(control As IRibbonControl, ByRef pressed)
    Select Case control.Id
        Case "xlMatchWidth":  pressed = (LtxExportSizeMode() = "width")
        Case "xlMatchHeight": pressed = (LtxExportSizeMode() = "height")
        Case "xlMatchSize":   pressed = (LtxExportSizeMode() = "size")
        Case "xlBlackToggle": pressed = LtxBlackIsOn()
        Case Else:            pressed = False
    End Select
End Sub

Public Sub LtxXl_OnToggle(control As IRibbonControl, pressed As Boolean)
    Dim m As String
    If control.Id = "xlBlackToggle" Then LtxBlackToggle: Exit Sub
    Select Case control.Id
        Case "xlMatchWidth":  m = "width"
        Case "xlMatchHeight": m = "height"
        Case "xlMatchSize":   m = "size"
    End Select
    If pressed Then LtxSetExportSizeMode m Else LtxSetExportSizeMode "none"
    On Error Resume Next
    gRibbon.Invalidate
End Sub

Public Sub LtxXl_OnAction(control As IRibbonControl)
    On Error GoTo fail
    Select Case control.Id
        ' Workbook
        Case "xlBackup":         LtxSaveBackup
        Case "xlOptimize":       LtxOptimizeWorkbook
        Case "xlHardcodeSheet":  LtxHardcodeWorkbook True
        Case "xlHardcodeBook":   LtxHardcodeWorkbook False
        Case "xlExportPng":      LtxExportPictureFile
        Case "xlHealth":         LtxHealthReport
        Case "xlHealthRemove":   LtxHealthReportRemove
        ' Color
        Case "xlAutoColor", "xlFmtAuto": LtxAutoColor
        Case "xlAutoColorSheet": LtxAutoColorSheet
        Case "xlAutoColorBook":  LtxAutoColorWorkbook
        Case "xlBlackBook":      LtxBlackToggleWorkbook
        Case "xlBlackSel":       LtxBlackSelection
        Case "xlFontCycle", "xlFmtFont": LtxCycleFontColor
        Case "xlFillCycle", "xlFmtFill": LtxCycleFill
        Case "xlBorderCycle", "xlFmtBorderCycle": LtxCycleBorder
        ' Format > Number
        Case "xlFmtNum":         LtxCycleNumber
        Case "xlFmtPct":         LtxCyclePercent
        Case "xlFmtMult":        LtxCycleMultiple
        Case "xlFmtDate":        LtxCycleDate
        Case "xlFmtCur":         LtxCycleCurrency
        Case "xlFmtNeg":         LtxNegativeStyleCycle
        ' Format > Border
        Case "xlFmtBottom":      LtxBorderBottomCycle
        Case "xlFmtTop":         LtxBorderTopCycle
        Case "xlFmtOutline":     LtxBorderOutline
        Case "xlFmtNoBorder":    LtxBorderRemove
        ' Format > Alignment
        Case "xlFmtCenterH":     LtxCenterHCycle
        Case "xlFmtAlignCycle":  LtxCycleAlign
        Case "xlFmtVAlign":      LtxVAlignCycle
        Case "xlFmtIndentMore":  LtxIndentMore
        Case "xlFmtIndentLess":  LtxIndentLess
        Case "xlFmtWrap":        LtxToggleWrap
        ' Format > Font / Case / Styles
        Case "xlFmtGrow":        LtxFontGrow
        Case "xlFmtShrink":      LtxFontShrink
        Case "xlFmtUnderline":   LtxUnderlineCycle
        Case "xlFmtUpper":       LtxCaseUpper
        Case "xlFmtLower":       LtxCaseLower
        Case "xlFmtProper":      LtxCaseProper
        Case "xlHeader":         LtxHeaderStyle
        Case "xlTotal":          LtxTotalStyle
        Case "xlSumRow":         LtxSumRow
        Case "xlCopyFmt":        LtxCopyFormat
        Case "xlPasteFmt":       LtxPasteFormat
        ' Formulas
        Case "xlAnchors":        LtxCycleAnchors
        Case "xlCaseSwitch":     LtxInsertCaseSwitch
        Case "xlFlip":           LtxFlipSign
        Case "xlValues":         LtxToValues
        Case "xlIfError":        LtxWrapIfError
        Case "xlSimplify":       LtxSimplify
        Case "xlCheck":          LtxInsertCheck
        Case "xlFillRight":      LtxFastFillRight
        Case "xlFillDown":       LtxFastFillDown
        Case "xlCopyExact":      LtxCopyExact
        Case "xlPasteExact":     LtxPasteExact
        Case "xlName":           LtxNameRange
        Case "xlTag":            LtxTagCell
        ' Charts
        Case "xlFootball":       LtxFootballField
        Case "xlFFLine":         LtxFootballFieldLine
        Case "xlStacked":        LtxStackedColumn
        Case "xlWaterfall":      LtxWaterfall
        Case "xlChartCol":       LtxQuickChart 51
        Case "xlChartLine":      LtxQuickChart 4
        Case "xlChartBar":       LtxQuickChart 57
        Case "xlCAGR":           LtxInsertCAGR
        Case "xlGrowthRow":      LtxInsertGrowthRow
        Case "xlMarginRow":      LtxInsertMarginRow
        Case "xlChartTotals":    LtxChartTotals
        Case "xlChartCAGR":      LtxChartCAGR
        Case "xlChartDelta":     LtxChartDelta
        Case "xlChartNotesRemove": LtxChartNotesRemove
        ' Library
        Case "xlLibLBO":         LtxInsertLBO
        Case "xlLibDCF":         LtxInsertDCF
        Case "xlLibDist":        LtxInsertDistributionWaterfall
        Case "xlLibAccDil":      LtxInsertAccretionDilution
        Case "xlLibReturns":     LtxInsertReturns
        Case "xlLibDebt":        LtxInsertDebtSchedule
        Case "xlLibCredit":      LtxInsertCreditStats
        Case "xlLibBS":          LtxInsertBlackScholes
        Case "xlLibSens":        LtxInsertSensitivity
        Case "xlLibScenario":    LtxInsertScenarioBlock
        Case "xlLibTimeline":    LtxInsertTimeline
        Case "xlRunCases":       LtxRunCases
        Case "xlLibInsert":      LtxLibInsert
        Case "xlLibSave":        LtxLibSave
        Case "xlLibOpen":        LtxLibOpen
        Case "xlLibPath":        LtxLibSetPath
        ' Cells / Sheets / View
        Case "xlCleanNumbers":   LtxCleanNumbers
        Case "xlTrimText":       LtxTrimText
        Case "xlUnmergeFill":    LtxUnmergeFill
        Case "xlBlankRows":      LtxRemoveBlankRows
        Case "xlToThousands":    LtxToThousands
        Case "xlToMillions":     LtxToMillions
        Case "xlFromThousands":  LtxFromThousands
        Case "xlTransposeVals":  LtxTransposeValues
        Case "xlSelStats":       LtxSelectionStats
        Case "xlUnhide":         LtxUnhideAll
        Case "xlHideInactive":   LtxHideInactiveSheets
        Case "xlCompare":        LtxCompareSheets
        Case "xlGridlines":      LtxGridlinesToggle
        Case "xlHeadings":       LtxToggleHeadings
        Case "xlFreeze":         LtxFreezeToggle
        Case "xlShowFormulas":   LtxToggleFormulas
        Case "xlZoom":           LtxZoomCycle
        Case "xlColWidth":       LtxColWidth
        Case "xlAutofit":        LtxAutofitCols
        ' Audit
        Case "xlModelCheck":     LtxModelCheckPane
        Case "xlAuditSheet":     LtxAuditHardcodes False
        Case "xlAuditBook":      LtxAuditHardcodes True
        Case "xlHcWhitelist":    XlSet "HardcodeWhitelist", InputBox("Numbers allowed inside formulas (separate with |):", "Lattice", XlGet("HardcodeWhitelist", "0|1|2|100|12|365|360|1000|1000000|4|52|30|31|0.5"))
        Case "xlTraceIn", "xlTraceInBig": LtxTraceIn
        Case "xlTraceOut":       LtxTraceOut
        Case "xlGoPrec":         LtxGoPrecedent
        Case "xlGoDep":          LtxGoDependent
        Case "xlGoBack":         LtxGoBack
        Case "xlPrec":           LtxPrecedents
        Case "xlDep":            LtxDependents
        Case "xlFlow":           LtxFormulaFlow
        Case "xlArrows":         LtxShowArrows
        Case "xlClearArrows":    LtxClearArrows
        Case "xlHlHardcodes":    LtxHighlightHardcodes
        Case "xlInconsistent":   LtxInconsistentFormulas
        Case "xlErrors":         LtxSelectErrors
        Case "xlHardcodes":      LtxSelectHardcodes
        Case "xlExtractHc":      LtxExtractHardcode
        ' Comps
        Case "xlQuickStats":     LtxQuickStats
        Case "xlFormatComps":    LtxFormatComps
        Case "xlRankPct":        LtxRankPercentile
        Case "xlOutliers":       LtxFlagOutliers
        ' Export
        Case "xlSendPic", "xlSendPic2": LtxSendToPpt "picture"
        Case "xlSendTable":      LtxSendToPpt "table"
        Case "xlSendChart":      LtxSendToPpt "chart"
        Case "xlSendNativeChart": LtxSendToPpt "nativechart"
        Case "xlSendEmbed":      LtxSendToPpt "embed"
        Case "xlSendWordPic", "xlSendWordPic2": LtxSendToWord "picture"
        Case "xlSendWordTable":  LtxSendToWord "table"
        Case "xlCopyPic":        LtxCopyForPpt "picture"
        Case "xlCopyTable":      LtxCopyForPpt "table"
        Case "xlCopyChart":      LtxCopyForPpt "chart"
        Case "xlRefreshPpt":     LtxRefreshPptFromHere
        Case "xlRefreshWord":    LtxRefreshWordFromHere
        ' Lattice
        Case "xlShortcutCard":   LtxShortcutCard
        Case "xlKeysEdit":       LtxKeysEdit
        Case "xlKeysReload":     LtxKeysLoad
        Case "xlKeysReset":      LtxKeysReset
        Case "xlSettingsDlg":    LtxOpenSettings
        Case "xlStyleFile", "xlStyleFile2": LtxEditStyleFile
        Case "xlAutoNameToggle": XlSet "AutoName", IIf(XlGet("AutoName", "1") = "1", "0", "1"): MsgBox "Auto-name exported ranges: " & IIf(XlGet("AutoName") = "1", "on", "off"), vbInformation, "Lattice"
        Case "xlExportActivateToggle": XlSet "ExportActivate", IIf(XlGet("ExportActivate", "1") = "1", "0", "1"): MsgBox "Switch to PowerPoint/Word after export: " & IIf(XlGet("ExportActivate") = "1", "on", "off"), vbInformation, "Lattice"
        Case "xlAbout":          MsgBox "Lattice for Excel" & vbLf & "Formatting, modelling, library, auditing, charts, comps, shortcuts and export to PowerPoint / Word." & vbLf & vbLf & "Files: " & Environ$("APPDATA") & "\Lattice", vbInformation, "Lattice"
        Case Else:               MsgBox "No handler for " & control.Id, vbExclamation, "Lattice"
    End Select
    Exit Sub
fail:
    MsgBox "Lattice: " & Err.Description, vbExclamation, "Lattice"
End Sub
