Attribute VB_Name = "modLtxRibbon"
' =====================================================================
'  Lattice for PowerPoint - ribbon callbacks
'  Every button in customUI14.xml calls LtxRibbon_OnAction and is routed
'  here by its control id, so the XML never needs to know macro names.
' =====================================================================
Option Explicit

Public gRibbon As IRibbonUI

' No-op used by the build to force this module to compile
Public Sub LtxRibbon_Ping()
End Sub

Public Sub LtxRibbon_OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

' Bring the Lattice tab to the front (used by the screenshot tests; handy from the QAT too)
Public Sub LtxShowTab()
    On Error Resume Next
    gRibbon.ActivateTab "ltxTab"
End Sub

Public Sub LtxRibbonInvalidate()
    On Error Resume Next
    If Not gRibbon Is Nothing Then gRibbon.Invalidate
End Sub

Public Sub LtxRibbon_GetPressed(control As IRibbonControl, ByRef pressed)
    Select Case control.Id
        Case "ltxAirplane":   pressed = LtxBool("AirplaneMode")
        Case "ltxBlackFonts": pressed = LtxBool("BlackFonts", True)
        Case "ltxHideGrid":   pressed = LtxBool("HideGridlines", True)
        Case "ltxChartNoBorder": pressed = LtxBool("ChartNoBorder", True)
        Case "ltxCheckErrors": pressed = LtxBool("CheckErrors", True)
        Case Else:            pressed = False
    End Select
End Sub

Public Sub LtxRibbon_OnToggle(control As IRibbonControl, pressed As Boolean)
    Select Case control.Id
        Case "ltxAirplane":   LtxSet "AirplaneMode", IIf(pressed, "1", "0")
        Case "ltxBlackFonts": LtxSet "BlackFonts", IIf(pressed, "1", "0")
        Case "ltxHideGrid":   LtxSet "HideGridlines", IIf(pressed, "1", "0")
        Case "ltxChartNoBorder": LtxSet "ChartNoBorder", IIf(pressed, "1", "0")
        Case "ltxCheckErrors": LtxSet "CheckErrors", IIf(pressed, "1", "0")
    End Select
End Sub

Public Sub LtxRibbon_GetLabel(control As IRibbonControl, ByRef label)
    Select Case control.Id
        Case "ltxAbout": label = "Lattice " & LTX_VERSION
        Case Else: label = control.Id
    End Select
End Sub

Public Sub LtxRibbon_OnAction(control As IRibbonControl)
    On Error GoTo fail
    Select Case control.Id
        ' ---- Presentation ----
        Case "ltxNew":          LtxNewFromTemplate
        Case "ltxBackup":       LtxSaveBackup
        Case "ltxPdf":          LtxExportPdf False
        Case "ltxPdfSel":       LtxExportPdf True
        Case "ltxSend":         LtxSendMail False
        Case "ltxSendPdf":      LtxSendMail True
        Case "ltxDupSlides":    LtxDuplicateSlides
        Case "ltxDelHidden":    LtxDeleteHiddenSlides
        Case "ltxClearNotes":   LtxRemoveNotes
        Case "ltxNumbersOn":    LtxSlideNumbers True
        Case "ltxNumbersOff":   LtxSlideNumbers False
        Case "ltxAgenda":       LtxAgenda
        Case "ltxDeckCheck":    LtxDeckCheck
        Case "ltxRemoveStamps": LtxRemoveStamps
        ' ---- Insert ----
        Case "ltxLibInsert":    LtxLibraryInsert
        Case "ltxLibSave":      LtxLibrarySave
        Case "ltxMasterInsert": LtxMasterInsert
        Case "ltxMasterSave":   LtxMasterSave
        Case "ltxMasterManage": LtxMasterManage
        Case "ltxStGreen":      LtxSticker "light_green"
        Case "ltxStAmber":      LtxSticker "light_amber"
        Case "ltxStRed":        LtxSticker "light_red"
        Case "ltxStLights":     LtxSticker "lights"
        Case "ltxStTick":       LtxSticker "tick"
        Case "ltxStCross":      LtxSticker "cross"
        Case "ltxStHarvey":     LtxSticker "harvey_cycle"
        Case "ltxStDraft":      LtxSticker "stamp_draft"
        Case "ltxStConf":       LtxSticker "stamp_confidential"
        Case "ltxStPrelim":     LtxSticker "stamp_preliminary"
        Case "ltxStDisc":       LtxSticker "stamp_discussion"
        Case "ltxStIllus":      LtxSticker "stamp_illustrative"
        Case "ltxStDraftAll":   LtxSticker "stamp_draft_all"
        Case "ltxStRemove":     LtxSticker "stamps_remove"
        Case "ltxStickyNote":   LtxTurbo "sticky"
        ' ---- Colour ----
        Case "ltxFontCycle":    LtxCycleFontColor
        Case "ltxFillCycle":    LtxCycleFill
        Case "ltxLineCycle":    LtxCycleLine
        Case "ltxHexFill":      LtxHexColor "fill"
        Case "ltxHexLine":      LtxHexColor "line"
        Case "ltxHexFont":      LtxHexColor "font"
        Case "ltxCopyFormat":   LtxCopyFormat
        Case "ltxPasteFormat":  LtxPasteFormat
        ' ---- Text ----
        Case "ltxGrow":         LtxText "grow"
        Case "ltxShrink":       LtxText "shrink"
        Case "ltxBullets":      LtxText "bullets"
        Case "ltxUpper":        LtxText "upper"
        Case "ltxLower":        LtxText "lower"
        Case "ltxTitleCase":    LtxText "title"
        Case "ltxSentence":     LtxText "sentence"
        Case "ltxSuper":        LtxText "super"
        Case "ltxSub":          LtxText "sub"
        Case "ltxClearFmt":     LtxText "clear"
        Case "ltxSpaceMore":    LtxText "spacemore"
        Case "ltxSpaceLess":    LtxText "spaceless"
        Case "ltxStripSpaces":  LtxText "stripspaces"
        Case "ltxAutofitNone":  LtxAutofit 0
        Case "ltxAutofitShape": LtxAutofit 1
        Case "ltxAutofitText":  LtxAutofit 2
        Case "ltxMargins0":     LtxMargins 0
        Case "ltxMargins4":     LtxMargins 3.6
        Case "ltxWordWrap":     LtxWordWrapToggle
        Case "ltxSelType":      LtxSelectMatch "type"
        Case "ltxSelFill":      LtxSelectMatch "fill"
        Case "ltxSelLine":      LtxSelectMatch "line"
        Case "ltxSelFont":      LtxSelectMatch "font"
        Case "ltxSelFontColor": LtxSelectMatch "fontcolor"
        Case "ltxSelSize":      LtxSelectMatch "size"
        Case "ltxSelLinked":    LtxSelectMatch "linked"
        Case "ltxSelText":      LtxSelectMatch "text"
        Case "ltxSelPictures":  LtxSelectMatch "pictures"
        Case "ltxArrangeGrid":  LtxArrangeGrid
        Case "ltxAutoAlign":    LtxAutoAlign
        Case "ltxContainer":    LtxArrangeInContainer
        Case "ltxStackV":       LtxStack "v"
        Case "ltxStackH":       LtxStack "h"
        Case "ltxBodyLeft":     LtxAlignToBody "left"
        Case "ltxBodyRight":    LtxAlignToBody "right"
        Case "ltxBodyTop":      LtxAlignToBody "top"
        Case "ltxBodyBottom":   LtxAlignToBody "bottom"
        Case "ltxBodyCenter":   LtxAlignToBody "center"
        Case "ltxBodyMiddle":   LtxAlignToBody "middle"
        Case "ltxBodyFill":     LtxAlignToBody "fill"
        Case "ltxBodyWidth":    LtxAlignToBody "width"
        Case "ltxSlideLinksRefresh": LtxRefreshSlideLinks
        Case "ltxSlideLinkUpdate": LtxUpdateSlideLink
        ' ---- Shapes ----
        Case "ltxAlignLeft":    LtxAlign "left"
        Case "ltxAlignCenter":  LtxAlign "center"
        Case "ltxAlignRight":   LtxAlign "right"
        Case "ltxAlignTop":     LtxAlign "top"
        Case "ltxAlignMiddle":  LtxAlign "middle"
        Case "ltxAlignBottom":  LtxAlign "bottom"
        Case "ltxDistH":        LtxDistribute "h"
        Case "ltxDistV":        LtxDistribute "v"
        Case "ltxGapH":         LtxSetGap "h"
        Case "ltxGapV":         LtxSetGap "v"
        Case "ltxCenterSlide":  LtxCenterOnSlide "both"
        Case "ltxCenterH":      LtxCenterOnSlide "h"
        Case "ltxCenterV":      LtxCenterOnSlide "v"
        Case "ltxSameW":        LtxMatchSize "width"
        Case "ltxSameH":        LtxMatchSize "height"
        Case "ltxSameSize":     LtxMatchSize "both"
        Case "ltxStretchL":     LtxStretch "left"
        Case "ltxStretchR":     LtxStretch "right"
        Case "ltxStretchT":     LtxStretch "top"
        Case "ltxStretchB":     LtxStretch "bottom"
        Case "ltxSquare":       LtxMakeSquare
        Case "ltxFitSlide":     LtxFitToSlide
        Case "ltxSwap":         LtxSwap
        Case "ltxCopyGeom":     LtxCopyGeometry
        Case "ltxRememberPos":  LtxRememberPosition
        Case "ltxApplyPos":     LtxApplyPosition True
        Case "ltxApplyPosOnly": LtxApplyPosition False
        Case "ltxFront":        LtxZ 0
        Case "ltxBack":         LtxZ 1
        Case "ltxForward":      LtxZ 2
        Case "ltxBackward":     LtxZ 3
        Case "ltxGroup":        LtxGroupToggle
        Case "ltxToPicture":    LtxConvertToPicture "emf"
        Case "ltxToPng":        LtxConvertToPicture "png"
        Case "ltxCropSlide":    LtxCropToSlide
        Case "ltxNudgeL":       LtxNudge -1, 0
        Case "ltxNudgeR":       LtxNudge 1, 0
        Case "ltxNudgeU":       LtxNudge 0, -1
        Case "ltxNudgeD":       LtxNudge 0, 1
        ' ---- Import / links ----
        Case "ltxImportPic", "ltxImportPic2": LtxImport "picture"
        Case "ltxImportTable":  LtxImport "table"
        Case "ltxImportChart":  LtxImport "chart"
        Case "ltxImportNativeChart": LtxImport "nativechart"
        Case "ltxImportEmbed":  LtxImport "embed"
        Case "ltxRefreshSel":   LtxRefreshSelected
        Case "ltxRefreshSlide": LtxRefreshSlide
        Case "ltxRefreshAll":   LtxRefreshAll
        Case "ltxSizeWidth":    LtxSetSizeMode "width"
        Case "ltxSizeHeight":   LtxSetSizeMode "height"
        Case "ltxSizeKeep":     LtxSetSizeMode "keep"
        Case "ltxSizeSource":   LtxSetSizeMode "source"
        Case "ltxGoSource":     LtxGoToSource
        Case "ltxChangeSource": LtxChangeSource
        Case "ltxBreakLink":    LtxBreakLink
        Case "ltxOpenSources":  LtxOpenAllSources
        Case "ltxExplorer":     frmLtxExplorer.Show vbModeless
        Case "ltxHighlight":    LtxHighlightLinks
        Case "ltxKeytips":      LtxKeytipCard
        ' ---- Text links ----
        Case "ltxLinkTextWhole": LtxLinkTextWhole
        Case "ltxLinkTextSel":   LtxLinkTextSelection
        Case "ltxTextLinkInfo":  LtxTextLinkInfo
        Case "ltxTextLinkEdit":  LtxTextLinkRetemplate
        ' ---- Styles ----
        Case "ltxStyleSave":     LtxStyleSave
        Case "ltxStyleApply":    LtxStyleApply
        Case "ltxStyleSimilar":  LtxStyleApplyToSimilar
        Case "ltxStyleUpdate":   LtxStyleUpdateAll
        Case "ltxStyleManage":   LtxStyleManage
        Case "ltxStyleClear":    LtxStyleClearTag
        ' ---- Deck tools ----
        Case "ltxHarmonizeFonts": LtxHarmonizeFonts
        Case "ltxFontReport":    LtxFontReport
        Case "ltxReplaceFont":   LtxReplaceFontDeck
        Case "ltxAlignTitles":   LtxAlignTitles
        Case "ltxConformTitles": LtxConformRole "title"
        Case "ltxConformSubs":   LtxConformRole "subtitle"
        Case "ltxConformBody":   LtxConformRole "body"
        Case "ltxConformLike":   LtxConformLikeSelected
        Case "ltxFontCheck":     LtxFontConsistencyCheck
        Case "ltxCtxRefresh", "ltxCtxRefreshPic", "ltxCtxRefreshTbl", "ltxRefreshSelOnly": LtxRefreshSelected
        Case "ltxCtxGoSourcePic": LtxGoToSource
        Case "ltxCtxGoSource":   LtxGoToSource
        Case "ltxCtxBreak":      LtxBreakLink
        Case "ltxCtxConformLike": LtxConformLikeSelected
        Case "ltxTblNumbers":    LtxTableNumberFormat
        Case "ltxTblDistribute": LtxTableDistribute
        Case "ltxTblFontSize":   LtxTableFontSize
        Case "ltxTblHeader":     LtxTableHeaderStyle
        Case "ltxChartPalette":  LtxChartPalette
        Case "ltxSetChartPalette": LtxSetChartPalette
        Case "ltxDeleteComments": LtxDeleteComments
        Case "ltxCollectNotes":  LtxCollectNotes
        Case "ltxSectionsPdf":   LtxExportSectionsPdf
        ' ---- Pagination ----
        Case "ltxPageNums":      LtxPageNumbers
        Case "ltxPageNumsOff":   LtxPageNumbersRemove
        Case "ltxPageSkip":      LtxPageSkipToggle
        Case "ltxFirstNumber":   LtxFirstSlideNumber
        Case "ltxBlankEvery":    LtxBlankPages "every"
        Case "ltxBlankSections": LtxBlankPages "sections"
        Case "ltxBlankOdd":      LtxBlankPages "odd"
        Case "ltxBlankRemove":   LtxBlankPagesRemove
        Case "ltxSyncSections":  LtxSyncSectionTitles
        Case "ltxSectionFromTitle": LtxSectionFromTitle
        Case "ltxToc":           LtxTableOfContents
        Case "ltxSetTableFont":  LtxSet "TableFontSize", InputBox("Font size for imported tables (0 = keep Excel's):", LTX_APP, LtxGet("TableFontSize", "10"))
        ' ---- Settings ----
        Case "ltxSettingsDlg":  LtxOpenSettings
        Case "ltxFmtEmf":       LtxSet "PictureFormat", "emf"
        Case "ltxFmtPng":       LtxSet "PictureFormat", "png"
        Case "ltxSetTemplate":  LtxSet "Template", "": LtxNewFromTemplate
        Case "ltxSetMargin":    LtxSet "MarginPt", InputBox("Slide margin for imports (points):", LTX_APP, LtxGet("MarginPt", "36"))
        Case "ltxSetGapDef":    LtxSet "Gap", InputBox("Default gap for distribute/grid (points):", LTX_APP, LtxGet("Gap", "8"))
        Case "ltxSetCheckWords": LtxSet "CheckWords", InputBox("Deck Check words (comma separated):", LTX_APP, LtxGet("CheckWords", "TBD,TBU,xx,XXX,[placeholder],lorem,DRAFT,TODO"))
        Case "ltxOpenLog":      Shell "notepad.exe """ & LtxDataDir() & "\lattice.log""", vbNormalFocus
        Case "ltxAbout":        MsgBox "Lattice " & LTX_VERSION & vbLf & "Excel/PowerPoint/Word productivity add-in." & vbLf & vbLf & "Data folder: " & LtxDataDir(), vbInformation, LTX_APP
        Case Else
            If Left$(control.Id, 8) = "ltxTurbo" Then
                LtxTurbo LCase$(Mid$(control.Id, 9))
            ElseIf Left$(control.Id, 8) = "ltxTheme" Then
                ' ltxThemeFill5 -> fill, theme colour 5
                Dim t As String, n As Long
                t = Mid$(control.Id, 9)
                n = Val(Right$(t, IIf(IsNumeric(Right$(t, 2)), 2, 1)))
                LtxThemeColor LCase$(Left$(t, Len(t) - Len(CStr(n)))), n
            Else
                MsgBox "No handler for " & control.Id, vbExclamation, LTX_APP
            End If
    End Select
    Exit Sub
fail:
    MsgBox "Lattice: " & Err.Description, vbExclamation, LTX_APP
End Sub
