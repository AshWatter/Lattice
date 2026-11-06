Attribute VB_Name = "modLtxCards"
' =====================================================================
'  Lattice for PowerPoint - keytip card (Alt, L, letter)
' =====================================================================
Option Explicit

Public Sub LtxKeytipCard()
    Dim r As String
    r = LtxCardGroup("Presentation") & _
        LtxCardRow("New deck from template", "Alt L N") & LtxCardRow("Agenda (sections, dividers, TOC)", "Alt L G") & _
        LtxCardRow("Pagination", "Alt L P") & LtxCardRow("Deck Check", "Alt L K") & LtxCardRow("Deck Tools (fonts, tables, charts, notes)", "Alt L D") & _
        LtxCardGroup("Insert") & _
        LtxCardRow("Slide Library", "Alt L B") & LtxCardRow("Turbo Shapes", "Alt L U") & LtxCardRow("Master Shapes", "Alt L M") & _
        LtxCardGroup("Color and text") & _
        LtxCardRow("Color Tools", "Alt L C") & LtxCardRow("Styles", "Alt L Y") & LtxCardRow("Select Match", "Alt L S") & LtxCardRow("Arrange Shapes", "Alt L W") & _
        LtxCardGroup("Shapes") & _
        LtxCardRow("Align left / centre / right", "Alt L 1 / 2 / 3") & LtxCardRow("Align top / middle / bottom", "Alt L 4 / 5 / 6") & _
        LtxCardRow("Distribute horizontally / vertically", "Alt L 7 / 8") & LtxCardRow("Same width / height / size", "Alt L 9 / 0 / Q") & _
        LtxCardGroup("Import and links") & _
        LtxCardRow("Import Excel selection (black fonts)", "Alt L E") & LtxCardRow("Link Text", "Alt L L") & _
        LtxCardRow("Refresh selected links", "Alt L R") & LtxCardRow("Explorer Pane", "Alt L X") & LtxCardRow("Airplane Mode", "Alt L A") & _
        LtxCardRow("Highlight links (slide sorter)", "Alt L H") & LtxCardRow("Settings", "Alt L O") & _
        LtxCardGroup("From Excel") & _
        LtxCardRow("Export selection to PowerPoint (picture)", "Ctrl Alt R") & LtxCardRow("Export as table / chart", "Ctrl Alt T / C") & _
        LtxCardRow("Export to Word", "Ctrl Alt W") & LtxCardRow("Refresh the open deck from Excel", "Ctrl Shift U")
    LtxOpenHtml LtxDataDir() & "\lattice_powerpoint_keytips.html", LtxCardHtml("PowerPoint", r)
End Sub
