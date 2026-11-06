Attribute VB_Name = "modLtxTheme"
' =====================================================================
'  Lattice design system for windows (Excel copy; src\ppt\modLtxTheme.bas is the same code)
'  Navy title band with an azure accent line, white body, Consolas for formulas, flat buttons
'  (primary navy, secondary sky, Close / Cancel ghost), slate footer behind a hairline.
'  Palette: navy 31,56,100  azure 46,117,182  sky 222,229,240  mist 243,245,249
'           ink 31,41,55    slate 107,114,128
' =====================================================================
Option Explicit

Public Sub LtxThemeForm(frm As Object, titleText As String)
    Dim ctl As Object, band As Object, accent As Object, hair As Object, w As Single, primary As Boolean, fresh As Boolean
    On Error Resume Next
    frm.BackColor = RGB(255, 255, 255)
    frm.Font.Name = "Segoe UI"
    w = frm.InsideWidth
    Set band = frm.Controls("lblBand")
    fresh = (band Is Nothing)
    If fresh Then
        Set band = frm.Controls.Add("Forms.Label.1", "lblBand", True)
        Set accent = frm.Controls.Add("Forms.Label.1", "lblAccent", True)
        Set hair = frm.Controls.Add("Forms.Label.1", "lblHair", True)
        For Each ctl In frm.Controls
            If ctl.Name <> "lblBand" And ctl.Name <> "lblAccent" And ctl.Name <> "lblHair" And ctl.Name <> "lblTitle" Then
                If ctl.Top < 44 Then ctl.Top = ctl.Top + 12: ctl.Height = ctl.Height - 12
            End If
        Next
        band.ZOrder 1
    Else
        Set accent = frm.Controls("lblAccent")
        Set hair = frm.Controls("lblHair")
    End If
    band.Left = 0: band.Top = 0: band.Width = w: band.Height = 40
    band.BackColor = RGB(31, 56, 100): band.Caption = ""
    accent.Left = 0: accent.Top = 40: accent.Width = w: accent.Height = 3
    accent.BackColor = RGB(46, 117, 182): accent.Caption = ""
    With frm.Controls("lblTitle")
        .Left = 14: .Top = 11: .Width = w - 28: .Height = 20
        .BackStyle = 0: .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI": .Font.Size = 12: .Font.Bold = True
        .Caption = titleText
        .ZOrder 0
    End With
    For Each ctl In frm.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.Font.Name = "Consolas": ctl.Font.Size = 9: ctl.BackColor = RGB(243, 245, 249)
                ctl.BorderStyle = 1: ctl.BorderColor = RGB(222, 229, 240): ctl.SpecialEffect = 0: ctl.ForeColor = RGB(31, 56, 100)
            Case "ListBox"
                ctl.Font.Name = "Consolas": ctl.Font.Size = 9
                ctl.BorderStyle = 1: ctl.BorderColor = RGB(222, 229, 240): ctl.SpecialEffect = 0: ctl.ForeColor = RGB(31, 41, 55)
            Case "ComboBox"
                ctl.Font.Name = "Segoe UI": ctl.Font.Size = 9
                ctl.BorderStyle = 1: ctl.BorderColor = RGB(222, 229, 240): ctl.SpecialEffect = 0: ctl.ForeColor = RGB(31, 41, 55)
            Case "CheckBox", "OptionButton"
                ctl.Font.Name = "Segoe UI": ctl.Font.Size = 9: ctl.ForeColor = RGB(31, 41, 55): ctl.BackStyle = 0
            Case "Label"
                If ctl.Name <> "lblBand" And ctl.Name <> "lblAccent" And ctl.Name <> "lblHair" Then ctl.BackStyle = 0
                If Left$(ctl.Name, 4) = "lbl_" Then ctl.Font.Name = "Segoe UI": ctl.Font.Size = 9: ctl.ForeColor = RGB(31, 41, 55)
            Case "CommandButton"
                ctl.Font.Name = "Segoe UI": ctl.Font.Size = 8: ctl.Font.Bold = True: ctl.Height = 26
                primary = (ctl.Name = "btnGo" Or ctl.Name = "btnRefresh" Or ctl.Name = "btnRefreshAll" Or ctl.Name = "btnSheet" _
                           Or ctl.Name = "btnBook" Or ctl.Name = "btnSave" Or ctl.Name = "btnEdit")
                If ctl.Name = "btnClose" Or ctl.Name = "btnCancel" Then
                    ctl.BackColor = RGB(243, 245, 249): ctl.ForeColor = RGB(107, 114, 128)
                ElseIf primary Then
                    ctl.BackColor = RGB(31, 56, 100): ctl.ForeColor = RGB(255, 255, 255)
                Else
                    ctl.BackColor = RGB(222, 229, 240): ctl.ForeColor = RGB(31, 56, 100)
                End If
        End Select
    Next
    With frm.Controls("lblInfo")
        .Font.Name = "Segoe UI": .Font.Size = 8: .ForeColor = RGB(107, 114, 128)
        hair.Left = .Left: hair.Width = .Width: hair.Top = .Top - 5: hair.Height = 1
        hair.BackColor = RGB(222, 229, 240): hair.Caption = ""
    End With
End Sub

' Settings dialog entry points (modeless, so the ribbon stays usable while it is open)
Public Sub LtxOpenSettings()
    frmLtxSettings.Show vbModeless
End Sub

Public Sub LtxCloseSettings()
    On Error Resume Next
    Unload frmLtxSettings
    Unload frmLtxExplorer
End Sub
