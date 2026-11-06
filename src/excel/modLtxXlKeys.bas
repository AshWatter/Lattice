Attribute VB_Name = "modLtxXlKeys"
' =====================================================================
'  Lattice for Excel - keyboard shortcuts
'  Shortcuts live in %APPDATA%\Lattice\excel_shortcuts.txt as
'  "key=macro" lines using Excel OnKey syntax (^ Ctrl, + Shift, % Alt).
'  Loaded when the add-in opens; "Reload shortcuts" after editing.
'  Every Lattice macro (any Public Sub without arguments) can be bound.
' =====================================================================
Option Explicit

Private mBound As Collection

Public Function KeysFilePath() As String
    KeysFilePath = Environ$("APPDATA") & "\Lattice\excel_shortcuts.txt"
End Function

Private Sub WriteDefaults()
    Dim f As Integer
    If Dir(Environ$("APPDATA") & "\Lattice", vbDirectory) = "" Then MkDir Environ$("APPDATA") & "\Lattice"
    f = FreeFile
    Open KeysFilePath() For Output As #f
    Print #f, "# Lattice for Excel shortcuts.  key=macro   (^ = Ctrl, + = Shift, % = Alt, {F1}..{F12}, letters lower case)"
    Print #f, "# Delete a line to free the key. Lines starting with # are ignored. Reload with Lattice > Settings > Reload shortcuts."
    Print #f, "# --- formatting ---"
    Print #f, "^+1=LtxCycleNumber"
    Print #f, "^+5=LtxCyclePercent"
    Print #f, "^+6=LtxCycleMultiple"
    Print #f, "^+3=LtxCycleDate"
    Print #f, "^+4=LtxCycleCurrency"
    Print #f, "^+f=LtxCycleFontColor"
    Print #f, "^+h=LtxCycleFill"
    Print #f, "^+b=LtxCycleBorder"
    Print #f, "^+a=LtxAutoColor"
    Print #f, "^+z=LtxBlackToggle"
    Print #f, "^+c=LtxCenterAcross"
    Print #f, "^+w=LtxToggleWrap"
    Print #f, "^+g=LtxGridlinesToggle"
    Print #f, "^+i=LtxIndentMore"
    Print #f, "^+o=LtxIndentLess"
    Print #f, "# --- modelling ---"
    Print #f, "^+q=LtxCycleAnchors"
    Print #f, "^+n=LtxFlipSign"
    Print #f, "^+v=LtxToValues"
    Print #f, "^+r=LtxFastFillRight"
    Print #f, "^+d=LtxFastFillDown"
    Print #f, "^+e=LtxWrapIfError"
    Print #f, "# --- auditing ---"
    Print #f, "^{[}=LtxGoPrecedent"
    Print #f, "^{]}=LtxGoDependent"
    Print #f, "^+~=LtxGoBack"
    Print #f, "^+{[}=LtxTraceIn"
    Print #f, "^+{]}=LtxTraceOut"
    Print #f, "^+l=LtxFormulaFlow"
    Print #f, "^+k=LtxHighlightHardcodes"
    Print #f, "# --- export ---"
    Print #f, "^%r=LtxSendToPptPicture"
    Print #f, "^%t=LtxSendToPptTable"
    Print #f, "^%c=LtxSendToPptChart"
    Print #f, "^%w=LtxSendToWordPicture"
    Print #f, "^+p=LtxCopyForPptPicture"
    Print #f, "^+t=LtxCopyForPptTable"
    Print #f, "^+u=LtxRefreshPptFromHere"
    Print #f, "# --- charts ---"
    Print #f, "^+m=LtxInsertCAGR"
    Print #f, "# --- file ---"
    Print #f, "^+s=LtxSaveBackup"
    Close #f
End Sub

' Wrappers so parameterised macros can be bound
Public Sub LtxCopyForPptPicture(): LtxCopyForPpt "picture": End Sub
Public Sub LtxCopyForPptTable(): LtxCopyForPpt "table": End Sub

' Older shortcut files predate the export keys: append them once
Private Sub UpgradeKeysFile()
    Dim txt As String, f As Integer
    f = FreeFile
    Open KeysFilePath() For Input As #f
    txt = Input$(LOF(f), f)
    Close #f
    txt = Replace(txt, "^+{[}=LtxPrecedents", "^+{[}=LtxTraceIn")
    txt = Replace(txt, "^+{]}=LtxDependents", "^+{]}=LtxTraceOut")
    If InStr(txt, "LtxGoPrecedent") = 0 Then
        txt = Replace(txt, "^+[=LtxPrecedents", "^+{[}=LtxTraceIn")
        txt = Replace(txt, "^+]=LtxDependents", "^+{]}=LtxTraceOut")
        f = FreeFile
        Open KeysFilePath() For Output As #f
        Print #f, txt;
        Print #f, "# --- navigation (added) ---"
        Print #f, "^{[}=LtxGoPrecedent"
        Print #f, "^{]}=LtxGoDependent"
        Print #f, "^+~=LtxGoBack"
        Close #f
    End If
    If InStr(txt, "LtxTraceIn") > 0 And InStr(txt, "LtxPrecedents") = 0 Then
        f = FreeFile: Open KeysFilePath() For Output As #f: Print #f, txt;: Close #f
    End If
    If InStr(txt, "LtxBlackToggle") = 0 Then
        f = FreeFile: Open KeysFilePath() For Append As #f: Print #f, "^+z=LtxBlackToggle": Close #f
    End If
    If InStr(txt, "LtxSendToPptPicture") = 0 Then
        f = FreeFile
        Open KeysFilePath() For Append As #f
        Print #f, "# --- export (added) ---"
        Print #f, "^%r=LtxSendToPptPicture"
        Print #f, "^%t=LtxSendToPptTable"
        Print #f, "^%c=LtxSendToPptChart"
        Print #f, "^%w=LtxSendToWordPicture"
        Close #f
    End If
End Sub

Public Sub LtxKeysLoad()
    Dim f As Integer, ln As String, i As Long, k As String, m As String, n As Long
    If Dir(KeysFilePath()) = "" Then WriteDefaults Else UpgradeKeysFile
    LtxKeysUnload
    Set mBound = New Collection
    f = FreeFile
    Open KeysFilePath() For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        ln = Trim$(ln)
        i = InStr(ln, "=")
        If i > 1 And Left$(ln, 1) <> "#" Then
            k = Trim$(Left$(ln, i - 1)): m = Trim$(Mid$(ln, i + 1))
            On Error Resume Next
            Application.OnKey k, "'" & ThisWorkbook.Name & "'!" & m
            If Err.Number = 0 Then mBound.Add k: n = n + 1
            On Error GoTo 0
        End If
    Loop
    Close #f
    Application.StatusBar = "Lattice: " & n & " shortcuts loaded"
End Sub

Public Sub LtxKeysUnload()
    Dim k As Variant
    If mBound Is Nothing Then Exit Sub
    On Error Resume Next
    For Each k In mBound
        Application.OnKey CStr(k)      ' restore Excel's default
    Next
    Set mBound = Nothing
End Sub

Public Sub LtxKeysEdit()
    If Dir(KeysFilePath()) = "" Then WriteDefaults
    Shell "notepad.exe """ & KeysFilePath() & """", vbNormalFocus
End Sub

Public Sub LtxKeysReset()
    If MsgBox("Overwrite your shortcut file with the Lattice defaults?", vbOKCancel + vbQuestion, LTX_APP_XL) <> vbOK Then Exit Sub
    WriteDefaults
    LtxKeysLoad
End Sub

Public Sub LtxKeysList()
    Dim f As Integer, ln As String, s As String, i As Long
    If Dir(KeysFilePath()) = "" Then WriteDefaults
    f = FreeFile
    Open KeysFilePath() For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        i = InStr(ln, "=")
        If i > 1 And Left$(Trim$(ln), 1) <> "#" Then
            s = s & Replace(Replace(Replace(Left$(ln, i - 1), "^", "Ctrl+"), "+", "Shift+"), "%", "Alt+") & vbTab & Mid$(ln, i + 1) & vbLf
        End If
    Loop
    Close #f
    MsgBox s, vbInformation, LTX_APP_XL & " - Shortcuts"
End Sub
