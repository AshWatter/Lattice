# Builds a Settings dialog from a spec: @{k=key; t='check'|'text'|'choice'; l=label; d=default; o='a|b'}.
# Controls are named chk_<key> / txt_<key> / cbo_<key> and carry the default in Tag, so one
# generic code-behind reads and writes every setting.
function Add-SettingsForm($vbproj, [string]$name, [string]$caption, $items, [string]$codePath) {
    $comp = $vbproj.VBComponents.Add(3)
    $comp.Name = $name
    $comp.Properties.Item('Caption').Value = $caption
    $comp.Properties.Item('Width').Value = 520
    try { $comp.Properties.Item('StartUpPosition').Value = 1 } catch {}
    $d = $comp.Designer
    $ttl = $d.Controls.Add('Forms.Label.1', 'lblTitle', $true)
    $ttl.Left = 8; $ttl.Top = 6; $ttl.Width = 500; $ttl.Height = 22; $ttl.Caption = $caption
    $y = 58
    foreach ($it in $items) {
        if ($it.t -eq 'check') {
            $c = $d.Controls.Add('Forms.CheckBox.1', "chk_$($it.k)", $true)
            $c.Left = 16; $c.Top = $y; $c.Width = 480; $c.Height = 18; $c.Caption = $it.l; $c.Tag = $it.d
        } else {
            $l = $d.Controls.Add('Forms.Label.1', "lbl_$($it.k)", $true)
            $l.Left = 16; $l.Top = $y + 3; $l.Width = 200; $l.Height = 16; $l.Caption = $it.l
            if ($it.t -eq 'choice') {
                $c = $d.Controls.Add('Forms.ComboBox.1', "cbo_$($it.k)", $true); $c.Style = 2; $c.Tag = "$($it.o);$($it.d)"
            } else {
                $c = $d.Controls.Add('Forms.TextBox.1', "txt_$($it.k)", $true); $c.Tag = $it.d
            }
            $c.Left = 220; $c.Top = $y; $c.Width = 280; $c.Height = 20
        }
        $y += 30
    }
    $lbl = $d.Controls.Add('Forms.Label.1', 'lblInfo', $true)
    $lbl.Left = 8; $lbl.Top = $y + 8; $lbl.Width = 500; $lbl.Height = 14
    $lbl.Caption = 'Saved under %APPDATA%\Lattice; changes apply immediately.   Enter = save   Esc = cancel'
    $x = 8
    foreach ($b in @(@{n='btnSave';c='Save'}, @{n='btnClose';c='Cancel'})) {
        $btn = $d.Controls.Add('Forms.CommandButton.1', $b.n, $true)
        $btn.Caption = $b.c; $btn.Left = $x; $btn.Top = $y + 28; $btn.Width = 84; $btn.Height = 24
        $x += 90
    }
    $comp.Properties.Item('Height').Value = $y + 28 + 24 + 44
    $comp.CodeModule.AddFromString((Get-Content $codePath -Raw))
    Write-Step "  built $name"
}

$script:LtxExcelSettings = @(
    @{k='AutoName';t='check';l='Name exported ranges automatically (keeps links alive when rows move)';d='1'},
    @{k='ExportActivate';t='check';l='Switch to PowerPoint / Word after an export';d='1'},
    @{k='ExportSizeMode';t='choice';l='Match mode for exports';o='none|width|height|size';d='width'},
    @{k='HardcodeWhitelist';t='text';l='Numbers allowed inside formulas';d='0|1|2|100|12|365|360|1000|1000000|4|52|30|31|0.5'},
    @{k='LibraryPath';t='text';l='Personal library folder';d=''},
    @{k='ChartPalette';t='text';l='Chart palette (hex, | separated)';d='1F3864|2E75B6|9DC3E6|BDD7EE|7F7F7F|BFBFBF|C00000|ED7D31'},
    @{k='ColInput';t='text';l='AutoColor: inputs (BGR number)';d='16711680'},
    @{k='ColFormula';t='text';l='AutoColor: formulas';d='0'},
    @{k='ColLink';t='text';l='AutoColor: links to other sheets';d='32768'},
    @{k='ColExternal';t='text';l='AutoColor: external links';d='8388736'},
    @{k='ColPartial';t='text';l='AutoColor: formulas with hard-codes';d='128'}
)
$script:LtxPptSettings = @(
    @{k='PictureFormat';t='choice';l='Picture format for imports';o='emf|png';d='emf'},
    @{k='SizeMode';t='choice';l='Default size mode for new links';o='width|height|size|none';d='width'},
    @{k='BlackFonts';t='check';l='Force fonts to black on import';d='1'},
    @{k='HideGridlines';t='check';l='Hide gridlines in imported pictures';d='1'},
    @{k='ChartNoBorder';t='check';l='Remove chart borders in pictures';d='1'},
    @{k='CheckErrors';t='check';l='Warn about formula errors on import';d='1'},
    @{k='QuietRefresh';t='check';l='No message after a successful refresh';d='1'},
    @{k='AirplaneMode';t='check';l='Airplane mode: never open workbooks from disk';d='0'},
    @{k='MarginPt';t='text';l='Import margin from the slide edge (pt)';d='36'},
    @{k='Gap';t='text';l='Default gap between shapes (pt)';d='8'},
    @{k='TableFontSize';t='text';l='Imported table font size';d='10'},
    @{k='TableHeaderHex';t='text';l='Imported table header colour (hex)';d='1F3864'},
    @{k='ChartPalette';t='text';l='Chart palette (hex, | separated)';d='1F3864|2E75B6|9DC3E6|BDD7EE|7F7F7F|BFBFBF|C00000|ED7D31'},
    @{k='Template';t='text';l='Deck template (.potx path)';d=''},
    @{k='CheckWords';t='text';l='Deck Check words to flag';d='TBD,TBU,xx,XXX,[placeholder],lorem,DRAFT,TODO'}
)
