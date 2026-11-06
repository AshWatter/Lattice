# Cuts the brand sheets in docs\brand into usable assets with headless Chrome (System.Drawing cannot read WebP):
# logo mark / wordmark in light and dark, and the three spot illustrations. Crops are percentages of each sheet.
param([string]$Brand = (Join-Path (Split-Path $PSScriptRoot -Parent) 'docs\brand'))
$chrome = @("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
$tmp = Join-Path $env:TEMP 'lattice_brand'; New-Item -ItemType Directory -Force $tmp | Out-Null
function Crop($src, $out, $w, $h, $sizeX, $sizeY, $posX, $posY, $bg = 'transparent') {
    $uri = (New-Object System.Uri((Join-Path $Brand $src))).AbsoluteUri
    $html = "<!doctype html><html><body style='margin:0;background:$bg'><div style='width:${w}px;height:${h}px;background:url($uri) $posX $posY / $sizeX $sizeY no-repeat'></div></body></html>"
    $hp = Join-Path $tmp ($out + '.html'); [IO.File]::WriteAllText($hp, $html)
    $args = @('--headless=new', '--disable-gpu', '--no-first-run', '--hide-scrollbars', "--user-data-dir=`"$tmp\profile`"", '--default-background-color=00000000', "--window-size=$w,$h", "--screenshot=`"$(Join-Path $Brand ($out + '.png'))`"", (New-Object System.Uri($hp)).AbsoluteUri)
    $p = Start-Process -FilePath $chrome -ArgumentList $args -WindowStyle Hidden -PassThru; [void]$p.WaitForExit(60000)
    "$out : $(if (Test-Path (Join-Path $Brand ($out + '.png'))) { 'ok' } else { 'failed' })"
}
# logo sheet: left column ~28% = mark, right ~72% = wordmark; top row light, bottom row dark
Crop 'logo_sheet.webp' 'mark_light'     560 404 '357%' '200%' '0' '0'
Crop 'logo_sheet.webp' 'wordmark_light' 1424 404 '139%' '200%' '100%' '0'
Crop 'logo_sheet.webp' 'mark_dark'      560 404 '357%' '212%' '0' '100%'
Crop 'logo_sheet.webp' 'wordmark_dark'  1424 404 '139%' '212%' '100%' '100%'
# illustrations: three across, central band
Crop 'illustrations.webp' 'ill_model' 540 330 '300%' '167%' '0' '50%'
Crop 'illustrations.webp' 'ill_link'  540 330 '300%' '167%' '50%' '50%'
Crop 'illustrations.webp' 'ill_slide' 540 330 '300%' '167%' '100%' '50%'
