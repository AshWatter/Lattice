# Fails (exit 1) when any git-tracked file carries a personal identifier: a name fragment, a local
# user path or a private email. The patterns are assembled from character codes at run time so this
# script does not itself contain them. Every file is checked as Latin-1 text and as UTF-16LE text;
# Office packages (.xlam .ppam .dotm .xlsx .pptx .docx) are opened and every zip entry is checked the
# same way; PDFs are checked raw, with every FlateDecode stream inflated, and their metadata read.
param([string]$Root = (Split-Path $PSScriptRoot -Parent))
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
function S([int[]]$codes) { -join ($codes | ForEach-Object { [char]$_ }) }
$patterns = @(
    (S 97,108,118,97,114,111),                                   # first name
    (S 103,111,114,110,101,115),                                 # surname, plain
    ((S 103,111,114,110) + [char]233),                           # surname with the accent (code point)
    ((S 103,111,114,110) + [char]195),                           # surname with the accent, UTF-8 bytes read as Latin-1
    (S 97,115,104,119,97,116,116,101,114,64,103,109,97,105,108,46,99,111,109),   # private email
    (S 117,115,101,114,115,92,97,108,118,97,114),                # local profile path, backslash
    (S 117,115,101,114,115,47,97,108,118,97,114)                 # local profile path, forward slash
)
$latin1 = [Text.Encoding]::GetEncoding(28591); $utf16 = [Text.Encoding]::Unicode
$hits = New-Object System.Collections.Generic.List[string]
function Test-Bytes([byte[]]$bytes, [string]$label) {
    if (-not $bytes -or $bytes.Length -eq 0) { return }
    foreach ($text in @($latin1.GetString($bytes), $utf16.GetString($bytes))) {
        foreach ($p in $patterns) {
            $i = $text.IndexOf($p, [StringComparison]::OrdinalIgnoreCase)
            if ($i -ge 0) {
                $from = [Math]::Max(0, $i - 30); $len = [Math]::Min(80, $text.Length - $from)
                $ctx = ($text.Substring($from, $len) -replace '[\x00-\x1f]', ' ')
                $script:hits.Add("$label : ...$ctx..."); break
            }
        }
    }
}
function Inflate([byte[]]$data) {
    try {
        $ms = New-Object IO.MemoryStream(,$data); $ms.Position = 2     # skip the zlib header
        $ds = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Decompress)
        $out = New-Object IO.MemoryStream; $ds.CopyTo($out); $ds.Close(); $out.ToArray()
    } catch { $null }
}
function Test-Pdf([byte[]]$bytes, [string]$label) {
    Test-Bytes $bytes "$label (raw)"
    $text = $latin1.GetString($bytes); $pos = 0; $n = 0
    while (($i = $text.IndexOf('stream', $pos)) -ge 0) {
        $start = $i + 6; if ($text[$start] -eq "`r") { $start++ }; if ($text[$start] -eq "`n") { $start++ }
        $end = $text.IndexOf('endstream', $start); if ($end -lt 0) { break }
        $seg = New-Object byte[] ($end - $start); [Array]::Copy($bytes, $start, $seg, 0, $end - $start)
        $inflated = Inflate $seg
        if ($inflated) { $n++; Test-Bytes $inflated "$label (stream $n inflated)" }
        $pos = $end + 9
    }
}
Push-Location $Root
$files = git ls-files
Pop-Location
$zipExt = '.xlam', '.ppam', '.dotm', '.xlsx', '.pptx', '.docx', '.zip'
$count = 0
foreach ($rel in $files) {
    $full = Join-Path $Root $rel
    if (-not (Test-Path $full -PathType Leaf)) { continue }
    $count++
    $ext = [IO.Path]::GetExtension($rel).ToLower()
    $bytes = [IO.File]::ReadAllBytes($full)
    if ($zipExt -contains $ext) {
        Test-Bytes $bytes "$rel (container)"
        try {
            $zip = [IO.Compression.ZipFile]::OpenRead($full)
            foreach ($e in $zip.Entries) {
                $es = $e.Open(); $ms = New-Object IO.MemoryStream; $es.CopyTo($ms); $es.Close()
                Test-Bytes $ms.ToArray() "$rel!$($e.FullName)"
            }
            $zip.Dispose()
        } catch { $hits.Add("$rel : could not open as zip ($($_.Exception.Message))") }
    } elseif ($ext -eq '.pdf') {
        Test-Pdf $bytes $rel
    } else {
        Test-Bytes $bytes $rel
    }
}
"privacy scan: $count tracked files checked"
if ($hits.Count -gt 0) {
    "FAIL: $($hits.Count) hit(s)"
    $hits | ForEach-Object { "  $_" }
    exit 1
}
"PASS: no personal identifiers found"
exit 0
