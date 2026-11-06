# Removes the identity Office stamps into every file it saves. Opens an .xlam / .ppam / .dotm / .xlsx
# with System.IO.Packaging (the same API the ribbon injection uses, which Office accepts) and:
#   - sets the package creator and last-modified-by (docProps/core.xml) to the project name;
#   - deletes any x15ac:absPath element in xl/workbook.xml (Excel records the folder the file was
#     saved from), together with its mc:AlternateContent wrapper when that is all the wrapper holds;
#   - leaves every other part, including vbaProject.bin, untouched.
# Dot-source this file, then call Invoke-OfficeScrub <path>.
function Invoke-OfficeScrub([string]$path, [string]$creator = 'Chiltern Capital Management') {
    Add-Type -AssemblyName WindowsBase
    $pkg = [System.IO.Packaging.Package]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite)
    try {
        $pkg.PackageProperties.Creator = $creator
        $pkg.PackageProperties.LastModifiedBy = $creator
        $wbUri = New-Object System.Uri -ArgumentList '/xl/workbook.xml', ([UriKind]::Relative)
        if ($pkg.PartExists($wbUri)) {
            $part = $pkg.GetPart($wbUri)
            $xml = New-Object System.Xml.XmlDocument
            $s = $part.GetStream([IO.FileMode]::Open, [IO.FileAccess]::Read); $xml.Load($s); $s.Close()
            $nodes = @($xml.SelectNodes("//*[local-name()='absPath']"))
            foreach ($n in $nodes) {
                $choice = $n.ParentNode; $alt = $choice.ParentNode
                if ($alt -and $alt.LocalName -eq 'AlternateContent' -and $alt.ChildNodes.Count -eq 1 -and $choice.ChildNodes.Count -eq 1) {
                    [void]$alt.ParentNode.RemoveChild($alt)
                } else {
                    [void]$choice.RemoveChild($n)
                }
            }
            if ($nodes.Count -gt 0) {
                $s = $part.GetStream([IO.FileMode]::Create, [IO.FileAccess]::Write); $xml.Save($s); $s.Close()
            }
        }
        $pkg.Flush()
    } finally { $pkg.Close() }
    "  scrubbed $(Split-Path $path -Leaf): creator set, $($nodes.Count) absPath element(s) removed"
}
