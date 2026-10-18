# Erzeugt tests/out/waterfall-template.pptx ohne PowerPoint: nimmt samples/template.pptx,
# behält den Folientitel und ersetzt Diagramme und Beschriftungen durch ein Rechteck "Bridge" als Platzhalter.
param([string]$Out = (Join-Path $PSScriptRoot 'out\waterfall-template.pptx'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
Copy-Item (Join-Path $PSScriptRoot '..\samples\template.pptx') $Out -Force

$zip = [IO.Compression.ZipFile]::Open($Out, 'Update')
try {
    $entry = $zip.GetEntry('ppt/slides/slide1.xml')
    $reader = New-Object IO.StreamReader($entry.Open())
    $xml = [xml]$reader.ReadToEnd()
    $reader.Close()

    $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
    $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    $tree = $xml.SelectSingleNode('/p:sld/p:cSld/p:spTree', $ns)
    foreach ($name in 'LeftChart', 'RightChart', 'LeftChartTitle', 'RightChartTitle') {
        $node = $tree.SelectSingleNode("*[*/p:cNvPr[@name='$name']]", $ns)
        if ($node) { [void]$tree.RemoveChild($node) }
    }
    $placeholder = @"
<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr><p:cNvPr id="90" name="Bridge"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="838200" y="1500000"/><a:ext cx="10515600" cy="4700000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr>
</p:sp>
"@
    $fragment = [xml]$placeholder
    [void]$tree.AppendChild($xml.ImportNode($fragment.DocumentElement, $true))

    $entry.Delete()
    $newEntry = $zip.CreateEntry('ppt/slides/slide1.xml')
    $writer = New-Object IO.StreamWriter($newEntry.Open(), (New-Object Text.UTF8Encoding($false)))
    $xml.Save($writer)
    $writer.Close()
} finally { $zip.Dispose() }
"Vorlage geschrieben: $Out"
