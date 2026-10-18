# Prüft eine von rethink erzeugte Präsentation:
#  1. eingebettete Arbeitsmappen (Rohdaten) direkt aus dem ZIP lesen
#  2. Datei in PowerPoint ohne Fenster öffnen, Texte und Diagrammreihen ausgeben
#  3. optional jede Folie als PNG exportieren
param(
    [Parameter(Mandatory)][string]$Pptx,
    [string]$PngDir,
    [int]$ExpectSlides = -1,   # bricht ab, wenn PowerPoint eine andere Anzahl sieht
    [int]$ExpectCharts = -1    # z. B. weil eine Reparatur Diagramme entfernt hat
)
$ErrorActionPreference = 'Stop'
$path = (Resolve-Path $Pptx).Path
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

"== Eingebettete Arbeitsmappen"
$zip = [IO.Compression.ZipFile]::OpenRead($path)
try {
    foreach ($entry in $zip.Entries | Where-Object FullName -like 'ppt/embeddings/rethink_data*.xlsx') {
        $inner = New-Object IO.Compression.ZipArchive($entry.Open())
        $sheet = $inner.GetEntry('xl/worksheets/sheet1.xml')
        $xml = [xml](New-Object IO.StreamReader($sheet.Open())).ReadToEnd()
        $rows = foreach ($r in $xml.worksheet.sheetData.row) {
            ($r.c | ForEach-Object { if ($_.is) { $_.is.t.'#text' } else { $_.v } }) -join ' | '
        }
        "  {0}: {1}" -f $entry.FullName, ($rows -join ' // ')
        $inner.Dispose()
    }
} finally { $zip.Dispose() }

"== PowerPoint"
$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$pp = New-Object -ComObject PowerPoint.Application
try {
    $pres = $pp.Presentations.Open($path, -1, 0, 0)   # schreibgeschützt, ohne Fenster
    "  Folien: $($pres.Slides.Count)"
    $slideCount = $pres.Slides.Count
    $chartCount = 0
    foreach ($slide in $pres.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasChart -eq -1) {
                $chartCount++
                $chart = $shape.Chart
                for ($i = 1; $i -le $chart.SeriesCollection().Count; $i++) {
                    $s = $chart.SeriesCollection($i)
                    "  Folie {0} | {1} | Reihe '{2}' = {3} | Kategorien {4}" -f $slide.SlideIndex, $shape.Name, $s.Name,
                        (($s.Values | ForEach-Object { $_ }) -join ', '), (($s.XValues | ForEach-Object { $_ }) -join ', ')
                }
            } elseif ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $text = $shape.TextFrame.TextRange.Text -replace '\s+', ' '
                if ($text.Length -gt 60) { $text = $text.Substring(0, 60) + '…' }
                "  Folie {0} | {1} | '{2}'" -f $slide.SlideIndex, $shape.Name, $text
            }
        }
        if ($PngDir) {
            New-Item -ItemType Directory -Force $PngDir | Out-Null
            $slide.Export((Join-Path (Resolve-Path $PngDir) ("slide{0}.png" -f $slide.SlideIndex)), 'PNG', 1280, 720)
        }
    }
    $pres.Close()
} finally {
    if (-not $wasRunning) { $pp.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
}
"  Diagramme: $chartCount"
if ($ExpectSlides -ge 0 -and $slideCount -ne $ExpectSlides) { throw "Erwartet $ExpectSlides Folien, PowerPoint sieht $slideCount" }
if ($ExpectCharts -ge 0 -and $chartCount -ne $ExpectCharts) { throw "Erwartet $ExpectCharts Diagramme, PowerPoint sieht $chartCount" }
"  OK"
