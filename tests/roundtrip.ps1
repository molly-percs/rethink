# Prüft "Daten bearbeiten": Öffnet für jedes Diagramm die eingebettete Arbeitsmappe in PowerPoint
# und vergleicht Reihen, Werte und Kategorien vor und nach dem Aktivieren. Ändert sich etwas,
# passen Formeln (Sheet1!$B$n) und Arbeitsmappe nicht zusammen.
# Arbeitet auf einer Kopie; PowerPoint und Excel öffnen sich dabei kurz (minimiert).
param([Parameter(Mandatory)][string]$Pptx)
$ErrorActionPreference = 'Stop'

$copy = Join-Path $env:TEMP ("rethink-roundtrip-{0}.pptx" -f $PID)
Copy-Item (Resolve-Path $Pptx).Path $copy -Force

function Snapshot($chart) {
    $out = @()
    for ($i = 1; $i -le $chart.SeriesCollection().Count; $i++) {
        $s = $chart.SeriesCollection($i)
        $out += "{0}={1}|{2}" -f $s.Name, (($s.Values | ForEach-Object { $_ }) -join ','), (($s.XValues | ForEach-Object { $_ }) -join ',')
    }
    $out -join ' ; '
}

$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$pp = New-Object -ComObject PowerPoint.Application
$failed = 0
try {
    $pres = $pp.Presentations.Open($copy, 0, 0, -1)   # beschreibbar, mit Fenster (für ChartData nötig)
    try { $pp.WindowState = 2 } catch { }
    foreach ($slide in $pres.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasChart -ne -1) { continue }
            $chart = $shape.Chart
            $before = Snapshot $chart
            $chart.ChartData.Activate()
            $book = $chart.ChartData.Workbook
            try { $book.Application.WindowState = -4140 } catch { }   # xlMinimized
            $rows = $book.Worksheets.Item(1).UsedRange.Rows.Count
            $after = Snapshot $chart
            $book.Close()
            $same = $before -eq $after
            if (-not $same) { $failed++ }
            "{0} Folie {1} {2}: {3} Zeilen im Blatt" -f ($(if ($same) { 'OK  ' } else { 'FEHL' })), $slide.SlideIndex, $shape.Name, $rows
            if (-not $same) { "     vorher:  $before"; "     nachher: $after" }
        }
    }
    $pres.Saved = -1
    $pres.Close()
} finally {
    if (-not $wasRunning) { $pp.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
    Remove-Item $copy -ErrorAction SilentlyContinue
}
if ($failed) { throw "$failed Diagramme haben nach dem Öffnen der Arbeitsmappe andere Werte" }
"Alle Diagramme unverändert."
