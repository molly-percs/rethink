# Erzeugt samples/template.pptx über das PowerPoint-Objektmodell:
# eine Folie mit Titel, zwei Beschriftungen und zwei nativen Diagrammen, deren Formnamen
# den Namen in sample.ppttc entsprechen.
param([string]$Out = (Join-Path $PSScriptRoot 'template.pptx'))
$ErrorActionPreference = 'Stop'

$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$pp = New-Object -ComObject PowerPoint.Application
try {
    # Diagramme lassen sich nur in einer Präsentation mit Fenster einfügen; das Fenster wird minimiert.
    $pres = $pp.Presentations.Add(-1)
    try { $pp.WindowState = 2 } catch { }
    $slide = $pres.Slides.Add(1, 11)          # ppLayoutTitleOnly
    $title = $slide.Shapes.Item(1)
    $title.Name = 'SlideTitle'
    $title.TextFrame.TextRange.Text = '<SlideTitle>'

    foreach ($label in @(@('LeftChartTitle', 40), @('RightChartTitle', 500))) {
        $box = $slide.Shapes.AddTextbox(1, $label[1], 120, 420, 30)
        $box.Name = $label[0]
        $box.TextFrame.TextRange.Text = "<$($label[0])>"
        $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2   # zentriert
    }

    foreach ($chart in @(@('LeftChart', 52, 40), @('RightChart', 51, 500))) {   # gestapelt / gruppiert
        $shape = $slide.Shapes.AddChart2(-1, $chart[1], $chart[2], 160, 420, 330)
        $shape.Name = $chart[0]
        try { $shape.Chart.ChartData.Workbook.Close() } catch { }
    }

    $pres.SaveAs($Out)
    $pres.Close()
    "Vorlage geschrieben: $Out"
} finally {
    if (-not $wasRunning) { $pp.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
}
