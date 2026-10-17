# Erzeugt tests/out/types.pptx: je ein Linien- (mit Markern), Flächen- und Kreisdiagramm.
param([string]$Out = (Join-Path $PSScriptRoot 'out\types.pptx'))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null

$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$pp = New-Object -ComObject PowerPoint.Application
try {
    $pres = $pp.Presentations.Add(-1)   # Diagramme brauchen ein Fenster; es wird minimiert
    try { $pp.WindowState = 2 } catch { }
    $slide = $pres.Slides.Add(1, 12)    # ppLayoutBlank
    $charts = @(@('LineChart', 65, 20), @('AreaChart', 1, 330), @('PieChart', 5, 640))   # xlLineMarkers, xlArea, xlPie
    foreach ($c in $charts) {
        $shape = $slide.Shapes.AddChart2(-1, $c[1], $c[2], 100, 300, 300)
        $shape.Name = $c[0]
        try { $shape.Chart.ChartData.Workbook.Close() } catch { }
    }
    $pres.SaveAs($Out)
    $pres.Close()
    "Vorlage geschrieben: $Out"
} finally {
    if (-not $wasRunning) { $pp.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
}
