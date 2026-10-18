# Exportiert alle Folien einer Präsentation als PNG (Sichtprüfung des Layouts), über PowerPoint-Automation.
#   tests\export-png.ps1 tests\out\gantt.pptx        → tests\out\gantt-1.png, gantt-2.png, …
param([Parameter(Mandatory)][string]$Pptx, [int]$Width = 1600, [int]$Height = 900)
$ErrorActionPreference = 'Stop'
$file = (Resolve-Path $Pptx).Path
$stem = Join-Path (Split-Path $file) ([IO.Path]::GetFileNameWithoutExtension($file))
$pptWasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$pp = New-Object -ComObject PowerPoint.Application
try {
    $pres = $pp.Presentations.Open($file, -1, 0, 0)  # schreibgeschützt, ohne Fenster
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $pres.Slides.Item($i).Export("$stem-$i.png", 'PNG', $Width, $Height)
        "$stem-$i.png"
    }
    $pres.Close()
} finally {
    if (-not $pptWasRunning) { $pp.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
}
