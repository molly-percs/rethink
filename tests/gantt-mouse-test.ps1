# Echte Mausbedienung am Zeitplan (SendInput): Griff am Balkenende ziehen, Balken verschieben, Klick wählt den
# Balken aus. Übernimmt für einige Sekunden den Mauszeiger. Bildschirmfoto während des Ziehens: tests\out\mouse-drag.png
param([string]$OutDir = (Join-Path $PSScriptRoot 'out'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class MouseTestWin {
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, int dx, int dy, uint data, UIntPtr extra);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    public const uint Down = 0x0002, Up = 0x0004;
}
"@
[void][MouseTestWin]::SetProcessDPIAware()  # Bildschirmpixel wie im Overlay
$failed = 0
function Expect($ok, $what) { if ($ok) { "ok    $what" } else { "FEHL  $what"; $script:failed++ } }
function Handle($info, $pattern) {
    $m = [regex]::Match($info, "\[$pattern (-?\d+),(-?\d+),(-?\d+),(-?\d+)\]")
    if (-not $m.Success) { throw "Griff nicht gefunden: $pattern" }
    @([int](([int]$m.Groups[1].Value + [int]$m.Groups[3].Value) / 2), [int](([int]$m.Groups[2].Value + [int]$m.Groups[4].Value) / 2))
}
function Drag([int]$x, [int]$y, [int]$dx, $shot) {
    [void][MouseTestWin]::SetCursorPos($x, $y); Start-Sleep -Milliseconds 150
    [MouseTestWin]::mouse_event([MouseTestWin]::Down, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 120
    for ($i = 1; $i -le 10; $i++) { [void][MouseTestWin]::SetCursorPos($x + [int]($dx * $i / 10), $y); Start-Sleep -Milliseconds 60 }
    Start-Sleep -Milliseconds 300
    if ($shot) {
        $bmp = New-Object Drawing.Bitmap 900, 500
        $g = [Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen([Math]::Max(0, $x - 450), [Math]::Max(0, $y - 250), 0, 0, $bmp.Size)
        $bmp.Save($shot); $g.Dispose(); $bmp.Dispose()
    }
    [MouseTestWin]::mouse_event([MouseTestWin]::Up, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 500
}
function Dates($api, $role) { [regex]::Match($api.GanttTargets(), [regex]::Escape($role) + '=([^;]+)').Groups[1].Value }

$pptWasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = -1
$pres = $pp.Presentations.Add(-1)
try {
    $slide = $pres.Slides.Add(1, 12)
    $hwnd = [IntPtr]$pp.HWND
    [void][MouseTestWin]::ShowWindow($hwnd, 3); [void][MouseTestWin]::SetForegroundWindow($hwnd)
    $api = $pp.COMAddIns.Item('rethink.addin').Object
    $name = $api.InsertElement('gantt', 'up', $false)
    $slide.Shapes.Item($name).Select(); Start-Sleep -Milliseconds 800
    $info = $api.OverlayInfo()

    $before = Dates $api 'bar r3'
    $end = Handle $info 'ende bar r3'
    Drag $end[0] $end[1] 70 (Join-Path $OutDir 'mouse-drag.png')
    $after = Dates $api 'bar r3'
    Expect ($after -ne $before -and $after.Split('..')[0] -eq $before.Split('..')[0]) "Maus: Balkenende ziehen ($before → $after)"

    $slide.Shapes.Item($name).Select(); Start-Sleep -Milliseconds 600
    $info = $api.OverlayInfo()
    $before = Dates $api 'bar r4'
    $body = Handle $info 'verschieben bar r4'
    Drag $body[0] $body[1] -50 $null
    $after = Dates $api 'bar r4'
    Expect ($after -ne $before) "Maus: Balken verschieben ($before → $after)"

    $slide.Shapes.Item($name).Select(); Start-Sleep -Milliseconds 600
    $info = $api.OverlayInfo()
    $body = Handle $info 'verschieben bar r2'
    [void][MouseTestWin]::SetCursorPos($body[0], $body[1]); Start-Sleep -Milliseconds 150
    [MouseTestWin]::mouse_event([MouseTestWin]::Down, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 80
    [MouseTestWin]::mouse_event([MouseTestWin]::Up, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 600
    $sel = $pp.ActiveWindow.Selection
    # Innerhalb einer Gruppe meldet PowerPoint die Gruppe in ShapeRange und die Form in ChildShapeRange
    Expect ($sel.HasChildShapeRange -and $sel.ChildShapeRange.Item(1).Name -eq 'bar r2') "Maus: Klick auf den Balken wählt ihn in der Gruppe aus ($($sel.ChildShapeRange.Item(1).Name))"
    Expect ($api.OverlayInfo() -match 'sichtbar=true') 'Overlay bleibt bei ausgewähltem Balken sichtbar'
} finally {
    $pres.Saved = -1; $pres.Close()
    if (-not $pptWasRunning) { $pp.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
}
if ($failed) { throw "$failed Prüfungen fehlgeschlagen" }
"Alle Prüfungen bestanden."
