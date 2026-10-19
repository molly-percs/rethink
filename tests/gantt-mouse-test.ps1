# Echte Mausbedienung am Zeitplan (SendInput): Griff am Balkenende ziehen, Balken verschieben, Klick wählt den
# Balken aus; Zeilengrenze überfahren und rechtsklicken, Beschriftungszelle anklicken und mit der echten Tastatur
# tippen (Enter, Tab, Esc, Klick auf die leere Folie). Übernimmt für einige Sekunden Maus UND Tastatur – PowerPoint
# muss vorn bleiben, nichts anfassen. Bildschirmfoto während des Ziehens: tests\out\mouse-drag.png
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
    public const uint Down = 0x0002, Up = 0x0004, RightDown = 0x0008, RightUp = 0x0010;
    // Tastatur über SendInput (Unicode-Zeichen als VK_PACKET, Sondertasten als virtuelle Tasten)
    [StructLayout(LayoutKind.Sequential)] public struct MOUSEINPUT { public int dx; public int dy; public uint mouseData; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }
    [StructLayout(LayoutKind.Sequential)] public struct KEYBDINPUT { public ushort wVk; public ushort wScan; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }
    [StructLayout(LayoutKind.Explicit)] public struct INPUTUNION { [FieldOffset(0)] public MOUSEINPUT mi; [FieldOffset(0)] public KEYBDINPUT ki; }
    [StructLayout(LayoutKind.Sequential)] public struct INPUT { public uint type; public INPUTUNION u; }
    [DllImport("user32.dll", SetLastError = true)] static extern uint SendInput(uint n, INPUT[] inputs, int size);
    static void Send(ushort vk, ushort scan, uint flags) {
        INPUT[] input = new INPUT[1];
        input[0].type = 1;  // INPUT_KEYBOARD
        input[0].u.ki.wVk = vk; input[0].u.ki.wScan = scan; input[0].u.ki.dwFlags = flags;
        SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }
    public static void Key(ushort vk) { Send(vk, 0, 0); Send(vk, 0, 2); }
    public static void Character(char c) { Send(0, c, 4); Send(0, c, 4 | 2); }  // KEYEVENTF_UNICODE, KEYUP
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
function Rect($info, $pattern) {
    $m = [regex]::Match($info, "\[$pattern (-?\d+),(-?\d+),(-?\d+),(-?\d+)\]")
    if (-not $m.Success) { throw "Griff nicht gefunden: $pattern" }
    @([int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value, [int]$m.Groups[4].Value)
}
function Info() { $api.OverlayInfo() }
# hover=… / menue=… / editor=… aus OverlayInfo
function Field($info, $key) { [regex]::Match($info, "$key=(.+?)( (hover|menue|editor)=|$)").Groups[1].Value }
function MouseTo([int]$x, [int]$y) { [void][MouseTestWin]::SetCursorPos($x, $y); Start-Sleep -Milliseconds 300 }
function Click([int]$x, [int]$y) {
    MouseTo $x $y
    [MouseTestWin]::mouse_event([MouseTestWin]::Down, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 80
    [MouseTestWin]::mouse_event([MouseTestWin]::Up, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 600
}
function RightClick() {   # an der aktuellen Mausposition
    [MouseTestWin]::mouse_event([MouseTestWin]::RightDown, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 80
    [MouseTestWin]::mouse_event([MouseTestWin]::RightUp, 0, 0, 0, [UIntPtr]::Zero)
}
function TypeText([string]$text) { foreach ($c in $text.ToCharArray()) { [MouseTestWin]::Character($c); Start-Sleep -Milliseconds 40 }; Start-Sleep -Milliseconds 200 }
function PressKey([int]$vk) { [MouseTestWin]::Key([uint16]$vk); Start-Sleep -Milliseconds 300 }
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

    # ---- Zeilengrenzen, Beschriftungszellen und Eingabefeld: echte Maus (SendInput) und echte Tastatur ----
    $slide.Shapes.Item($name).Select(); Start-Sleep -Milliseconds 600
    $api.GanttCommand('column:remark=1')   # zweite Beschriftungsspalte (Spalte 5): Tab springt dorthin
    $api.OpenDatasheet($false)
    $slide.Shapes.Item($name).Select(); Start-Sleep -Milliseconds 800
    [void][MouseTestWin]::SetForegroundWindow($hwnd)
    $win = $pp.ActiveWindow
    $info = Info
    $c3 = Rect $info 'zelle r3 c0'; $c4 = Rect $info 'zelle r4 c0'
    $xB = [int](($c3[0] + $c3[2]) / 2); $yB = [int](($c3[3] + $c4[1]) / 2)   # Grenze zwischen Tabellenzeile 3 und 4

    # Hover: leuchtet genau an der Grenze, nicht weit davon
    MouseTo $xB $yB
    Expect ((Field (Info) 'hover') -eq 'waagerecht zeile=4') "Maus auf der Zeilengrenze: Band zwischen Zeile 3 und 4 ($(Field (Info) 'hover'))"
    MouseTo $xB ($yB + 14)
    Expect ((Field (Info) 'hover') -eq '-') 'Maus 14 px von der Grenze: kein Band'
    $off = 0
    foreach ($o in 4, 3) { MouseTo $xB ($yB + $o); if ((Field (Info) 'hover') -eq 'waagerecht zeile=4') { $off = $o; break } }
    Expect ($off -gt 0) "Maus $off px neben der Grenze: Band leuchtet noch (Reichweite ~4 px)"

    # Rechtsklick dort, wo es leuchtet: das Menü arbeitet an genau dieser Grenze (nicht am Rundmenü / an keiner Grenze)
    if ($off -gt 0) {
        RightClick
        Start-Sleep -Milliseconds 500
        PressKey 0x1B   # Esc schließt das Menü (und die Minileiste), bevor die Automation gefragt wird
        Start-Sleep -Milliseconds 500
        Expect ((Field (Info) 'menue') -eq 'trennzeile=4 trenndatum=-') "Rechtsklick $off px neben der Grenze: Menü mit den Trennlinien-Einträgen ($(Field (Info) 'menue'))"
        $px = $xB; $py = $yB
        $ptX = ($px - $win.PointsToScreenPixelsX(0)) * 1000 / ($win.PointsToScreenPixelsX(1000) - $win.PointsToScreenPixelsX(0))
        $ptY = ($py - $win.PointsToScreenPixelsY(0)) * 1000 / ($win.PointsToScreenPixelsY(1000) - $win.PointsToScreenPixelsY(0))
        Expect ($api.GanttMenu($ptX, $ptY) -match 'Waagerechte Trennlinie hier (einfügen|entfernen)=rowLine=4') 'Menü an der Grenze enthält die Trennlinien-Einträge'
    }

    # Zelle anklicken, tippen, Enter
    $slide.Shapes.Item($name).Select(); Start-Sleep -Milliseconds 500
    $info = Info
    $a2 = Handle $info 'zelle r2 c0'
    Click $a2[0] $a2[1]
    Expect ((Field (Info) 'editor') -eq 'true') 'Klick auf eine Beschriftungszelle öffnet das Eingabefeld'
    TypeText 'Maus Eins'
    PressKey 0x0D
    Start-Sleep -Milliseconds 700
    Expect ($api.DatasheetCell(2, 0) -eq 'Maus Eins' -and (Field (Info) 'editor') -eq 'false') "Tippen + Enter: Wert steht im Datenblatt ($($api.DatasheetCell(2, 0)))"

    # Tab übernimmt und springt in die nächste Spalte derselben Zeile
    $a3 = Handle (Info) 'zelle r3 c0'
    Click $a3[0] $a3[1]
    TypeText 'Tab Eins'
    PressKey 0x09
    Start-Sleep -Milliseconds 700
    Expect ($api.DatasheetCell(3, 0) -eq 'Tab Eins') "Tab: Wert der ersten Spalte übernommen ($($api.DatasheetCell(3, 0)))"
    Expect ((Field (Info) 'editor') -eq 'true') 'Tab: Eingabefeld der nächsten Spalte ist offen'
    TypeText 'Hinweis Tab'
    PressKey 0x0D
    Start-Sleep -Milliseconds 700
    Expect ($api.DatasheetCell(3, 5) -eq 'Hinweis Tab') "Tab, tippen, Enter: Wert der zweiten Spalte im Datenblatt ($($api.DatasheetCell(3, 5)))"

    # Esc verwirft; Markieren per Ziehen im Eingabefeld verschiebt das Overlay nicht
    $a2 = Handle (Info) 'zelle r2 c0'
    Click $a2[0] $a2[1]
    $before = [regex]::Match((Info), 'fenster=[-\d,]+').Value
    Drag ($a2[0] - 30) $a2[1] 60 $null
    Start-Sleep -Milliseconds 300
    Expect ((Field (Info) 'editor') -eq 'true' -and [regex]::Match((Info), 'fenster=[-\d,]+').Value -eq $before) 'Ziehen im Eingabefeld (Text markieren) verschiebt das Overlay nicht'
    TypeText 'Verworfen'
    PressKey 0x1B
    Start-Sleep -Milliseconds 600
    Expect ($api.DatasheetCell(2, 0) -eq 'Maus Eins' -and (Field (Info) 'editor') -eq 'false') "Esc verwirft den getippten Text ($($api.DatasheetCell(2, 0)))"

    # Tippen, dann auf die leere Folie klicken: PowerPoints Auswahlwechsel darf den Text nicht verschlucken
    $slide.Shapes.Item($name).Select(); Start-Sleep -Milliseconds 500
    $a4 = Handle (Info) 'zelle r4 c0'
    Click $a4[0] $a4[1]
    Expect ((Field (Info) 'editor') -eq 'true') 'Eingabefeld in Zeile 4 offen'
    TypeText 'Leerklick'
    $emptyX = [int]$win.PointsToScreenPixelsX(480); $emptyY = [int]$win.PointsToScreenPixelsY(50)   # oberhalb des Zeitplans (der beginnt bei Top = 135 pt)
    Click $emptyX $emptyY
    Start-Sleep -Milliseconds 700
    Expect ($api.DatasheetCell(4, 0) -eq 'Leerklick' -and (Field (Info) 'editor') -eq 'false') "Klick auf die leere Folie übernimmt den getippten Text ($($api.DatasheetCell(4, 0)))"
    $api.CloseDatasheet()
} finally {
    $pres.Saved = -1; $pres.Close()
    if (-not $pptWasRunning) { $pp.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
}
if ($failed) { throw "$failed Prüfungen fehlgeschlagen" }
"Alle Prüfungen bestanden."
