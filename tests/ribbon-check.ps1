# Prüft, ob PowerPoint das Ribbon-XML des Add-ins annimmt (onLoad wurde gerufen), aktiviert den Tab "rethink"
# und speichert ein Bildschirmfoto des Menübands nach tests\out\ribbon.png.
param([string]$Out = (Join-Path $PSScriptRoot 'out\ribbon.png'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class RibbonCheckWin {
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
    public struct RECT { public int Left, Top, Right, Bottom; }
}
"@
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
$pptWasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = -1
$pres = $pp.Presentations.Add(-1)
try {
    [void]$pres.Slides.Add(1, 12)
    # Jede "mso:"-ID in den Tags des Ribbon-XML muss für ExecuteMso gültig sein (GetLabelMso wirft sonst).
    $xml = Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot '..\addin\ribbon.xml')
    $badMso = @()
    foreach ($id in ([regex]::Matches($xml, '(?<=\btag="[^"]*)mso:([A-Za-z0-9_]+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)) {
        try { "mso:$id -> " + $pp.CommandBars.GetLabelMso($id) } catch { $badMso += $id }
    }
    $api = $pp.COMAddIns.Item('rethink.addin').Object
    $status = $api.RibbonStatus()
    "Ribbon: $status"
    if ($status -eq 'geladen') { $api.ActivateTab() }
    $hwnd = [IntPtr]$pp.HWND
    [void][RibbonCheckWin]::ShowWindow($hwnd, 3)
    [void][RibbonCheckWin]::SetForegroundWindow($hwnd)
    Start-Sleep -Seconds 2
    $r = New-Object RibbonCheckWin+RECT
    [void][RibbonCheckWin]::GetWindowRect($hwnd, [ref]$r)
    $bmp = New-Object Drawing.Bitmap ($r.Right - $r.Left), 230
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
    $bmp.Save($Out)
    $g.Dispose(); $bmp.Dispose()
    $Out
    if ($status -ne 'geladen') { throw 'PowerPoint hat das Ribbon-XML nicht angenommen' }
    if ($badMso) { throw "Ungültige idMso in den Tags von ribbon.xml (GetLabelMso schlägt fehl): $($badMso -join ', ')" }
} finally {
    $pres.Saved = -1; $pres.Close()
    if (-not $pptWasRunning) { $pp.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
}
