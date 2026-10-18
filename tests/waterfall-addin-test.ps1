# Ende-zu-Ende-Test des Wasserfalls im Add-in (registriertes Add-in nötig: scripts\register.ps1).
# Automation über COMAddIns("rethink.addin").Object; echte Mausklicks auf das Overlay sind so nicht möglich –
# die Griffe werden über AdjustSegment ausgelöst und ihre Lage über OverlayInfo geprüft.
param([string]$OutDir = (Join-Path $PSScriptRoot 'out\addin-waterfall'), [switch]$KeepOpen)
$ErrorActionPreference = 'Stop'
$failed = 0
function Expect($ok, $what) { if ($ok) { "ok    $what" } else { "FEHL  $what"; $script:failed++ } }
function Texts($group) { @($group.GroupItems | Where-Object { $_.HasTextFrame -eq -1 -and $_.Type -eq 17 } | ForEach-Object { $_.TextFrame.TextRange.Text }) }
function Ids($group) { (@($group.GroupItems | ForEach-Object { $_.Id }) -join ',') }
function WaitFor([scriptblock]$cond, [int]$seconds = 6) {
    $deadline = (Get-Date).AddSeconds($seconds)
    do { if (& $cond) { return $true }; Start-Sleep -Milliseconds 250 } until ((Get-Date) -gt $deadline)
    return $false
}

New-Item -ItemType Directory -Force $OutDir | Out-Null
$pptWasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$excelBefore = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
$pp = New-Object -ComObject PowerPoint.Application
$pres = $null
try {
    $pres = $pp.Presentations.Add(-1)
    $slide = $pres.Slides.Add(1, 12)
    $pp.ActiveWindow.View.GotoSlide(1)
    $addin = $pp.COMAddIns.Item('rethink.addin')
    Expect ($addin.Connect) 'Add-in geladen'
    $api = $addin.Object

    # 1. Einfügen – ohne Excel
    $name = $api.InsertWaterfall($false)
    $group = $slide.Shapes.Item($name)
    Expect ($group.Type -eq 6 -and $group.Tags.Item('RETHINK_MODEL') -match '"waterfall"') "Element '$name' mit Modell"
    Expect (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq $excelBefore) 'Einfügen startet kein Excel'
    Expect ((Texts $group) -contains '115') 'Zwischensumme 115'
    $roles = @($group.GroupItems | ForEach-Object { $_.Name })
    Expect ($roles -contains 'segment r1 c2' -and $roles -contains 'connector c1' -and $roles -contains 'table cell r1 c4') 'Formen heißen nach ihrer Rolle'

    # 2. Overlay: Griffe an Segmenten (nicht an e) + Datenblatt-Symbol
    $group.Select(); Start-Sleep -Milliseconds 700
    $info = $api.OverlayInfo()
    Expect ($info -match 'sichtbar=true') 'Overlay sichtbar'
    Expect ($info -match 'griffe=9') "4 Segmente × ▲▼ + Datenblatt-Symbol ($(([regex]::Match($info,'griffe=\d+')).Value))"
    Expect (-not ($info -match 'r1 c4 ') -and -not ($info -match 'r1 c6 ')) 'keine Griffe an e-Segmenten'
    $el = [regex]::Match($info, 'element=(-?\d+),(-?\d+),(-?\d+),(-?\d+)')
    $icon = [regex]::Match($info, '\[datenblatt r0 c0 (-?\d+),(-?\d+),(-?\d+),(-?\d+)\]')
    Expect ($icon.Success -and [int]$icon.Groups[1].Value -gt [int]$el.Groups[3].Value -and [int]$icon.Groups[2].Value -gt [int]$el.Groups[4].Value) 'Datenblatt-Symbol rechts unten außerhalb des Elements'

    # 3. ▲ an Ort und Stelle (ohne Datenblatt)
    $idsBefore = Ids $group
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $api.AdjustSegment(1, 2, 1, $false)   # Umsatz 40 → 41
    $sw.Stop()
    $group = $slide.Shapes.Item($name)
    Expect ($group.Tags.Item('RETHINK_MODEL') -match '"number":41') "▲ ändert den Wert 40 → 41 ($($sw.ElapsedMilliseconds) ms)"
    Expect ((Texts $group) -contains '116') 'Zwischensumme rechnet mit (116)'
    Expect ((Ids $group) -eq $idsBefore) 'Formen bleiben erhalten (Änderung an Ort und Stelle)'
    $api.AdjustSegment(1, 3, -1, $true)   # Kosten -25 → -35 (Umschalt = ×10)
    Expect ($slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL') -match '"number":-35') '▼ mit Umschalt: zehnfacher Schritt'
    $eBlocked = $false
    try { $api.AdjustSegment(1, 4, 1, $false) } catch { $eBlocked = $_.Exception.Message -match 'berechnet' }
    Expect $eBlocked 'e-Segment lässt sich nicht direkt ändern'
    # eigene Formatierung bleibt erhalten
    $seg = $slide.Shapes.Item($name).GroupItems.Item('segment r1 c1')
    $seg.Fill.ForeColor.RGB = 0x00B050
    $api.AdjustSegment(1, 1, 1, $false)
    Expect ($slide.Shapes.Item($name).GroupItems.Item('segment r1 c1').Fill.ForeColor.RGB -eq 0x00B050) 'eigene Segmentfarbe bleibt bei Wertänderung'

    # 3b. ▲▼ ziehen: dieselben Schritte wie mit der Maus (Maßstab eingefroren, Schritt nach Genauigkeit des Werts)
    $slide.Shapes.Item($name).Select(); Start-Sleep -Milliseconds 400
    $api.OpenDatasheet($false)
    $idsBefore = Ids $slide.Shapes.Item($name)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $r = $api.DragSegment(1, 2, '6;15;30')   # Umsatz 41, 30 px nach oben
    $sw.Stop()
    $ppu = [double]([regex]::Match($r, 'pixelProEinheit=([\d.]+)').Groups[1].Value)
    $expected = 41 + [math]::Round(30 / $ppu, [MidpointRounding]::AwayFromZero)
    $cell = $api.DatasheetCell(1, 2)
    Expect ($ppu -gt 0 -and $expected -gt 41 -and $cell -eq [string]$expected) "Ziehen um 30 px: 41 → $expected (Datenblatt: $cell, $r, $($sw.ElapsedMilliseconds) ms)"
    Expect ($slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL') -match ('"number":' + $expected + '[,}]')) '… Modell am Element gespeichert'
    Expect ((Ids $slide.Shapes.Item($name)) -eq $idsBefore) '… an Ort und Stelle'
    [void]$api.DragSegment(1, 2, '30;0')
    Expect ($api.DatasheetCell(1, 2) -eq [string]$expected) 'hin und zurück ziehen: Wert unverändert'
    $r = $api.DragSegment(1, 3, '-20')         # Kosten -35, 20 px nach unten: Balken wächst nach unten
    $ppu = [double]([regex]::Match($r, 'pixelProEinheit=([\d.]+)').Groups[1].Value)
    $down = -35 - [math]::Round(20 / $ppu, [MidpointRounding]::AwayFromZero)
    Expect ($down -lt -35 -and $api.DatasheetCell(1, 3) -eq [string]$down) "negatives Segment nach unten: -35 → $down"
    $eBlocked = $false
    try { [void]$api.DragSegment(1, 4, '20') } catch { $eBlocked = $_.Exception.Message -match 'berechnet' }
    Expect $eBlocked 'e-Segment lässt sich nicht ziehen'
    $api.CloseDatasheet()
    # ein Rückgängig nimmt das ganze Ziehen zurück (StartNewUndoEntry zu Beginn)
    $slide.Shapes.Item($name).Select()
    [void]$api.DragSegment(1, 2, '10;20;40')
    $pp.CommandBars.ExecuteMso('Undo'); Start-Sleep -Milliseconds 300
    Expect ($slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL') -match ('"number":' + $expected + '[,}]')) 'ein Rückgängig nimmt das ganze Ziehen zurück'
    # Werte für die folgenden Teile zurücksetzen (Umsatz 41, Kosten -35)
    $slide.Shapes.Item($name).Select(); $api.OpenDatasheet($false)
    $api.DatasheetEnter(1, 2, '41'); $api.DatasheetEnter(1, 3, '-35')
    $api.CloseDatasheet()

    # 4a. Eingebautes Datenblatt (Symbol unten rechts / Doppelklick)
    $slide.Shapes.Item($name).Select()
    $api.OpenDatasheet($false)
    Expect ($api.DatasheetCount() -eq 1) 'eingebautes Datenblatt geöffnet'
    Expect ($api.DatasheetCell(1, 2) -eq '41' -and $api.DatasheetCell(0, 2) -eq 'Umsatz' -and $api.DatasheetCell(1, 4) -eq 'e') 'Datenblatt zeigt Kategorien, Werte und e'
    $idsBefore = Ids $slide.Shapes.Item($name)
    $api.DatasheetEnter(1, 2, '45')
    Expect ($slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL') -match '"number":45') 'Eingabe aktualisiert das Diagramm sofort'
    Expect ((Ids $slide.Shapes.Item($name)) -eq $idsBefore) '… an Ort und Stelle'
    $api.DatasheetEnter(1, 3, '-12,5')
    Expect ($slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL') -match '"number":-12.5') 'Eingabe mit deutschem Komma (-12,5)'
    $api.DatasheetPaste(1, 5, "-8`t")
    Expect ($slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL') -match '"number":-8') 'Einfügen (Strg+V-Weg) übernimmt Werte'
    $slide.Shapes.Item($name).Select()
    $api.AdjustSegment(1, 2, 1, $false)
    $cell = $api.DatasheetCell(1, 2)
    # Schritt nach der Genauigkeit des Werts selbst: 45 → 46, obwohl das Diagramm wegen -12,5 eine Nachkommastelle zeigt
    Expect ($cell -eq '46') "▲ (Schritt 1) aktualisiert das offene Datenblatt (Zelle: $cell)"
    $api.CloseDatasheet()
    Expect ($api.DatasheetCount() -eq 0) 'Datenblatt geschlossen (nichts zu speichern)'
    # Werte für den Excel-Teil zurücksetzen (Umsatz 41)
    $slide.Shapes.Item($name).Select(); $api.OpenDatasheet($false)
    $api.DatasheetEnter(1, 2, '41'); $api.DatasheetEnter(1, 3, '-35'); $api.DatasheetEnter(1, 5, '-10'); $api.DatasheetEnter(1, 6, 'e')
    $api.CloseDatasheet()

    # 4b. Alternativ: Datenblatt per Excel-Automation (Kontextmenü „In Excel bearbeiten“): live ohne Speichern
    $slide.Shapes.Item($name).Select()
    $file = $api.OpenExcelDatasheet($false)
    $book = $api.DatasheetWorkbook()
    Expect ($book -ne $null) "Datenblatt in Excel geöffnet ($([IO.Path]::GetFileName($file)))"
    $sheet = $book.Worksheets.Item(1)
    Expect ($sheet.Range('C2').Value2 -eq 41) 'Datenblatt zeigt den per ▲ geänderten Wert'
    $sheet.Range('C2').Value2 = 55
    Expect (WaitFor { $slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL') -match '"number":55' }) 'Änderung in Excel kommt ohne Speichern an'
    # ▲ bei offenem Datenblatt schreibt in die Excel-Zelle
    $slide.Shapes.Item($name).Select()
    $api.AdjustSegment(1, 2, 1, $false)
    $xlValue = $sheet.Range('C2').Value2
    Expect ($xlValue -eq 56) "▲ bei offenem Datenblatt ändert die Excel-Zelle (Zelle: $xlValue)"
    Expect (WaitFor { $slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL') -match '"number":56' }) '… und das Diagramm'
    # Farbe (kein Änderungsereignis in Excel) wird beim Schließen übernommen
    $sheet.Range('D2').Interior.Color = 192
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    Expect (WaitFor { $slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL') -match '"fill":"#C00000"' }) 'Schließen übernimmt auch die Zellfarbe'
    Expect (WaitFor { $api.EventStatus() -match 'datenblaetter=0' }) 'Datenblatt nach dem Schließen abgemeldet'
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Expect (WaitFor { @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq $excelBefore } 10) 'eigene Excel-Instanz beendet'

    # 5. Größe ändern → Neu-Layout nach dem Loslassen
    $group = $slide.Shapes.Item($name)
    $group.Select(); Start-Sleep -Milliseconds 400
    $group.Width = $group.Width * 0.8
    $target = $group.Width
    Expect (WaitFor { $l = $slide.Shapes.Item($name).Tags.Item('RETHINK_LAYOUT'); $w = [double]($l.Split(';')[6]); [math]::Abs($w - $slide.Shapes.Item($name).Width) -lt 1 -and [math]::Abs($w - $target) -lt 8 }) 'Größenänderung zeichnet neu (Layout-Rahmen folgt)'

    # 6. Ausrichtung (Einfügemenü: Pfeile um das Symbol) und abbauender Wasserfall
    $slide2 = $pres.Slides.Add(2, 12)
    $pp.ActiveWindow.View.GotoSlide(2)
    $right = $api.InsertElement('waterfall', 'right', $false)
    $g2 = $slide2.Shapes.Item($right)
    Expect ($g2.Tags.Item('RETHINK_KIND') -eq 'waterfall right' -and $g2.Tags.Item('RETHINK_MODEL') -match '"orientation":"right"') 'Wasserfall nach rechts eingefügt'
    $s1 = $g2.GroupItems.Item('segment r1 c1'); $s2 = $g2.GroupItems.Item('segment r1 c2')
    Expect ($s2.Left -gt $s1.Left -and $s2.Top -gt $s1.Top -and $s1.Width -gt $s1.Height) 'Balken wachsen nach rechts, Kategorien untereinander'
    $g2.Select(); Start-Sleep -Milliseconds 600
    $info = $api.OverlayInfo()
    Expect ($info -match '\[auf r1 c2 ' -and $info -match '\[ab r1 c2 ') 'Griffe ◀▶ an den Balken'
    $r = $api.DragSegment(1, 2, '20;40')   # Umsatz 40, 40 px nach rechts = mehr
    $ppu = [double]([regex]::Match($r, 'pixelProEinheit=([\d.]+)').Groups[1].Value)
    $grown = 40 + [math]::Round(40 / $ppu, [MidpointRounding]::AwayFromZero)
    Expect ($grown -gt 40 -and $slide2.Shapes.Item($right).Tags.Item('RETHINK_MODEL') -match ('"number":' + $grown + '[,}]')) "waagerecht ziehen: 40 → $grown ($r)"
    $down = $api.InsertElement('waterfall-down', 'up', $false)
    $g3 = $slide2.Shapes.Item($down)
    $texts = @($g3.GroupItems | Where-Object { $_.HasTextFrame -eq -1 } | ForEach-Object { $_.TextFrame.TextRange.Text })
    Expect ($g3.Tags.Item('RETHINK_MODEL') -match '"string":"e"' -and $texts -contains '100') 'abbauender Wasserfall: Summe links (e rückwärts = 100)'
    $slide2.Export((Join-Path $OutDir 'slide2.png'), 'PNG', 1280, 720)

    # 7. Bild
    $slide.Export((Join-Path $OutDir 'slide1.png'), 'PNG', 1280, 720)
    if ($KeepOpen) { $slide.Shapes.Item($name).Select() }
} finally {
    if (-not $KeepOpen) {
        if ($pres) { $pres.Saved = -1; $pres.Close() }
        if (-not $pptWasRunning) { $pp.Quit() }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
}
if ($failed) { throw "$failed Prüfungen fehlgeschlagen" }
"Alle Prüfungen bestanden."
