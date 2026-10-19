# Ende-zu-Ende-Test des Zeitplans (Gantt) im Add-in (registriertes Add-in nötig: scripts\register.ps1).
# Automation über COMAddIns("rethink.addin").Object; Ziehen läuft über GanttDrag denselben Weg wie die Maus
# (Overlay-Gesten), Kontextmenübefehle über GanttCommand mit Folienkoordinaten statt Mausposition.
param([string]$OutDir = (Join-Path $PSScriptRoot 'out\addin-gantt'), [switch]$KeepOpen)
$ErrorActionPreference = 'Stop'
$failed = 0
function Expect($ok, $what) { if ($ok) { "ok    $what" } else { "FEHL  $what"; $script:failed++ } }
function Ids($group) { (@($group.GroupItems | ForEach-Object { $_.Id }) -join ',') }
function Roles($group) { @($group.GroupItems | ForEach-Object { $_.Name }) }
function D([string]$s) { [datetime]::ParseExact($s, 'yyyy-MM-dd', $null) }
# "range=a..b; scales=…; bar r2=a..b; milestone r3=a; …" → Hashtable
function Targets($api) {
    $t = @{}
    foreach ($part in ($api.GanttTargets() -split '; ')) {
        $kv = $part -split '=', 2
        $t[$kv[0]] = $kv[1]
    }
    $t
}
function StartOf($t, $role) { D (($t[$role] -split '\.\.')[0]) }
function EndOf($t, $role) { D (($t[$role] -split '\.\.')[1]) }
function Center($shape) { @(($shape.Left + $shape.Width / 2), ($shape.Top + $shape.Height / 2)) }

New-Item -ItemType Directory -Force $OutDir | Out-Null
$pptWasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$pp = New-Object -ComObject PowerPoint.Application
$pres = $null
try {
    $pres = $pp.Presentations.Add(-1)
    $slide = $pres.Slides.Add(1, 12)
    $pp.ActiveWindow.View.GotoSlide(1)
    $addin = $pp.COMAddIns.Item('rethink.addin')
    Expect ($addin.Connect) 'Add-in geladen'
    $api = $addin.Object
    Expect ($api.RibbonStatus() -eq 'geladen') 'Ribbon-XML (mit Kontextmenü „rethink-Zeitplan“) angenommen'

    # 1. Einfügen: aktueller und die zwei folgenden Monate
    $name = $api.InsertElement('gantt', 'up', $false)
    $group = $slide.Shapes.Item($name)
    Expect ($group.Type -eq 6 -and $group.Tags.Item('RETHINK_MODEL') -match '"gantt"' -and $group.Tags.Item('RETHINK_KIND') -eq 'gantt') "Zeitplan '$name' mit Modell"
    $roles = Roles $group
    Expect ($roles -contains 'bar r2' -and $roles -contains 'milestone r3' -and $roles -contains 'milestone r1' -and $roles -contains 'activity r2' -and $roles -contains 'headline activity') 'Formen heißen nach ihrer Rolle (Balken, Meilensteine, Beschriftungen)'
    $scheme = $slide.ThemeColorScheme
    Expect ($group.GroupItems.Item('bar r2').Fill.ForeColor.RGB -eq $scheme.Colors(5).RGB -and $group.GroupItems.Item('milestone r3').Fill.ForeColor.RGB -eq $scheme.Colors(6).RGB) 'Farben aus dem Master: Balken Akzent 1, Meilensteine Akzent 2'
    $t = Targets $api
    $first = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    $expectedRange = '{0:yyyy-MM-dd}..{1:yyyy-MM-dd}' -f $first, $first.AddMonths(3).AddDays(-1)
    Expect ($t['range'] -eq $expectedRange) "Bereich: aktueller und zwei folgende Monate ($($t['range']))"
    Expect ($t['scales'] -match 'Month') "automatische Skalen ($($t['scales']))"

    # 2. Overlay: Griffe an Balkenenden, Flächen zum Verschieben, Pfeile an den Skalenenden
    $group.Select(); Start-Sleep -Milliseconds 700
    $info = $api.OverlayInfo()
    Expect ($info -match 'sichtbar=true') 'Overlay sichtbar'
    Expect ($info -match '\[anfang bar r2 ' -and $info -match '\[ende bar r2 ' -and $info -match '\[verschieben bar r2 ') 'Balken: Griffe an Anfang und Ende, Fläche zum Verschieben'
    Expect ($info -match '\[verschieben milestone r3 ') 'Meilenstein verschiebbar'
    Expect ($info -match '\[bereich-anfang range ' -and $info -match '\[bereich-ende range ') 'Pfeile an beiden Skalenenden'
    Expect ($info -match '\[datenblatt ') 'Datenblatt-Symbol'
    Expect ($info -match '\[zeile activity r2 ' -and $info -match '\[zeile activity r3 ' -and $info -match '\[zeile activity r4 ') 'Zeilengriff „⋮⋮“ je Aktivitätszeile'
    Expect ($info -match '\[zelle r2 c0 ' -and $info -match '\[zelle r3 c0 ') 'Zellen zum Bearbeiten je Zeile und Beschriftungsspalte'
    $el = [regex]::Match($info, 'element=(-?\d+),(-?\d+),(-?\d+),(-?\d+)'); $gr = [regex]::Match($info, '\[zeile activity r2 (-?\d+),(-?\d+),(-?\d+),(-?\d+)\]')
    Expect ([int]$gr.Groups[3].Value -lt [int]$el.Groups[1].Value) 'Zeilengriff liegt links neben dem Diagramm (nicht auf PowerPoints Anfassern)'

    # Maßstab für Erwartungen: Folienpunkte je Bildschirmpixel und je Tag
    $win = $pp.ActiveWindow
    $ptPerPx = 1000 / ($win.PointsToScreenPixelsX(1000) - $win.PointsToScreenPixelsX(0))
    $geo = $group.Tags.Item('RETHINK_GANTT') -split ';' | ForEach-Object { [double]::Parse($_, [Globalization.CultureInfo]::InvariantCulture) }
    $plotW = $geo[1] - $geo[0]  # plotLeft;plotRight;… in Folienpunkten
    $days = ((D (($t['range'] -split '\.\.')[1])) - (D (($t['range'] -split '\.\.')[0]))).TotalDays + 1
    $ptPerDay = $plotW / $days
    function Px([double]$d) { [int][math]::Round($d * $ptPerDay / $ptPerPx) }  # Mausweg für d Tage bei der aktuellen Zoomstufe
    function Steps([double[]]$ds) { ($ds | ForEach-Object { Px $_ }) -join ';' }

    # 3. Ende ziehen: an Ort und Stelle, gerastet, Anfang bleibt
    $s0 = StartOf $t 'bar r2'; $e0 = EndOf $t 'bar r2'
    $ids = Ids $group
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $tips = $api.GanttDrag('bar r2', 1, (Steps 1, 3, 5.3), $false)
    $sw.Stop()
    $t = Targets $api
    $delta = ((EndOf $t 'bar r2') - $e0).TotalDays
    $expected = (Px 5.3) * $ptPerPx / $ptPerDay
    Expect ((StartOf $t 'bar r2') -eq $s0 -and $delta -gt 0 -and [math]::Abs($delta - $expected) -le 7) ("Ende ziehen um gut 5 Tage: +{0} Tage (erwartet ~{1:N1}, {2} ms)" -f $delta, $expected, $sw.ElapsedMilliseconds)
    Expect ($tips -match 'KW') "Kurzinfo beim Ziehen ($(($tips -split ' \| ')[-1]))"
    Expect ((Ids $slide.Shapes.Item($name)) -eq $ids) '… an Ort und Stelle (dieselben Formen)'
    Expect ($slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL') -match ('"date":"{0:yyyy-MM-dd}"' -f (EndOf $t 'bar r2'))) '… Modell am Element gespeichert'

    # 4. Verschieben erhält die Dauer
    $s1 = StartOf $t 'bar r2'; $e1 = EndOf $t 'bar r2'
    [void]$api.GanttDrag('bar r2', 2, (Steps -1, -1.7), $false)
    $t = Targets $api
    Expect ((StartOf $t 'bar r2') -lt $s1 -and ((EndOf $t 'bar r2') - (StartOf $t 'bar r2')) -eq ($e1 - $s1)) ('Verschieben: {0:dd.MM.} → {1:dd.MM.}, Dauer bleibt' -f $s1, (StartOf $t 'bar r2'))

    # 5. Esc bricht ab
    $before = $t['bar r2']
    [void]$api.GanttDrag('bar r2', 2, (Steps 3, 6), $true)
    $t = Targets $api
    Expect ($t['bar r2'] -eq $before) 'Esc beim Ziehen: Daten unverändert'

    # 6. Meilenstein ziehen
    $m0 = D $t['milestone r3']
    [void]$api.GanttDrag('milestone r3', 2, (Steps 2, 3.2), $false)
    $t = Targets $api
    Expect ((D $t['milestone r3']) -gt $m0) ('Meilenstein ziehen: {0:dd.MM.} → {1:dd.MM.}' -f $m0, (D $t['milestone r3']))

    # 7. Ein Rückgängig nimmt das ganze Ziehen zurück
    $before = $t['bar r2']
    $slide.Shapes.Item($name).Select()
    [void]$api.GanttDrag('bar r2', 1, (Steps 1, 3, 5.2), $false)
    $pp.CommandBars.ExecuteMso('Undo'); Start-Sleep -Milliseconds 400
    $slide.Shapes.Item($name).Select()
    Expect ((Targets $api)['bar r2'] -eq $before) 'ein Rückgängig nimmt das ganze Ziehen zurück'

    # 8. Datenblatt: zeigt gezogene Daten, Eingabe zeichnet neu
    $api.OpenDatasheet($false)
    Expect ($api.DatasheetCell(0, 1) -eq 'Start' -and $api.DatasheetCell(2, 0) -eq 'Analyse') 'Datenblatt: Überschriften und Aktivitäten'
    [void]$api.GanttDrag('bar r2', 1, (Steps 2.2), $false)
    $t = Targets $api
    Expect ($api.DatasheetCell(2, 2) -eq ('{0:dd.MM.yyyy}' -f (EndOf $t 'bar r2'))) "offenes Datenblatt folgt dem Ziehen ($($api.DatasheetCell(2, 2)))"
    $newEnd = (StartOf $t 'bar r2').AddDays(10)
    $api.DatasheetEnter(2, 2, ('{0:d.M.yy}' -f $newEnd))
    $t = Targets $api
    Expect ((EndOf $t 'bar r2') -eq $newEnd) ('Eingabe im Datenblatt (T.M.JJ) setzt das Ende ({0:dd.MM.yyyy})' -f $newEnd)
    $api.CloseDatasheet()

    # 9. Skalenende ziehen: Bereich in Einheiten der feinsten Skala
    $slide.Shapes.Item($name).Select()
    $range0 = (Targets $api)['range']
    [void]$api.GanttDrag('range', 1, (Steps 10, 20), $false)
    $t = Targets $api
    $rangeEnd = D (($t['range'] -split '\.\.')[1])
    Expect ($t['range'] -ne $range0 -and $rangeEnd -gt (D (($range0 -split '\.\.')[1]))) "Skalenende ziehen verlängert den Bereich ($range0 → $($t['range']))"

    # 10. Kontextmenübefehle an einer Mausposition (Folienkoordinaten)
    $group = $slide.Shapes.Item($name)
    $row3 = $group.GroupItems.Item('activity r3')
    $bar3 = $group.GroupItems.Item('bar r3')
    $y3 = $row3.Top + $row3.Height / 2
    $x3 = $bar3.Left + $bar3.Width / 2
    Expect ($api.GanttHit($x3, $y3) -match 'zeile=3 .*element=bar r3') "Treffer: Zeile 3, Balken ($($api.GanttHit($x3, $y3)))"
    $api.GanttCommand('rtkG_insertRow', $x3, $y3)
    $group = $slide.Shapes.Item($name)
    Expect ($group.GroupItems.Item('activity r4').TextFrame.TextRange.Text -eq 'Neue Aktivität' -and $group.GroupItems.Item('activity r5').TextFrame.TextRange.Text -eq 'Umsetzung') 'Zeile einfügen (unter der angeklickten)'
    $row4 = $group.GroupItems.Item('activity r4')
    $y4 = $row4.Top + $row4.Height / 2
    $api.GanttCommand('rtkG_newBar', $x3, $y4)
    $api.GanttCommand('rtkG_newMilestone', $x3 + 60, $y4)
    $t = Targets $api
    Expect ($t.ContainsKey('bar r4') -and $t.ContainsKey('milestone r4')) "Neuer Balken und Meilenstein in der leeren Zeile → Datenblatt ($($t['bar r4']), $($t['milestone r4']))"
    $row2 = $slide.Shapes.Item($name).GroupItems.Item('activity r2')
    $api.GanttCommand('rtkG_newChevron', $x3, $row2.Top + $row2.Height / 2)
    $group = $slide.Shapes.Item($name)
    Expect ((Roles $group) -contains 'item 0' -and $group.GroupItems.Item('item 0').AutoShapeType -eq 52) 'Neues Chevron als frei gesetztes Element'
    $c = Center $group.GroupItems.Item('item 0')
    $api.GanttCommand('rtkG_deleteItem', $c[0], $c[1])
    Expect (-not ((Roles $slide.Shapes.Item($name)) -contains 'item 0')) 'Element löschen'
    $api.GanttCommand('rtkG_rowUp', $x3, $y4)
    Expect ($slide.Shapes.Item($name).GroupItems.Item('activity r3').TextFrame.TextRange.Text -eq 'Neue Aktivität') 'Zeile nach oben'
    $api.GanttCommand('rtkG_deleteRow', $x3, $y3)
    $texts = @($slide.Shapes.Item($name).GroupItems | Where-Object { $_.HasTextFrame -eq -1 } | ForEach-Object { $_.TextFrame.TextRange.Text })
    Expect (-not ($texts -contains 'Neue Aktivität') -and ($texts -contains 'Konzept')) 'Zeile löschen'
    $api.GanttCommand('rtkG_fit')
    $t = Targets $api
    $starts = @($t.Keys | Where-Object { $_ -match '^(bar|milestone) r\d+$' } | ForEach-Object { D (($t[$_] -split '\.\.')[0]) })
    Expect ((D (($t['range'] -split '\.\.')[0])) -eq ($starts | Measure-Object -Minimum).Minimum) "Skala an Daten anpassen ($($t['range']))"

    # 11. Anzeige über die Umschalter des Kontextmenüs
    Expect ($api.GanttPressed('rtkT_separators')) 'Trennlinien sind an'
    $api.GanttToggle('rtkT_separators', $false)
    Expect (-not ((Roles $slide.Shapes.Item($name)) -match '^separator')) 'Trennlinien aus'
    $api.GanttToggle('rtkT_separators', $true)
    $api.GanttToggle('rtkT_barLabels', $true)
    Expect ((Roles $slide.Shapes.Item($name)) -contains 'bar r2 label') 'Datumsbeschriftung an Balken'
    $api.GanttToggle('rtkT_responsible', $true)
    Expect ((Roles $slide.Shapes.Item($name)) -contains 'headline responsible' -and $api.GanttPressed('rtkT_responsible')) 'Spalte „Verantwortlich“ eingeblendet'
    $api.GanttToggle('rtkT_scaleYear', $true)
    Expect ((Targets $api)['scales'] -match '^Year') ("Skala Jahr von Hand dazu (" + (Targets $api)['scales'] + ")")
    $api.GanttCommand('rtkG_autoScales')
    Expect (-not ((Targets $api)['scales'] -match 'Year')) 'Automatische Skalen wieder an'
    $api.GanttToggle('rtkT_weekendsHidden', $true)
    Expect ($api.GanttPressed('rtkT_weekendsHidden') -and -not $api.GanttPressed('rtkT_weekendsShaded')) 'Wochenenden ausgeblendet'
    $api.GanttToggle('rtkT_weekendsShaded', $true)

    # 12. Bereich wie im Kalender: zwei Wochen → Tagesskala, Wochenenden schattiert
    $bar = StartOf (Targets $api) 'bar r2'
    $api.SetGanttDates('', ('{0:yyyy-MM-dd}' -f $bar), ('{0:yyyy-MM-dd}' -f $bar.AddDays(13)))
    $roles = Roles $slide.Shapes.Item($name)
    Expect (($roles -contains 'scale day 1') -and ($roles -match '^weekend')) "zwei Wochen: Tagesskala und schattierte Wochenenden ($((Targets $api)['scales']))"
    $api.OpenCalendar('bar r2')
    $s0 = StartOf (Targets $api) 'bar r2'
    $sum = $api.CalendarPick(('{0:yyyy-MM-dd}' -f $s0), ('{0:yyyy-MM-dd}' -f $s0.AddDays(29)))
    Expect ($sum -eq ('{0:dd.MM.yyyy} bis {1:dd.MM.yyyy} · 4 Wochen 2 Tage' -f $s0, $s0.AddDays(29)) -and (EndOf (Targets $api) 'bar r2') -eq $s0.AddDays(30)) "Kalender: Zeitraum wählen setzt den Balken ($sum)"
    Expect ($api.CloseCalendar()) 'Kalender öffnet und schließt'
    $api.GanttCommand('rtkG_fit')

    # 13. Größe ändern → Neu-Layout
    $group = $slide.Shapes.Item($name)
    $group.Select(); Start-Sleep -Milliseconds 400
    $group.Width = $group.Width * 0.85
    $target = $group.Width
    $deadline = (Get-Date).AddSeconds(6); $ok = $false
    do { Start-Sleep -Milliseconds 250; $l = $slide.Shapes.Item($name).Tags.Item('RETHINK_LAYOUT'); $w = [double]($l.Split(';')[6]); $ok = [math]::Abs($w - $slide.Shapes.Item($name).Width) -lt 1 -and [math]::Abs($w - $target) -lt 8 } until ($ok -or (Get-Date) -gt $deadline)
    Expect $ok 'Größenänderung zeichnet neu'

    # 14. Eigenes Kontextmenü (Rechtsklick) und seine Einstellungen
    $group = $slide.Shapes.Item($name)
    $group.Select(); Start-Sleep -Milliseconds 300
    $c = Center $group.GroupItems.Item('bar r2')
    $menu = $api.GanttMenu($c[0], $c[1])
    Expect ($menu -match 'Form > Chevron=rowShape=chevron' -and $menu -match 'Verbinden mit \(Vorgänger\) > Konzept=link=Konzept') 'Rechtsklick auf Balken: Form und Verbinden mit …'
    Expect ($menu -match 'Skalen > Tag=scale:day=1' -and $menu -match 'Format der Skalen > Monat > ' -and $menu -match 'Spalten > Kategorie=column:category=1') 'Menü: Skalen, Skalenformate, Spalten'
    Expect ($menu -match 'Trennlinien > Senkrecht.*\[x\]=separators=0' -and $menu -match 'Trennlinien > Waagerecht.*=rowSeparators=1' -and $menu -match 'Wochenenden > Leicht grau \[x\]') 'Menü: Trennlinien senkrecht/waagerecht, Wochenenden'
    Expect ($menu -match 'Design > Standard \[x\]=design=standard' -and $menu -match 'Balkenbreite und Formen > Breit=barHeight=0.75') 'Menü: Design und Balkenbreite'
    $api.GanttCommand('rowShape=chevron', $c[0], $c[1])
    Expect ($slide.Shapes.Item($name).GroupItems.Item('bar r2').AutoShapeType -eq 52) 'Form per Rechtsklick: Balken wird Chevron'
    $api.GanttCommand('rowSeparators=1')
    Expect ((Roles $slide.Shapes.Item($name)) -contains 'row line r3') 'waagerechte Trennlinien an'
    $api.GanttCommand('weekends=plain')
    Expect (-not ((Roles $slide.Shapes.Item($name)) -match '^weekend')) 'Wochenenden normal (ohne Grau)'
    $api.GanttCommand('weekends=shaded')
    $h0 = $slide.Shapes.Item($name).GroupItems.Item('bar r3').Height
    $api.GanttCommand('barHeight=0.90')
    Expect ($slide.Shapes.Item($name).GroupItems.Item('bar r3').Height -gt $h0) 'Balkenbreite einstellbar'
    $api.GanttCommand('format:month=MMM')
    $months = @($slide.Shapes.Item($name).GroupItems | Where-Object { $_.Name -match '^scale month \d+$' } | ForEach-Object { $_.TextFrame.TextRange.Text })
    Expect ($months.Count -gt 0 -and ($months | Where-Object { $_.Length -gt 4 }).Count -eq 0) "Format der Monatsskala: $($months -join ' ')"
    # Kategorien im Datenblatt eintragen, dann nach Kategorie sortieren (gewählte Reihenfolge)
    $api.OpenDatasheet($false)
    $api.DatasheetEnter(2, 6, 'Planung'); $api.DatasheetEnter(3, 6, 'Bau'); $api.DatasheetEnter(4, 6, 'Planung'); $api.DatasheetEnter(5, 6, 'Bau')
    $roles = Roles $slide.Shapes.Item($name)
    Expect ($roles -contains 'headline category' -and $roles -contains 'category label r2') 'Kategoriespalte erscheint mit den Einträgen'
    $api.GanttCommand('sortCategories=Bau|Planung')
    $g = $slide.Shapes.Item($name)
    Expect ($g.GroupItems.Item('activity r2').TextFrame.TextRange.Text -eq 'Konzept' -and $g.GroupItems.Item('activity r4').TextFrame.TextRange.Text -eq 'Analyse') 'nach Kategorie sortiert: Bau zuerst, Reihenfolge innerhalb bleibt'
    Expect ($api.DatasheetCell(2, 0) -eq 'Konzept' -and $api.DatasheetCell(2, 6) -eq 'Bau') 'offenes Datenblatt zeigt die sortierte Tabelle'
    # Zeile im Diagramm zwei Zeilen nach unten ziehen: landet in „Planung“, Datenblatt folgt
    $rowPx = [int][math]::Round(($g.GroupItems.Item('activity r3').Top - $g.GroupItems.Item('activity r2').Top) / $ptPerPx)
    $g.Select()
    $tips = $api.GanttDrag('row r2', 2, "$([int]($rowPx / 2));$(2 * $rowPx)", $false)
    Expect ($tips -match 'Zeile verschieben' -and $tips -match '→ nach Planung') "Kurzinfo beim Ziehen einer Zeile: Ziel-Kategorie ($tips)"
    $g = $slide.Shapes.Item($name)
    Expect ($g.GroupItems.Item('activity r4').TextFrame.TextRange.Text -eq 'Konzept') 'Zeile per Ziehen verschoben'
    Expect ($api.DatasheetCell(4, 0) -eq 'Konzept' -and $api.DatasheetCell(4, 6) -eq 'Planung') "Zeile übernimmt die Kategorie der Ablagestelle ($($api.DatasheetCell(4, 6)))"
    # Zeile im Datenblatt einfügen/löschen: frei gesetzte Elemente bleiben bei ihrer Zeile
    $row3 = $g.GroupItems.Item('activity r3'); $label3 = $row3.TextFrame.TextRange.Text
    $api.GanttCommand('new=chevron', (Center $g.GroupItems.Item('bar r3'))[0], $row3.Top + $row3.Height / 2)
    $item = @(Roles $slide.Shapes.Item($name) | Where-Object { $_ -match '^item \d+$' })[-1]
    function SameRow($r) { $gg = $slide.Shapes.Item($name); [math]::Abs((Center $gg.GroupItems.Item($item))[1] - (Center $gg.GroupItems.Item("activity r$r"))[1]) -lt 3 }
    $api.DatasheetMenu(3, 'insertRow')
    Expect ((SameRow 4) -and $api.DatasheetCell(4, 0) -eq $label3) 'Zeile im Datenblatt eingefügt: frei gesetztes Element rückt mit'
    $api.DatasheetMenu(3, 'deleteRow')
    Expect ((SameRow 3) -and $api.DatasheetCell(3, 0) -eq $label3) 'Zeile im Datenblatt gelöscht: Element rückt zurück'
    $ic = Center $slide.Shapes.Item($name).GroupItems.Item($item); $api.GanttCommand('deleteItem', $ic[0], $ic[1])
    $api.CloseDatasheet()
    $api.GanttCommand('design=modern')
    $roles = Roles $slide.Shapes.Item($name)
    Expect ($roles -contains 'header band' -and $slide.Shapes.Item($name).GroupItems.Item('category r2').Type -eq 1) 'Design Modern: Kopfband, Kategorie-Blöcke'
    $row = $slide.Shapes.Item($name).GroupItems.Item('bar r3'); $c = Center $row
    $api.GanttCommand('link=Konzept', $c[0], $c[1])
    Expect ((Roles $slide.Shapes.Item($name)) -match '^link r3 1 arrow$') 'Verbinden mit Vorgänger per Rechtsklick'
    $api.GanttCommand('design=standard')

    # Minileiste per Rechtsklick: Balken (Form, Füllung, Linie, Stärke), Farben aus dem Design der Folie
    $c = Center $slide.Shapes.Item($name).GroupItems.Item('bar r3')
    $bar = $api.GanttToolbar($c[0], $c[1])
    Expect ($bar -match '\[Balken\]' -and $bar -match '\[Füllung #' -and $bar -match 'Füllung > Akzent 2=fill=accent2' -and $bar -match 'Balken > Gestrichelter Balken=barStyle=dashed' -and $bar -match 'Füllung > Automatisch \(Master\)( \[x\])?=fill=') 'Minileiste am Balken: Form, Füllung mit Designfarben'
    $accent2 = $slide.ThemeColorScheme.Colors(6).RGB
    $api.GanttCommand('fill=accent2', $c[0], $c[1])
    $api.GanttCommand('lineWeight=2.25', $c[0], $c[1])
    $s = $slide.Shapes.Item($name).GroupItems.Item('bar r3')
    Expect ($s.Fill.ForeColor.RGB -eq $accent2 -and $s.Line.Visible -and [math]::Abs($s.Line.Weight - 2.25) -lt 0.01) ("Füllung Akzent 2 und Linienstärke 2,25 pt (Füllung $($s.Fill.ForeColor.RGB) statt $accent2, Linie $($s.Line.Visible) $($s.Line.Weight), Stil $($s.Tags.Item('RETHINK_STYLE')), " + [regex]::Match($slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL'), '"rowStyles":[^\]]*\]').Value + ')')
    $api.GanttCommand('barStyle=dashed', $c[0], $c[1])
    $s = $slide.Shapes.Item($name).GroupItems.Item('bar r3')
    Expect (-not $s.Fill.Visible -and $s.Line.DashStyle -ne 1) 'Gestrichelter Balken: nur gestrichelter Umriss'
    $api.GanttCommand('barStyle=solid', $c[0], $c[1])
    Expect ($slide.Shapes.Item($name).GroupItems.Item('bar r3').Fill.ForeColor.RGB -eq $accent2) 'wieder durchgezogen, Füllung bleibt'
    $api.GanttCommand('rowUp', $c[0], $c[1])
    Expect ($slide.Shapes.Item($name).GroupItems.Item('bar r2').Fill.ForeColor.RGB -eq $accent2) 'Aussehen wandert mit der Zeile'
    $ms = @(Roles $slide.Shapes.Item($name) | Where-Object { $_ -match '^milestone r\d+$' })[0]
    $m = Center $slide.Shapes.Item($name).GroupItems.Item($ms)
    Expect (($api.GanttToolbar($m[0], $m[1])) -match 'Datum > Datum zeigen') 'Minileiste am Meilenstein: Form, Farben, Datum'
    $sc = Center $slide.Shapes.Item($name).GroupItems.Item(@(Roles $slide.Shapes.Item($name) | Where-Object { $_ -match '^scale ' })[0])
    Expect (($api.GanttToolbar($sc[0], $sc[1])) -match '\[Skalen\]' -and ($api.GanttToolbar($sc[0], $sc[1])) -match 'Format > Monat > ') 'Minileiste an der Skala: Skalen und Format'

    # Rundmenü: Rechtsklick auf eine freie Stelle im Zeitbereich, neues Element an genau dieser Stelle
    $api.GanttCommand('insertRow')
    $g = $slide.Shapes.Item($name)
    $last = @(Roles $g | Where-Object { $_ -match '^activity r\d+$' } | ForEach-Object { [int]($_ -replace '\D', '') } | Sort-Object)[-1]
    $geo = $g.Tags.Item('RETHINK_GANTT') -split ';' | ForEach-Object { [double]::Parse($_, [Globalization.CultureInfo]::InvariantCulture) }
    $lx = ($geo[0] + $geo[1]) / 2; $ly = (Center $g.GroupItems.Item("activity r$last"))[1]
    if ($api.GanttHit($lx, $ly) -match 'trenndatum=\d') { $lx += 11 }  # nicht auf einer Skalengrenze (dort käme das Menü)
    $radial = $api.GanttRightClick($lx, $ly, '')
    Expect ($radial -match 'Balken=new=bar' -and $radial -match 'Meilenstein=new=milestone' -and $radial -notmatch '\(aus\)') 'Rundmenü an einer freien Stelle der Zeile'
    [void]$api.GanttRightClick($lx, $ly, 'new=bar')
    $t = Targets $api
    Expect ($t.ContainsKey("bar r$last")) "Rundmenü: neuer Balken an der Stelle ($($t["bar r$last"]))"
    $bc = Center $slide.Shapes.Item($name).GroupItems.Item("bar r$last")
    Expect ($api.GanttRightClick($bc[0], $bc[1], '') -eq '') 'auf einem Balken kein Rundmenü (dort Minileiste und Menü)'

    # Einzelne Trennlinien: Hover an der Grenze, Rechtsklick-Menü, Befehle, Wandern mit der Zeile
    $g = $slide.Shapes.Item($name); $g.Select(); Start-Sleep -Milliseconds 500
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $tok = $g.Tags.Item('RETHINK_GANTT') -split ';'
    $lay = $g.Tags.Item('RETHINK_LAYOUT') -split ';' | ForEach-Object { [double]::Parse($_, $inv) }
    $dx = $g.Left - $lay[4]; $dy = $g.Top - $lay[5]
    $rowsTop = [double]::Parse($tok[4], $inv) + $dy; $rowH = [double]::Parse($tok[5], $inv); $labelsLeft = [double]::Parse($tok[7], $inv) + $dx
    $k = 9 + 3 * [int]$tok[8]; $edgeCount = [int]$tok[$k + 1]
    $edges = @(0..($edgeCount - 1) | ForEach-Object { [pscustomobject]@{ X = [double]::Parse($tok[$k + 2 + 2 * $_], $inv) + $dx; Date = ([datetime]'1970-01-01').AddDays([double]::Parse($tok[$k + 3 + 2 * $_], $inv)).ToString('yyyy-MM-dd') } })
    Expect ($edgeCount -ge 4) "Grenzen der feinsten Skala am Element vermerkt ($edgeCount)"
    $yB = $rowsTop + 2 * $rowH; $xl = $labelsLeft + 8   # Grenze zwischen Tabellenzeile 3 und 4, im Beschriftungsbereich
    $hov = $api.GanttHover($xl, $yB)
    Expect ($hov -match '^waagerecht zeile=4 band=') "Hover zwischen zwei Zeilen: waagerechtes Band ($hov)"
    $hc = $api.GanttHover($xl, $rowsTop + $rowH / 2)   # Zeilenmitte der ersten Aktivitätszeile (fern von PowerPoints Anfassern)
    Expect ($hc -notmatch 'waagerecht' -and $hc -match 'griff=zelle r2 c\d') "Hover in der Zeilenmitte: kein Band, dafür leuchtet die Zelle ($hc)"
    Expect ($api.GanttHover($g.Left - 19 * $ptPerPx, $yB + $rowH / 2) -match 'griff=zeile activity r4') 'Hover am Zeilengriff'
    $e = $edges[2]; $yMid = $rowsTop + $rowH / 2
    Expect ($api.GanttHover($e.X, $yMid) -match "^senkrecht datum=$($e.Date) band=") "Hover an einer Skalengrenze: senkrechtes Band ($($e.Date))"
    $hit = $api.GanttHit($e.X, $yMid)
    Expect ($hit -match "trenndatum=$($e.Date)" -and $api.GanttHit($xl, $yB) -match 'trennzeile=4') "Treffer enthält die Grenze ($hit)"
    $far = $api.GanttHover($e.X, $rowsTop - 6 * $ptPerPx)   # über den Zeilen (im Skalenband) gehört die Grenze nicht dazu
    Expect ($far -notmatch 'senkrecht') "im Skalenband kein Band ($far)"
    $menu = $api.GanttMenu($xl, $yB)
    Expect ($menu -match 'Waagerechte Trennlinie hier einfügen=rowLine=4' -and $menu -notmatch 'Senkrechte Trennlinie') 'Menü an der Zeilengrenze: Waagerechte Trennlinie hier einfügen'
    Expect ($api.GanttMenu($e.X, $yMid) -match "Senkrechte Trennlinie hier einfügen=dateLine=$($e.Date)") 'Menü an der Skalengrenze: Senkrechte Trennlinie hier einfügen'
    Expect ($api.GanttRightClick($e.X + 8 * $ptPerPx, $yB, '') -eq '') 'an der Grenze kein Rundmenü'
    $api.GanttCommand('rowLine=4')
    $g = $slide.Shapes.Item($name)
    Expect ((Roles $g) -contains 'line row r4' -and [math]::Abs($g.GroupItems.Item('line row r4').Top - $yB) -lt 1.5 -and $g.GroupItems.Item('line row r4').Line.Weight -eq 0.75) 'Befehl rowLine=4: Linie „line row r4“ genau an der Grenze'
    Expect ($api.GanttMenu($xl, $yB) -match 'Waagerechte Trennlinie hier entfernen=rowLine=4') 'Menü bietet jetzt „entfernen“ an'
    $api.GanttCommand('rowUp', $xl, $yB + $rowH / 2)
    Expect (((Roles $slide.Shapes.Item($name)) -contains 'line row r3') -and -not ((Roles $slide.Shapes.Item($name)) -contains 'line row r4')) 'Zeile nach oben: die Linie geht mit ihrer Zeile'
    $api.GanttCommand('rowLine=3')
    Expect (-not ((Roles $slide.Shapes.Item($name)) -match '^line row')) 'Befehl noch einmal: Linie entfernt'
    $api.GanttCommand("dateLine=$($e.Date)")
    $g = $slide.Shapes.Item($name)
    Expect ((Roles $g) -contains "line date $($e.Date)" -and [math]::Abs($g.GroupItems.Item("line date $($e.Date)").Left - $e.X) -lt 1.5) "Befehl dateLine: Linie „line date $($e.Date)“ an der Skalengrenze"
    $api.GanttCommand("dateLine=$($e.Date)")
    Expect (-not ((Roles $slide.Shapes.Item($name)) -match '^line date')) 'Datumslinie noch einmal: entfernt'
    $slide.Shapes.Item($name).Select()

    # 15. Bild
    $slide.Export((Join-Path $OutDir 'slide1.png'), 'PNG', 1600, 900)
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
