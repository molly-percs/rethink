# Ende-zu-Ende-Test des Wasserfalls im Add-in (registriertes Add-in nötig: scripts\register.ps1):
#  1. Wasserfall einfügen (Automation: COMAddIns("rethink.addin").Object.InsertWaterfall)
#  2. Datenblatt schreiben lassen, mit echtem Excel ändern (Wert + Zellfarbe) und speichern
#  3. Übernehmen (ApplyDatasheets) und prüfen, dass das Element neu gezeichnet wurde
#  4. Folie als PNG exportieren
param([string]$PngDir = (Join-Path $PSScriptRoot 'out\addin-waterfall'), [switch]$KeepOpen)
$ErrorActionPreference = 'Stop'
$failed = 0
function Expect($ok, $what) { if ($ok) { "ok    $what" } else { "FEHL  $what"; $script:failed++ } }

$pptWasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$xlWasRunning = [bool](Get-Process EXCEL -ErrorAction SilentlyContinue)
$pp = New-Object -ComObject PowerPoint.Application
$pres = $null
try {
    $pres = $pp.Presentations.Add(-1)
    $slide = $pres.Slides.Add(1, 12)   # leer
    $pp.ActiveWindow.View.GotoSlide(1)
    $addin = $pp.COMAddIns.Item('rethink.addin')
    Expect ($addin.Connect) 'Add-in geladen'
    $api = $addin.Object

    # 1. Einfügen
    $name = $api.InsertWaterfall($false)
    $group = $slide.Shapes.Item($name)
    $model = $group.Tags.Item('RETHINK_MODEL')
    Expect ($group.Type -eq 6) "Element '$name' ist eine Gruppe"
    Expect ($model -match '"waterfall"' -and $model -match '"datatable":true') 'Modell-Tag mit Typ und Datentabelle'
    Expect ($group.AlternativeText -eq "rethink:$name") 'Alternativtext rethink:<Name>'
    $texts = @($group.GroupItems | Where-Object { $_.HasTextFrame -eq -1 } | ForEach-Object { $_.TextFrame.TextRange.Text })
    Expect ($texts -contains '105') "berechnete Zwischensumme 105 sichtbar (Texte: $($texts -join ' '))"

    # 2. Datenblatt
    $group.Select()
    $file = $api.OpenDatasheet($false)
    Expect (Test-Path $file) "Datenblatt geschrieben: $file"
    $xl = New-Object -ComObject Excel.Application
    try {
        $book = $xl.Workbooks.Open($file)
        $sheet = $book.Worksheets.Item(1)
        Expect ($sheet.Name -eq 'Datenblatt') 'Blattname "Datenblatt"'
        Expect ($sheet.Range('E2').Value2 -eq 'e') 'e steht im Datenblatt'
        $sheet.Range('C2').Value2 = 55               # Umsatz 40 → 55
        $sheet.Range('D2').Interior.Color = 192      # RGB(192,0,0) als BGR
        $book.Save()
        $book.Close($false)
    } finally {
        if (-not $xlWasRunning) { $xl.Quit() }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }

    # 3. Übernehmen
    $redrawn = $api.ApplyDatasheets()
    Expect ($redrawn -eq 1) "ApplyDatasheets zeichnet 1 Element neu (geliefert: $redrawn)"
    $group = $slide.Shapes.Item($name)
    $model = $group.Tags.Item('RETHINK_MODEL')
    Expect ($model -match '"number":55') 'Modell enthält den neuen Wert 55'
    Expect ($model -match '"fill":"#C00000"') 'Modell enthält die Zellfarbe aus Excel'
    $texts = @($group.GroupItems | Where-Object { $_.HasTextFrame -eq -1 } | ForEach-Object { $_.TextFrame.TextRange.Text })
    Expect ($texts -contains '120') "neue Zwischensumme 120 sichtbar"
    $red = @($group.GroupItems | Where-Object { $_.Type -eq 1 -and $_.Fill.ForeColor.RGB -eq 192 })   # 1 = AutoShape (Segment)
    Expect ($red.Count -eq 1) 'Kosten-Segment ist rot'
    Expect ($api.ApplyDatasheets() -eq 0) 'ohne neue Änderung kein erneutes Zeichnen'

    # 3b. Wie im echten Ablauf: in Excel speichern, rethink bemerkt es selbst (sekündliche Prüfung)
    $xl = New-Object -ComObject Excel.Application
    try {
        $book = $xl.Workbooks.Open($file)
        $book.Worksheets.Item(1).Range('B2').Value2 = 90   # Start 100 → 90
        $book.Save()
        $book.Close($false)
    } finally {
        if (-not $xlWasRunning) { $xl.Quit() }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    $deadline = (Get-Date).AddSeconds(8)
    do { Start-Sleep -Milliseconds 500; $model = $slide.Shapes.Item($name).Tags.Item('RETHINK_MODEL') } until ($model -match '"number":90' -or (Get-Date) -gt $deadline)
    Expect ($model -match '"number":90') 'Speichern in Excel zeichnet das Element ohne Anstoß neu'

    # 3c. Nach dem Neuzeichnen (neue Form-ID) dasselbe Datenblatt wieder öffnen, kein zweites
    $slide.Shapes.Item($name).Select()
    $again = $api.OpenDatasheet($false)
    Expect ($again -eq $file) 'erneutes Öffnen nach Neuzeichnen nutzt dieselbe Datei'
    Expect ($api.EventStatus() -match 'datenblaetter=1') 'genau ein überwachtes Datenblatt'

    # 4. Auswahl-Leuchten (Overlay-Fenster des Add-ins) und Ereignisse
    Add-Type -Namespace RethinkTest -Name Win -MemberDefinition @'
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindowW(string c, string t);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, int[] r);
'@ -ErrorAction SilentlyContinue
    function OverlayMatches($shape) {
        $h = [RethinkTest.Win]::FindWindowW('rethinkSelectionOverlay', $null)
        $r = New-Object int[] 4; [void][RethinkTest.Win]::GetWindowRect($h, $r)
        $w = $pp.ActiveWindow
        $l = $w.PointsToScreenPixelsX($shape.Left); $t = $w.PointsToScreenPixelsY($shape.Top)
        $rr = $w.PointsToScreenPixelsX($shape.Left + $shape.Width); $b = $w.PointsToScreenPixelsY($shape.Top + $shape.Height)
        # Leuchten 10 px rundum, Etikett 18 px oben
        return ([math]::Abs($r[0] - ($l - 10)) -le 1) -and ([math]::Abs($r[1] - ($t - 28)) -le 1) -and
               ([math]::Abs($r[2] - ($rr + 10)) -le 1) -and ([math]::Abs($r[3] - ($b + 10)) -le 1)
    }
    $group = $slide.Shapes.Item($name)
    $group.Select(); Start-Sleep -Milliseconds 800
    $overlay = [RethinkTest.Win]::FindWindowW('rethinkSelectionOverlay', $null)
    "      Ereignisse: $($api.EventStatus())"
    Expect ($overlay -ne [IntPtr]::Zero -and [RethinkTest.Win]::IsWindowVisible($overlay)) 'Leuchtrahmen sichtbar bei Auswahl'
    Expect (OverlayMatches $group) 'Leuchtrahmen liegt genau um das Element'
    $zoom = $pp.ActiveWindow.View.Zoom
    $pp.ActiveWindow.View.Zoom = [math]::Max(20, $zoom - 15); Start-Sleep -Milliseconds 600
    Expect (OverlayMatches $group) 'Leuchtrahmen folgt dem Zoom'
    $pp.ActiveWindow.View.Zoom = $zoom; Start-Sleep -Milliseconds 400
    $pp.ActiveWindow.Selection.Unselect(); Start-Sleep -Milliseconds 800
    Expect (-not [RethinkTest.Win]::IsWindowVisible($overlay)) 'Leuchtrahmen verschwindet bei Abwahl'
    $group.GroupItems.Item(1).Select(); Start-Sleep -Milliseconds 800
    Expect ([RethinkTest.Win]::IsWindowVisible($overlay)) 'Leuchtrahmen auch bei Auswahl einer Form in der Gruppe'
    $pp.ActiveWindow.Selection.Unselect(); Start-Sleep -Milliseconds 500

    # 5. Bild
    New-Item -ItemType Directory -Force $PngDir | Out-Null
    $slide.Export((Join-Path $PngDir 'slide1.png'), 'PNG', 1280, 720)

    # 6. Austausch mit der Kommandozeile
    #  a) Add-in-Element per .ppttc neu befüllen
    $deck = Join-Path $PngDir 'addin-deck.pptx'
    $pres.SaveCopyAs($deck)
    $ppttc = Join-Path $PngDir 'addin-interop.ppttc'
    $json = '[{"template":"addin-deck.pptx","data":[{"name":"' + $name + '","table":[[null,{"string":"A"},{"string":"B"},{"string":"Summe"}],[{"string":"Wert"},{"number":10},{"number":7},{"string":"e"}]]}]}]'
    [IO.File]::WriteAllText($ppttc, $json, (New-Object Text.UTF8Encoding($false)))
    $cliOut = Join-Path $PngDir 'addin-interop.pptx'
    & (Join-Path $PSScriptRoot '..\bin\rethink.exe') build $ppttc -o $cliOut | Out-Null
    Expect ($LASTEXITCODE -eq 0) 'Kommandozeile befüllt das im Add-in erzeugte Element'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($cliOut)
    $tagText = ($zip.Entries | Where-Object FullName -like 'ppt/tags/*.xml' | ForEach-Object { (New-Object IO.StreamReader($_.Open())).ReadToEnd() }) -join ''
    $zip.Dispose()
    Expect ($tagText -match '&quot;number&quot;:7' -and $tagText -match '&quot;datatable&quot;:true') 'neue Daten im Modell, gespeicherte Optionen bleiben'
    #  b) Von der Kommandozeile erzeugtes Element im Add-in: Leuchtrahmen und Datenblatt
    $cliDeck = Join-Path $PSScriptRoot 'out\waterfall.pptx'
    if (Test-Path $cliDeck) {
        $cliPres = $pp.Presentations.Open($cliDeck, -1, 0, -1)
        $pp.ActiveWindow.View.GotoSlide(1)
        $bridge = $cliPres.Slides.Item(1).Shapes.Item('Bridge')
        $bridge.Select(); Start-Sleep -Milliseconds 800
        Expect ([RethinkTest.Win]::IsWindowVisible([RethinkTest.Win]::FindWindowW('rethinkSelectionOverlay', $null))) 'Leuchtrahmen auch für Kommandozeilen-Element'
        $cliSheet = $api.OpenDatasheet($false)
        Expect (Test-Path $cliSheet) 'Datenblatt für Kommandozeilen-Element'
        $cliPres.Close()
    } else { "      (tests\out\waterfall.pptx fehlt – Teil b übersprungen)" }
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
