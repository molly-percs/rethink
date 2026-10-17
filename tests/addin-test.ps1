# Prüft das registrierte Add-in in PowerPoint über COM-Automation
# (Application.COMAddIns("rethink.addin").Object ruft dieselben Befehle auf wie das Ribbon, nur ohne Dialoge):
#   1. Diagramm einfügen: Chart1/Chart2, natives gruppiertes Säulendiagramm, Alternativtext rethink:<Name>
#   2. Elementname setzen: Alternativtext rethink:<Name>, Mehrfachauswahl, Entfernen, Fehler ohne Auswahl
#   3. Aus .ppttc erzeugen: die so vorbereitete Vorlage wird mit rethink::build gefüllt und geöffnet
# Eigene Präsentationen werden wieder geschlossen; andere geöffnete Präsentationen bleiben unberührt.
param([string]$OutDir = (Join-Path $PSScriptRoot 'out\addin'))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path

$script:failures = 0
function Check([bool]$ok, [string]$what) {
    if ($ok) { "ok      $what" } else { "FEHLER  $what"; $script:failures++ }
}
# Das Automationsobjekt hat keine Typbibliothek; Aufrufe gehen direkt über IDispatch.
function Call($obj, [string]$name, [object[]]$arguments = @()) {
    [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments)
}
function CallError($obj, [string]$name, [object[]]$arguments = @()) {
    try { Call $obj $name $arguments | Out-Null; return '' } catch { return $_.Exception.InnerException.Message + $_.Exception.Message }
}

$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$pp = New-Object -ComObject PowerPoint.Application
$own = @()
try {
    $addin = $pp.COMAddIns | Where-Object ProgId -eq 'rethink.addin'
    if (-not $addin) { $pp.COMAddIns.Update(); $addin = $pp.COMAddIns | Where-Object ProgId -eq 'rethink.addin' }
    Check ([bool]$addin) 'Add-in ist in PowerPoint angemeldet'
    if (-not $addin) { throw 'rethink.addin nicht gefunden - scripts\register.ps1 ausführen' }
    if (-not $addin.Connect) { $addin.Connect = $true }
    Check ([bool]$addin.Connect) 'Add-in ist geladen (Connect)'
    $rt = $addin.Object
    Check ($null -ne $rt) 'COMAddIn.Object ist gesetzt'
    "        Version: $(Call $rt 'Version')"

    # --- 1. Diagramm einfügen ---
    $pres = $pp.Presentations.Add(-1); $own += $pres
    $pres.Windows.Item(1).Activate()
    $slide = $pres.Slides.Add(1, 11)                 # ppLayoutTitleOnly
    $title = $slide.Shapes.Item(1)
    $n1 = Call $rt 'InsertChart'
    $n2 = Call $rt 'InsertChart'
    Check ($n1 -eq 'Chart1' -and $n2 -eq 'Chart2') "Diagramm einfügen vergibt Chart1, Chart2 ($n1, $n2)"
    $c1 = $slide.Shapes.Item('Chart1'); $c2 = $slide.Shapes.Item('Chart2')
    Check ($c1.HasChart -eq -1 -and $c1.Chart.ChartType -eq 51) 'Chart1 ist natives gruppiertes Säulendiagramm (xlColumnClustered)'
    Check ($c1.AlternativeText -eq 'rethink:Chart1') "Alternativtext von Chart1: '$($c1.AlternativeText)'"
    $setup = $pres.PageSetup
    Check ([math]::Abs($c1.Left + $c1.Width / 2 - $setup.SlideWidth / 2) -lt 1) 'Diagramm ist waagrecht zentriert'
    Check (($pp.Windows | Measure-Object).Count -ge 1) 'Datenblatt geschlossen, Präsentationsfenster bleibt'

    # --- 2. Elementname setzen ---
    $title.Select()
    Check ((Call $rt 'SetName' @('  Titel  ')) -eq 1) 'Elementname für eine Form gesetzt'
    Check ($title.AlternativeText -eq 'rethink:Titel') "Alternativtext getrimmt: '$($title.AlternativeText)'"
    Check ((Call $rt 'GetName') -eq 'Titel') 'GetName liest den Namen zurück'
    $c2.Select()
    Call $rt 'SetName' @('Rechts') | Out-Null
    Check ($c2.AlternativeText -eq 'rethink:Rechts' -and $c2.Name -eq 'Chart2') 'Diagramm umbenannt, Formname bleibt'

    $a = $slide.Shapes.AddTextbox(1, 10, 10, 100, 20); $b = $slide.Shapes.AddTextbox(1, 10, 40, 100, 20)
    $a.AlternativeText = 'Logo der Firma'
    $a.Select(); $b.Select(0)                         # 0 = msoFalse: zur Auswahl hinzufügen
    Check ((Call $rt 'SetName' @('Doppelt')) -eq 2) 'Mehrfachauswahl: beide Formen benannt'
    Check ($a.AlternativeText -eq 'rethink:Doppelt' -and $b.AlternativeText -eq 'rethink:Doppelt') 'beide Alternativtexte gesetzt'
    Check ((Call $rt 'SetName' @('')) -eq 2 -and $a.AlternativeText -eq '' -and $b.AlternativeText -eq '') 'leerer Name entfernt den rethink-Namen'
    $a.Delete(); $b.Delete()

    $pp.ActiveWindow.Selection.Unselect()
    $err = CallError $rt 'SetName' @('X')
    Check ($err -match 'Formen auswählen') "ohne Auswahl: verständliche Fehlermeldung ('$err')"

    $template = Join-Path $OutDir 'template.pptx'
    $pres.SaveAs($template)
    $pres.Close(); $own = @()

    # --- 3. Aus .ppttc erzeugen ---
    $ppttc = Join-Path $OutDir 'addin-test.ppttc'
    @'
[
  { "template": "template.pptx", "data": [
    { "name": "Titel",  "table": [[{ "string": "Aus dem Add-in erzeugt" }]] },
    { "name": "Chart1", "table": [[null, { "string": "2024" }, { "string": "2025" }],
                                  [{ "string": "Umsatz" }, { "number": 10 }, { "number": 20 }]] },
    { "name": "Rechts", "table": [[null, { "string": "A" }, { "string": "B" }, { "string": "C" }],
                                  [{ "string": "Reihe 1" }, { "number": 1 }, { "number": 2 }, { "number": 3 }],
                                  [{ "string": "Reihe 2" }, { "number": 4 }, { "number": 5 }, { "number": 6 }]] }
  ] }
]
'@ | Set-Content -LiteralPath $ppttc -Encoding utf8
    $expected = Join-Path $OutDir 'addin-test.pptx'
    if (Test-Path $expected) { Remove-Item -LiteralPath $expected }

    $opened = Call $rt 'BuildFrom' @($ppttc)
    Check ($opened -eq $expected) "BuildFrom erzeugt und öffnet $opened"
    $out = $pp.Presentations | Where-Object FullName -eq $opened; $own += $out
    $s = $out.Slides.Item(1)
    Check ($s.Shapes.Item(1).TextFrame.TextRange.Text -eq 'Aus dem Add-in erzeugt') 'Titel über Alternativtext rethink:Titel gefüllt'
    $ch1 = $s.Shapes.Item('Chart1').Chart
    Check ($ch1.SeriesCollection().Count -eq 1 -and $ch1.SeriesCollection(1).Name -eq 'Umsatz' -and
           (($ch1.SeriesCollection(1).Values | ForEach-Object { $_ }) -join ',') -eq '10,20') 'Chart1 über Formnamen gefüllt'
    $ch2 = $s.Shapes.Item('Chart2').Chart
    Check ($ch2.SeriesCollection().Count -eq 2 -and (($ch2.SeriesCollection(2).Values | ForEach-Object { $_ }) -join ',') -eq '4,5,6') 'Chart2 über Alternativtext rethink:Rechts gefüllt'
    Check ((Call $rt 'LastWarnings') -eq '') 'keine Warnungen'

    # Zieldatei ist geöffnet -> freier Name statt Überschreiben
    $second = Call $rt 'BuildFrom' @($ppttc)
    Check ($second -eq (Join-Path $OutDir 'addin-test (2).pptx')) "geöffnete Zieldatei wird nicht überschrieben: $(Split-Path $second -Leaf)"
    $own += ($pp.Presentations | Where-Object FullName -eq $second)

    $err = CallError $rt 'BuildFrom' @((Join-Path $OutDir 'gibtsnicht.ppttc'))
    Check ($err -match 'gibtsnicht') "fehlende Datei: Fehlermeldung ('$err')"
} finally {
    foreach ($p in $own) { try { $p.Saved = -1; $p.Close() } catch { } }
    $second = Join-Path $OutDir 'addin-test (2).pptx'
    if (Test-Path -LiteralPath $second) { Remove-Item -LiteralPath $second }
    if (-not $wasRunning) { $pp.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
}
if ($script:failures) { "FEHLGESCHLAGEN: $script:failures"; exit 1 } else { 'Alle Add-in-Prüfungen bestanden.' }
