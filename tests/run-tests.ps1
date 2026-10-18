# Prüft Validierung und Fehlermeldungen der CLI (ohne PowerPoint).
# Aufruf: tests\run-tests.ps1 -Exe <pfad\zu\rethink.exe>
param([Parameter(Mandatory)][string]$Exe)
$ErrorActionPreference = 'Stop'
$dir = Join-Path $PSScriptRoot 'out\cases'
New-Item -ItemType Directory -Force $dir | Out-Null
$template = (Resolve-Path (Join-Path $PSScriptRoot '..\samples\template.pptx')).Path -replace '\\', '/'
# Vorlage mit Platzhalter "Bridge" für gezeichnete Diagramme
& (Join-Path $PSScriptRoot 'make-waterfall-template.ps1') | Out-Null
$chartTemplate = (Resolve-Path (Join-Path $PSScriptRoot 'out\waterfall-template.pptx')).Path -replace '\\', '/'

$cases = @(
    @{ Name = 'kein-json';       Json = '{ das ist kein json';                                                   Exit = 2; Expect = 'Keine gültige JSON-Datei' }
    @{ Name = 'kein-array';      Json = '{}';                                                                    Exit = 2; Expect = 'JSON-Array' }
    @{ Name = 'ohne-template';   Json = '[{"data":[]}]';                                                         Exit = 2; Expect = '[0]: Schlüssel "template" fehlt' }
    @{ Name = 'zwei-typen';      Json = "[{`"template`":`"$template`",`"data`":[{`"name`":`"X`",`"table`":[[{`"string`":`"a`",`"number`":1}]]}]}]"; Exit = 2; Expect = '[0].data[0].table[0][0]: Zelle braucht genau einen' }
    @{ Name = 'datum';           Json = "[{`"template`":`"$template`",`"data`":[{`"name`":`"X`",`"table`":[[{`"date`":`"01.02.2024`"}]]}]}]"; Exit = 2; Expect = 'ISO 8601' }
    @{ Name = 'farbe';           Json = "[{`"template`":`"$template`",`"data`":[{`"name`":`"X`",`"table`":[[{`"number`":1,`"fill`":`"rgb(300,0,0)`"}]]}]}]"; Exit = 2; Expect = 'zwischen 0 und 255' }
    @{ Name = 'vorlage-fehlt';   Json = '[{"template":"gibtsnicht.pptx","data":[]}]';                            Exit = 2; Expect = 'Vorlage nicht gefunden' }
    @{ Name = 'unbekannter-name'; Json = "[{`"template`":`"$template`",`"data`":[{`"name`":`"Nirgends`",`"table`":[[{`"string`":`"x`"}]]}]}]"; Exit = 0; Expect = "Element 'Nirgends' wurde in der Vorlage nicht gefunden" }
    @{ Name = 'gantt-option';    Json = "[{`"template`":`"$chartTemplate`",`"data`":[{`"name`":`"Bridge`",`"chart`":{`"type`":`"gantt`",`"weekends`":`"bunt`"},`"table`":[[null]]}]}]"; Exit = 2; Expect = '"weekends" muss einer der Werte "plain", "shaded", "hidden" sein' }
    @{ Name = 'gantt-element';   Json = "[{`"template`":`"$chartTemplate`",`"data`":[{`"name`":`"Bridge`",`"chart`":{`"type`":`"gantt`",`"items`":[{`"kind`":`"bar`",`"row`":2,`"start`":`"2026-01-05`"}]},`"table`":[[null]]}]}]"; Exit = 2; Expect = 'chart.items[0]: "end" fehlt' }
)

$failed = 0
foreach ($case in $cases) {
    $file = Join-Path $dir "$($case.Name).ppttc"
    [IO.File]::WriteAllText($file, $case.Json, (New-Object Text.UTF8Encoding($false)))
    # Windows PowerShell 5.1 macht aus stderr-Zeilen nativer Programme Fehler; hier sind sie erwartete Ausgabe.
    $output = & { $ErrorActionPreference = 'Continue'; & $Exe build $file -o (Join-Path $dir "$($case.Name).pptx") 2>&1 | Out-String }
    $ok = ($LASTEXITCODE -eq $case.Exit) -and $output.Contains($case.Expect)
    if (-not $ok) { $failed++ }
    "{0,-5} {1,-17} exit={2} {3}" -f ($(if ($ok) { 'OK' } else { 'FEHL' })), $case.Name, $LASTEXITCODE, ($output.Trim() -split "`n")[0]
}
if ($failed) { throw "$failed Testfälle fehlgeschlagen" }
"Alle $($cases.Count) Testfälle bestanden."
