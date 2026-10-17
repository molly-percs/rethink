# Erzeugt aus einer .ppttc-Datei eine Präsentation mit rethink und öffnet sie in PowerPoint.
# Aufruf:  .\ausprobieren.ps1                    (nimmt samples\sample.ppttc)
#          .\ausprobieren.ps1 -Ppttc daten.ppttc
param([string]$Ppttc = (Join-Path $PSScriptRoot 'samples\sample.ppttc'))
$ErrorActionPreference = 'Stop'

$exe = Join-Path $PSScriptRoot 'bin\rethink.exe'
$source = (Resolve-Path $Ppttc).Path
$output = Join-Path $PSScriptRoot ("samples\out\{0}.pptx" -f [IO.Path]::GetFileNameWithoutExtension($source))

# Die Beispielvorlage wird bei Bedarf über PowerPoint erzeugt.
if ($source -eq (Join-Path $PSScriptRoot 'samples\sample.ppttc') -and -not (Test-Path (Join-Path $PSScriptRoot 'samples\template.pptx'))) {
    & (Join-Path $PSScriptRoot 'samples\make-template.ps1')
}

& $exe build $source -o $output
if ($LASTEXITCODE -ne 0) { throw "rethink ist mit Code $LASTEXITCODE abgebrochen (ist die Ausgabe noch in PowerPoint geöffnet?)" }

Invoke-Item $output
"Geöffnet: $output"
