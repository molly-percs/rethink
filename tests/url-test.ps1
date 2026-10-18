# Prüft Vorlagen per http-URL: Ein lokaler Webserver (nur localhost) liefert samples/template.pptx aus.
param([Parameter(Mandatory)][string]$Exe, [int]$Port = 18731)
$ErrorActionPreference = 'Stop'
$template = (Resolve-Path (Join-Path $PSScriptRoot '..\samples\template.pptx')).Path
$dir = Join-Path $PSScriptRoot 'out\url'
New-Item -ItemType Directory -Force $dir | Out-Null

$server = Start-Job -ArgumentList $template, $Port -ScriptBlock {
    param($file, $port)
    $listener = New-Object Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$port/")
    $listener.Start()
    try {
        $context = $listener.GetContext()          # genau eine Anfrage bedienen
        $bytes = [IO.File]::ReadAllBytes($file)
        $context.Response.ContentType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $context.Response.Close()
        "served $($context.Request.Url.AbsolutePath)"
    } finally { $listener.Stop() }
}
Start-Sleep 2

$ppttc = Join-Path $dir 'remote.ppttc'
$json = "[{`"template`":`"http://localhost:$Port/template.pptx`",`"data`":[{`"name`":`"SlideTitle`",`"table`":[[{`"string`":`"Vorlage per URL`"}]]}]}]"
[IO.File]::WriteAllText($ppttc, $json, (New-Object Text.UTF8Encoding($false)))
& $Exe build $ppttc -o (Join-Path $dir 'remote.pptx')
$exit = $LASTEXITCODE
$log = Receive-Job $server -Wait -AutoRemoveJob
"Server: $log"
if ($exit -ne 0 -or -not (Test-Path (Join-Path $dir 'remote.pptx'))) { throw "URL-Vorlage fehlgeschlagen (exit $exit)" }
"URL-Vorlage OK."
