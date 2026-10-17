# Registriert das rethink-Add-in für den aktuellen Benutzer (nur HKCU, keine Administratorrechte).
#
#   HKCU\Software\Classes\CLSID\{68EFFB93-...}            COM-Klasse (InprocServer32, ThreadingModel, ProgID)
#   HKCU\Software\Classes\rethink.addin                   ProgID -> CLSID (PowerPoint sucht über diesen Namen)
#   HKCU\Software\Microsoft\Office\PowerPoint\Addins\rethink.addin   Anmeldung bei PowerPoint (LoadBehavior 3)
#
# Rückgängig: scripts\unregister.ps1
param([string]$Dll = (Join-Path $PSScriptRoot '..\bin\rethink-addin.dll'))
$ErrorActionPreference = 'Stop'

$Clsid = '{68EFFB93-253D-4A83-AA32-87F9C46D9023}'
$ProgId = 'rethink.addin'
$Name = 'rethink PowerPoint-Add-in'

$path = (Resolve-Path -LiteralPath $Dll).Path
if ($path.Length -ge 260) { throw "Pfad zu lang für LoadLibrary ($($path.Length) Zeichen): $path" }

# Nur 64-Bit-DLLs passen zu 64-Bit-Office (PE-Kopf: Machine 0x8664).
$bytes = [IO.File]::ReadAllBytes($path)
$pe = [BitConverter]::ToInt32($bytes, 0x3C)
if ([BitConverter]::ToUInt16($bytes, $pe + 4) -ne 0x8664) { throw "Keine x64-DLL: $path" }

$hkcu = [Microsoft.Win32.Registry]::CurrentUser
function Set-Key([string]$sub, [hashtable]$values) {
    $key = $hkcu.CreateSubKey($sub)
    try {
        foreach ($n in $values.Keys) {
            $v = $values[$n]
            if ($v -is [int]) { $key.SetValue($n, $v, [Microsoft.Win32.RegistryValueKind]::DWord) }
            else { $key.SetValue($n, [string]$v, [Microsoft.Win32.RegistryValueKind]::String) }
        }
    } finally { $key.Close() }
}

Set-Key "Software\Classes\CLSID\$Clsid" @{ '' = $Name }
Set-Key "Software\Classes\CLSID\$Clsid\InprocServer32" @{ '' = $path; 'ThreadingModel' = 'Apartment' }
Set-Key "Software\Classes\CLSID\$Clsid\ProgID" @{ '' = $ProgId }
Set-Key "Software\Classes\$ProgId" @{ '' = $Name }
Set-Key "Software\Classes\$ProgId\CLSID" @{ '' = $Clsid }
Set-Key "Software\Microsoft\Office\PowerPoint\Addins\$ProgId" @{
    'FriendlyName' = 'rethink'
    'Description'  = 'Präsentationen aus .ppttc-Daten erzeugen, Diagramme einfügen und Elemente benennen'
    'LoadBehavior' = 3
}

"Registriert: $ProgId $Clsid"
"DLL: $path"
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    'PowerPoint läuft bereits: Der Tab erscheint nach einem Neustart von PowerPoint oder über'
    'Datei > Optionen > Add-Ins > COM-Add-Ins > rethink aktivieren.'
}
