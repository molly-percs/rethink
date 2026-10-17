# Entfernt alle Registrierungseinträge des rethink-Add-ins aus HKCU: die drei von register.ps1
# angelegten Schlüssel und die Werte, die PowerPoint selbst zum Add-in anlegt (Ladezeiten, Resilienz).
$ErrorActionPreference = 'Stop'

$Clsid = '{68EFFB93-253D-4A83-AA32-87F9C46D9023}'
$ProgId = 'rethink.addin'
$hkcu = [Microsoft.Win32.Registry]::CurrentUser

foreach ($sub in "Software\Classes\CLSID\$Clsid", "Software\Classes\$ProgId", "Software\Microsoft\Office\PowerPoint\Addins\$ProgId") {
    if ($hkcu.OpenSubKey($sub)) {
        $hkcu.DeleteSubKeyTree($sub)
        "Entfernt: HKCU\$sub"
    }
}

# Von PowerPoint geschriebene Werte: Name beginnt mit der ProgID, oder (DisabledItems) Binärdaten mit dem DLL-Namen.
$marker = [Text.Encoding]::Unicode.GetBytes('rethink-addin.dll')
function Test-Contains([byte[]]$data, [byte[]]$part) {
    $hex = [BitConverter]::ToString($data)
    return $hex.Contains([BitConverter]::ToString($part))
}
$office = $hkcu.OpenSubKey('Software\Microsoft\Office')
$versions = if ($office) { $office.GetSubKeyNames() | Where-Object { $_ -match '^\d+\.\d+$' } } else { @() }
foreach ($ver in $versions) {
    foreach ($sub in 'AddInLoadTimes', 'Resiliency\DoNotDisableAddinList', 'Resiliency\NotificationReminderAddinData',
                     'Resiliency\DisabledItems', 'Resiliency\CrashingAddinList') {
        $keyPath = "Software\Microsoft\Office\$ver\PowerPoint\$sub"
        $key = $hkcu.OpenSubKey($keyPath, $true)
        if (-not $key) { continue }
        try {
            foreach ($value in $key.GetValueNames()) {
                $data = $key.GetValue($value)
                $match = $value -like "$ProgId*" -or ($data -is [byte[]] -and (Test-Contains $data $marker))
                if ($match) {
                    $key.DeleteValue($value)
                    "Entfernt: HKCU\$keyPath -> $value"
                }
            }
        } finally { $key.Close() }
    }
}

# Office merkt sich das geprüfte Ribbon-XML je Add-in (Wert "rethink.addin.Microsoft.PowerPoint.Presentation").
foreach ($ver in $versions) {
    $keyPath = "Software\Microsoft\Office\$ver\Common\CustomUIValidationCache"
    $key = $hkcu.OpenSubKey($keyPath, $true)
    if (-not $key) { continue }
    try {
        foreach ($value in $key.GetValueNames() | Where-Object { $_ -like "$ProgId.*" }) {
            $key.DeleteValue($value)
            "Entfernt: HKCU\$keyPath -> $value"
        }
    } finally { $key.Close() }
}

# Kontrolle: nichts mehr vorhanden?
$left = @("Software\Classes\CLSID\$Clsid", "Software\Classes\$ProgId", "Software\Microsoft\Office\PowerPoint\Addins\$ProgId") |
    Where-Object { $hkcu.OpenSubKey($_) }
if ($left) { throw "Nicht entfernt: $($left -join ', ')" }
'rethink-Add-in ist abgemeldet.'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    'Hinweis: In einem laufenden PowerPoint bleibt das Add-in bis zum Beenden geladen.'
}
