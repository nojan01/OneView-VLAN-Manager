#Requires -Version 7.5

$tool = Join-Path $PSScriptRoot 'OneView-RackmountFirmware-GUI.ps1'
$applianceFile = Join-Path $PSScriptRoot 'Synergy.txt'
if (-not (Test-Path -LiteralPath $tool -PathType Leaf)) {
    throw "Firmware-Update-Kern nicht gefunden: $tool"
}

& $tool -ServerMode Synergy -ApplianceFile $applianceFile
