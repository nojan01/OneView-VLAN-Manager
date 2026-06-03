# ============================================================================
#  OneView_VersionMap.ps1
#  Zentrale Versions-/Modul-Tabelle fuer alle Backup- und Update-Skripte.
#  Wird per Dot-Sourcing eingebunden:  . "$PSScriptRoot\OneView_VersionMap.ps1"
#
#  Pflege bei neuer OneView-Minor-/Major-Version:
#    - Bevorzugt einen exakten Eintrag mit MajorMinor='<x>.<y>' anlegen
#      (z.B. '11.30'). Jede Minor-Version kann ihr eigenes PowerShell-Modul
#      UND ihre eigene X-API-Version mitbringen.
#    - Falls fuer eine ganze Major-Reihe nur ein Modul existiert, kann
#      zusaetzlich ein Wildcard-Eintrag mit MajorMinor='<x>.*' als Fallback
#      dienen (greift, wenn keine exakte Minor-Zeile passt).
#    - Reihenfolge in der Tabelle ist egal: exakte Treffer haben immer
#      Vorrang vor Wildcard-Treffern.
# ============================================================================

# -----------------------------------------------------------------------------
# Mapping-Tabelle
# -----------------------------------------------------------------------------
# Felder pro Eintrag:
#   MajorMinor : '<Major>.<Minor>' (exakt) ODER '<Major>.*' (Wildcard fuer Major-Fallback)
#   Module     : Name des HPEOneView-PowerShell-Moduls (PSGallery)
#   ApiVersion : Wert fuer den HTTP-Header 'X-API-Version' (int)
#   Notes      : freier Kommentar
#
# Hinweis zu den X-API-Versionen / Modulnamen:
#   Die nachstehenden Werte entsprechen der zum Erstellungszeitpunkt
#   verfuegbaren HPE-Dokumentation. Bitte bei einem OneView-Upgrade
#   durch einen Blick in die HPE Release Notes / PSGallery verifizieren
#   und ggf. korrigieren - die Tabelle ist genau dafuer da.
# -----------------------------------------------------------------------------
$script:OvVersionMap = @(
    # OneView 6.x (Generation vor v11)
    [PSCustomObject]@{ MajorMinor = '6.60';  Module = 'HPEOneView.660';  ApiVersion = 4600; Notes = 'OV 6.60 LTS' }

    # OneView 11.x  - jede Minor hat eigenes Modul + eigene X-API-Version
    [PSCustomObject]@{ MajorMinor = '11.10'; Module = 'HPEOneView.1000'; ApiVersion = 5200; Notes = 'OV 11.10' }
    [PSCustomObject]@{ MajorMinor = '11.20'; Module = 'HPEOneView.1100'; ApiVersion = 5400; Notes = 'OV 11.20' }
    [PSCustomObject]@{ MajorMinor = '11.30'; Module = 'HPEOneView.1200'; ApiVersion = 5600; Notes = 'OV 11.30 (Werte beim Release pruefen)' }

    # Wildcard-Fallback: unbekannte 11.x-Minor -> letzte bekannte 11er-Zeile
    [PSCustomObject]@{ MajorMinor = '11.*';  Module = 'HPEOneView.1200'; ApiVersion = 5600; Notes = 'Fallback fuer unbekannte 11.x-Minor' }

    # Platzhalter fuer kuenftige Major-Versionen:
    # [PSCustomObject]@{ MajorMinor = '12.0';  Module = 'HPEOneView.1300'; ApiVersion = 6000; Notes = 'OV 12.0' }
    # [PSCustomObject]@{ MajorMinor = '12.*';  Module = 'HPEOneView.1300'; ApiVersion = 6000; Notes = 'Fallback 12.x' }
)

# Fallback-Hinweis fuer die Anzeige, falls keine Version erkannt werden konnte.
$script:OvVersionFallback = '?'

function Resolve-OvModule {
    <#
        .SYNOPSIS
        Liefert Modul- und API-Info fuer eine erkannte OneView-Software-Version.
        Sucht zuerst nach exakter Major.Minor-Zuordnung (z.B. '11.20'),
        anschliessend nach Wildcard-Eintrag '<Major>.*'.

        .PARAMETER Version
        OneView Software-Version-String, z.B. "11.20" oder "11.20.00-1234567".

        .OUTPUTS
        PSCustomObject mit MajorMinor/Module/ApiVersion/Notes - oder $null,
        wenn weder ein exakter noch ein Wildcard-Eintrag passt.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$Version)

    if ($Version -notmatch '^\s*(\d+)\.(\d+)') { return $null }
    $major    = [int]$Matches[1]
    $minor    = [int]$Matches[2]
    $exactKey = "$major.$minor"
    $wildKey  = "$major.*"

    # 1) exakter Minor-Treffer hat Vorrang
    foreach ($entry in $script:OvVersionMap) {
        if ($entry.MajorMinor -eq $exactKey) { return $entry }
    }
    # 2) Wildcard fuer dieselbe Major-Version
    foreach ($entry in $script:OvVersionMap) {
        if ($entry.MajorMinor -eq $wildKey) { return $entry }
    }
    return $null
}

function Get-OvVersionInfo {
    <#
        .SYNOPSIS
        Fragt /rest/version einer Appliance ab und liefert
        einen Datensatz mit Software-Version + zugehoerigem Modul / X-API.
        .OUTPUTS
        PSCustomObject  Appliance, SoftwareVersion, MajorMinor, Module, ApiVersion, Source, Error
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Appliance,
        [int]$TimeoutSec = 10
    )

    $result = [PSCustomObject]@{
        Appliance       = $Appliance
        SoftwareVersion = $null
        MajorMinor      = $null
        Module          = $null
        ApiVersion      = $null
        Source          = $null
        Error           = $null
    }

    # SSL-Bypass nur fuer den aktuellen Aufruf
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch {}

    $uri = "https://$Appliance/rest/version"
    try {
        $resp = if ($PSVersionTable.PSVersion.Major -ge 7) {
            Invoke-RestMethod -Uri $uri -Method Get -TimeoutSec $TimeoutSec -SkipCertificateCheck -ErrorAction Stop
        } else {
            Invoke-RestMethod -Uri $uri -Method Get -TimeoutSec $TimeoutSec -ErrorAction Stop
        }

        # OneView liefert je nach Version unterschiedliche Felder.
        # Bevorzugt softwareVersion (z.B. "11.20.00-1234567"),
        # ansonsten currentVersion (API-Version) als Heuristik.
        $sw = $null
        foreach ($prop in @('softwareVersion','SoftwareVersion','applianceVersion')) {
            if ($resp.PSObject.Properties.Name -contains $prop -and $resp.$prop) {
                $sw = [string]$resp.$prop
                break
            }
        }

        if (-not $sw) {
            # Fallback ueber currentVersion (X-API-Version): suche besten
            # exakten Eintrag (kein Wildcard) mit ApiVersion <= currentVersion.
            if ($resp.PSObject.Properties.Name -contains 'currentVersion' -and $resp.currentVersion) {
                $cv = [int]$resp.currentVersion
                $best = $null
                foreach ($entry in $script:OvVersionMap) {
                    if ($entry.MajorMinor -like '*.\*') { continue } # Wildcards ueberspringen
                    if ($entry.ApiVersion -le $cv) {
                        if ($null -eq $best -or $entry.ApiVersion -gt $best.ApiVersion) { $best = $entry }
                    }
                }
                if ($best) {
                    $sw = $best.MajorMinor
                    $result.Source = "currentVersion=$cv -> ~$sw (heuristisch)"
                }
            }
        } else {
            $result.Source = '/rest/version softwareVersion'
        }

        if (-not $sw) { throw "Antwort enthaelt keine erkennbare Software-Version." }

        $result.SoftwareVersion = $sw
        if ($sw -match '^\s*(\d+)\.(\d+)') { $result.MajorMinor = "$($Matches[1]).$($Matches[2])" }

        $entry = Resolve-OvModule -Version $sw
        if ($entry) {
            $result.Module     = $entry.Module
            $result.ApiVersion = $entry.ApiVersion
        } else {
            $result.Error = "Keine Modul-Zuordnung in OvVersionMap fuer Version '$sw'."
        }
    }
    catch {
        $result.Error = "Versionsabfrage fehlgeschlagen: $($_.Exception.Message)"
    }

    return $result
}
