<#
.SYNOPSIS
    Exportiert Server Profile Templates aus HPE OneView als JSON-Dateien.

.DESCRIPTION
    Dieses Skript liest alle Server Profile Templates über die HPE OneView RESTful API
    aus und speichert sie als individuelle JSON-Dateien in einem Verzeichnis.
    Zusätzlich wird eine Übersichtsdatei (_index.json) erstellt.

    Ablauf:
    1. Konfiguration aus config.json laden
    2. Authentifizierung an der OneView Appliance (REST API)
    3. Server Profile Templates abrufen (GET /rest/server-profile-templates)
    4. Jedes Template als JSON-Datei speichern
    5. Index-Datei erstellen
    6. Session abmelden

.PARAMETER ConfigPath
    Pfad zur Konfigurationsdatei (Standard: .\config.json)

.PARAMETER OutputPath
    Pfad zum Ausgabeverzeichnis (Standard: .\ServerProfileTemplates_Export_<Datum>)

.PARAMETER LogPath
    Optionaler Pfad zur Log-Datei

.EXAMPLE
    .\Export-ServerProfileTemplates.ps1
    Exportiert mit Standardkonfiguration.

.EXAMPLE
    .\Export-ServerProfileTemplates.ps1 -OutputPath "C:\Exports\SPT_Backup"
    Exportiert in ein benutzerdefiniertes Verzeichnis.

.NOTES
    Autor:   OneView VLAN Projekt
    Datum:   2026-04-01
    Benötigt: PowerShell 7.x
    API-Ref:  HPE OneView RESTful API
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),

    [Parameter()]
    [string]$OutputPath = "",

    [Parameter()]
    [string]$LogPath = ""
)

$ErrorActionPreference = "Stop"

# ============================================================================
#  Logging
# ============================================================================
$script:logFile = $null

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Message"
    Write-Output $line
    if ($script:logFile) {
        $line | Out-File -FilePath $script:logFile -Append -Encoding UTF8
    }
}

# ============================================================================
#  Konfiguration laden
# ============================================================================
if (-not (Test-Path $ConfigPath)) {
    Write-Log "Konfigurationsdatei nicht gefunden: $ConfigPath" "ERROR"
    exit 1
}

$config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json

if (-not $config.OneViewAppliances -or $config.OneViewAppliances.Count -eq 0) {
    Write-Log "Keine Appliance in der Konfiguration definiert." "ERROR"
    exit 1
}

$appliance  = $config.OneViewAppliances[0]
$apiVersion = if ($config.ApiVersion) { [int]$config.ApiVersion } else { 8000 }

# ============================================================================
#  Ausgabeverzeichnis
# ============================================================================
if (-not $OutputPath) {
    $OutputPath = Join-Path $PSScriptRoot ("ServerProfileTemplates_Export_{0}" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
}
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# ============================================================================
#  Log-Datei
# ============================================================================
if ($LogPath) {
    $script:logFile = $LogPath
} else {
    $logsDir = Join-Path $PSScriptRoot "Logs"
    if (-not (Test-Path $logsDir)) { New-Item -Path $logsDir -ItemType Directory -Force | Out-Null }
    $script:logFile = Join-Path $logsDir ("SPT_Export_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
}

Write-Log "Server Profile Template Export gestartet"
Write-Log "Appliance: $($appliance.Hostname)"
Write-Log "Ausgabe:   $OutputPath"

# ============================================================================
#  API-Version ermitteln
# ============================================================================
$hostname = $appliance.Hostname
$baseUri  = "https://$hostname"

try {
    $verResponse = Invoke-RestMethod -Uri "$baseUri/rest/version" -Method Get -SkipCertificateCheck
    $apiVersion = [int]$verResponse.currentVersion
    Write-Log "API-Version: $apiVersion"
} catch {
    Write-Log "API-Version konnte nicht ermittelt werden, verwende $apiVersion" "WARN"
}

# ============================================================================
#  Authentifizierung
# ============================================================================
if ($env:OV_USERNAME -and $env:OV_PASSWORD) {
    $username = $env:OV_USERNAME
    $password = $env:OV_PASSWORD
} else {
    $cred = Get-Credential -Message "OneView Anmeldung für $hostname"
    $username = $cred.UserName
    $password = $cred.GetNetworkCredential().Password
}

$loginBody = @{
    userName        = $username
    password        = $password
    authLoginDomain = "Local"
} | ConvertTo-Json

$headers = @{
    "Content-Type"  = "application/json"
    "X-API-Version" = $apiVersion
}

try {
    $loginResponse = Invoke-RestMethod -Uri "$baseUri/rest/login-sessions" `
        -Method Post -Headers $headers -Body $loginBody -SkipCertificateCheck
    $sessionId = $loginResponse.sessionID
    if ([string]::IsNullOrEmpty($sessionId)) { throw "Keine sessionID erhalten" }
    Write-Log "Authentifizierung erfolgreich" "SUCCESS"
} catch {
    Write-Log "Authentifizierung fehlgeschlagen: $_" "ERROR"
    exit 1
}

$authHeaders = @{
    "Content-Type"  = "application/json"
    "X-API-Version" = $apiVersion
    "Auth"          = $sessionId
}

# ============================================================================
#  Server Profile Templates abrufen (paginiert)
# ============================================================================
try {
    Write-Log "Lade Server Profile Templates..."
    $allTemplates = [System.Collections.Generic.List[object]]::new()
    $uri = "$baseUri/rest/server-profile-templates?start=0&count=100"

    do {
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $authHeaders -SkipCertificateCheck
        if ($response.members) {
            $allTemplates.AddRange([object[]]$response.members)
        }
        $uri = if ($response.nextPageUri) { "$baseUri$($response.nextPageUri)" } else { $null }
    } while ($uri)

    Write-Log "$($allTemplates.Count) Server Profile Template(s) gefunden" "SUCCESS"
} catch {
    Write-Log "Fehler beim Laden der Server Profile Templates: $_" "ERROR"
    Invoke-RestMethod -Uri "$baseUri/rest/login-sessions" -Method Delete -Headers $authHeaders -SkipCertificateCheck -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# ============================================================================
#  Templates als JSON speichern
# ============================================================================
$index = @()

foreach ($template in $allTemplates) {
    $safeName = ($template.name -replace '[\\/:*?\"<>|\s]', '_')
    $filePath = Join-Path $OutputPath "${safeName}.json"

    $template | ConvertTo-Json -Depth 20 | Set-Content -Path $filePath -Encoding UTF8
    Write-Log "Exportiert: $($template.name) → $(Split-Path $filePath -Leaf)"

    $index += [PSCustomObject]@{
        Name                     = $template.name
        Uri                      = $template.uri
        Status                   = $template.status
        State                    = $template.state
        ServerHardwareTypeUri    = $template.serverHardwareTypeUri
        EnclosureGroupUri        = $template.enclosureGroupUri
        FileName                 = "${safeName}.json"
    }
}

# Index-Datei speichern
$indexPath = Join-Path $OutputPath "_index.json"
$index | ConvertTo-Json -Depth 5 | Set-Content -Path $indexPath -Encoding UTF8
Write-Log "Index-Datei: $(Split-Path $indexPath -Leaf)"

# ============================================================================
#  Session beenden
# ============================================================================
try {
    Invoke-RestMethod -Uri "$baseUri/rest/login-sessions" -Method Delete `
        -Headers $authHeaders -SkipCertificateCheck | Out-Null
} catch { }

Write-Log "Export abgeschlossen – $($allTemplates.Count) Templates exportiert" "SUCCESS"
