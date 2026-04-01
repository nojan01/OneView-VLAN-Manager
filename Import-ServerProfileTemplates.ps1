<#
.SYNOPSIS
    Importiert Server Profile Templates in HPE OneView aus JSON-Dateien.

.DESCRIPTION
    Dieses Skript liest Server Profile Template-Definitionen aus JSON-Dateien und
    erstellt oder aktualisiert die entsprechenden Templates über die
    HPE OneView RESTful API.

    Modi:
    - Auto:   Prüft ob ein Template mit gleichem Namen existiert.
              Existiert es → Update (PUT), sonst → Create (POST).
    - Create: Erstellt neue Templates (Fehler wenn Name existiert).
    - Update: Aktualisiert bestehende Templates (Fehler wenn Name nicht existiert).

    Bei der Erstellung werden schreibgeschützte Felder (uri, eTag, created,
    modified, etc.) automatisch entfernt.

    Ablauf:
    1. Konfiguration aus config.json laden
    2. JSON-Dateien einlesen
    3. Authentifizierung an der OneView Appliance (REST API)
    4. Existierende Templates abrufen (für Auto/Update-Modus)
    5. Templates erstellen oder aktualisieren
    6. Session abmelden

.PARAMETER ConfigPath
    Pfad zur Konfigurationsdatei (Standard: .\config.json)

.PARAMETER InputPath
    Pfad zu einer JSON-Datei oder einem Verzeichnis mit JSON-Dateien.

.PARAMETER Mode
    Importmodus: Auto, Create oder Update (Standard: Auto)

.PARAMETER LogPath
    Optionaler Pfad zur Log-Datei

.EXAMPLE
    .\Import-ServerProfileTemplates.ps1 -InputPath ".\SPT_Backup\MyTemplate.json"
    Importiert ein einzelnes Template (Auto-Modus).

.EXAMPLE
    .\Import-ServerProfileTemplates.ps1 -InputPath ".\SPT_Backup" -Mode Create
    Erstellt alle Templates aus dem Verzeichnis.

.NOTES
    Autor:   OneView VLAN Projekt
    Datum:   2026-04-01
    Benötigt: PowerShell 7.x
    API-Ref:  HPE OneView RESTful API
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),

    [Parameter(Mandatory)]
    [string]$InputPath,

    [Parameter()]
    [ValidateSet("Create", "Update", "Auto")]
    [string]$Mode = "Auto",

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
#  Log-Datei
# ============================================================================
if ($LogPath) {
    $script:logFile = $LogPath
} else {
    $logsDir = Join-Path $PSScriptRoot "Logs"
    if (-not (Test-Path $logsDir)) { New-Item -Path $logsDir -ItemType Directory -Force | Out-Null }
    $script:logFile = Join-Path $logsDir ("SPT_Import_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
}

# ============================================================================
#  JSON-Dateien einlesen
# ============================================================================
if (-not (Test-Path $InputPath)) {
    Write-Log "Eingabepfad nicht gefunden: $InputPath" "ERROR"
    exit 1
}

$jsonFiles = @()
if ((Get-Item $InputPath).PSIsContainer) {
    $jsonFiles = Get-ChildItem -Path $InputPath -Filter "*.json" |
        Where-Object { $_.Name -ne "_index.json" } |
        Select-Object -ExpandProperty FullName
} else {
    $jsonFiles = @($InputPath)
}

if ($jsonFiles.Count -eq 0) {
    Write-Log "Keine JSON-Dateien gefunden in: $InputPath" "ERROR"
    exit 1
}

Write-Log "Server Profile Template Import gestartet (Modus: $Mode)"
Write-Log "Appliance: $($appliance.Hostname)"
Write-Log "$($jsonFiles.Count) JSON-Datei(en) gefunden"

# ============================================================================
#  Templates aus JSON laden
# ============================================================================
$templatesToImport = @()
foreach ($file in $jsonFiles) {
    try {
        $templateData = Get-Content -Path $file -Raw | ConvertFrom-Json
        $templatesToImport += @{
            FileName = (Split-Path $file -Leaf)
            Data     = $templateData
        }
        Write-Log "Geladen: $(Split-Path $file -Leaf) → $($templateData.name)"
    } catch {
        Write-Log "Fehler beim Lesen von $(Split-Path $file -Leaf): $_" "ERROR"
    }
}

if ($templatesToImport.Count -eq 0) {
    Write-Log "Keine gültigen Templates geladen." "ERROR"
    exit 1
}

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
#  Existierende Templates abrufen (für Auto/Update)
# ============================================================================
$existingTemplates = @{}

if ($Mode -ne "Create") {
    try {
        Write-Log "Lade existierende Server Profile Templates..."
        $start    = 0
        $pageSize = 100
        do {
            $uri = "$baseUri/rest/server-profile-templates?start=$start&count=$pageSize"
            $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $authHeaders -SkipCertificateCheck
            foreach ($member in $response.members) {
                $existingTemplates[$member.name] = $member
            }
            $total = $response.total
            $start += $response.members.Count
        } while ($existingTemplates.Count -lt $total)
        Write-Log "$($existingTemplates.Count) existierende Templates gefunden"
    } catch {
        Write-Log "Fehler beim Laden existierender Templates: $_" "ERROR"
    }
}

# ============================================================================
#  Schreibgeschützte Felder für Neuerstellung entfernen
# ============================================================================
function Remove-ReadOnlyFields {
    param([object]$TemplateObject)

    $json = $TemplateObject | ConvertTo-Json -Depth 20
    $ht = $json | ConvertFrom-Json -AsHashtable

    $readOnlyFields = @(
        "uri", "eTag", "created", "modified", "uuid",
        "taskUri", "stateReason", "refreshState",
        "inProgress", "scopesUri"
    )

    foreach ($field in $readOnlyFields) {
        $ht.Remove($field)
    }

    return $ht
}

# ============================================================================
#  Templates importieren
# ============================================================================
$successCount = 0
$errorCount   = 0

foreach ($item in $templatesToImport) {
    $templateData = $item.Data
    $templateName = $templateData.name
    $fileName     = $item.FileName

    Write-Log "Verarbeite: $templateName ($fileName)"

    try {
        $existing = $existingTemplates[$templateName]

        if ($Mode -eq "Create" -and $existing) {
            Write-Log "Template '$templateName' existiert bereits (Modus: Create) – übersprungen" "WARN"
            $errorCount++
            continue
        }

        if ($Mode -eq "Update" -and -not $existing) {
            Write-Log "Template '$templateName' nicht gefunden (Modus: Update) – übersprungen" "WARN"
            $errorCount++
            continue
        }

        if ($existing) {
            # ── Update (PUT) ──
            Write-Log "  Aktualisiere existierendes Template: $templateName"

            $templateData.eTag = $existing.eTag
            $templateData.uri  = $existing.uri

            $body = $templateData | ConvertTo-Json -Depth 20
            $updateUri = "$baseUri$($existing.uri)"

            $response = Invoke-RestMethod -Uri $updateUri -Method Put `
                -Headers $authHeaders -Body $body -SkipCertificateCheck

            Write-Log "  Template aktualisiert: $templateName" "SUCCESS"
            $successCount++
        } else {
            # ── Create (POST) ──
            Write-Log "  Erstelle neues Template: $templateName"

            $cleanTemplate = Remove-ReadOnlyFields -TemplateObject $templateData
            $body = $cleanTemplate | ConvertTo-Json -Depth 20

            $response = Invoke-RestMethod -Uri "$baseUri/rest/server-profile-templates" `
                -Method Post -Headers $authHeaders -Body $body -SkipCertificateCheck

            Write-Log "  Template erstellt: $templateName" "SUCCESS"
            $successCount++
        }
    } catch {
        Write-Log "  Fehler bei $templateName : $_" "ERROR"
        $errorCount++
    }
}

# ============================================================================
#  Session beenden
# ============================================================================
try {
    Invoke-RestMethod -Uri "$baseUri/rest/login-sessions" -Method Delete `
        -Headers $authHeaders -SkipCertificateCheck | Out-Null
} catch { }

Write-Log ""
Write-Log "Import abgeschlossen: $successCount erfolgreich, $errorCount fehlgeschlagen" "SUCCESS"
