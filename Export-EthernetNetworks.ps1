<#
.SYNOPSIS
    Exportiert alle Ethernet Networks aus HPE OneView in eine Excel-Datei.

.DESCRIPTION
    Dieses Skript liest alle Ethernet Networks, Network Sets, Scopes und
    Connection Templates über die HPE OneView RESTful API aus und exportiert
    die Daten in eine Excel-Datei mit identischem Aufbau wie die Import-Vorlage.

    Die exportierte Datei kann direkt als Eingabe für Create-EthernetNetworks.ps1
    verwendet werden (z.B. um Netzwerke auf eine andere Appliance zu übertragen).

    Ablauf:
    1. Konfiguration aus config.json laden
    2. Authentifizierung an der OneView Appliance (REST API)
    3. Ethernet Networks abrufen     (GET /rest/ethernet-networks)
    4. Network Sets abrufen          (GET /rest/network-sets)
    5. Scopes abrufen                (GET /rest/scopes)
    6. Connection Templates abrufen  (GET /rest/connection-templates)
    7. Daten zusammenführen und in Excel exportieren

.PARAMETER ConfigPath
    Pfad zur Konfigurationsdatei (Standard: .\config.json)

.PARAMETER OutputPath
    Pfad für die Ausgabe-Excel-Datei (Standard: .\OneView_Networks_Export_<Datum>.xlsx)

.EXAMPLE
    .\Export-EthernetNetworks.ps1
    Exportiert mit Standardkonfiguration.

.EXAMPLE
    .\Export-EthernetNetworks.ps1 -OutputPath "C:\Exports\Networks.xlsx"
    Exportiert in eine benutzerdefinierte Datei.

.NOTES
    Autor:   OneView VLAN Projekt
    Datum:   2026-02-06
    Benötigt: PowerShell 7.x, Modul "ImportExcel"
    API-Ref:  https://support.hpe.com/docs/display/public/dp00006616en_us/index.html
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

# ============================================================================
#  Initialisierung
# ============================================================================
$ErrorActionPreference = "Stop"
$script:LogEntries = [System.Collections.Generic.List[string]]::new()
$script:LogPath = $LogPath

# ============================================================================
#  Hilfsfunktionen
# ============================================================================

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"
    $script:LogEntries.Add($entry)

    switch ($Level) {
        "INFO"    { Write-Host $entry -ForegroundColor Gray }
        "WARN"    { Write-Host $entry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $entry -ForegroundColor Red }
        "SUCCESS" { Write-Host $entry -ForegroundColor Green }
    }
}

function Save-Log {
    param([string]$LogDir = $PSScriptRoot)
    if ($script:LogPath) {
        $logFile = $script:LogPath
        $parent = Split-Path $logFile -Parent
        if (-not (Test-Path $parent)) { New-Item -Path $parent -ItemType Directory -Force | Out-Null }
        $script:LogEntries | Add-Content -Path $logFile -Encoding UTF8
    } else {
        $logsDir = Join-Path $LogDir "Logs"
        if (-not (Test-Path $logsDir)) { New-Item -Path $logsDir -ItemType Directory -Force | Out-Null }
        $logFile = Join-Path $logsDir ("VLAN_Export_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
        $script:LogEntries | Set-Content -Path $logFile -Encoding UTF8
    }
    Write-Host "`nProtokoll gespeichert: $logFile" -ForegroundColor Cyan
}

# ============================================================================
#  OneView REST API – Session Management
# ============================================================================

function Connect-OneViewAPI {
    param(
        [Parameter(Mandatory)][string]$Hostname,
        [Parameter(Mandatory)][PSCredential]$Credential,
        [Parameter(Mandatory)][int]$ApiVersion
    )

    $baseUri  = "https://$Hostname"
    $loginUri = "$baseUri/rest/login-sessions"

    $body = @{
        userName        = $Credential.UserName
        password        = $Credential.GetNetworkCredential().Password
        authLoginDomain = "Local"
    } | ConvertTo-Json

    $headers = @{
        "Content-Type"  = "application/json"
        "X-API-Version" = $ApiVersion
    }

    Write-Log "Verbinde zu OneView Appliance: $Hostname (API-Version: $ApiVersion)"

    try {
        $response = Invoke-RestMethod -Uri $loginUri `
            -Method Post `
            -Headers $headers `
            -Body $body `
            -SkipCertificateCheck

        $sessionId = $response.sessionID
        if ([string]::IsNullOrEmpty($sessionId)) {
            throw "Keine sessionID in der Antwort erhalten."
        }

        Write-Log "Erfolgreich authentifiziert an $Hostname" -Level SUCCESS

        return @{
            BaseUri   = $baseUri
            SessionId = $sessionId
            Headers   = @{
                "Content-Type"  = "application/json"
                "X-API-Version" = $ApiVersion
                "Auth"          = $sessionId
            }
        }
    }
    catch {
        Write-Log "Authentifizierung fehlgeschlagen für ${Hostname}: $_" -Level ERROR
        throw
    }
}

function Disconnect-OneViewAPI {
    param([Parameter(Mandatory)][hashtable]$Session)

    try {
        Invoke-RestMethod -Uri "$($Session.BaseUri)/rest/login-sessions" `
            -Method Delete `
            -Headers $Session.Headers `
            -SkipCertificateCheck | Out-Null
        Write-Log "Session abgemeldet von $($Session.BaseUri)" -Level INFO
    }
    catch {
        Write-Log "Fehler beim Abmelden: $_" -Level WARN
    }
}

function Get-OneViewApiVersion {
    <#
    .SYNOPSIS  Ermittelt die aktuelle API-Version einer OneView Appliance via GET /rest/version
    .DESCRIPTION
        Der Endpoint /rest/version erfordert KEINE Authentifizierung und liefert
        die aktuelle (currentVersion) und minimale (minimumVersion) API-Version.
        Gibt bei Fehler den übergebenen Fallback-Wert zurück.
    #>
    param(
        [Parameter(Mandatory)][string]$Hostname,
        [int]$FallbackVersion = 8000
    )

    $uri = "https://$Hostname/rest/version"

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Get -SkipCertificateCheck
        $detected = [int]$response.currentVersion
        Write-Log "API-Version von $Hostname automatisch erkannt: $detected" -Level SUCCESS
        return $detected
    }
    catch {
        Write-Log "API-Version von $Hostname konnte nicht ermittelt werden – verwende Fallback: $FallbackVersion" -Level WARN
        return $FallbackVersion
    }
}

# ============================================================================
#  OneView REST API – GET Funktionen (mit Paginierung)
# ============================================================================

function Get-AllPaginated {
    <#
    .SYNOPSIS
        Ruft alle Ergebnisse einer paginierten OneView API-Ressource ab.
    .DESCRIPTION
        Die OneView API liefert standardmäßig max. 100 Einträge pro Seite.
        Diese Funktion iteriert über alle Seiten und gibt sämtliche Members zurück.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][string]$ResourcePath,
        [int]$PageSize = 200
    )

    $allMembers = [System.Collections.Generic.List[object]]::new()
    $uri = "$($Session.BaseUri)$ResourcePath" + "?start=0&count=$PageSize"

    do {
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck

        if ($response.members) {
            $allMembers.AddRange([object[]]$response.members)
        }

        Write-Log "  Seite abgerufen: $($allMembers.Count) / $($response.total) Einträge" -Level INFO

        $uri = if ($response.nextPageUri) { "$($Session.BaseUri)$($response.nextPageUri)" } else { $null }
    } while ($uri)

    return $allMembers
}

function Get-AllEthernetNetworks {
    param([Parameter(Mandatory)][hashtable]$Session)

    Write-Log "Rufe Ethernet Networks ab (paginiert)..."
    $members = Get-AllPaginated -Session $Session -ResourcePath "/rest/ethernet-networks"
    Write-Log "  $($members.Count) Ethernet Networks gefunden." -Level INFO
    return $members
}

function Get-AllNetworkSets {
    param([Parameter(Mandatory)][hashtable]$Session)

    Write-Log "Rufe Network Sets ab (paginiert)..."
    try {
        $members = Get-AllPaginated -Session $Session -ResourcePath "/rest/network-sets"
        Write-Log "  $($members.Count) Network Sets gefunden." -Level INFO
        return $members
    }
    catch {
        Write-Log "  Network Sets konnten nicht abgerufen werden: $_" -Level WARN
        return @()
    }
}

function Get-AllScopes {
    param([Parameter(Mandatory)][hashtable]$Session)

    Write-Log "Rufe Scopes ab (paginiert)..."
    try {
        $members = Get-AllPaginated -Session $Session -ResourcePath "/rest/scopes"
        Write-Log "  $($members.Count) Scopes gefunden." -Level INFO
        return $members
    }
    catch {
        Write-Log "  Scopes konnten nicht abgerufen werden: $_" -Level WARN
        return @()
    }
}

function Get-ConnectionTemplate {
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][string]$ConnectionTemplateUri
    )

    $uri = "$($Session.BaseUri)$ConnectionTemplateUri"
    try {
        $template = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck
        return $template
    }
    catch {
        Write-Log "  Connection Template konnte nicht abgerufen werden: $ConnectionTemplateUri" -Level WARN
        return $null
    }
}

# ============================================================================
#  Hauptprogramm
# ============================================================================

function Main {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   HPE OneView – Ethernet Networks nach Excel exportieren    ║" -ForegroundColor Cyan
    Write-Host "║   Über die RESTful API (ohne HPE OneView PowerShell Modul)  ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    # -------------------------------------------
    # 1. Konfiguration laden
    # -------------------------------------------
    Write-Log "Lade Konfiguration aus: $ConfigPath"
    if (-not (Test-Path $ConfigPath)) {
        Write-Log "Konfigurationsdatei nicht gefunden: $ConfigPath" -Level ERROR
        return
    }
    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json

    # ImportExcel-Modul prüfen
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log "Installiere Modul 'ImportExcel'..." -Level WARN
        Install-Module -Name ImportExcel -Scope CurrentUser -Force
    }
    Import-Module ImportExcel

    # -------------------------------------------
    # 2. Anmeldedaten abfragen
    # -------------------------------------------
    $credential = Get-Credential -Message "OneView Anmeldedaten eingeben (Benutzername & Kennwort)"
    if (-not $credential) {
        Write-Log "Keine Anmeldedaten eingegeben – Abbruch." -Level ERROR
        return
    }

    # -------------------------------------------
    # 3. Für jede Appliance: Netzwerke exportieren
    # -------------------------------------------
    foreach ($appliance in $config.OneViewAppliances) {
        Write-Host ""
        Write-Log "============================================================"
        Write-Log "Appliance: $($appliance.Name) ($($appliance.Hostname))"
        Write-Log "============================================================"

        $session = $null

        try {
            # 3a. API-Version automatisch erkennen & Anmelden
            $detectedApiVersion = Get-OneViewApiVersion -Hostname $appliance.Hostname -FallbackVersion $config.ApiVersion
            $session = Connect-OneViewAPI -Hostname $appliance.Hostname `
                -Credential $credential `
                -ApiVersion $detectedApiVersion

            # 3b. Alle Daten abrufen
            $networks    = Get-AllEthernetNetworks -Session $session
            $networkSets = Get-AllNetworkSets -Session $session
            $scopes      = Get-AllScopes -Session $session

            # -------------------------------------------
            # Lookup-Tabellen erstellen
            # -------------------------------------------

            # Network-URI → Liste der Network Set Namen
            $networkToSetsMap = @{}
            foreach ($ns in $networkSets) {
                if ($ns.networkUris) {
                    foreach ($netUri in $ns.networkUris) {
                        if (-not $networkToSetsMap.ContainsKey($netUri)) {
                            $networkToSetsMap[$netUri] = [System.Collections.Generic.List[string]]::new()
                        }
                        $networkToSetsMap[$netUri].Add($ns.name)
                    }
                }
            }

            # Network-URI → Liste der Scope Namen
            $networkToScopesMap = @{}
            foreach ($sc in $scopes) {
                if ($sc.resourceUris) {
                    foreach ($resUri in $sc.resourceUris) {
                        if (-not $networkToScopesMap.ContainsKey($resUri)) {
                            $networkToScopesMap[$resUri] = [System.Collections.Generic.List[string]]::new()
                        }
                        $networkToScopesMap[$resUri].Add($sc.name)
                    }
                }
            }

            # -------------------------------------------
            # 3c. Daten zusammenführen
            # -------------------------------------------
            Write-Log "Verarbeite $($networks.Count) Ethernet Networks..."

            $exportData = [System.Collections.Generic.List[PSCustomObject]]::new()

            foreach ($net in $networks) {
                # Bandwidth aus Connection Template lesen
                $bwPreferred = 2.5   # Default
                $bwMaximum   = 50    # Default

                if ($net.connectionTemplateUri) {
                    $ct = Get-ConnectionTemplate -Session $session -ConnectionTemplateUri $net.connectionTemplateUri
                    if ($ct -and $ct.bandwidth) {
                        $bwPreferred = [math]::Round($ct.bandwidth.typicalBandwidth / 1000, 2)
                        $bwMaximum   = [math]::Round($ct.bandwidth.maximumBandwidth / 1000, 2)
                    }
                }

                # Scope(s) ermitteln
                $scopeNames = ""
                if ($networkToScopesMap.ContainsKey($net.uri)) {
                    $scopeNames = ($networkToScopesMap[$net.uri] | Sort-Object) -join "; "
                }

                # Network Set(s) ermitteln
                $setNames = ""
                if ($networkToSetsMap.ContainsKey($net.uri)) {
                    $setNames = ($networkToSetsMap[$net.uri] | Sort-Object) -join "; "
                }

                # IPv4 / IPv6 Subnet
                $ipv4Subnet = if ($net.subnetUri) { $net.subnetUri } else { "" }
                $ipv6Subnet = if ($net.ipv6SubnetUri) { $net.ipv6SubnetUri } else { "" }

                # Description
                $description = if ($net.description) { $net.description } else { "" }

                $exportData.Add([PSCustomObject]@{
                    NetworkName          = $net.name
                    VlanId               = $net.vlanId
                    EthernetNetworkType  = $net.ethernetNetworkType
                    Purpose              = $net.purpose
                    SmartLink            = $net.smartLink
                    PrivateNetwork       = $net.privateNetwork
                    PreferredBandwidthGb = $bwPreferred
                    MaximumBandwidthGb   = $bwMaximum
                    Scope                = $scopeNames
                    NetworkSet           = $setNames
                    IPv4SubnetId         = $ipv4Subnet
                    IPv6SubnetId         = $ipv6Subnet
                    Description          = $description
                })
            }

            # -------------------------------------------
            # 3d. Excel exportieren
            # -------------------------------------------
            $outFile = $OutputPath
            if ([string]::IsNullOrWhiteSpace($outFile)) {
                $safeName = $appliance.Name -replace '[\\/:*?"<>|]', '_'
                $outFile = Join-Path $PSScriptRoot ("VLAN_{0}_{1}.xlsx" -f $safeName, (Get-Date -Format "yyyy-MM-dd"))
            }

            $exportData | Export-Excel -Path $outFile `
                -WorksheetName "VLANs" `
                -AutoSize `
                -FreezeTopRow `
                -BoldTopRow

            Write-Log "Export abgeschlossen: $($exportData.Count) Netzwerke" -Level SUCCESS
            Write-Log "Datei: $outFile" -Level SUCCESS

            Write-Host ""
            Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
            Write-Host ("║  Exportiert: {0,-47}║" -f "$($exportData.Count) Ethernet Networks") -ForegroundColor Green
            Write-Host ("║  Appliance:  {0,-47}║" -f $appliance.Name) -ForegroundColor Gray
            Write-Host ("║  Datei:      {0,-47}║" -f (Split-Path $outFile -Leaf)) -ForegroundColor Gray
            Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        }
        catch {
            Write-Log "Kritischer Fehler bei Appliance $($appliance.Hostname): $_" -Level ERROR
        }
        finally {
            if ($session) {
                Disconnect-OneViewAPI -Session $session
            }
        }
    }

    Save-Log
}

# Skript starten
Main
