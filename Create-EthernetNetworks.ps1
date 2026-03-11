<#
.SYNOPSIS
    Erstellt Ethernet Networks in HPE OneView basierend auf einer Excel-Datei.

.DESCRIPTION
    Dieses Skript liest VLAN-Definitionen aus einer Excel-Datei und erstellt
    die entsprechenden Ethernet Networks über die HPE OneView RESTful API.

    Unterstützte Felder (gemäss OneView "Create Network" Dialog):
    - Name, VLAN ID, EthernetNetworkType (Tagged/Untagged/Tunnel)
    - Purpose (General, Management, VMMigration, FaultTolerance, ISCSI)
    - SmartLink, PrivateNetwork
    - Preferred / Maximum Bandwidth (in Gb/s, wie in der GUI)
    - Scope (Zuweisung zu vorhandenem Scope)
    - Network Set (Zuweisung zu vorhandenem Network Set)
    - IPv4 / IPv6 Subnet ID (Assoziation mit vorhandenem Subnet)

    Ablauf:
    1. Konfiguration aus config.json laden
    2. VLAN-Daten aus Excel importieren & validieren
    3. Authentifizierung an der OneView Appliance (REST API)
    4. Existierende Netzwerke, Network Sets und Scopes abrufen
    5. Ethernet Networks erstellen (POST /rest/ethernet-networks)
    6. Bandwidth setzen (PUT /rest/connection-templates/{id})
    7. Network Set aktualisieren (PUT /rest/network-sets/{id})
    8. Scope zuweisen (PATCH /rest/scopes/{id})
    9. Session abmelden & Protokoll speichern

.PARAMETER ConfigPath
    Pfad zur Konfigurationsdatei (Standard: .\config.json)

.PARAMETER WhatIf
    Simuliert die Erstellung ohne tatsächliche Änderungen in OneView.

.EXAMPLE
    .\Create-EthernetNetworks.ps1
    Führt das Skript mit Standardkonfiguration aus.

.EXAMPLE
    .\Create-EthernetNetworks.ps1 -ConfigPath "C:\Config\myconfig.json" -WhatIf
    Simuliert die Erstellung mit benutzerdefinierter Konfiguration.

.NOTES
    Autor:   OneView VLAN Projekt
    Datum:   2026-02-06
    Benötigt: PowerShell 7.x, Modul "ImportExcel"
    API-Ref:  https://support.hpe.com/docs/display/public/dp00006616en_us/index.html
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json")
)

# ============================================================================
#  Konstanten & Initialisierung
# ============================================================================
$ErrorActionPreference = "Stop"
$script:LogEntries = [System.Collections.Generic.List[string]]::new()
$script:CreatedNetworks = 0
$script:UpdatedNetworks = 0
$script:SkippedNetworks = 0
$script:FailedNetworks  = 0

# Gültige Werte gemäss HPE OneView API
$ValidPurposes             = @("General", "Management", "VMMigration", "FaultTolerance", "ISCSI")
$ValidEthernetNetworkTypes = @("Tagged", "Untagged", "Tunnel")

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
    $logsDir = Join-Path $LogDir "Logs"
    if (-not (Test-Path $logsDir)) { New-Item -Path $logsDir -ItemType Directory -Force | Out-Null }
    $logFile = Join-Path $logsDir ("VLAN_Import_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $script:LogEntries | Set-Content -Path $logFile -Encoding UTF8
    Write-Host "`nProtokoll gespeichert: $logFile" -ForegroundColor Cyan
}

# ============================================================================
#  OneView REST API – Session Management
# ============================================================================

function Connect-OneViewAPI {
    <#
    .SYNOPSIS  Authentifizierung via POST /rest/login-sessions
    #>
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
        Write-Log "Authentifizierung fehlgeschlagen für $Hostname : $_" -Level ERROR
        throw
    }
}

function Disconnect-OneViewAPI {
    <#
    .SYNOPSIS  Abmeldung via DELETE /rest/login-sessions
    #>
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
    $start = 0

    do {
        $uri = "$($Session.BaseUri)$ResourcePath" + "?start=$start&count=$PageSize"
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck

        if ($response.members) {
            $allMembers.AddRange([object[]]$response.members)
        }

        $total = $response.total
        $start += $response.members.Count

        Write-Log "  Seite abgerufen: $($allMembers.Count) / $total Einträge" -Level INFO

    } while ($allMembers.Count -lt $total)

    return $allMembers
}

function Get-ExistingEthernetNetworks {
    <#
    .SYNOPSIS  GET /rest/ethernet-networks – alle existierenden Ethernet Networks (paginiert)
    #>
    param([Parameter(Mandatory)][hashtable]$Session)

    Write-Log "Rufe existierende Ethernet Networks ab (paginiert)..."

    try {
        $networks = Get-AllPaginated -Session $Session -ResourcePath "/rest/ethernet-networks"
        Write-Log "  $($networks.Count) existierende Ethernet Networks gefunden." -Level INFO
        return $networks
    }
    catch {
        Write-Log "Fehler beim Abrufen der Ethernet Networks: $_" -Level ERROR
        throw
    }
}

function Get-ExistingNetworkSets {
    <#
    .SYNOPSIS  GET /rest/network-sets – alle existierenden Network Sets (paginiert)
    #>
    param([Parameter(Mandatory)][hashtable]$Session)

    Write-Log "Rufe existierende Network Sets ab (paginiert)..."

    try {
        $sets = Get-AllPaginated -Session $Session -ResourcePath "/rest/network-sets"
        Write-Log "  $($sets.Count) Network Sets gefunden." -Level INFO
        return $sets
    }
    catch {
        Write-Log "Fehler beim Abrufen der Network Sets: $_" -Level WARN
        return @()
    }
}

function Get-ExistingScopes {
    <#
    .SYNOPSIS  GET /rest/scopes – alle existierenden Scopes (paginiert)
    #>
    param([Parameter(Mandatory)][hashtable]$Session)

    Write-Log "Rufe existierende Scopes ab (paginiert)..."

    try {
        $scopes = Get-AllPaginated -Session $Session -ResourcePath "/rest/scopes"
        Write-Log "  $($scopes.Count) Scopes gefunden." -Level INFO
        return $scopes
    }
    catch {
        Write-Log "Fehler beim Abrufen der Scopes: $_" -Level WARN
        return @()
    }
}

# ============================================================================
#  OneView REST API – CREATE / UPDATE Funktionen
# ============================================================================

function New-EthernetNetwork {
    <#
    .SYNOPSIS  POST /rest/ethernet-networks – neues Ethernet Network erstellen
    .DESCRIPTION
        Verwendet Invoke-WebRequest (statt Invoke-RestMethod) um HTTP Status Code
        und Location Header auswerten zu können.

        Mögliche API-Antworten:
        - 201 Created  + Body mit vollständigem Netzwerk-Objekt
        - 201 Created  + Location Header (Body leer oder minimal)
        - 202 Accepted + Task-Objekt im Body oder Location → Task-URI

        Fallback: Netzwerk per Name-Filter suchen.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][hashtable]$NetworkDefinition
    )

    $uri = "$($Session.BaseUri)/rest/ethernet-networks"

    $body = @{
        type                = "ethernet-networkV4"
        name                = $NetworkDefinition.Name
        vlanId              = [int]$NetworkDefinition.VlanId
        purpose             = $NetworkDefinition.Purpose
        ethernetNetworkType = $NetworkDefinition.EthernetNetworkType
        smartLink           = [bool]$NetworkDefinition.SmartLink
        privateNetwork      = [bool]$NetworkDefinition.PrivateNetwork
    }

    # Optionale Felder
    if (-not [string]::IsNullOrWhiteSpace($NetworkDefinition.Description)) {
        $body["description"] = $NetworkDefinition.Description
    }
    if (-not [string]::IsNullOrWhiteSpace($NetworkDefinition.IPv4SubnetId)) {
        $body["subnetUri"] = $NetworkDefinition.IPv4SubnetId
    }
    if (-not [string]::IsNullOrWhiteSpace($NetworkDefinition.IPv6SubnetId)) {
        $body["ipv6SubnetUri"] = $NetworkDefinition.IPv6SubnetId
    }

    $jsonBody = $body | ConvertTo-Json -Depth 5

    Write-Log "  Erstelle Ethernet Network: $($NetworkDefinition.Name) (VLAN $($NetworkDefinition.VlanId))..."

    try {
        # Invoke-WebRequest verwenden um HTTP Status Code und Header auszuwerten
        $webResponse = Invoke-WebRequest -Uri $uri `
            -Method Post `
            -Headers $Session.Headers `
            -Body $jsonBody `
            -SkipCertificateCheck

        $statusCode = $webResponse.StatusCode

        # Response Body parsen (kann leer sein)
        $responseBody = $null
        if ($webResponse.Content -and $webResponse.Content.Trim().Length -gt 0) {
            $responseBody = $webResponse.Content | ConvertFrom-Json
        }

        # Location Header auslesen
        $locationHeader = $null
        if ($webResponse.Headers.ContainsKey("Location")) {
            $locationValues = $webResponse.Headers["Location"]
            $locationHeader = if ($locationValues -is [array]) { $locationValues[0] } else { $locationValues }
            # Absolute URI → relative URI extrahieren
            if ($locationHeader -match "^https?://") {
                $locationHeader = ([System.Uri]$locationHeader).AbsolutePath
            }
        }

        Write-Log ("  API-Antwort: HTTP {0}, Body-Typ: {1}, Location: {2}" -f `
            $statusCode, `
            $(if ($responseBody) { $responseBody.type } else { "(leer)" }), `
            $(if ($locationHeader) { $locationHeader } else { "(keine)" })) -Level INFO

        # --- Fall 1: Synchrone Antwort mit vollständigem Netzwerk-Objekt im Body ---
        if ($responseBody -and $responseBody.uri -and $responseBody.uri -like "/rest/ethernet-networks/*") {
            Write-Log "  Erfolgreich erstellt: $($responseBody.name) (URI: $($responseBody.uri))" -Level SUCCESS
            return $responseBody
        }

        # --- Fall 2: Task-Objekt im Body (asynchrone Verarbeitung, HTTP 202) ---
        $taskUri = $null
        if ($responseBody) {
            if ($responseBody.type -like "*Task*")       { $taskUri = $responseBody.uri }
            elseif ($responseBody.uri -like "/rest/tasks/*") { $taskUri = $responseBody.uri }
            elseif ($responseBody.taskUri)                { $taskUri = $responseBody.taskUri }
        }
        if ($taskUri) {
            Write-Log "  Task erkannt im Body ($taskUri) – warte auf Abschluss..." -Level INFO
            $networkUri = Wait-OneViewTask -Session $Session -TaskUri $taskUri
            $network = Invoke-RestMethod -Uri "$($Session.BaseUri)$networkUri" `
                -Method Get -Headers $Session.Headers -SkipCertificateCheck
            Write-Log "  Erfolgreich erstellt: $($network.name) (URI: $($network.uri))" -Level SUCCESS
            return $network
        }

        # --- Fall 3: Location Header vorhanden (REST-typisch bei 201 ohne Body) ---
        if ($locationHeader) {
            if ($locationHeader -like "/rest/tasks/*") {
                Write-Log "  Task in Location Header ($locationHeader) – warte auf Abschluss..." -Level INFO
                $networkUri = Wait-OneViewTask -Session $Session -TaskUri $locationHeader
                $network = Invoke-RestMethod -Uri "$($Session.BaseUri)$networkUri" `
                    -Method Get -Headers $Session.Headers -SkipCertificateCheck
            }
            else {
                Write-Log "  Rufe Netzwerk von Location Header ab: $locationHeader" -Level INFO
                $network = Invoke-RestMethod -Uri "$($Session.BaseUri)$locationHeader" `
                    -Method Get -Headers $Session.Headers -SkipCertificateCheck
            }
            Write-Log "  Erfolgreich erstellt: $($network.name) (URI: $($network.uri))" -Level SUCCESS
            return $network
        }

        # --- Fall 4: Fallback – Netzwerk per Name suchen ---
        Write-Log "  Kein verwertbarer Body/Location – suche erstelltes Netzwerk per Name..." -Level WARN
        Start-Sleep -Milliseconds 500  # kurz warten damit OneView das Netzwerk indexiert
        $encodedName = [System.Uri]::EscapeDataString($NetworkDefinition.Name)
        $searchUri = "$($Session.BaseUri)/rest/ethernet-networks?filter=name%3D'$encodedName'"
        $searchResult = Invoke-RestMethod -Uri $searchUri -Method Get `
            -Headers $Session.Headers -SkipCertificateCheck

        if ($searchResult.members -and $searchResult.members.Count -gt 0) {
            $network = $searchResult.members[0]
            Write-Log "  Per Suche gefunden: $($network.name) (URI: $($network.uri))" -Level SUCCESS
            return $network
        }

        throw "Netzwerk wurde mglw. erstellt (HTTP $statusCode), konnte aber nicht abgerufen werden."
    }
    catch {
        $errorDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
        $errMsg = if ($errorDetail.message) { $errorDetail.message } else { $_.Exception.Message }
        Write-Log "  Fehler beim Erstellen von '$($NetworkDefinition.Name)': $errMsg" -Level ERROR
        throw
    }
}

function Update-EthernetNetwork {
    <#
    .SYNOPSIS  Aktualisiert ein bestehendes Ethernet Network via PUT /rest/ethernet-networks/{id}
    .DESCRIPTION
        Vergleicht die Soll-Werte aus der Excel-Datei mit dem Ist-Zustand.
        Nur bei Abweichungen wird ein PUT ausgeführt.
        Gibt $true zurück wenn Änderungen vorgenommen wurden, $false wenn identisch.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][PSObject]$ExistingNetwork,
        [Parameter(Mandatory)][hashtable]$NetworkDefinition
    )

    $changes = @()

    # Vergleich der Eigenschaften
    if ([int]$ExistingNetwork.vlanId -ne [int]$NetworkDefinition.VlanId) {
        $changes += "VlanId: $($ExistingNetwork.vlanId) -> $($NetworkDefinition.VlanId)"
    }
    if ($ExistingNetwork.purpose -ne $NetworkDefinition.Purpose) {
        $changes += "Purpose: $($ExistingNetwork.purpose) -> $($NetworkDefinition.Purpose)"
    }
    if ($ExistingNetwork.ethernetNetworkType -ne $NetworkDefinition.EthernetNetworkType) {
        $changes += "EthernetNetworkType: $($ExistingNetwork.ethernetNetworkType) -> $($NetworkDefinition.EthernetNetworkType)"
    }
    if ([bool]$ExistingNetwork.smartLink -ne [bool]$NetworkDefinition.SmartLink) {
        $changes += "SmartLink: $($ExistingNetwork.smartLink) -> $($NetworkDefinition.SmartLink)"
    }
    if ([bool]$ExistingNetwork.privateNetwork -ne [bool]$NetworkDefinition.PrivateNetwork) {
        $changes += "PrivateNetwork: $($ExistingNetwork.privateNetwork) -> $($NetworkDefinition.PrivateNetwork)"
    }

    if ($changes.Count -eq 0) {
        Write-Log "  KEINE ÄNDERUNG: '$($NetworkDefinition.Name)' ist bereits aktuell." -Level INFO
        return $false
    }

    Write-Log "  AKTUALISIERE: '$($NetworkDefinition.Name)' – $($changes.Count) Änderung(en):" -Level INFO
    foreach ($c in $changes) {
        Write-Log "    - $c" -Level INFO
    }

    # Bestehendes Netzwerk-Objekt klonen und Werte aktualisieren
    $uri = "$($Session.BaseUri)$($ExistingNetwork.uri)"

    $updateBody = $ExistingNetwork | ConvertTo-Json -Depth 10 | ConvertFrom-Json
    $updateBody.vlanId              = [int]$NetworkDefinition.VlanId
    $updateBody.purpose             = $NetworkDefinition.Purpose
    $updateBody.ethernetNetworkType = $NetworkDefinition.EthernetNetworkType
    $updateBody.smartLink           = [bool]$NetworkDefinition.SmartLink
    $updateBody.privateNetwork      = [bool]$NetworkDefinition.PrivateNetwork

    $jsonBody = $updateBody | ConvertTo-Json -Depth 10

    $putHeaders = $Session.Headers.Clone()
    $putHeaders["If-Match"] = "*"

    try {
        Invoke-RestMethod -Uri $uri `
            -Method Put `
            -Headers $putHeaders `
            -Body $jsonBody `
            -SkipCertificateCheck | Out-Null

        Write-Log "  Erfolgreich aktualisiert: '$($NetworkDefinition.Name)'" -Level SUCCESS
        return $true
    }
    catch {
        $errorDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
        $errMsg = if ($errorDetail.message) { $errorDetail.message } else { $_.Exception.Message }
        Write-Log "  Fehler beim Aktualisieren von '$($NetworkDefinition.Name)': $errMsg" -Level ERROR
        throw
    }
}

function Set-ConnectionTemplateBandwidth {
    <#
    .SYNOPSIS  Bandwidth setzen via GET + PUT /rest/connection-templates/{id}
    .DESCRIPTION
        Die API erwartet Bandwidth in Mbps.
        Im Excel werden Gb/s angegeben (wie in der OneView GUI).
        Umrechnung: Mbps = Gb/s * 1000
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][string]$ConnectionTemplateUri,
        [Parameter(Mandatory)][double]$PreferredBandwidthGb,
        [Parameter(Mandatory)][double]$MaximumBandwidthGb
    )

    $uri = "$($Session.BaseUri)$ConnectionTemplateUri"
    $typicalMbps = [int]($PreferredBandwidthGb * 1000)
    $maximumMbps = [int]($MaximumBandwidthGb * 1000)

    try {
        # Aktuelles Template abrufen
        $template = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck

        # Bandwidth aktualisieren
        $template.bandwidth.typicalBandwidth = $typicalMbps
        $template.bandwidth.maximumBandwidth = $maximumMbps

        $jsonBody = $template | ConvertTo-Json -Depth 10

        # If-Match: * erzwingt Update ohne ETag-Prüfung (laut API-Doku unterstützt)
        $putHeaders = $Session.Headers.Clone()
        $putHeaders["If-Match"] = "*"

        $updated = Invoke-RestMethod -Uri $uri `
            -Method Put `
            -Headers $putHeaders `
            -Body $jsonBody `
            -SkipCertificateCheck

        Write-Log ("  Bandwidth gesetzt: Preferred={0} Gb/s ({1} Mbps), Maximum={2} Gb/s ({3} Mbps)" -f `
            $PreferredBandwidthGb, $typicalMbps, $MaximumBandwidthGb, $maximumMbps) -Level SUCCESS
        return $updated
    }
    catch {
        $errDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
        $errMsg = if ($errDetail.message) { $errDetail.message } else { $_.Exception.Message }
        Write-Log "  Fehler beim Setzen der Bandwidth: $errMsg" -Level WARN
    }
}

function Add-NetworkToNetworkSet {
    <#
    .SYNOPSIS  Fügt ein Ethernet Network einem bestehenden Network Set hinzu.
    .DESCRIPTION
        1. GET  /rest/network-sets/{id}   – aktuellen Stand abrufen
        2. networkUris um die neue Network-URI ergänzen
        3. PUT  /rest/network-sets/{id}   – aktualisiertes Network Set speichern
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][PSObject]$NetworkSet,
        [Parameter(Mandatory)][string]$NetworkUri
    )

    $setUri = "$($Session.BaseUri)$($NetworkSet.uri)"

    try {
        # Aktuellen Stand abrufen
        $currentSet = Invoke-RestMethod -Uri $setUri -Method Get -Headers $Session.Headers -SkipCertificateCheck

        # Prüfen ob Network bereits enthalten
        if ($currentSet.networkUris -contains $NetworkUri) {
            Write-Log "  Network bereits im Network Set '$($NetworkSet.name)' enthalten." -Level INFO
            return
        }

        # Network URI hinzufügen
        $updatedUris = @($currentSet.networkUris) + $NetworkUri
        $currentSet.networkUris = $updatedUris

        $jsonBody = $currentSet | ConvertTo-Json -Depth 10

        # If-Match: * erzwingt Update ohne ETag-Prüfung (laut API-Doku unterstützt)
        $putHeaders = $Session.Headers.Clone()
        $putHeaders["If-Match"] = "*"

        Invoke-RestMethod -Uri $setUri `
            -Method Put `
            -Headers $putHeaders `
            -Body $jsonBody `
            -SkipCertificateCheck | Out-Null

        Write-Log "  Zum Network Set '$($NetworkSet.name)' hinzugefügt." -Level SUCCESS
    }
    catch {
        $errorDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
        $errMsg = if ($errorDetail.message) { $errorDetail.message } else { $_.Exception.Message }
        Write-Log "  Fehler beim Hinzufügen zum Network Set '$($NetworkSet.name)': $errMsg" -Level WARN
    }
}

function Wait-OneViewTask {
    <#
    .SYNOPSIS  Wartet auf den Abschluss einer asynchronen OneView Task.
    .DESCRIPTION
        POST /rest/ethernet-networks (und andere Ressourcen) antwortet bei längeren
        Operationen mit HTTP 202 Accepted und einem Task-Objekt statt dem fertigen
        Ressourcen-Objekt. Diese Funktion pollt GET /rest/tasks/{id} bis die Task
        den Status "Completed" oder einen Fehlerstatus erreicht.
        Bei Erfolg gibt sie die URI der erstellten Ressource zurück.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][string]$TaskUri,
        [int]$TimeoutSeconds = 120,
        [int]$PollingIntervalMs = 1000
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $uri = "$($Session.BaseUri)$TaskUri"

    while ((Get-Date) -lt $deadline) {
        $task = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck

        switch ($task.taskState) {
            "Completed" {
                # Ressourcen-URI des erstellten Objekts zurückgeben
                $resourceUri = $task.associatedResource.resourceUri
                if ([string]::IsNullOrWhiteSpace($resourceUri)) {
                    throw "Task abgeschlossen, aber keine Ressourcen-URI in 'associatedResource.resourceUri' gefunden."
                }
                return $resourceUri
            }
            { $_ -in @("Error", "Warning", "Terminated", "Killed") } {
                $errMsg = if ($task.taskErrors) { ($task.taskErrors | ForEach-Object { $_.message }) -join "; " } else { $task.taskState }
                throw "OneView Task fehlgeschlagen ($($task.taskState)): $errMsg"
            }
            # "Running", "Starting", "Pending" → weiter warten
        }

        Start-Sleep -Milliseconds $PollingIntervalMs
    }

    throw "Timeout: OneView Task '$TaskUri' nach $TimeoutSeconds Sekunden nicht abgeschlossen."
}

function Add-ResourceToScope {
    <#
    .SYNOPSIS  Weist eine Ressource einem Scope zu.
    .DESCRIPTION
        PATCH /rest/scopes/{id}
        Body: {
            "type":              "ScopePatchDto",
            "addedResourceUris": [ "<network-uri>" ]
        }
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][PSObject]$Scope,
        [Parameter(Mandatory)][string]$ResourceUri
    )

    $scopePatchUri = "$($Session.BaseUri)$($Scope.uri)/resource-assignments"

    $body = @{
        addedResourceUris   = @($ResourceUri)
        removedResourceUris = @()
    } | ConvertTo-Json -Depth 5

    try {
        Invoke-RestMethod -Uri $scopePatchUri `
            -Method Patch `
            -Headers $Session.Headers `
            -Body $body `
            -SkipCertificateCheck | Out-Null

        Write-Log "  Scope '$($Scope.name)' zugewiesen." -Level SUCCESS
    }
    catch {
        $errorDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
        $errMsg = if ($errorDetail.message) { $errorDetail.message } else { $_.Exception.Message }
        Write-Log "  Warnung: Scope-Zuweisung fehlgeschlagen: $errMsg" -Level WARN
    }
}

# ============================================================================
#  Excel Import & Validierung
# ============================================================================

function Import-VlanDataFromExcel {
    param(
        [Parameter(Mandatory)][string]$ExcelPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][hashtable]$Defaults
    )

    if (-not (Test-Path $ExcelPath)) {
        throw "Excel-Datei nicht gefunden: $ExcelPath"
    }

    Write-Log "Lese Excel-Datei: $ExcelPath (Sheet: $SheetName)"

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Log "Installiere Modul 'ImportExcel'..." -Level WARN
        Install-Module -Name ImportExcel -Scope CurrentUser -Force
    }
    Import-Module ImportExcel

    # Versuch 1: Standard-Import (Header in Zeile 1)
    $rawData = Import-Excel -Path $ExcelPath -WorksheetName $SheetName -ErrorAction SilentlyContinue
    
    # Prüfen ob die Spalte "NetworkName" erkannt wurde
    if ($rawData -and $rawData.Count -gt 0) {
        $headers = @($rawData[0].PSObject.Properties.Name)
        if ($headers -notcontains "NetworkName") {
            # Header nicht in Zeile 1 → vermutlich Titelzeile vorhanden, ab Zeile 2 lesen
            Write-Log "  Titelzeile erkannt – lese ab Zeile 2." -Level INFO
            $rawData = Import-Excel -Path $ExcelPath -WorksheetName $SheetName -StartRow 2 -ErrorAction SilentlyContinue
        }
    }

    if (-not $rawData -or $rawData.Count -eq 0) {
        throw "Keine Daten in der Excel-Datei gefunden (Sheet: $SheetName). Bitte prüfen Sie, dass die Spalte 'NetworkName' vorhanden ist."
    }

    Write-Log "  $($rawData.Count) Zeilen aus Excel gelesen."

    $validNetworks = [System.Collections.Generic.List[hashtable]]::new()
    $rowNum = 1

    foreach ($row in $rawData) {
        $rowNum++

        # Name ist Pflicht
        if ([string]::IsNullOrWhiteSpace($row.NetworkName)) {
            Write-Log "  Zeile ${rowNum}: NetworkName fehlt – übersprungen." -Level WARN
            continue
        }

        # VLAN-ID validieren
        $vlanId = 0
        if ($null -ne $row.VlanId -and $row.VlanId.ToString().Trim() -ne "") {
            $vlanId = [int]$row.VlanId
        }

        # EthernetNetworkType
        $ethType = if (-not [string]::IsNullOrWhiteSpace($row.EthernetNetworkType)) {
            $row.EthernetNetworkType.Trim()
        } else { $Defaults.EthernetNetworkType }
        if ($ethType -notin $ValidEthernetNetworkTypes) {
            Write-Log "  Zeile $rowNum ($($row.NetworkName)): Ungültiger EthernetNetworkType '$ethType' – verwende Default." -Level WARN
            $ethType = $Defaults.EthernetNetworkType
        }

        # Für Tagged muss VLAN-ID zwischen 1-4094 sein
        if ($ethType -eq "Tagged" -and ($vlanId -lt 1 -or $vlanId -gt 4094)) {
            Write-Log "  Zeile $rowNum ($($row.NetworkName)): VLAN-ID $vlanId ungültig für Tagged (1-4094) – übersprungen." -Level ERROR
            continue
        }

        # Purpose
        $purpose = if (-not [string]::IsNullOrWhiteSpace($row.Purpose)) {
            $row.Purpose.Trim()
        } else { $Defaults.Purpose }
        if ($purpose -notin $ValidPurposes) {
            Write-Log "  Zeile $rowNum ($($row.NetworkName)): Ungültiger Purpose '$purpose' – verwende Default." -Level WARN
            $purpose = $Defaults.Purpose
        }

        # SmartLink
        $smartLink = if ($null -ne $row.SmartLink -and $row.SmartLink.ToString().Trim() -ne "") {
            [System.Convert]::ToBoolean($row.SmartLink)
        } else { $Defaults.SmartLink }

        # PrivateNetwork
        $privateNetwork = if ($null -ne $row.PrivateNetwork -and $row.PrivateNetwork.ToString().Trim() -ne "") {
            [System.Convert]::ToBoolean($row.PrivateNetwork)
        } else { $Defaults.PrivateNetwork }

        # Bandwidth (Gb/s – wie in der OneView GUI)
        $bwPreferred = if ($null -ne $row.PreferredBandwidthGb -and $row.PreferredBandwidthGb.ToString().Trim() -ne "") {
            [double]$row.PreferredBandwidthGb
        } else { $Defaults.PreferredBandwidthGb }

        $bwMaximum = if ($null -ne $row.MaximumBandwidthGb -and $row.MaximumBandwidthGb.ToString().Trim() -ne "") {
            [double]$row.MaximumBandwidthGb
        } else { $Defaults.MaximumBandwidthGb }

        # Scope (optional)
        $scope = if (-not [string]::IsNullOrWhiteSpace($row.Scope)) {
            $row.Scope.Trim()
        } else { "" }

        # Network Set (optional)
        $networkSet = if (-not [string]::IsNullOrWhiteSpace($row.NetworkSet)) {
            $row.NetworkSet.Trim()
        } else { "" }

        # IPv4 / IPv6 Subnet ID (optional)
        $ipv4SubnetId = if (-not [string]::IsNullOrWhiteSpace($row.IPv4SubnetId)) {
            $row.IPv4SubnetId.Trim()
        } else { "" }

        $ipv6SubnetId = if (-not [string]::IsNullOrWhiteSpace($row.IPv6SubnetId)) {
            $row.IPv6SubnetId.Trim()
        } else { "" }

        # Description
        $description = if (-not [string]::IsNullOrWhiteSpace($row.Description)) {
            $row.Description.Trim()
        } else { "" }

        $validNetworks.Add(@{
            Name                 = $row.NetworkName.Trim()
            VlanId               = $vlanId
            Purpose              = $purpose
            EthernetNetworkType  = $ethType
            SmartLink            = $smartLink
            PrivateNetwork       = $privateNetwork
            PreferredBandwidthGb = $bwPreferred
            MaximumBandwidthGb   = $bwMaximum
            Scope                = $scope
            NetworkSet           = $networkSet
            IPv4SubnetId         = $ipv4SubnetId
            IPv6SubnetId         = $ipv6SubnetId
            Description          = $description
        })
    }

    Write-Log "  $($validNetworks.Count) gültige Netzwerk-Definitionen nach Validierung." -Level INFO
    return $validNetworks
}

# ============================================================================
#  Hauptprogramm
# ============================================================================

function Main {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   HPE OneView – Ethernet Networks aus Excel erstellen       ║" -ForegroundColor Cyan
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

    # -------------------------------------------
    # 2. Excel-Pfad auflösen
    # -------------------------------------------
    $excelPath = $config.ExcelFilePath
    if (-not [System.IO.Path]::IsPathRooted($excelPath)) {
        $excelPath = Join-Path $PSScriptRoot $excelPath
    }
    $excelPath = [System.IO.Path]::GetFullPath($excelPath)

    $defaults = @{
        Purpose              = $config.DefaultSettings.Purpose
        SmartLink            = $config.DefaultSettings.SmartLink
        PrivateNetwork       = $config.DefaultSettings.PrivateNetwork
        EthernetNetworkType  = $config.DefaultSettings.EthernetNetworkType
        PreferredBandwidthGb = $config.DefaultSettings.PreferredBandwidthGb
        MaximumBandwidthGb   = $config.DefaultSettings.MaximumBandwidthGb
    }

    # -------------------------------------------
    # 3. VLAN-Daten aus Excel importieren
    # -------------------------------------------
    try {
        $networks = Import-VlanDataFromExcel -ExcelPath $excelPath `
            -SheetName $config.ExcelSheetName `
            -Defaults $defaults
    }
    catch {
        Write-Log "Fehler beim Excel-Import: $_" -Level ERROR
        Save-Log
        return
    }

    if (-not $networks -or $networks.Count -eq 0) {
        Write-Log "Keine Netzwerke zum Erstellen gefunden." -Level WARN
        Save-Log
        return
    }

    # Zusammenfassung anzeigen
    Write-Host ""
    Write-Host "Folgende Ethernet Networks werden erstellt:" -ForegroundColor Yellow
    Write-Host ("-" * 105)
    Write-Host ("{0,-30} {1,-7} {2,-10} {3,-13} {4,-8} {5,-8} {6,-15} {7,-15}" -f `
        "Name", "VLAN", "Typ", "Purpose", "BW(Gb)", "Max(Gb)", "Scope", "Network Set")
    Write-Host ("-" * 105)
    foreach ($net in $networks) {
        Write-Host ("{0,-30} {1,-7} {2,-10} {3,-13} {4,-8} {5,-8} {6,-15} {7,-15}" -f `
            $net.Name, $net.VlanId, $net.EthernetNetworkType, $net.Purpose, `
            $net.PreferredBandwidthGb, $net.MaximumBandwidthGb, `
            $(if ($net.Scope) { $net.Scope } else { "-" }), `
            $(if ($net.NetworkSet) { $net.NetworkSet } else { "-" }))
    }
    Write-Host ("-" * 105)
    Write-Host "Gesamt: $($networks.Count) Netzwerk(e)`n" -ForegroundColor Yellow

    # -------------------------------------------
    # 4. Benutzerbestätigung
    # -------------------------------------------
    if (-not $WhatIfPreference) {
        $confirm = Read-Host "Möchten Sie fortfahren? (J/N)"
        if ($confirm -notmatch "^[JjYy]") {
            Write-Log "Abbruch durch Benutzer." -Level WARN
            Save-Log
            return
        }
    }

    # -------------------------------------------
    # 5. Anmeldedaten abfragen
    # -------------------------------------------
    Write-Host ""
    $credential = Get-Credential -Message "OneView Anmeldedaten eingeben (Benutzername & Kennwort)"
    if (-not $credential) {
        Write-Log "Keine Anmeldedaten eingegeben – Abbruch." -Level ERROR
        Save-Log
        return
    }

    # -------------------------------------------
    # 6. Für jede Appliance: Netzwerke erstellen
    # -------------------------------------------
    foreach ($appliance in $config.OneViewAppliances) {
        Write-Host ""
        Write-Log "============================================================"
        Write-Log "Appliance: $($appliance.Name) ($($appliance.Hostname))"
        Write-Log "============================================================"

        $session = $null

        try {
            # 6a. API-Version automatisch erkennen & Anmelden
            $detectedApiVersion = Get-OneViewApiVersion -Hostname $appliance.Hostname -FallbackVersion $config.ApiVersion
            $session = Connect-OneViewAPI -Hostname $appliance.Hostname `
                -Credential $credential `
                -ApiVersion $detectedApiVersion

            # 6b. Existierende Ressourcen abrufen
            $existingNetworks = Get-ExistingEthernetNetworks -Session $session
            $existingNetworkLookup = @{}
            foreach ($en in $existingNetworks) {
                $existingNetworkLookup[$en.name] = $en
            }

            $existingNetworkSets = Get-ExistingNetworkSets -Session $session
            $existingScopes      = Get-ExistingScopes -Session $session

            # Lookup-Tabellen für Network Sets und Scopes (Name → Objekt)
            $networkSetLookup = @{}
            foreach ($ns in $existingNetworkSets) {
                $networkSetLookup[$ns.name] = $ns
            }
            $scopeLookup = @{}
            foreach ($sc in $existingScopes) {
                $scopeLookup[$sc.name] = $sc
            }

            # 6c. Netzwerke erstellen / aktualisieren
            foreach ($netDef in $networks) {

                if ($WhatIfPreference) {
                    Write-Log "  [WhatIf] Würde verarbeiten: $($netDef.Name) (VLAN $($netDef.VlanId))" -Level INFO
                    $script:SkippedNetworks++
                    continue
                }

                # Prüfen ob Netzwerk bereits existiert
                $existingNetwork = $null
                if ($existingNetworkLookup.ContainsKey($netDef.Name)) {
                    $existingNetwork = $existingNetworkLookup[$netDef.Name]
                }

                if ($existingNetwork) {
                    # --- Bestehendes Netzwerk: Vergleichen und ggf. aktualisieren ---
                    try {
                        $wasUpdated = Update-EthernetNetwork -Session $session `
                            -ExistingNetwork $existingNetwork `
                            -NetworkDefinition $netDef

                        Start-Sleep -Milliseconds 1500

                        # Bandwidth auch für bestehende Netzwerke prüfen/aktualisieren
                        if ($existingNetwork.connectionTemplateUri) {
                            Set-ConnectionTemplateBandwidth -Session $session `
                                -ConnectionTemplateUri $existingNetwork.connectionTemplateUri `
                                -PreferredBandwidthGb $netDef.PreferredBandwidthGb `
                                -MaximumBandwidthGb $netDef.MaximumBandwidthGb
                        }

                        # Network Set(s) zuweisen (add-only)
                        if (-not [string]::IsNullOrWhiteSpace($netDef.NetworkSet)) {
                            $setNames = $netDef.NetworkSet -split '\s*;\s*'
                            foreach ($setName in $setNames) {
                                $setName = $setName.Trim()
                                if ([string]::IsNullOrWhiteSpace($setName)) { continue }
                                if ($networkSetLookup.ContainsKey($setName)) {
                                    Add-NetworkToNetworkSet -Session $session `
                                        -NetworkSet $networkSetLookup[$setName] `
                                        -NetworkUri $existingNetwork.uri
                                }
                                else {
                                    Write-Log "  Warnung: Network Set '$setName' nicht gefunden – Zuweisung übersprungen." -Level WARN
                                }
                            }
                        }

                        # Scope(s) zuweisen (add-only)
                        if (-not [string]::IsNullOrWhiteSpace($netDef.Scope)) {
                            $scopeNames = $netDef.Scope -split '\s*;\s*'
                            foreach ($scName in $scopeNames) {
                                $scName = $scName.Trim()
                                if ([string]::IsNullOrWhiteSpace($scName)) { continue }
                                if ($scopeLookup.ContainsKey($scName)) {
                                    Add-ResourceToScope -Session $session `
                                        -Scope $scopeLookup[$scName] `
                                        -ResourceUri $existingNetwork.uri
                                }
                                else {
                                    Write-Log "  Warnung: Scope '$scName' nicht gefunden – Zuweisung übersprungen." -Level WARN
                                }
                            }
                        }

                        if ($wasUpdated) {
                            $script:UpdatedNetworks++
                        }
                        else {
                            $script:SkippedNetworks++
                        }
                    }
                    catch {
                        $script:FailedNetworks++
                        continue
                    }
                }
                else {
                    # --- Neues Netzwerk erstellen ---
                    try {
                        # I. Ethernet Network erstellen
                        $createdNetwork = New-EthernetNetwork -Session $session -NetworkDefinition $netDef

                        # Sicherheitsprüfung: Netzwerk-Objekt muss URI enthalten
                        if (-not $createdNetwork -or -not $createdNetwork.uri) {
                            Write-Log "  FEHLER: Kein gültiges Netzwerk-Objekt zurückerhalten – Folgeschritte übersprungen." -Level ERROR
                            $script:FailedNetworks++
                            continue
                        }

                        Start-Sleep -Milliseconds 1500

                        # II. Bandwidth setzen
                        if ($createdNetwork.connectionTemplateUri) {
                            Set-ConnectionTemplateBandwidth -Session $session `
                                -ConnectionTemplateUri $createdNetwork.connectionTemplateUri `
                                -PreferredBandwidthGb $netDef.PreferredBandwidthGb `
                                -MaximumBandwidthGb $netDef.MaximumBandwidthGb
                        }
                        else {
                            Write-Log "  WARNUNG: connectionTemplateUri fehlt im Netzwerk-Objekt – Bandwidth nicht gesetzt." -Level WARN
                        }

                        # III. Network Set(s) zuweisen
                        if (-not [string]::IsNullOrWhiteSpace($netDef.NetworkSet)) {
                            $setNames = $netDef.NetworkSet -split '\s*;\s*'
                            foreach ($setName in $setNames) {
                                $setName = $setName.Trim()
                                if ([string]::IsNullOrWhiteSpace($setName)) { continue }
                                if ($networkSetLookup.ContainsKey($setName)) {
                                    Add-NetworkToNetworkSet -Session $session `
                                        -NetworkSet $networkSetLookup[$setName] `
                                        -NetworkUri $createdNetwork.uri
                                }
                                else {
                                    Write-Log "  Warnung: Network Set '$setName' nicht gefunden – Zuweisung übersprungen." -Level WARN
                                }
                            }
                        }

                        # IV. Scope(s) zuweisen
                        if (-not [string]::IsNullOrWhiteSpace($netDef.Scope)) {
                            $scopeNames = $netDef.Scope -split '\s*;\s*'
                            foreach ($scName in $scopeNames) {
                                $scName = $scName.Trim()
                                if ([string]::IsNullOrWhiteSpace($scName)) { continue }
                                if ($scopeLookup.ContainsKey($scName)) {
                                    Add-ResourceToScope -Session $session `
                                        -Scope $scopeLookup[$scName] `
                                        -ResourceUri $createdNetwork.uri
                                }
                                else {
                                    Write-Log "  Warnung: Scope '$scName' nicht gefunden – Zuweisung übersprungen." -Level WARN
                                }
                            }
                        }

                        $script:CreatedNetworks++
                    }
                    catch {
                        $script:FailedNetworks++
                        continue
                    }
                }
            }
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

    # -------------------------------------------
    # 7. Zusammenfassung
    # -------------------------------------------
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                    Zusammenfassung                           ║" -ForegroundColor Cyan
    Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
    Write-Host ("║  Erstellt:       {0,-43}║" -f $script:CreatedNetworks) -ForegroundColor Green
    Write-Host ("║  Aktualisiert:   {0,-43}║" -f $script:UpdatedNetworks) -ForegroundColor Cyan
    Write-Host ("║  Übersprungen:   {0,-43}║" -f $script:SkippedNetworks) -ForegroundColor Yellow
    Write-Host ("║  Fehlgeschlagen: {0,-43}║" -f $script:FailedNetworks) -ForegroundColor Red
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    Save-Log
}

# Skript starten
Main
