<#
.SYNOPSIS
    Erstellt und aktualisiert Network Sets in HPE OneView basierend auf einer Excel-Datei.

.DESCRIPTION
    Dieses Skript liest Network-Set-Definitionen aus einer Excel-Datei und erstellt
    bzw. aktualisiert die entsprechenden Network Sets über die HPE OneView RESTful API.

    Unterstützte Felder:
    - NetworkSetName (Pflicht)
    - Networks (Pflicht – Semikolon-getrennte Liste von Ethernet Network Namen)
    - NativeNetwork (optional – muss eines der Networks sein)
    - Preferred / Maximum Bandwidth (in Gb/s, wie in der GUI)
    - Scope (Zuweisung zu vorhandenem Scope)
    - Description

    Ablauf:
    1. Konfiguration aus config.json laden
    2. Network-Set-Daten aus Excel importieren & validieren
    3. Authentifizierung an der OneView Appliance (REST API)
    4. Existierende Network Sets, Ethernet Networks und Scopes abrufen
    5. Network Sets erstellen (POST /rest/network-sets) oder aktualisieren (PUT)
    6. Bandwidth setzen (PUT /rest/connection-templates/{id})
    7. Scope zuweisen (PATCH /rest/scopes/{id})
    8. Session abmelden & Protokoll speichern

.PARAMETER ConfigPath
    Pfad zur Konfigurationsdatei (Standard: .\config.json)

.PARAMETER WhatIf
    Simuliert die Erstellung ohne tatsächliche Änderungen in OneView.

.EXAMPLE
    .\Create-NetworkSets.ps1
    Führt das Skript mit Standardkonfiguration aus.

.EXAMPLE
    .\Create-NetworkSets.ps1 -ConfigPath "C:\Config\myconfig.json" -WhatIf
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
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),

    [Parameter()]
    [string]$LogPath = ""
)

# ============================================================================
#  Konstanten & Initialisierung
# ============================================================================
$ErrorActionPreference = "Stop"
$script:LogEntries = [System.Collections.Generic.List[string]]::new()
$script:LogPath = $LogPath
$script:CreatedSets  = 0
$script:UpdatedSets  = 0
$script:SkippedSets  = 0
$script:FailedSets   = 0

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
        $logFile = Join-Path $logsDir ("NetworkSet_Import_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
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
        Write-Log "Authentifizierung fehlgeschlagen für $Hostname : $_" -Level ERROR
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

function Get-ExistingEthernetNetworks {
    param([Parameter(Mandatory)][hashtable]$Session)

    Write-Log "Rufe existierende Ethernet Networks ab (paginiert)..."

    try {
        $networks = Get-AllPaginated -Session $Session -ResourcePath "/rest/ethernet-networks"
        Write-Log "  $($networks.Count) Ethernet Networks gefunden." -Level INFO
        return $networks
    }
    catch {
        Write-Log "Fehler beim Abrufen der Ethernet Networks: $_" -Level ERROR
        throw
    }
}

function Get-ExistingNetworkSets {
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

function New-NetworkSet {
    <#
    .SYNOPSIS  POST /rest/network-sets – neues Network Set erstellen
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][hashtable]$SetDefinition
    )

    $uri = "$($Session.BaseUri)/rest/network-sets"

    $body = @{
        type            = "network-setV5"
        name            = $SetDefinition.Name
        networkUris     = @($SetDefinition.NetworkUris)
        networkSetType  = $SetDefinition.NetworkSetType
    }

    if (-not [string]::IsNullOrWhiteSpace($SetDefinition.NativeNetworkUri)) {
        $body["nativeNetworkUri"] = $SetDefinition.NativeNetworkUri
    }

    if (-not [string]::IsNullOrWhiteSpace($SetDefinition.Description)) {
        $body["description"] = $SetDefinition.Description
    }

    $jsonBody = $body | ConvertTo-Json -Depth 5

    Write-Log "  Erstelle Network Set: $($SetDefinition.Name) ($($SetDefinition.NetworkUris.Count) Netzwerke)..."

    try {
        $webResponse = Invoke-WebRequest -Uri $uri `
            -Method Post `
            -Headers $Session.Headers `
            -Body $jsonBody `
            -SkipCertificateCheck

        $statusCode = $webResponse.StatusCode

        $responseBody = $null
        if ($webResponse.Content -and $webResponse.Content.Trim().Length -gt 0) {
            $responseBody = $webResponse.Content | ConvertFrom-Json
        }

        $locationHeader = $null
        if ($webResponse.Headers.ContainsKey("Location")) {
            $locationValues = $webResponse.Headers["Location"]
            $locationHeader = if ($locationValues -is [array]) { $locationValues[0] } else { $locationValues }
            if ($locationHeader -match "^https?://") {
                $locationHeader = ([System.Uri]$locationHeader).AbsolutePath
            }
        }

        Write-Log ("  API-Antwort: HTTP {0}, Body-Typ: {1}, Location: {2}" -f `
            $statusCode, `
            $(if ($responseBody) { $responseBody.type } else { "(leer)" }), `
            $(if ($locationHeader) { $locationHeader } else { "(keine)" })) -Level INFO

        # Fall 1: Synchrone Antwort mit vollständigem Objekt im Body
        if ($responseBody -and $responseBody.uri -and $responseBody.uri -like "/rest/network-sets/*") {
            Write-Log "  Erfolgreich erstellt: $($responseBody.name) (URI: $($responseBody.uri))" -Level SUCCESS
            return $responseBody
        }

        # Fall 2: Task-Objekt im Body (asynchrone Verarbeitung)
        $taskUri = $null
        if ($responseBody) {
            if ($responseBody.type -like "*Task*")           { $taskUri = $responseBody.uri }
            elseif ($responseBody.uri -like "/rest/tasks/*")  { $taskUri = $responseBody.uri }
            elseif ($responseBody.taskUri)                    { $taskUri = $responseBody.taskUri }
        }
        if ($taskUri) {
            Write-Log "  Task erkannt im Body ($taskUri) – warte auf Abschluss..." -Level INFO
            $nsUri = Wait-OneViewTask -Session $Session -TaskUri $taskUri
            $ns = Invoke-RestMethod -Uri "$($Session.BaseUri)$nsUri" `
                -Method Get -Headers $Session.Headers -SkipCertificateCheck
            Write-Log "  Erfolgreich erstellt: $($ns.name) (URI: $($ns.uri))" -Level SUCCESS
            return $ns
        }

        # Fall 3: Location Header
        if ($locationHeader) {
            if ($locationHeader -like "/rest/tasks/*") {
                Write-Log "  Task in Location Header ($locationHeader) – warte auf Abschluss..." -Level INFO
                $nsUri = Wait-OneViewTask -Session $Session -TaskUri $locationHeader
                $ns = Invoke-RestMethod -Uri "$($Session.BaseUri)$nsUri" `
                    -Method Get -Headers $Session.Headers -SkipCertificateCheck
            }
            else {
                Write-Log "  Rufe Network Set von Location Header ab: $locationHeader" -Level INFO
                $ns = Invoke-RestMethod -Uri "$($Session.BaseUri)$locationHeader" `
                    -Method Get -Headers $Session.Headers -SkipCertificateCheck
            }
            Write-Log "  Erfolgreich erstellt: $($ns.name) (URI: $($ns.uri))" -Level SUCCESS
            return $ns
        }

        # Fall 4: Fallback – per Name suchen
        Write-Log "  Kein verwertbarer Body/Location – suche erstelltes Network Set per Name..." -Level WARN
        Start-Sleep -Milliseconds 500
        $encodedName = [System.Uri]::EscapeDataString($SetDefinition.Name)
        $searchUri = "$($Session.BaseUri)/rest/network-sets?filter=name%3D'$encodedName'"
        $searchResult = Invoke-RestMethod -Uri $searchUri -Method Get `
            -Headers $Session.Headers -SkipCertificateCheck

        if ($searchResult.members -and $searchResult.members.Count -gt 0) {
            $ns = $searchResult.members[0]
            Write-Log "  Per Suche gefunden: $($ns.name) (URI: $($ns.uri))" -Level SUCCESS
            return $ns
        }

        throw "Network Set wurde mglw. erstellt (HTTP $statusCode), konnte aber nicht abgerufen werden."
    }
    catch {
        $errorDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
        $errMsg = if ($errorDetail.message) { $errorDetail.message } else { $_.Exception.Message }
        Write-Log "  Fehler beim Erstellen von '$($SetDefinition.Name)': $errMsg" -Level ERROR
        throw
    }
}

function Update-NetworkSet {
    <#
    .SYNOPSIS  Aktualisiert ein bestehendes Network Set via PUT /rest/network-sets/{id}
    .DESCRIPTION
        Vergleicht die Soll-Werte aus der Excel-Datei mit dem Ist-Zustand.
        Nur bei Abweichungen wird ein PUT ausgeführt.
        Gibt $true zurück wenn Änderungen vorgenommen wurden, $false wenn identisch.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][PSObject]$ExistingSet,
        [Parameter(Mandatory)][hashtable]$SetDefinition
    )

    $changes = @()

    # Vergleich networkUris (sortierte Mengenvergleich)
    $existingUris = @($ExistingSet.networkUris | Sort-Object)
    $desiredUris  = @($SetDefinition.NetworkUris | Sort-Object)
    $existingJoined = ($existingUris -join ",")
    $desiredJoined  = ($desiredUris -join ",")

    if ($existingJoined -ne $desiredJoined) {
        $addedCount   = ($desiredUris | Where-Object { $_ -notin $existingUris }).Count
        $removedCount = ($existingUris | Where-Object { $_ -notin $desiredUris }).Count
        $changes += "Networks: $addedCount hinzugefügt, $removedCount entfernt"
    }

    # Vergleich nativeNetworkUri
    $existingNative = if ($ExistingSet.nativeNetworkUri) { $ExistingSet.nativeNetworkUri } else { "" }
    $desiredNative  = if ($SetDefinition.NativeNetworkUri) { $SetDefinition.NativeNetworkUri } else { "" }
    if ($existingNative -ne $desiredNative) {
        $changes += "NativeNetwork: geändert"
    }

    # Vergleich Description
    $existingDesc = if ($ExistingSet.description) { $ExistingSet.description } else { "" }
    $desiredDesc  = if ($SetDefinition.Description) { $SetDefinition.Description } else { "" }
    if ($existingDesc -ne $desiredDesc) {
        $changes += "Description: '$existingDesc' -> '$desiredDesc'"
    }

    # Vergleich networkSetType
    $existingType = if ($ExistingSet.networkSetType) { $ExistingSet.networkSetType } else { "Regular" }
    $desiredType  = if ($SetDefinition.NetworkSetType) { $SetDefinition.NetworkSetType } else { "Regular" }
    if ($existingType -ne $desiredType) {
        $changes += "NetworkSetType: '$existingType' -> '$desiredType'"
    }

    if ($changes.Count -eq 0) {
        Write-Log "  KEINE ÄNDERUNG: '$($SetDefinition.Name)' ist bereits aktuell." -Level INFO
        return $false
    }

    Write-Log "  AKTUALISIERE: '$($SetDefinition.Name)' – $($changes.Count) Änderung(en):" -Level INFO
    foreach ($c in $changes) {
        Write-Log "    - $c" -Level INFO
    }

    # Bestehendes Objekt klonen und aktualisieren
    $uri = "$($Session.BaseUri)$($ExistingSet.uri)"

    $updateBody = $ExistingSet | ConvertTo-Json -Depth 10 | ConvertFrom-Json
    $updateBody.networkUris = @($SetDefinition.NetworkUris)

    if (-not [string]::IsNullOrWhiteSpace($SetDefinition.NativeNetworkUri)) {
        $updateBody.nativeNetworkUri = $SetDefinition.NativeNetworkUri
    }
    else {
        $updateBody.nativeNetworkUri = $null
    }

    if (-not [string]::IsNullOrWhiteSpace($SetDefinition.Description)) {
        $updateBody.description = $SetDefinition.Description
    }
    else {
        $updateBody.description = ""
    }

    # NetworkSetType aktualisieren
    $updateBody.networkSetType = if ($SetDefinition.NetworkSetType) { $SetDefinition.NetworkSetType } else { "Regular" }

    $jsonBody = $updateBody | ConvertTo-Json -Depth 10

    $putHeaders = $Session.Headers.Clone()
    $putHeaders["If-Match"] = "*"

    try {
        Invoke-RestMethod -Uri $uri `
            -Method Put `
            -Headers $putHeaders `
            -Body $jsonBody `
            -SkipCertificateCheck | Out-Null

        Write-Log "  Erfolgreich aktualisiert: '$($SetDefinition.Name)'" -Level SUCCESS
        return $true
    }
    catch {
        $errorDetail = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
        $errMsg = if ($errorDetail.message) { $errorDetail.message } else { $_.Exception.Message }
        Write-Log "  Fehler beim Aktualisieren von '$($SetDefinition.Name)': $errMsg" -Level ERROR
        throw
    }
}

function Set-ConnectionTemplateBandwidth {
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
        $template = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck

        $template.bandwidth.typicalBandwidth = $typicalMbps
        $template.bandwidth.maximumBandwidth = $maximumMbps

        $jsonBody = $template | ConvertTo-Json -Depth 10

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

function Wait-OneViewTask {
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
        }

        Start-Sleep -Milliseconds $PollingIntervalMs
    }

    throw "Timeout: OneView Task '$TaskUri' nach $TimeoutSeconds Sekunden nicht abgeschlossen."
}

function Add-ResourceToScope {
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

function Import-NetworkSetDataFromExcel {
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

    # Prüfen ob die Spalte "NetworkSetName" erkannt wurde
    if ($rawData -and $rawData.Count -gt 0) {
        $headers = @($rawData[0].PSObject.Properties.Name)
        if ($headers -notcontains "NetworkSetName") {
            Write-Log "  Titelzeile erkannt – lese ab Zeile 2." -Level INFO
            $rawData = Import-Excel -Path $ExcelPath -WorksheetName $SheetName -StartRow 2 -ErrorAction SilentlyContinue
        }
    }

    if (-not $rawData -or $rawData.Count -eq 0) {
        throw "Keine Daten in der Excel-Datei gefunden (Sheet: $SheetName). Bitte prüfen Sie, dass die Spalte 'NetworkSetName' vorhanden ist."
    }

    Write-Log "  $($rawData.Count) Zeilen aus Excel gelesen."

    $validSets = [System.Collections.Generic.List[hashtable]]::new()
    $rowNum = 1

    foreach ($row in $rawData) {
        $rowNum++

        # Name ist Pflicht
        if ([string]::IsNullOrWhiteSpace($row.NetworkSetName)) {
            Write-Log "  Zeile ${rowNum}: NetworkSetName fehlt – übersprungen." -Level WARN
            continue
        }

        # Networks ist Pflicht (Semikolon-getrennte Liste)
        if ([string]::IsNullOrWhiteSpace($row.Networks)) {
            Write-Log "  Zeile ${rowNum} ($($row.NetworkSetName)): Networks fehlt – übersprungen." -Level WARN
            continue
        }

        $networkNames = @($row.Networks -split '\s*;\s*' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
        if ($networkNames.Count -eq 0) {
            Write-Log "  Zeile ${rowNum} ($($row.NetworkSetName)): Keine gültigen Netzwerk-Namen – übersprungen." -Level WARN
            continue
        }

        # NativeNetwork (optional)
        $nativeNetwork = if (-not [string]::IsNullOrWhiteSpace($row.NativeNetwork)) {
            $row.NativeNetwork.Trim()
        } else { "" }

        # Bandwidth (Gb/s)
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

        # Description
        $description = if (-not [string]::IsNullOrWhiteSpace($row.Description)) {
            $row.Description.Trim()
        } else { "" }

        # NetworkSetType (Regular oder Large, Standard: Regular)
        $networkSetType = if (-not [string]::IsNullOrWhiteSpace($row.NetworkSetType)) {
            $row.NetworkSetType.Trim()
        } else { "Regular" }

        $validSets.Add(@{
            Name                 = $row.NetworkSetName.Trim()
            NetworkNames         = $networkNames
            NativeNetwork        = $nativeNetwork
            PreferredBandwidthGb = $bwPreferred
            MaximumBandwidthGb   = $bwMaximum
            Scope                = $scope
            Description          = $description
            NetworkSetType       = $networkSetType
        })
    }

    Write-Log "  $($validSets.Count) gültige Network-Set-Definitionen nach Validierung." -Level INFO
    return $validSets
}

# ============================================================================
#  Hauptprogramm
# ============================================================================

function Main {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   HPE OneView – Network Sets aus Excel erstellen            ║" -ForegroundColor Cyan
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
    $excelPath = if ($config.NetworkSetExcelFilePath) { $config.NetworkSetExcelFilePath } else { $config.ExcelFilePath }
    if (-not [System.IO.Path]::IsPathRooted($excelPath)) {
        $excelPath = Join-Path $PSScriptRoot $excelPath
    }
    $excelPath = [System.IO.Path]::GetFullPath($excelPath)

    $sheetName = if ($config.NetworkSetExcelSheetName) { $config.NetworkSetExcelSheetName } else { "NetworkSets" }

    $defaults = @{
        PreferredBandwidthGb = 2.5
        MaximumBandwidthGb   = 20
    }
    if ($config.NetworkSetDefaultSettings) {
        if ($null -ne $config.NetworkSetDefaultSettings.PreferredBandwidthGb) {
            $defaults.PreferredBandwidthGb = $config.NetworkSetDefaultSettings.PreferredBandwidthGb
        }
        if ($null -ne $config.NetworkSetDefaultSettings.MaximumBandwidthGb) {
            $defaults.MaximumBandwidthGb = $config.NetworkSetDefaultSettings.MaximumBandwidthGb
        }
    }

    # -------------------------------------------
    # 3. Network-Set-Daten aus Excel importieren
    # -------------------------------------------
    try {
        $networkSets = Import-NetworkSetDataFromExcel -ExcelPath $excelPath `
            -SheetName $sheetName `
            -Defaults $defaults
    }
    catch {
        Write-Log "Fehler beim Excel-Import: $_" -Level ERROR
        Save-Log
        return
    }

    if (-not $networkSets -or $networkSets.Count -eq 0) {
        Write-Log "Keine Network Sets zum Erstellen gefunden." -Level WARN
        Save-Log
        return
    }

    # Zusammenfassung anzeigen
    Write-Host ""
    Write-Host "Folgende Network Sets werden verarbeitet:" -ForegroundColor Yellow
    Write-Host ("-" * 105)
    Write-Host ("{0,-30} {1,-40} {2,-20} {3,-8}" -f `
        "Name", "Networks", "Native", "BW(Gb)")
    Write-Host ("-" * 105)
    foreach ($setDef in $networkSets) {
        $netStr = ($setDef.NetworkNames -join "; ")
        if ($netStr.Length -gt 37) { $netStr = $netStr.Substring(0, 37) + "..." }
        $nativeStr = if ($setDef.NativeNetwork) { $setDef.NativeNetwork } else { "-" }
        if ($nativeStr.Length -gt 17) { $nativeStr = $nativeStr.Substring(0, 17) + "..." }
        Write-Host ("{0,-30} {1,-40} {2,-20} {3,-8}" -f `
            $setDef.Name, $netStr, $nativeStr, $setDef.PreferredBandwidthGb)
    }
    Write-Host ("-" * 105)
    Write-Host "Gesamt: $($networkSets.Count) Network Set(s)`n" -ForegroundColor Yellow

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
    # 6. Für jede Appliance: Network Sets erstellen / aktualisieren
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
            $existingNetworkSets = Get-ExistingNetworkSets -Session $session
            $existingScopes = Get-ExistingScopes -Session $session

            # Lookup-Tabellen
            $networkNameToUri = @{}
            foreach ($net in $existingNetworks) {
                $networkNameToUri[$net.name] = $net.uri
            }

            $networkSetLookup = @{}
            foreach ($ns in $existingNetworkSets) {
                $networkSetLookup[$ns.name] = $ns
            }

            $scopeLookup = @{}
            foreach ($sc in $existingScopes) {
                $scopeLookup[$sc.name] = $sc
            }

            # 6c. Network Sets verarbeiten
            foreach ($setDef in $networkSets) {

                if ($WhatIfPreference) {
                    Write-Log "  [WhatIf] Würde verarbeiten: $($setDef.Name)" -Level INFO
                    $script:SkippedSets++
                    continue
                }

                # Netzwerk-Namen zu URIs auflösen
                $resolvedUris = [System.Collections.Generic.List[string]]::new()
                $unresolvedNames = @()
                foreach ($netName in $setDef.NetworkNames) {
                    if ($networkNameToUri.ContainsKey($netName)) {
                        $resolvedUris.Add($networkNameToUri[$netName])
                    }
                    else {
                        $unresolvedNames += $netName
                    }
                }

                if ($unresolvedNames.Count -gt 0) {
                    Write-Log "  Warnung: $($unresolvedNames.Count) Netzwerk(e) nicht gefunden: $($unresolvedNames -join ', ')" -Level WARN
                }

                if ($resolvedUris.Count -eq 0) {
                    Write-Log "  FEHLER: Kein einziges Netzwerk aufgelöst für '$($setDef.Name)' – übersprungen." -Level ERROR
                    $script:FailedSets++
                    continue
                }

                # Native Network auflösen
                $nativeUri = ""
                if (-not [string]::IsNullOrWhiteSpace($setDef.NativeNetwork)) {
                    if ($networkNameToUri.ContainsKey($setDef.NativeNetwork)) {
                        $nativeUri = $networkNameToUri[$setDef.NativeNetwork]
                        # Sicherstellen, dass Native in der Netzwerk-Liste ist
                        if ($nativeUri -notin $resolvedUris) {
                            $resolvedUris.Add($nativeUri)
                            Write-Log "  Native Network '$($setDef.NativeNetwork)' wurde automatisch zur Netzwerk-Liste hinzugefügt." -Level INFO
                        }
                    }
                    else {
                        Write-Log "  Warnung: Native Network '$($setDef.NativeNetwork)' nicht gefunden – wird nicht gesetzt." -Level WARN
                    }
                }

                $setDefResolved = @{
                    Name                 = $setDef.Name
                    NetworkUris          = @($resolvedUris)
                    NativeNetworkUri     = $nativeUri
                    PreferredBandwidthGb = $setDef.PreferredBandwidthGb
                    MaximumBandwidthGb   = $setDef.MaximumBandwidthGb
                    Scope                = $setDef.Scope
                    Description          = $setDef.Description
                }

                # Prüfen ob Network Set bereits existiert
                $existingSet = $null
                if ($networkSetLookup.ContainsKey($setDef.Name)) {
                    $existingSet = $networkSetLookup[$setDef.Name]
                }

                if ($existingSet) {
                    # --- Bestehendes Network Set: Vergleichen und ggf. aktualisieren ---
                    try {
                        $wasUpdated = Update-NetworkSet -Session $session `
                            -ExistingSet $existingSet `
                            -SetDefinition $setDefResolved

                        Start-Sleep -Milliseconds 1500

                        # Bandwidth prüfen/aktualisieren
                        if ($existingSet.connectionTemplateUri) {
                            Set-ConnectionTemplateBandwidth -Session $session `
                                -ConnectionTemplateUri $existingSet.connectionTemplateUri `
                                -PreferredBandwidthGb $setDef.PreferredBandwidthGb `
                                -MaximumBandwidthGb $setDef.MaximumBandwidthGb
                        }

                        # Scope(s) zuweisen (add-only)
                        if (-not [string]::IsNullOrWhiteSpace($setDef.Scope)) {
                            $scopeNames = $setDef.Scope -split '\s*;\s*'
                            foreach ($scName in $scopeNames) {
                                $scName = $scName.Trim()
                                if ([string]::IsNullOrWhiteSpace($scName)) { continue }
                                if ($scopeLookup.ContainsKey($scName)) {
                                    Add-ResourceToScope -Session $session `
                                        -Scope $scopeLookup[$scName] `
                                        -ResourceUri $existingSet.uri
                                }
                                else {
                                    Write-Log "  Warnung: Scope '$scName' nicht gefunden – Zuweisung übersprungen." -Level WARN
                                }
                            }
                        }

                        if ($wasUpdated) {
                            $script:UpdatedSets++
                        }
                        else {
                            $script:SkippedSets++
                        }
                    }
                    catch {
                        $script:FailedSets++
                        continue
                    }
                }
                else {
                    # --- Neues Network Set erstellen ---
                    try {
                        # I. Network Set erstellen
                        $createdSet = New-NetworkSet -Session $session -SetDefinition $setDefResolved

                        if (-not $createdSet -or -not $createdSet.uri) {
                            Write-Log "  FEHLER: Kein gültiges Network-Set-Objekt zurückerhalten – Folgeschritte übersprungen." -Level ERROR
                            $script:FailedSets++
                            continue
                        }

                        Start-Sleep -Milliseconds 1500

                        # II. Bandwidth setzen
                        if ($createdSet.connectionTemplateUri) {
                            Set-ConnectionTemplateBandwidth -Session $session `
                                -ConnectionTemplateUri $createdSet.connectionTemplateUri `
                                -PreferredBandwidthGb $setDef.PreferredBandwidthGb `
                                -MaximumBandwidthGb $setDef.MaximumBandwidthGb
                        }
                        else {
                            Write-Log "  WARNUNG: connectionTemplateUri fehlt im Network-Set-Objekt – Bandwidth nicht gesetzt." -Level WARN
                        }

                        # III. Scope(s) zuweisen
                        if (-not [string]::IsNullOrWhiteSpace($setDef.Scope)) {
                            $scopeNames = $setDef.Scope -split '\s*;\s*'
                            foreach ($scName in $scopeNames) {
                                $scName = $scName.Trim()
                                if ([string]::IsNullOrWhiteSpace($scName)) { continue }
                                if ($scopeLookup.ContainsKey($scName)) {
                                    Add-ResourceToScope -Session $session `
                                        -Scope $scopeLookup[$scName] `
                                        -ResourceUri $createdSet.uri
                                }
                                else {
                                    Write-Log "  Warnung: Scope '$scName' nicht gefunden – Zuweisung übersprungen." -Level WARN
                                }
                            }
                        }

                        $script:CreatedSets++
                    }
                    catch {
                        $script:FailedSets++
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
    Write-Host ("║  Erstellt:       {0,-43}║" -f $script:CreatedSets) -ForegroundColor Green
    Write-Host ("║  Aktualisiert:   {0,-43}║" -f $script:UpdatedSets) -ForegroundColor Cyan
    Write-Host ("║  Übersprungen:   {0,-43}║" -f $script:SkippedSets) -ForegroundColor Yellow
    Write-Host ("║  Fehlgeschlagen: {0,-43}║" -f $script:FailedSets) -ForegroundColor Red
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    Save-Log
}

# Skript starten
Main
