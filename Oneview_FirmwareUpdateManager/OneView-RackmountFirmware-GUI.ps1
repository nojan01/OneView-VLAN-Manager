#Requires -Version 7.5
param(
    [ValidateSet('Rackmount', 'Synergy')]
    [string]$ServerMode = 'Rackmount',
    [string]$ApplianceFile
)

# ============================================================================
#  HPE ProLiant Rackmount Firmware Update GUI  (Gen10 / Gen10 Plus / Gen11 / Gen12+)
#  ------------------------------------------------------------------------
#  Aktualisiert Firmware von STANDALONE Rackmount-Servern direkt ueber die
#  iLO-Redfish-API (iLO 5 = Gen10/Gen10 Plus, iLO 6 = Gen11, iLO 7 = Gen12).
#
#  Hintergrund:
#   Die Server sind in HPE OneView nur im MONITORING-Mode eingebunden. OneView
#   kann monitored Server NICHT flashen (das erfordert Managed-Mode + Profil).
#   Daher wird hier direkt gegen das iLO geflasht - der von HPE empfohlene,
#   sichere Weg fuer Standalone-Updates ueber MultipartHttpPushUri.
#
#  Funktionen:
#   - Inventar pruefen (read-only): Erreichbarkeit, Modell/Generation,
#     aktuelle iLO-/BIOS-/Komponenten-Versionen pro Server
#   - Firmware flashen: eine einzelne .fwpkg ODER ein Ordner mit mehreren
#     .fwpkg nacheinander, auf einem Server oder einer ganzen Charge parallel
#   - Optionale SHA-256-Verifikation der Firmware-Datei vor dem Upload
#   - Parallele Verarbeitung (Runspace-Pool) mit einstellbarer Parallelitaet
#
#  Sicherheit:
#   - Es wird NICHT automatisch rebootet. BIOS/Systemfirmware wird in den
#     Pending/Redundant-Bereich geschrieben und beim naechsten regulaeren
#     Neustart (Wartungsfenster) aktiv.
#   - iLO-Firmware wird (falls im Stapel) zuletzt geflasht; nach dem iLO-Reset
#     wird die Session automatisch neu aufgebaut.
#
#  Redfish-Endpunkte:
#   POST   /redfish/v1/SessionService/Sessions        (Login -> X-Auth-Token)
#   GET    /redfish/v1/Systems/1                       (Modell / Generation)
#   GET    /redfish/v1/UpdateService                   (State / MultipartHttpPushUri)
#   GET    /redfish/v1/UpdateService/FirmwareInventory (aktuelle Versionen)
#   POST   {MultipartHttpPushUri}                      (Upload + Flash, multipart)
#   DELETE /redfish/v1/SessionService/Sessions/{id}    (Logout)
# ============================================================================

$scriptFolder = $PSScriptRoot
$script:IsSynergyMode = $ServerMode -eq 'Synergy'
# Die Build-Kennung wird in jeder gestarteten Aktion geloggt. So ist bei
# Diagnoseexporten eindeutig, dass die GUI nach einer Aenderung wirklich
# beendet und aus dieser Skriptversion neu gestartet wurde.
$script:BuildId = '2026.08.14.50'

# =============================
# Konsolenfenster ausblenden + DPI
# =============================
if (-not ([System.Management.Automation.PSTypeName]::new("Win32.NativeMethods").Type)) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace Win32 {
    public static class NativeMethods {
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();
    }
}
"@
}
try { [Win32.NativeMethods]::SetProcessDPIAware() | Out-Null } catch {}
$consolePtr = [Win32.NativeMethods]::GetConsoleWindow()
if ($consolePtr -ne [System.IntPtr]::Zero) {
    [Win32.NativeMethods]::ShowWindow($consolePtr, 0)
}

# =============================
# Assemblies
# =============================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net.Http

# =============================
# Logs-Verzeichnis
# =============================
$script:logFolder = Join-Path $scriptFolder "Logs"
if (-not (Test-Path $script:logFolder)) { New-Item -ItemType Directory -Path $script:logFolder -Force | Out-Null }
$logPrefix = if ($script:IsSynergyMode) { 'SynergyFirmware' } else { 'RackmountFirmware' }
$script:logFile = Join-Path $script:logFolder ("{0}_{1:yyyy-MM-dd_HHmmss}.log" -f $logPrefix, (Get-Date))
function Write-LogFile {
    param([string]$Text, [datetime]$Timestamp = (Get-Date))
    try { Add-Content -LiteralPath $script:logFile -Value ("{0:yyyy-MM-dd HH:mm:ss}  {1}" -f $Timestamp, $Text) -Encoding utf8 } catch {}
}

# ============================================================================
#  iLO-Redfish-Helper (wird im Runspace per Invoke-Expression geladen)
# ============================================================================
$script:iloCode = @'
# --- Globaler Zertifikats-Bypass (interne iLO-Zertifikate sind selbstsigniert) ---
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { param($s,$c,$ch,$e) $true }
try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13 } catch {
    try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {}
}

$script:IloTimeoutSec = 40

# Login: liefert @{ Token=<X-Auth-Token>; SessionUri=<relative Uri> }
function ILO-Login {
    param([string]$Ilo,[string]$User,[string]$Pass)
    $body = @{ UserName = $User; Password = $Pass } | ConvertTo-Json
    $resp = Invoke-WebRequest -Uri "https://$Ilo/redfish/v1/SessionService/Sessions/" -Method Post `
        -Body $body -ContentType 'application/json' -Headers @{ 'OData-Version' = '4.0' } `
        -SkipCertificateCheck -TimeoutSec $script:IloTimeoutSec -ErrorAction Stop
    $token = $resp.Headers['X-Auth-Token']
    if ($token -is [array]) { $token = $token[0] }
    if ([string]::IsNullOrEmpty($token)) { throw "Kein X-Auth-Token von $Ilo erhalten (Login fehlgeschlagen)" }
    $loc = $resp.Headers['Location']
    if ($loc -is [array]) { $loc = $loc[0] }
    $sessUri = $null
    if ($loc) { try { $sessUri = ([Uri]$loc).AbsolutePath } catch { $sessUri = "$loc" } }
    @{ Token = "$token"; SessionUri = $sessUri }
}

function ILO-Logout {
    param([string]$Ilo,[string]$Token,[string]$SessionUri)
    if ([string]::IsNullOrEmpty($SessionUri)) { return }
    try {
        Invoke-RestMethod -Uri "https://$Ilo$SessionUri" -Method Delete -Headers @{ 'X-Auth-Token' = $Token } `
            -SkipCertificateCheck -TimeoutSec 15 -ErrorAction SilentlyContinue | Out-Null
    } catch {}
}

function ILO-Get {
    param([string]$Ilo,[string]$Token,[string]$Path,[int]$TimeoutSec = 0)
    $to = if ($TimeoutSec -gt 0) { $TimeoutSec } else { $script:IloTimeoutSec }
    Invoke-RestMethod -Uri "https://$Ilo$Path" -Method Get -Headers @{ 'X-Auth-Token' = $Token } `
        -SkipCertificateCheck -TimeoutSec $to -ErrorAction Stop
}

# Ermittelt Modell + Generation. Liefert @{ Model=...; Gen=<int>; Serial=...; iLO=<int> }
function ILO-GetSystemInfo {
    param([string]$Ilo,[string]$Token)
    $sys = ILO-Get -Ilo $Ilo -Token $Token -Path "/redfish/v1/Systems/1"
    $model = "$($sys.Model)"
    $gen = 0
    if ($model -match '(?i)Gen(\d+)') { $gen = [int]$Matches[1] }
    $serial = "$($sys.SerialNumber)"
    # Aktive + Backup(Redundant)-System-ROM-Version direkt aus der
    # ComputerSystem-Ressource (Oem.Hpe.Bios). WICHTIG: Die iLO-Overview liest
    # genau diese Felder - anders als das FirmwareInventory zeigt 'Backup' nach
    # einem Online-Flash SOFORT die frisch in die Redundant-ROM geschriebene,
    # noch nicht per Reboot aktivierte Version. Damit laesst sich ein bereits
    # geflashtes, aber noch nicht aktiviertes BIOS iLO-nativ erkennen.
    $biosCur = ''; $biosBak = ''
    try {
        $bios = $sys.Oem.Hpe.Bios
        if ($bios) {
            if ($bios.Current -and $bios.Current.VersionString) { $biosCur = "$($bios.Current.VersionString)" }
            if ($bios.Backup  -and $bios.Backup.VersionString)  { $biosBak = "$($bios.Backup.VersionString)" }
        }
    } catch {}
    # iLO-Generation aus Manager ableiten
    $iloGen = 0
    try {
        $mgr = ILO-Get -Ilo $Ilo -Token $Token -Path "/redfish/v1/Managers/1"
        $fw = "$($mgr.FirmwareVersion)"   # z.B. "iLO 5 v3.09" / "iLO 6 v1.64"
        if ($fw -match '(?i)iLO\s*(\d+)') { $iloGen = [int]$Matches[1] }
    } catch {}
    $lastResetTime = ''
    try { $lastResetTime = "$($sys.LastResetTime)".Trim() } catch {}
    @{ Model = $model; Gen = $gen; Serial = $serial; iLO = $iloGen; BiosCurrent = $biosCur; BiosBackup = $biosBak; LastResetTime = $lastResetTime }
}

# Wandelt einen Redfish-FirmwareInventory-Eintrag in die kompakte Form um, die
# der Update-Entscheidung zugrunde liegt. HPE nutzt Oem.Hpe.Targets als
# PLDM-Gerätekennung; genau diese Kennung steht auch in den Devices.Device[]
# der passenden .fwpkg. Sie ist damit belastbarer als ein Komponentenname.
function ConvertTo-FirmwareInventoryRecord {
    param($Item)
    $identifiers = @()
    # HPEs dokumentierter Korrelationsschluessel ist ausschliesslich
    # Oem.Hpe.Targets. DeviceInstance ist eine Instanzkennung mit anderer
    # Semantik und darf niemals gegen Paket-Targets querverglichen werden.
    $targets = $null
    try { $targets = $Item.Oem.Hpe.Targets } catch {}
    foreach ($raw in @($targets)) {
        if ($null -eq $raw -or $raw -is [System.Collections.IDictionary] -or $raw -is [System.Management.Automation.PSCustomObject]) { continue }
        $id = "$raw".Trim()
        if ($id -and $identifiers -notcontains $id) { $identifiers += $id }
    }
    $deviceClass = ''
    try { $deviceClass = "$($Item.Oem.Hpe.DeviceClass)".Trim() } catch {}
    # Updateable ist die von iLO gelieferte Freigabe fuer einen Online-Flash.
    # Ein explizites $false darf nicht durch einen Paket-Upload umgangen werden:
    # genau das fuehrt bei z.B. einigen NIC-/Controller-Firmwarepaketen zu einem
    # Fehlerzustand des UpdateService.
    @{ Uri = "$($Item.'@odata.id')"; Name = "$($Item.Name)"; Version = "$($Item.Version)"; Updateable = $Item.Updateable; DeviceClass = $deviceClass; DeviceIdentifiers = @($identifiers) }
}

# Liest das Firmware-Inventar (Name -> Version und HPE-Geraetekenungen).
function ILO-GetFirmwareInventory {
    param([string]$Ilo,[string]$Token)
    $inv = @()
    try {
        $col = ILO-Get -Ilo $Ilo -Token $Token -Path "/redfish/v1/UpdateService/FirmwareInventory/?`$expand=."
        foreach ($m in @($col.Members)) {
            $inv += ConvertTo-FirmwareInventoryRecord -Item $m
        }
    } catch {
        # Fallback ohne $expand: Member-URIs einzeln laden
        try {
            $col = ILO-Get -Ilo $Ilo -Token $Token -Path "/redfish/v1/UpdateService/FirmwareInventory/"
            foreach ($ref in @($col.Members)) {
                $u = $ref."@odata.id"
                if (-not $u) { continue }
                try { $m = ILO-Get -Ilo $Ilo -Token $Token -Path $u; $inv += ConvertTo-FirmwareInventoryRecord -Item $m } catch {}
            }
        } catch {}
    }
    ,$inv
}

# Liefert das vollstaendige Firmware-Inventar fuer den Diagnoseexport. Anders
# als ILO-GetFirmwareInventory bleiben hier auch HPE-OEM-Felder erhalten,
# insbesondere Targets und DeviceClass. Sie sind fuer eine eindeutige Paketzuordnung
# wesentlich aussagekraeftiger als der sichtbare Komponentenname allein.
function ILO-GetFirmwareInventoryDetails {
    param([string]$Ilo,[string]$Token)
    $members = @()
    try {
        $col = ILO-Get -Ilo $Ilo -Token $Token -Path "/redfish/v1/UpdateService/FirmwareInventory/?`$expand=."
        $members = @($col.Members)
        # Manche iLO-Versionen ignorieren $expand und liefern nur Referenzen.
        if ($members.Count -and -not (@($members[0].PSObject.Properties.Name) -contains 'Name')) { $members = @() }
    } catch { $members = @() }
    if (-not $members.Count) {
        try {
            $col = ILO-Get -Ilo $Ilo -Token $Token -Path "/redfish/v1/UpdateService/FirmwareInventory/"
            foreach ($ref in @($col.Members)) {
                $uri = "$($ref."@odata.id")"
                if (-not $uri) { continue }
                try { $members += ILO-Get -Ilo $Ilo -Token $Token -Path $uri } catch {}
            }
        } catch {}
    }
    $details = @()
    foreach ($m in @($members)) {
        if (-not $m) { continue }
        $hpeOem = $null
        try { $hpeOem = $m.Oem.Hpe } catch {}
        $details += [pscustomobject]@{
            Uri         = "$($m."@odata.id")"
            Id          = "$($m.Id)"
            Name        = "$($m.Name)"
            Version     = "$($m.Version)"
            Description = "$($m.Description)"
            Updateable  = $m.Updateable
            RelatedItem = @($m.RelatedItem)
            DeviceClass = "$($hpeOem.DeviceClass)"
            OemHpe      = $hpeOem
        }
    }
    return @($details)
}

# iLO stellt je nach Generation und Lizenz unterschiedliche Device-Inventory-
# Endpunkte bereit. Der Diagnoseexport prueft sie nur lesend und notiert auch
# nicht verfuegbare Pfade, damit die spaetere Zuordnung reproduzierbar bleibt.
function ILO-GetDeviceInventoryDiagnostic {
    param([string]$Ilo,[string]$Token)
    $result = @()
    foreach ($path in @(
        '/redfish/v1/Systems/1/DeviceInventory/?$expand=.',
        '/redfish/v1/Systems/1/Oem/Hpe/DeviceInventory/?$expand=.',
        '/redfish/v1/Chassis/1/DeviceInventory/?$expand=.',
        '/redfish/v1/Systems/1/PCIeDevices/?$expand=.'
    )) {
        try {
            $data = ILO-Get -Ilo $Ilo -Token $Token -Path $path
            $result += [pscustomobject]@{ Path = $path; Available = $true; Data = $data }
        } catch {
            $result += [pscustomobject]@{ Path = $path; Available = $false; Error = $_.Exception.Message }
        }
    }
    return @($result)
}

# Liefert aus einer Redfish-Referenz oder einem Referenzarray ausschliesslich
# relative Redfish-URIs. Die Hilfsfunktion ist absichtlich schema-neutral,
# weil die Storage-Struktur zwischen iLO 5 (HPE SmartStorage) und iLO 6/7
# (DMTF Storage/PLDM-RDE) unterschiedlich aufgebaut ist.
function Get-RedfishLinkUris {
    param($Value)
    $uris = New-Object System.Collections.ArrayList
    foreach ($entry in @($Value)) {
        if ($null -eq $entry) { continue }
        $uri = ''
        if ($entry -is [string]) { $uri = "$entry" }
        else { try { $uri = "$($entry."@odata.id")" } catch {} }
        if ($uri -match '^https?://') {
            try { $uri = ([Uri]$uri).AbsolutePath } catch {}
        }
        if ($uri -and $uri.StartsWith('/') -and -not ($uris -contains $uri)) { [void]$uris.Add($uri) }
    }
    return @($uris)
}

# Formt einen physischen Datentraeger in einen generationsneutralen,
# auswertbaren Diagnoseeintrag um. Die vollstaendige Redfish-Antwort bleibt
# ebenfalls erhalten, damit modellspezifische OEM-Felder nicht verloren gehen.
function ConvertTo-StorageDriveDiagnostic {
    param($Drive,[string]$Uri,[string]$Source)
    $hpeOem = $null
    try { $hpeOem = $Drive.Oem.Hpe } catch {}
    $firmware = ''
    foreach ($value in @($Drive.FirmwareVersion, $Drive.Version, $hpeOem.FirmwareVersion, $hpeOem.Firmware.Current.VersionString)) {
        if ($value) { $firmware = "$value"; break }
    }
    [pscustomobject]@{
        Uri             = $Uri
        Source          = $Source
        Id              = "$($Drive.Id)"
        Name            = "$($Drive.Name)"
        Model           = "$($Drive.Model)"
        Manufacturer    = "$($Drive.Manufacturer)"
        SerialNumber    = "$($Drive.SerialNumber)"
        PartNumber      = "$($Drive.PartNumber)"
        FirmwareVersion = $firmware
        Protocol        = "$($Drive.Protocol)"
        MediaType       = "$($Drive.MediaType)"
        CapacityBytes   = $Drive.CapacityBytes
        Status          = $Drive.Status
        OemHpe          = $hpeOem
        Data            = $Drive
    }
}

# Liest das Storage-/Laufwerksinventar nur zur Diagnose. Der Ablauf ist
# generationsbewusst: iLO 6/7 und RDE-Controller verwenden die DMTF-Links
# unter ComputerSystem.Storage; iLO 5 kann zusaetzlich das alte HPE
# SmartStorage-Schema bereitstellen. Anstelle fester IDs wird immer zuerst den
# vom jeweiligen iLO gelieferten @odata.id-Links gefolgt.
function ILO-GetStorageInventoryDiagnostic {
    param([string]$Ilo,[string]$Token,[int]$IloGeneration = 0)
    $probes = New-Object System.Collections.ArrayList
    $controllers = New-Object System.Collections.ArrayList
    $drives = New-Object System.Collections.ArrayList
    $seenDrives = @{}
    $read = {
        param([string]$Path,[string]$Purpose)
        if (-not $Path) { return $null }
        try {
            $data = ILO-Get -Ilo $Ilo -Token $Token -Path $Path
            [void]$probes.Add([pscustomobject]@{ Path = $Path; Purpose = $Purpose; Available = $true })
            return $data
        } catch {
            [void]$probes.Add([pscustomobject]@{ Path = $Path; Purpose = $Purpose; Available = $false; Error = $_.Exception.Message })
            return $null
        }
    }.GetNewClosure()
    $addController = {
        param($Controller,[string]$Uri,[string]$Source)
        if (-not $Controller) { return }
        $key = if ($Uri) { $Uri } else { "$($Controller.Id)|$Source" }
        if (@($controllers | Where-Object { $_.Uri -eq $key }).Count) { return }
        [void]$controllers.Add([pscustomobject]@{
            Uri             = $Uri
            Source          = $Source
            Id              = "$($Controller.Id)"
            Name            = "$($Controller.Name)"
            Model           = "$($Controller.Model)"
            Manufacturer    = "$($Controller.Manufacturer)"
            FirmwareVersion = "$($Controller.FirmwareVersion)"
            Status          = $Controller.Status
            Data            = $Controller
        })
    }.GetNewClosure()
    $addDrive = {
        param($Drive,[string]$Uri,[string]$Source)
        if (-not $Drive) { return }
        $key = if ($Uri) { $Uri } else { "$($Drive.Id)|$($Drive.SerialNumber)|$Source" }
        if ($seenDrives.ContainsKey($key)) { return }
        $seenDrives[$key] = $true
        [void]$drives.Add((ConvertTo-StorageDriveDiagnostic -Drive $Drive -Uri $Uri -Source $Source))
    }.GetNewClosure()

    # DMTF Storage: ab Gen10 vorhanden und fuer iLO 6/7 der einzige
    # zukunftssichere Weg. Das Link-Feld wird bevorzugt; der feste Pfad ist nur
    # ein Fallback fuer aeltere iLO-5-Firmware ohne Link im Systemobjekt.
    $system = & $read '/redfish/v1/Systems/1' 'ComputerSystem (Storage-Discovery)'
    $storageRoots = @()
    try { $storageRoots += Get-RedfishLinkUris $system.Storage } catch {}
    if (-not $storageRoots.Count) { $storageRoots = @('/redfish/v1/Systems/1/Storage') }
    foreach ($rootUri in @($storageRoots | Select-Object -Unique)) {
        $root = & $read $rootUri 'DMTF Storage collection'
        if (-not $root) { continue }
        $storageUris = @(Get-RedfishLinkUris $root.Members)
        # Manche Implementierungen geben statt einer Collection direkt eine
        # einzelne Storage-Ressource zurueck.
        if (-not $storageUris.Count -and ($root.PSObject.Properties.Name -contains 'Drives')) { $storageUris = @($rootUri) }
        foreach ($storageUri in $storageUris) {
            $storage = if ($storageUri -eq $rootUri) { $root } else { & $read $storageUri 'DMTF Storage subsystem' }
            if (-not $storage) { continue }
            & $addController $storage $storageUri 'DMTF-Storage'
            foreach ($driveUri in @(Get-RedfishLinkUris $storage.Drives)) {
                $drive = & $read $driveUri 'DMTF physical drive'
                & $addDrive $drive $driveUri 'DMTF-Storage'
            }
        }
    }

    # iLO 5 / Gen10 kann das alte OEM-Schema bereitstellen. HPE hat es ab iLO
    # 6/Gen11 entfernt; daher wird es dort nur probiert, wenn die Standard-API
    # keinerlei Controller geliefert hat. Fehlende Legacy-Endpunkte sind im
    # Export sichtbar, aber niemals ein Fehler fuer das Inventar selbst.
    $useLegacy = ($IloGeneration -le 5 -or $controllers.Count -eq 0)
    if ($useLegacy) {
        $legacyRoots = @()
        try { $legacyRoots += Get-RedfishLinkUris $system.Oem.Hpe.SmartStorage } catch {}
        if (-not $legacyRoots.Count) { $legacyRoots = @('/redfish/v1/Systems/1/SmartStorage') }
        foreach ($legacyUri in @($legacyRoots | Select-Object -Unique)) {
            $legacy = & $read $legacyUri 'HPE SmartStorage root (iLO 5 legacy)'
            if (-not $legacy) { continue }
            $controllerCollections = @()
            $controllerCollections += Get-RedfishLinkUris $legacy.ArrayControllers
            $controllerCollections += Get-RedfishLinkUris $legacy.HostBusAdapters
            if (-not $controllerCollections.Count) {
                $controllerCollections = @(
                    "$legacyUri/ArrayControllers",
                    "$legacyUri/HostBusAdapters"
                )
            }
            foreach ($collectionUri in @($controllerCollections | Select-Object -Unique)) {
                $collection = & $read $collectionUri 'HPE SmartStorage controller collection'
                if (-not $collection) { continue }
                foreach ($controllerUri in @(Get-RedfishLinkUris $collection.Members)) {
                    $controller = & $read $controllerUri 'HPE SmartStorage controller'
                    if (-not $controller) { continue }
                    & $addController $controller $controllerUri 'HPE-SmartStorage'
                    $driveCollections = @()
                    # iLO 5 liefert bei HostBusAdapters die Laufwerks-Collection
                    # nicht direkt als DiskDrives, sondern unter Links.Drives.
                    # Dieser Link ist server-/controllerabhaengig und muss daher
                    # dem vom Controller bereitgestellten @odata.id folgen.
                    $driveCollections += Get-RedfishLinkUris $controller.Links.Drives
                    $driveCollections += Get-RedfishLinkUris $controller.Links.DiskDrives
                    $driveCollections += Get-RedfishLinkUris $controller.DiskDrives
                    $driveCollections += Get-RedfishLinkUris $controller.UnconfiguredDrives
                    foreach ($driveCollectionUri in @($driveCollections | Select-Object -Unique)) {
                        $driveCollection = & $read $driveCollectionUri 'HPE SmartStorage drive collection'
                        if (-not $driveCollection) { continue }
                        foreach ($driveUri in @(Get-RedfishLinkUris $driveCollection.Members)) {
                            $drive = & $read $driveUri 'HPE SmartStorage physical drive'
                            & $addDrive $drive $driveUri 'HPE-SmartStorage'
                        }
                    }
                }
            }
        }
    }
    [pscustomobject]@{
        Strategy    = if ($IloGeneration -ge 6) { 'DMTF Storage (iLO 6/7), Legacy nur bei Bedarf' } else { 'DMTF Storage + HPE SmartStorage (iLO 5)' }
        Probes      = @($probes)
        Controllers = @($controllers)
        Drives      = @($drives)
    }
}

# Liefert @{ State=<HpeState>; Percent=<int>; PushUri=<Upload-URI>; Detail=<iLO-Fehlerdetails> }
function ILO-GetUpdateState {
    param([string]$Ilo,[string]$Token)
    $us = ILO-Get -Ilo $Ilo -Token $Token -Path "/redfish/v1/UpdateService"
    $state = $null; $pct = $null
    if ($us.Oem -and $us.Oem.Hpe) {
        $state = "$($us.Oem.Hpe.State)"
        if ($null -ne $us.Oem.Hpe.FlashProgressPercent) { $pct = [int]$us.Oem.Hpe.FlashProgressPercent }
    }
    # Upload-URI ermitteln. Unser Multipart-Payload (sessionKey/parameters/file)
    # entspricht EXAKT dem offiziellen HPE-Beispiel, das 'HttpPushUri'
    # (= /cgi-bin/uploadFile) verwendet - identisch fuer iLO 5/6/7. Daher
    # 'HttpPushUri' ZUERST. Der DMTF-Endpunkt 'MultipartHttpPushUri'
    # (/redfish/v1/UpdateService/upload) erwartet andere Part-Namen
    # (UpdateParameters/UpdateFile) und passt NICHT zu diesem Payload -> nur als
    # Notnagel, wenn HttpPushUri fehlt.
    $pushUri = ''
    if ($us.PSObject.Properties.Name -contains 'HttpPushUri') { $pushUri = "$($us.HttpPushUri)".Trim() }
    if (-not $pushUri -and $us.PSObject.Properties.Name -contains 'MultipartHttpPushUri') { $pushUri = "$($us.MultipartHttpPushUri)".Trim() }
    if (-not $pushUri -and $us.PSObject.Properties.Name -contains 'MultiPartHttpPushUri') { $pushUri = "$($us.MultiPartHttpPushUri)".Trim() }
    # Fallback: HPE-Upload-Endpunkt von iLO 5/6/7.
    if (-not $pushUri) { $pushUri = '/cgi-bin/uploadFile' }
    # Bei State=Error stellt iLO die Ursache je nach Generation unter
    # unterschiedlichen OEM-Feldern bereit. Fuer das Log wird sie kompakt
    # zusammengefasst, ohne eine generationsspezifische Feldannahme zu treffen.
    $detailParts = @()
    if ($us.Oem -and $us.Oem.Hpe) {
        foreach ($field in @('Result','Message','ErrorCode','Status','ExtendedInfo')) {
            try {
                $value = $us.Oem.Hpe.$field
                if ($null -ne $value -and "$value".Trim()) { $detailParts += ("{0}={1}" -f $field, "$value".Trim()) }
            } catch {}
        }
    }
    foreach ($field in @('Status','Message')) {
        try {
            $value = $us.$field
            if ($null -ne $value -and "$value".Trim()) { $detailParts += ("{0}={1}" -f $field, "$value".Trim()) }
        } catch {}
    }
    @{ State = $state; Percent = $pct; PushUri = $pushUri; Detail = (@($detailParts | Select-Object -Unique) -join '; ') }
}

# Laedt eine .fwpkg via MultipartHttpPushUri hoch und startet den Flash.
function ILO-UploadComponent {
    param(
        [string]$Ilo,
        [string]$Token,
        [string]$PushUri,
        [string]$FilePath,
        [bool]$UpdateRepository = $false,
        [bool]$UpdateTarget = $true,
        [scriptblock]$ProgressCb = $null
    )
    if ([string]::IsNullOrWhiteSpace($PushUri)) { throw "MultipartHttpPushUri nicht verfuegbar (zu altes iLO?)" }
    if (-not (Test-Path -LiteralPath $FilePath)) { throw "Datei nicht gefunden: $FilePath" }
    $fi = Get-Item -LiteralPath $FilePath
    $fileName = $fi.Name
    $totalSize = $fi.Length
    $uri = "https://$Ilo$PushUri"

    $handler = New-Object System.Net.Http.HttpClientHandler
    try { $handler.ServerCertificateCustomValidationCallback = [System.Net.Http.HttpClientHandler]::DangerousAcceptAnyServerCertificateValidator }
    catch { $handler.ServerCertificateCustomValidationCallback = { param($a,$b,$c,$d) $true } }
    try { $handler.CheckCertificateRevocationList = $false } catch {}

    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromHours(2)
    $client.DefaultRequestHeaders.Add("X-Auth-Token", $Token)
    $client.DefaultRequestHeaders.Add("OData-Version", "4.0")
    $client.DefaultRequestHeaders.Add("Accept", "application/json")
    # WICHTIG: 'Expect: 100-continue' abschalten. .NET sendet das per Default;
    # iLO lehnt es mit 400 Bad Request ab (HPE-Beispiel setzt 'Expect:' leer).
    $client.DefaultRequestHeaders.ExpectContinue = $false
    # Cookie-basierte Session zusaetzlich mitsenden (von HPE-Doku empfohlen).
    try { $client.DefaultRequestHeaders.Add("Cookie", "sessionKey=$Token") } catch {}

    $fs = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        # Eigene Boundary ohne Anführungszeichen. .NET setzt sonst
        # boundary="..." (gequotet); der alte CGI-Parser von iLO (/cgi-bin/uploadFile)
        # lehnt die gequotete Boundary mit 400 Bad Request ab. curl sendet sie ungequotet.
        $boundary = "---------------------------" + [DateTime]::Now.Ticks.ToString("x")
        $multi = New-Object System.Net.Http.MultipartFormDataContent($boundary)
        # Quotes aus dem Boundary-Parameter des Content-Type-Headers entfernen.
        foreach ($prm in $multi.Headers.ContentType.Parameters) {
            if ($prm.Name -eq 'boundary') { $prm.Value = $prm.Value.Trim('"') }
        }

        # 1) sessionKey
        $skc = New-Object System.Net.Http.StringContent($Token)
        $cdSk = New-Object System.Net.Http.Headers.ContentDispositionHeaderValue("form-data"); $cdSk.Name = "sessionKey"
        $skc.Headers.ContentDisposition = $cdSk
        $multi.Add($skc)

        # 2) parameters (JSON)
        #   UpdateRepository=$false / UpdateTarget=$true  -> direkt online flashen
        #     (iLO, System-ROM, ...).
        #   UpdateRepository=$true  / UpdateTarget=$false -> nur ins iLO-Repository
        #     legen (deferred, z.B. SPS/ME/CPLD). Aktivierung erfolgt anschliessend
        #     ueber einen Task in der UpdateTaskQueue beim naechsten Reboot/POST.
        #     UpdateTarget=$true wuerde hier den Zustand 'Error' ausloesen!
        # Der HPE-HttpPushUri akzeptiert hier nur seine dokumentierten
        # Uploadparameter. TPMOverride gehoert nicht in diesen Multipart-POST:
        # iLO kann unbekannte Felder annehmen, die Komponente danach aber mit
        # einer irrefuehrenden Zielgeraete-Meldung ablehnen. Falls ein Queue-Task
        # TPMOverride verlangt, behandelt ILO-CreateUpdateTask das separat.
        $paramObj = @{ UpdateRepository = $UpdateRepository; UpdateTarget = $UpdateTarget; ETag = "atoken"; Section = 0 } | ConvertTo-Json -Compress
        $pc = New-Object System.Net.Http.StringContent($paramObj)
        $pc.Headers.ContentType = New-Object System.Net.Http.Headers.MediaTypeHeaderValue("application/json")
        $cdP = New-Object System.Net.Http.Headers.ContentDispositionHeaderValue("form-data"); $cdP.Name = "parameters"
        $pc.Headers.ContentDisposition = $cdP
        $multi.Add($pc)

        # 3) file (binaer). Name/FileName werden von .NET selbst korrekt gequotet.
        $streamContent = New-Object System.Net.Http.StreamContent($fs, 1MB)
        $streamContent.Headers.ContentType = New-Object System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream")
        $cdF = New-Object System.Net.Http.Headers.ContentDispositionHeaderValue("form-data")
        $cdF.Name = "file"; $cdF.FileName = $fileName
        $streamContent.Headers.ContentDisposition = $cdF
        $multi.Add($streamContent)

        $task = $client.PostAsync($uri, $multi)
        if ($ProgressCb) {
            while (-not $task.IsCompleted) {
                Start-Sleep -Milliseconds 750
                try {
                    $sent = $fs.Position
                    if ($totalSize -gt 0) { & $ProgressCb ([int](($sent / $totalSize) * 100)) }
                } catch {}
            }
        }
        $resp = $task.GetAwaiter().GetResult()
        $body = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        if (-not $resp.IsSuccessStatusCode) {
            throw "Upload/Flash abgelehnt ($([int]$resp.StatusCode) $($resp.ReasonPhrase)): $body"
        }
        if ($ProgressCb) { & $ProgressCb 100 }
        return $body
    }
    finally {
        $fs.Dispose(); $client.Dispose(); $handler.Dispose()
    }
}

# Wartet bis der Flash abgeschlossen ist. Robust gegen iLO-Reset (re-Login).
# Liefert @{ Token=<evtl. neu>; State=<Endzustand> }
function ILO-WaitForFlash {
    param(
        [string]$Ilo,[string]$Token,[string]$User,[string]$Pass,
        [int]$TimeoutSec = 1800,
        [scriptblock]$ProgressCb = $null
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $lastPct = -1
    $reachedActive = $false
    while ((Get-Date) -lt $deadline) {
        try {
            $st = ILO-GetUpdateState -Ilo $Ilo -Token $Token
            $state = $st.State
            if ($null -ne $st.Percent -and $st.Percent -ne $lastPct) {
                $lastPct = $st.Percent
                if ($ProgressCb) { & $ProgressCb $st.Percent $state }
            }
            switch -Regex ($state) {
                '^(Uploading|Verifying|Writing|Updating)$' { $reachedActive = $true; Start-Sleep -Seconds 5 }
                '^(Complete|Idle)$' {
                    if ($reachedActive) { return @{ Token = $Token; State = $state } }
                    Start-Sleep -Seconds 4   # Flash evtl. noch nicht gestartet
                }
                '^Error$' { throw "iLO meldet Flash-State 'Error'" }
                default { Start-Sleep -Seconds 5 }
            }
        }
        catch {
            if ($_.Exception.Message -like '*Flash-State*') { throw }
            # Verbindungsverlust -> iLO resettet evtl. (iLO-Firmware). Re-Login versuchen.
            Start-Sleep -Seconds 15
            try {
                $s = ILO-Login -Ilo $Ilo -User $User -Pass $Pass
                $Token = $s.Token
                $reachedActive = $true
            } catch {
                # iLO noch nicht zurueck - weiter warten
            }
        }
    }
    throw "Timeout nach $TimeoutSec s - Flash-Status nicht bestaetigt"
}

# Wartet bis der Repository-Upload (deferred) verarbeitet ist.
# Anders als ILO-WaitForFlash wird hier KEIN aktiver Flash erwartet - die States
# laufen Uploading/Verifying/Writing -> Complete/Idle (KEIN 'Updating').
# Liefert den End-State; wirft bei 'Error'.
function ILO-WaitForRepository {
    param(
        [string]$Ilo,[string]$Token,
        [int]$TimeoutSec = 600,
        [scriptblock]$ProgressCb = $null
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    # Ein noch anstehender 'Error'-Rest eines VORHERIGEN Vorgangs soll unseren
    # neuen Upload nicht sofort abbrechen: bis zu dieser Frist auf den Start
    # (Uploading/Verifying/Writing) warten, bevor 'Error' als echter Fehler zaehlt.
    $graceDeadline = (Get-Date).AddSeconds(30)
    $lastPct = -1
    $sawActive = $false
    while ((Get-Date) -lt $deadline) {
        $st = ILO-GetUpdateState -Ilo $Ilo -Token $Token
        $state = $st.State
        if ($null -ne $st.Percent -and $st.Percent -ne $lastPct) {
            $lastPct = $st.Percent
            if ($ProgressCb) { & $ProgressCb $st.Percent $state }
        }
        switch -Regex ($state) {
            '^(Uploading|Verifying|Writing)$' { $sawActive = $true; Start-Sleep -Seconds 3 }
            '^Error$' {
                if ($sawActive -or (Get-Date) -gt $graceDeadline) { throw "iLO meldet Repository-State 'Error'" }
                Start-Sleep -Seconds 3
            }
            '^(Complete|Idle|)$' {
                # Kurz gegenpruefen, damit ein noch nicht gestarteter Schreibvorgang
                # nicht faelschlich als 'fertig' gilt.
                if ($sawActive) { return $state }
                Start-Sleep -Seconds 2
                $st2 = ILO-GetUpdateState -Ilo $Ilo -Token $Token
                if ($st2.State -match '^(Complete|Idle|)$') { return $st2.State }
            }
            default { Start-Sleep -Seconds 3 }
        }
    }
    throw "Timeout nach $TimeoutSec s - Repository-Upload nicht bestaetigt"
}

# Wartet auf einen stabil freien UpdateService. Ein Queue-Task kann bereits
# "Complete" melden, waehrend iLO global noch Verifying/Updating ausfuehrt.
# Vor dem naechsten Upload/Task muessen daher BEIDE Abfragen abgeschlossen sein.
function ILO-WaitForUpdateServiceReady {
    param(
        [string]$Ilo,[string]$Token,
        [int]$TimeoutSec = 2400,
        [scriptblock]$LogCb = $null
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $lastState = $null
    $stableReady = 0
    while ((Get-Date) -lt $deadline) {
        $st = ILO-GetUpdateState -Ilo $Ilo -Token $Token
        $state = "$($st.State)"
        if ($LogCb -and $state -ne $lastState) {
            try { & $LogCb "UpdateService-Wartezustand: State='$state'." } catch {}
            $lastState = $state
        }
        if ($state -match '^(Error|Exception|Failed)$') {
            $detail = if ($st.Detail) { ": $($st.Detail)" } else { '' }
            throw "iLO-UpdateService meldet State '$state'$detail"
        }
        if ($state -match '^(Idle|Complete|)$') {
            $stableReady++
            if ($stableReady -ge 2) { return $st }
        } else {
            $stableReady = 0
        }
        Start-Sleep -Seconds 3
    }
    throw "Timeout nach $TimeoutSec s - iLO-UpdateService wurde nicht frei"
}

# BMC-Tasks aus der UpdateTaskQueue werden bei eingeschaltetem Server sofort
# verarbeitet. Die Queue-Bestaetigung allein ist aber noch kein Nachweis des
# erfolgreichen Flashs; deshalb wird die konkrete Zielversion im
# FirmwareInventory abgewartet. Die Zuordnung verwendet dieselben Targets und
# DeviceClasses wie Auswahl und Vorschau.
#
# WICHTIG: Zusaetzlich wird der QUEUE-TASK selbst ueberwacht. Manche Firmware
# (v.a. NIC/PLDM wie Mellanox MCX) wird vom Bmc zwar erfolgreich geschrieben
# (Task='Complete'), erscheint im FirmwareInventory aber erst NACH dem Reboot
# mit neuer Version. Ohne Task-Pruefung wuerde die Schleife hier bis zum
# Timeout haengen und alle Folgepakete blockieren. Rueckgabe daher:
#   @{ Status='Confirmed';     Inventory=<Eintrag>; TaskState=... }  -> Version aktiv
#   @{ Status='PendingReboot'; Inventory=$null;     TaskState=... }  -> geflasht/queued, Aktivierung beim Reboot
# Task-Fehler (Exception/Canceled/Expired) fuehren zu einem Throw mit Detail.
function ILO-WaitForFirmwareInventoryVersion {
    param(
        [string]$Ilo,[string]$Token,[string]$TargetVersion,
        [string]$InvPattern = '',[string[]]$InventoryTerms = @(),
        [string[]]$DeviceIdentifiers = @(),[string[]]$DeviceClasses = @(),
        [string]$TaskUri = '',[string]$ComponentFileName = '',
        [int]$TimeoutSec = 2400,[scriptblock]$ProgressCb = $null,
        [scriptblock]$LogCb = $null
    )
    $targetNorm = Get-NormFwVersion $TargetVersion
    if (-not $targetNorm) { throw 'BMC-Queue kann nicht bestaetigt werden: Zielversion aus Firmwarepaket nicht lesbar.' }
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    # Nachfristen: Task fertig/wartend, aber Inventar (noch) unveraendert.
    $completeGraceSec = 90    # Task 'Complete' -> kurz nachpruefen, dann PendingReboot
    $pendingGraceSec  = 300   # Task bleibt 'Pending' -> iLO wartet auf Reboot
    $completeSince = $null; $pendingSince = $null
    $lastReportedTaskState = $null
    while ((Get-Date) -lt $deadline) {
        $state = ILO-GetUpdateState -Ilo $Ilo -Token $Token
        if ($state.State -match '^(Error|Exception|Failed)$') {
            $detail = if ($state.Detail) { ": $($state.Detail)" } else { '' }
            throw "iLO meldet waehrend BMC-Queue den UpdateService-State '$($state.State)'$detail"
        }
        $inv = ILO-GetFirmwareInventory -Ilo $Ilo -Token $Token
        $matches = @(ILO-FindInstalledFirmware -Inventory $inv -InvPattern $InvPattern -InventoryTerms $InventoryTerms -DeviceIdentifiers $DeviceIdentifiers -DeviceClasses $DeviceClasses)
        $hit = @($matches | Where-Object { (Compare-FwVersion $_.Version $TargetVersion) -eq 0 }) | Select-Object -First 1
        if ($hit) { return [pscustomobject]@{ Status = 'Confirmed'; Inventory = $hit; TaskState = '' } }

        # Task-Status ermitteln: bevorzugt direkt ueber die Task-URI, sonst
        # ueber die Queue anhand des Repository-Dateinamens.
        $taskState = ''; $taskDetail = ''
        if ($TaskUri) {
            try {
                $t = ILO-Get -Ilo $Ilo -Token $Token -Path $TaskUri
                $taskState = "$($t.State)"
                if ($t.PSObject.Properties.Name -contains 'Result' -and $t.Result) { $taskDetail = ("$($t.Result | ConvertTo-Json -Compress -Depth 4)") }
            } catch {}
        }
        # Manche iLO-Versionen liefern zwar eine Task-URI, erlauben danach aber
        # keinen Einzelabruf. Dann den Task ueber den Dateinamen in der Queue suchen.
        if (-not $taskState -and $ComponentFileName) {
            try {
                $qt = @(ILO-GetUpdateTaskQueue -Ilo $Ilo -Token $Token) | Where-Object { "$($_.Filename)" -ieq $ComponentFileName } | Select-Object -First 1
                if ($qt) {
                    $taskState = "$($qt.State)"
                    if ($qt.Result) { $taskDetail = ("$($qt.Result | ConvertTo-Json -Compress -Depth 4)") }
                }
            } catch {}
        }

        if ($LogCb -and $taskState -ne $lastReportedTaskState) {
            try { & $LogCb "Bmc-Queue-Status '$ComponentFileName': Task='$taskState', UpdateService='$($state.State)'." } catch {}
            $lastReportedTaskState = $taskState
        }

        if ($taskState -match '^(Exception|Canceled|Cancelled|Expired|Error|Failed)$') {
            if ($taskDetail -match 'NoMatchingTargetDeviceFound') {
                return [pscustomobject]@{ Status = 'NotApplicable'; Inventory = $null; TaskState = $taskState; TaskDetail = $taskDetail }
            }
            $d = if ($taskDetail) { " | Task-Result: $taskDetail" } else { '' }
            throw "BMC-Queue-Task meldet State '$taskState'$d"
        }
        if ($taskState -eq 'Complete' -and $state.State -match '^(Idle|Complete|)$') {
            if (-not $completeSince) { $completeSince = Get-Date }
            if (((Get-Date) - $completeSince).TotalSeconds -ge $completeGraceSec) {
                return [pscustomobject]@{ Status = 'PendingReboot'; Inventory = $null; TaskState = $taskState }
            }
        } else { $completeSince = $null }
        if ($taskState -eq 'Pending' -and $state.State -notmatch '^(Uploading|Verifying|Writing|Updating|InProgress)$') {
            if (-not $pendingSince) { $pendingSince = Get-Date }
            if (((Get-Date) - $pendingSince).TotalSeconds -ge $pendingGraceSec) {
                return [pscustomobject]@{ Status = 'PendingReboot'; Inventory = $null; TaskState = $taskState }
            }
        } else { $pendingSince = $null }

        if ($ProgressCb) { & $ProgressCb $state.Percent $state.State }
        Start-Sleep -Seconds 5
    }
    throw "Timeout nach $TimeoutSec s - BMC-Queue-Zielversion '$TargetVersion' nicht im FirmwareInventory bestaetigt"
}

# Liest das serverseitige iLO ComponentRepository. Im Strict-Modus muss die
# Sammlung einschliesslich aller Member vollstaendig lesbar sein; andernfalls
# darf keine Updateentscheidung getroffen werden.
function ILO-GetComponentRepository {
    param([string]$Ilo,[string]$Token,[switch]$Strict)
    $errors = @()
    foreach ($repoPath in @('/redfish/v1/UpdateService/ComponentRepository/?$expand=.', '/redfish/v1/UpdateService/ComponentRepository/')) {
        try {
            $repo = ILO-Get -Ilo $Ilo -Token $Token -Path $repoPath
            $members = @($repo.Members)
            $full = @()
            foreach ($member in $members) {
                if ($member.PSObject.Properties.Name -contains 'Filename') {
                    $full += $member
                    continue
                }
                $u = "$($member.'@odata.id')"
                if (-not $u) {
                    if ($Strict) { throw 'Repository-Member ohne @odata.id' }
                    continue
                }
                try { $full += ILO-Get -Ilo $Ilo -Token $Token -Path $u }
                catch {
                    if ($Strict) { throw "Repository-Member '$u' konnte nicht gelesen werden: $($_.Exception.Message)" }
                }
            }
            return @($full)
        } catch {
            $errors += $_.Exception.Message
        }
    }
    if ($Strict) { throw "ComponentRepository konnte nicht vollstaendig gelesen werden: $($errors -join ' | ')" }
    return @()
}

# Sucht im iLO ComponentRepository nach dem exakten Dateinamen. Der Filename ist
# laut HPE der eindeutige Repository-Schluessel; Teilstring-Matches sind fuer
# Updateentscheidungen nicht sicher genug.
function ILO-FindRepositoryComponent {
    param([string]$Ilo,[string]$Token,[string]$ComponentFileName,[switch]$Strict)
    $members = @(ILO-GetComponentRepository -Ilo $Ilo -Token $Token -Strict:$Strict)
    return $members | Where-Object { "$($_.Filename)" -ieq $ComponentFileName } | Select-Object -First 1
}

# Prueft NACH dem reinen Repository-Upload, aber VOR ApplyUpdate, die von iLO
# selbst ermittelte Paketzuordnung. Nur ein exakter Schnitt zwischen den
# Repository-Targets und Oem.Hpe.Targets des Live-Inventars darf
# einen Queue-Task freigeben. Ein Paketname oder eine lokale Metadaten-Heuristik
# allein reicht dafuer nicht aus.
function Test-RepositoryComponentApplicable {
    param($Component,[object[]]$InventoryMatches = @(),[string]$TargetVersion = '')
    if (-not $Component) {
        return [pscustomobject]@{ Applicable = $false; Reason = 'Repository-Komponente fehlt nach dem Upload' }
    }
    $repoTargets = @($Component.Targets | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Select-Object -Unique)
    if (-not $repoTargets.Count) {
        return [pscustomobject]@{ Applicable = $false; Reason = 'iLO liefert fuer die Repository-Komponente keine Targets' }
    }
    $repoClass = "$($Component.DeviceClass)".Trim()
    $targetNorm = Get-NormFwVersion $TargetVersion
    $repoNorm = Get-NormFwVersion "$($Component.Version)"
    if ($targetNorm -and $repoNorm -and (Compare-FwVersion $repoNorm $targetNorm) -ne 0) {
        return [pscustomobject]@{ Applicable = $false; Reason = "Repository-Version '$($Component.Version)' stimmt nicht mit Ziel '$TargetVersion' ueberein" }
    }
    foreach ($item in @($InventoryMatches)) {
        $inventoryClass = "$($item.DeviceClass)".Trim()
        if ($repoClass -and $inventoryClass -and $repoClass -ine $inventoryClass) { continue }
        $inventoryIds = @($item.DeviceIdentifiers | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
        $matchingTargets = @($repoTargets | Where-Object { $inventoryIds -contains $_ })
        if ($matchingTargets.Count) {
            return [pscustomobject]@{
                Applicable = $true
                Reason = "iLO-Repository-Target '$($matchingTargets[0])' stimmt mit Live-Inventar '$($item.Name)' ueberein"
            }
        }
    }
    return [pscustomobject]@{
        Applicable = $false
        Reason = "keines der iLO-Repository-Targets [$($repoTargets -join ', ')] passt zu den Live-Inventar-Targets"
    }
}

# Liest die iLO UpdateTaskQueue aus und liefert je Task ein Objekt
# @{ Name; Filename; Id; Uri; State }. Wird zur Verifikation genutzt, dass ein
# angelegter Task tatsaechlich in der Queue liegt.
function ILO-GetUpdateTaskQueue {
    param([string]$Ilo,[string]$Token,[string]$QueueUri = $null,[switch]$Strict)
    if ([string]::IsNullOrWhiteSpace($QueueUri)) {
        $QueueUri = '/redfish/v1/UpdateService/UpdateTaskQueue'
        try {
            $us = ILO-Get -Ilo $Ilo -Token $Token -Path '/redfish/v1/UpdateService'
            if ($us.Oem -and $us.Oem.Hpe -and $us.Oem.Hpe.UpdateTaskQueue -and $us.Oem.Hpe.UpdateTaskQueue."@odata.id") {
                $QueueUri = "$($us.Oem.Hpe.UpdateTaskQueue."@odata.id")"
            }
        } catch {
            if ($Strict) { throw "UpdateTaskQueue-URI konnte nicht ermittelt werden: $($_.Exception.Message)" }
        }
    }
    $tasks = @()
    try {
        $q = ILO-Get -Ilo $Ilo -Token $Token -Path $QueueUri
        foreach ($ref in @($q.Members)) {
            $u = $ref."@odata.id"
            if (-not $u) { continue }
            try {
                $t = ILO-Get -Ilo $Ilo -Token $Token -Path $u
                $fn = ''; $componentUri = ''
                if ($t.PSObject.Properties.Name -contains 'Filename' -and $t.Filename) { $fn = "$($t.Filename)" }
                if ($t.PSObject.Properties.Name -contains 'Component' -and $t.Component) {
                    if ($t.Component -is [string]) { $componentUri = "$($t.Component)" }
                    elseif ($t.Component.'@odata.id') { $componentUri = "$($t.Component.'@odata.id')" }
                }
                $tasks += [pscustomobject]@{
                    Name = "$($t.Name)"; Filename = $fn; ComponentUri = $componentUri
                    Id = "$($t.Id)"; Uri = "$u"; State = "$($t.State)"
                    Command = "$($t.Command)"; UpdatableBy = @($t.UpdatableBy)
                    Targets = @($t.Targets); Result = $t.Result
                }
            } catch {
                if ($Strict) { throw "UpdateTaskQueue-Member '$u' konnte nicht gelesen werden: $($_.Exception.Message)" }
            }
        }
    } catch {
        if ($Strict) { throw "UpdateTaskQueue konnte nicht vollstaendig gelesen werden: $($_.Exception.Message)" }
    }
    return @($tasks)
}

function Test-UpdateTaskNoMatchingTarget {
    param($Task)
    if (-not $Task -or -not $Task.Result) { return $false }
    try {
        return (($Task.Result | ConvertTo-Json -Compress -Depth 8) -match 'NoMatchingTargetDeviceFound')
    } catch {
        return ("$($Task.Result)" -match 'NoMatchingTargetDeviceFound')
    }
}

function ILO-DeleteUpdateTask {
    param([string]$Ilo,[string]$Token,[string]$TaskUri,[scriptblock]$LogCb = $null)
    if ([string]::IsNullOrWhiteSpace($TaskUri) -or $TaskUri -notmatch '^/redfish/v1/UpdateService/UpdateTaskQueue/') {
        throw "Ungueltige UpdateTaskQueue-URI zum Loeschen: '$TaskUri'"
    }
    Invoke-RestMethod -Uri "https://$Ilo$TaskUri" -Method Delete `
        -Headers @{ 'X-Auth-Token' = $Token; 'OData-Version' = '4.0' } `
        -SkipCertificateCheck -TimeoutSec $script:IloTimeoutSec -ErrorAction Stop | Out-Null
    $remaining = @(ILO-GetUpdateTaskQueue -Ilo $Ilo -Token $Token -Strict)
    if (@($remaining | Where-Object { "$($_.Uri)" -ieq $TaskUri }).Count) {
        throw "Update-Task '$TaskUri' ist nach DELETE weiterhin in der Queue vorhanden."
    }
    if ($LogCb) {
        try { & $LogCb "Abgeschlossenen Fehler-Task aus der Queue entfernt: $TaskUri" } catch {}
    }
}

function ILO-RetireFailedUpdateTasks {
    param([string]$Ilo,[string]$Token,[scriptblock]$LogCb = $null)
    $retired = @()
    $tasks = @(ILO-GetUpdateTaskQueue -Ilo $Ilo -Token $Token -Strict)
    foreach ($task in $tasks) {
        # Nur endgueltig fehlgeschlagene Tasks entfernen. Pending, laufende und
        # erfolgreich abgeschlossene Tasks bleiben als autoritativer Zustand
        # unangetastet.
        if ("$($task.State)" -notmatch '^(Exception|Expired|Canceled|Cancelled|Error|Failed)$') { continue }
        if ($LogCb) {
            try { & $LogCb "Queue-Bereinigung: alter Fehler-Task '$($task.Name)' (State='$($task.State)') wird entfernt." } catch {}
        }
        ILO-DeleteUpdateTask -Ilo $Ilo -Token $Token -TaskUri "$($task.Uri)" -LogCb $LogCb
        $retired += $task
    }
    return @($retired)
}

function ILO-RetireNoMatchingTargetTasks {
    param([string]$Ilo,[string]$Token,[scriptblock]$LogCb = $null)
    $retired = @()
    $tasks = @(ILO-GetUpdateTaskQueue -Ilo $Ilo -Token $Token -Strict)
    foreach ($task in $tasks) {
        if ("$($task.State)" -notmatch '^(Exception|Error|Failed)$') { continue }
        if (-not (Test-UpdateTaskNoMatchingTarget -Task $task)) { continue }
        ILO-DeleteUpdateTask -Ilo $Ilo -Token $Token -TaskUri "$($task.Uri)" -LogCb $LogCb
        $retired += $task
    }
    return @($retired)
}

function Test-UpdateTaskMatchesComponent {
    param($Task,[string]$ComponentFileName,[string]$ComponentUri = '')
    if (-not $Task -or -not $ComponentFileName) { return $false }
    if ("$($Task.Filename)" -ieq $ComponentFileName) { return $true }
    if ($ComponentUri -and "$($Task.ComponentUri)" -ieq $ComponentUri) { return $true }
    # Alte Complete-Tasks koennen Filename und Component verlieren. Nur dann ist
    # der von uns erzeugte stabile Name ein zulaessiger Ersatzschluessel.
    if ($Task.Filename -or $Task.ComponentUri) { return $false }
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ComponentFileName)
    $safeBaseName = (($baseName -replace '[^A-Za-z0-9_-]', '-') -replace '-{2,}', '-')
    $stablePrefix = "Update-$safeBaseName"
    $identity = if ($stablePrefix.Length -gt 63) { $stablePrefix.Substring(0, 63) } else { $stablePrefix }
    $taskName = "$($Task.Name)"
    return $taskName -ieq $identity -or
        $taskName.StartsWith("$identity-", [System.StringComparison]::OrdinalIgnoreCase)
}

function ILO-WaitForUpdateTaskVisibility {
    param(
        [string]$Ilo,[string]$Token,[string]$QueueUri,
        [string]$TaskName,[string]$ComponentFileName,
        [string[]]$BeforeNames = @(),[string[]]$BeforeUris = @(),
        [int]$TimeoutSec = 90
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $lastError = ''
    $queueWasReadable = $false
    while ((Get-Date) -lt $deadline) {
        try {
            $tasks = @(ILO-GetUpdateTaskQueue -Ilo $Ilo -Token $Token -QueueUri $QueueUri -Strict)
            $queueWasReadable = $true
            $mine = $tasks | Where-Object {
                $_.Name -eq $TaskName -or
                ($_.Name -notin $BeforeNames -and "$($_.Filename)" -ieq $ComponentFileName)
            } | Select-Object -First 1
            if ($mine) {
                return [pscustomobject]@{ Task = $mine; QueueWasReadable = $true; LastError = '' }
            }
        } catch {
            $lastError = $_.Exception.Message
            # Bei sofort ausgefuehrten iLO-6-Tasks kann der Collection-Member
            # bereits angelegt sein, waehrend sein Detail-Endpunkt noch HTTP 500
            # liefert. Eine genau eine neue Member-URI beweist die Task-Erstellung
            # eindeutig, ohne einen riskanten zweiten POST.
            try {
                $collection = ILO-Get -Ilo $Ilo -Token $Token -Path $QueueUri
                $newUris = @($collection.Members | ForEach-Object { "$($_.'@odata.id')" } | Where-Object {
                    $_ -and $_ -notin $BeforeUris
                } | Select-Object -Unique)
                if ($newUris.Count -eq 1) {
                    $task = [pscustomobject]@{
                        Name = $TaskName; Filename = $ComponentFileName
                        Uri = $newUris[0]; State = ''
                    }
                    return [pscustomobject]@{ Task = $task; QueueWasReadable = $false; LastError = $lastError }
                }
            } catch {}
        }
        Start-Sleep -Seconds 2
    }
    return [pscustomobject]@{ Task = $null; QueueWasReadable = $queueWasReadable; LastError = $lastError }
}

# Bildet HPEs serverseitiges Zustandsmodell nach:
#   FirmwareInventory = aktive Version
#   ComponentRepository = bereitgestellte Version + Aktivierungsart
#   UpdateTaskQueue = Ausfuehrungszustand
# Nur Status 'New' darf einen neuen Upload/Task ausloesen.
function ILO-GetFirmwareServerDisposition {
    param(
        [string]$Ilo,[string]$Token,[string]$ComponentFileName,
        [string]$TargetVersion,[object[]]$InventoryMatches = @()
    )
    $repo = @(ILO-GetComponentRepository -Ilo $Ilo -Token $Token -Strict)
    $queue = @(ILO-GetUpdateTaskQueue -Ilo $Ilo -Token $Token -Strict)
    $globalStall = $queue | Where-Object { "$($_.State)" -match '^(Exception|Expired|Error|Failed)$' } | Select-Object -First 1
    $globalRunning = $queue | Where-Object { "$($_.State)" -match '^(Pending|InProgress|Running|Verifying|Updating)$' } | Select-Object -First 1
    $components = @($repo | Where-Object { "$($_.Filename)" -ieq $ComponentFileName })
    if ($components.Count -gt 1) {
        return [pscustomobject]@{ Status = 'Blocked'; Reason = 'mehrere Repository-Eintraege mit identischem Filename'; Component = $null; Task = $null }
    }
    if (-not $components.Count) {
        # Tasks koennen laenger sichtbar bleiben als ihr Repository-Verweis. Ein
        # passender Task ist auch ohne Repository-Member autoritativ und darf nie
        # durch einen erneuten Upload/POST dupliziert werden.
        $orphanTasks = @($queue | Where-Object {
            Test-UpdateTaskMatchesComponent -Task $_ -ComponentFileName $ComponentFileName
        })
        $orphanFailed = $orphanTasks | Where-Object { "$($_.State)" -match '^(Exception|Expired|Error|Failed)$' } | Select-Object -First 1
        if ($orphanFailed) {
            if (Test-UpdateTaskNoMatchingTarget -Task $orphanFailed) {
                return [pscustomobject]@{ Status = 'NotApplicable'; Reason = 'iLO bestaetigt: kein passendes Zielgeraet fuer diese Firmware'; Component = $null; Task = $orphanFailed }
            }
            return [pscustomobject]@{ Status = 'Blocked'; Reason = "bestehender Task ohne Repository-Verweis hat State='$($orphanFailed.State)'"; Component = $null; Task = $orphanFailed }
        }
        $orphanRunning = $orphanTasks | Where-Object { "$($_.State)" -match '^(InProgress|Running|Verifying|Updating)$' } | Select-Object -First 1
        if ($orphanRunning) {
            return [pscustomobject]@{ Status = 'InProgress'; Reason = "bestehender Task ohne Repository-Verweis hat State='$($orphanRunning.State)'"; Component = $null; Task = $orphanRunning }
        }
        $orphanPending = $orphanTasks | Where-Object { "$($_.State)" -eq 'Pending' } | Select-Object -First 1
        if ($orphanPending) {
            return [pscustomobject]@{ Status = 'Pending'; Reason = 'bestehender Task wartet ohne lesbaren Repository-Verweis in der Queue'; Component = $null; Task = $orphanPending }
        }
        $orphanComplete = $orphanTasks | Where-Object { "$($_.State)" -eq 'Complete' } | Select-Object -First 1
        if ($orphanComplete) {
            return [pscustomobject]@{ Status = 'Blocked'; Reason = 'bestehender Complete-Task ohne Repository-Verweis; Aktivierungsstatus ist nicht sicher bestimmbar'; Component = $null; Task = $orphanComplete }
        }
        if ($globalStall) {
            return [pscustomobject]@{ Status = 'QueueBlocked'; Reason = "Queue enthaelt Task '$($globalStall.Name)' mit State='$($globalStall.State)'"; Component = $null; Task = $globalStall }
        }
        if ($globalRunning) {
            return [pscustomobject]@{ Status = 'QueueBusy'; Reason = "Queue verarbeitet Task '$($globalRunning.Name)'"; Component = $null; Task = $globalRunning }
        }
        return [pscustomobject]@{ Status = 'New'; Reason = 'nicht im Repository'; Component = $null; Task = $null }
    }

    $component = $components[0]
    $componentUri = "$($component.'@odata.id')"
    $repoVersion = "$($component.Version)"
    $targetNorm = Get-NormFwVersion $TargetVersion
    $repoNorm = Get-NormFwVersion $repoVersion
    if ($targetNorm -and $repoNorm -and (Compare-FwVersion $repoNorm $targetNorm) -ne 0) {
        return [pscustomobject]@{ Status = 'Blocked'; Reason = "Repository-Version '$repoVersion' stimmt nicht mit Ziel '$TargetVersion' ueberein"; Component = $component; Task = $null }
    }

    # Repository-Targets/DeviceClass sind die autoritative HPE-Korrelation.
    $repoTargets = @($component.Targets | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
    $repoClass = "$($component.DeviceClass)".Trim()
    $correlated = @($InventoryMatches)
    if ($repoTargets.Count) {
        $correlated = @($InventoryMatches | Where-Object {
            $invClass = "$($_.DeviceClass)".Trim()
            if ($repoClass -and $invClass -and $repoClass -ine $invClass) { return $false }
            $invIds = @($_.DeviceIdentifiers | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
            return @($repoTargets | Where-Object { $invIds -contains $_ }).Count -gt 0
        })
    }
    if (-not $correlated.Count) {
        return [pscustomobject]@{ Status = 'Blocked'; Reason = 'Repository-Targets passen zu keinem Live-Inventarziel'; Component = $component; Task = $null }
    }

    $activeTarget = $false
    if ($targetNorm) {
        $activeTarget = @($correlated | Where-Object { (Compare-FwVersion $_.Version $TargetVersion) -eq 0 }).Count -gt 0
    }
    $tasks = @($queue | Where-Object {
        Test-UpdateTaskMatchesComponent -Task $_ -ComponentFileName $ComponentFileName -ComponentUri $componentUri
    })
    $failed = $tasks | Where-Object { "$($_.State)" -match '^(Exception|Expired|Canceled|Cancelled|Error|Failed)$' } | Select-Object -First 1
    if ($failed) {
        if (Test-UpdateTaskNoMatchingTarget -Task $failed) {
            return [pscustomobject]@{ Status = 'NotApplicable'; Reason = 'iLO bestaetigt: kein passendes Zielgeraet fuer diese Firmware'; Component = $component; Task = $failed }
        }
        return [pscustomobject]@{ Status = 'Blocked'; Reason = "bestehender Task State='$($failed.State)'"; Component = $component; Task = $failed }
    }
    $running = $tasks | Where-Object { "$($_.State)" -match '^(InProgress|Running|Verifying|Updating)$' } | Select-Object -First 1
    if ($running) {
        return [pscustomobject]@{ Status = 'InProgress'; Reason = "Task State='$($running.State)'"; Component = $component; Task = $running }
    }
    $pending = $tasks | Where-Object { "$($_.State)" -eq 'Pending' } | Select-Object -First 1
    if ($pending) {
        return [pscustomobject]@{ Status = 'Pending'; Reason = 'Task wartet in der UpdateTaskQueue'; Component = $component; Task = $pending }
    }
    $complete = $tasks | Where-Object { "$($_.State)" -eq 'Complete' } | Select-Object -First 1
    if ($complete) {
        if ($activeTarget) {
            return [pscustomobject]@{ Status = 'Active'; Reason = 'Task Complete und Zielversion aktiv'; Component = $component; Task = $complete }
        }
        $activates = "$($component.Activates)"
        return [pscustomobject]@{ Status = 'PendingActivation'; Reason = "Task Complete; Aktivierung='$activates'; aktive Version noch abweichend"; Component = $component; Task = $complete }
    }
    if ($activeTarget) {
        return [pscustomobject]@{ Status = 'Active'; Reason = 'Zielversion im FirmwareInventory aktiv'; Component = $component; Task = $null }
    }
    if ($globalStall) {
        return [pscustomobject]@{ Status = 'QueueBlocked'; Reason = "Queue enthaelt Task '$($globalStall.Name)' mit State='$($globalStall.State)'"; Component = $component; Task = $globalStall }
    }
    if ($globalRunning) {
        return [pscustomobject]@{ Status = 'QueueBusy'; Reason = "Queue verarbeitet Task '$($globalRunning.Name)'"; Component = $component; Task = $globalRunning }
    }
    if ([bool]$component.Locked) {
        return [pscustomobject]@{ Status = 'Blocked'; Reason = 'Repository-Komponente ist Locked, aber kein auswertbarer Task vorhanden'; Component = $component; Task = $null }
    }
    # Repository ist nur die Bereitstellung, nicht der Nachweis eines Updates.
    # Stimmen Filename, Version, DeviceClass und Targets exakt mit Paket und
    # Live-Inventar ueberein, kann die vorhandene Komponente ohne erneuten Upload
    # gezielt in die Queue gestellt werden.
    return [pscustomobject]@{ Status = 'RepositoryReady'; Reason = 'passende Komponente liegt im Repository; aktive Version ist aelter und kein Task vorhanden'; Component = $component; Task = $null }
}

# Legt einen Update-Task in der iLO UpdateTaskQueue an (deferred-Aktivierung beim
# naechsten Reboot/POST). Liefert die Location/URI des erzeugten Tasks (oder '').
#
# WICHTIG: iLO quittiert einen UpdateTaskQueue-POST je nach Firmware auch dann mit
# 2xx, wenn der Task NICHT angelegt wird (z.B. unbekanntes Zusatzfeld wie
# 'TPMOverride', oder ein Filename der nicht exakt einem Repository-Eintrag
# entspricht). Genau das fuehrt zum Symptom "liegt im Repository, aber nicht in
# der Queue". Diese Funktion verifiziert daher NACH jedem POST durch erneutes
# Lesen der Queue, ob der Task wirklich existiert, und probiert andernfalls die
# naechste Payload-Variante.
function ILO-CreateUpdateTask {
    param(
        [string]$Ilo,[string]$Token,
        [string]$ComponentFileName,
        [string[]]$UpdatableBy = @('Uefi'),
        [string[]]$Targets = @(),
        [string]$TaskName = $null,
        [scriptblock]$LogCb = $null
    )
    $say = { param($m) if ($LogCb) { try { & $LogCb $m } catch {} } }
    $Targets = @($Targets | ForEach-Object { "$_".Trim() } | Where-Object {
        $_ -match '^/redfish/v1/UpdateService/FirmwareInventory/'
    } | Select-Object -Unique)
    if ($Targets.Count -gt 10) { throw "Ein Update-Task darf hoechstens 10 FirmwareInventory-Targets enthalten; erhalten: $($Targets.Count)." }

    # UpdateTaskQueue-URI ermitteln (Oem/Hpe-Link, sonst Standardpfad).
    $queueUri = '/redfish/v1/UpdateService/UpdateTaskQueue'
    try {
        $us = ILO-Get -Ilo $Ilo -Token $Token -Path "/redfish/v1/UpdateService"
        if ($us.Oem -and $us.Oem.Hpe -and $us.Oem.Hpe.UpdateTaskQueue -and $us.Oem.Hpe.UpdateTaskQueue."@odata.id") {
            $queueUri = "$($us.Oem.Hpe.UpdateTaskQueue."@odata.id")"
        }
    } catch {}

    # Den EXAKTEN Repository-Dateinamen ermitteln: iLO nutzt den Filename als
    # eindeutigen Schluessel; der Task muss genau diesen Namen referenzieren -
    # sonst wird kein Task angelegt (POST evtl. trotzdem 2xx).
    $match = ILO-FindRepositoryComponent -Ilo $Ilo -Token $Token -ComponentFileName $ComponentFileName -Strict
    if ($match -and $match.Filename) { $ComponentFileName = "$($match.Filename)" }
    else { throw "Repository-Komponente '$ComponentFileName' wurde nicht exakt gefunden. Kein Update-Task wird angelegt." }

    # Task-Name muss eindeutig sein UND wird Teil der iLO-Task-URI -> nur
    # URI-sichere Zeichen (keine Leerzeichen/Punkte), sonst 400 Bad Request.
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ComponentFileName)
    $safeBaseName = (($baseName -replace '[^A-Za-z0-9_-]', '-') -replace '-{2,}', '-')
    $stableTaskPrefix = "Update-$safeBaseName"
    if (-not $TaskName) { $TaskName = "$stableTaskPrefix-$((Get-Date).ToString('yyyyMMddHHmmss'))" }
    $TaskName = ($TaskName -replace '[^A-Za-z0-9_-]', '-') -replace '-{2,}', '-'
    if ($TaskName.Length -gt 63) { $TaskName = $TaskName.Substring(0, 63) }

    # Bestehende Tasks merken -> hinterher pruefen, ob unserer NEU dazukam.
    # Fail-closed: Kann die bestehende Queue nicht VOLLSTAENDIG gelesen werden,
    # darf niemals ein neuer Task erzeugt werden - sonst waere ein Duplikat nicht
    # sicher auszuschliessen.
    $before = @(ILO-GetUpdateTaskQueue -Ilo $Ilo -Token $Token -QueueUri $queueUri -Strict)
    $beforeNames = @($before | ForEach-Object { $_.Name })
    $beforeUris = @($before | ForEach-Object { "$($_.Uri)" } | Where-Object { $_ })

    # IDEMPOTENZ: Existiert bereits ein Task fuer genau diese Datei (z.B. aus einem
    # frueheren, noch nicht per Reboot aktivierten Lauf)? Dann ist nichts zu tun.
    # Abgeschlossene iLO-Tasks liefern je nach Firmware kein Filename/Component
    # mehr. Deshalb zusaetzlich den von uns stabil erzeugten Tasknamen-Prefix
    # pruefen. Bei sehr langen Dateinamen entspricht der auf 63 Zeichen gekuerzte
    # Prefix dem vollstaendigen gespeicherten Tasknamen.
    $already = $before | Where-Object { Test-UpdateTaskMatchesComponent -Task $_ -ComponentFileName $ComponentFileName } | Select-Object -First 1
    if ($already) {
        & $say "  Task fuer '$ComponentFileName' ist bereits in der Queue (State='$($already.State)', Name='$($already.Name)') - kein erneuter POST noetig."
        if ($already.Uri) { return "$($already.Uri)" }
        return ''
    }
    $queueBlocker = $before | Where-Object { "$($_.State)" -match '^(Pending|InProgress|Running|Verifying|Updating|Exception|Expired|Error|Failed)$' } | Select-Object -First 1
    if ($queueBlocker) {
        throw "UpdateTaskQueue ist blockiert: Task '$($queueBlocker.Name)' hat State='$($queueBlocker.State)'. Kein neuer Task wird angelegt."
    }

    # Zuerst ausschliesslich das von HPE dokumentierte Payload senden. iLO 5
    # kann einen Task mit unbekanntem Zusatzfeld wie TPMOverride zwar mit 2xx
    # quittieren, ihn aber nicht in der Queue speichern. Nur wenn iLO den ersten
    # POST EINDEUTIG mit FWFlashTPMOverrideFlagRequired ablehnt, folgt genau ein
    # zweiter POST mit TPMOverride=true. Andere Probe-POSTs sind unzulaessig.
    # GetNewClosure() besitzt einen eigenen Script-Scope. Deshalb muss der
    # globale Timeout vorher als lokaler Wert erfasst werden.
    $requestTimeoutSec = if ([int]$script:IloTimeoutSec -gt 0) { [int]$script:IloTimeoutSec } else { 40 }
    $sendTaskPost = {
        param([bool]$WithTpmOverride)
        $bodyObj = [ordered]@{
            Name        = $TaskName
            Filename    = $ComponentFileName
            Command     = 'ApplyUpdate'
            UpdatableBy = @($UpdatableBy)
        }
        if ($WithTpmOverride) { $bodyObj['TPMOverride'] = $true }
        $postBody = $bodyObj | ConvertTo-Json -Compress
        try {
            $resp = Invoke-WebRequest -Uri "https://$Ilo$queueUri" -Method Post -Body $postBody `
                -ContentType 'application/json' -Headers @{ 'X-Auth-Token' = $Token; 'OData-Version' = '4.0' } `
                -SkipCertificateCheck -TimeoutSec $requestTimeoutSec -ErrorAction Stop
            $postLoc = $resp.Headers['Location']; if ($postLoc -is [array]) { $postLoc = $postLoc[0] }
            if (-not $postLoc) { try { $j = $resp.Content | ConvertFrom-Json; if ($j."@odata.id") { $postLoc = "$($j."@odata.id")" } } catch {} }
            return [pscustomobject]@{ Accepted = $true; Body = $postBody; Location = $postLoc; Detail = ''; ErrorMessage = ''; OutcomeUnclear = $false }
        } catch {
            $detail = ''
            try { if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detail = $_.ErrorDetails.Message } } catch {}
            if (-not $detail) {
                try {
                    $rs = $_.Exception.Response.GetResponseStream()
                    if ($rs) { $sr = New-Object System.IO.StreamReader($rs); $detail = $sr.ReadToEnd(); $sr.Close() }
                } catch {}
            }
            return [pscustomobject]@{
                Accepted      = $false
                Body          = $postBody
                Location      = $null
                Detail        = $detail
                ErrorMessage  = $_.Exception.Message
                OutcomeUnclear = -not $detail
            }
        }
    }.GetNewClosure()

    $post = & $sendTaskPost $false
    if (-not $post.Accepted) {
        & $say "  Task-POST abgelehnt: $($post.ErrorMessage)$(if($post.Detail){" | $($post.Detail)"})"
        $requiresTpmOverride = "$($post.ErrorMessage) $($post.Detail)" -match '(?i)FWFlashTPMOverrideFlagRequired|TPMOverride.*Required'
        if ($requiresTpmOverride) {
            & $say '  iLO fordert TPMOverride explizit an; wiederhole den dokumentierten Task-POST einmalig mit TPMOverride=true.'
            $post = & $sendTaskPost $true
            if (-not $post.Accepted) {
                & $say "  Task-POST mit TPMOverride abgelehnt: $($post.ErrorMessage)$(if($post.Detail){" | $($post.Detail)"})"
            }
        }
    }
    $body = $post.Body
    $lastDetail = $post.Detail
    $loc = $post.Location
    $postOutcomeUnclear = [bool]$post.OutcomeUnclear

    $visibility = ILO-WaitForUpdateTaskVisibility -Ilo $Ilo -Token $Token `
        -QueueUri $queueUri -TaskName $TaskName -ComponentFileName $ComponentFileName `
        -BeforeNames $beforeNames -BeforeUris $beforeUris -TimeoutSec 90
    $mine = $visibility.Task
    if ($mine) {
        if ($mine.Uri) { return "$($mine.Uri)" }
        if ($loc) { try { return ([Uri]$loc).AbsolutePath } catch { return "$loc" } }
        return ''
    }
    if ($postOutcomeUnclear -or -not $visibility.QueueWasReadable) {
        $visibilityDetail = if ($visibility.LastError) { " Letzter Queue-Fehler: $($visibility.LastError)" } else { '' }
        $unclear = New-Object System.InvalidOperationException("Ausgang des einzigen Task-POST fuer '$ComponentFileName' ist nicht sicher bestimmbar. Kein weiterer POST wird gesendet.$visibilityDetail")
        $unclear.Data['KeepSubmittingMarker'] = $true
        throw $unclear
    }
    $qNow = @(ILO-GetUpdateTaskQueue -Ilo $Ilo -Token $Token -QueueUri $queueUri -Strict)
    $qTxt = if ($qNow.Count) { ($qNow | ForEach-Object { "$($_.Name)=$($_.Filename)" }) -join '; ' } else { '(leer)' }
    throw "UpdateTaskQueue-Task konnte nicht angelegt werden fuer '$ComponentFileName'. Payload=$body. iLO-Antwort=$lastDetail. Queue jetzt: $qTxt"
}


# Waehlt den passenden Firmware-Unterordner. Bevorzugt wird die neue,
# generationsbasierte Struktur des Downloaders: gen10, gen11 oder gen12.
# Darin liegen alle passenden Firmwarepakete; nicht verbaute Komponenten werden
# anschliessend pro Server anhand des iLO-Firmware-Inventars uebersprungen.
# Die bisherige modellbasierte Struktur bleibt als Rueckwaertskompatibilitaet.
# Liefert den vollstaendigen Ordnerpfad oder wirft einen Fehler.
function Resolve-FirmwareFolder {
    param([string]$BaseDir,[string]$Model,[int]$Gen)
    if (-not (Test-Path -LiteralPath $BaseDir -PathType Container)) { throw "Firmware-Verzeichnis nicht gefunden: $BaseDir" }
    $subs = @(Get-ChildItem -LiteralPath $BaseDir -Directory -ErrorAction Stop)
    if ($subs.Count -eq 0) { throw "Keine Typ-Unterordner im Firmware-Verzeichnis: $BaseDir" }

    # Neue Standardstruktur des Firmware-Downloaders: exakt ein Ordner je
    # Generation (gen10/gen11/gen12; optional auch fwpp-gen12 akzeptieren).
    if ($Gen -gt 0) {
        $genFolder = @($subs | Where-Object { $_.Name -match ("(?i)^(?:fwpp[ _-]?)?gen[ _-]?0*{0}$" -f $Gen) }) | Select-Object -First 1
        if ($genFolder) { return $genFolder.FullName }
    }

    # Legacy-Regel 1: Ordnername == kompletter Modellname (normalisiert, ohne Trenner).
    #   "ProLiant DL380a Gen11" == Ordner "ProLiant DL380a Gen11" oder "ProLiant_DL380a_Gen11"
    $normModel = ($Model -replace '[^A-Za-z0-9]', '').ToLower()
    if ($normModel) {
        foreach ($d in $subs) {
            $normDir = ($d.Name -replace '[^A-Za-z0-9]', '').ToLower()
            if ($normDir -eq $normModel) { return $d.FullName }
        }
    }

    # Legacy-Regel 2: Schema <TYP>_Gen<NN> - exakte Uebereinstimmung Typ + Generation.
    $modelType = $null
    if ($Model -match '\b([A-Z]{2}\d{2,4}[A-Za-z]?)\b') { $modelType = $Matches[1] }
    $normType = if ($modelType) { $modelType.ToLower() } else { '' }
    if ($normType -and $Gen -gt 0) {
        foreach ($d in $subs) {
            if ($d.Name -match '(?i)^(.+?)[ _-]*Gen[ _]?0*(\d+)\b') {
                $fType = ($Matches[1] -replace '[^A-Za-z0-9]', '').ToLower()
                $fGen  = [int]$Matches[2]
                if ($fType -eq $normType -and $fGen -eq $Gen) { return $d.FullName }
            }
        }
    }
    throw "Kein passender Typ-Ordner fuer Modell '$Model' (Typ '$modelType', Gen$Gen). Vorhandene Ordner: $((@($subs | ForEach-Object Name)) -join ', ')"
}

# Bestimmt anhand des Dateinamens den Komponententyp.
# Liefert @{ Kind; InvPattern; IsIlo; Deferred; UpdatableBy }
#  - IsIlo:    iLO-eigene Firmware. Nach dem Flash macht das iLO einen Selbst-Reset
#              und ist sofort aktiv -> KEIN Server-Reboot noetig.
#  - Deferred: Komponente kann NICHT sofort (online) geflasht werden. Sie wird ins
#              iLO-Repository hochgeladen (UpdateTarget=false) und ueber einen Task
#              in der UpdateTaskQueue beim naechsten Reboot/POST aktiviert
#              (z.B. SPS/ME, CPLD). Direktes Flashen liefert hier den Zustand 'Error'.
#  - UpdatableBy: Update-Agent fuer den UpdateTaskQueue-Task (Uefi=beim POST,
#              Bmc=durch iLO, RuntimeAgent=durch SUM/SUT im OS).
#  - InvPattern: Regex zum Finden des passenden Eintrags im Firmware-Inventar.
function Get-ComponentKind {
    param([string]$FileName)
    $n = "$FileName"
    if ($n -match '(?i)ilo')                                          { return @{ Kind = 'iLO';  InvPattern = '(?i)iLO';                                            IsIlo = $true;  Deferred = $false; UpdatableBy = @('Bmc')  } }
    if ($n -match '(?i)sps|server.?platform|(^|[^a-z])me[_\-]')       { return @{ Kind = 'SPS';  InvPattern = '(?i)SPS|Server Platform Services|Management Engine'; IsIlo = $false; Deferred = $true;  UpdatableBy = @('Uefi') } }
    if ($n -match '(?i)cpld|programmable.?logic')                     { return @{ Kind = 'CPLD'; InvPattern = '(?i)CPLD|Programmable Logic';                        IsIlo = $false; Deferred = $true;  UpdatableBy = @('Uefi') } }
    if ($n -match '(?i)(^|[^a-z])ie[_\-]|innovation.?engine')         { return @{ Kind = 'IE';   InvPattern = '(?i)Innovation Engine|(^|\W)IE(\W|$)';                 IsIlo = $false; Deferred = $true;  UpdatableBy = @('Uefi') } }
    if ($n -match '(?i)system.?rom|bios|romflash|^[a-z]\d{2}[_\-]')   { return @{ Kind = 'ROM';  InvPattern = '(?i)System ROM|BIOS';                                IsIlo = $false; Deferred = $false; UpdatableBy = @('Uefi') } }
    if ($n -match '(?i)power.?management|(^|[^a-z])pmc')              { return @{ Kind = 'PMC';  InvPattern = '(?i)Power Management';                               IsIlo = $false; Deferred = $false; UpdatableBy = @('Bmc')  } }
    # Der Dateiname liefert nur Typ und Inventarmuster. Deferred/UpdatableBy
    # sind hier keine Routing-Entscheidung; diese kommt ausschliesslich aus JSON.
    return @{ Kind = 'Other'; InvPattern = ''; IsIlo = $false; Deferred = $true; UpdatableBy = @('Uefi') }
}

# Verbindliche Update-Reihenfolge: iLO-Pakete kommen immer nach allen anderen
# Komponenten, da der iLO-Selbstreset insbesondere bei Synergy die Verbindung
# ueber OneView fuer laengere Zeit blockieren kann.
function Sort-FirmwareComponentsForUpdate {
    param([string[]]$Components)
    return @($Components | Sort-Object `
        @{ Expression = {
            $name = [System.IO.Path]::GetFileName($_)
            if ((Get-ComponentKind -FileName $name).IsIlo) { 1 } else { 0 }
        }; Ascending = $true }, `
        @{ Expression = { [System.IO.Path]::GetFileName($_) }; Ascending = $true })
}

# Liest Metadaten zu einer .fwpkg-Firmwaredatei.
# Zwei HPE-Liefervarianten werden unterstuetzt:
#   Gen10/10Plus/11: Metadaten als 'payload.json' IM .fwpkg-ZIP eingebettet.
#   Gen12:           Metadaten als SEPARATE Sidecar-Datei NEBEN der .fwpkg mit
#                    gleichem Basisnamen (z.B. ilo7_1.22.00.fwpkg ->
#                    ilo7_1.22.00.json). Diese wird BEVORZUGT ausgewertet.
# Liefert @{ Version; Deferred; Reboot; UpdatableBy; InventoryTerms;
# DeviceIdentifiers; DeviceClasses } mit
# Tri-State-Logik fuer Deferred/Reboot, dem vom Paket vorgegebenen Update-Agenten
# und den Geraetebezeichnungen fuer den Inventar-Abgleich:
#   Deferred: $true = nicht direkt flashbar (Repository + Task Queue)
#             $false = direkt online flashbar
#             $null  = keine eindeutige JSON-Auskunft (Paket wird gesperrt)
#   Reboot:   $true/$false/$null analog (Aktivierung erfordert Reboot?)
# Maßgebliche Felder (Schreibweise variiert je Generation, (?i)/Property-Zugriff
# decken beide ab):
#   "DirectFlashOk"/"DirectFlashOK": true/false -> false => Deferred
#   "UefiFlashable"/"UEFIFlashable": true/false -> true  => Deferred
#   "ResetRequired": true/false   -> Reboot (Alias: RebootRequired)
#   "ServerPowerOff": true/false  -> vollstaendiges Power Off zur Aktivierung
#   "UpdatableBy": ["Bmc"|"Uefi"|"RuntimeAgent"] -> nennt den zulaessigen
#       Ausfuehrungsagenten, entscheidet aber NICHT allein ueber den Flash-Weg.
#       Ausschliesslich DirectFlashOk/UefiFlashable bestimmen, ob direkt
#       geflasht oder ueber Repository + UpdateTaskQueue gearbeitet wird.
#   "Version"/"FirmwareVersion"/"TargetVersion"
function Get-FwpkgMeta {
    param([string]$FilePath)
    if (-not $script:FwpkgMetaCache) { $script:FwpkgMetaCache = @{} }
    if ($script:FwpkgMetaCache.ContainsKey($FilePath)) { return $script:FwpkgMetaCache[$FilePath] }

    # Nimmt nur konkrete, menschenlesbare Geraetebezeichnungen auf. IDs,
    # Versionsnummern und allgemeine Begriffe sind fuer den Abgleich gegen das
    # FirmwareInventory nicht eindeutig genug und duerfen KEIN Update ausloesen.
    $addInventoryTerms = {
        param($Value, $res)
        foreach ($raw in @($Value)) {
            if ($null -eq $raw -or $raw -is [System.Collections.IDictionary] -or $raw -is [System.Management.Automation.PSCustomObject]) { continue }
            $term = "$raw".Trim()
            if ($term.Length -lt 6 -or $term -match '(?i)^(hpe|hewlett[ -]?packard|firmware|component|device|unknown|n/?a|system|server|platform|management|controller|adapter|storage|network|power)$') { continue }
            # Ein Servermodell ist keine Hardwarekomponente. Eine Generation in
            # einem konkreten Geraetenamen (z.B. "Gen12 Boot Controller NS204i-u")
            # bleibt dagegen erhalten, damit sie den Namens-Fallback einschraenkt.
            if ($term -match '(?i)\bproliant\b|\bsystem\s*(rom|firmware)\b|\bserver\s*platform\b') { continue }
            if ($term -match '^https?://' -or $term -match '^[0-9a-f]{4,}([:-][0-9a-f]{2,})+$') { continue }
            if ($term -notmatch '[A-Za-z]') { continue }
            if (-not ($res.InventoryTerms -contains $term)) { $res.InventoryTerms += $term }
        }
    }
    # Hardware-IDs werden nicht heuristisch gematcht, aber im Diagnoseexport
    # ausgegeben. Fuer die Updateentscheidung werden jedoch ausschliesslich
    # Paket-Target/Targets gegen FirmwareInventory.Oem.Hpe.Targets verglichen.
    $addDeviceIdentifiers = {
        param($Value, $res)
        foreach ($raw in @($Value)) {
            if ($null -eq $raw -or $raw -is [System.Collections.IDictionary] -or $raw -is [System.Management.Automation.PSCustomObject]) { continue }
            $id = "$raw".Trim()
            if (-not $id -or $id.Length -lt 2) { continue }
            if (-not ($res.DeviceIdentifiers -contains $id)) { $res.DeviceIdentifiers += $id }
        }
    }
    $addDeviceClasses = {
        param($Value, $res)
        foreach ($raw in @($Value)) {
            if ($null -eq $raw -or $raw -is [System.Collections.IDictionary] -or $raw -is [System.Management.Automation.PSCustomObject]) { continue }
            $class = "$raw".Trim()
            if (-not $class) { continue }
            if (-not ($res.DeviceClasses -contains $class)) { $res.DeviceClasses += $class }
        }
    }

    # Wertet ein einzelnes JSON-Objekt aus und setzt die Rueckgabefelder.
    $parseObj = {
        param($o, $res, [bool]$IsDevice = $false)
        if (-not $o) { return }
        # Version (mehrere moegliche Schluesselnamen)
        foreach ($vk in @('Version', 'FirmwareVersion', 'TargetVersion')) {
            if (-not $res.Version -and $o.PSObject.Properties.Name -contains $vk) {
                $cand = "$($o.$vk)".Trim()
                if ($cand -match '\d') { $res.Version = $cand; $res.VersionSource = 'JSON'; break }
            }
        }
        # NUR die konkreten Device-Objekte enthalten eine belastbare Zuordnung
        # zur Hardware. Top-Level-/Image-Namen sind oft generisch (z.B. "System"
        # oder "Firmware") und wuerden sonst fast jedes Inventar falsch matchen.
        if ($IsDevice) {
            foreach ($nk in @('DeviceName', 'ComponentName')) {
                if ($o.PSObject.Properties.Name -contains $nk) { & $addInventoryTerms $o.$nk $res }
            }
            # Ausschliesslich Target/Targets entspricht HPEs dokumentierter
            # Korrelation gegen FirmwareInventory.Oem.Hpe.Targets.
            # DeviceInstance sowie einzelne PCI-Felder sind keine Paket-Targets.
            foreach ($ik in @('Target', 'Targets')) {
                if ($o.PSObject.Properties.Name -contains $ik) { & $addDeviceIdentifiers $o.$ik $res }
            }
        }
        # DeviceClass ist die von HPE dokumentierte zusaetzliche Schranke vor
        # dem Target-Abgleich. Je nach Paketformat steht sie im Device-, Image-
        # oder Top-Level-Objekt; daher wird sie in allen Ebenen gelesen.
        if ($o.PSObject.Properties.Name -contains 'DeviceClass') { & $addDeviceClasses $o.DeviceClass $res }
        # DirectFlashOk ist die autoritative Routing-Angabe. UefiFlashable wird
        # nur verwendet, wenn DirectFlashOk im Paket vollstaendig fehlt.
        if ($o.PSObject.Properties.Name -contains 'DirectFlashOk') {
            try {
                $route = -not [bool]$o.DirectFlashOk
                if ($res.HasDirectFlashOk -and $res.DirectFlashRoute -ne $route) { $res.DirectRouteConflict = $true }
                $res.HasDirectFlashOk = $true
                $res.DirectFlashRoute = $route
                $res.Deferred = $route
                $res.RouteConflict = $res.DirectRouteConflict
                $res.RouteSources += "DirectFlashOk=$($o.DirectFlashOk)"
            } catch {}
        }
        # UefiFlashable -> Deferred (true bedeutet Flash via UEFI beim Reboot),
        # jedoch nur als Fallback ohne DirectFlashOk.
        if ($o.PSObject.Properties.Name -contains 'UefiFlashable') {
            try {
                $route = [bool]$o.UefiFlashable
                if ($res.HasUefiFlashable -and $res.UefiFlashRoute -ne $route) { $res.UefiRouteConflict = $true }
                $res.HasUefiFlashable = $true
                $res.UefiFlashRoute = $route
                if (-not $res.HasDirectFlashOk) {
                    $res.Deferred = $route
                    $res.RouteConflict = $res.UefiRouteConflict
                }
                $res.RouteSources += "UefiFlashable=$($o.UefiFlashable)"
            } catch {}
        }
        # ResetRequired / RebootRequired -> Reboot
        if ($null -eq $res.Reboot) {
            foreach ($rk in @('ResetRequired', 'RebootRequired')) {
                if ($o.PSObject.Properties.Name -contains $rk) {
                    try { $res.Reboot = [bool]$o.$rk; break } catch {}
                }
            }
        }
        # Ein erforderliches Power Off ist staerker als ein normaler Reboot.
        # Sobald eine anwendbare Paket-Payload es verlangt, bleibt die Anforderung
        # gesetzt; spaetere false-Werte anderer Image-Eintraege heben sie nicht auf.
        if ($o.PSObject.Properties.Name -contains 'ServerPowerOff') {
            try {
                $powerOff = [bool]$o.ServerPowerOff
                if ($powerOff -or $null -eq $res.PowerOff) { $res.PowerOff = $powerOff }
            } catch {}
        }
        # UpdatableBy aus dem Paket uebernehmen. Das ist insbesondere fuer
        # weitere Komponenten (NIC/HBA/Storage/PSU/...) wichtig: ihre fwpkg
        # bestimmt, ob iLO, UEFI beim POST oder der Runtime-Agent aktiv wird.
        if ($o.PSObject.Properties.Name -contains 'UpdatableBy' -and @($o.UpdatableBy).Count) {
            $agents = @($o.UpdatableBy | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
            if ($agents.Count -and -not $res.UpdatableBy.Count) { $res.UpdatableBy = $agents }
        }
    }

    # Sammelt aus einem payload.json-Objekt alle relevanten Teil-Objekte und
    # bringt sie in Auswertungsreihenfolge (spezifischste zuerst). HPE FWPKG-v2
    # verschachtelt die Flags: Devices.Device[].FirmwareImages[] (DirectFlashOk,
    # UefiFlashable, ResetRequired), Version steht in Devices.Device[], und
    # UpdatableBy auf der obersten Ebene. Aelteres/flaches Format wird ebenso
    # unterstuetzt (Top-Objekt bzw. Components[]).
    $flatten = {
        param($o)
        $list = New-Object System.Collections.ArrayList
        if (-not $o) { return @() }
        $tops = @($o)
        foreach ($t in $tops) {
            if (-not $t) { continue }
            # FWPKG-v2: Devices.Device[] -> FirmwareImages[]
            if ($t.PSObject.Properties.Name -contains 'Devices' -and $t.Devices -and ($t.Devices.PSObject.Properties.Name -contains 'Device')) {
                foreach ($dev in @($t.Devices.Device)) {
                    if (-not $dev) { continue }
                    if ($dev.PSObject.Properties.Name -contains 'FirmwareImages') {
                        foreach ($img in @($dev.FirmwareImages)) { if ($img) { [void]$list.Add([pscustomobject]@{ Data = $img; IsDevice = $false }) } }
                    }
                    [void]$list.Add([pscustomobject]@{ Data = $dev; IsDevice = $true })
                }
            }
            # Anderes Format: Components[]
            if ($t.PSObject.Properties.Name -contains 'Components') {
                foreach ($c in @($t.Components)) { if ($c) { [void]$list.Add([pscustomobject]@{ Data = $c; IsDevice = $true }) } }
            }
            # Zuletzt das Top-Objekt selbst (UpdatableBy etc.)
            [void]$list.Add([pscustomobject]@{ Data = $t; IsDevice = $false })
        }
        return $list.ToArray()
    }

    $res = @{
        Version = ''; VersionSource = ''; Deferred = $null; RouteConflict = $false; RouteSources = @()
        HasDirectFlashOk = $false; DirectFlashRoute = $null; DirectRouteConflict = $false
        HasUefiFlashable = $false; UefiFlashRoute = $null; UefiRouteConflict = $false
        Reboot = $null; PowerOff = $null; UpdatableBy = @(); InventoryTerms = @(); DeviceIdentifiers = @(); DeviceClasses = @()
    }

    # Wertet einen JSON-Text strukturiert aus (flatten + parseObj). Faellt bei
    # nicht-parsebarem JSON auf gezielte Regex-Suche zurueck. (?i) deckt die
    # verschiedenen HPE-Schreibweisen ab (DirectFlashOk/DirectFlashOK,
    # UefiFlashable/UEFIFlashable).
    $applyText = {
        param($txt)
        if (-not $txt) { return }
        try {
            $obj = $txt | ConvertFrom-Json -ErrorAction Stop
            foreach ($it in @(& $flatten $obj)) { & $parseObj $it.Data $res $it.IsDevice }
        } catch {
            if (-not $res.Version -and $txt -match '(?i)"(FirmwareVersion|TargetVersion|Version)"\s*:\s*"([^"]+)"') {
                $cand = $Matches[2].Trim(); if ($cand -match '\d') { $res.Version = $cand; $res.VersionSource = 'JSON' }
            }
            if ($null -eq $res.Deferred -and $txt -match '(?i)"DirectFlashOk"\s*:\s*(true|false)') {
                $res.Deferred = ($Matches[1] -ieq 'false'); $res.RouteSources += "DirectFlashOk=$($Matches[1])"
            }
            if ($null -eq $res.Deferred -and $txt -match '(?i)"UefiFlashable"\s*:\s*(true|false)') {
                $res.Deferred = ($Matches[1] -ieq 'true'); $res.RouteSources += "UefiFlashable=$($Matches[1])"
            }
            if ($null -eq $res.Reboot   -and $txt -match '(?i)"(ResetRequired|RebootRequired)"\s*:\s*(true|false)') { $res.Reboot = ($Matches[2] -ieq 'true') }
            if ($null -eq $res.PowerOff -and $txt -match '(?i)"ServerPowerOff"\s*:\s*(true|false)') { $res.PowerOff = ($Matches[1] -ieq 'true') }
        }
    }
    # Sind alle drei Felder bestimmt?
    $complete = { ($res.Version -and $null -ne $res.Deferred -and $null -ne $res.Reboot) }

    # 0) Gen12: separate Sidecar-JSON NEBEN der .fwpkg (gleicher Basisname).
    try {
        $sidecar = [System.IO.Path]::ChangeExtension($FilePath, '.json')
        if ($sidecar -and (Test-Path -LiteralPath $sidecar)) {
            & $applyText ([System.IO.File]::ReadAllText($sidecar))
        }
    } catch {}

    # 1)/2) Gen10/10Plus/11: eingebettete payload.json bzw. andere *.json im ZIP.
    if (-not (& $complete)) {
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
            $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
            try {
                # payload.json bevorzugt (HPE-dokumentierte Quelle).
                $payload = @($zip.Entries | Where-Object { $_.Name -match '(?i)^payload\.json$' })[0]
                if ($payload) {
                    $sr = New-Object System.IO.StreamReader($payload.Open())
                    try { $txt = $sr.ReadToEnd() } finally { $sr.Dispose() }
                    & $applyText $txt
                }
                # Fallback: andere *.json-Eintraege im Paket.
                if (-not (& $complete)) {
                    foreach ($e in $zip.Entries) {
                        if ($e.Name -notmatch '(?i)\.json$') { continue }
                        if ($e.Name -match '(?i)^payload\.json$') { continue }
                        if ($e.Length -gt 4MB) { continue }
                        $sr = New-Object System.IO.StreamReader($e.Open())
                        try { $txt = $sr.ReadToEnd() } finally { $sr.Dispose() }
                        & $applyText $txt
                        if (& $complete) { break }
                    }
                }
            } finally { $zip.Dispose() }
        } catch {}
    }

    if (-not $res.Version) {
        $fn = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        # iLO: ilo5_318 -> 3.18 / ilo6_164 -> 1.64
        if ($fn -match '(?i)ilo\d?[ _\-]?(\d)(\d{2})$') { $res.Version = "$($Matches[1]).$($Matches[2])"; $res.VersionSource = 'Filename' }
        elseif ($fn -match '(\d+\.\d+(?:\.\d+)*)')      { $res.Version = $Matches[1]; $res.VersionSource = 'Filename' }
    }
    $script:FwpkgMetaCache[$FilePath] = $res
    return $res
}

# Ermittelt die zu einer Firmwaredatei passenden, auf diesem Server wirklich
# vorhandenen Firmware-Inventar-Eintraege. Die bekannten Basiskomponenten werden
# ueber feste Muster gematcht; bei allen anderen Komponenten verwendet die
# Funktion die Geraetebezeichnungen aus payload.json bzw. der Sidecar-JSON.
function ILO-FindInstalledFirmware {
    param(
        [object[]]$Inventory,
        [string]$InvPattern = '',
        [string[]]$InventoryTerms = @(),
        [string[]]$DeviceIdentifiers = @(),
        [string[]]$DeviceClasses = @()
    )
    # Nicht '$matches' nennen: PowerShell verwendet '$Matches' automatisch fuer
    # Regex-Treffer und unterscheidet bei Variablennamen nicht nach Gross-/Kleinschreibung.
    $matchedItems = New-Object System.Collections.ArrayList
    $add = {
        param($item)
        if (-not $item) { return }
        if (-not (@($matchedItems | Where-Object { "$($_.Name)" -ieq "$($item.Name)" }).Count)) { [void]$matchedItems.Add($item) }
    }

    # Primaerer Abgleich: die exakten HPE/PLDM-Geraetekennungen aus dem Paket
    # gegen Oem.Hpe.Targets des Server-Inventars. Bei Basis-Firmware teilen
    # mehrere Inventarelemente manchmal eine Plattformkennung. Dann schraenkt
    # das bekannte InvPattern (ROM/iLO/SPS/IE/...) den exakten ID-Treffer auf
    # die richtige Firmwarekomponente ein.
    $packageIds = @($DeviceIdentifiers | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
    $packageClasses = @($DeviceClasses | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
    $packageGens = @($InventoryTerms | ForEach-Object {
        [regex]::Matches("$_", '(?i)\bGen\s*(1[0-2])\b') | ForEach-Object { $_.Groups[1].Value }
    } | Select-Object -Unique)
    if ($packageIds.Count) {
        $idHits = @()
        foreach ($item in @($Inventory)) {
            # Entspricht dem HPE-Referenzalgorithmus: Sobald iLO eine
            # DeviceClass liefert, muss eine vorhandene Paketklasse dazu
            # passen. Bei alten Paketformaten ohne DeviceClass bleibt der
            # exakte Target-Abgleich weiterhin die belastbare Grundlage.
            $inventoryClass = "$($item.DeviceClass)".Trim()
            if ($inventoryClass -and $packageClasses.Count -and -not (@($packageClasses | Where-Object { $_ -ieq $inventoryClass }).Count)) { continue }
            $inventoryGens = @([regex]::Matches("$($item.Name)", '(?i)\bGen\s*(1[0-2])\b') | ForEach-Object { $_.Groups[1].Value })
            if ($packageGens.Count -and $inventoryGens.Count -and -not @($packageGens | Where-Object { $inventoryGens -contains $_ }).Count) { continue }
            $itemIds = @($item.DeviceIdentifiers | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
            $found = $false
            foreach ($id in $itemIds) {
                if (@($packageIds | Where-Object { $_ -ieq $id }).Count) { $found = $true; break }
            }
            if ($found) { $idHits += $item }
        }
        if ($InvPattern) {
            foreach ($item in @($idHits | Where-Object { "$($_.Name)" -match $InvPattern })) { & $add $item }
        } else {
            foreach ($item in @($idHits)) { & $add $item }
        }
        # Eine Paket-ID, die im Serverinventar nicht vorkommt, ist ein
        # eindeutiger Ausschluss. Kein Rueckfall auf Namen: das waere wieder
        # die Ursache fuer die zuvor beobachteten 195 Fehlkandidaten.
        return @($matchedItems)
    }

    # Legacy-Fallback fuer alte Pakete ohne Targets/Device-ID.
    if ($InvPattern) {
        foreach ($item in @($Inventory | Where-Object { "$($_.Name)" -match $InvPattern })) { & $add $item }
    }
    foreach ($rawTerm in @($InventoryTerms)) {
        $term = "$rawTerm".Trim()
        if ($term.Length -lt 6) { continue }
        $words = @($term -split '[^A-Za-z0-9]+' | Where-Object {
            $_.Length -ge 3 -and $_ -notmatch '(?i)^(hpe|firmware|adapter|controller|device|component|system|server|platform|management|storage|network|power|services)$'
        })
        # Eine komponentenspezifische Kennung braucht mehrere relevante Woerter,
        # eine Modellnummer oder einen Hersteller. Allgemeine Bezeichnungen wie
        # "System Firmware" duerfen nie als vorhandene Hardware gelten.
        $hasVendor = $term -match '(?i)broadcom|mellanox|emulex|qlogic|marvell|intel|amd|nvidia|smart.?array|microchip|lsi|seagate|samsung|micron|kingston'
        $hasNumber = $term -match '\d'
        if ($words.Count -lt 2 -or (-not $hasVendor -and -not $hasNumber -and $words.Count -lt 3)) { continue }
        foreach ($item in @($Inventory)) {
            $invName = "$($item.Name)".Trim()
            if (-not $invName) { continue }
            # Wenn Paket- und Inventarname beide eine Geraetegeneration nennen,
            # muss mindestens eine Generation uebereinstimmen. So darf ein reines
            # Gen12-NS204i-Paket nicht auf einen Gen11-Controller geroutet werden.
            $termGens = @([regex]::Matches($term, '(?i)\bGen\s*(1[0-2])\b') | ForEach-Object { $_.Groups[1].Value })
            $invGens = @([regex]::Matches($invName, '(?i)\bGen\s*(1[0-2])\b') | ForEach-Object { $_.Groups[1].Value })
            if ($termGens.Count -and $invGens.Count -and -not @($termGens | Where-Object { $invGens -contains $_ }).Count) { continue }
            # Nur der vollstaendige, konkrete Paketname darf als unmittelbarer
            # Teil des iLO-Inventarnamens gelten. Die umgekehrte Richtung waere
            # zu breit (z.B. Inventar "System" in fast jedem Firmwarepaket).
            if ($invName.IndexOf($term, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                & $add $item; continue
            }
            # Fallback fuer leicht unterschiedliche Bezeichnungen: bei einer
            # Modellnummer/Herstellerkennung reichen zwei, sonst drei relevante
            # Woerter. Damit bleiben falsch-positive Gen-/System-Treffer aus.
            $requiredWords = if ($hasVendor -or $hasNumber) { 2 } else { 3 }
            $hitWords = @($words | Where-Object { $invName.IndexOf($_, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 })
            if ($hitWords.Count -ge $requiredWords) { & $add $item }
        }
    }
    return @($matchedItems)
}

# Erkennt uneindeutige Zuordnungen VOR einem Upload. Ein Firmware-Inventar-
# Eintrag darf nur zu genau einer Paketdatei passen; treffen mehrere Pakete
# zu, ist die Metadatenzuordnung nicht belastbar genug und alle Beteiligten
# werden fail-closed uebersprungen.
function Get-AmbiguousFirmwareMatches {
    param(
        [object[]]$Inventory,
        [string[]]$Components,
        [string[]]$KindFilter = @()
    )
    $owners = @{}
    foreach ($comp in @($Components)) {
        if (-not (Test-Path -LiteralPath $comp -PathType Leaf)) { continue }
        $name = [System.IO.Path]::GetFileName($comp)
        $kind = Get-ComponentKind -FileName $name
        if (@($KindFilter).Count -and (@($KindFilter) -notcontains $kind.Kind)) { continue }
        $meta = Get-FwpkgMeta -FilePath $comp
        # Unbekannte/geraetespezifische Komponenten duerfen ausschliesslich ueber
        # JSON Target/Targets gegen FirmwareInventory.Oem.Hpe.Targets zugeordnet
        # werden. Ohne diese IDs gibt es keinen Namens-Fallback und keinen Upload.
        if ($kind.Kind -in @('Other','SPS','CPLD','IE','PMC') -and -not @($meta.DeviceIdentifiers).Count) {
            continue
        }
        $matches = @(ILO-FindInstalledFirmware -Inventory $Inventory -InvPattern $kind.InvPattern -InventoryTerms $meta.InventoryTerms -DeviceIdentifiers $meta.DeviceIdentifiers -DeviceClasses $meta.DeviceClasses)
        foreach ($item in $matches) {
            $key = "$($item.Name)".Trim()
            if (-not $key) { continue }
            if (-not $owners.ContainsKey($key)) { $owners[$key] = @() }
            if ($owners[$key] -notcontains $name) { $owners[$key] += $name }
        }
    }
    $ambiguous = @{}
    foreach ($key in @($owners.Keys)) {
        if (@($owners[$key]).Count -gt 1) { $ambiguous[$key] = @($owners[$key]) }
    }
    return $ambiguous
}

# Erzeugt ausschliesslich aus lokal vorhandenen Firmwaredateien und dem bereits
# gelesenen Server-Inventar eine Update-Vorschau. Diese Funktion schreibt nichts
# auf den Server und wird sowohl fuer die Inventar-Pruefung als auch fuer die
# eigentliche Update-Entscheidung nach denselben Matching-Regeln verwendet.
function Get-FirmwareUpdatePreview {
    param(
        [object[]]$Inventory,
        $SystemInfo,
        [string[]]$Components,
        [string[]]$KindFilter = @(),
        [bool]$AllowDowngrade = $false,
        [bool]$StrictVersionComparison = $false
    )
    $updates = New-Object System.Collections.ArrayList
    $current = New-Object System.Collections.ArrayList
    $absent = New-Object System.Collections.ArrayList
    $ambiguous = New-Object System.Collections.ArrayList
    $notOnlineUpdateable = New-Object System.Collections.ArrayList
    $repositoryQueue = New-Object System.Collections.ArrayList
    $downgrades = New-Object System.Collections.ArrayList
    $ambiguousMap = Get-AmbiguousFirmwareMatches -Inventory $Inventory -Components $Components -KindFilter $KindFilter
    foreach ($comp in @($Components)) {
        if (-not (Test-Path -LiteralPath $comp -PathType Leaf)) { continue }
        $name = [System.IO.Path]::GetFileName($comp)
        $kind = Get-ComponentKind -FileName $name
        if (@($KindFilter).Count -and (@($KindFilter) -notcontains $kind.Kind)) { continue }
        $meta = Get-FwpkgMeta -FilePath $comp
        $target = $meta.Version
        if ($StrictVersionComparison -and (-not $target -or $meta.VersionSource -ne 'JSON')) {
            [void]$notOnlineUpdateable.Add("$name (strikter Versionsschutz: keine Zielversion aus Paket-JSON)")
            continue
        }
        if ($meta.RouteConflict -or $null -eq $meta.Deferred) {
            $routeProblem = if ($meta.RouteConflict) { 'widerspruechliche DirectFlashOk/UefiFlashable-Angaben' } else { 'DirectFlashOk/UefiFlashable fehlt' }
            [void]$notOnlineUpdateable.Add("$name (JSON-Routing unklar: $routeProblem)")
            continue
        }
        $installed = @(ILO-FindInstalledFirmware -Inventory $Inventory -InvPattern $kind.InvPattern -InventoryTerms $meta.InventoryTerms -DeviceIdentifiers $meta.DeviceIdentifiers -DeviceClasses $meta.DeviceClasses)
        if (-not $installed.Count) {
            [void]$absent.Add($name)
            continue
        }
        if (@($installed | Where-Object { $ambiguousMap.ContainsKey("$($_.Name)") }).Count) {
            [void]$ambiguous.Add($name)
            continue
        }
        $notUpdateableMatches = @($installed | Where-Object { $null -ne $_.Updateable -and -not [bool]$_.Updateable })
        $allNotUpdateable = $installed.Count -and $notUpdateableMatches.Count -eq $installed.Count
        $routeText = if ($meta.PowerOff) { ' [Power Off erforderlich]' } else { '' }
        $queueRoute = [bool]$meta.Deferred
        $queueAgents = @($meta.UpdatableBy | Where-Object { "$_" -match '^(?i:Bmc|Uefi|RuntimeAgent)$' })
        if ($queueRoute -and -not $queueAgents.Count) {
            [void]$notOnlineUpdateable.Add("$name (JSON verlangt Repository/Queue, aber UpdatableBy fehlt)")
            continue
        }
        if ($queueRoute) {
            $routeText += ' [Repository/Queue]'
        } elseif ($allNotUpdateable) {
            $items = (@($notUpdateableMatches | ForEach-Object { "$($_.Name)=$($_.Version)" }) -join ', ')
            [void]$notOnlineUpdateable.Add("$name ($items; JSON verlangt Direct-Flash, iLO meldet Updateable=false)")
            continue
        }
        $targetNorm = Get-NormFwVersion $target
        $alreadyCurrent = $false
        if ($targetNorm) {
            $alreadyCurrent = @($installed | Where-Object { (Compare-FwVersion $_.Version $target) -eq 0 }).Count -gt 0
            # Beim BIOS zaehlt auch die bereits geschriebene Backup-ROM als
            # aktuell/vorgemerkt, damit sie nicht erneut als Update erscheint.
            if ($kind.Kind -eq 'ROM' -and -not $alreadyCurrent) {
                $currentRom = Get-NormFwVersion $SystemInfo.BiosCurrent
                $backupRom = Get-NormFwVersion $SystemInfo.BiosBackup
                $alreadyCurrent = ((Compare-FwVersion $targetNorm $currentRom) -eq 0 -or
                    (Compare-FwVersion $targetNorm $backupRom) -eq 0)
            }
        }
        if ($alreadyCurrent) {
            [void]$current.Add($name)
        } else {
            # Downgrade-Schutz: Ist die Zielversion EINDEUTIG aelter als die
            # hoechste bereits installierte Version der zugeordneten Komponente,
            # wird das Paket nur mit gesetzter Option 'Downgrades zulassen'
            # geflasht. Nicht numerisch vergleichbare Versionen bleiben unberuehrt.
            $downgradeFrom = $null
            $unverifiableVersion = $null
            if ($target -and -not $AllowDowngrade) {
                foreach ($inst in @($installed)) {
                    $cmp = if ($StrictVersionComparison) {
                        Compare-FwVersionStrict $target $inst.Version
                    } else {
                        Compare-FwVersion $target $inst.Version
                    }
                    if ($StrictVersionComparison -and $null -eq $cmp) {
                        $unverifiableVersion = "$($inst.Version)"
                        break
                    }
                    if ($null -ne $cmp -and $cmp -lt 0) {
                        if (-not $downgradeFrom -or ((Compare-FwVersion $inst.Version $downgradeFrom) -gt 0)) { $downgradeFrom = $inst.Version }
                    }
                }
                if ($kind.Kind -eq 'ROM' -and -not $downgradeFrom -and $SystemInfo.BiosCurrent) {
                    $cmpRom = if ($StrictVersionComparison) {
                        Compare-FwVersionStrict $target $SystemInfo.BiosCurrent
                    } else {
                        Compare-FwVersion $target $SystemInfo.BiosCurrent
                    }
                    if ($null -ne $cmpRom -and $cmpRom -lt 0) { $downgradeFrom = $SystemInfo.BiosCurrent }
                }
            }
            if ($unverifiableVersion) {
                [void]$notOnlineUpdateable.Add("$name (strikter Versionsschutz: Ziel '$target' nicht sicher mit installiert '$unverifiableVersion' vergleichbar)")
            } elseif ($downgradeFrom) {
                [void]$downgrades.Add("$name (Ziel $target ist aelter als installiert $downgradeFrom)")
            } else {
                $targetText = if ($target) { " -> $target" } else { ' (Zielversion nicht lesbar)' }
                [void]$updates.Add("$name$targetText$routeText")
                if ($queueRoute) {
                    $items = (@($installed | ForEach-Object { "$($_.Name)=$($_.Version)" }) -join ', ')
                    [void]$repositoryQueue.Add("$name ($items; Repository/Queue)")
                }
            }
        }
    }
    [pscustomobject]@{ Updates = @($updates); Current = @($current); Absent = @($absent); Ambiguous = @($ambiguous); NotOnlineUpdateable = @($notOnlineUpdateable); RepositoryQueue = @($repositoryQueue); Downgrades = @($downgrades) }
}

# Bereitet die verwendeten Paketmetadaten fuer einen lesbaren Diagnoseexport
# auf. Die DeviceIdentifiers werden bewusst nicht fuer einen heuristischen
# Flash-Entscheid verwendet: Erst der Vergleich mit dem realen iLO-Inventar
# zeigt, welche IDs sich generationuebergreifend belastbar zuordnen lassen.
function Get-FirmwarePackageDiagnostics {
    param([string[]]$Components)
    $packages = @()
    foreach ($comp in @($Components)) {
        if (-not (Test-Path -LiteralPath $comp -PathType Leaf)) { continue }
        $meta = Get-FwpkgMeta -FilePath $comp
        $kind = Get-ComponentKind -FileName ([System.IO.Path]::GetFileName($comp))
        $packages += [pscustomobject]@{
            File              = [System.IO.Path]::GetFileName($comp)
            FullPath          = $comp
            Kind              = $kind.Kind
            InventoryPattern  = $kind.InvPattern
            TargetVersion     = $meta.Version
            VersionSource     = $meta.VersionSource
            Deferred          = $meta.Deferred
            RouteConflict     = $meta.RouteConflict
            RouteSources      = @($meta.RouteSources)
            RebootRequired    = $meta.Reboot
            PowerOffRequired  = $meta.PowerOff
            UpdatableBy       = @($meta.UpdatableBy)
            InventoryTerms    = @($meta.InventoryTerms)
            DeviceIdentifiers = @($meta.DeviceIdentifiers)
            DeviceClasses     = @($meta.DeviceClasses)
        }
    }
    return @($packages)
}

# Protokolliert den Hardware-Abgleich fuer jedes lokale Paket. Damit ist im
# Diagnoseexport nachvollziehbar, warum ein Paket ausgewaehlt, ausgeschlossen
# oder wegen Updateable=false gesperrt wurde. Es ist dieselbe strikte
# ID-/DeviceClass-Pruefung wie im Updatepfad; diese Funktion selbst ist rein
# lesend und aendert weder iLO noch Repository.
function Get-FirmwareCorrelationDiagnostics {
    param([object[]]$Inventory,[string[]]$Components,[string[]]$KindFilter = @())
    $rows = @()
    foreach ($comp in @($Components)) {
        if (-not (Test-Path -LiteralPath $comp -PathType Leaf)) { continue }
        $file = [System.IO.Path]::GetFileName($comp)
        $kind = Get-ComponentKind -FileName $file
        if (@($KindFilter).Count -and (@($KindFilter) -notcontains $kind.Kind)) { continue }
        $meta = Get-FwpkgMeta -FilePath $comp
        $matches = @(ILO-FindInstalledFirmware -Inventory $Inventory -InvPattern $kind.InvPattern -InventoryTerms $meta.InventoryTerms -DeviceIdentifiers $meta.DeviceIdentifiers -DeviceClasses $meta.DeviceClasses)
        $rows += [pscustomobject]@{
            File              = $file
            Kind              = $kind.Kind
            TargetVersion     = $meta.Version
            VersionSource     = $meta.VersionSource
            DeviceIdentifiers = @($meta.DeviceIdentifiers)
            DeviceClasses     = @($meta.DeviceClasses)
            Matches           = @($matches | ForEach-Object {
                [pscustomobject]@{ Name = $_.Name; Version = $_.Version; Updateable = $_.Updateable; DeviceClass = $_.DeviceClass; DeviceIdentifiers = @($_.DeviceIdentifiers) }
            })
            MatchCount        = $matches.Count
            OnlineFlashAllowed = if ($matches.Count) { @($matches | Where-Object { $null -ne $_.Updateable -and -not [bool]$_.Updateable }).Count -ne $matches.Count } else { $false }
        }
    }
    return @($rows)
}

# Schreibt eine reine Diagnoseaufnahme auf den lokalen Rechner. Weder iLO noch
# Firmware-Repository oder Task Queue werden dabei veraendert. Die Datei ist
# absichtlich pro Server/Zeitschnitt getrennt, damit Vergleichsdaten mehrerer
# Generationen oder Hardware-Konfigurationen nicht vermischt werden.
function Write-FirmwareDiagnostic {
    param(
        [string]$StateDir,
        [string]$Ilo,
        $SystemInfo,
        [object[]]$FirmwareInventory,
        [object[]]$DeviceInventory,
        $StorageInventory,
        $UpdateDecision,
        [object[]]$FirmwarePackages
    )
    if (-not $StateDir) { return '' }
    try {
        $dir = Join-Path $StateDir 'Diagnostics'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $identity = "$($SystemInfo.Serial)"
        if (-not $identity) { $identity = $Ilo }
        $identity = $identity -replace '[^A-Za-z0-9._-]', '_'
        if (-not $identity) { $identity = 'unknown' }
        $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
        $path = Join-Path $dir ("FirmwareDiagnostic_{0}_{1}.json" -f $identity, $stamp)
        $report = [ordered]@{
            GeneratedAt            = (Get-Date).ToString('s')
            Ilo                    = $Ilo
            SystemInfo             = $SystemInfo
            FirmwareInventory      = @($FirmwareInventory)
            DeviceInventoryProbes  = @($DeviceInventory)
            StorageInventory       = $StorageInventory
            UpdateDecision         = $UpdateDecision
            FirmwarePackages       = @($FirmwarePackages)
        }
        $report | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $path -Encoding UTF8
        return $path
    } catch { return '' }
}

# Normalisiert eine Version fuer den Vergleich: erstes numerisches Segment,
# fuehrende Nullen entfernt (z.B. '04.01.05.201' -> '4.1.5.201').
function Get-NormFwVersion {
    param([string]$V)
    if (-not $V) { return '' }
    $m = [regex]::Match("$V", '\d+(?:\.\d+)+')
    if (-not $m.Success) { $m = [regex]::Match("$V", '\d+') }
    if (-not $m.Success) { return ("$V".Trim().ToLower()) }
    return (($m.Value -split '\.') | ForEach-Object { [int]$_ }) -join '.'
}

# Numerischer Versionsvergleich zweier (normalisierter) Firmwareversionen.
# Liefert -1 (A < B), 0 (gleich), 1 (A > B) oder $null, wenn mindestens eine
# Version nicht rein numerisch vergleichbar ist. Wird fuer den Downgrade-
# Schutz verwendet: nur bei einem EINDEUTIG aelteren Ziel wird gesperrt.
function Compare-FwVersion {
    param([string]$A, [string]$B)
    $na = Get-NormFwVersion $A
    $nb = Get-NormFwVersion $B
    if (-not $na -or -not $nb) { return $null }
    if ($na -notmatch '^\d+(?:\.\d+)*$' -or $nb -notmatch '^\d+(?:\.\d+)*$') { return $null }
    $pa = @($na -split '\.'); $pb = @($nb -split '\.')
    $len = [Math]::Max($pa.Count, $pb.Count)
    for ($i = 0; $i -lt $len; $i++) {
        $va = if ($i -lt $pa.Count) { [long]$pa[$i] } else { 0 }
        $vb = if ($i -lt $pb.Count) { [long]$pb[$i] } else { 0 }
        if ($va -lt $vb) { return -1 }
        if ($va -gt $vb) { return 1 }
    }
    return 0
}

function Get-StrictFwVersionDescriptor {
    param([string]$Version)

    $raw = "$Version".Trim()
    if (-not $raw) { return $null }

    # Unterstützte HPE-Darstellungen mit mehrteiliger Nummer:
    # 3.21, v3.66 (04/01/2026), U30 v3.10 (02/22/2024),
    # 3.21 Jul 30 2026 sowie iLO-7-JSON 1.24.00_pass_17.
    # Nur der von HPE verwendete _pass_<Build>-Zusatz ist zugelassen;
    # beliebige Qualifier wie 1.2-beta bleiben absichtlich gesperrt.
    $dotted = [regex]::Match(
        $raw,
        '(?ix)^(?:[A-Z]\d+\s+)?v?(?<version>\d+(?:\.\d+)+)(?:' +
        '\s*\(\d{1,2}/\d{1,2}/\d{4}\)|' +
        '\s+[A-Z]{3}\s+\d{1,2}\s+\d{4}|' +
        '_pass_\d+' +
        ')?$'
    )
    if ($dotted.Success) {
        return [pscustomobject]@{ Numeric = $dotted.Groups['version'].Value; Signature = 'Dotted' }
    }

    # Einzelnummern sind nur vergleichbar, wenn der alphanumerische Rahmen
    # identisch ist, z.B. HPD6 < HPD8. HPD6 gegen ABC8 ist nicht belastbar.
    $single = [regex]::Match($raw, '^(?<prefix>[A-Za-z._-]*)(?<version>\d+)(?<suffix>[A-Za-z._-]*)$')
    if ($single.Success) {
        $signature = ("{0}|{1}" -f $single.Groups['prefix'].Value, $single.Groups['suffix'].Value).ToLowerInvariant()
        return [pscustomobject]@{ Numeric = $single.Groups['version'].Value; Signature = "Single:$signature" }
    }
    return $null
}

function Compare-FwVersionStrict {
    param([string]$A, [string]$B)

    $left = Get-StrictFwVersionDescriptor -Version $A
    $right = Get-StrictFwVersionDescriptor -Version $B
    if (-not $left -or -not $right -or $left.Signature -ne $right.Signature) { return $null }
    return Compare-FwVersion -A $left.Numeric -B $right.Numeric
}

# Ermittelt aus allen Paketen einer Generationsablage die tatsaechlich fuer
# diesen Server passenden Dateien. Pro Inventarkomponente bleibt hoechstens
# das Paket mit der hoechsten Zielversion uebrig; bei gleicher hoechster
# Version bleibt die Zuordnung absichtlich gesperrt statt beliebig zu flashen.
function Select-FirmwarePackagesForServer {
    param(
        [object[]]$Inventory,
        [string[]]$Components,
        [string[]]$KindFilter = @(),
        [bool]$StrictMetadata = $false
    )
    $owners = @{}
    $notInstalled = New-Object System.Collections.ArrayList
    $notOnlineUpdateable = New-Object System.Collections.ArrayList
    $repositoryQueue = New-Object System.Collections.ArrayList
    foreach ($comp in @($Components)) {
        if (-not (Test-Path -LiteralPath $comp -PathType Leaf)) { continue }
        $name = [System.IO.Path]::GetFileName($comp)
        $kind = Get-ComponentKind -FileName $name
        if (@($KindFilter).Count -and (@($KindFilter) -notcontains $kind.Kind)) { continue }
        $meta = Get-FwpkgMeta -FilePath $comp
        if ($StrictMetadata -and (-not $meta.Version -or $meta.VersionSource -ne 'JSON')) {
            [void]$notOnlineUpdateable.Add("$name (strikter Versionsschutz: keine Zielversion aus Paket-JSON)")
            continue
        }
        if ($meta.RouteConflict -or $null -eq $meta.Deferred) {
            $routeProblem = if ($meta.RouteConflict) { 'widerspruechliche DirectFlashOk/UefiFlashable-Angaben' } else { 'DirectFlashOk/UefiFlashable fehlt' }
            [void]$notOnlineUpdateable.Add("$name (JSON-Routing unklar: $routeProblem)")
            continue
        }
        if ($kind.Kind -in @('Other','SPS','CPLD','IE','PMC') -and -not @($meta.DeviceIdentifiers).Count) {
            if (-not ($notInstalled -contains $name)) { [void]$notInstalled.Add($name) }
            continue
        }
        $inventoryMatches = @(ILO-FindInstalledFirmware -Inventory $Inventory -InvPattern $kind.InvPattern -InventoryTerms $meta.InventoryTerms -DeviceIdentifiers $meta.DeviceIdentifiers -DeviceClasses $meta.DeviceClasses)
        if (-not $inventoryMatches.Count) {
            if (-not ($notInstalled -contains $name)) { [void]$notInstalled.Add($name) }
            continue
        }
        # Eine exakte Target-ID allein berechtigt noch nicht zum Upload. Die
        # Online-Freigabe wird mit dem Kandidaten mitgefuehrt und erst NACH der
        # Versionsauswahl ausgewertet. Dadurch wird bei mehreren Versionen nur
        # die neueste gesperrte Variante angezeigt, nicht jede Altversion.
        $notUpdateableMatches = @($inventoryMatches | Where-Object { $null -ne $_.Updateable -and -not [bool]$_.Updateable })
        $allNotUpdateable = $inventoryMatches.Count -and $notUpdateableMatches.Count -eq $inventoryMatches.Count
        $queueRoute = [bool]$meta.Deferred
        $queueAgents = @($meta.UpdatableBy | Where-Object { "$_" -match '^(?i:Bmc|Uefi|RuntimeAgent)$' })
        if ($queueRoute -and -not $queueAgents.Count) {
            [void]$notOnlineUpdateable.Add("$name (JSON verlangt Repository/Queue, aber UpdatableBy fehlt)")
            continue
        }
        $onlineFlashAllowed = (-not $queueRoute) -and (-not $allNotUpdateable)
        $queueByBmc = $queueRoute -and
            (@($meta.UpdatableBy | Where-Object { "$_" -ieq 'Bmc' }).Count -gt 0)
        foreach ($item in $inventoryMatches) {
            $key = "$($item.Name)".Trim()
            if (-not $key) { continue }
            if (-not $owners.ContainsKey($key)) { $owners[$key] = @() }
            $owners[$key] += [pscustomobject]@{ Path = $comp; Name = $name; Version = $meta.Version; OnlineFlashAllowed = $onlineFlashAllowed; QueueRoute = $queueRoute; QueueByBmc = $queueByBmc; InventoryName = $item.Name; InventoryVersion = $item.Version }
        }
    }

    $selected = @{}
    $blocked = @{}
    $ambiguous = New-Object System.Collections.ArrayList
    foreach ($key in @($owners.Keys)) {
        $candidates = @($owners[$key])
        $bestRank = $null
        foreach ($candidate in $candidates) {
            $norm = Get-NormFwVersion $candidate.Version
            # Jede numerische Version wird auf vier gepaddete Segmente
            # gebracht, damit 8.00 korrekt groesser als 6.52 einsortiert wird.
            $rank = '0000000000.0000000000.0000000000.0000000000'
            if ($norm -match '^\d+(?:\.\d+)*$') {
                $parts = @($norm -split '\.')
                $padded = @()
                for ($i = 0; $i -lt 4; $i++) {
                    $part = if ($i -lt $parts.Count) { [int64]$parts[$i] } else { 0 }
                    $padded += ('{0:D10}' -f $part)
                }
                $rank = $padded -join '.'
            }
            $candidate | Add-Member -NotePropertyName Rank -NotePropertyValue $rank -Force
            if ($null -eq $bestRank -or $rank -gt $bestRank) { $bestRank = $rank }
        }
        $best = @($candidates | Where-Object { $_.Rank -eq $bestRank })
        if ($best.Count -eq 1) {
            if ($best[0].OnlineFlashAllowed -or $best[0].QueueRoute) {
                $selected[$best[0].Path] = $true
                if ($best[0].QueueRoute) {
                    [void]$repositoryQueue.Add("$($best[0].Name) ($($best[0].InventoryName)=$($best[0].InventoryVersion); Repository/Queue)")
                }
            } else {
                $blocked[$best[0].Path] = $true
                [void]$notOnlineUpdateable.Add("$($best[0].Name) ($($best[0].InventoryName)=$($best[0].InventoryVersion); iLO: Updateable=false)")
            }
        } else {
            [void]$ambiguous.Add("${key}: $((@($best | ForEach-Object { $_.Name })) -join ', ')")
        }
    }
    $selectedComponents = @($Components | Where-Object { $selected.ContainsKey($_) })
    $superseded = @()
    foreach ($key in @($owners.Keys)) {
        foreach ($candidate in @($owners[$key])) {
            if (-not $selected.ContainsKey($candidate.Path) -and -not $blocked.ContainsKey($candidate.Path) -and $candidate.Name -notin $superseded) { $superseded += $candidate.Name }
        }
    }
    [pscustomobject]@{
        Components   = $selectedComponents
        NotInstalled = @($notInstalled)
        NotOnlineUpdateable = @($notOnlineUpdateable)
        RepositoryQueue = @($repositoryQueue)
        Superseded   = @($superseded)
        Ambiguous    = @($ambiguous)
    }
}

# ─────────────────────────────────────────
# Lokaler Marker fuer bereits geflashte, aber noch NICHT per Reboot
# aktivierte Komponenten (v.a. BIOS/System ROM).
# Hintergrund: HPE iLO schreibt ein Online-BIOS-Update in die Redundant/
# Backup-ROM und aktiviert es erst beim naechsten Reboot/POST. Das
# FirmwareInventory (Redfish) meldet die neue Version - weder unter
# 'System ROM' noch unter 'Redundant System ROM' - erst NACH dem Reboot.
# Ein zweiter Update-Lauf VOR dem Reboot wuerde den bereits erfolgten
# Flash daher nicht erkennen und unnoetig erneut flashen. Wir merken den
# Flash deshalb lokal (StagedFlashes.json, je Server-Seriennummer + Typ).
# Der Marker ist selbstheilend: sobald die Zielversion aktiv im Inventar
# erscheint (Reboot erfolgt), wird er entfernt.
function Get-StagedFlashPath {
    param([string]$StateDir)
    if (-not $StateDir) { $StateDir = $env:TEMP }
    Join-Path $StateDir 'StagedFlashes.json'
}
function Read-StagedFlashes {
    param([string]$StateDir)
    $path = Get-StagedFlashPath -StateDir $StateDir
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        if (-not $raw -or -not $raw.Trim()) { return @() }
        return @($raw | ConvertFrom-Json -ErrorAction Stop)
    } catch { return @() }
}
function Test-StagedFlash {
    # $true, wenn fuer (Serial, Kind) bereits ein Flash mit derselben
    # normalisierten Zielversion vorgemerkt ist.
    param([string]$StateDir,[string]$Serial,[string]$Kind,[string]$TargetNorm)
    if (-not $Serial -or -not $Kind -or -not $TargetNorm) { return $false }
    $mtx = New-Object System.Threading.Mutex($false,'Global\OV_RackmountFw_StagedFlash')
    try { [void]$mtx.WaitOne() } catch {}
    try {
        foreach ($e in (Read-StagedFlashes -StateDir $StateDir)) {
            if ("$($e.Serial)" -ieq $Serial -and "$($e.Kind)" -ieq $Kind -and "$($e.TargetNorm)" -eq $TargetNorm) { return $true }
        }
        return $false
    } finally { try { [void]$mtx.ReleaseMutex() } catch {}; $mtx.Dispose() }
}
function Get-StagedFlash {
    # Liefert den aktuellen Marker fuer (Serial, Kind), unabhaengig von der
    # Zielversion. Wird fuer die Reparatur eines versehentlich in das redundante
    # ROM geschriebenen BIOS-Downgrades verwendet.
    param([string]$StateDir,[string]$Serial,[string]$Kind)
    if (-not $Serial -or -not $Kind) { return $null }
    $mtx = New-Object System.Threading.Mutex($false,'Global\OV_RackmountFw_StagedFlash')
    try { [void]$mtx.WaitOne() } catch {}
    try {
        return Read-StagedFlashes -StateDir $StateDir |
            Where-Object { "$($_.Serial)" -ieq $Serial -and "$($_.Kind)" -ieq $Kind } |
            Select-Object -First 1
    } finally { try { [void]$mtx.ReleaseMutex() } catch {}; $mtx.Dispose() }
}
function Set-StagedFlash {
    # Legt/aktualisiert den Marker fuer (Serial, Kind) auf die Zielversion.
    param(
        [string]$StateDir,[string]$Serial,[string]$Kind,
        [string]$File,[string]$TargetNorm,
        [string]$State = 'PendingReboot'
    )
    if (-not $Serial -or -not $Kind) { return }
    $mtx = New-Object System.Threading.Mutex($false,'Global\OV_RackmountFw_StagedFlash')
    try { [void]$mtx.WaitOne() } catch {}
    try {
        $list = @(Read-StagedFlashes -StateDir $StateDir | Where-Object { -not ("$($_.Serial)" -ieq $Serial -and "$($_.Kind)" -ieq $Kind) })
        $now = Get-Date
        $list += [pscustomobject]@{
            Serial = $Serial; Kind = $Kind; File = $File; TargetNorm = $TargetNorm
            State = $State; StagedAt = $now.ToString('s')
            StagedAtUtc = $now.ToUniversalTime().ToString('o')
        }
        $path = Get-StagedFlashPath -StateDir $StateDir
        $dir = Split-Path -Parent $path
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        (@($list) | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $path -Encoding UTF8
    } catch {} finally { try { [void]$mtx.ReleaseMutex() } catch {}; $mtx.Dispose() }
}
function Clear-StagedFlash {
    # Entfernt den Marker fuer (Serial, Kind) - z.B. nachdem die neue
    # Version aktiv im Inventar erscheint (Reboot erfolgt).
    param([string]$StateDir,[string]$Serial,[string]$Kind)
    if (-not $Serial -or -not $Kind) { return }
    $mtx = New-Object System.Threading.Mutex($false,'Global\OV_RackmountFw_StagedFlash')
    try { [void]$mtx.WaitOne() } catch {}
    try {
        $before = @(Read-StagedFlashes -StateDir $StateDir)
        $list = @($before | Where-Object { -not ("$($_.Serial)" -ieq $Serial -and "$($_.Kind)" -ieq $Kind) })
        if ($list.Count -ne $before.Count) {
            $path = Get-StagedFlashPath -StateDir $StateDir
            if ($list.Count -eq 0) { Remove-Item -LiteralPath $path -ErrorAction SilentlyContinue }
            else { (@($list) | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $path -Encoding UTF8 }
        }
    } catch {} finally { try { [void]$mtx.ReleaseMutex() } catch {}; $mtx.Dispose() }
}
function Test-RebootAfterStagedFlash {
    param($Marker,[string]$LastResetTime)
    if (-not $Marker -or -not $LastResetTime) { return $false }
    $stagedValue = $Marker.StagedAtUtc
    if ($null -eq $stagedValue -or [string]::IsNullOrWhiteSpace("$stagedValue")) {
        $stagedValue = $Marker.StagedAt
    }
    if ($null -eq $stagedValue -or [string]::IsNullOrWhiteSpace("$stagedValue")) { return $false }
    try {
        $staged = if ($stagedValue -is [datetimeoffset]) {
            $stagedValue.ToUniversalTime()
        } elseif ($stagedValue -is [datetime]) {
            ([datetimeoffset]$stagedValue).ToUniversalTime()
        } else {
            [datetimeoffset]::Parse(
                "$stagedValue",
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::RoundtripKind
            ).ToUniversalTime()
        }
        $reset = [datetimeoffset]::Parse(
            $LastResetTime,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind
        ).ToUniversalTime()
        return $reset -gt $staged
    } catch {
        return $false
    }
}
'@

# Die Paketmetadaten werden im UI-Prozess einmal eingelesen und anschließend
# read-only an alle Server-Runspaces übergeben. Das verhindert bei großen
# Synergy-Flotten, dass jedes Paket für jeden Server erneut entpackt wird.
Invoke-Expression $script:iloCode
function Get-FirmwarePackageIndex {
    param([string[]]$Components)

    $index = @{}
    foreach ($component in @($Components)) {
        if (-not (Test-Path -LiteralPath $component -PathType Leaf)) { continue }
        if ($script:FwpkgMetaCache -and $script:FwpkgMetaCache.ContainsKey($component)) {
            [void]$script:FwpkgMetaCache.Remove($component)
        }
        $index[$component] = Get-FwpkgMeta -FilePath $component
    }
    return $index
}

# ============================================================================
#  OneView-REST-Helfer (UI-Thread) - liest iLO-Adressen + Modell der
#  monitored Server aus OneView (server-hardware). Keine HPE-Module noetig.
# ============================================================================
function OV-GetApiVersion {
    param([string]$A)
    $r = Invoke-RestMethod -Uri "https://$A/rest/version" -Method Get -SkipCertificateCheck -TimeoutSec 30 -ErrorAction Stop
    [int]$r.currentVersion
}
function OV-Login {
    param([string]$A,[string]$U,[string]$P,[int]$V)
    $b = @{ userName = $U; password = $P; authLoginDomain = "Local" } | ConvertTo-Json
    $h = @{ "Content-Type" = "application/json"; "X-API-Version" = "$V" }
    $r = Invoke-RestMethod -Uri "https://$A/rest/login-sessions" -Method Post -Body $b -Headers $h -SkipCertificateCheck -TimeoutSec 30 -ErrorAction Stop
    if ([string]::IsNullOrEmpty($r.sessionID)) { throw "Keine sessionID erhalten von $A" }
    $r.sessionID
}
function OV-Logout {
    param([string]$A,[string]$S,[int]$V)
    $h = @{ Auth = $S; "X-API-Version" = "$V" }
    try { Invoke-RestMethod -Uri "https://$A/rest/login-sessions" -Method Delete -Headers $h -SkipCertificateCheck -TimeoutSec 10 -EA SilentlyContinue | Out-Null } catch {}
}
function OV-Rest {
    param([string]$A,[string]$S,[int]$V,[string]$M,[string]$E)
    $h = @{ Auth = $S; "X-API-Version" = "$V" }
    Invoke-RestMethod -Uri "https://$A$E" -Method $M -Headers $h -ContentType "application/json" -SkipCertificateCheck -TimeoutSec 60 -ErrorAction Stop
}
function OV-RestAll {
    param([string]$A,[string]$S,[int]$V,[string]$E)
    $items = @(); $endpoint = $E
    while ($endpoint) {
        $page = OV-Rest -A $A -S $S -V $V -M Get -E $endpoint
        if ($page.members) { $items += $page.members }
        $endpoint = if ($page.nextPageUri) { $page.nextPageUri } else { $null }
    }
    ,$items
}

# Liefert eine Hashtable scopeUri -> scopeName (aus /rest/scopes).
function OV-GetScopeMap {
    param([string]$A,[string]$S,[int]$V)
    $map = @{}
    try {
        $scopes = OV-RestAll -A $A -S $S -V $V -E "/rest/scopes?count=1000"
        foreach ($sc in $scopes) { if ($sc.uri) { $map[$sc.uri] = $sc.name } }
    } catch {}
    $map
}

# Liefert eine Hashtable serverUri -> @(scopeUris) ueber den Index.
function OV-GetServerScopeUris {
    param([string]$A,[string]$S,[int]$V)
    $map = @{}
    try {
        $idx = OV-RestAll -A $A -S $S -V $V -E "/rest/index/resources?category=server-hardware&count=1000"
        foreach ($r in $idx) {
            if ($r.uri) {
                $uris = @()
                if ($r.PSObject.Properties.Name -contains 'scopeUris' -and $r.scopeUris) { $uris = @($r.scopeUris) }
                $map[$r.uri] = $uris
            }
        }
    } catch {}
    $map
}

# Zerlegt einen beliebigen iLO-Adresswert (String, Array, Objekt mit .address) in Strings.
function Expand-IloAddressValue {
    param($v)
    if ($null -eq $v) { return }
    if ($v -is [string]) {
        foreach ($t in ($v -split '[\s,;]+')) { $t2 = ($t + '').Trim().Trim('[', ']'); if ($t2) { Write-Output $t2 } }
        return
    }
    if ($v -is [System.Collections.IDictionary]) { if ($v.Contains('address')) { Expand-IloAddressValue -v $v['address'] }; return }
    if ($v -is [System.Collections.IEnumerable]) { foreach ($e in $v) { Expand-IloAddressValue -v $e }; return }
    if ($v.PSObject -and ($v.PSObject.Properties.Name -contains 'address')) { Expand-IloAddressValue -v $v.address; return }
    $s = "$v".Trim(); if ($s) { Write-Output $s }
}

# Liefert ausschliesslich eine gueltige IPv4-Adresse des iLO. IPv6 und
# DNS-Namen sind in den verwalteten Netzen nicht als Redfish-Ziel nutzbar.
function Get-IloIpFromHardware {
    param($sh)
    if (-not $sh) { return $null }
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $scored = New-Object System.Collections.Generic.List[object]
    $addOne = {
        param([string]$addr)
        if ([string]::IsNullOrWhiteSpace($addr)) { return }
        if ($addr -match '\s') { return }
        $parsed = $null
        if (-not [System.Net.IPAddress]::TryParse($addr, [ref]$parsed)) { return }
        if ($parsed.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { return }
        $addr = $parsed.IPAddressToString
        if ($seen.Contains($addr)) { return }
        [void]$seen.Add($addr)
        $prio = if ($addr -match '^169\.254\.') { 1 } else { 0 }
        [void]$scored.Add([pscustomobject]@{ Addr = $addr; Prio = $prio })
    }
    if ($sh.PSObject.Properties.Name -contains 'mpHostInfo' -and $sh.mpHostInfo) {
        $mh = $sh.mpHostInfo
        if ($mh.PSObject.Properties.Name -contains 'mpIpAddresses' -and $mh.mpIpAddresses) {
            foreach ($ip in @($mh.mpIpAddresses)) { foreach ($a in @(Expand-IloAddressValue -v $ip)) { & $addOne $a } }
        }
    }
    foreach ($f in @('mpDnsName', 'mpHostName', 'mpIpAddress')) {
        if ($sh.PSObject.Properties.Name -contains $f -and $sh.$f) { foreach ($a in @(Expand-IloAddressValue -v $sh.$f)) { & $addOne $a } }
    }
    $ordered = @($scored | Sort-Object Prio)
    if ($ordered.Count -gt 0) { return [string]$ordered[0].Addr }
    return $null
}

# ============================================================================
#  Haupt-Formular
# ============================================================================
$form = New-Object System.Windows.Forms.Form
$null = $form.Handle
$form.Text = if ($script:IsSynergyMode) {
    "© 2025 N.J. Airbus D&S - HPE Synergy Firmware Update (iLO Redfish)"
} else {
    "© 2025 N.J. Airbus D&S - HPE ProLiant Rackmount Firmware Update (Gen10+, iLO Redfish)"
}
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$wWidth  = [Math]::Min(1180, $screen.Width  - 40)
$wHeight = [Math]::Min(940,  $screen.Height - 40)
$form.Size = New-Object System.Drawing.Size($wWidth, $wHeight)
# Mindestgroesse klein genug fuer kleine Laptop-Displays; bei zu kleinem Fenster
# sorgt AutoScroll fuer Scrollleisten statt abgeschnittenem Inhalt.
$form.MinimumSize = New-Object System.Drawing.Size(820, 480)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
# Inhalt ist absolut positioniert (bis ~858 px hoch, ~1115 px breit). AutoScroll
# blendet automatisch Scrollleisten ein, wenn das Fenster kleiner als der Inhalt
# ist -> der untere Bereich bleibt auf kleinen Monitoren erreichbar.
$form.AutoScroll = $true
$form.AutoScrollMinSize = New-Object System.Drawing.Size(1130, 892)

$boldFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

# ─────────────────────────────────────────
# Credentials: iLO + OneView
# ─────────────────────────────────────────
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.AutoSize = $true; $lblUser.Location = '12,15'; $lblUser.Text = "iLO Benutzer:"; $lblUser.Font = $boldFont
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = '110,12'; $txtUser.Size = '150,22'; $txtUser.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.AutoSize = $true; $lblPass.Location = '270,15'; $lblPass.Text = "Pwd:"
$form.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = '315,12'; $txtPass.Size = '150,22'; $txtPass.UseSystemPasswordChar = $true; $txtPass.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtPass)

$lblOvUser = New-Object System.Windows.Forms.Label
$lblOvUser.AutoSize = $true; $lblOvUser.Location = '485,15'; $lblOvUser.Text = "OneView User:"; $lblOvUser.Font = $boldFont
$form.Controls.Add($lblOvUser)

$txtOvUser = New-Object System.Windows.Forms.TextBox
$txtOvUser.Location = '590,12'; $txtOvUser.Size = '150,22'; $txtOvUser.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtOvUser)

$lblOvPass = New-Object System.Windows.Forms.Label
$lblOvPass.AutoSize = $true; $lblOvPass.Location = '750,15'; $lblOvPass.Text = "Pwd:"
$form.Controls.Add($lblOvPass)

$txtOvPass = New-Object System.Windows.Forms.TextBox
$txtOvPass.Location = '795,12'; $txtOvPass.Size = '150,22'; $txtOvPass.UseSystemPasswordChar = $true; $txtOvPass.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtOvPass)

$tipCred = New-Object System.Windows.Forms.ToolTip
$tipCred.AutoPopDelay = 15000
$tipCred.SetToolTip($txtUser, "iLO-Account mit Update-Berechtigung (Configure iLO Settings).`r`nIdealerweise ein einheitlicher, ueber iLO-Federation ausgerollter Account fuer die ganze Flotte.")
$tipCred.SetToolTip($txtOvUser, "OneView-Account (nur Lesen noetig). Wird fuer den Import der iLO-Adressen, Modelle und Scopes der monitored Server verwendet.")

# ─────────────────────────────────────────
# Server-Quelle: OneView (Serverliste wird aus OneView erzeugt)
# ─────────────────────────────────────────
$grpSrc = New-Object System.Windows.Forms.GroupBox
$grpSrc.Location = '12,40'; $grpSrc.Size = '1103,62'; $grpSrc.Text = "Server-Quelle: OneView (Serverliste wird aus OneView erzeugt)"
$form.Controls.Add($grpSrc)

$lblOv = New-Object System.Windows.Forms.Label
$lblOv.AutoSize = $true; $lblOv.Location = '12,28'; $lblOv.Text = "Appliances:"; $lblOv.Font = $boldFont
$grpSrc.Controls.Add($lblOv)

$txtOvFile = New-Object System.Windows.Forms.TextBox
$applianceFileName = if ($script:IsSynergyMode) { 'Synergy.ini' } else { 'Oneview.ini' }
$defaultApplianceFile = if ([string]::IsNullOrWhiteSpace($ApplianceFile)) {
    Join-Path $scriptFolder $applianceFileName
} else {
    [System.IO.Path]::GetFullPath($ApplianceFile)
}
$txtOvFile.Location = '90,25'; $txtOvFile.Size = '300,22'; $txtOvFile.Text = $defaultApplianceFile; $txtOvFile.BorderStyle = 'FixedSingle'
$grpSrc.Controls.Add($txtOvFile)

$btnOvBrowse = New-Object System.Windows.Forms.Button
$btnOvBrowse.Location = '395,24'; $btnOvBrowse.Size = '40,24'; $btnOvBrowse.Text = "..."
$grpSrc.Controls.Add($btnOvBrowse)

$btnLoadOV = New-Object System.Windows.Forms.Button
$btnLoadOV.Location = '445,24'; $btnLoadOV.Size = '215,24'; $btnLoadOV.Text = "Server aus OneView laden"
$btnLoadOV.BackColor = [System.Drawing.Color]::FromArgb(60, 90, 160); $btnLoadOV.ForeColor = [System.Drawing.Color]::White
$btnLoadOV.FlatStyle = 'Flat'
$grpSrc.Controls.Add($btnLoadOV)

$btnLoadCache = New-Object System.Windows.Forms.Button
$serverCacheName = if ($script:IsSynergyMode) { 'SynergyServers.ini' } else { 'Servers.ini' }
$btnLoadCache.Location = '665,24'; $btnLoadCache.Size = '150,24'; $btnLoadCache.Text = "Aus $serverCacheName laden"
$grpSrc.Controls.Add($btnLoadCache)

$lblOvHint = New-Object System.Windows.Forms.Label
$lblOvHint.AutoSize = $true; $lblOvHint.Location = '825,28'; $lblOvHint.Text = "(1 Appliance/Zeile)"; $lblOvHint.ForeColor = [System.Drawing.Color]::Gray
$grpSrc.Controls.Add($lblOvHint)

$tipOv = New-Object System.Windows.Forms.ToolTip
$tipOv.AutoPopDelay = 15000
$applianceFileTip = if ($script:IsSynergyMode) {
    "Textdatei mit einer OneView-Appliance pro Zeile: Hostname/IP oder Hostname/IP;Anzeigename.`r`nDas Tool liest server-hardware + Scopes aus und erzeugt daraus die Serverliste ($serverCacheName)."
} else {
    "Textdatei mit einer OneView-Appliance (Hostname/IP) pro Zeile.`r`nDas Tool liest server-hardware + Scopes aus und erzeugt daraus die Serverliste ($serverCacheName)."
}
$tipOv.SetToolTip($txtOvFile, $applianceFileTip)
$tipOv.SetToolTip($btnLoadCache, "Laedt die zuletzt aus OneView erzeugte $serverCacheName (offline, inkl. Modell + Scope).")

# Modusgetrennte generierte/gecachte Liste (Output des OneView-Imports)
$serversFile = Join-Path $scriptFolder $serverCacheName

# ─────────────────────────────────────────
# Server-Auswahl (Filter: Text, Servertyp, Scope)
# ─────────────────────────────────────────
$lblSel = New-Object System.Windows.Forms.Label
$lblSel.AutoSize = $true; $lblSel.Location = '12,114'; $lblSel.Text = "Auswahl:"; $lblSel.Font = $boldFont
$form.Controls.Add($lblSel)

$btnSelAll = New-Object System.Windows.Forms.Button
$btnSelAll.Location = '80,110'; $btnSelAll.Size = '60,24'; $btnSelAll.Text = "Alle"
$form.Controls.Add($btnSelAll)

$btnSelNone = New-Object System.Windows.Forms.Button
$btnSelNone.Location = '143,110'; $btnSelNone.Size = '60,24'; $btnSelNone.Text = "Keine"
$form.Controls.Add($btnSelNone)

$lblType = New-Object System.Windows.Forms.Label
$lblType.AutoSize = $true; $lblType.Location = '220,114'; $lblType.Text = "Servertyp:"
$form.Controls.Add($lblType)

$cboType = New-Object System.Windows.Forms.ComboBox
$cboType.Location = '290,111'; $cboType.Size = '220,22'; $cboType.DropDownStyle = 'DropDownList'
$cboType.Items.Add("(alle)") | Out-Null; $cboType.SelectedIndex = 0
$form.Controls.Add($cboType)

$lblScope = New-Object System.Windows.Forms.Label
$lblScope.AutoSize = $true; $lblScope.Location = '525,114'; $lblScope.Text = "Scope:"
$form.Controls.Add($lblScope)

$cboScope = New-Object System.Windows.Forms.ComboBox
$cboScope.Location = '575,111'; $cboScope.Size = '180,22'; $cboScope.DropDownStyle = 'DropDownList'
$cboScope.Items.Add("(alle)") | Out-Null; $cboScope.SelectedIndex = 0
$form.Controls.Add($cboScope)

$lblAppliance = New-Object System.Windows.Forms.Label
$lblAppliance.AutoSize = $true; $lblAppliance.Location = '625,114'; $lblAppliance.Text = "Appliance:"
$lblAppliance.Visible = $script:IsSynergyMode
$form.Controls.Add($lblAppliance)

$cboAppliance = New-Object System.Windows.Forms.ComboBox
$cboAppliance.Location = '695,111'; $cboAppliance.Size = '150,22'; $cboAppliance.DropDownStyle = 'DropDownList'
$cboAppliance.Items.Add("(alle)") | Out-Null; $cboAppliance.SelectedIndex = 0
$cboAppliance.Visible = $script:IsSynergyMode
$form.Controls.Add($cboAppliance)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.AutoSize = $true; $lblFilter.Location = '770,114'; $lblFilter.Text = "Suche:"
$form.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = '815,111'; $txtFilter.Size = '180,22'; $txtFilter.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtFilter)
$tipFilter = New-Object System.Windows.Forms.ToolTip
$tipFilter.SetToolTip($txtFilter, "Filter auf iLO-Adresse ODER Servername. Wildcards * und ? moeglich (z.B. esxprod*, *db?? ).")

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.AutoSize = $true; $lblCount.Location = '1005,114'; $lblCount.Text = ""; $lblCount.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblCount)

if ($script:IsSynergyMode) {
    $cboType.Location = '280,111'; $cboType.Size = '150,22'
    $lblScope.Location = '440,114'
    $cboScope.Location = '485,111'; $cboScope.Size = '130,22'
    $lblFilter.Location = '855,114'
    $txtFilter.Location = '900,111'; $txtFilter.Size = '100,22'
}

$gridServers = New-Object System.Windows.Forms.DataGridView
$gridServers.Location = '12,140'; $gridServers.Size = '1103,146'
$gridServers.AllowUserToAddRows = $false; $gridServers.AllowUserToDeleteRows = $false
$gridServers.AllowUserToResizeRows = $false; $gridServers.RowHeadersVisible = $false
$gridServers.MultiSelect = $false; $gridServers.SelectionMode = 'FullRowSelect'
$gridServers.AutoGenerateColumns = $false; $gridServers.BackgroundColor = [System.Drawing.Color]::White
$gridServers.BorderStyle = 'FixedSingle'; $gridServers.EditMode = 'EditOnEnter'

$colSelected = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSelected.Name = 'Selected'; $colSelected.HeaderText = 'Auswahl'; $colSelected.Width = 58
$colSelected.SortMode = 'NotSortable'
[void]$gridServers.Columns.Add($colSelected)
$serverColumns = if ($script:IsSynergyMode) {
    @(
        @{ Name = 'Ilo';       Header = 'iLO / IP-Adresse'; Width = 140 },
        @{ Name = 'Name';      Header = 'Servername';       Width = 180 },
        @{ Name = 'Appliance'; Header = 'OneView Appliance'; Width = 165 },
        @{ Name = 'Model';     Header = 'Servertyp';        Width = 280 },
        @{ Name = 'Scope';     Header = 'Scope';            Width = 165 }
    )
} else {
    @(
        @{ Name = 'Ilo';   Header = 'iLO / IP-Adresse'; Width = 145 },
        @{ Name = 'Name';  Header = 'Servername';       Width = 255 },
        @{ Name = 'Model'; Header = 'Servertyp';        Width = 285 },
        @{ Name = 'Scope'; Header = 'Scope';            Width = 245 }
    )
}
foreach ($column in $serverColumns) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $column.Name; $col.HeaderText = $column.Header; $col.Width = $column.Width
    $col.ReadOnly = $true
    [void]$gridServers.Columns.Add($col)
}
$colIloLink = New-Object System.Windows.Forms.DataGridViewLinkColumn
$colIloLink.Name = 'IloLink'; $colIloLink.HeaderText = 'iLO Link'; $colIloLink.Width = 85
$colIloLink.Text = 'Oeffnen'; $colIloLink.UseColumnTextForLinkValue = $true
$colIloLink.TrackVisitedState = $false; $colIloLink.ReadOnly = $true
[void]$gridServers.Columns.Add($colIloLink)
$form.Controls.Add($gridServers)

# Voller (ungefilterter) Bestand: Liste von Objekten
# @{ Raw='host[;user;pass]'; Name; Appliance; Model; Gen; Scope }
$script:allServers = @()

function Test-ServerEligibleForSelection {
    param($Server)

    if (-not $Server) { return $false }
    $isSynergy = "$($Server.Model)" -match '(?i)\bSynergy\b'
    if ($script:IsSynergyMode) { return $isSynergy }
    if ($isSynergy) { return $false }

    $gen = 0
    if ($Server.PSObject.Properties.Name -contains 'Gen') {
        try { $gen = [int]$Server.Gen } catch { $gen = 0 }
    }
    if ($gen -le 0 -and "$($Server.Model)" -match '(?i)\bGen\s*(\d+)\b') {
        $gen = [int]$Matches[1]
    }
    return ($gen -le 0 -or $gen -ge 10)
}

function Resolve-SynergyFirmwareFolder {
    param([string]$BaseDir)

    if (-not (Test-Path -LiteralPath $BaseDir -PathType Container)) {
        throw "Firmware-Basisverzeichnis nicht gefunden: $BaseDir"
    }
    $folder = Join-Path $BaseDir 'synergy'
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        throw "Synergy-Firmwareordner nicht gefunden: $folder"
    }
    return $folder
}

# Schreibt den aktuellen Bestand als reloadbare Serverliste.
function Save-ServersFile {
    param([string]$Path)
    $lines = @(
        "# Automatisch aus OneView erzeugt - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "# Format: ilo[;user;pass]    # Name=...; Appliance=...; Modell=...; Gen=NN; Scope=...",
        "# Diese Datei kann ueber die Cache-Ladefunktion offline wieder eingelesen werden."
    )
    foreach ($obj in ($script:allServers | Sort-Object Raw)) {
        $meta = "# Name=$($obj.Name); Appliance=$($obj.Appliance); Modell=$($obj.Model); Gen=$($obj.Gen); Scope=$($obj.Scope)"
        $lines += ("{0}    {1}" -f $obj.Raw, $meta)
    }
    try { Set-Content -LiteralPath $Path -Value $lines -Encoding UTF8 } catch { Add-Log "Serverliste '$Path' konnte nicht geschrieben werden: $($_.Exception.Message)" ([System.Drawing.Color]::Red) }
}

# Laedt eine zuvor erzeugte Serverliste (inkl. Metadaten) zurueck in den Bestand.
function Load-ServersFile {
    param([string]$Path)
    $script:allServers = @()
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        [System.Windows.Forms.MessageBox]::Show("Serverliste nicht gefunden:`n$Path", "Hinweis", 'OK', 'Warning'); return
    }
    foreach ($line in Get-Content -LiteralPath $Path) {
        $t = $line.Trim()
        if ($t -eq '' -or $t.StartsWith('#')) { continue }
        $name = ''; $appliance = ''; $model = ''; $gen = 0; $scope = ''
        $hostPart = $t
        $hashIdx = $t.IndexOf('#')
        if ($hashIdx -ge 0) {
            $hostPart = $t.Substring(0, $hashIdx).Trim()
            $meta = $t.Substring($hashIdx + 1)
            if ($meta -match '(?i)Name=([^;]*)')   { $name = $Matches[1].Trim() }
            if ($meta -match '(?i)Appliance=([^;]*)') { $appliance = $Matches[1].Trim() }
            if ($meta -match '(?i)Modell=([^;]*)') { $model = $Matches[1].Trim() }
            if ($meta -match '(?i)Gen=(\d+)')      { $gen = [int]$Matches[1] }
            if ($meta -match '(?i)Scope=(.*)$')    { $scope = $Matches[1].Trim() }
        }
        if ($hostPart -eq '') { continue }
        $script:allServers += [pscustomobject]@{ Raw = $hostPart; Name = $name; Appliance = $appliance; Model = $model; Gen = $gen; Scope = $scope }
    }
    Rebuild-FilterChoices
    Apply-ServerFilter
    Add-Log "Aus '$([System.IO.Path]::GetFileName($Path))' geladen: $($script:allServers.Count) Server." ([System.Drawing.Color]::DarkBlue)
}

# Befuellt die Servertyp- und Scope-Dropdowns aus dem aktuellen Bestand.
function Rebuild-FilterChoices {
    $selType = if ($cboType.SelectedItem) { "$($cboType.SelectedItem)" } else { "(alle)" }
    $selScope = if ($cboScope.SelectedItem) { "$($cboScope.SelectedItem)" } else { "(alle)" }
    $selAppliance = if ($cboAppliance.SelectedItem) { "$($cboAppliance.SelectedItem)" } else { "(alle)" }

    $selectableServers = @($script:allServers | Where-Object { Test-ServerEligibleForSelection $_ })
    $models = @($selectableServers | Where-Object { $_.Model } | ForEach-Object { $_.Model } | Sort-Object -Unique)
    $scopes = @($selectableServers | Where-Object { $_.Scope } | ForEach-Object { $_.Scope -split '\s*,\s*' } | Where-Object { $_ } | Sort-Object -Unique)
    $appliances = @($selectableServers | Where-Object { $_.Appliance } | ForEach-Object { $_.Appliance } | Sort-Object -Unique)

    $cboType.BeginUpdate(); $cboType.Items.Clear(); $cboType.Items.Add("(alle)") | Out-Null
    foreach ($m in $models) { $cboType.Items.Add($m) | Out-Null }
    $i = $cboType.Items.IndexOf($selType); $cboType.SelectedIndex = if ($i -ge 0) { $i } else { 0 }
    $cboType.EndUpdate()

    $cboScope.BeginUpdate(); $cboScope.Items.Clear(); $cboScope.Items.Add("(alle)") | Out-Null
    foreach ($s in $scopes) { $cboScope.Items.Add($s) | Out-Null }
    $i = $cboScope.Items.IndexOf($selScope); $cboScope.SelectedIndex = if ($i -ge 0) { $i } else { 0 }
    $cboScope.EndUpdate()

    $cboAppliance.BeginUpdate(); $cboAppliance.Items.Clear(); $cboAppliance.Items.Add("(alle)") | Out-Null
    foreach ($a in $appliances) { $cboAppliance.Items.Add($a) | Out-Null }
    $i = $cboAppliance.Items.IndexOf($selAppliance); $cboAppliance.SelectedIndex = if ($i -ge 0) { $i } else { 0 }
    $cboAppliance.EndUpdate()
}

function Apply-ServerFilter {
    $flt = $txtFilter.Text.Trim()
    $selType = if ($cboType.SelectedItem) { "$($cboType.SelectedItem)" } else { "(alle)" }
    $selScope = if ($cboScope.SelectedItem) { "$($cboScope.SelectedItem)" } else { "(alle)" }
    $selAppliance = if ($cboAppliance.SelectedItem) { "$($cboAppliance.SelectedItem)" } else { "(alle)" }
    $gridServers.SuspendLayout()
    $gridServers.Rows.Clear()
    foreach ($obj in $script:allServers) {
        if (-not (Test-ServerEligibleForSelection $obj)) { continue }
        $iloHost = ($obj.Raw -split ';')[0].Trim()
        $srvName = "$($obj.Name)".Trim()
        $show = $true
        if (-not [string]::IsNullOrWhiteSpace($flt)) {
            # Filter auf iLO-Adresse ODER Servername. * und ? sind Platzhalter.
            # Damit sich die Suche wie ohne Platzhalter (Teilstring/"enthaelt")
            # verhaelt, wird das Muster mit * umschlossen, sofern es nicht bereits
            # mit * beginnt/endet. So findet z.B. 'prod*' weiterhin 'esxprod01'.
            if ($flt -match '[\*\?]') {
                $pat = $flt
                if ($pat -notlike '`**') { $pat = "*$pat" }
                if ($pat -notlike '*`*') { $pat = "$pat*" }
                $show = ($iloHost -like $pat) -or ($srvName -like $pat)
            } else {
                $show = ($iloHost -like "*$flt*") -or ($srvName -like "*$flt*")
            }
        }
        if ($show -and $selType -ne '(alle)') { $show = ($obj.Model -eq $selType) }
        if ($show -and $selScope -ne '(alle)') {
            $sList = @($obj.Scope -split '\s*,\s*')
            $show = ($sList -contains $selScope)
        }
        if ($show -and $script:IsSynergyMode -and $selAppliance -ne '(alle)') {
            $show = ("$($obj.Appliance)" -eq $selAppliance)
        }
        if ($show) {
            $rowIndex = $gridServers.Rows.Add()
            $row = $gridServers.Rows[$rowIndex]
            $row.Tag = $obj
            $row.Cells.Item('Selected').Value = $false
            $row.Cells.Item('Ilo').Value = $iloHost
            $row.Cells.Item('Name').Value = $srvName
            if ($script:IsSynergyMode) { $row.Cells.Item('Appliance').Value = "$($obj.Appliance)" }
            $row.Cells.Item('Model').Value = "$($obj.Model)"
            $row.Cells.Item('Scope').Value = "$($obj.Scope)"
        }
    }
    $gridServers.ResumeLayout()
    $selectableCount = @($script:allServers | Where-Object { Test-ServerEligibleForSelection $_ }).Count
    $lblCount.Text = "$($gridServers.Rows.Count) angezeigt / $selectableCount auswaehlbar"
}

# Liefert ausgewaehlte Server als @{ Ilo; User; Pass; Name }.
function Get-CheckedServers {
    $result = New-Object System.Collections.ArrayList
    $seen = @{}
    # EndEdit() liefert einen Boolean. Ohne Unterdrueckung wird dieser von
    # PowerShell als zusaetzlicher "Server" aus der Funktion ausgegeben.
    [void]$gridServers.EndEdit()
    for ($rowIndex = 0; $rowIndex -lt $gridServers.Rows.Count; $rowIndex++) {
        $row = $gridServers.Rows.Item($rowIndex)
        if ([bool]$row.Cells.Item(0).Value) {
            $u = $txtUser.Text; $pw = $txtPass.Text
            $obj = $row.Tag
            # Die sichtbare iLO-Spalte ist pro Tabellenzeile ein einzelner Wert.
            # Numerische Indexer vermeiden PowerShells Member-Enumeration bei
            # benannten DataGridView-Collections.
            $ilo = [string]$row.Cells.Item(1).Value
            $ilo = $ilo.Trim()
            if ($obj) {
                if (-not (Test-ServerEligibleForSelection $obj)) {
                    [void](Add-Log "Server '$ilo' mit Generation Gen$($obj.Gen) ist nicht auswaehlbar (unterstuetzt: Gen10+)." ([System.Drawing.Color]::DarkOrange))
                    continue
                }
                $parts = @("$($obj.Raw)" -split ';')
                if ($parts.Count -ge 2 -and $parts[1].Trim() -ne '') { $u = $parts[1].Trim() }
                if ($parts.Count -ge 3 -and $parts[2].Trim() -ne '') { $pw = $parts[2].Trim() }
            }
            if (-not $ilo -or $ilo -match '\s') {
                [void](Add-Log "Ausgewaehlte Serverzeile mit ungueltiger iLO-Adresse '$ilo' wurde uebersprungen." ([System.Drawing.Color]::DarkOrange))
                continue
            }
            $key = $ilo.ToLowerInvariant()
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            [void]$result.Add([pscustomobject]@{
                Ilo  = $ilo
                User = $u
                Pass = $pw
                Name = if ($obj) { "$($obj.Name)" } else { '' }
            })
        }
    }
    $selected = @($result.ToArray())
    if ($selected.Count) {
        [void](Add-Log "Serverauswahl: $($selected.Count) einzelne iLO-Ziele: $((@($selected | ForEach-Object { $_.Ilo })) -join ', ')" ([System.Drawing.Color]::DarkBlue))
    }
    return $selected
}

# Importiert iLO-Adressen + Modelle + Scopes der monitored Server aus OneView.
function Load-FromOneView {
    if ([string]::IsNullOrWhiteSpace($txtOvUser.Text) -or [string]::IsNullOrWhiteSpace($txtOvPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte OneView-Benutzer und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    $ovFile = $txtOvFile.Text
    if ([string]::IsNullOrWhiteSpace($ovFile) -or -not (Test-Path -LiteralPath $ovFile)) {
        [System.Windows.Forms.MessageBox]::Show("OneView-Appliance-Datei nicht gefunden:`n$ovFile", "Fehler", 'OK', 'Error'); return
    }
    $appliances = @()
    foreach ($line in Get-Content -LiteralPath $ovFile) {
        $t = $line.Trim(); if ($t -eq '' -or $t.StartsWith('#')) { continue }
        $parts = @($t -split ';', 2)
        $address = "$($parts[0])".Trim()
        if (-not $address) { continue }
        $displayName = if ($parts.Count -gt 1 -and -not [string]::IsNullOrWhiteSpace($parts[1])) {
            "$($parts[1])".Trim()
        } else {
            $address
        }
        $appliances += [pscustomobject]@{ Address = $address; Name = $displayName }
    }
    if ($appliances.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Keine Appliances in der Datei.", "Hinweis", 'OK', 'Warning'); return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnLoadOV.Enabled = $false; $btnLoadCache.Enabled = $false; $btnInventory.Enabled = $false; $btnFlash.Enabled = $false
    $collected = @{}
    try {
        foreach ($applianceEntry in $appliances) {
            $appl = $applianceEntry.Address
            $applianceName = $applianceEntry.Name
            Add-Log "OneView $applianceName ($appl) : Abfrage server-hardware + Scopes..." ([System.Drawing.Color]::DarkBlue)
            [System.Windows.Forms.Application]::DoEvents()
            $sess = $null; $apiV = $null
            try {
                $apiV = OV-GetApiVersion -A $appl
                $sess = OV-Login -A $appl -U $txtOvUser.Text -P $txtOvPass.Text -V $apiV
                $scopeMap = OV-GetScopeMap -A $appl -S $sess -V $apiV
                $srvScopes = OV-GetServerScopeUris -A $appl -S $sess -V $apiV
                $members = OV-RestAll -A $appl -S $sess -V $apiV -E "/rest/server-hardware?start=0&count=1000"
                $added = 0; $skipped = 0
                foreach ($m in $members) {
                    $model = "$($m.model)"
                    $gen = 0; if ($model -match '(?i)Gen\s*(\d+)') { $gen = [int]$Matches[1] }
                    $ilo = Get-IloIpFromHardware -sh $m
                    if ([string]::IsNullOrWhiteSpace($ilo)) { $skipped++; continue }
                    # Scope-Namen ermitteln
                    $scopeNames = @()
                    $uris = @()
                    if ($m.PSObject.Properties.Name -contains 'scopeUris' -and $m.scopeUris) { $uris = @($m.scopeUris) }
                    elseif ($m.uri -and $srvScopes.ContainsKey($m.uri)) { $uris = @($srvScopes[$m.uri]) }
                    foreach ($u in $uris) { if ($scopeMap.ContainsKey($u)) { $scopeNames += $scopeMap[$u] } elseif ($u) { $scopeNames += ($u -replace '.*/', '') } }
                    $scope = (@($scopeNames | Sort-Object -Unique) -join ', ')
                    # Servername ermitteln (bevorzugt serverName, sonst Hardware-Name)
                    $srvName = ''
                    if ($m.PSObject.Properties.Name -contains 'serverName' -and -not [string]::IsNullOrWhiteSpace("$($m.serverName)")) { $srvName = "$($m.serverName)".Trim() }
                    elseif ($m.PSObject.Properties.Name -contains 'name' -and -not [string]::IsNullOrWhiteSpace("$($m.name)")) { $srvName = "$($m.name)".Trim() }
                    $candidate = [pscustomobject]@{ Raw = $ilo; Name = $srvName; Appliance = $applianceName; Model = $model; Gen = $gen; Scope = $scope }
                    if (-not (Test-ServerEligibleForSelection $candidate)) { $skipped++; continue }
                    $collected[$ilo] = $candidate
                    $added++
                }
                Add-Log "OneView $appl : $added uebernommen, $skipped uebersprungen." ([System.Drawing.Color]::DarkGreen)
            }
            catch { Add-Log "OneView $appl : FEHLER - $($_.Exception.Message)" ([System.Drawing.Color]::Red) }
            finally { if ($sess -and $apiV) { try { OV-Logout -A $appl -S $sess -V $apiV } catch {} } }
            [System.Windows.Forms.Application]::DoEvents()
        }
        $script:allServers = @($collected.Values | Sort-Object Raw)
        Save-ServersFile -Path $serversFile
        Rebuild-FilterChoices
        Apply-ServerFilter
        Add-Log "OneView-Import fertig: $($script:allServers.Count) Server -> $serversFile gespeichert." ([System.Drawing.Color]::DarkBlue)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnLoadOV.Enabled = $true; $btnLoadCache.Enabled = $true; $btnInventory.Enabled = $true; $btnFlash.Enabled = $true
    }
}

$btnSelAll.Add_Click({ foreach ($row in $gridServers.Rows) { $row.Cells.Item('Selected').Value = $true } })
$btnSelNone.Add_Click({ foreach ($row in $gridServers.Rows) { $row.Cells.Item('Selected').Value = $false } })
$gridServers.Add_CurrentCellDirtyStateChanged({
    if ($gridServers.IsCurrentCellDirty) {
        [void]$gridServers.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})
$gridServers.Add_CellContentClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
    if ($gridServers.Columns[$e.ColumnIndex].Name -ne 'IloLink') { return }
    $ilo = "$($gridServers.Rows[$e.RowIndex].Cells.Item('Ilo').Value)".Trim()
    if (-not $ilo) { return }
    try {
        $uri = [System.UriBuilder]::new('https', $ilo).Uri.AbsoluteUri
        Start-Process -FilePath $uri
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Die iLO-Webseite konnte nicht geoeffnet werden:`n$ilo`n`n$($_.Exception.Message)",
            "iLO Link", 'OK', 'Error') | Out-Null
    }
})
$txtFilter.Add_TextChanged({ Apply-ServerFilter })
$cboType.Add_SelectedIndexChanged({ Apply-ServerFilter })
$cboScope.Add_SelectedIndexChanged({ Apply-ServerFilter })
$cboAppliance.Add_SelectedIndexChanged({ Apply-ServerFilter })
$btnOvBrowse.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "Textdateien (*.ini)|*.ini|Alle (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') { $txtOvFile.Text = $ofd.FileName }
})
$btnLoadOV.Add_Click({ Load-FromOneView })
$btnLoadCache.Add_Click({ Load-ServersFile -Path $serversFile })


# ─────────────────────────────────────────
# Firmware-Auswahl
# ─────────────────────────────────────────
$grpFw = New-Object System.Windows.Forms.GroupBox
$grpFw.Location = '12,292'; $grpFw.Size = '1103,116'; $grpFw.Text = "Firmware (.fwpkg)"
$form.Controls.Add($grpFw)

$rbFile = New-Object System.Windows.Forms.RadioButton
$rbFile.Location = '12,22'; $rbFile.Size = '160,22'; $rbFile.Text = "Einzeldatei (.fwpkg)"
$grpFw.Controls.Add($rbFile)

$rbFolder = New-Object System.Windows.Forms.RadioButton
$rbFolder.Location = '180,22'; $rbFolder.Size = '300,22'; $rbFolder.Text = "Ordner (mehrere .fwpkg pro Server)"
$grpFw.Controls.Add($rbFolder)

$rbBaseDir = New-Object System.Windows.Forms.RadioButton
$rbBaseDir.Location = '490,22'; $rbBaseDir.Size = '595,22'; $rbBaseDir.Text = "Firmware-Verzeichnis mit Typ-Unterordnern (autom. Zuordnung je Server)"; $rbBaseDir.Checked = $true
$grpFw.Controls.Add($rbBaseDir)

$lblPath = New-Object System.Windows.Forms.Label
$lblPath.AutoSize = $true; $lblPath.Location = '12,53'; $lblPath.Text = "Pfad:"; $lblPath.Font = $boldFont
$grpFw.Controls.Add($lblPath)

$txtFw = New-Object System.Windows.Forms.TextBox
$txtFw.Location = '60,50'; $txtFw.Size = '855,22'; $txtFw.BorderStyle = 'FixedSingle'; $txtFw.Text = (Join-Path $scriptFolder "Firmware")
$grpFw.Controls.Add($txtFw)

$btnBrowseFw = New-Object System.Windows.Forms.Button
$btnBrowseFw.Location = '925,49'; $btnBrowseFw.Size = '90,24'; $btnBrowseFw.Text = "Browse..."
$grpFw.Controls.Add($btnBrowseFw)

$lblSha = New-Object System.Windows.Forms.Label
$lblSha.AutoSize = $true; $lblSha.Location = '12,84'; $lblSha.Text = "SHA-256 (optional):"; $lblSha.Enabled = $false
$grpFw.Controls.Add($lblSha)

$txtSha = New-Object System.Windows.Forms.TextBox
$txtSha.Location = '140,81'; $txtSha.Size = '360,22'; $txtSha.BorderStyle = 'FixedSingle'; $txtSha.Enabled = $false
$grpFw.Controls.Add($txtSha)
$tipSha = New-Object System.Windows.Forms.ToolTip
$tipSha.AutoPopDelay = 15000
$tipSha.SetToolTip($txtSha, "Optional: erwarteter SHA-256 der Einzeldatei. Wird vor dem Upload lokal geprueft.`r`n(Nur bei 'Einzeldatei' aktiv.)")

$lblFwHint = New-Object System.Windows.Forms.Label
$lblFwHint.AutoSize = $false; $lblFwHint.Location = '515,78'; $lblFwHint.Size = '576,32'; $lblFwHint.ForeColor = [System.Drawing.Color]::Gray
$lblFwHint.Text = "Standard: Basisverzeichnis mit gen10, gen11 und gen12. Modellordner werden weiterhin unterstuetzt. Dateinamen NICHT aendern."
$grpFw.Controls.Add($lblFwHint)

$tipFw = New-Object System.Windows.Forms.ToolTip
$tipFw.AutoPopDelay = 20000
$tipFw.SetToolTip($rbBaseDir, "Firmware-Verzeichnis mit einem Unterordner je Generation (Standard):`r`n  Firmware\gen10\*.fwpkg`r`n  Firmware\gen11\*.fwpkg`r`n  Firmware\gen12\*.fwpkg`r`nDer passende Generationsordner wird am iLO ermittelt. Nicht verbaute Komponenten werden anhand des Firmware-Inventars uebersprungen.`r`nBestehende Modellordner (z.B. 'ProLiant DL380a Gen11') werden weiterhin als Fallback unterstuetzt.")
if ($script:IsSynergyMode) {
    $grpFw.Text = 'Synergy Firmware (.fwpkg)'
    $rbFile.Visible = $false
    $rbFolder.Visible = $false
    $rbBaseDir.Location = '12,22'
    $rbBaseDir.Size = '1075,22'
    $rbBaseDir.Text = "Firmware-Basisverzeichnis (Pakete aus Unterordner 'synergy')"
    $lblFwHint.Text = "Alle Synergy-Pakete gemeinsam unter Firmware\synergy ablegen. Die Zuordnung erfolgt ueber Paket-JSON und Live-Inventar."
    $tipFw.SetToolTip($rbBaseDir, "Basisverzeichnis auswaehlen. Verwendet werden ausschliesslich die .fwpkg-Dateien direkt unter:`r`n  Firmware\synergy\*.fwpkg")
}

$btnBrowseFw.Add_Click({
    if ($rbFile.Checked) {
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "HPE Firmware (*.fwpkg)|*.fwpkg|Alle (*.*)|*.*"
        if ($ofd.ShowDialog() -eq 'OK') { $txtFw.Text = $ofd.FileName }
    } else {
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = if ($script:IsSynergyMode) {
            "Firmware-Basisverzeichnis auswaehlen; Pakete werden aus dem Unterordner 'synergy' gelesen"
        } elseif ($rbBaseDir.Checked) {
            "Basisverzeichnis mit gen10/gen11/gen12-Unterordnern auswaehlen"
        } else {
            "Ordner mit .fwpkg-Dateien auswaehlen"
        }
        if ($fbd.ShowDialog() -eq 'OK') { $txtFw.Text = $fbd.SelectedPath }
    }
})
$rbFolder.Add_CheckedChanged({ if ($rbFolder.Checked) { $txtSha.Enabled = $false; $lblSha.Enabled = $false } })
$rbBaseDir.Add_CheckedChanged({ if ($rbBaseDir.Checked) { $txtSha.Enabled = $false; $lblSha.Enabled = $false } })
$rbFile.Add_CheckedChanged({ if ($rbFile.Checked) { $txtSha.Enabled = $true; $lblSha.Enabled = $true } })

# ─────────────────────────────────────────
# Aktionen + Parallelitaet
# ─────────────────────────────────────────
$lblPar = New-Object System.Windows.Forms.Label
$lblPar.AutoSize = $true; $lblPar.Location = '12,422'; $lblPar.Text = "Parallel:"; $lblPar.Font = $boldFont
$form.Controls.Add($lblPar)

$numPar = New-Object System.Windows.Forms.NumericUpDown
$numPar.Location = '75,419'; $numPar.Size = '55,22'; $numPar.Minimum = 1; $numPar.Maximum = 50; $numPar.Value = 8
$form.Controls.Add($numPar)
$tipPar = New-Object System.Windows.Forms.ToolTip
$tipPar.SetToolTip($numPar, "Anzahl gleichzeitig verarbeiteter Server. 8 ist ein sicherer Standard.")

$btnInventory = New-Object System.Windows.Forms.Button
$btnInventory.Location = '150,414'; $btnInventory.Size = '180,34'; $btnInventory.Text = "Inventar pruefen (read-only)"
$btnInventory.BackColor = [System.Drawing.Color]::FromArgb(60, 90, 160); $btnInventory.ForeColor = [System.Drawing.Color]::White
$btnInventory.FlatStyle = 'Flat'; $btnInventory.Font = $boldFont
$btnInventory.TextAlign = 'MiddleCenter'; $btnInventory.AutoEllipsis = $false; $btnInventory.Padding = '0,0,0,0'
$form.Controls.Add($btnInventory)

$btnFlash = New-Object System.Windows.Forms.Button
$btnFlash.Location = '340,414'; $btnFlash.Size = '180,34'; $btnFlash.Text = "Firmware aktualisieren"
$btnFlash.BackColor = [System.Drawing.Color]::FromArgb(40, 100, 60); $btnFlash.ForeColor = [System.Drawing.Color]::White
$btnFlash.FlatStyle = 'Flat'; $btnFlash.Font = $boldFont
$btnFlash.TextAlign = 'MiddleCenter'; $btnFlash.AutoEllipsis = $false; $btnFlash.Padding = '0,0,0,0'
$form.Controls.Add($btnFlash)

# Komponentenfilter: steuert, welche Firmware-Typen geflasht werden.
#   'Alle Komponenten'       -> kein Filter (Standardverhalten)
#   'Nur iLO'                -> ausschliesslich iLO-Firmware
#   'Nur iLO + BIOS + SPS'   -> iLO, System-ROM/BIOS und SPS/ME; Rest ignoriert
$cmbFwFilter = New-Object System.Windows.Forms.ComboBox
$cmbFwFilter.Location = '528,420'; $cmbFwFilter.Size = '172,22'; $cmbFwFilter.DropDownStyle = 'DropDownList'
[void]$cmbFwFilter.Items.Add('Alle Komponenten')
[void]$cmbFwFilter.Items.Add('Nur iLO')
[void]$cmbFwFilter.Items.Add('Nur iLO + BIOS + SPS')
$cmbFwFilter.SelectedIndex = 0
$form.Controls.Add($cmbFwFilter)
# Liefert den aktiven Kind-Filter als Array von Get-ComponentKind-Kinds.
# Leeres Array = alle Komponenten.
$script:GetKindFilter = {
    switch ($cmbFwFilter.SelectedIndex) {
        1 { @('iLO') }
        2 { @('iLO','ROM','SPS') }
        default { @() }
    }
}.GetNewClosure()
$tipIlo = New-Object System.Windows.Forms.ToolTip
$tipIlo.AutoPopDelay = 15000
$fwFilterTooltip = if ($script:IsSynergyMode) {
    "Alle Komponenten: alle eindeutig passenden Pakete aus Firmware\synergy.`r`nOptional kann auf iLO oder iLO + BIOS + SPS eingeschraenkt werden.`r`nNur eindeutig neuere, per JSON und Live-Inventar zuordenbare Versionen werden verarbeitet."
} else {
    "Alle Komponenten: alle .fwpkg der Quelle werden verarbeitet (nur Gen10+).`r`nNur iLO: ausschliesslich iLO-Firmware (nur Gen10+).`r`nNur iLO + BIOS + SPS: nur iLO-, System-ROM/BIOS- und SPS/ME-Pakete; andere Komponenten (CPLD/IE/PMC/...) werden ignoriert (nur Gen10+).`r`nBereits aktuelle Versionen werden immer uebersprungen. HPE Synergy bleibt blockiert."
}
$tipIlo.SetToolTip($cmbFwFilter, $fwFilterTooltip)

$chkInventoryReport = New-Object System.Windows.Forms.CheckBox
$chkInventoryReport.Location = '150,451'; $chkInventoryReport.Size = '250,22'
$chkInventoryReport.Text = "HTML-Bericht nach Inventarpruefung"
$form.Controls.Add($chkInventoryReport)
$tipReport = New-Object System.Windows.Forms.ToolTip
$tipReport.AutoPopDelay = 15000
$tipReport.SetToolTip($chkInventoryReport, "Erstellt nach Abschluss einen zusammengefassten, filterbaren HTML-Bericht im Unterordner 'Reports'. Es werden keine zusaetzlichen iLO-Abfragen ausgefuehrt.")

$chkDiagnosticReport = New-Object System.Windows.Forms.CheckBox
$chkDiagnosticReport.Location = '410,451'; $chkDiagnosticReport.Size = '300,22'
$chkDiagnosticReport.Text = "Technische Diagnoseexporte erstellen"
$chkDiagnosticReport.Checked = $false
$form.Controls.Add($chkDiagnosticReport)
$tipDiagnostic = New-Object System.Windows.Forms.ToolTip
$tipDiagnostic.AutoPopDelay = 15000
$tipDiagnostic.SetToolTip($chkDiagnosticReport, "Optional fuer Fehlersuche: schreibt pro Server einen detaillierten JSON-Diagnoseexport. Fuehrt zusaetzliche read-only Redfish-Abfragen fuer Device- und Storage-Inventar aus.")

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = '705,418'; $progress.Size = '410,26'; $progress.Minimum = 0; $progress.Maximum = 100
$form.Controls.Add($progress)

# ─────────────────────────────────────────
# Ergebnis-Liste
# ─────────────────────────────────────────
$lv = New-Object System.Windows.Forms.ListView
$lv.Location = '12,478'; $lv.Size = '1103,222'; $lv.View = 'Details'; $lv.FullRowSelect = $true; $lv.GridLines = $true
$lv.Scrollable = $true   # horizontale Scrollleiste, sobald Spalten breiter als Steuerelement sind
$lv.Columns.Add("Server (iLO)", 145) | Out-Null
$lv.Columns.Add("Modell", 170) | Out-Null
$lv.Columns.Add("Phase", 260) | Out-Null
$lv.Columns.Add("Fortschritt", 75) | Out-Null
$lv.Columns.Add("Ergebnis", 70) | Out-Null
# Details bewusst breit, damit der vollstaendige Text (z.B. Fehlermeldungen) lesbar
# ist; Summe der Spaltenbreiten > Steuerelementbreite -> horizontaler Scrollbalken.
$lv.Columns.Add("Details", 900) | Out-Null
$form.Controls.Add($lv)

# ─────────────────────────────────────────
# Log + Status
# ─────────────────────────────────────────
$panelLog = New-Object System.Windows.Forms.Panel
$panelLog.Location = '12,708'; $panelLog.Size = '1103,150'; $panelLog.BorderStyle = 'FixedSingle'
$form.Controls.Add($panelLog)

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Dock = 'Fill'; $logBox.ReadOnly = $true; $logBox.BorderStyle = 'None'
$logBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$panelLog.Controls.Add($logBox)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.Dock = 'Bottom'
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel; $statusLabel.Text = "Bereit"
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

# ── Anpassung an kleine (Laptop-)Bildschirme ─────────────────────────────────
# Die GUI ist fuer eine feste Groesse entworfen. Auf niedrigen Aufloesungen war
# der untere Teil (Tabelle, Log, Statusleiste) nicht sichtbar. Daher:
#  - Fenster groessenveraenderbar + maximierbar
#  - Bildlauf (AutoScroll) ueber die volle entworfene Inhaltsflaeche
#  - Startgroesse auf den nutzbaren Bildschirmbereich begrenzen
$designClient = $form.ClientSize
$form.FormBorderStyle = 'Sizable'
$form.MaximizeBox = $true
$form.MinimumSize = New-Object System.Drawing.Size(700, 500)
$form.AutoScroll = $true
# Scrollbereich auf die urspruenglich entworfene Inhaltsgroesse fixieren, damit
# alle Steuerelemente per Bildlauf erreichbar bleiben.
$form.AutoScrollMinSize = $designClient
# Startgroesse auf den nutzbaren Bildschirmbereich (ohne Taskleiste) begrenzen.
$wa = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
$startW = [Math]::Min($form.Width, $wa.Width)
$startH = [Math]::Min($form.Height, $wa.Height)
$form.Size = New-Object System.Drawing.Size($startW, $startH)
$form.StartPosition = 'CenterScreen'

function Add-Log { param([string]$Text, [System.Drawing.Color]$Color = [System.Drawing.Color]::Black)
    $timestamp = Get-Date
    $logBox.SelectionColor = $Color
    $logBox.AppendText(("[{0:HH:mm:ss}] {1}`r`n" -f $timestamp, $Text))
    $logBox.SelectionColor = $logBox.ForeColor
    $logBox.ScrollToCaret()
    Write-LogFile -Text $Text -Timestamp $timestamp
}

function ConvertTo-InventoryReportHtml {
    param($Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Write-InventoryHtmlReport {
    param(
        [string]$Path,
        [object[]]$Rows,
        [string]$BuildId,
        [string]$FirmwareSource,
        [int]$ExpectedServerCount = 0
    )
    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }

    $ordered = @($Rows | Sort-Object @{ Expression = { "$($_.ServerName)" } }, @{ Expression = { "$($_.Ilo)" } })
    $unsupported = @($ordered | Where-Object { "$($_.Status)" -eq 'Nicht unterstuetzt' }).Count
    $failed = @($ordered | Where-Object { -not $_.Success -and "$($_.Status)" -ne 'Nicht unterstuetzt' }).Count
    $reachable = $ordered.Count - $failed
    if ($ExpectedServerCount -le 0) { $ExpectedServerCount = $ordered.Count }
    $missing = [Math]::Max(0, $ExpectedServerCount - $ordered.Count)
    $serversWithUpdates = @($ordered | Where-Object { @($_.Updates).Count -gt 0 }).Count
    $updateCount = 0
    $componentCount = 0
    foreach ($row in $ordered) {
        $updateCount += @($row.Updates).Count
        $componentCount += @($row.Inventory).Count
    }

    $sb = [System.Text.StringBuilder]::new([Math]::Max(65536, $ordered.Count * 4096))
    [void]$sb.AppendLine('<!doctype html><html lang="de"><head><meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width,initial-scale=1"><title>HPE Firmware-Inventar</title>')
    [void]$sb.AppendLine('<style>body{font:14px Segoe UI,Arial,sans-serif;margin:0;background:#f4f6f8;color:#18212b}header{background:#17365d;color:#fff;padding:22px 28px}main{padding:20px 28px}.cards{display:flex;flex-wrap:wrap;gap:12px;margin:0 0 18px}.card{background:#fff;border:1px solid #d7dde5;border-radius:6px;padding:12px 18px;min-width:130px}.card b{display:block;font-size:24px}.tools{position:sticky;top:0;background:#f4f6f8;padding:8px 0 12px;z-index:2}.tools #q{width:min(520px,70%);padding:8px;border:1px solid #aeb8c4;border-radius:4px}details.server{background:#fff;border:1px solid #ccd4dd;border-left:5px solid #3f7f4f;border-radius:5px;margin:0 0 10px}details.updates{border-left-color:#d98400}details.error{border-left-color:#b3261e}summary{cursor:pointer;padding:12px 14px;font-weight:600}.meta{color:#536273;font-weight:400;margin-left:8px}.content{padding:0 14px 14px}.notice{padding:8px 10px;background:#fff4d6;border-radius:4px;margin:6px 0 12px}.errorText{background:#fde7e5}table{border-collapse:collapse;width:100%;margin:8px 0 16px}th,td{text-align:left;border:1px solid #d8dee6;padding:6px 8px;vertical-align:top}th{background:#eaf0f6}.muted{color:#657383}.hidden{display:none}footer{color:#657383;padding:4px 28px 24px}@media print{.tools{display:none}details.server{break-inside:avoid}details.server>summary{list-style:none}details.server:not([open])>.content{display:block}}</style>')
    [void]$sb.AppendLine('<script>function applyFilter(){const q=document.getElementById("q").value.toLowerCase();const only=document.getElementById("only").checked;for(const e of document.querySelectorAll(".server")){e.classList.toggle("hidden",!e.dataset.search.includes(q)||(only&&!e.classList.contains("updates")));}}function setOpen(v){for(const e of document.querySelectorAll(".server:not(.hidden)"))e.open=v;}</script></head><body>')
    [void]$sb.AppendLine("<header><h1>HPE Firmware-Inventar</h1><div>Erstellt: $(ConvertTo-InventoryReportHtml (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) &nbsp;|&nbsp; Build $(ConvertTo-InventoryReportHtml $BuildId)</div></header><main>")
    [void]$sb.AppendLine("<div class='cards'><div class='card'><b>$ExpectedServerCount</b>Ausgewaehlt</div><div class='card'><b>$reachable</b>Erreichbar</div><div class='card'><b>$failed</b>Fehler</div><div class='card'><b>$unsupported</b>Nicht unterstuetzt</div><div class='card'><b>$serversWithUpdates</b>Server mit Updates</div><div class='card'><b>$updateCount</b>Updates</div><div class='card'><b>$componentCount</b>Firmwareeintraege</div></div>")
    if ($missing -gt 0) {
        [void]$sb.AppendLine("<div class='notice errorText'><b>Bericht unvollstaendig:</b> $missing von $ExpectedServerCount ausgewaehlten Servern haben keinen Ergebnisdatensatz geliefert.</div>")
    }
    $sourceText = if ($FirmwareSource) { ConvertTo-InventoryReportHtml $FirmwareSource } else { 'Keine Firmwarequelle ausgewaehlt; der Bericht enthaelt nur den Ist-Stand.' }
    [void]$sb.AppendLine("<p><b>Firmwarequelle:</b> $sourceText</p>")
    [void]$sb.AppendLine('<div class="tools"><input id="q" oninput="applyFilter()" placeholder="Server, IP, Modell oder Firmware suchen..."> <label><input id="only" type="checkbox" onchange="applyFilter()"> nur Server mit Updates</label> <button onclick="setOpen(true)">Alle aufklappen</button> <button onclick="setOpen(false)">Alle einklappen</button></div>')

    foreach ($row in $ordered) {
        $name = if ([string]::IsNullOrWhiteSpace("$($row.ServerName)")) { "$($row.Ilo)" } else { "$($row.ServerName)" }
        $updates = @($row.Updates)
        $inventory = @($row.Inventory | Sort-Object @{ Expression = { "$($_.Name)" } }, @{ Expression = { "$($_.Version)" } })
        $cssClass = if (-not $row.Success) { 'server error' } elseif ($updates.Count) { 'server updates' } else { 'server' }
        $badge = if (-not $row.Success) { 'Fehler' } elseif ($updates.Count) { "$($updates.Count) Update(s)" } else { 'aktuell/keine Updates erkannt' }
        $searchParts = @($name, $row.Ilo, $row.Model, $row.Serial, $row.Status)
        $searchParts += @($inventory | ForEach-Object { "$($_.Name) $($_.Version)" })
        $searchParts += @($updates | ForEach-Object { "$($_.Package) $($_.Component) $($_.CurrentVersion) $($_.TargetVersion)" })
        $search = ConvertTo-InventoryReportHtml (($searchParts -join ' ').ToLowerInvariant())
        [void]$sb.AppendLine("<details class='$cssClass' data-search='$search'><summary>$(ConvertTo-InventoryReportHtml $name)<span class='meta'>$(ConvertTo-InventoryReportHtml $row.Ilo) | $(ConvertTo-InventoryReportHtml $row.Model) | $badge</span></summary><div class='content'>")
        [void]$sb.AppendLine("<table><tr><th>Servername</th><td>$(ConvertTo-InventoryReportHtml $name)</td><th>IP/iLO</th><td>$(ConvertTo-InventoryReportHtml $row.Ilo)</td></tr><tr><th>Modell</th><td>$(ConvertTo-InventoryReportHtml $row.Model)</td><th>Generation</th><td>$(ConvertTo-InventoryReportHtml $row.Generation)</td></tr><tr><th>Seriennummer</th><td>$(ConvertTo-InventoryReportHtml $row.Serial)</td><th>Status</th><td>$(ConvertTo-InventoryReportHtml $row.Status)</td></tr></table>")
        if ($row.Message) {
            $messageClass = if ($row.Success) { 'notice' } else { 'notice errorText' }
            [void]$sb.AppendLine("<div class='$messageClass'>$(ConvertTo-InventoryReportHtml $row.Message)</div>")
        }
        [void]$sb.AppendLine("<h3>Erforderliche Firmwareupdates ($($updates.Count))</h3>")
        if ($updates.Count) {
            [void]$sb.AppendLine('<table><thead><tr><th>Komponente</th><th>Aktueller Stand</th><th>Zielstand</th><th>Paket</th><th>Installationsweg</th></tr></thead><tbody>')
            foreach ($update in $updates) {
                [void]$sb.AppendLine("<tr><td>$(ConvertTo-InventoryReportHtml $update.Component)</td><td>$(ConvertTo-InventoryReportHtml $update.CurrentVersion)</td><td>$(ConvertTo-InventoryReportHtml $update.TargetVersion)</td><td>$(ConvertTo-InventoryReportHtml $update.Package)</td><td>$(ConvertTo-InventoryReportHtml $update.Route)</td></tr>")
            }
            [void]$sb.AppendLine('</tbody></table>')
        } else {
            [void]$sb.AppendLine("<p class='muted'>Keine erforderlichen Updates aus der ausgewaehlten Firmwarequelle ermittelt.</p>")
        }
        [void]$sb.AppendLine("<h3>Installierte Firmwarestaende ($($inventory.Count))</h3>")
        if ($inventory.Count) {
            [void]$sb.AppendLine('<table><thead><tr><th>Komponente</th><th>Version</th><th>Online aktualisierbar</th></tr></thead><tbody>')
            foreach ($item in $inventory) {
                $updateable = if ($null -eq $item.Updateable) { '-' } elseif ([bool]$item.Updateable) { 'Ja' } else { 'Nein' }
                [void]$sb.AppendLine("<tr><td>$(ConvertTo-InventoryReportHtml $item.Name)</td><td>$(ConvertTo-InventoryReportHtml $item.Version)</td><td>$updateable</td></tr>")
            }
            [void]$sb.AppendLine('</tbody></table>')
        } else {
            [void]$sb.AppendLine("<p class='muted'>Kein Firmwareinventar verfuegbar.</p>")
        }
        [void]$sb.AppendLine('</div></details>')
    }
    [void]$sb.AppendLine('</main><footer>Lokaler Bericht ohne externe Ressourcen. Es wurden keine zusaetzlichen iLO-Abfragen fuer die Berichtserstellung ausgefuehrt.</footer></body></html>')
    Set-Content -LiteralPath $Path -Value $sb.ToString() -Encoding utf8
    return $Path
}

# ─────────────────────────────────────────
# Async-Engine: ConcurrentQueue + Timer (UI-Thread)
# ─────────────────────────────────────────
$script:uiQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
$script:doneCount = 0
$script:totalCount = 0
$script:inventoryReportEnabled = $false
$script:inventoryReportPath = ''
$script:inventoryReportFirmwareSource = ''
$script:inventoryReportRows = New-Object System.Collections.ArrayList

function Find-Row { param([string]$Ilo)
    foreach ($it in $lv.Items) { if ($it.Name -eq $Ilo) { return $it } }
    return $null
}

$script:guiTimer = New-Object System.Windows.Forms.Timer
$script:guiTimer.Interval = 200
$script:guiTimer.Add_Tick({
    $msg = $null
    while ($script:uiQueue.TryDequeue([ref]$msg)) {
        switch ($msg.Type) {
            'INIT' {
                $lv.Items.Clear()
                foreach ($ilo in $msg.Servers) {
                    $li = New-Object System.Windows.Forms.ListViewItem($ilo)
                    $li.Name = $ilo
                    $li.SubItems.Add("-") | Out-Null     # Modell
                    $li.SubItems.Add("Wartet...") | Out-Null  # Phase
                    $li.SubItems.Add("0%") | Out-Null    # Fortschritt
                    $li.SubItems.Add("-") | Out-Null     # Ergebnis
                    $li.SubItems.Add("") | Out-Null      # Details
                    $lv.Items.Add($li) | Out-Null
                }
                $progress.Value = 0
                $script:doneCount = 0
                $script:totalCount = $msg.Servers.Count
                $statusLabel.Text = "0 / $($script:totalCount)"
            }
            'REPORT' {
                [void]$script:inventoryReportRows.Add($msg.Data)
            }
            'MODEL' {
                $li = Find-Row $msg.Ilo
                if ($li) { $li.SubItems[1].Text = $msg.Model }
            }
            'PHASE' {
                $li = Find-Row $msg.Ilo
                if ($li) { $li.SubItems[2].Text = $msg.Phase; $li.EnsureVisible() }
            }
            'PROGRESS' {
                $li = Find-Row $msg.Ilo
                if ($li) { $li.SubItems[3].Text = "$($msg.Percent)%" }
            }
            'LOG' { Add-Log $msg.Text }
            'DONE' {
                $li = Find-Row $msg.Ilo
                if ($li) {
                    $li.SubItems[2].Text = $msg.Phase
                    $li.SubItems[3].Text = if ($msg.Success) { "100%" } else { "-" }
                    $li.SubItems[4].Text = if ($msg.Success) { "OK" } else { "Fehler" }
                    $li.SubItems[5].Text = $msg.Detail
                    $li.ForeColor = if ($msg.Success) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkRed }
                    $li.EnsureVisible()
                }
                $script:doneCount++
                if ($script:totalCount -gt 0) {
                    $progress.Value = [Math]::Min(100, [int](($script:doneCount / $script:totalCount) * 100))
                }
                $statusLabel.Text = "$($script:doneCount) / $($script:totalCount)"
                $col = if ($msg.Success) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkRed }
                Add-Log "$($msg.Ilo): $($msg.Detail)" $col
            }
            'FINISHED' {
                Add-Log "=== Vorgang abgeschlossen ($($script:doneCount)/$($script:totalCount)) ===" ([System.Drawing.Color]::DarkBlue)
                $statusLabel.Text = "Fertig ($($script:doneCount)/$($script:totalCount))"
                if ($script:inventoryReportEnabled) {
                    try {
                        $reportPath = Write-InventoryHtmlReport -Path $script:inventoryReportPath -Rows @($script:inventoryReportRows) -BuildId $script:BuildId -FirmwareSource $script:inventoryReportFirmwareSource -ExpectedServerCount $script:totalCount
                        Add-Log "HTML-Inventarbericht erstellt: $reportPath" ([System.Drawing.Color]::DarkGreen)
                        $statusLabel.Text = "Fertig - Bericht: $reportPath"
                    } catch {
                        Add-Log "HTML-Inventarbericht konnte nicht erstellt werden: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                    } finally {
                        $script:inventoryReportEnabled = $false
                    }
                }
                $btnInventory.Enabled = $true; $btnFlash.Enabled = $true
            }
        }
    }
})
$script:guiTimer.Start()

# ─────────────────────────────────────────
# Runspace-Pool-Start (generisch)
# ─────────────────────────────────────────
function Start-Batch {
    param([array]$Servers, [scriptblock]$Worker, [hashtable]$ExtraArgs, [int]$MaxParallel)

    $script:uiQueue.Enqueue(@{ Type = 'INIT'; Servers = ($Servers | ForEach-Object { $_.Ilo }) })

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $pool = [runspacefactory]::CreateRunspacePool(1, $MaxParallel, $iss, $Host)
    $pool.Open()
    $script:pool = $pool
    $script:jobs = @()

    foreach ($entry in $Servers) {
        $argTable = @{
            iloCode      = $script:iloCode
            ilo          = $entry.Ilo
            user         = $entry.User
            pass         = $entry.Pass
            serverName    = $entry.Name
            uiQueue      = $script:uiQueue
            scriptFolder = $scriptFolder
        }
        if ($ExtraArgs) { foreach ($k in $ExtraArgs.Keys) { $argTable[$k] = $ExtraArgs[$k] } }
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        $null = $ps.AddScript($Worker).AddArgument($argTable)
        $handle = $ps.BeginInvoke()
        $script:jobs += [PSCustomObject]@{ PS = $ps; Handle = $handle; Ilo = $entry.Ilo }
    }

    if ($script:watch) { $script:watch.Stop(); $script:watch.Dispose() }
    $script:watch = New-Object System.Windows.Forms.Timer
    $script:watch.Interval = 1000
    $script:watch.Add_Tick({
        $allDone = $true
        foreach ($j in $script:jobs) { if (-not $j.Handle.IsCompleted) { $allDone = $false; break } }
        if ($allDone) {
            $script:watch.Stop()
            foreach ($j in $script:jobs) {
                try { $null = $j.PS.EndInvoke($j.Handle) } catch {}
                try { $j.PS.Dispose() } catch {}
            }
            try { $script:pool.Close(); $script:pool.Dispose() } catch {}
            $script:jobs = @()
            $script:uiQueue.Enqueue(@{ Type = 'FINISHED' })
        }
    })
    $script:watch.Start()
}

# ─────────────────────────────────────────
# Validierung gemeinsam
# ─────────────────────────────────────────
function Test-CommonInput {
    [OutputType([bool])]
    param()

    if ([string]::IsNullOrWhiteSpace($txtUser.Text) -or [string]::IsNullOrWhiteSpace($txtPass.Text)) {
        $null = [System.Windows.Forms.MessageBox]::Show(
            "Bitte iLO-Benutzer und Passwort eingeben.",
            "Credentials fehlen",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    return $true
}

# ═══════════════════════════════════════════════════════════════════
#  Aktion: Inventar pruefen (read-only)
# ═══════════════════════════════════════════════════════════════════
$btnInventory.Add_Click({
    if (-not (Test-CommonInput)) { return }
    $servers = Get-CheckedServers
    if ($servers.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Keine Server ausgewaehlt.", "Hinweis", 'OK', 'Warning'); return }

    $reportRequested = [bool]$chkInventoryReport.Checked
    $script:inventoryReportEnabled = $false
    $script:inventoryReportRows.Clear()
    $script:inventoryReportFirmwareSource = ''
    $script:inventoryReportPath = ''

    # Ist eine Firmwarequelle gewaehlt, erweitert die read-only Inventar-Pruefung
    # das Ergebnis um eine serverspezifische Update-Vorschau. Ohne Quelle bleibt
    # sie weiterhin eine reine Bestandsaufnahme.
    $previewEnabled = -not [string]::IsNullOrWhiteSpace($txtFw.Text)
    $previewMode = 'none'
    $previewBaseDir = ''
    $previewComponents = @()
    $packageMetaCache = @{}
    if ($previewEnabled) {
        $previewBaseDir = $txtFw.Text.Trim()
        if ($script:IsSynergyMode) {
            $previewMode = 'folder'
            try { $previewBaseDir = Resolve-SynergyFirmwareFolder -BaseDir $previewBaseDir } catch {
                [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Fehler", 'OK', 'Error'); return
            }
            $previewComponents = @(Get-ChildItem -LiteralPath $previewBaseDir -Filter *.fwpkg -File | Sort-Object Name | ForEach-Object { $_.FullName })
        } elseif ($rbBaseDir.Checked) {
            $previewMode = 'basedir'
            if (-not (Test-Path -LiteralPath $previewBaseDir -PathType Container)) {
                [System.Windows.Forms.MessageBox]::Show("Firmware-Verzeichnis nicht gefunden:`n$previewBaseDir", "Fehler", 'OK', 'Error'); return
            }
        }
        elseif ($rbFolder.Checked) {
            $previewMode = 'folder'
            if (-not (Test-Path -LiteralPath $previewBaseDir -PathType Container)) {
                [System.Windows.Forms.MessageBox]::Show("Firmware-Ordner nicht gefunden:`n$previewBaseDir", "Fehler", 'OK', 'Error'); return
            }
            $previewComponents = @(Get-ChildItem -LiteralPath $previewBaseDir -Filter *.fwpkg -File | Sort-Object Name | ForEach-Object { $_.FullName })
        }
        else {
            $previewMode = 'file'
            if (-not (Test-Path -LiteralPath $previewBaseDir -PathType Leaf)) {
                [System.Windows.Forms.MessageBox]::Show("Firmware-Datei nicht gefunden:`n$previewBaseDir", "Fehler", 'OK', 'Error'); return
            }
            $previewComponents = @($previewBaseDir)
        }
        if ($previewMode -ne 'basedir' -and -not $previewComponents.Count) {
            [System.Windows.Forms.MessageBox]::Show("Keine .fwpkg-Dateien fuer die Update-Vorschau gefunden.", "Hinweis", 'OK', 'Warning'); return
        }
        if ($script:IsSynergyMode) {
            $packageMetaCache = Get-FirmwarePackageIndex -Components $previewComponents
        }
    }
    $script:inventoryReportEnabled = $reportRequested
    if ($script:inventoryReportEnabled) {
        $reportsDir = Join-Path $scriptFolder 'Reports'
        $script:inventoryReportPath = Join-Path $reportsDir ("FirmwareInventory_{0:yyyy-MM-dd_HHmmss}.html" -f (Get-Date))
        if ($previewEnabled) { $script:inventoryReportFirmwareSource = $previewBaseDir }
    }

    $btnInventory.Enabled = $false; $btnFlash.Enabled = $false
    $previewText = if ($previewEnabled) { ' mit Update-Vorschau' } else { '' }
    Add-Log "=== Inventar-Pruefung$previewText fuer $($servers.Count) Server gestartet (Build $script:buildId) ===" ([System.Drawing.Color]::DarkBlue)

    $worker = {
        param($p)
        $iloCode = $p.iloCode; $ilo = $p.ilo; $user = $p.user; $pass = $p.pass; $uiQueue = $p.uiQueue
        $serverName = "$($p.serverName)"; $reportEnabled = [bool]$p.reportEnabled
        $diagnosticEnabled = [bool]$p.diagnosticEnabled
        $previewEnabled = [bool]$p.previewEnabled; $previewMode = $p.previewMode; $previewBaseDir = $p.previewBaseDir
        $previewComponents = @($p.previewComponents); $kindFilter = @($p.kindFilter); $scriptFolder = $p.scriptFolder
        $synergyMode = [bool]$p.synergyMode; $strictNoDowngrade = [bool]$p.strictNoDowngrade
        $log = { param($t) $uiQueue.Enqueue(@{ Type = 'LOG'; Text = "$ilo : $t" }) }.GetNewClosure()
        $sendReport = {
            param([bool]$Success, [string]$Status, $Info, [object[]]$Inventory, [object[]]$Updates, [string]$Message)
            if (-not $reportEnabled) { return }
            $inventoryRows = @()
            foreach ($item in @($Inventory)) {
                $inventoryRows += @{
                    Name       = "$($item.Name)"
                    Version    = "$($item.Version)"
                    Updateable = $item.Updateable
                }
            }
            $model = ''; $generation = ''; $serial = ''
            if ($Info) {
                $model = "$($Info.Model)"
                if ($Info.Gen) { $generation = "Gen$($Info.Gen)" }
                $serial = "$($Info.Serial)"
            }
            $uiQueue.Enqueue(@{
                Type = 'REPORT'
                Data = @{
                    ServerName = $serverName
                    Ilo        = $ilo
                    Model      = $model
                    Generation = $generation
                    Serial     = $serial
                    Success    = $Success
                    Status     = $Status
                    Message    = $Message
                    Inventory  = @($inventoryRows)
                    Updates    = @($Updates)
                }
            })
        }.GetNewClosure()
        $sess = $null
        $info = $null
        $inv = @()
        $fullInv = @()
        $deviceInv = @()
        $storageInv = $null
        try {
            Invoke-Expression $iloCode
            if ($p.packageMetaCache) { $script:FwpkgMetaCache = $p.packageMetaCache }
            & $log "Login am iLO..."
            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = 'Login...' })
            $sess = ILO-Login -Ilo $ilo -User $user -Pass $pass
            & $log "Login OK, lese Inventar..."
            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = 'Lese Inventar...' })
            $info = ILO-GetSystemInfo -Ilo $ilo -Token $sess.Token
            $uiQueue.Enqueue(@{ Type = 'MODEL'; Ilo = $ilo; Model = $info.Model })
            & $log "Systeminfo: Modell='$($info.Model)', Gen=$($info.Gen), iLO-Gen=$($info.iLO), SN=$($info.Serial)"

            if ($synergyMode -and $info.Model -notmatch '(?i)\bSynergy\b') {
                & $log "Uebersprungen: '$($info.Model)' ist kein Synergy Compute Module."
                & $sendReport $false 'Nicht unterstuetzt' $info @() @() "Kein Synergy Compute Module"
                $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Kein Synergy'; Detail = "Modell '$($info.Model)' ist kein Synergy Compute Module" })
                return
            }
            if (-not $synergyMode -and $info.Model -match '(?i)\bSynergy\b') {
                & $log "Uebersprungen: Synergy Compute Module ist im Rackmount-Modus nicht zulaessig."
                & $sendReport $false 'Nicht unterstuetzt' $info @() @() "Synergy im Rackmount-Modus blockiert"
                $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Synergy blockiert'; Detail = "Synergy Compute Module '$($info.Model)' im Rackmount-Modus blockiert" })
                return
            }
            if ($info.Gen -gt 0 -and $info.Gen -lt 10) {
                & $log "Uebersprungen: Gen$($info.Gen) < Gen10 (nicht unterstuetzt)"
                & $sendReport $false 'Nicht unterstuetzt' $info @() @() "Generation Gen$($info.Gen) < Gen10"
                $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Nicht unterstuetzt'; Detail = "Generation Gen$($info.Gen) < Gen10 - uebersprungen" })
                return
            }

            $inv = ILO-GetFirmwareInventory -Ilo $ilo -Token $sess.Token
            if ($diagnosticEnabled) {
                try { $fullInv = @(ILO-GetFirmwareInventoryDetails -Ilo $ilo -Token $sess.Token) } catch { $fullInv = @() }
                if (-not $fullInv.Count) {
                    # Der Export bleibt auch bei aelteren iLO-Versionen nuetzlich,
                    # die nur die reduzierte Inventaransicht liefern.
                    $fullInv = @($inv | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Version = $_.Version } })
                }
                try { $deviceInv = @(ILO-GetDeviceInventoryDiagnostic -Ilo $ilo -Token $sess.Token) } catch { $deviceInv = @() }
                try { $storageInv = ILO-GetStorageInventoryDiagnostic -Ilo $ilo -Token $sess.Token -IloGeneration $info.iLO } catch { $storageInv = $null }
                if ($storageInv) {
                    & $log "Diagnose-Storage-Inventar: $($storageInv.Controllers.Count) Controller, $($storageInv.Drives.Count) Laufwerk(e) [$($storageInv.Strategy)]"
                } else {
                    & $log 'Diagnose-Storage-Inventar konnte nicht gelesen werden.'
                }
            }
            $bios = ($inv | Where-Object { $_.Name -match '(?i)System ROM|BIOS' } | Select-Object -First 1)
            $iloFw = ($inv | Where-Object { $_.Name -match '(?i)iLO' } | Select-Object -First 1)
            $biosV = if ($bios) { $bios.Version } else { '?' }
            $iloV = if ($iloFw) { $iloFw.Version } else { '?' }
            $detail = "BIOS: $biosV | iLO: $iloV | Komponenten: $($inv.Count)"

            $serverComponents = @()
            $updateDecision = [pscustomobject]@{
                Status             = if ($previewEnabled) { 'Firmwarequelle wird ausgewertet' } else { 'Keine Firmwarequelle ausgewaehlt' }
                FirmwareSource     = if ($previewEnabled) { $previewBaseDir } else { '' }
                SelectedPackages   = @()
                WouldUpdate        = @()
                AlreadyCurrent     = @()
                NotInstalled       = @()
                Superseded         = @()
                Ambiguous          = @()
                NotOnlineUpdateable = @()
                RepositoryQueue    = @()
                PackageCorrelation  = @()
            }
            if ($previewEnabled) {
                $serverComponents = @($previewComponents)
                if ($previewMode -eq 'basedir') {
                    try {
                        $folder = Resolve-FirmwareFolder -BaseDir $previewBaseDir -Model $info.Model -Gen $info.Gen
                        $serverComponents = @(Get-ChildItem -LiteralPath $folder -Filter *.fwpkg -File | Sort-Object Name | ForEach-Object { $_.FullName })
                        & $log "Update-Vorschau: Typ-Ordner '$([System.IO.Path]::GetFileName($folder))' mit $($serverComponents.Count) Paket(en) gefunden."
                    } catch {
                        $detail += " | Update-Vorschau nicht moeglich: $($_.Exception.Message)"
                        $updateDecision.Status = 'Firmwareordner nicht aufloesbar'
                        $updateDecision | Add-Member -NotePropertyName Error -NotePropertyValue $_.Exception.Message -Force
                        if ($diagnosticEnabled) {
                            $diagPath = Write-FirmwareDiagnostic -StateDir $scriptFolder -Ilo $ilo -SystemInfo $info -FirmwareInventory $fullInv -DeviceInventory $deviceInv -StorageInventory $storageInv -UpdateDecision $updateDecision -FirmwarePackages @()
                            if ($diagPath) { $detail += " | Diagnose: $diagPath" }
                        }
                        & $log $detail
                        & $sendReport $true $updateDecision.Status $info $inv @() $detail
                        $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $true; Phase = "OK (Gen$($info.Gen))"; Detail = $detail })
                        return
                    }
                }
                $selection = Select-FirmwarePackagesForServer -Inventory $inv -Components $serverComponents -KindFilter $kindFilter -StrictMetadata $strictNoDowngrade
                $preview = Get-FirmwareUpdatePreview -Inventory $inv -SystemInfo $info -Components $selection.Components -KindFilter $kindFilter -StrictVersionComparison $strictNoDowngrade
                $preview.Absent = @($selection.NotInstalled)
                $preview.Ambiguous = @($preview.Ambiguous + $selection.Ambiguous)
                $correlation = Get-FirmwareCorrelationDiagnostics -Inventory $inv -Components $serverComponents -KindFilter $kindFilter
                $updateDecision = [pscustomobject]@{
                    Status              = 'Ausgewertet (read-only)'
                    FirmwareSource      = $previewBaseDir
                    SelectedPackages    = @($selection.Components | ForEach-Object { [System.IO.Path]::GetFileName($_) })
                    WouldUpdate         = @($preview.Updates)
                    AlreadyCurrent      = @($preview.Current)
                    NotInstalled        = @($selection.NotInstalled)
                    Superseded          = @($selection.Superseded)
                    Ambiguous           = @($preview.Ambiguous)
                    NotOnlineUpdateable = @($selection.NotOnlineUpdateable + $preview.NotOnlineUpdateable | Select-Object -Unique)
                    RepositoryQueue     = @($preview.RepositoryQueue)
                    PackageCorrelation  = @($correlation)
                }
                if ($preview.Updates.Count) {
                    $detail += " | Updates ($($preview.Updates.Count)): $($preview.Updates -join '; ')"
                } else {
                    $detail += ' | Updates: keine'
                }
                if ($preview.Current.Count) { $detail += " | Aktuell: $($preview.Current.Count)" }
                if ($preview.Absent.Count) { $detail += " | Nicht verbaut/nicht zuordenbar: $($preview.Absent.Count)" }
                if ($preview.Ambiguous.Count) { $detail += " | Uneindeutig (gesperrt): $($preview.Ambiguous.Count)" }
                $notOnlineCount = @($selection.NotOnlineUpdateable + $preview.NotOnlineUpdateable | Select-Object -Unique).Count
                if ($notOnlineCount) { $detail += " | Nicht online flashbar: $notOnlineCount" }
                $queueCount = @($preview.RepositoryQueue).Count
                if ($queueCount) { $detail += " | Repository/Queue-Updates: $queueCount" }
            } else {
                $detail += ' | Update-Vorschau: Firmwarequelle auswaehlen'
            }
            $reportUpdates = @()
            if ($previewEnabled -and $preview -and $correlation) {
                foreach ($corr in @($correlation)) {
                    $updateText = ''
                    foreach ($candidate in @($preview.Updates)) {
                        if ("$candidate".StartsWith("$($corr.File) ", [System.StringComparison]::OrdinalIgnoreCase) -or "$candidate" -eq "$($corr.File)") {
                            $updateText = "$candidate"
                            break
                        }
                    }
                    if (-not $updateText) { continue }
                    $componentNames = @($corr.Matches | ForEach-Object { "$($_.Name)" } | Where-Object { $_ } | Select-Object -Unique)
                    $installedVersions = @($corr.Matches | ForEach-Object { "$($_.Name): $($_.Version)" } | Select-Object -Unique)
                    $reportUpdates += @{
                        Component      = if ($componentNames.Count) { $componentNames -join '; ' } else { "$($corr.Kind)" }
                        CurrentVersion = if ($installedVersions.Count) { $installedVersions -join '; ' } else { '-' }
                        TargetVersion  = "$($corr.TargetVersion)"
                        Package        = "$($corr.File)"
                        Route          = if ($updateText -match '\[Repository/Queue\]') { 'iLO Repository/Update Queue' } else { 'Direktes Firmwareupdate' }
                    }
                }
            }
            if ($diagnosticEnabled) {
                $packageDiag = Get-FirmwarePackageDiagnostics -Components $serverComponents
                $diagPath = Write-FirmwareDiagnostic -StateDir $scriptFolder -Ilo $ilo -SystemInfo $info -FirmwareInventory $fullInv -DeviceInventory $deviceInv -StorageInventory $storageInv -UpdateDecision $updateDecision -FirmwarePackages $packageDiag
                if ($diagPath) {
                    $detail += " | Diagnose: $diagPath"
                    & $log "Diagnoseexport geschrieben: $diagPath"
                } else {
                    & $log 'Diagnoseexport konnte nicht geschrieben werden.'
                }
            }
            & $log "Inventar gelesen: $detail"
            & $sendReport $true $updateDecision.Status $info $inv $reportUpdates $detail
            $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $true; Phase = "OK (Gen$($info.Gen))"; Detail = $detail })
        }
        catch {
            $errMsg = $_.Exception.Message
            if ($_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber) { $errMsg = "$errMsg (Zeile $($_.InvocationInfo.ScriptLineNumber))" }
            & $log "FEHLER: $errMsg"
            & $sendReport $false 'Fehler' $info $inv @() $errMsg
            $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Fehler'; Detail = $_.Exception.Message })
        }
        finally {
            if ($sess) { try { ILO-Logout -Ilo $ilo -Token $sess.Token -SessionUri $sess.SessionUri; & $log "Logout OK" } catch {} }
        }
    }

    Start-Batch -Servers $servers -Worker $worker -ExtraArgs @{
        previewEnabled = $previewEnabled; previewMode = $previewMode; previewBaseDir = $previewBaseDir
        previewComponents = $previewComponents; kindFilter = @(& $script:GetKindFilter)
        reportEnabled = $script:inventoryReportEnabled
        diagnosticEnabled = [bool]$chkDiagnosticReport.Checked
        synergyMode = $script:IsSynergyMode; strictNoDowngrade = $script:IsSynergyMode
        packageMetaCache = $packageMetaCache
    } -MaxParallel ([int]$numPar.Value)
})

# ═══════════════════════════════════════════════════════════════════
#  Aktion: Firmware aktualisieren
# ═══════════════════════════════════════════════════════════════════
$btnFlash.Add_Click({
    if (-not (Test-CommonInput)) { return }

    # Firmware-Quelle bestimmen
    $fwPath = $txtFw.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($fwPath)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Firmware-Datei, -Ordner oder Basisverzeichnis auswaehlen.", "Fehler", 'OK', 'Error'); return
    }
    $mode = 'file'
    $components = @()
    if ($script:IsSynergyMode) {
        $mode = 'folder'
        try { $synergyFolder = Resolve-SynergyFirmwareFolder -BaseDir $fwPath } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Fehler", 'OK', 'Error'); return
        }
        $components = @(Get-ChildItem -LiteralPath $synergyFolder -Filter *.fwpkg -File | Sort-Object Name | ForEach-Object { $_.FullName })
        if (-not $components.Count) {
            [System.Windows.Forms.MessageBox]::Show("Keine .fwpkg-Dateien im Synergy-Firmwareordner gefunden:`n$synergyFolder", "Fehler", 'OK', 'Warning'); return
        }
        $fwPath = $synergyFolder
    } elseif ($rbBaseDir.Checked) {
        # Firmware-Verzeichnis: Komponenten werden je Server (nach Modell) erst im Worker ermittelt.
        $mode = 'basedir'
        if (-not (Test-Path -LiteralPath $fwPath -PathType Container)) {
            [System.Windows.Forms.MessageBox]::Show("Firmware-Verzeichnis nicht gefunden:`n$fwPath", "Fehler", 'OK', 'Error'); return
        }
        $subDirs = @(Get-ChildItem -LiteralPath $fwPath -Directory -ErrorAction SilentlyContinue)
        if ($subDirs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Im Firmware-Verzeichnis sind keine Typ-Unterordner vorhanden.`n`nErwartet z.B. 'DL360_Gen11', 'DL380_Gen11' oder 'DL380a_Gen11'.", "Fehler", 'OK', 'Warning'); return
        }
    }
    elseif ($rbFolder.Checked) {
        $mode = 'folder'
        if (-not (Test-Path -LiteralPath $fwPath -PathType Container)) {
            [System.Windows.Forms.MessageBox]::Show("Ordner nicht gefunden:`n$fwPath", "Fehler", 'OK', 'Error'); return
        }
        $components = @(Get-ChildItem -LiteralPath $fwPath -Filter *.fwpkg -File | Sort-Object Name | ForEach-Object { $_.FullName })
        if ($components.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine .fwpkg-Dateien im Ordner gefunden.", "Fehler", 'OK', 'Warning'); return
        }
    } else {
        $mode = 'file'
        if (-not (Test-Path -LiteralPath $fwPath -PathType Leaf)) {
            [System.Windows.Forms.MessageBox]::Show("Datei nicht gefunden:`n$fwPath", "Fehler", 'OK', 'Error'); return
        }
        $components = @($fwPath)
    }

    # iLO-eigene Firmware ans Ende sortieren (Reset erst am Schluss) - nur fuer feste Komponentenliste
    if ($mode -ne 'basedir') {
        $components = @(Sort-FirmwareComponentsForUpdate -Components $components)
    }
    $packageMetaCache = if ($script:IsSynergyMode) { Get-FirmwarePackageIndex -Components $components } else { @{} }

    $servers = Get-CheckedServers
    if ($servers.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Keine Server ausgewaehlt.", "Hinweis", 'OK', 'Warning'); return }

    # Optionale lokale SHA-256-Pruefung (nur Einzeldatei)
    $expectedSha = $txtSha.Text.Trim()
    if ($rbFile.Checked -and -not [string]::IsNullOrWhiteSpace($expectedSha)) {
        Add-Log "Berechne lokalen SHA-256 (kann etwas dauern)..."
        $statusLabel.Text = "SHA-256 wird geprueft..."
        try { $localHash = (Get-FileHash -LiteralPath $components[0] -Algorithm SHA256).Hash } catch {
            [System.Windows.Forms.MessageBox]::Show("SHA-256 Berechnung fehlgeschlagen: $($_.Exception.Message)", "Fehler", 'OK', 'Error'); return
        }
        if ($localHash -ine $expectedSha) {
            $m = "SHA-256 stimmt NICHT ueberein!`nErwartet:  $expectedSha`nBerechnet: $localHash`n`nAbgebrochen."
            [System.Windows.Forms.MessageBox]::Show($m, "SHA-256 Mismatch", 'OK', 'Error')
            Add-Log $m ([System.Drawing.Color]::Red); return
        }
        Add-Log "SHA-256 OK ($localHash)" ([System.Drawing.Color]::DarkGreen)
    }

    $fwFilterMsg = switch ($cmbFwFilter.SelectedIndex) {
        1 { "NUR-iLO-Modus: Es wird ausschliesslich die iLO-Firmware verarbeitet.`nAndere Komponenten werden ignoriert.`n`n" }
        2 { "FILTER-Modus 'Nur iLO + BIOS + SPS': Es werden ausschliesslich iLO-, System-ROM/BIOS- und SPS/ME-Pakete verarbeitet.`nAndere Komponenten (CPLD/IE/PMC/...) werden ignoriert.`n`n" }
        default { "" }
    }

    # Eigener Sicherheitsdialog mit Standardauswahl "Nein", damit der Hinweis
    # nicht im allgemeinen (teilweise sehr langen) Bestaetigungstext untergeht.
    $bitlockerText = "ACHTUNG - BITLOCKER / TPM`n`n" +
        "Falls iLO bei einem Repository-Task TPMOverride verlangt, wird dieser " +
        "Parameter nach erneuter Bestaetigung durch iLO gezielt verwendet.`n`n" +
        "BitLocker oder eine andere TPM-basierte Laufwerksverschluesselung " +
        "muss auf ALLEN ausgewaehlten Servern VOR dem Update ausgeschaltet " +
        "oder angehalten (Suspend) sein. Andernfalls kann Windows beim " +
        "naechsten Start den Recovery-Key verlangen.`n`n" +
        "Bestaetigen Sie mit JA, dass BitLocker ausgeschaltet bzw. angehalten ist."
    $bitlockerConfirmed = [System.Windows.Forms.MessageBox]::Show(
        $bitlockerText,
        "BitLocker vor Firmwareupdate pruefen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($bitlockerConfirmed -ne [System.Windows.Forms.DialogResult]::Yes) {
        Add-Log "Firmware-Update abgebrochen: BitLocker-/TPM-Warnung wurde nicht bestaetigt." ([System.Drawing.Color]::DarkOrange)
        return
    }
    Add-Log "BitLocker-/TPM-Warnung bestaetigt; ein von iLO ausdruecklich verlangtes TPMOverride darf fuer Queue-Tasks verwendet werden." ([System.Drawing.Color]::DarkOrange)

    if ($mode -eq 'basedir') {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "ACHTUNG - Firmware wird auf $($servers.Count) Server geflasht (max. $([int]$numPar.Value) parallel).`n`n" +
            $fwFilterMsg +
            "Automatische Typ-Zuordnung aus Basisverzeichnis:`n$fwPath`n`n" +
            "Je Server wird das Modell live am iLO ermittelt und der passende Typ-Unterordner (.fwpkg) verwendet.`n" +
            "Server ohne passenden Unterordner werden uebersprungen.`n" +
            "Bereits aktuelle Komponenten werden automatisch uebersprungen.`n`n" +
            "Es wird NICHT automatisch rebootet.`n" +
            "Reine iLO-Updates sind nach dem iLO-Selbstreset sofort aktiv (kein Reboot).`n" +
            "BIOS/SPS/Systemfirmware wird erst beim naechsten regulaeren Neustart aktiv.`n`nFortfahren?",
            "Firmware-Update bestaetigen", 'YesNo', 'Warning')
        if ($res -ne 'Yes') { return }
        $btnInventory.Enabled = $false; $btnFlash.Enabled = $false
        Add-Log "=== Firmware-Update (autom. Typ-Zuordnung) gestartet auf $($servers.Count) Server (Build $script:buildId) ===" ([System.Drawing.Color]::DarkBlue)
    } else {
        $compNames = ($components | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join ", "
        $res = [System.Windows.Forms.MessageBox]::Show(
            "ACHTUNG - Firmware wird auf $($servers.Count) Server geflasht (max. $([int]$numPar.Value) parallel).`n`n" +
            $fwFilterMsg +
            "Komponenten ($($components.Count)):`n$compNames`n`n" +
            $(if ($script:IsSynergyMode) { "Strikter Versionsschutz: Nur eindeutig neuere JSON-Zielversionen werden verarbeitet.`n" } else { "Bereits aktuelle Komponenten werden automatisch uebersprungen.`n" }) +
            "Es wird NICHT automatisch rebootet.`n" +
            "Reine iLO-Updates sind nach dem iLO-Selbstreset sofort aktiv (kein Reboot).`n" +
            "BIOS/SPS/Systemfirmware wird erst beim naechsten regulaeren Neustart aktiv.`n`nFortfahren?",
            "Firmware-Update bestaetigen", 'YesNo', 'Warning')
        if ($res -ne 'Yes') { return }
        $btnInventory.Enabled = $false; $btnFlash.Enabled = $false
        Add-Log "=== Firmware-Update gestartet: $($components.Count) Komponente(n) auf $($servers.Count) Server (Build $script:buildId) ===" ([System.Drawing.Color]::DarkBlue)
    }

    $worker = {
        param($p)
        $iloCode = $p.iloCode; $ilo = $p.ilo; $user = $p.user; $pass = $p.pass; $uiQueue = $p.uiQueue
        $components = $p.components; $mode = $p.mode; $baseDir = $p.baseDir; $scriptFolder = $p.scriptFolder
        $kindFilter = @($p.kindFilter)
        $synergyMode = [bool]$p.synergyMode
        $strictNoDowngrade = [bool]$p.strictNoDowngrade
        Invoke-Expression $iloCode
        if ($p.packageMetaCache) { $script:FwpkgMetaCache = $p.packageMetaCache }
        # Lokaler Log-Helfer: schreibt in das gemeinsame UI-/Datei-Log.
        $log = { param($t) $uiQueue.Enqueue(@{ Type = 'LOG'; Text = "$ilo : $t" }) }.GetNewClosure()
        $sess = $null
        try {
            & $log "Login am iLO..."
            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = 'Login...' })
            $sess = ILO-Login -Ilo $ilo -User $user -Pass $pass
            $token = $sess.Token
            & $log "Login OK (Session: $($sess.SessionUri))"

            $info = ILO-GetSystemInfo -Ilo $ilo -Token $token
            $uiQueue.Enqueue(@{ Type = 'MODEL'; Ilo = $ilo; Model = $info.Model })
            & $log "Systeminfo: Modell='$($info.Model)', Gen=$($info.Gen), iLO-Gen=$($info.iLO), SN=$($info.Serial)"
            if ($synergyMode -and $info.Model -notmatch '(?i)\bSynergy\b') {
                & $log "BLOCKIERT: '$($info.Model)' ist kein Synergy Compute Module."
                $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Kein Synergy'; Detail = "Modell '$($info.Model)' ist kein Synergy Compute Module" })
                return
            }
            if (-not $synergyMode -and $info.Model -match '(?i)\bSynergy\b') {
                & $log "BLOCKIERT: '$($info.Model)' ist ein Synergy Compute Module und im Rackmount-Modus nicht zulaessig."
                $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Synergy - blockiert'; Detail = "Synergy Compute Module '$($info.Model)' im Rackmount-Modus blockiert" })
                return
            }
            if ($info.Gen -gt 0 -and $info.Gen -lt 10) {
                & $log "Uebersprungen: Gen$($info.Gen) < Gen10 (nicht unterstuetzt)"
                $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Nicht unterstuetzt'; Detail = "Gen$($info.Gen) < Gen10 - uebersprungen" })
                return
            }

            # Basisverzeichnis-Modus: passenden Typ-Unterordner je Server ermitteln
            if ($mode -eq 'basedir') {
                $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = 'Typ-Ordner suchen...' })
                & $log "Suche Typ-Ordner im Firmware-Verzeichnis '$baseDir'..."
                $folder = $null
                try { $folder = Resolve-FirmwareFolder -BaseDir $baseDir -Model $info.Model -Gen $info.Gen } catch { $folder = $null; & $log "Typ-Ordner-Fehler: $($_.Exception.Message)" }
                if (-not $folder) {
                    $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Kein Typ-Ordner'; Detail = "Kein passender Unterordner fuer '$($info.Model)' (Gen$($info.Gen)) im Basisverzeichnis" })
                    return
                }
                $components = @(Get-ChildItem -LiteralPath $folder -Filter *.fwpkg -File | Sort-Object Name | ForEach-Object { $_.FullName })
                if ($components.Count -eq 0) {
                    & $log "Keine .fwpkg-Dateien im Typ-Ordner '$([System.IO.Path]::GetFileName($folder))'"
                    $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Keine Firmware'; Detail = "Im Typ-Ordner '$([System.IO.Path]::GetFileName($folder))' keine .fwpkg-Dateien" })
                    return
                }
                # iLO-eigene Firmware ans Ende sortieren (Reset erst am Schluss)
                $components = @(Sort-FirmwareComponentsForUpdate -Components $components)
                & $log "Typ-Ordner '$([System.IO.Path]::GetFileName($folder))', $($components.Count) Komponente(n): $((@($components | ForEach-Object { [System.IO.Path]::GetFileName($_) })) -join ', ')"
            }

            # Komponentenfilter: Dateiliste auf die gewaehlten Firmware-Typen
            # reduzieren. Greift fuer alle Quell-Modi (Einzeldatei/Ordner/
            # Basisverzeichnis), da hier die effektive Dateiliste feststeht.
            if (@($kindFilter).Count) {
                $preCount = @($components).Count
                $components = @($components | Where-Object { $kindFilter -contains (Get-ComponentKind -FileName ([System.IO.Path]::GetFileName($_))).Kind })
                $filterTxt = if ($components.Count) { (@($components | ForEach-Object { [System.IO.Path]::GetFileName($_) })) -join ', ' } else { '(keine)' }
                & $log "Komponentenfilter [$($kindFilter -join '+')]: $($components.Count) von $preCount Komponente(n) passen: $filterTxt"
                if ($components.Count -eq 0) {
                    $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Keine passende Firmware'; Detail = "Filter '$($kindFilter -join '+')': keine passende .fwpkg in der Firmware-Quelle gefunden" })
                    return
                }
            }

            # Firmware-Inventar einmal lesen. Es ist fuer den Versionsvergleich
            # UND den serverspezifischen Hardware-Abgleich verpflichtend.
            $inv = @()
            try { $inv = ILO-GetFirmwareInventory -Ilo $ilo -Token $token } catch { & $log "Inventar konnte nicht gelesen werden: $($_.Exception.Message)" }
            if (-not $inv.Count) {
                & $log 'Abbruch: Firmware-Inventar ist leer oder nicht verfuegbar; kein komponentenspezifisches Update ohne Hardware-Abgleich.'
                $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Inventar fehlt'; Detail = 'Firmware-Inventar leer/nicht verfuegbar - Update aus Sicherheitsgruenden nicht gestartet' })
                return
            }
            & $log "Firmware-Inventar ($($inv.Count) Eintraege): $((@($inv | ForEach-Object { "$($_.Name)=$($_.Version)" })) -join ' | ')"
            # Endgueltig fehlgeschlagene Queue-Tasks koennen nicht mehr ausgefuehrt
            # werden und blockieren auf iLO 5 trotzdem jeden nachfolgenden Task.
            # Wartende, laufende und erfolgreiche Tasks werden nicht entfernt.
            try {
                $retiredFailed = @(ILO-RetireFailedUpdateTasks -Ilo $ilo -Token $token -LogCb $log)
                if ($retiredFailed.Count) {
                    & $log "Queue-Bereinigung abgeschlossen: $($retiredFailed.Count) nicht mehr ausfuehrbare(r) Fehler-Task(s) entfernt."
                }
            } catch {
                & $log "Abbruch: UpdateTaskQueue konnte vor der Paketauswahl nicht sicher bereinigt/gelesen werden: $($_.Exception.Message)"
                $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Queue nicht lesbar'; Detail = 'Keine Firmwareauswahl ohne vollstaendig lesbare UpdateTaskQueue' })
                return
            }
            $selection = Select-FirmwarePackagesForServer -Inventory $inv -Components $components -KindFilter $kindFilter -StrictMetadata $strictNoDowngrade
            # Nach der serverspezifischen Paketwahl nochmals verbindlich ordnen.
            # Damit bleibt iLO auch dann zuletzt, wenn Filter oder Paketersetzung
            # die zuvor aufgebaute Komponentenliste veraendert haben.
            $components = @(Sort-FirmwareComponentsForUpdate -Components $selection.Components)
            if ($selection.NotInstalled.Count) {
                & $log "Hardware-Filter: $($selection.NotInstalled.Count) Paket(e) passen zu keiner verbauten Komponente und werden nicht verarbeitet."
            }
            if ($selection.NotOnlineUpdateable.Count) {
                & $log "Sicherheitsfilter: $($selection.NotOnlineUpdateable.Count) exakt zugeordnete Paket(e) sind laut iLO nicht online updatefaehig und werden nicht hochgeladen: $($selection.NotOnlineUpdateable -join '; ')"
            }
            if ($selection.RepositoryQueue.Count) {
                & $log "Update-Routing vor Versionsvergleich: $($selection.RepositoryQueue.Count) Paket(e) sind laut JSON/Live-Inventar fuer ComponentRepository + UpdateTaskQueue vorgesehen: $($selection.RepositoryQueue -join '; ')"
            }
            if ($selection.Superseded.Count) {
                & $log "Versionsfilter: $($selection.Superseded.Count) aeltere Paketversion(en) werden zugunsten der jeweils neuesten passenden Version uebersprungen."
            }
            if ($selection.Ambiguous.Count) {
                & $log "Sicherheitsfilter: $($selection.Ambiguous.Count) Komponente(n) haben mehrere gleich neue Paketdateien und bleiben gesperrt: $($selection.Ambiguous -join '; ')"
            }
            if (-not $components.Count) {
                & $log 'Keine eindeutig passende Firmwarekomponente fuer diesen Server gefunden; kein Upload gestartet.'
                $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $true; Phase = 'Keine Updates'; Detail = 'Keine eindeutig passende Firmwarekomponente in der gewaehlten Quelle' })
                return
            }
            $total = $components.Count
            $idx = 0
            & $log "Nach Hardware-/Versionsfilter (iLO immer zuletzt): $total Firmwarepaket(e) zur Pruefung: $((@($components | ForEach-Object { [System.IO.Path]::GetFileName($_) })) -join ', ')"
            $ambiguousMap = Get-AmbiguousFirmwareMatches -Inventory $inv -Components $components -KindFilter $kindFilter
            if ($ambiguousMap.Count) {
                & $log "Sicherheitsfilter: $($ambiguousMap.Count) Inventar-Komponente(n) passen zu mehreren Firmwarepaketen; diese Pakete werden nicht installiert."
            }

            $okNames = @(); $skipNames = @(); $absentNames = @(); $stagedNames = @(); $powerOffNames = @(); $bmcQueuedNames = @(); $downgradeNames = @(); $failItems = @()
            $rebootNeeded = $false
            $powerOffNeeded = $false

            foreach ($comp in $components) {
                $idx++
                $name = [System.IO.Path]::GetFileName($comp)
                $sizeMb = [Math]::Round(((Get-Item -LiteralPath $comp).Length / 1MB), 1)
                $kind = Get-ComponentKind -FileName $name
                $queueByBmc = $false
                $repositoryReady = $false

                # .fwpkg-Metadaten lesen: Version sowie (falls vorhanden) verbindliche
                # Angaben zum Flash-Weg, Update-Agenten und Geraetebezeichnung.
                # Letztere wird zwingend gegen das Live-Inventar dieses Servers
                # abgeglichen, damit keine nicht verbaute Komponente hochgeladen wird.
                $meta = Get-FwpkgMeta -FilePath $comp
                $tgtVer = $meta.Version
                $pendingMarkerKind = "Pending::$name"
                if ($meta.PowerOff) {
                    & $log "Aktivierungsanforderung '$name': Paket-JSON meldet ServerPowerOff=true; ein normaler Reboot reicht nicht, vollstaendiges Power Off erforderlich."
                }
                if ($strictNoDowngrade -and (-not $tgtVer -or $meta.VersionSource -ne 'JSON')) {
                    & $log "SICHERHEITSSPERRE $idx/$total : '$name' - Synergy-Modus verlangt eine Zielversion direkt aus dem Paket-JSON. Kein Upload."
                    $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "JSON-Zielversion fehlt $idx/$total : $name" })
                    $failItems += "$name (keine JSON-Zielversion)"
                    continue
                }
                if ($kind.Kind -in @('Other','SPS','CPLD','IE','PMC') -and -not @($meta.DeviceIdentifiers).Count) {
                    & $log "SICHERHEITSSPERRE $idx/$total : '$name' - Paket-JSON enthaelt keine eindeutigen Target/Targets fuer den Abgleich mit FirmwareInventory.Oem.Hpe.Targets. Kein Upload."
                    $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Target-ID fehlt $idx/$total : $name" })
                    $absentNames += $name
                    continue
                }
                if ($meta.RouteConflict -or $null -eq $meta.Deferred) {
                    $routeProblem = if ($meta.RouteConflict) { "widerspruechliche Angaben: $($meta.RouteSources -join ', ')" } else { 'DirectFlashOk/UefiFlashable fehlt' }
                    & $log "SICHERHEITSSPERRE $idx/$total : '$name' - JSON-Flashweg ist nicht eindeutig ($routeProblem). Kein Upload."
                    $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "JSON-Routing unklar $idx/$total : $name" })
                    $failItems += "$name (JSON-Flashweg unklar)"
                    continue
                }
                $kind.Deferred = [bool]$meta.Deferred
                $kind.UpdatableBy = @($meta.UpdatableBy | Where-Object { $_ })
                if ($kind.Deferred -and -not @($kind.UpdatableBy | Where-Object { "$_" -match '^(?i:Bmc|Uefi|RuntimeAgent)$' }).Count) {
                    & $log "SICHERHEITSSPERRE $idx/$total : '$name' - JSON verlangt Repository/Queue, enthaelt aber keinen gueltigen UpdatableBy-Agenten. Kein Upload."
                    $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Queue-Agent fehlt $idx/$total : $name" })
                    $failItems += "$name (JSON-Queue-Agent fehlt)"
                    continue
                }
                $jsonRoute = if ($kind.Deferred) { 'ComponentRepository + UpdateTaskQueue' } else { 'Direct-Flash' }
                & $log "JSON-Update-Route '$name': $jsonRoute; Angaben: $($meta.RouteSources -join ', ')$(if($kind.UpdatableBy.Count){"; UpdatableBy=$($kind.UpdatableBy -join '+')"})"
                # Nicht verbaute Hardware ueberspringen. BIOS, iLO und SPS haben
                # feste Inventarmuster; alle weiteren Komponenten werden mit den
                # aus der .fwpkg gelesenen Geraetebezeichnungen abgeglichen.
                $instMatches = @(ILO-FindInstalledFirmware -Inventory $inv -InvPattern $kind.InvPattern -InventoryTerms $meta.InventoryTerms -DeviceIdentifiers $meta.DeviceIdentifiers -DeviceClasses $meta.DeviceClasses)
                $instItem = $instMatches | Select-Object -First 1
                if (-not $instItem) {
                    $matchInfo = if ($kind.InvPattern) { "Muster '$($kind.InvPattern)'" } elseif (@($meta.InventoryTerms).Count) { "Paketkennungen '$($meta.InventoryTerms -join ' | ')'" } else { 'keine auswertbare Geraetekennung in den Paketmetadaten' }
                    & $log "Uebersprungen $idx/$total : '$name' - Hardware nicht verbaut oder nicht eindeutig im Server-Inventar gefunden ($matchInfo)."
                    $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Nicht verbaut $idx/$total : $name" })
                    $absentNames += $name
                    continue
                }
                if (@($instMatches | Where-Object { $ambiguousMap.ContainsKey("$($_.Name)") }).Count) {
                    & $log "Uebersprungen $idx/$total : '$name' - uneindeutige Paketzuordnung fuer Inventar '$($instItem.Name)' (mehrere Firmwaredateien passen)."
                    $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Uneindeutig $idx/$total : $name" })
                    $absentNames += $name
                    continue
                }
                $notUpdateableMatches = @($instMatches | Where-Object { $null -ne $_.Updateable -and -not [bool]$_.Updateable })
                $allNotUpdateable = $instMatches.Count -and $notUpdateableMatches.Count -eq $instMatches.Count
                $queueRoute = [bool]$meta.Deferred
                $queueByBmc = $queueRoute -and
                    (@($kind.UpdatableBy | Where-Object { "$_" -ieq 'Bmc' }).Count -gt 0)
                if ($queueRoute) {
                    & $log "Update-Route '$name': JSON verlangt ComponentRepository + UpdateTaskQueue mit Paket-Agent(en) $($kind.UpdatableBy -join '+')."
                } elseif ($allNotUpdateable) {
                    $invTxt = (@($notUpdateableMatches | ForEach-Object { "$($_.Name)=$($_.Version)" }) -join ', ')
                    & $log "Uebersprungen $idx/$total : '$name' - JSON verlangt Direct-Flash, aber iLO meldet fuer [$invTxt] Updateable=false. Der Flash-Weg wird nicht auf Queue umgedeutet."
                    $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Direct-Flash nicht freigegeben $idx/$total : $name" })
                    $skipNames += $name
                    continue
                } else {
                    & $log "Update-Route '$name': JSON verlangt Direct-Flash."
                }

                # Versionsvergleich: bereits aktuelle Komponente ueberspringen.
                # WICHTIG (BIOS): es kann MEHRERE passende Inventar-Eintraege geben
                # (z.B. 'System ROM' = aktiv UND 'Redundant System ROM' = Backup).
                # Nach einem Flash liegt die neue Version zunaechst NUR im Redundant
                # ROM und wird erst beim Reboot aktiv. Wir ueberspringen daher, wenn
                # die Zielversion in IRGENDEINEM passenden Eintrag bereits vorliegt -
                # sonst wuerde bei jedem Lauf erneut (unnoetig) geflasht.
                if ($tgtVer -and $instMatches.Count) {
                    $tn = Get-NormFwVersion $tgtVer
                    $hit = $null
                    if ($tn) { $hit = $instMatches | Where-Object { (Compare-FwVersion $_.Version $tgtVer) -eq 0 } | Select-Object -First 1 }
                    # Recovery-Fall: Ein frueherer Lauf hat ein aelteres BIOS in
                    # das redundante ROM geschrieben, waehrend die neuere Version
                    # noch aktiv ist. Dann muss ein Paket der aktiven Version
                    # erneut geflasht werden duerfen, um das redundante ROM zu
                    # reparieren; die normale Gleichheitspruefung wuerde es sonst
                    # faelschlich als "bereits aktuell" ueberspringen.
                    $repairStagedBiosDowngrade = $false
                    $stagedBios = $null
                    if (-not $strictNoDowngrade -and $kind.Kind -eq 'ROM' -and $tn) {
                        $stagedBios = Get-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $kind.Kind
                        $currentNorm = Get-NormFwVersion $info.BiosCurrent
                        if ($stagedBios -and $currentNorm -and $tn -eq $currentNorm) {
                            $stagedCmp = Compare-FwVersion $stagedBios.TargetNorm $currentNorm
                            $repairStagedBiosDowngrade = ($null -ne $stagedCmp -and $stagedCmp -lt 0)
                        }
                    }
                    if ($hit -and -not $repairStagedBiosDowngrade) {
                        & $log "Uebersprungen $idx/$total : '$name' - Version $tgtVer bereits vorhanden (Inventar '$($hit.Name)' = $($hit.Version))"
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Aktuell $idx/$total : $name" })
                        # Zielversion ist aktiv im Inventar -> evtl. gesetzter
                        # Staged-Marker (Reboot erfolgt) kann entfernt werden.
                        if (-not $kind.Deferred -and -not $kind.IsIlo) { Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $kind.Kind }
                        if ($queueByBmc) { Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name" }
                        Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $pendingMarkerKind
                        $skipNames += $name
                        continue
                    }
                    if ($repairStagedBiosDowngrade) {
                        & $log "BIOS-RECOVERY $idx/$total : '$name' - aktives ROM ist '$($info.BiosCurrent)', aber der lokale Staging-Marker zeigt das aeltere redundante ROM '$($stagedBios.TargetNorm)'. Version '$tgtVer' wird erneut geflasht, um das redundante ROM zu reparieren."
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "BIOS-Recovery $idx/$total : $name" })
                    }

                    # Verbindlicher Downgrade-Schutz im Flash-Worker. Die Vorschau
                    # allein ist keine Sicherheitsgrenze, da ein Update auch ohne
                    # vorherige Inventarpruefung gestartet werden kann. Verglichen
                    # wird mit ALLEN passenden Inventareintraegen; beim BIOS kommen
                    # aktive und redundante ROM-Version aus ComputerSystem hinzu.
                    $rawInstalledVersions = @($instMatches | ForEach-Object { "$($_.Version)".Trim() })
                    if ($kind.Kind -eq 'ROM') {
                        if ($info.BiosCurrent) { $rawInstalledVersions += "$($info.BiosCurrent)".Trim() }
                        if ($info.BiosBackup)  { $rawInstalledVersions += "$($info.BiosBackup)".Trim() }
                    }
                    $installedVersions = @($rawInstalledVersions | Where-Object { $_ })
                    if ($strictNoDowngrade) {
                        $missingVersion = $rawInstalledVersions.Count -ne $installedVersions.Count
                        $unverifiableVersions = @($installedVersions | Where-Object {
                            $null -eq (Compare-FwVersionStrict $tgtVer $_)
                        })
                        if ($missingVersion -or -not $installedVersions.Count -or $unverifiableVersions.Count) {
                            $installedText = if ($installedVersions.Count) { $installedVersions -join ', ' } else { '(keine lesbare installierte Version)' }
                            & $log "SICHERHEITSSPERRE $idx/$total : '$name' - Ziel '$tgtVer' ist nicht eindeutig numerisch mit [$installedText] vergleichbar. Im Synergy-Modus wird nur eine sicher neuere Version geflasht."
                            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Version nicht vergleichbar $idx/$total : $name" })
                            $downgradeNames += "$name (Versionsvergleich nicht eindeutig)"
                            continue
                        }
                    }
                    $newerInstalled = @($installedVersions | Where-Object {
                        $cmp = if ($strictNoDowngrade) {
                            Compare-FwVersionStrict $tgtVer $_
                        } else {
                            Compare-FwVersion $tgtVer $_
                        }
                        $null -ne $cmp -and $cmp -lt 0
                    })
                    if ($newerInstalled.Count) {
                        $highestInstalled = $newerInstalled[0]
                        foreach ($candidateVersion in $newerInstalled) {
                            $cmpHighest = Compare-FwVersion $candidateVersion $highestInstalled
                            if ($null -ne $cmpHighest -and $cmpHighest -gt 0) { $highestInstalled = $candidateVersion }
                        }
                        & $log "SICHERHEITSSPERRE $idx/$total : '$name' - Downgrade von installiert '$highestInstalled' auf Ziel '$tgtVer' wird NICHT durchgefuehrt."
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Downgrade gesperrt $idx/$total : $name" })
                        $downgradeNames += "$name ($highestInstalled -> $tgtVer)"
                        continue
                    }

                    $invTxt = ($instMatches | ForEach-Object { "$($_.Name)=$($_.Version)" }) -join ', '
                    & $log "Versionsvergleich '$name': Ziel='$tgtVer' (norm '$tn') nicht in [$invTxt] -> wird geflasht."
                } else {
                    $whyNo = if (-not $instMatches.Count) { "kein Inventar-Eintrag fuer [$($kind.Kind)] (Muster '$($kind.InvPattern)')" } else { "keine Zielversion aus fwpkg lesbar" }
                    & $log "Versionspruefung '$name' nicht moeglich: $whyNo -> wird geflasht."
                }

                # iLO-NATIVE Erkennung fuer System ROM (BIOS): Die ComputerSystem-
                # Ressource (Oem.Hpe.Bios) fuehrt die aktive (Current) UND die
                # Backup/Redundant-ROM-Version. Anders als das FirmwareInventory
                # zeigt 'Backup' unmittelbar nach einem Online-Flash bereits die
                # frisch geflashte, noch nicht per Reboot aktivierte Version (genau
                # wie die iLO-Overview). Damit wird ein zweiter Lauf VOR dem Reboot
                # zuverlaessig erkannt, ohne auf einen lokalen Marker angewiesen zu
                # sein. Stimmt die Zielversion mit Current -> bereits aktiv; mit
                # Backup -> bereits geflasht, Aktivierung beim Reboot ausstehend.
                if ($kind.Kind -eq 'ROM' -and $tgtVer) {
                    $tnRom = Get-NormFwVersion $tgtVer
                    $curN  = Get-NormFwVersion $info.BiosCurrent
                    $bakN  = Get-NormFwVersion $info.BiosBackup
                    if ($tnRom -and $curN -and $tnRom -eq $curN -and -not $repairStagedBiosDowngrade) {
                        & $log "Uebersprungen $idx/$total : '$name' - Version $tgtVer bereits aktiv (System ROM = $($info.BiosCurrent))"
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Aktuell $idx/$total : $name" })
                        if (-not $kind.IsIlo) { Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $kind.Kind }
                        $skipNames += $name
                        continue
                    }
                    if ($tnRom -and $bakN -and $tnRom -eq $bakN) {
                        & $log "Uebersprungen $idx/$total : '$name' - Version $tgtVer bereits in Redundant System ROM geflasht (Backup = $($info.BiosBackup)), Aktivierung beim Reboot ausstehend."
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name (Reboot)" })
                        $stagedNames += $name
                        $rebootNeeded = $true
                        continue
                    }
                }

                # HPE-serverseitige Pending-Erkennung fuer alle Repository-/Queue-
                # Komponenten. Repository, TaskQueue und aktives Inventar werden
                # gemeinsam ausgewertet. 'New' darf hochgeladen werden;
                # 'RepositoryReady' verwendet das bereits vorhandene, exakt
                # passende Paket ohne erneuten Upload. Alle anderen Zustaende
                # bleiben fail-closed.
                if ($kind.Deferred) {
                    try {
                        $serverDisposition = ILO-GetFirmwareServerDisposition -Ilo $ilo -Token $token -ComponentFileName $name -TargetVersion $tgtVer -InventoryMatches $instMatches
                    } catch {
                        $serverStateError = "HPE-Updatezustand konnte nicht sicher gelesen werden: $($_.Exception.Message)"
                        & $log "ABBRUCH bei '$name': $serverStateError. Kein weiterer Upload/Task."
                        $failItems += "$name ($serverStateError)"
                        break
                    }
                    & $log "HPE-Serverzustand '$name': Status='$($serverDisposition.Status)' - $($serverDisposition.Reason)"

                    if ($serverDisposition.Status -eq 'RepositoryReady') {
                        $repositoryReady = $true
                        & $log "Repository-Wiederverwendung '$name': kein erneuter Upload; vorhandene, exakt passende Repository-Komponente wird in die Queue gestellt."
                    }
                    if ($serverDisposition.Status -eq 'NotApplicable') {
                        if ($serverDisposition.Task -and $serverDisposition.Task.Uri) {
                            try {
                                ILO-DeleteUpdateTask -Ilo $ilo -Token $token -TaskUri "$($serverDisposition.Task.Uri)" -LogCb $log
                            } catch {
                                $cleanupError = "Nicht anwendbarer Task konnte nicht sicher entfernt werden: $($_.Exception.Message)"
                                & $log "ABBRUCH bei '$name': $cleanupError"
                                $failItems += "$name ($cleanupError)"
                                break
                            }
                        }
                        $tnNotApplicable = Get-NormFwVersion $tgtVer
                        if ($tnNotApplicable) {
                            Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name" -File $name -TargetNorm $tnNotApplicable -State 'NoMatchingTarget'
                        }
                        & $log "Uebersprungen $idx/$total : '$name' - iLO bestaetigt, dass auf diesem Server kein kompatibles Zielgeraet vorhanden ist. Kein erneuter Versuch."
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Nicht anwendbar $idx/$total : $name" })
                        $skipNames += "$name (kein Zielgeraet)"
                        continue
                    }
                    if ($serverDisposition.Status -eq 'Active') {
                        Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name"
                        Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $pendingMarkerKind
                        & $log "Uebersprungen $idx/$total : '$name' - Zielversion ist serverseitig als aktiv bestaetigt."
                        $skipNames += $name
                        continue
                    }
                    if ($serverDisposition.Status -in @('Pending','PendingActivation')) {
                        $activation = "$($serverDisposition.Component.Activates)"
                        $tnServerPending = Get-NormFwVersion $tgtVer
                        if ($tnServerPending) {
                            Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $pendingMarkerKind -File $name -TargetNorm $tnServerPending -State $serverDisposition.Status
                        }
                        if ($queueByBmc -and $tnServerPending) {
                            Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name" -File $name -TargetNorm $tnServerPending -State $serverDisposition.Status
                        }
                        & $log "Uebersprungen $idx/$total : '$name' - iLO bestaetigt bereits '$($serverDisposition.Status)' (Activates='$activation'). Kein erneuter Upload/Task."
                        if ($meta.PowerOff) {
                            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name (Power Off)" })
                            $powerOffNames += $name
                            $powerOffNeeded = $true
                        } else {
                            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name (iLO)" })
                            $stagedNames += $name
                            if ($activation -notmatch '^Immediately$') { $rebootNeeded = $true }
                        }
                        continue
                    }
                    if ($serverDisposition.Status -in @('InProgress','QueueBusy')) {
                        & $log "STOPP: '$name' bzw. ein vorheriger iLO-Task wird noch verarbeitet ($($serverDisposition.Reason)). Keine Folgekomponente wird gestartet."
                        $stagedNames += $name
                        break
                    }
                    if ($serverDisposition.Status -in @('Blocked','QueueBlocked')) {
                        & $log "SICHERHEITSSPERRE $idx/$total : '$name' - $($serverDisposition.Reason). Kein automatischer Wiederholungsversuch."
                        $failItems += "$name ($($serverDisposition.Reason))"
                        if ($serverDisposition.Status -eq 'QueueBlocked') { break }
                        continue
                    }
                    if ($serverDisposition.Status -notin @('New','RepositoryReady')) {
                        & $log "SICHERHEITSSPERRE $idx/$total : '$name' - unbekannter HPE-Serverstatus '$($serverDisposition.Status)'."
                        $failItems += "$name (unbekannter HPE-Serverstatus '$($serverDisposition.Status)')"
                        break
                    }
                }

                # Paketgenauer Fallback fuer erfolgreich geflashte oder sicher in
                # die Queue gestellte Firmware, deren aktive Version erst nach dem
                # Reboot im Inventar erscheint. Der Marker ist absichtlich nach
                # Dateiname getrennt; ein grober Kind='Other'-Marker wuerde bei
                # mehreren Komponenten im selben Lauf gegenseitig ueberschrieben.
                if (-not $kind.IsIlo -and $tgtVer) {
                    $tnStg = Get-NormFwVersion $tgtVer
                    $pendingMarker = Get-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $pendingMarkerKind
                    if ($tnStg -and $pendingMarker -and "$($pendingMarker.TargetNorm)" -eq $tnStg) {
                        $markerNeedsPowerOff = "$($pendingMarker.State)" -eq 'PendingPowerOff'
                        if (-not $markerNeedsPowerOff -and (Test-RebootAfterStagedFlash -Marker $pendingMarker -LastResetTime $info.LastResetTime)) {
                            & $log "Pending-Marker fuer '$name' stammt von vor dem letzten Server-Reboot; aktive Version ist weiterhin abweichend. Marker wird entfernt und das Update darf erneut geprueft werden."
                            Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $pendingMarkerKind
                        } elseif ("$($pendingMarker.State)" -match '^(PendingPowerOff|PendingReboot|Pending|PendingActivation|Queued|Flashed|Complete)$') {
                            if ($markerNeedsPowerOff -or $meta.PowerOff) {
                                & $log "Uebersprungen $idx/$total : '$name' - Version $tgtVer wurde bereits erfolgreich geflasht; Aktivierung durch vollstaendiges Power Off ausstehend."
                                $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name (Power Off)" })
                                $powerOffNames += $name
                                $powerOffNeeded = $true
                            } else {
                                & $log "Uebersprungen $idx/$total : '$name' - Version $tgtVer wurde bereits erfolgreich geflasht/eingeplant; Aktivierung beim Reboot ausstehend."
                                $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name (Reboot)" })
                                $stagedNames += $name
                                $rebootNeeded = $true
                            }
                            continue
                        }
                    }
                }

                # Nur ein zuvor serverseitig bestaetigter Complete-Task, dessen neue
                # Version erst nach dem Reboot sichtbar wird, darf lokal gegen einen
                # Doppel-POST schuetzen. Unklare Submitting/Submitted-Marker werden
                # durch den aktuellen iLO-Zustand ersetzt und blockieren kein Update.
                if ($queueByBmc -and $tgtVer) {
                    $tnQueue = Get-NormFwVersion $tgtVer
                    $queueMarker = Get-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name"
                    if ($tnQueue -and $queueMarker -and "$($queueMarker.TargetNorm)" -eq $tnQueue) {
                        $markerState = if ($queueMarker.State) { "$($queueMarker.State)" } else { 'PendingReboot' }
                        if ($markerState -in @('PendingReboot','PendingPowerOff')) {
                            if ($markerState -eq 'PendingPowerOff' -or $meta.PowerOff) {
                                & $log "Uebersprungen $idx/$total : '$name' - ein vorheriger iLO-Task wurde bereits als Complete bestaetigt; Aktivierung durch vollstaendiges Power Off ausstehend. Kein doppelter Task."
                                $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name (Power Off)" })
                                $powerOffNames += $name
                                $powerOffNeeded = $true
                            } else {
                                & $log "Uebersprungen $idx/$total : '$name' - ein vorheriger iLO-Task wurde bereits als Complete bestaetigt; Aktivierung/Inventarwechsel nach Reboot ausstehend. Kein doppelter Task."
                                $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name (Reboot)" })
                                $stagedNames += $name
                                $rebootNeeded = $true
                            }
                            continue
                        }
                        & $log "Lokaler, nicht abgeschlossener Marker '$markerState' fuer '$name' wird durch den aktuellen iLO-Status '$($serverDisposition.Status)' ersetzt und entfernt."
                        Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name"
                    }
                }

                try {
                    # Vor JEDEM Upload/Task auf einen stabil freien UpdateService
                    # warten. So kann kein Folgepaket starten, solange der vorherige
                    # Task global noch Verifying/Updating/Writing ausfuehrt.
                    $st = ILO-WaitForUpdateServiceReady -Ilo $ilo -Token $token -TimeoutSec 2400 -LogCb $log
                    $stateDetail = if ($st.Detail) { ", Detail='$($st.Detail)'" } else { '' }
                    & $log "UpdateService: State='$($st.State)', PushUri='$($st.PushUri)'$stateDetail"
                    $pushUri = $st.PushUri

                    $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Upload $idx/$total : $name" })
                    $modeTxt = if ($kind.Deferred) { "Repository/deferred [$($kind.Kind)]" } else { "direkt [$($kind.Kind)]" }
                    & $log "Upload $idx/$($total): '$name' ($sizeMb MB, $modeTxt) -> $pushUri"
                    $cb = { param($pct) $uiQueue.Enqueue(@{ Type = 'PROGRESS'; Ilo = $ilo; Percent = $pct }) }.GetNewClosure()

                    if ($kind.Deferred) {
                        # WICHTIG bei WIEDERHOLTEM Lauf: Liegt die Komponente bereits
                        # im iLO-Repository (vom letzten, noch nicht per Reboot
                        # aktivierten Update), wuerde ein erneuter Repository-Upload
                        # derselben Datei den UpdateService in den Zustand 'Error'
                        # versetzen ("liegt im Repository, aber Fehler"). Daher zuerst
                        # pruefen und den Upload ueberspringen; der Queue-Task wird
                        # anschliessend ohnehin (idempotent) sichergestellt.
                        $existing = ILO-FindRepositoryComponent -Ilo $ilo -Token $token -ComponentFileName $name
                        if ($existing) {
                            & $log "Repository $idx/$total : '$name' liegt bereits im iLO-Repository (vom vorherigen Lauf, noch nicht aktiviert) - kein erneuter Upload."
                            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Repository $idx/$total : $name (vorhanden)" })
                        } else {
                            # SPS/CPLD/IE: NUR ins iLO-Repository hochladen (UpdateTarget=false).
                            # Sofortiges Flashen (UpdateTarget=true) wuerde 'Error' liefern.
                            ILO-UploadComponent -Ilo $ilo -Token $token -PushUri $pushUri -FilePath $comp -UpdateRepository:$true -UpdateTarget:$false -ProgressCb $cb | Out-Null
                            & $log "Upload $idx/$total abgeschlossen - schreibe ins iLO-Repository..."
                            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Repository $idx/$total : $name" })
                            $rcb = { param($pct, $state) $uiQueue.Enqueue(@{ Type = 'PROGRESS'; Ilo = $ilo; Percent = $pct }) }.GetNewClosure()
                            $rState = ILO-WaitForRepository -Ilo $ilo -Token $token -TimeoutSec 600 -ProgressCb $rcb
                            & $log "Repository-Upload $idx/$total fertig: '$name' (State='$rState')."
                        }

                        # Verbindliche iLO-native Anwendbarkeitspruefung NACH dem
                        # Repository-Import und VOR ApplyUpdate. Erst jetzt stehen
                        # die von iLO verifizierten Repository-Targets zur
                        # Verfuegung. Ohne exakte Uebereinstimmung mit dem
                        # Live-Inventar darf kein Queue-Task entstehen.
                        $repoComponent = ILO-FindRepositoryComponent -Ilo $ilo -Token $token -ComponentFileName $name -Strict
                        $repoApplicability = Test-RepositoryComponentApplicable -Component $repoComponent -InventoryMatches $instMatches -TargetVersion $tgtVer
                        if (-not $repoApplicability.Applicable) {
                            $tnRepoMismatch = Get-NormFwVersion $tgtVer
                            if ($queueByBmc -and $tnRepoMismatch) {
                                Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name" -File $name -TargetNorm $tnRepoMismatch -State 'RepositoryTargetMismatch'
                            }
                            & $log "SICHERHEITSSPERRE $idx/$total : '$name' - $($repoApplicability.Reason). Kein ApplyUpdate-Task; kein Device wird angesprochen."
                            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Nicht kompatibel $idx/$total : $name" })
                            $failItems += "$name (Repository-Targets passen nicht zum Live-Inventar)"
                            continue
                        }
                        & $log "iLO-Anwendbarkeit bestaetigt fuer '$name': $($repoApplicability.Reason)."

                        $inventoryTargetUris = @($instMatches | ForEach-Object { "$($_.Uri)".Trim() } | Where-Object {
                            $_ -match '^/redfish/v1/UpdateService/FirmwareInventory/'
                        } | Select-Object -Unique)
                        if ($kind.Kind -in @('Other','SPS','CPLD','IE','PMC') -and -not $inventoryTargetUris.Count) {
                            & $log "SICHERHEITSSPERRE $idx/$total : '$name' - passendes FirmwareInventory-Target hat keine verwertbare @odata.id. Kein ApplyUpdate-Task."
                            $failItems += "$name (FirmwareInventory-Target-URI fehlt)"
                            continue
                        }
                        # HPE Component-Tasks aktualisieren alle fuer das Paket
                        # anwendbaren Ziele; die Task-API bietet keine Auswahl einer
                        # einzelnen FirmwareInventory-URI. Die exakte Schranke ist
                        # bereits durch Paket- und Repository-Targets gegen das
                        # Live-Inventar geprueft.
                        $taskTargets = @()

                        # Task in der UpdateTaskQueue anlegen/sicherstellen. ILO-CreateUpdateTask ist idempotent:
                        # existiert bereits ein Task fuer die Datei, wird dieser akzeptiert;
                        # sonst wird er angelegt (mit Verifikation + Payload-Fallbacks).
                        & $log "Stelle Update-Task in der Queue sicher fuer '$name'..."
                        # Marker VOR dem ersten moeglichen POST setzen. Selbst bei
                        # Prozessabbruch zwischen POST und Statusauswertung darf ein
                        # spaeterer Lauf keinen zweiten Task erzeugen.
                        $tnQueueSubmitting = Get-NormFwVersion $tgtVer
                        if ($queueByBmc -and $tnQueueSubmitting) {
                            Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name" -File $name -TargetNorm $tnQueueSubmitting -State 'Submitting'
                        }
                        try {
                            $taskUri = ILO-CreateUpdateTask -Ilo $ilo -Token $token -ComponentFileName $name -UpdatableBy $kind.UpdatableBy -Targets $taskTargets -LogCb $log
                        } catch {
                            # Eindeutig abgelehnte POST-Varianten haben keinen Task
                            # erzeugt; ihr vorsorglicher Submitting-Marker darf den
                            # korrigierten naechsten Lauf nicht blockieren. Nur bei
                            # unklarem HTTP-Ausgang bleibt der Absturzschutz bestehen.
                            if ($queueByBmc -and $tnQueueSubmitting -and
                                -not [bool]$_.Exception.Data['KeepSubmittingMarker']) {
                                Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name"
                            }
                            throw
                        }
                        if ($queueByBmc -and $tnQueueSubmitting) {
                            Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name" -File $name -TargetNorm $tnQueueSubmitting -State 'Submitted'
                        }
                        & $log "Update-Task in Queue bestaetigt fuer '$name' (UpdatableBy=$(($kind.UpdatableBy) -join '+'))$(if($taskUri){" -> $taskUri"})"
                        $qPendingReboot = $false
                        if ($queueByBmc) {
                            & $log "Bmc-Queue fuer '$name' ist angelegt; iLO verarbeitet diesen Task bei eingeschaltetem Server sofort."
                            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "BMC-Queue $idx/$total : $name" })
                            $qcb = { param($pct, $state) if ($null -ne $pct) { $uiQueue.Enqueue(@{ Type = 'PROGRESS'; Ilo = $ilo; Percent = $pct }) } }.GetNewClosure()
                            $qres = ILO-WaitForFirmwareInventoryVersion -Ilo $ilo -Token $token -TargetVersion $tgtVer -InvPattern $kind.InvPattern -InventoryTerms $meta.InventoryTerms -DeviceIdentifiers $meta.DeviceIdentifiers -DeviceClasses $meta.DeviceClasses -TaskUri $taskUri -ComponentFileName $name -TimeoutSec 2400 -ProgressCb $qcb -LogCb $log
                            if ($qres.Status -eq 'Confirmed') {
                                Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name"
                                Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $pendingMarkerKind
                                & $log "Bmc-Queue abgeschlossen: '$name' im Firmware-Inventar bestaetigt ($($qres.Inventory.Name) = $($qres.Inventory.Version))."
                            } elseif ($qres.Status -eq 'PendingReboot') {
                                $qPendingReboot = $true
                                $tnQueueDone = Get-NormFwVersion $tgtVer
                                $pendingState = if ($meta.PowerOff) { 'PendingPowerOff' } else { 'PendingReboot' }
                                if ($tnQueueDone) {
                                    Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name" -File $name -TargetNorm $tnQueueDone -State $pendingState
                                    Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $pendingMarkerKind -File $name -TargetNorm $tnQueueDone -State $pendingState
                                }
                                if ($meta.PowerOff) {
                                    & $log "Bmc-Queue-Task fuer '$name' meldet '$($qres.TaskState)'; die neue Version erscheint erst nach vollstaendigem Power Off im Firmware-Inventar. Fahre mit der naechsten Komponente fort."
                                    $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name (Power Off)" })
                                } else {
                                    & $log "Bmc-Queue-Task fuer '$name' meldet '$($qres.TaskState)'; die neue Version erscheint erst nach dem Reboot im Firmware-Inventar. Fahre mit der naechsten Komponente fort."
                                    $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name (Reboot)" })
                                }
                            } elseif ($qres.Status -eq 'NotApplicable') {
                                ILO-DeleteUpdateTask -Ilo $ilo -Token $token -TaskUri $taskUri -LogCb $log
                                $tnNoTarget = Get-NormFwVersion $tgtVer
                                if ($tnNoTarget) {
                                    Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind "BmcQueue::$name" -File $name -TargetNorm $tnNoTarget -State 'NoMatchingTarget'
                                }
                                & $log "Uebersprungen $idx/$total : '$name' - iLO meldet NoMatchingTargetDeviceFound. Der Exception-Task wurde entfernt; kein erneuter Versuch auf diesem Server."
                                $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Nicht anwendbar $idx/$total : $name" })
                                $skipNames += "$name (kein Zielgeraet)"
                                continue
                            } else {
                                throw "Unbekannter Rueckgabestatus der BMC-Queue-Ueberwachung: '$($qres.Status)'"
                            }
                        } else {
                            $tnQueued = Get-NormFwVersion $tgtVer
                            if ($tnQueued) {
                                Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $pendingMarkerKind -File $name -TargetNorm $tnQueued -State 'Queued'
                            }
                            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name (Reboot)" })
                        }
                    } else {
                        # Direkt online flashen. iLO verarbeitet den signierten
                        # HPE-.fwpkg-Container und waehlt die interne Payload selbst.
                        ILO-UploadComponent -Ilo $ilo -Token $token -PushUri $pushUri -FilePath $comp -UpdateRepository:$false -UpdateTarget:$true -ProgressCb $cb | Out-Null
                        & $log "Upload $idx/$total abgeschlossen, starte Flash..."
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Flash $idx/$total : $name" })
                        $fcb = { param($pct, $state) $uiQueue.Enqueue(@{ Type = 'PROGRESS'; Ilo = $ilo; Percent = $pct }) }.GetNewClosure()
                        $wf = ILO-WaitForFlash -Ilo $ilo -Token $token -User $user -Pass $pass -TimeoutSec 2400 -ProgressCb $fcb
                        $token = $wf.Token
                        & $log "Flash $idx/$total fertig: '$name' (End-State='$($wf.State)')"
                        # Direkt geflashte Nicht-iLO-Komponente paketgenau vormerken,
                        # wenn die Payload eine Aktivierung per Reboot verlangt.
                        if (-not $kind.IsIlo -and $tgtVer) {
                            $tnMk = Get-NormFwVersion $tgtVer
                            if ($tnMk -and ($meta.PowerOff -or $null -eq $meta.Reboot -or $meta.Reboot)) {
                                $activationState = if ($meta.PowerOff) { 'PendingPowerOff' } else { 'Flashed' }
                                Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $pendingMarkerKind -File $name -TargetNorm $tnMk -State $activationState
                                # Der ROM-Typmarker bleibt ausschliesslich fuer die
                                # spezielle BIOS-Downgrade-Recovery erhalten.
                                if ($kind.Kind -eq 'ROM') {
                                    Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $kind.Kind -File $name -TargetNorm $tnMk -State 'PendingReboot'
                                }
                            }
                        }
                    }

                    # iLO ist nach Selbst-Reset sofort aktiv; alles andere braucht einen Reboot.
                    if ($meta.PowerOff) {
                        $powerOffNames += $name
                    } elseif ($kind.Deferred) {
                        if ($queueByBmc -and -not $qPendingReboot) { $bmcQueuedNames += $name } else { $stagedNames += $name }
                    } else { $okNames += $name }
                    # Reboot-Bedarf: deferred-Komponenten und Nicht-iLO immer; bei
                    # direkt geflashten Nicht-iLO-Komponenten entscheidet - falls
                    # vorhanden - die fwpkg-Metadaten (RebootRequired).
                    if ($meta.PowerOff) {
                        $powerOffNeeded = $true
                    } elseif ($kind.Deferred -and (-not $queueByBmc -or $qPendingReboot)) {
                        $rebootNeeded = $true
                    } elseif (-not $kind.IsIlo) {
                        if ($null -eq $meta.Reboot -or $meta.Reboot) { $rebootNeeded = $true }
                    }
                }
                catch {
                    $cErr = $_.Exception.Message
                    $abortUpdateRun = $cErr -match '^iLO-UpdateService meldet State|iLO-UpdateService wurde nicht frei'
                    if ($abortUpdateRun) {
                        & $log "ABBRUCH bei '$name': $cErr - keine weiteren Firmwarepakete werden gestartet."
                    } else {
                        & $log "FEHLER bei '$name': $cErr - Komponente uebersprungen, weiter mit naechster."
                    }
                    $failItems += "$name ($cErr)"
                    if ($abortUpdateRun) { break }
                }

                # kurze Pause, bis State wieder Idle ist (vor naechster Komponente)
                if ($idx -lt $total) {
                    for ($w = 0; $w -lt 12; $w++) {
                        try { $s2 = ILO-GetUpdateState -Ilo $ilo -Token $token; if ($s2.State -match '^(Idle|Complete)$') { break } } catch {}
                        Start-Sleep -Seconds 5
                    }
                }
            }

            # Ergebnis zusammenfassen.
            $parts = @()
            if ($okNames.Count)     { $parts += "$($okNames.Count) verarbeitet" }
            if ($bmcQueuedNames.Count) { $parts += "$($bmcQueuedNames.Count) BMC-Queue (sofort)" }
            if ($stagedNames.Count) { $parts += "$($stagedNames.Count) vorgemerkt (Reboot)" }
            if ($powerOffNames.Count) { $parts += "$($powerOffNames.Count) vorgemerkt (Power Off)" }
            if ($skipNames.Count)   { $parts += "$($skipNames.Count) aktuell" }
            if ($downgradeNames.Count) { $parts += "$($downgradeNames.Count) Downgrade gesperrt" }
            if ($absentNames.Count) { $parts += "$($absentNames.Count) nicht verbaut" }
            if ($failItems.Count)   { $parts += "$($failItems.Count) Fehler" }
            $summary = if ($parts.Count) { ($parts -join ', ') } else { 'nichts zu tun' }
            $rebootTxt = if ($powerOffNeeded) { 'Power Off fuer Aktivierung erforderlich' } elseif ($rebootNeeded) { 'Reboot fuer Aktivierung noetig' } else { 'Kein Reboot noetig' }
            $success = ($failItems.Count -eq 0)
            $phase = if (-not $success) { 'Teilweise' } elseif (($okNames.Count + $stagedNames.Count + $powerOffNames.Count + $bmcQueuedNames.Count) -eq 0 -and $downgradeNames.Count) { 'Downgrade gesperrt' } elseif (($okNames.Count + $stagedNames.Count + $powerOffNames.Count + $bmcQueuedNames.Count) -eq 0) { 'Aktuell' } else { 'Fertig' }
            $detail = "$summary - $rebootTxt"
            if ($downgradeNames.Count) { $detail += " | Gesperrt: $($downgradeNames -join '; ')" }
            if ($failItems.Count) { $detail += " | Fehler: $($failItems -join '; ')" }
            & $log "Server fertig: $summary. $rebootTxt."
            $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $success; Phase = $phase; Detail = $detail })
        }
        catch {
            $errMsg = $_.Exception.Message
            $detailLine = $errMsg
            if ($_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber) {
                $detailLine = "$errMsg (Zeile $($_.InvocationInfo.ScriptLineNumber))"
            }
            & $log "FEHLER: $detailLine"
            if ($_.Exception.InnerException) { & $log "FEHLER (Detail): $($_.Exception.InnerException.Message)" }
            $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Fehler'; Detail = $errMsg })
        }
        finally {
            if ($sess) { try { ILO-Logout -Ilo $ilo -Token $token -SessionUri $sess.SessionUri; & $log "Logout OK" } catch { & $log "Logout-Fehler: $($_.Exception.Message)" } }
        }
    }

    Start-Batch -Servers $servers -Worker $worker -ExtraArgs @{
        components = $components; mode = $mode; baseDir = $fwPath; kindFilter = @(& $script:GetKindFilter)
        synergyMode = $script:IsSynergyMode; strictNoDowngrade = $script:IsSynergyMode
        packageMetaCache = $packageMetaCache
    } -MaxParallel ([int]$numPar.Value)
})

# ─────────────────────────────────────────
# Exit
# ─────────────────────────────────────────
$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Location = '1040,12'; $btnExit.Size = '75,24'; $btnExit.Text = "Exit"
$form.Controls.Add($btnExit)
$btnExit.Add_Click({ $form.Close() })

$form.Add_FormClosing({
    try { if ($script:watch) { $script:watch.Stop() } } catch {}
    try { if ($script:guiTimer) { $script:guiTimer.Stop() } } catch {}
    try { if ($script:pool) { $script:pool.Close(); $script:pool.Dispose() } } catch {}
})

# Initiales Laden: zuletzt aus OneView erzeugte Serverliste (falls vorhanden)
if (Test-Path -LiteralPath $serversFile) {
    Load-ServersFile -Path $serversFile
    Add-Log "Bereit. $serverCacheName geladen ($($script:allServers.Count) Server). Fuer aktuelle Daten 'Server aus OneView laden'." ([System.Drawing.Color]::DarkBlue)
} else {
    Add-Log "Bereit. Noch keine $serverCacheName - bitte 'Server aus OneView laden'. Log: $script:logFile"
}

[void]$form.ShowDialog()
