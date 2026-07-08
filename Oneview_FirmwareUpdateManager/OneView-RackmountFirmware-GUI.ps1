#Requires -Version 7.5
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
$script:logFile = Join-Path $script:logFolder ("RackmountFirmware_{0:yyyy-MM-dd_HHmmss}.log" -f (Get-Date))
function Write-LogFile { param([string]$Text)
    try { Add-Content -LiteralPath $script:logFile -Value ("{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $Text) -Encoding utf8 } catch {}
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
    @{ Model = $model; Gen = $gen; Serial = $serial; iLO = $iloGen; BiosCurrent = $biosCur; BiosBackup = $biosBak }
}

# Liest das Firmware-Inventar (Name -> Version). Liefert Array von @{ Name; Version }
function ILO-GetFirmwareInventory {
    param([string]$Ilo,[string]$Token)
    $inv = @()
    try {
        $col = ILO-Get -Ilo $Ilo -Token $Token -Path "/redfish/v1/UpdateService/FirmwareInventory/?`$expand=."
        foreach ($m in @($col.Members)) {
            $inv += @{ Name = "$($m.Name)"; Version = "$($m.Version)" }
        }
    } catch {
        # Fallback ohne $expand: Member-URIs einzeln laden
        try {
            $col = ILO-Get -Ilo $Ilo -Token $Token -Path "/redfish/v1/UpdateService/FirmwareInventory/"
            foreach ($ref in @($col.Members)) {
                $u = $ref.'@odata.id'
                if (-not $u) { continue }
                try { $m = ILO-Get -Ilo $Ilo -Token $Token -Path $u; $inv += @{ Name = "$($m.Name)"; Version = "$($m.Version)" } } catch {}
            }
        } catch {}
    }
    ,$inv
}

# Liefert @{ State=<HpeState>; Percent=<int>; PushUri=<Upload-URI> }
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
    @{ State = $state; Percent = $pct; PushUri = $pushUri }
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

# Sucht im iLO ComponentRepository nach einem Eintrag, dessen Filename zu
# $ComponentFileName passt (exakt, sonst per Teilstring). Liefert das gematchte
# Member-Objekt (mit .Filename und @odata.id) oder $null. Nutzt $expand, faellt
# sonst auf Einzel-Abruf der Member-URIs zurueck. Wird genutzt, um bei einem
# WIEDERHOLTEN Lauf zu erkennen, dass eine deferred-Komponente bereits im
# Repository liegt (erneuter Upload wuerde den UpdateService in 'Error' versetzen).
function ILO-FindRepositoryComponent {
    param([string]$Ilo,[string]$Token,[string]$ComponentFileName)
    foreach ($repoPath in @('/redfish/v1/UpdateService/ComponentRepository/?$expand=.', '/redfish/v1/UpdateService/ComponentRepository/')) {
        try {
            $repo = ILO-Get -Ilo $Ilo -Token $Token -Path $repoPath
            $members = @($repo.Members)
            if ($members.Count -and -not ($members[0].PSObject.Properties.Name -contains 'Filename')) {
                $full = @()
                foreach ($ref in $members) { $u = $ref.'@odata.id'; if ($u) { try { $full += ILO-Get -Ilo $Ilo -Token $Token -Path $u } catch {} } }
                $members = $full
            }
            $match = $members | Where-Object { "$($_.Filename)" -ieq $ComponentFileName } | Select-Object -First 1
            if (-not $match) { $match = $members | Where-Object { "$($_.Filename)" -like "*$ComponentFileName*" -or $ComponentFileName -like "*$($_.Filename)*" } | Select-Object -First 1 }
            if ($match) { return $match }
        } catch {}
    }
    return $null
}

# Liest die iLO UpdateTaskQueue aus und liefert je Task ein Objekt
# @{ Name; Filename; Id; Uri; State }. Wird zur Verifikation genutzt, dass ein
# angelegter Task tatsaechlich in der Queue liegt.
function ILO-GetUpdateTaskQueue {
    param([string]$Ilo,[string]$Token,[string]$QueueUri = $null)
    if ([string]::IsNullOrWhiteSpace($QueueUri)) {
        $QueueUri = '/redfish/v1/UpdateService/UpdateTaskQueue'
        try {
            $us = ILO-Get -Ilo $Ilo -Token $Token -Path '/redfish/v1/UpdateService'
            if ($us.Oem -and $us.Oem.Hpe -and $us.Oem.Hpe.UpdateTaskQueue -and $us.Oem.Hpe.UpdateTaskQueue.'@odata.id') {
                $QueueUri = "$($us.Oem.Hpe.UpdateTaskQueue.'@odata.id')"
            }
        } catch {}
    }
    $tasks = @()
    try {
        $q = ILO-Get -Ilo $Ilo -Token $Token -Path $QueueUri
        foreach ($ref in @($q.Members)) {
            $u = $ref.'@odata.id'
            if (-not $u) { continue }
            try {
                $t = ILO-Get -Ilo $Ilo -Token $Token -Path $u
                $fn = ''
                if ($t.PSObject.Properties.Name -contains 'Filename' -and $t.Filename) { $fn = "$($t.Filename)" }
                elseif ($t.PSObject.Properties.Name -contains 'Component' -and $t.Component) { $fn = "$($t.Component)" }
                $tasks += [pscustomobject]@{ Name = "$($t.Name)"; Filename = $fn; Id = "$($t.Id)"; Uri = "$u"; State = "$($t.State)" }
            } catch {}
        }
    } catch {}
    ,$tasks
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
        [string]$TaskName = $null,
        [scriptblock]$LogCb = $null
    )
    $say = { param($m) if ($LogCb) { try { & $LogCb $m } catch {} } }

    # UpdateTaskQueue-URI ermitteln (Oem/Hpe-Link, sonst Standardpfad).
    $queueUri = '/redfish/v1/UpdateService/UpdateTaskQueue'
    try {
        $us = ILO-Get -Ilo $Ilo -Token $Token -Path "/redfish/v1/UpdateService"
        if ($us.Oem -and $us.Oem.Hpe -and $us.Oem.Hpe.UpdateTaskQueue -and $us.Oem.Hpe.UpdateTaskQueue.'@odata.id') {
            $queueUri = "$($us.Oem.Hpe.UpdateTaskQueue.'@odata.id')"
        }
    } catch {}

    # Den EXAKTEN Repository-Dateinamen ermitteln: iLO nutzt den Filename als
    # eindeutigen Schluessel; der Task muss genau diesen Namen referenzieren -
    # sonst wird kein Task angelegt (POST evtl. trotzdem 2xx).
    $match = ILO-FindRepositoryComponent -Ilo $Ilo -Token $Token -ComponentFileName $ComponentFileName
    if ($match -and $match.Filename) { $ComponentFileName = "$($match.Filename)" }
    else { & $say "  WARN: '$ComponentFileName' nicht eindeutig im iLO-Repository gefunden - lege Task trotzdem mit diesem Namen an." }

    # Task-Name muss eindeutig sein UND wird Teil der iLO-Task-URI -> nur
    # URI-sichere Zeichen (keine Leerzeichen/Punkte), sonst 400 Bad Request.
    if (-not $TaskName) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ComponentFileName)
        $TaskName = "Update-$baseName-$((Get-Date).ToString('yyyyMMddHHmmss'))"
    }
    $TaskName = ($TaskName -replace '[^A-Za-z0-9_-]', '-') -replace '-{2,}', '-'
    if ($TaskName.Length -gt 63) { $TaskName = $TaskName.Substring(0, 63) }

    # Bestehende Tasks merken -> hinterher pruefen, ob unserer NEU dazukam.
    $before = @(ILO-GetUpdateTaskQueue -Ilo $Ilo -Token $Token -QueueUri $queueUri)
    $beforeNames = @($before | ForEach-Object { $_.Name })

    # IDEMPOTENZ: Existiert bereits ein Task fuer genau diese Datei (z.B. aus einem
    # frueheren, noch nicht per Reboot aktivierten Lauf)? Dann ist nichts zu tun -
    # ein erneuter POST wuerde nur ein Duplikat erzeugen bzw. abgelehnt werden.
    $already = $before | Where-Object { "$($_.Filename)" -ieq $ComponentFileName } | Select-Object -First 1
    if ($already) {
        & $say "  Task fuer '$ComponentFileName' ist bereits in der Queue - kein erneuter POST noetig."
        if ($already.Uri) { return "$($already.Uri)" }
        return ''
    }

    # Payload-Varianten in Reihenfolge der Zuverlaessigkeit:
    #  - Feldname 'Filename' zuerst (iLO 5/6/7 laut HPE-Referenz), 'Component' als
    #    Fallback fuer abweichende Firmwares.
    #  - OHNE 'TPMOverride' zuerst: HPE-Referenz und iLO-Doku senden dieses Feld im
    #    Task NICHT. Manche iLO-Firmwares akzeptieren den POST mit unbekanntem Feld
    #    (2xx), legen aber KEINEN Task an -> genau der beobachtete Fehler. MIT
    #    'TPMOverride' nur als Fallback (relevant wenn TPM aktiv Direktflash blockt).
    $variants = @()
    foreach ($fileKey in @('Filename', 'Component')) {
        foreach ($withTpm in @($false, $true)) {
            $variants += @{ FileKey = $fileKey; Tpm = $withTpm }
        }
    }

    $lastDetail = ''
    $lastBody = ''
    foreach ($v in $variants) {
        $bodyObj = [ordered]@{
            Name        = $TaskName
            Command     = 'ApplyUpdate'
            UpdatableBy = @($UpdatableBy)
        }
        $bodyObj[$v.FileKey] = $ComponentFileName   # iLO5='Filename', iLO6/7 alt.='Component'
        if ($v.Tpm) { $bodyObj['TPMOverride'] = $true }
        $body = $bodyObj | ConvertTo-Json -Compress
        $lastBody = $body

        $loc = $null
        try {
            $resp = Invoke-WebRequest -Uri "https://$Ilo$queueUri" -Method Post -Body $body `
                -ContentType 'application/json' -Headers @{ 'X-Auth-Token' = $Token; 'OData-Version' = '4.0' } `
                -SkipCertificateCheck -TimeoutSec $script:IloTimeoutSec -ErrorAction Stop
            $loc = $resp.Headers['Location']; if ($loc -is [array]) { $loc = $loc[0] }
            if (-not $loc) { try { $j = $resp.Content | ConvertFrom-Json; if ($j.'@odata.id') { $loc = "$($j.'@odata.id')" } } catch {} }
        } catch {
            # iLO-Fehlertext (MessageId/Detail) aus dem Response-Body herausziehen.
            $detail = ''
            try { if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detail = $_.ErrorDetails.Message } } catch {}
            if (-not $detail) {
                try {
                    $rs = $_.Exception.Response.GetResponseStream()
                    if ($rs) { $sr = New-Object System.IO.StreamReader($rs); $detail = $sr.ReadToEnd(); $sr.Close() }
                } catch {}
            }
            $lastDetail = $detail
            # Unbekanntes/nicht schreibbares Feld -> naechste Variante probieren.
            if ($detail -match 'PropertyNotWritableOrUnknown|PropertyUnknown|PropertyNotWritable') { continue }
            & $say "  Task-POST ($($v.FileKey), TPMOverride=$($v.Tpm)) abgelehnt: $($_.Exception.Message)$(if($detail){" | $detail"})"
            continue
        }

        # VERIFIKATION: Queue erneut lesen und pruefen, ob unser Task nun existiert.
        Start-Sleep -Milliseconds 400
        $after = @(ILO-GetUpdateTaskQueue -Ilo $Ilo -Token $Token -QueueUri $queueUri)
        $mine = $after | Where-Object {
            $_.Name -eq $TaskName -or
            ($_.Name -notin $beforeNames -and "$($_.Filename)" -ieq $ComponentFileName)
        } | Select-Object -First 1
        if ($mine) {
            if ($mine.Uri) { return "$($mine.Uri)" }
            if ($loc) { try { return ([Uri]$loc).AbsolutePath } catch { return "$loc" } }
            return ''
        }

        # POST quittiert, Task aber NICHT in der Queue -> naechste Payload-Variante.
        & $say "  Task-POST ($($v.FileKey), TPMOverride=$($v.Tpm)) quittiert, aber Task nicht in Queue - probiere naechste Variante."
    }

    $qNow = @(ILO-GetUpdateTaskQueue -Ilo $Ilo -Token $Token -QueueUri $queueUri)
    $qTxt = if ($qNow.Count) { ($qNow | ForEach-Object { "$($_.Name)=$($_.Filename)" }) -join '; ' } else { '(leer)' }
    throw "UpdateTaskQueue-Task konnte nicht angelegt werden fuer '$ComponentFileName'. Letzte Payload=$lastBody. iLO-Antwort=$lastDetail. Queue jetzt: $qTxt"
}


# Waehlt anhand des Modells den passenden Typ-Unterordner im Firmware-Verzeichnis.
# Empfohlene (eindeutige) Benennung: Ordnername == kompletter Modellname aus
# Servers.txt (Feld 'Modell='), z.B. "ProLiant DL380a Gen11". Leer-/Unterstrich-/
# Bindestrich-Unterschiede werden dabei ignoriert (Normalisierung).
# Alternativ wird das Schema <TYP>_Gen<NN> unterstuetzt (z.B. DL380a_Gen11).
# WICHTIG: Der Typ wird buchstabengenau verglichen - 'DL380' und 'DL380a' sind
#          unterschiedliche Server und werden NIE verwechselt.
# Liefert den vollstaendigen Ordnerpfad oder wirft einen Fehler.
function Resolve-FirmwareFolder {
    param([string]$BaseDir,[string]$Model,[int]$Gen)
    if (-not (Test-Path -LiteralPath $BaseDir -PathType Container)) { throw "Firmware-Verzeichnis nicht gefunden: $BaseDir" }
    $subs = @(Get-ChildItem -LiteralPath $BaseDir -Directory -ErrorAction Stop)
    if ($subs.Count -eq 0) { throw "Keine Typ-Unterordner im Firmware-Verzeichnis: $BaseDir" }

    # Regel 1: Ordnername == kompletter Modellname (normalisiert, ohne Trenner).
    #   "ProLiant DL380a Gen11" == Ordner "ProLiant DL380a Gen11" oder "ProLiant_DL380a_Gen11"
    $normModel = ($Model -replace '[^A-Za-z0-9]', '').ToLower()
    if ($normModel) {
        foreach ($d in $subs) {
            $normDir = ($d.Name -replace '[^A-Za-z0-9]', '').ToLower()
            if ($normDir -eq $normModel) { return $d.FullName }
        }
    }

    # Regel 2: Schema <TYP>_Gen<NN> - exakte Uebereinstimmung Typ + Generation.
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
    return @{ Kind = 'Other'; InvPattern = ''; IsIlo = $false; Deferred = $false; UpdatableBy = @('Uefi') }
}

# Liest Metadaten zu einer .fwpkg-Firmwaredatei.
# Zwei HPE-Liefervarianten werden unterstuetzt:
#   Gen10/10Plus/11: Metadaten als 'payload.json' IM .fwpkg-ZIP eingebettet.
#   Gen12:           Metadaten als SEPARATE Sidecar-Datei NEBEN der .fwpkg mit
#                    gleichem Basisnamen (z.B. ilo7_1.22.00.fwpkg ->
#                    ilo7_1.22.00.json). Diese wird BEVORZUGT ausgewertet.
# Liefert @{ Version; Deferred; Reboot } mit Tri-State-Logik fuer Deferred/Reboot:
#   Deferred: $true = nicht direkt flashbar (Repository + Task Queue)
#             $false = direkt online flashbar
#             $null  = keine Auskunft (Dateiname-Heuristik entscheidet)
#   Reboot:   $true/$false/$null analog (Aktivierung erfordert Reboot?)
# Maßgebliche Felder (Schreibweise variiert je Generation, (?i)/Property-Zugriff
# decken beide ab):
#   "DirectFlashOk"/"DirectFlashOK": true/false -> false => Deferred
#   "UefiFlashable"/"UEFIFlashable": true/false -> true  => Deferred
#   "ResetRequired": true/false   -> Reboot (Alias: RebootRequired)
#   "UpdatableBy": ["Bmc"|"Uefi"|"RuntimeAgent"] -> Bmc/iLO=direkt, sonst deferred
#   "Version"/"FirmwareVersion"/"TargetVersion"
function Get-FwpkgMeta {
    param([string]$FilePath)

    # Wertet ein einzelnes JSON-Objekt aus und setzt die Rueckgabefelder.
    $parseObj = {
        param($o, $res)
        if (-not $o) { return }
        # Version (mehrere moegliche Schluesselnamen)
        foreach ($vk in @('Version', 'FirmwareVersion', 'TargetVersion')) {
            if (-not $res.Version -and $o.PSObject.Properties.Name -contains $vk) {
                $cand = "$($o.$vk)".Trim()
                if ($cand -match '\d') { $res.Version = $cand; break }
            }
        }
        # DirectFlashOk -> Deferred (false bedeutet deferred)
        if ($null -eq $res.Deferred -and $o.PSObject.Properties.Name -contains 'DirectFlashOk') {
            try { $res.Deferred = -not [bool]$o.DirectFlashOk } catch {}
        }
        # UefiFlashable -> Deferred (true bedeutet Flash via UEFI beim Reboot)
        if ($null -eq $res.Deferred -and $o.PSObject.Properties.Name -contains 'UefiFlashable') {
            try { $res.Deferred = [bool]$o.UefiFlashable } catch {}
        }
        # ResetRequired / RebootRequired -> Reboot
        if ($null -eq $res.Reboot) {
            foreach ($rk in @('ResetRequired', 'RebootRequired')) {
                if ($o.PSObject.Properties.Name -contains $rk) {
                    try { $res.Reboot = [bool]$o.$rk; break } catch {}
                }
            }
        }
        # UpdatableBy -> Deferred-Hinweis (nur falls DirectFlashOk fehlte)
        if ($null -eq $res.Deferred -and $o.PSObject.Properties.Name -contains 'UpdatableBy') {
            $ub = (@($o.UpdatableBy) -join ',').ToLower()
            $hasBmc = ($ub -match 'bmc|ilo')
            $hasDef = ($ub -match 'uefi|runtimeagent|sut|sum')
            if     ($hasBmc -and -not $hasDef) { $res.Deferred = $false }
            elseif ($hasDef -and -not $hasBmc) { $res.Deferred = $true }
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
                        foreach ($img in @($dev.FirmwareImages)) { if ($img) { [void]$list.Add($img) } }
                    }
                    [void]$list.Add($dev)
                }
            }
            # Anderes Format: Components[]
            if ($t.PSObject.Properties.Name -contains 'Components') {
                foreach ($c in @($t.Components)) { if ($c) { [void]$list.Add($c) } }
            }
            # Zuletzt das Top-Objekt selbst (UpdatableBy etc.)
            [void]$list.Add($t)
        }
        return $list.ToArray()
    }

    $res = @{ Version = ''; Deferred = $null; Reboot = $null }

    # Wertet einen JSON-Text strukturiert aus (flatten + parseObj). Faellt bei
    # nicht-parsebarem JSON auf gezielte Regex-Suche zurueck. (?i) deckt die
    # verschiedenen HPE-Schreibweisen ab (DirectFlashOk/DirectFlashOK,
    # UefiFlashable/UEFIFlashable).
    $applyText = {
        param($txt)
        if (-not $txt) { return }
        try {
            $obj = $txt | ConvertFrom-Json -ErrorAction Stop
            foreach ($it in @(& $flatten $obj)) { & $parseObj $it $res }
        } catch {
            if (-not $res.Version -and $txt -match '(?i)"(FirmwareVersion|TargetVersion|Version)"\s*:\s*"([^"]+)"') {
                $cand = $Matches[2].Trim(); if ($cand -match '\d') { $res.Version = $cand }
            }
            if ($null -eq $res.Deferred -and $txt -match '(?i)"DirectFlashOk"\s*:\s*(true|false)') { $res.Deferred = ($Matches[1] -ieq 'false') }
            if ($null -eq $res.Deferred -and $txt -match '(?i)"UefiFlashable"\s*:\s*(true|false)') { $res.Deferred = ($Matches[1] -ieq 'true') }
            if ($null -eq $res.Reboot   -and $txt -match '(?i)"(ResetRequired|RebootRequired)"\s*:\s*(true|false)') { $res.Reboot = ($Matches[2] -ieq 'true') }
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
        if ($fn -match '(?i)ilo\d?[ _\-]?(\d)(\d{2})$') { $res.Version = "$($Matches[1]).$($Matches[2])" }
        elseif ($fn -match '(\d+\.\d+(?:\.\d+)*)')      { $res.Version = $Matches[1] }
    }
    return $res
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
function Set-StagedFlash {
    # Legt/aktualisiert den Marker fuer (Serial, Kind) auf die Zielversion.
    param([string]$StateDir,[string]$Serial,[string]$Kind,[string]$File,[string]$TargetNorm)
    if (-not $Serial -or -not $Kind) { return }
    $mtx = New-Object System.Threading.Mutex($false,'Global\OV_RackmountFw_StagedFlash')
    try { [void]$mtx.WaitOne() } catch {}
    try {
        $list = @(Read-StagedFlashes -StateDir $StateDir | Where-Object { -not ("$($_.Serial)" -ieq $Serial -and "$($_.Kind)" -ieq $Kind) })
        $list += [pscustomobject]@{ Serial = $Serial; Kind = $Kind; File = $File; TargetNorm = $TargetNorm; StagedAt = (Get-Date).ToString('s') }
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
'@

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

# Liefert die beste iLO-Adresse aus einem server-hardware-Objekt (LinkLocal hinten).
function Get-IloIpFromHardware {
    param($sh)
    if (-not $sh) { return $null }
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $scored = New-Object System.Collections.Generic.List[object]
    $addOne = {
        param([string]$addr)
        if ([string]::IsNullOrWhiteSpace($addr)) { return }
        if ($addr -match '\s') { return }
        if ($addr -match '^(.+)%[^%]+$') { $addr = $matches[1] }
        if ($seen.Contains($addr)) { return }
        [void]$seen.Add($addr)
        $isV6 = $addr -match ':'
        $prio = 0
        if ($addr -match '^(?i)fe80:') { $prio = 3 }
        elseif ($addr -match '^169\.254\.') { $prio = 2 }
        elseif ($isV6) { $prio = 1 }
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
$form.Text = "© 2025 N.J. Airbus D&S - HPE ProLiant Rackmount Firmware Update (Gen10+, iLO Redfish)"
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
$txtOvFile.Location = '90,25'; $txtOvFile.Size = '300,22'; $txtOvFile.Text = (Join-Path $scriptFolder "Oneview.txt"); $txtOvFile.BorderStyle = 'FixedSingle'
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
$btnLoadCache.Location = '665,24'; $btnLoadCache.Size = '150,24'; $btnLoadCache.Text = "Aus Servers.txt laden"
$grpSrc.Controls.Add($btnLoadCache)

$chkOnlyGen10 = New-Object System.Windows.Forms.CheckBox
$chkOnlyGen10.Location = '825,26'; $chkOnlyGen10.Size = '110,20'; $chkOnlyGen10.Text = [char]0x2265 + " Gen10"; $chkOnlyGen10.Checked = $true
$grpSrc.Controls.Add($chkOnlyGen10)

$lblOvHint = New-Object System.Windows.Forms.Label
$lblOvHint.AutoSize = $true; $lblOvHint.Location = '945,28'; $lblOvHint.Text = "(1 Appliance/Zeile)"; $lblOvHint.ForeColor = [System.Drawing.Color]::Gray
$grpSrc.Controls.Add($lblOvHint)

$tipOv = New-Object System.Windows.Forms.ToolTip
$tipOv.AutoPopDelay = 15000
$tipOv.SetToolTip($txtOvFile, "Textdatei mit einer OneView-Appliance (Hostname/IP) pro Zeile.`r`nDas Tool liest server-hardware + Scopes aus und erzeugt daraus die Serverliste (Servers.txt).")
$tipOv.SetToolTip($btnLoadCache, "Laedt die zuletzt aus OneView erzeugte Servers.txt (offline, inkl. Modell + Scope).")

# Servers.txt = generierte/gecachte Liste (Output des OneView-Imports)
$serversFile = Join-Path $scriptFolder "Servers.txt"

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

$chkServers = New-Object System.Windows.Forms.CheckedListBox
$chkServers.Location = '12,140'; $chkServers.Size = '1103,146'; $chkServers.CheckOnClick = $true; $chkServers.BorderStyle = 'FixedSingle'
$chkServers.IntegralHeight = $false
$chkServers.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($chkServers)

# Voller (ungefilterter) Bestand: Liste von Objekten @{ Raw='host[;user;pass]'; Name; Model; Gen; Scope }
$script:allServers = @()

# Hilfsfunktion: erzeugt den Anzeigetext eines Server-Objekts in festen Spalten:
#   IP/iLO 16 | Servername 40 | Servertyp 38 | Scope (Rest), jeweils 1 Leerzeichen getrennt.
function Get-ServerDisplay {
    param($obj)
    $ilo = ($obj.Raw -split ';')[0].Trim()
    $name  = if ($obj.Name)  { "$($obj.Name)" }  else { '' }
    $model = if ($obj.Model) { "$($obj.Model)" } else { '' }
    if ($name.Length  -gt 40) { $name  = $name.Substring(0, 40) }
    if ($model.Length -gt 38) { $model = $model.Substring(0, 38) }
    $line = '{0} {1} {2}' -f $ilo.PadRight(16), $name.PadRight(40), $model.PadRight(38)
    if ($obj.Scope) { $line += ' {' + $obj.Scope + '}' }
    return $line
}

# Schreibt den aktuellen Bestand als Servers.txt (reloadbar inkl. Modell/Gen/Scope).
function Save-ServersFile {
    param([string]$Path)
    $lines = @(
        "# Automatisch aus OneView erzeugt - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "# Format: ilo[;user;pass]    # Name=...; Modell=...; Gen=NN; Scope=...",
        "# Diese Datei kann ueber 'Aus Servers.txt laden' offline wieder eingelesen werden."
    )
    foreach ($obj in ($script:allServers | Sort-Object Raw)) {
        $meta = "# Name=$($obj.Name); Modell=$($obj.Model); Gen=$($obj.Gen); Scope=$($obj.Scope)"
        $lines += ("{0}    {1}" -f $obj.Raw, $meta)
    }
    try { Set-Content -LiteralPath $Path -Value $lines -Encoding UTF8 } catch { Add-Log "Servers.txt konnte nicht geschrieben werden: $($_.Exception.Message)" ([System.Drawing.Color]::Red) }
}

# Laedt eine zuvor erzeugte Servers.txt (inkl. Metadaten) zurueck in den Bestand.
function Load-ServersFile {
    param([string]$Path)
    $script:allServers = @()
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        [System.Windows.Forms.MessageBox]::Show("Servers.txt nicht gefunden:`n$Path", "Hinweis", 'OK', 'Warning'); return
    }
    foreach ($line in Get-Content -LiteralPath $Path) {
        $t = $line.Trim()
        if ($t -eq '' -or $t.StartsWith('#')) { continue }
        $name = ''; $model = ''; $gen = 0; $scope = ''
        $hostPart = $t
        $hashIdx = $t.IndexOf('#')
        if ($hashIdx -ge 0) {
            $hostPart = $t.Substring(0, $hashIdx).Trim()
            $meta = $t.Substring($hashIdx + 1)
            if ($meta -match '(?i)Name=([^;]*)')   { $name = $Matches[1].Trim() }
            if ($meta -match '(?i)Modell=([^;]*)') { $model = $Matches[1].Trim() }
            if ($meta -match '(?i)Gen=(\d+)')      { $gen = [int]$Matches[1] }
            if ($meta -match '(?i)Scope=(.*)$')    { $scope = $Matches[1].Trim() }
        }
        if ($hostPart -eq '') { continue }
        $script:allServers += [pscustomobject]@{ Raw = $hostPart; Name = $name; Model = $model; Gen = $gen; Scope = $scope }
    }
    Rebuild-FilterChoices
    Apply-ServerFilter
    Add-Log "Aus Servers.txt geladen: $($script:allServers.Count) Server." ([System.Drawing.Color]::DarkBlue)
}

# Befuellt die Servertyp- und Scope-Dropdowns aus dem aktuellen Bestand.
function Rebuild-FilterChoices {
    $selType = if ($cboType.SelectedItem) { "$($cboType.SelectedItem)" } else { "(alle)" }
    $selScope = if ($cboScope.SelectedItem) { "$($cboScope.SelectedItem)" } else { "(alle)" }

    $models = @($script:allServers | Where-Object { $_.Model } | ForEach-Object { $_.Model } | Sort-Object -Unique)
    $scopes = @($script:allServers | Where-Object { $_.Scope } | ForEach-Object { $_.Scope -split '\s*,\s*' } | Where-Object { $_ } | Sort-Object -Unique)

    $cboType.BeginUpdate(); $cboType.Items.Clear(); $cboType.Items.Add("(alle)") | Out-Null
    foreach ($m in $models) { $cboType.Items.Add($m) | Out-Null }
    $i = $cboType.Items.IndexOf($selType); $cboType.SelectedIndex = if ($i -ge 0) { $i } else { 0 }
    $cboType.EndUpdate()

    $cboScope.BeginUpdate(); $cboScope.Items.Clear(); $cboScope.Items.Add("(alle)") | Out-Null
    foreach ($s in $scopes) { $cboScope.Items.Add($s) | Out-Null }
    $i = $cboScope.Items.IndexOf($selScope); $cboScope.SelectedIndex = if ($i -ge 0) { $i } else { 0 }
    $cboScope.EndUpdate()
}

function Apply-ServerFilter {
    $flt = $txtFilter.Text.Trim()
    $selType = if ($cboType.SelectedItem) { "$($cboType.SelectedItem)" } else { "(alle)" }
    $selScope = if ($cboScope.SelectedItem) { "$($cboScope.SelectedItem)" } else { "(alle)" }
    $chkServers.BeginUpdate()
    $chkServers.Items.Clear()
    foreach ($obj in $script:allServers) {
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
        if ($show) { $chkServers.Items.Add((Get-ServerDisplay $obj), $false) | Out-Null }
    }
    $chkServers.EndUpdate()
    $lblCount.Text = "$($chkServers.Items.Count) angezeigt / $($script:allServers.Count) gesamt"
}

# Liefert ausgewaehlte Server als @{ Ilo; User; Pass }. iLO = erstes Token der Zeile.
function Get-CheckedServers {
    $result = @()
    for ($i = 0; $i -lt $chkServers.Items.Count; $i++) {
        if ($chkServers.GetItemChecked($i)) {
            $txt = $chkServers.Items[$i].ToString().Trim()
            $ilo = (($txt -split '\s+', 2)[0]).Trim()   # iLO-Adresse = erstes Token (vor Name/Modell/Scope)
            $u = $txtUser.Text; $pw = $txtPass.Text
            $obj = $script:allServers | Where-Object { (($_.Raw -split ';')[0].Trim()) -eq $ilo } | Select-Object -First 1
            if ($obj) {
                $parts = $obj.Raw -split ';'
                if ($parts.Count -ge 2 -and $parts[1].Trim() -ne '') { $u = $parts[1].Trim() }
                if ($parts.Count -ge 3 -and $parts[2].Trim() -ne '') { $pw = $parts[2].Trim() }
            }
            $result += @{ Ilo = $ilo; User = $u; Pass = $pw }
        }
    }
    ,$result
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
        $appliances += ($t -split ';')[0].Trim()
    }
    if ($appliances.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Keine Appliances in der Datei.", "Hinweis", 'OK', 'Warning'); return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnLoadOV.Enabled = $false; $btnLoadCache.Enabled = $false; $btnInventory.Enabled = $false; $btnFlash.Enabled = $false
    $collected = @{}
    try {
        foreach ($appl in $appliances) {
            Add-Log "OneView $appl : Abfrage server-hardware + Scopes..." ([System.Drawing.Color]::DarkBlue)
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
                    if ($chkOnlyGen10.Checked -and $gen -lt 10) { $skipped++; continue }
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
                    $collected[$ilo] = [pscustomobject]@{ Raw = $ilo; Name = $srvName; Model = $model; Gen = $gen; Scope = $scope }
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

$btnSelAll.Add_Click({ for ($i = 0; $i -lt $chkServers.Items.Count; $i++) { $chkServers.SetItemChecked($i, $true) } })
$btnSelNone.Add_Click({ for ($i = 0; $i -lt $chkServers.Items.Count; $i++) { $chkServers.SetItemChecked($i, $false) } })
$txtFilter.Add_TextChanged({ Apply-ServerFilter })
$cboType.Add_SelectedIndexChanged({ Apply-ServerFilter })
$cboScope.Add_SelectedIndexChanged({ Apply-ServerFilter })
$btnOvBrowse.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "Textdateien (*.txt)|*.txt|Alle (*.*)|*.*"
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
$lblFwHint.Text = "Typ-Ordner = Modellname aus Servers.txt (z.B. 'ProLiant DL380a Gen11'). DL380 != DL380a. Dateinamen NICHT aendern."
$grpFw.Controls.Add($lblFwHint)

$tipFw = New-Object System.Windows.Forms.ToolTip
$tipFw.AutoPopDelay = 20000
$tipFw.SetToolTip($rbBaseDir, "Firmware-Verzeichnis mit je einem Unterordner pro Servertyp.`r`nEmpfohlen: Ordnername == Modellname aus Servers.txt (Feld 'Modell='), z.B.:`r`n  Firmware\ProLiant DL360 Gen10\*.fwpkg`r`n  Firmware\ProLiant DL380 Gen11\*.fwpkg`r`n  Firmware\ProLiant DL380a Gen11\*.fwpkg`r`nAlternativ Schema <TYP>_Gen<NN> (z.B. DL380a_Gen11).`r`nDL380 und DL380a sind verschiedene Typen - jeder Buchstabe zaehlt.`r`nDie Zuordnung erfolgt automatisch anhand des am iLO ausgelesenen Modells.")

$btnBrowseFw.Add_Click({
    if ($rbFile.Checked) {
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "HPE Firmware (*.fwpkg)|*.fwpkg|Alle (*.*)|*.*"
        if ($ofd.ShowDialog() -eq 'OK') { $txtFw.Text = $ofd.FileName }
    } else {
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = if ($rbBaseDir.Checked) { "Basisverzeichnis mit Typ-Unterordnern auswaehlen" } else { "Ordner mit .fwpkg-Dateien auswaehlen" }
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

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = '540,418'; $progress.Size = '575,26'; $progress.Minimum = 0; $progress.Maximum = 100
$form.Controls.Add($progress)

# ─────────────────────────────────────────
# Ergebnis-Liste
# ─────────────────────────────────────────
$lv = New-Object System.Windows.Forms.ListView
$lv.Location = '12,452'; $lv.Size = '1103,248'; $lv.View = 'Details'; $lv.FullRowSelect = $true; $lv.GridLines = $true
$lv.Scrollable = $true   # horizontale Scrollleiste, sobald Spalten breiter als Steuerelement sind
$lv.Columns.Add("Server (iLO)", 220) | Out-Null
$lv.Columns.Add("Modell", 200) | Out-Null
$lv.Columns.Add("Phase", 160) | Out-Null
$lv.Columns.Add("Fortschritt", 90) | Out-Null
$lv.Columns.Add("Ergebnis", 90) | Out-Null
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
    $logBox.SelectionColor = $Color
    $logBox.AppendText("$Text`r`n")
    $logBox.SelectionColor = $logBox.ForeColor
    $logBox.ScrollToCaret()
    Write-LogFile $Text
}

# ─────────────────────────────────────────
# Async-Engine: ConcurrentQueue + Timer (UI-Thread)
# ─────────────────────────────────────────
$script:uiQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
$script:doneCount = 0
$script:totalCount = 0

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
    if ([string]::IsNullOrWhiteSpace($txtUser.Text) -or [string]::IsNullOrWhiteSpace($txtPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte iLO-Benutzer und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return $false
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

    $btnInventory.Enabled = $false; $btnFlash.Enabled = $false
    Add-Log "=== Inventar-Pruefung fuer $($servers.Count) Server gestartet ===" ([System.Drawing.Color]::DarkBlue)

    $worker = {
        param($p)
        $iloCode = $p.iloCode; $ilo = $p.ilo; $user = $p.user; $pass = $p.pass; $uiQueue = $p.uiQueue
        Invoke-Expression $iloCode
        $log = { param($t) $uiQueue.Enqueue(@{ Type = 'LOG'; Text = "$ilo : $t" }) }.GetNewClosure()
        $sess = $null
        try {
            & $log "Login am iLO..."
            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = 'Login...' })
            $sess = ILO-Login -Ilo $ilo -User $user -Pass $pass
            & $log "Login OK, lese Inventar..."
            $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = 'Lese Inventar...' })
            $info = ILO-GetSystemInfo -Ilo $ilo -Token $sess.Token
            $uiQueue.Enqueue(@{ Type = 'MODEL'; Ilo = $ilo; Model = $info.Model })
            & $log "Systeminfo: Modell='$($info.Model)', Gen=$($info.Gen), iLO-Gen=$($info.iLO), SN=$($info.Serial)"

            if ($info.Gen -gt 0 -and $info.Gen -lt 10) {
                & $log "Uebersprungen: Gen$($info.Gen) < Gen10 (nicht unterstuetzt)"
                $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Nicht unterstuetzt'; Detail = "Generation Gen$($info.Gen) < Gen10 - uebersprungen" })
                return
            }

            $inv = ILO-GetFirmwareInventory -Ilo $ilo -Token $sess.Token
            $bios = ($inv | Where-Object { $_.Name -match '(?i)System ROM|BIOS' } | Select-Object -First 1)
            $iloFw = ($inv | Where-Object { $_.Name -match '(?i)iLO' } | Select-Object -First 1)
            $biosV = if ($bios) { $bios.Version } else { '?' }
            $iloV = if ($iloFw) { $iloFw.Version } else { '?' }
            $detail = "BIOS: $biosV | iLO: $iloV | Komponenten: $($inv.Count)"
            & $log "Inventar gelesen: $detail"
            $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $true; Phase = "OK (Gen$($info.Gen))"; Detail = $detail })
        }
        catch {
            $errMsg = $_.Exception.Message
            if ($_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber) { $errMsg = "$errMsg (Zeile $($_.InvocationInfo.ScriptLineNumber))" }
            & $log "FEHLER: $errMsg"
            $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Fehler'; Detail = $_.Exception.Message })
        }
        finally {
            if ($sess) { try { ILO-Logout -Ilo $ilo -Token $sess.Token -SessionUri $sess.SessionUri; & $log "Logout OK" } catch {} }
        }
    }

    Start-Batch -Servers $servers -Worker $worker -ExtraArgs @{} -MaxParallel ([int]$numPar.Value)
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
    if ($rbBaseDir.Checked) {
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
        $components = @($components | Sort-Object @{ Expression = { [System.IO.Path]::GetFileName($_) -match '(?i)ilo' }; Ascending = $true }, @{ Expression = { $_ } })
    }

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

    if ($mode -eq 'basedir') {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "ACHTUNG - Firmware wird auf $($servers.Count) Server geflasht (max. $([int]$numPar.Value) parallel).`n`n" +
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
        Add-Log "=== Firmware-Update (autom. Typ-Zuordnung) gestartet auf $($servers.Count) Server ===" ([System.Drawing.Color]::DarkBlue)
    } else {
        $compNames = ($components | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join ", "
        $res = [System.Windows.Forms.MessageBox]::Show(
            "ACHTUNG - Firmware wird auf $($servers.Count) Server geflasht (max. $([int]$numPar.Value) parallel).`n`n" +
            "Komponenten ($($components.Count)):`n$compNames`n`n" +
            "Bereits aktuelle Komponenten werden automatisch uebersprungen.`n" +
            "Es wird NICHT automatisch rebootet.`n" +
            "Reine iLO-Updates sind nach dem iLO-Selbstreset sofort aktiv (kein Reboot).`n" +
            "BIOS/SPS/Systemfirmware wird erst beim naechsten regulaeren Neustart aktiv.`n`nFortfahren?",
            "Firmware-Update bestaetigen", 'YesNo', 'Warning')
        if ($res -ne 'Yes') { return }
        $btnInventory.Enabled = $false; $btnFlash.Enabled = $false
        Add-Log "=== Firmware-Update gestartet: $($components.Count) Komponente(n) auf $($servers.Count) Server ===" ([System.Drawing.Color]::DarkBlue)
    }

    $worker = {
        param($p)
        $iloCode = $p.iloCode; $ilo = $p.ilo; $user = $p.user; $pass = $p.pass; $uiQueue = $p.uiQueue
        $components = $p.components; $mode = $p.mode; $baseDir = $p.baseDir; $scriptFolder = $p.scriptFolder
        Invoke-Expression $iloCode
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
            # HPE Synergy Compute-Module werden zentral ueber OneView-Firmware-
            # Baselines (SPP/Custom SPP) verwaltet. Direktes iLO-Flashen kann mit
            # dem Baseline-Management kollidieren -> hier bewusst hart blockieren.
            if ($info.Model -match '(?i)Synergy') {
                & $log "BLOCKIERT: '$($info.Model)' ist ein HPE Synergy Compute-Modul - Firmware wird ueber OneView (Baseline/SPP) verwaltet. Kein direktes iLO-Update. Server uebersprungen."
                $uiQueue.Enqueue(@{ Type = 'DONE'; Ilo = $ilo; Success = $false; Phase = 'Synergy - blockiert'; Detail = "Synergy Compute-Modul '$($info.Model)' - Firmware via OneView verwalten, direktes iLO-Flashen unterbunden" })
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
                $components = @($components | Sort-Object @{ Expression = { [System.IO.Path]::GetFileName($_) -match '(?i)ilo' }; Ascending = $true }, @{ Expression = { $_ } })
                & $log "Typ-Ordner '$([System.IO.Path]::GetFileName($folder))', $($components.Count) Komponente(n): $((@($components | ForEach-Object { [System.IO.Path]::GetFileName($_) })) -join ', ')"
            }

            $total = $components.Count
            $idx = 0

            # Firmware-Inventar einmal lesen (fuer Versionsvergleich / Skip).
            $inv = @()
            try { $inv = ILO-GetFirmwareInventory -Ilo $ilo -Token $token } catch { & $log "Inventar konnte nicht gelesen werden: $($_.Exception.Message)" }
            if ($inv.Count) { & $log "Firmware-Inventar ($($inv.Count) Eintraege): $((@($inv | ForEach-Object { "$($_.Name)=$($_.Version)" })) -join ' | ')" }

            $okNames = @(); $skipNames = @(); $absentNames = @(); $stagedNames = @(); $failItems = @()
            $rebootNeeded = $false

            foreach ($comp in $components) {
                $idx++
                $name = [System.IO.Path]::GetFileName($comp)
                $sizeMb = [Math]::Round(((Get-Item -LiteralPath $comp).Length / 1MB), 1)
                $kind = Get-ComponentKind -FileName $name

                # Nicht verbaute Hardware ueberspringen: gibt es im Firmware-Inventar
                # ueberhaupt einen passenden Eintrag? Wenn nein (und der Typ ist
                # eindeutig erkannt), ist die Komponente nicht vorhanden -> nicht flashen.
                $instItem = $null
                $instMatches = @()
                if ($kind.InvPattern) {
                    $instMatches = @($inv | Where-Object { $_.Name -match $kind.InvPattern })
                    $instItem = $instMatches[0]
                    if ($inv.Count -gt 0 -and -not $instItem) {
                        & $log "Uebersprungen $idx/$total : '$name' - Hardware nicht verbaut (kein passender Inventar-Eintrag fuer [$($kind.Kind)])"
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Nicht verbaut $idx/$total : $name" })
                        $absentNames += $name
                        continue
                    }
                }

                # .fwpkg-Metadaten lesen: Version + (falls vorhanden) verbindlicher
                # Deferred-Hinweis. Die Metadaten sind autoritativer als der Dateiname.
                $meta = Get-FwpkgMeta -FilePath $comp
                $tgtVer = $meta.Version
                if ($null -ne $meta.Deferred -and $meta.Deferred -ne $kind.Deferred) {
                    $mTxt = if ($meta.Deferred) { 'deferred (Repository/Queue)' } else { 'direkt flashbar' }
                    $nTxt = if ($kind.Deferred) { 'deferred' } else { 'direkt' }
                    & $log "Hinweis '$name': fwpkg-Metadaten sagen '$mTxt', Dateiname-Heuristik '$nTxt' - verwende Metadaten."
                    $kind.Deferred = $meta.Deferred
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
                    if ($tn) { $hit = $instMatches | Where-Object { (Get-NormFwVersion $_.Version) -eq $tn } | Select-Object -First 1 }
                    if ($hit) {
                        & $log "Uebersprungen $idx/$total : '$name' - Version $tgtVer bereits vorhanden (Inventar '$($hit.Name)' = $($hit.Version))"
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Aktuell $idx/$total : $name" })
                        # Zielversion ist aktiv im Inventar -> evtl. gesetzter
                        # Staged-Marker (Reboot erfolgt) kann entfernt werden.
                        if (-not $kind.Deferred -and -not $kind.IsIlo) { Clear-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $kind.Kind }
                        $skipNames += $name
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
                    if ($tnRom -and $curN -and $tnRom -eq $curN) {
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

                # Fallback (lokaler Marker) fuer bereits geflashte, aber noch nicht
                # per Reboot aktivierte, direkt geflashte Nicht-iLO-Komponenten,
                # falls die iLO-native Backup-ROM-Erkennung oben nicht greift (z.B.
                # aeltere iLO-Firmware ohne Oem.Hpe.Bios.Backup). iLO meldet die neue
                # Version erst NACH dem Reboot im Inventar; der Marker verhindert
                # erneutes Flashen bei einem zweiten Lauf VOR dem Reboot.
                if (-not $kind.Deferred -and -not $kind.IsIlo -and $tgtVer) {
                    $tnStg = Get-NormFwVersion $tgtVer
                    if ($tnStg -and (Test-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $kind.Kind -TargetNorm $tnStg)) {
                        & $log "Uebersprungen $idx/$total : '$name' - Version $tgtVer wurde bereits geflasht (Redundant ROM), Aktivierung beim Reboot ausstehend."
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name (Reboot)" })
                        $stagedNames += $name
                        $rebootNeeded = $true
                        continue
                    }
                }

                try {
                    # UpdateService bereit? (nur bei wirklich laufendem Flash abbrechen)
                    $st = ILO-GetUpdateState -Ilo $ilo -Token $token
                    & $log "UpdateService: State='$($st.State)', PushUri='$($st.PushUri)'"
                    if ($st.State -match '^(Uploading|Verifying|Writing|Updating)$') {
                        throw "iLO ist nicht bereit (State='$($st.State)') - anderer Flash laeuft?"
                    }
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

                        # Task in der UpdateTaskQueue anlegen/sicherstellen -> Aktivierung
                        # beim naechsten Reboot/POST. ILO-CreateUpdateTask ist idempotent:
                        # existiert bereits ein Task fuer die Datei, wird dieser akzeptiert;
                        # sonst wird er angelegt (mit Verifikation + Payload-Fallbacks).
                        & $log "Stelle Update-Task in der Queue sicher fuer '$name'..."
                        $taskUri = ILO-CreateUpdateTask -Ilo $ilo -Token $token -ComponentFileName $name -UpdatableBy $kind.UpdatableBy -LogCb $log
                        & $log "Update-Task in Queue bestaetigt fuer '$name' (UpdatableBy=$(($kind.UpdatableBy) -join '+'))$(if($taskUri){" -> $taskUri"})"
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Vorgemerkt $idx/$total : $name" })
                    } else {
                        # iLO/ROM: direkt online flashen.
                        ILO-UploadComponent -Ilo $ilo -Token $token -PushUri $pushUri -FilePath $comp -UpdateRepository:$false -UpdateTarget:$true -ProgressCb $cb | Out-Null
                        & $log "Upload $idx/$total abgeschlossen, starte Flash..."
                        $uiQueue.Enqueue(@{ Type = 'PHASE'; Ilo = $ilo; Phase = "Flash $idx/$total : $name" })
                        $fcb = { param($pct, $state) $uiQueue.Enqueue(@{ Type = 'PROGRESS'; Ilo = $ilo; Percent = $pct }) }.GetNewClosure()
                        $wf = ILO-WaitForFlash -Ilo $ilo -Token $token -User $user -Pass $pass -TimeoutSec 2400 -ProgressCb $fcb
                        $token = $wf.Token
                        & $log "Flash $idx/$total fertig: '$name' (End-State='$($wf.State)')"
                        # Direkt geflashte Nicht-iLO-Komponente (v.a. BIOS/System ROM):
                        # neue Version liegt in der Redundant/Backup-ROM und wird erst
                        # beim Reboot aktiv - lokal vormerken, damit ein zweiter Lauf
                        # VOR dem Reboot nicht erneut flasht.
                        if (-not $kind.IsIlo -and $tgtVer) {
                            $tnMk = Get-NormFwVersion $tgtVer
                            if ($tnMk) { Set-StagedFlash -StateDir $scriptFolder -Serial $info.Serial -Kind $kind.Kind -File $name -TargetNorm $tnMk }
                        }
                    }

                    # iLO ist nach Selbst-Reset sofort aktiv; alles andere braucht einen Reboot.
                    if ($kind.Deferred) { $stagedNames += $name } else { $okNames += $name }
                    # Reboot-Bedarf: deferred-Komponenten und Nicht-iLO immer; bei
                    # direkt geflashten Nicht-iLO-Komponenten entscheidet - falls
                    # vorhanden - die fwpkg-Metadaten (RebootRequired).
                    if ($kind.Deferred) {
                        $rebootNeeded = $true
                    } elseif (-not $kind.IsIlo) {
                        if ($null -eq $meta.Reboot -or $meta.Reboot) { $rebootNeeded = $true }
                    }
                }
                catch {
                    $cErr = $_.Exception.Message
                    & $log "FEHLER bei '$name': $cErr - Komponente uebersprungen, weiter mit naechster."
                    $failItems += "$name ($cErr)"
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
            if ($stagedNames.Count) { $parts += "$($stagedNames.Count) vorgemerkt (Reboot)" }
            if ($skipNames.Count)   { $parts += "$($skipNames.Count) aktuell" }
            if ($absentNames.Count) { $parts += "$($absentNames.Count) nicht verbaut" }
            if ($failItems.Count)   { $parts += "$($failItems.Count) Fehler" }
            $summary = if ($parts.Count) { ($parts -join ', ') } else { 'nichts zu tun' }
            $rebootTxt = if ($rebootNeeded) { 'Reboot fuer Aktivierung noetig' } else { 'Kein Reboot noetig' }
            $success = ($failItems.Count -eq 0)
            $phase = if (-not $success) { 'Teilweise' } elseif (($okNames.Count + $stagedNames.Count) -eq 0) { 'Aktuell' } else { 'Fertig' }
            $detail = "$summary - $rebootTxt"
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

    Start-Batch -Servers $servers -Worker $worker -ExtraArgs @{ components = $components; mode = $mode; baseDir = $fwPath } -MaxParallel ([int]$numPar.Value)
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

# Initiales Laden: zuletzt aus OneView erzeugte Servers.txt (falls vorhanden)
if (Test-Path -LiteralPath $serversFile) {
    Load-ServersFile -Path $serversFile
    Add-Log "Bereit. Servers.txt geladen ($($script:allServers.Count) Server). Fuer aktuelle Daten 'Server aus OneView laden'." ([System.Drawing.Color]::DarkBlue)
} else {
    Add-Log "Bereit. Noch keine Servers.txt - bitte 'Server aus OneView laden'. Log: $script:logFile"
}

[void]$form.ShowDialog()
