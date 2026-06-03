#Requires -Version 7.0
<#
.SYNOPSIS
    Headless Runner für HPE OneView Config-Backups (geplanter Task).

.DESCRIPTION
    Liest die Konfiguration aus backup_task_config.json (im gleichen Ordner),
    entschlüsselt die OneView-Credentials und die Backup-Passphrase mit dem
    AES-Key aus backup_task_key.bin und erzeugt parallel (Start-Job, eigene
    Prozesse = Modul-Isolation) Backups der HPE OneView Appliances.

    Unterstützt:
      - Versionsuebergreifend (OV 6.60, OV 11.x, ...): die OneView-Software-
        Version wird pro Appliance per /rest/version automatisch erkannt und
        das passende HPEOneView-PowerShell-Modul aus OneView_VersionMap.ps1
        gewaehlt. Pro benoetigtem Modul laeuft ein eigener Start-Job.
      - Optionale Übertragung der Backups per PSCP an einen Zielhost
      - Optionale Remote-Bereinigung per PLINK
      - Lokale Bereinigung älter als X Tage
      - Optionaler E-Mail-Versand einer Zusammenfassung (inkl. Fehler)

.NOTES
    Erfordert: PowerShell 7.x (Windows), HPEOneView-Module je nach Bedarf
    (z.B. HPEOneView.660 fuer OV 6.x, HPEOneView.1000 fuer OV 11.x).
#>

param(
    [string]$ConfigPath
)

# ---------------------------------------------------------------------------
# Pfade & Logging
# ---------------------------------------------------------------------------
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $MyInvocation.MyCommand.Path -Parent }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
if (-not $ConfigPath) { $ConfigPath = Join-Path $scriptDir 'backup_task_config.json' }
$keyFile = Join-Path $scriptDir 'backup_task_key.bin'
$credFile = Join-Path $scriptDir 'backup_task_cred.xml'

# Versions-/Modul-Tabelle einbinden (Resolve-OvModule, Get-OvVersionInfo)
. (Join-Path $scriptDir 'OneView_VersionMap.ps1')

function Resolve-ScriptPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if ([System.IO.Path]::IsPathRooted($Path)) { return $Path }
    return (Join-Path $scriptDir $Path)
}

$logDir = Join-Path $scriptDir 'Logs'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$runLog = Join-Path $logDir ("BackupTask_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Write-RunLog {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts][$Level] $Message"
    Add-Content -Path $runLog -Value $line -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Konfiguration laden
# ---------------------------------------------------------------------------
if (-not (Test-Path $ConfigPath)) {
    Write-RunLog "Konfigurationsdatei nicht gefunden: $ConfigPath" -Level ERROR
    throw "Konfigurationsdatei nicht gefunden: $ConfigPath"
}
$config = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
Write-RunLog "Konfiguration geladen: $ConfigPath"

if (-not (Test-Path $keyFile) -or -not (Test-Path $credFile)) {
    Write-RunLog "Credentials/Schlüssel fehlen (backup_task_key.bin / backup_task_cred.xml). Bitte Config-GUI ausführen." -Level ERROR
    throw "Credentials nicht konfiguriert."
}
$aesKey = [IO.File]::ReadAllBytes($keyFile)
$credXml = Import-Clixml -Path $credFile

try {
    $secPw = ConvertTo-SecureString -String $credXml.EncryptedPassword -Key $aesKey
}
catch {
    Write-RunLog "Passwort-Entschlüsselung fehlgeschlagen: $($_.Exception.Message)" -Level ERROR
    throw
}
$credential = New-Object System.Management.Automation.PSCredential($credXml.Username, $secPw)

# Backup-Passphrase entschlüsseln
if (-not $credXml.EncryptedPassphrase) {
    Write-RunLog "Keine Backup-Passphrase konfiguriert." -Level ERROR
    throw "Backup-Passphrase fehlt in $credFile"
}
try {
    $securePassphrase = ConvertTo-SecureString -String $credXml.EncryptedPassphrase -Key $aesKey
    $plainPassphrase = (New-Object System.Management.Automation.PSCredential('backup', $securePassphrase)).GetNetworkCredential().Password
}
catch {
    Write-RunLog "Passphrase-Entschlüsselung fehlgeschlagen: $($_.Exception.Message)" -Level ERROR
    throw
}

# ---------------------------------------------------------------------------
# TLS / Zertifikat
# ---------------------------------------------------------------------------
try {
    [System.Net.ServicePointManager]::SecurityProtocol = `
        [System.Net.SecurityProtocolType]::Tls12 -bor `
        [System.Net.SecurityProtocolType]::Tls13
}
catch {
    try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {}
}
try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch {}
$Global:SetLibraryBypassCertificatePolicy = $true

# ---------------------------------------------------------------------------
# Appliance-Liste einlesen (eine Datei, Versionserkennung erfolgt unten)
# ---------------------------------------------------------------------------
# Fallback: ggf. alte Konfigschluessel IPFile660/IPFile1110 mit auswerten,
# damit bestehende backup_task_config.json kompatibel bleibt.
$ipFileNew  = Resolve-ScriptPath $config.IPFile
$ipFileOld1 = Resolve-ScriptPath $config.IPFile660
$ipFileOld2 = Resolve-ScriptPath $config.IPFile1110

$rawAppliances = New-Object System.Collections.Generic.List[string]
foreach ($f in @($ipFileNew, $ipFileOld1, $ipFileOld2)) {
    if ($f -and (Test-Path $f)) {
        $entries = @(Get-Content $f | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') })
        foreach ($e in $entries) { if (-not $rawAppliances.Contains($e)) { $rawAppliances.Add($e) | Out-Null } }
    }
}
if (-not $ipFileNew -or -not (Test-Path $ipFileNew)) {
    if ($ipFileNew) { Write-RunLog "OneView IP-Datei nicht gefunden: $ipFileNew" -Level WARN }
    else            { Write-RunLog "Konfigschluessel 'IPFile' nicht gesetzt - nutze ggf. legacy IPFile660/IPFile1110." -Level WARN }
}

if ($rawAppliances.Count -eq 0) {
    Write-RunLog "Keine Appliances zu sichern." -Level ERROR
    throw "Keine Appliances konfiguriert."
}
Write-RunLog ("Appliances aus Konfiguration: {0}" -f $rawAppliances.Count)

# ---------------------------------------------------------------------------
# Versions-Auto-Erkennung + Gruppierung pro Modul
# ---------------------------------------------------------------------------
$applianceGroups = @{}   # Module -> @{ Module; Label; List }
$unresolved      = @()
foreach ($ip in $rawAppliances) {
    $info = Get-OvVersionInfo -Appliance $ip -TimeoutSec 10
    if ($info.Module) {
        $verLabel = if ($info.MajorMinor) { $info.MajorMinor } else { $info.SoftwareVersion }
        if (-not $applianceGroups.ContainsKey($info.Module)) {
            $applianceGroups[$info.Module] = [PSCustomObject]@{ Module = $info.Module; Label = $verLabel; List = @() }
        } else {
            $existing = $applianceGroups[$info.Module]
            if ($existing.Label -ne $verLabel) {
                $major = ($verLabel -split '\.')[0]
                $existing.Label = "$major.x"
            }
        }
        $applianceGroups[$info.Module].List += $ip
        Write-RunLog ("{0,-25} -> OV {1} (Modul {2})" -f $ip, $verLabel, $info.Module)
    } else {
        $unresolved += [PSCustomObject]@{ Appliance = $ip; Reason = $info.Error }
        Write-RunLog ("{0,-25} -> Versionserkennung fehlgeschlagen: {1}" -f $ip, $info.Error) -Level WARN
    }
}

# Modul-Verfuegbarkeit pruefen, andernfalls Gruppe verwerfen
foreach ($key in @($applianceGroups.Keys)) {
    $g = $applianceGroups[$key]
    if (-not (Get-Module -ListAvailable -Name $g.Module)) {
        Write-RunLog ("Modul '{0}' nicht installiert - {1} Appliance(s) (OV {2}) werden uebersprungen." -f $g.Module, $g.List.Count, $g.Label) -Level WARN
        foreach ($ip in $g.List) { $unresolved += [PSCustomObject]@{ Appliance = $ip; Reason = "Modul $($g.Module) nicht installiert" } }
        $applianceGroups.Remove($key)
    }
}

if ($applianceGroups.Count -eq 0) {
    Write-RunLog "Keine Appliance mit ermittelbarer Version / verfuegbarem Modul." -Level ERROR
    throw "Keine sicherungsfaehigen Appliances."
}

$totalCount = 0
foreach ($g in $applianceGroups.Values) { $totalCount += $g.List.Count }
Write-RunLog ("Gruppen: {0} | Sichere insgesamt {1} Appliance(s) | uebersprungen: {2}" -f $applianceGroups.Count, $totalCount, $unresolved.Count)

# ---------------------------------------------------------------------------
# Zielordner vorbereiten
# ---------------------------------------------------------------------------
$baseBackupDir = if ($config.BackupBaseDir) { Resolve-ScriptPath ([string]$config.BackupBaseDir) } else { Join-Path $scriptDir 'OneView_Backup' }
if (-not (Test-Path $baseBackupDir)) { New-Item -ItemType Directory -Path $baseBackupDir -Force | Out-Null }
$date = Get-Date -Format 'yyyy-MM-dd'
$folderPath = Join-Path $baseBackupDir $date
if (-not (Test-Path $folderPath)) { New-Item -ItemType Directory -Path $folderPath -Force | Out-Null }
$backupLogFile = Join-Path $baseBackupDir "Backup_Log_${date}.txt"
$errorLogFile = Join-Path $baseBackupDir "Error_Log_${date}.txt"

Write-RunLog "Zielordner: $folderPath"

# ---------------------------------------------------------------------------
# Batch-Job-Block (pro Modul, eigener Prozess)
# ---------------------------------------------------------------------------
$batchScript = {
    param(
        [string]$ApplianceListStr,
        [string]$ModuleName,
        [string]$VersionLabel,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$FolderPath,
        [string]$BaseBackupDir,
        [string]$Date,
        [string]$Passphrase,
        [string]$ErrorLogFile
    )

    $ApplianceList = @($ApplianceListStr -split '\|' | Where-Object { $_ })
    if ($ApplianceList.Count -eq 0) { return }

    try {
        [System.Net.ServicePointManager]::SecurityProtocol = `
            [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
    }
    catch { try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {} }
    try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch {}
    $Global:SetLibraryBypassCertificatePolicy = $true

    try {
        Import-Module $ModuleName -Force -ErrorAction Stop
        [PSCustomObject]@{ Type = 'LOG'; Message = "=== $ModuleName geladen - Backup für $($ApplianceList.Count) Appliance(s) (OV $VersionLabel) ===" }
    }
    catch {
        [PSCustomObject]@{ Type = 'MODULE_FAIL'; Message = "FEHLER: Konnte $ModuleName nicht laden: $($_.Exception.Message)" }
        foreach ($a in $ApplianceList) {
            [PSCustomObject]@{ Type = 'UPDATE'; Appliance = $a; Status = 'Fehler'; Detail = "Modul $ModuleName nicht ladbar" }
        }
        return
    }

    foreach ($appliance in $ApplianceList) {
        [PSCustomObject]@{ Type = 'PROGRESS'; Appliance = $appliance; VersionLabel = $VersionLabel }

        $currentFolder = Join-Path $FolderPath $appliance
        if (-not (Test-Path $currentFolder)) {
            try { New-Item -ItemType Directory -Path $currentFolder -Force -ErrorAction Stop | Out-Null }
            catch {
                [PSCustomObject]@{ Type = 'UPDATE'; Appliance = $appliance; Status = 'Fehler'; Detail = 'Ordner konnte nicht erstellt werden.' }
                continue
            }
        }
        Set-Location -Path $currentFolder

        [PSCustomObject]@{ Type = 'LOG'; Message = "Verbinde zu Appliance: $appliance" }

        $maxRetries = 2
        for ($attempt = 1; $attempt -le ($maxRetries + 1); $attempt++) {
            try {
                $passphraseSecure = ConvertTo-SecureString $Passphrase -AsPlainText -Force
                $backupJob = Start-Job -ScriptBlock {
                    param($h, $c, $loc, $pp, $m)
                    $Global:SetLibraryBypassCertificatePolicy = $true
                    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                    try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch {}
                    Import-Module $m -Force -ErrorAction Stop
                    Connect-OVMgmt -Hostname $h -Credential $c -ErrorAction Stop
                    New-OVBackup -Location $loc -Force -Passphrase $pp -ErrorAction Stop
                    Disconnect-OVMgmt
                } -ArgumentList $appliance, $Credential, $currentFolder, $passphraseSecure, $ModuleName
                Remove-Variable passphraseSecure -ErrorAction SilentlyContinue

                $backupDone = $backupJob | Wait-Job -Timeout 600
                if (-not $backupDone) {
                    $backupJob | Stop-Job -PassThru | Remove-Job -Force
                    throw "Timeout nach 600 Sekunden (Appliance antwortet nicht)"
                }
                if ($backupJob.State -eq 'Failed') {
                    $jobErr = ''
                    if ($backupJob.ChildJobs.Count -gt 0 -and $backupJob.ChildJobs[0].JobStateInfo.Reason) {
                        $jobErr = $backupJob.ChildJobs[0].JobStateInfo.Reason.Message
                    }
                    Remove-Job $backupJob -Force
                    if (-not $jobErr) { $jobErr = "Unbekannter Fehler im Backup-Job" }
                    throw $jobErr
                }
                # Fehlerstream und Output einsammeln (Error-Records aus Modul-Cmdlets)
                $jobErrors = @()
                $null = Receive-Job $backupJob -ErrorVariable jobErrors -ErrorAction SilentlyContinue
                Remove-Job $backupJob -Force
                if ($jobErrors -and $jobErrors.Count -gt 0) {
                    $errText = ($jobErrors | ForEach-Object { $_.ToString() }) -join ' | '
                    throw "Backup-Job meldete Fehler: $errText"
                }

                # Verifikation: Es muss tatsächlich eine Backup-Datei im Zielordner liegen (>0 Byte)
                $backupFiles = @(Get-ChildItem -Path $currentFolder -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.Extension -ne '.log' -and $_.Length -gt 0 })
                if ($backupFiles.Count -eq 0) {
                    throw "Backup-Job abgeschlossen, aber keine Backup-Datei im Ordner '$currentFolder' gefunden."
                }
                $backupFile = $backupFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1

                # Ab OV 11.20 enthaelt der HPE-Backup-Dateiname den Appliance-Hostname nicht mehr.
                # Damit die Zuordnung der Datei zur Appliance erhalten bleibt, wird der Hostname als Prefix vorangestellt.
                if ($backupFile.Name -notlike "$appliance*") {
                    $safeName = ($appliance -replace '[\\/:*?"<>|]', '_')
                    $newName  = "${safeName}_$($backupFile.Name)"
                    try {
                        $renamed = Rename-Item -LiteralPath $backupFile.FullName -NewName $newName -PassThru -ErrorAction Stop
                        $backupFile = Get-Item -LiteralPath $renamed.FullName
                    } catch {
                        [PSCustomObject]@{ Type = 'LOG'; Message = "WARNUNG: Backup-Datei konnte nicht umbenannt werden: $($_.Exception.Message)" }
                    }
                }
                $sizeMB = [math]::Round($backupFile.Length / 1MB, 2)

                [PSCustomObject]@{ Type = 'UPDATE'; Appliance = $appliance; Status = 'Erfolgreich'; Detail = "Backup erstellt: $($backupFile.Name) ($sizeMB MB, Versuch $attempt)." }
                break
            }
            catch {
                $errMsg = $_.Exception.Message
                if ($attempt -le $maxRetries) {
                    [PSCustomObject]@{ Type = 'LOG'; Message = "WARNUNG: $appliance Versuch $attempt fehlgeschlagen: $errMsg - Retry in 15s..." }
                    Start-Sleep -Seconds 15
                }
                else {
                    [PSCustomObject]@{ Type = 'UPDATE'; Appliance = $appliance; Status = 'Fehler'; Detail = "$errMsg (nach $attempt Versuchen)" }
                    try {
                        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Fehler bei Appliance ${appliance}: ${errMsg} (nach $attempt Versuchen)" |
                            Out-File -Append -FilePath $ErrorLogFile -Encoding UTF8
                    }
                    catch {}
                }
            }
            finally {
                Remove-Item -Path (Join-Path $currentFolder '*.log') -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Jobs starten (pro benoetigtem Modul ein eigener Start-Job)
# ---------------------------------------------------------------------------
$jobs = @()
foreach ($g in $applianceGroups.Values) {
    if ($g.List.Count -eq 0) { continue }
    $jobs += Start-Job -Name ("OV " + $g.Label) -ScriptBlock $batchScript -ArgumentList @(
        ($g.List -join '|'), $g.Module, $g.Label,
        $credential, $folderPath, $baseBackupDir, $date, $plainPassphrase, $errorLogFile
    )
}
Remove-Variable plainPassphrase -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# Ergebnisse sammeln
# ---------------------------------------------------------------------------
$results = @{}  # Appliance -> [PSCustomObject]@{ Status, Detail, Version }
foreach ($g in $applianceGroups.Values) {
    foreach ($a in $g.List) { $results[$a] = [PSCustomObject]@{ Status = 'Ausstehend'; Detail = ''; Version = $g.Label } }
}
foreach ($u in $unresolved) {
    if (-not $results.ContainsKey($u.Appliance)) {
        $results[$u.Appliance] = [PSCustomObject]@{ Status = 'Fehler'; Detail = $u.Reason; Version = '?' }
    }
}

$jobs | ForEach-Object { $_ | Wait-Job | Out-Null }
foreach ($job in $jobs) {
    $messages = @(Receive-Job $job -ErrorAction SilentlyContinue)
    foreach ($msg in $messages) {
        if ($null -eq $msg -or $null -eq $msg.Type) { continue }
        switch ($msg.Type) {
            'LOG' { Write-RunLog $msg.Message }
            'MODULE_FAIL' { Write-RunLog $msg.Message -Level ERROR }
            'PROGRESS' { Write-RunLog "Bearbeite $($msg.Appliance) (OV $($msg.VersionLabel))" }
            'UPDATE' {
                if ($results.ContainsKey($msg.Appliance)) {
                    $results[$msg.Appliance].Status = $msg.Status
                    $results[$msg.Appliance].Detail = $msg.Detail
                }
                $lvl = if ($msg.Status -ieq 'Erfolgreich') { 'INFO' } else { 'ERROR' }
                Write-RunLog "Appliance $($msg.Appliance): $($msg.Status) - $($msg.Detail)" -Level $lvl
            }
        }
    }
    if ($job.State -eq 'Failed' -and $job.ChildJobs.Count -gt 0 -and $job.ChildJobs[0].JobStateInfo.Reason) {
        Write-RunLog "FEHLER im Job $($job.Name): $($job.ChildJobs[0].JobStateInfo.Reason.Message)" -Level ERROR
    }
    Write-RunLog "--- $($job.Name) Backup-Batch abgeschlossen ---"
    Remove-Job $job -Force
}

# Zusammenfassung
$okCount = ($results.Values | Where-Object { $_.Status -ieq 'Erfolgreich' }).Count
$errCount = ($results.Values | Where-Object { $_.Status -ieq 'Fehler' }).Count
$pendCount = ($results.Values | Where-Object { $_.Status -ieq 'Ausstehend' }).Count
Write-RunLog "Backup-Ergebnis: $okCount erfolgreich, $errCount Fehler, $pendCount ohne Ergebnis."

# Log in Backup-Log-Datei
try {
    $logBody = @()
    $logBody += "OneView Backup Task Run - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $logBody += "Erfolgreich: $okCount | Fehler: $errCount | Ausstehend: $pendCount"
    $logBody += ''
    foreach ($key in $results.Keys | Sort-Object) {
        $r = $results[$key]
        $logBody += "[$($r.Status)] $key (OV $($r.Version)) - $($r.Detail)"
    }
    $logBody | Out-File -Append -FilePath $backupLogFile -Encoding UTF8
}
catch {}

# ---------------------------------------------------------------------------
# Optional: PSCP-Übertragung
# ---------------------------------------------------------------------------
if ($config.TransferEnabled) {
    $pscpExe = Resolve-ScriptPath ([string]$config.PscpPath)
    if (-not $pscpExe) { $pscpExe = Join-Path $scriptDir 'tools\pscp.exe' }
    if (-not (Test-Path $pscpExe)) {
        Write-RunLog "pscp.exe nicht gefunden: $pscpExe - Übertragung übersprungen." -Level WARN
    }
    elseif (-not $config.TransferHost -or -not $config.TransferUser -or -not $config.TransferRemotePath) {
        Write-RunLog "Transfer aktiviert, aber Host/User/RemotePath nicht vollständig - übersprungen." -Level WARN
    }
    else {
        try {
            # Transfer-Passwort (falls eigenes gesetzt, sonst OneView-Passwort)
            $transferPw = $null
            if ($credXml.EncryptedTransferPassword) {
                try {
                    $secTp = ConvertTo-SecureString -String $credXml.EncryptedTransferPassword -Key $aesKey
                    $transferPw = (New-Object System.Management.Automation.PSCredential('t', $secTp)).GetNetworkCredential().Password
                }
                catch { Write-RunLog "Transfer-Passwort konnte nicht entschlüsselt werden - fallback auf OV-Passwort." -Level WARN }
            }
            if (-not $transferPw) {
                $transferPw = $credential.GetNetworkCredential().Password
            }

            # Nur den heutigen Datumsordner übertragen (nicht das BaseDir mit Log-Dateien)
            $source = Join-Path $baseBackupDir $date
            if (-not (Test-Path $source -PathType Container)) {
                throw "Quell-Ordner für Transfer nicht gefunden: $source"
            }
            $remotePath = ([string]$config.TransferRemotePath).TrimEnd('/')
            $destination = "$($config.TransferUser)@$($config.TransferHost):$remotePath/"
            $pscpArgs = @('-r', '-batch', '-pw', $transferPw, $source, $destination)
            Write-RunLog "Starte PSCP-Übertragung: $source -> $destination"
            $pscpOut = Join-Path $env:TEMP "pscp_out_$PID.log"
            $pscpErr = Join-Path $env:TEMP "pscp_err_$PID.log"
            $p = Start-Process -FilePath $pscpExe -ArgumentList $pscpArgs -NoNewWindow -PassThru `
                    -RedirectStandardOutput $pscpOut -RedirectStandardError $pscpErr `
                    -WorkingDirectory (Split-Path $pscpExe -Parent)
            if (-not $p.WaitForExit(600000)) {
                $p.Kill()
                throw "PSCP Timeout nach 10 Minuten"
            }
            $pscpStdOut = (Test-Path $pscpOut) ? ((Get-Content $pscpOut -Raw -ErrorAction SilentlyContinue) ?? '') : ''
            $pscpStdErr = (Test-Path $pscpErr) ? ((Get-Content $pscpErr -Raw -ErrorAction SilentlyContinue) ?? '') : ''
            Remove-Item $pscpOut, $pscpErr -Force -ErrorAction SilentlyContinue
            if ($p.ExitCode -ne 0) {
                $detail = ($pscpStdErr + $pscpStdOut).Trim()
                if (-not $detail) { $detail = '(keine Ausgabe)' }
                throw "PSCP ExitCode $($p.ExitCode): $detail"
            }
            Write-RunLog "PSCP-Übertragung abgeschlossen."

            # Remote-Bereinigung per PLINK
            if ($config.RemoteCleanupEnabled) {
                $plinkExe = Resolve-ScriptPath ([string]$config.PlinkPath)
                if (-not $plinkExe) { $plinkExe = Join-Path $scriptDir 'tools\plink.exe' }
                if (-not (Test-Path $plinkExe)) {
                    Write-RunLog "plink.exe nicht gefunden: $plinkExe - Remote-Cleanup übersprungen." -Level WARN
                }
                else {
                    $days = if ($config.RemoteRetentionDays) { [int]$config.RemoteRetentionDays } else { 30 }
                    $remotePath = [string]$config.TransferRemotePath
                    $remoteCmd = "find $remotePath -mindepth 1 -depth -mtime +$days -exec rm -rf {} \;"
                    $plinkArgs = @('-batch', '-ssh', '-pw', $transferPw, "$($config.TransferUser)@$($config.TransferHost)", $remoteCmd)
                    Write-RunLog "Starte Remote-Cleanup (>$days Tage) auf $($config.TransferHost)"
                    $plinkOut = Join-Path $env:TEMP "plink_out_$PID.log"
                    $plinkErr = Join-Path $env:TEMP "plink_err_$PID.log"
                    $pl = Start-Process -FilePath $plinkExe -ArgumentList $plinkArgs -NoNewWindow -PassThru `
                            -RedirectStandardOutput $plinkOut -RedirectStandardError $plinkErr `
                            -WorkingDirectory (Split-Path $plinkExe -Parent)
                    if (-not $pl.WaitForExit(120000)) {
                        $pl.Kill()
                        Write-RunLog "PLINK Timeout nach 2 Minuten." -Level WARN
                    }
                    else {
                        $plErrText = (Test-Path $plinkErr) ? ((Get-Content $plinkErr -Raw -ErrorAction SilentlyContinue) ?? '') : ''
                        if ($pl.ExitCode -ne 0) {
                            Write-RunLog "PLINK ExitCode $($pl.ExitCode): $($plErrText.Trim())" -Level WARN
                        }
                        else {
                            Write-RunLog "Remote-Cleanup abgeschlossen."
                        }
                    }
                    Remove-Item $plinkOut, $plinkErr -Force -ErrorAction SilentlyContinue
                }
            }
        }
        catch {
            Write-RunLog "Transfer-Fehler: $($_.Exception.Message)" -Level ERROR
        }
        finally {
            Remove-Variable transferPw -ErrorAction SilentlyContinue
        }
    }
}

# ---------------------------------------------------------------------------
# Lokale Bereinigung
# ---------------------------------------------------------------------------
try {
    $retention = if ($config.LocalRetentionDays) { [int]$config.LocalRetentionDays } else { 5 }
    Get-ChildItem -Path $baseBackupDir -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$retention) } |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Write-RunLog "Lokale Bereinigung: Dateien älter als $retention Tage entfernt."
}
catch {
    Write-RunLog "Fehler bei lokaler Bereinigung: $($_.Exception.Message)" -Level WARN
}

# ---------------------------------------------------------------------------
# E-Mail
# ---------------------------------------------------------------------------
function ConvertTo-HtmlEncoded {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Text)
}

function Build-BackupHtmlBody {
    param([hashtable]$Results, [int]$Ok, [int]$Err, [int]$Pend)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html><html><head><meta charset="utf-8"/><style>')
    [void]$sb.AppendLine('body{font-family:Segoe UI,Arial,sans-serif;font-size:12px;color:#222}')
    [void]$sb.AppendLine('h2{margin:0 0 6px 0;font-size:15px}')
    [void]$sb.AppendLine('.kpi span{display:inline-block;padding:2px 8px;margin-right:6px;border-radius:3px;font-weight:bold;color:#fff}')
    [void]$sb.AppendLine('.k-ok{background:#1e8449}.k-err{background:#c0392b}.k-pend{background:#7f8c8d}')
    [void]$sb.AppendLine('table.at{border-collapse:collapse;width:100%;margin-top:8px}')
    [void]$sb.AppendLine('table.at th,table.at td{border:1px solid #ccc;padding:4px 6px;text-align:left;font-size:11.5px}')
    [void]$sb.AppendLine('table.at th{background:#34495e;color:#fff}')
    [void]$sb.AppendLine('tr.err td{background:#fdecea}tr.ok td{background:#eafaf1}')
    [void]$sb.AppendLine('</style></head><body>')
    [void]$sb.AppendLine("<h2>OneView Backup - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</h2>")
    [void]$sb.AppendLine('<div class="kpi">')
    [void]$sb.AppendLine("<span class='k-ok'>Erfolgreich: $Ok</span>")
    [void]$sb.AppendLine("<span class='k-err'>Fehler: $Err</span>")
    if ($Pend -gt 0) { [void]$sb.AppendLine("<span class='k-pend'>Ausstehend: $Pend</span>") }
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('<table class="at"><tr><th>Appliance</th><th>OV-Version</th><th>Status</th><th>Details</th></tr>')
    foreach ($key in $Results.Keys | Sort-Object) {
        $r = $Results[$key]
        $cls = if ($r.Status -ieq 'Fehler') { ' class="err"' } elseif ($r.Status -ieq 'Erfolgreich') { ' class="ok"' } else { '' }
        [void]$sb.AppendLine("<tr$cls><td>$(ConvertTo-HtmlEncoded $key)</td><td>$(ConvertTo-HtmlEncoded $r.Version)</td><td>$(ConvertTo-HtmlEncoded $r.Status)</td><td>$(ConvertTo-HtmlEncoded $r.Detail)</td></tr>")
    }
    [void]$sb.AppendLine('</table></body></html>')
    return $sb.ToString()
}

if ($config.SendEmail) {
    if ($config.OnlyOnErrors -and $errCount -eq 0 -and $pendCount -eq 0) {
        Write-RunLog "OnlyOnErrors aktiv und keine Fehler - keine E-Mail."
    }
    elseif (-not $config.SmtpServer -or -not $config.MailFrom -or -not $config.MailTo) {
        Write-RunLog "SMTP-Parameter unvollständig - keine E-Mail." -Level WARN
    }
    else {
        try {
            try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { param($s, $c, $ch, $e) return $true } } catch {}
            $smtpPort = if ($config.SmtpPort) { [int]$config.SmtpPort } else { 25 }
            $smtpClient = New-Object Net.Mail.SmtpClient($config.SmtpServer, $smtpPort)
            $smtpClient.EnableSsl = [bool]$config.UseSsl
            if ($config.SmtpUser -and $config.SmtpPasswordEncrypted) {
                try {
                    $sp = ConvertTo-SecureString -String $config.SmtpPasswordEncrypted -Key $aesKey
                    $smtpClient.Credentials = (New-Object System.Management.Automation.PSCredential($config.SmtpUser, $sp)).GetNetworkCredential()
                }
                catch { Write-RunLog "SMTP-Passwort nicht entschlüsselbar." -Level WARN }
            }
            $subjectPrefix = if ($config.SubjectPrefix) { $config.SubjectPrefix } else { '[OneView Backup]' }
            $hostName = try { $env:COMPUTERNAME } catch { 'unknown' }
            $mailMessage = New-Object System.Net.Mail.MailMessage
            $mailMessage.From = New-Object System.Net.Mail.MailAddress($config.MailFrom)
            foreach ($rcpt in @($config.MailTo -split '\s*;\s*|\s*,\s*' | Where-Object { $_ })) {
                $mailMessage.To.Add($rcpt)
            }
            $mailMessage.Subject = "$subjectPrefix OK=$okCount Err=$errCount ($hostName)"
            $mailMessage.Body = Build-BackupHtmlBody -Results $results -Ok $okCount -Err $errCount -Pend $pendCount
            $mailMessage.IsBodyHtml = $true
            $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
            $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8
            if ($errCount -gt 0) { $mailMessage.Priority = [System.Net.Mail.MailPriority]::High }
            $smtpClient.Send($mailMessage)
            $mailMessage.Dispose()
            $smtpClient.Dispose()
            Write-RunLog "E-Mail versendet an: $(($config.MailTo -split '\s*;\s*|\s*,\s*' | Where-Object { $_ }) -join ', ')"
        }
        catch {
            Write-RunLog "E-Mail-Versand fehlgeschlagen: $($_.Exception.Message)" -Level ERROR
        }
    }
}
else {
    Write-RunLog "E-Mail-Versand deaktiviert."
}

Write-RunLog "Lauf beendet."
