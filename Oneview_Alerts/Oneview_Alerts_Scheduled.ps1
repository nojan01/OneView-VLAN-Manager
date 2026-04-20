<#
.SYNOPSIS
    Headless Runner für OneView Alerts-Abfrage (geplanter Task).

.DESCRIPTION
    Liest die Konfiguration aus alerts_task_config.json (im gleichen Ordner),
    entschlüsselt die OneView-Credentials mit dem AES-Key aus alerts_task_key.bin,
    fragt Alerts aller Appliances ab (Modul-Isolation per Start-Job für
    HPEOneView.660 / HPEOneView.1000), schreibt Log-Dateien und verschickt
    eine E-Mail mit Zusammenfassung und Fehlern.

    Das Script ist explizit für unattended-Ausführung als geplanter Task
    (auch als SYSTEM-User) konzipiert.

.NOTES
    Erfordert: PowerShell 7.x (Windows), Module HPEOneView.660 und/oder
    HPEOneView.1000 (je nach Appliance-Versionen).
#>

param(
    [string]$ConfigPath
)

# ---------------------------------------------------------------------------
# Pfade
# ---------------------------------------------------------------------------
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $MyInvocation.MyCommand.Path -Parent }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
if (-not $ConfigPath) { $ConfigPath = Join-Path $scriptDir 'alerts_task_config.json' }
$keyFile = Join-Path $scriptDir 'alerts_task_key.bin'
$credFile = Join-Path $scriptDir 'alerts_task_cred.xml'
$knownIssuesFile = Join-Path $scriptDir 'KnownIssues.txt'
$logDir = Join-Path $scriptDir 'Logs'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$runLog = Join-Path $logDir ("AlertsTask_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

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

# Credentials laden (AES-verschlüsselt, systemuser-tauglich)
if (-not (Test-Path $keyFile) -or -not (Test-Path $credFile)) {
    Write-RunLog "Credentials/Schlüssel fehlen (alerts_task_key.bin / alerts_task_cred.xml). Bitte Config-GUI ausführen." -Level ERROR
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

# ---------------------------------------------------------------------------
# Zertifikat & TLS
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
# Appliance-Dateien auswerten
# ---------------------------------------------------------------------------
$applianceFiles = @()
switch ($config.ApplianceMode) {
    'GOV' { $applianceFiles = @(Join-Path $scriptDir 'Oneview_GOV.txt') }
    'DIV' { $applianceFiles = @(Join-Path $scriptDir 'Oneview_DIV.txt') }
    'BOTH' { $applianceFiles = @((Join-Path $scriptDir 'Oneview_GOV.txt'), (Join-Path $scriptDir 'Oneview_DIV.txt')) }
    default {
        Write-RunLog "Unbekannter ApplianceMode '$($config.ApplianceMode)'" -Level ERROR
        throw "Unbekannter ApplianceMode"
    }
}
$appliances = @()
foreach ($file in $applianceFiles) {
    if (-not (Test-Path $file)) {
        Write-RunLog "Appliance-Datei fehlt: $file" -Level WARN
        continue
    }
    $appliances += @(Get-Content -Path $file -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') })
}
if ($appliances.Count -eq 0) {
    Write-RunLog "Keine Appliances zum Abfragen gefunden." -Level ERROR
    throw "Keine Appliances"
}
Write-RunLog ("Abzufragende Appliances: {0}" -f ($appliances -join ', '))

# ---------------------------------------------------------------------------
# Bekannte Issues laden
# ---------------------------------------------------------------------------
$knownIssues = @()
if ((Test-Path $knownIssuesFile) -and $config.HideKnown) {
    $knownIssues = @(Get-Content -Path $knownIssuesFile |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') })
}

function Test-HideAlert {
    param([object]$Alert, [string]$Appliance)
    if (-not $knownIssues -or $knownIssues.Count -eq 0) { return $false }
    $vals = @()
    foreach ($p in 'description', 'message', 'alertTypeID', 'resourceName', 'name') {
        try {
            if ($Alert.PSObject.Properties.Match($p).Count -gt 0 -and $Alert.$p) { $vals += [string]$Alert.$p }
        }
        catch {}
    }
    try {
        if ($Alert.PSObject.Properties.Match('associatedResource').Count -gt 0 -and $Alert.associatedResource -and $Alert.associatedResource.resourceName) {
            $vals += [string]$Alert.associatedResource.resourceName
        }
    }
    catch {}
    try { $vals += ($Alert | Out-String) } catch {}
    $norm = (($vals -join ' ') -replace '\s+', ' ').Trim()
    $normLower = $norm.ToLowerInvariant()
    foreach ($pattern in $knownIssues) {
        $p = $pattern.Trim(); if (-not $p) { continue }
        if ($p.StartsWith('~')) {
            try { if ($norm -imatch $p.Substring(1)) { return $true } } catch {}
        }
        elseif ($p.StartsWith('^') -or $p.EndsWith('$')) {
            try { if ($norm -imatch $p) { return $true } } catch {}
        }
        else {
            if ($normLower.Contains(($p -replace '\s+', ' ').ToLowerInvariant())) { return $true }
        }
    }
    return $false
}

# ---------------------------------------------------------------------------
# OV-Versionserkennung
# ---------------------------------------------------------------------------
function Get-ApplianceOVModule {
    param([string]$Appliance)
    $module = 'HPEOneView.660'
    $softwareVersion = $null
    try {
        $headers = @{ 'X-API-Version' = '3800' }
        $resp = Invoke-RestMethod -Uri "https://$Appliance/rest/version" -Method Get -Headers $headers -TimeoutSec 10 -SkipCertificateCheck -ErrorAction Stop
        if ($resp.softwareVersion) {
            $softwareVersion = [string]$resp.softwareVersion
            if ($softwareVersion -match '^(\d+)') {
                $major = [int]$Matches[1]
                $module = if ($major -ge 7) { 'HPEOneView.1000' } else { 'HPEOneView.660' }
            }
        }
        elseif ($resp.currentVersion) {
            $api = 0; [int]::TryParse([string]$resp.currentVersion, [ref]$api) | Out-Null
            $module = if ($api -gt 3800) { 'HPEOneView.1000' } else { 'HPEOneView.660' }
        }
    }
    catch {}
    [PSCustomObject]@{ Appliance = $Appliance; Module = $module; SoftwareVersion = $softwareVersion }
}

$have660 = [bool](Get-Module -ListAvailable -Name 'HPEOneView.660')
$have1000 = [bool](Get-Module -ListAvailable -Name 'HPEOneView.1000')
if (-not ($have660 -or $have1000)) {
    Write-RunLog "Kein HPEOneView-Modul installiert." -Level ERROR
    throw "Kein HPEOneView-Modul installiert."
}

# Zeitraum
$timespan = if ($config.RangeUnit -eq 'Hours') { New-TimeSpan -Hours ([int]$config.RangeValue) } else { New-TimeSpan -Days ([int]$config.RangeValue) }
$cutoff = (Get-Date) - $timespan

# Versionen ermitteln
$applianceInfos = $appliances | ForEach-Object { Get-ApplianceOVModule -Appliance $_ }
foreach ($i in $applianceInfos) {
    $v = if ($i.SoftwareVersion) { $i.SoftwareVersion } else { '(Fallback)' }
    Write-RunLog "$($i.Appliance): Version $v -> $($i.Module)"
}
$byModule = $applianceInfos | Group-Object -Property Module

# ---------------------------------------------------------------------------
# Batch-Job-Block (pro Modul)
# ---------------------------------------------------------------------------
$batchScript = {
    param([string]$ModuleName, [string[]]$ApplianceList, [System.Management.Automation.PSCredential]$Credential, [timespan]$Timespan, [datetime]$Cutoff, [bool]$OwnerUnknownOnly)
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = `
            [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
    }
    catch { try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {} }
    try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch {}
    $Global:SetLibraryBypassCertificatePolicy = $true
    try { Import-Module $ModuleName -Force -ErrorAction Stop } catch {
        foreach ($a in $ApplianceList) {
            [PSCustomObject]@{ Type = 'RESULT'; Appliance = $a; Module = $ModuleName; Alerts = @(); Error = "Modul $ModuleName Ladefehler: $($_.Exception.Message)" }
        }
        return
    }
    function Get-Ts2 { param($A) foreach ($p in 'created', 'date', 'EventTimestamp') { if ($A.PSObject.Properties.Match($p).Count -gt 0 -and $A.$p) { try { return [datetime]$A.$p } catch {} } } return $null }
    function Test-OwnerUnknown {
        param($A)
        # Owner in OneView = assignedToUser. "unknown" = leer / None / Unassigned / unknown
        $val = $null
        foreach ($p in 'assignedToUser', 'owner', 'assignedTo') {
            if ($A.PSObject.Properties.Match($p).Count -gt 0 -and $null -ne $A.$p) {
                $val = [string]$A.$p
                break
            }
        }
        if ([string]::IsNullOrWhiteSpace($val)) { return $true }
        return ($val -match '^(unknown|unassigned|none)$')
    }
    foreach ($appliance in $ApplianceList) {
        $alerts = @(); $errorText = $null; $connected = $false
        try {
            Connect-OVMgmt -Hostname $appliance -Credential $Credential -ErrorAction Stop | Out-Null
            $connected = $true
            try {
                $alerts = @((Get-OVAlert -severity Critical -Timespan $Timespan -AlertState active)) +
                @((Get-OVAlert -severity Warning  -Timespan $Timespan -AlertState active)) +
                @((Get-OVAlert -severity Critical -Timespan $Timespan -AlertState locked)) +
                @((Get-OVAlert -severity Warning  -Timespan $Timespan -AlertState locked))
            }
            catch {
                $base = @(Get-OVAlert -severity Critical, Warning -AlertState active, locked)
                $alerts = @($base | Where-Object { (Get-Ts2 $_) -ge $Cutoff -and ($_.Severity -ieq 'Critical' -or $_.Severity -ieq 'Warning') -and ($_.AlertState -ieq 'Active' -or $_.AlertState -ieq 'Locked') })
            }
            if ($OwnerUnknownOnly) {
                # Nur Alerts ohne zugewiesenen Owner (Owner = "unknown")
                $alerts = @($alerts | Where-Object { Test-OwnerUnknown $_ })
            }
        }
        catch { $errorText = "$($_.Exception.Message)" }
        finally { if ($connected) { try { Disconnect-OVMgmt -ErrorAction SilentlyContinue | Out-Null } catch {} } }
        [PSCustomObject]@{ Type = 'RESULT'; Appliance = $appliance; Module = $ModuleName; Alerts = $alerts; Error = $errorText }
    }
}

$jobs = @()
foreach ($grp in $byModule) {
    $mod = $grp.Name
    $installed = if ($mod -eq 'HPEOneView.660') { $have660 } elseif ($mod -eq 'HPEOneView.1000') { $have1000 } else { $false }
    if (-not $installed) {
        Write-RunLog "Modul $mod nicht installiert - Appliances übersprungen: $($grp.Group.Appliance -join ', ')" -Level WARN
        continue
    }
    $ownerUnknownOnly = if ($null -ne $config.OwnerUnknownOnly) { [bool]$config.OwnerUnknownOnly } else { $false }
    $jobs += Start-Job -Name $mod -ScriptBlock $batchScript -ArgumentList @($mod, @($grp.Group.Appliance), $credential, $timespan, $cutoff, $ownerUnknownOnly)
}

$results = @{}
$jobs | ForEach-Object { $_ | Wait-Job | Out-Null }
foreach ($job in $jobs) {
    $out = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
    foreach ($o in $out) {
        if ($o -and $o.Type -eq 'RESULT') { $results[$o.Appliance] = $o }
    }
    try { Remove-Job -Job $job -Force } catch {}
}

# ---------------------------------------------------------------------------
# Auswertung & Zusammenfassung
# ---------------------------------------------------------------------------
$allAlerts = @()
$errorDetails = @()
$summaryLines = @()
foreach ($appliance in $appliances) {
    $summaryLines += "============================================"
    $summaryLines += "Appliance: $appliance"
    $res = $results[$appliance]
    if (-not $res) {
        $msg = "Fehler: Kein Ergebnis für $appliance (Modul fehlt?)"
        $summaryLines += $msg; $errorDetails += $msg
        Write-RunLog $msg -Level ERROR
        continue
    }
    if ($res.Error) {
        $msg = "Fehler: $appliance - $($res.Error)"
        $summaryLines += $msg; $errorDetails += $msg
        Write-RunLog $msg -Level ERROR
        continue
    }
    $alerts = @($res.Alerts)
    if ($alerts.Count -eq 0) {
        $summaryLines += "Keine Alarme für $appliance gefunden."
        continue
    }
    $filtered = $alerts
    $hiddenCount = 0
    if ($config.HideKnown) {
        $filtered = foreach ($al in $alerts) {
            if (Test-HideAlert -Alert $al -Appliance $appliance) { $hiddenCount++; continue }
            $al
        }
    }
    $allAlerts += ($filtered | ForEach-Object {
            $res2 = $null
            if ($_.PSObject.Properties.Match('resourceName').Count -gt 0 -and $_.resourceName) { $res2 = $_.resourceName }
            elseif ($_.PSObject.Properties.Match('associatedResource').Count -gt 0 -and $_.associatedResource -and $_.associatedResource.resourceName) { $res2 = $_.associatedResource.resourceName }
            [PSCustomObject]@{
                Appliance = $appliance
                Severity  = $_.Severity
                State     = $_.AlertState
                Created   = $_.created
                Owner     = if ($_.PSObject.Properties.Match('assignedToUser').Count -gt 0) { $_.assignedToUser } else { $null }
                Resource  = $res2
                Type      = $_.alertTypeID
                Message   = if ($_.message) { $_.message } else { $_.description }
            }
        })
    $grouped = $filtered | Group-Object -Property { "$($_.Severity) - $($_.AlertState)" }
    $summaryLines += "Alarm-Zusammenfassung für $($appliance):"
    foreach ($g in $grouped) { $summaryLines += "[$($g.Name)] : $($g.Count)" }
    $summaryLines += "Details (gefiltert): $($filtered.Count)"
    if ($hiddenCount -gt 0) { $summaryLines += "($hiddenCount bekannte Issue(s) ausgeblendet)" }
}

# Limit auf MaxDetails
$ordered = $allAlerts | Sort-Object `
@{ Expression = { if ($_.Severity -ieq 'Critical') { 0 } elseif ($_.Severity -ieq 'Warning') { 1 } else { 2 } } }, `
@{ Expression = { if ($_.State -ieq 'Active') { 0 } else { 1 } } }, `
@{ Expression = { try { [datetime]$_.Created } catch { [datetime]::MinValue } }; Descending = $true }
$limit = [int]$config.MaxDetails
if ($limit -gt 0 -and $ordered.Count -gt $limit) {
    $ordered = $ordered | Select-Object -First $limit
}

$detailLines = @()
foreach ($a in $ordered) {
    $ts = ''
    try { if ($a.Created) { $ts = ([datetime]$a.Created).ToString('yyyy-MM-dd HH:mm:ss') } } catch {}
    $detailLines += "[{0}][{1}] {2} {3} {4} ({5}): {6}" -f $a.Severity, $a.State, $ts, $a.Appliance, $a.Resource, $a.Type, $a.Message
}

$critCount = ($allAlerts | Where-Object { $_.Severity -ieq 'Critical' }).Count
$warnCount = ($allAlerts | Where-Object { $_.Severity -ieq 'Warning' }).Count
Write-RunLog "Abfrage fertig: Critical=$critCount, Warning=$warnCount, Fehler=$($errorDetails.Count)"

# Log-Dateien schreiben
$logBody = @()
$logBody += "OneView Alerts Task Run - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$logBody += ""
$logBody += $summaryLines
$logBody += ""
if ($detailLines.Count -gt 0) {
    $logBody += "=== Alert-Details ==="
    $logBody += $detailLines
}
if ($errorDetails.Count -gt 0) {
    $logBody += ""
    $logBody += "=== Fehlerdetails ==="
    $logBody += $errorDetails
}
$logBody | Set-Content -Path (Join-Path $logDir 'Alerts_current.txt') -Encoding UTF8

# ---------------------------------------------------------------------------
# E-Mail senden
# ---------------------------------------------------------------------------
function ConvertTo-HtmlEncoded {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Text)
}

function Build-AlertsHtmlBody {
    param(
        [array]$Alerts,
        [string[]]$Summary,
        [string[]]$Errors,
        [int]$Crit,
        [int]$Warn,
        [int]$Hidden
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html><html><head><meta charset="utf-8"/>')
    [void]$sb.AppendLine('<style>')
    [void]$sb.AppendLine('body { font-family: Segoe UI, Arial, sans-serif; font-size: 12px; color: #222; }')
    [void]$sb.AppendLine('h2 { margin: 0 0 6px 0; font-size: 15px; }')
    [void]$sb.AppendLine('h3 { margin: 14px 0 4px 0; font-size: 13px; color: #444; }')
    [void]$sb.AppendLine('.kpi { margin: 4px 0 10px 0; }')
    [void]$sb.AppendLine('.kpi span { display: inline-block; padding: 2px 8px; margin-right: 6px; border-radius: 3px; font-weight: bold; color: #fff; }')
    [void]$sb.AppendLine('.k-crit { background: #c0392b; } .k-warn { background: #d68910; } .k-err { background: #6c3483; } .k-ok { background: #1e8449; } .k-hide { background: #7f8c8d; }')
    [void]$sb.AppendLine('table.at { border-collapse: collapse; width: 100%; margin-top: 4px; }')
    [void]$sb.AppendLine('table.at th, table.at td { border: 1px solid #ccc; padding: 4px 6px; vertical-align: top; text-align: left; font-size: 11.5px; }')
    [void]$sb.AppendLine('table.at th { background: #34495e; color: #fff; }')
    [void]$sb.AppendLine('tr.sev-crit td { background: #fdecea; }')
    [void]$sb.AppendLine('tr.sev-warn td { background: #fef5e7; }')
    [void]$sb.AppendLine('td.sev { font-weight: bold; }')
    [void]$sb.AppendLine('td.sev-crit { color: #c0392b; } td.sev-warn { color: #b9770e; }')
    [void]$sb.AppendLine('pre.sum { background: #f4f6f7; border: 1px solid #d5dbdb; padding: 6px; white-space: pre-wrap; font-size: 11px; }')
    [void]$sb.AppendLine('ul.err { color: #6c3483; }')
    [void]$sb.AppendLine('</style></head><body>')

    [void]$sb.AppendLine("<h2>OneView Alerts - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</h2>")
    [void]$sb.AppendLine('<div class="kpi">')
    [void]$sb.AppendLine("<span class='k-crit'>Critical: $Crit</span>")
    [void]$sb.AppendLine("<span class='k-warn'>Warning: $Warn</span>")
    if ($Errors.Count -gt 0) { [void]$sb.AppendLine("<span class='k-err'>Errors: $($Errors.Count)</span>") }
    if ($Hidden -gt 0) { [void]$sb.AppendLine("<span class='k-hide'>Ausgeblendet: $Hidden</span>") }
    if ($Crit -eq 0 -and $Warn -eq 0 -and $Errors.Count -eq 0) { [void]$sb.AppendLine("<span class='k-ok'>Keine Alarme</span>") }
    [void]$sb.AppendLine('</div>')

    # Alerts werden bereits nach Severity (Critical > Warning) und Datum sortiert geliefert.
    if ($Alerts.Count -gt 0) {
        [void]$sb.AppendLine('<h3>Alert-Details</h3>')
        [void]$sb.AppendLine('<table class="at">')
        [void]$sb.AppendLine('<tr><th>Severity</th><th>State</th><th>Zeit</th><th>Appliance</th><th>Owner</th><th>Resource</th><th>Type</th><th>Message</th></tr>')
        foreach ($a in $Alerts) {
            $sev = [string]$a.Severity
            $cls = ''
            $tdCls = ''
            if ($sev -ieq 'Critical') { $cls = ' class="sev-crit"'; $tdCls = ' class="sev sev-crit"' }
            elseif ($sev -ieq 'Warning') { $cls = ' class="sev-warn"'; $tdCls = ' class="sev sev-warn"' }
            $ts = ''
            try { if ($a.Created) { $ts = ([datetime]$a.Created).ToString('yyyy-MM-dd HH:mm:ss') } } catch {}
            $ownerVal = if ([string]::IsNullOrWhiteSpace([string]$a.Owner)) { 'unknown' } else { [string]$a.Owner }
            $row = "<tr$cls>" +
            "<td$tdCls>$(ConvertTo-HtmlEncoded $sev)</td>" +
            "<td>$(ConvertTo-HtmlEncoded ([string]$a.State))</td>" +
            "<td>$(ConvertTo-HtmlEncoded $ts)</td>" +
            "<td>$(ConvertTo-HtmlEncoded ([string]$a.Appliance))</td>" +
            "<td>$(ConvertTo-HtmlEncoded $ownerVal)</td>" +
            "<td>$(ConvertTo-HtmlEncoded ([string]$a.Resource))</td>" +
            "<td>$(ConvertTo-HtmlEncoded ([string]$a.Type))</td>" +
            "<td>$(ConvertTo-HtmlEncoded ([string]$a.Message))</td>" +
            '</tr>'
            [void]$sb.AppendLine($row)
        }
        [void]$sb.AppendLine('</table>')
    }

    if ($Errors.Count -gt 0) {
        [void]$sb.AppendLine('<h3>Fehlerdetails</h3><ul class="err">')
        foreach ($e in $Errors) { [void]$sb.AppendLine("<li>$(ConvertTo-HtmlEncoded $e)</li>") }
        [void]$sb.AppendLine('</ul>')
    }

    if ($Summary -and $Summary.Count -gt 0) {
        [void]$sb.AppendLine('<h3>Zusammenfassung je Appliance</h3>')
        [void]$sb.AppendLine("<pre class='sum'>$(ConvertTo-HtmlEncoded ($Summary -join [Environment]::NewLine))</pre>")
    }

    [void]$sb.AppendLine('</body></html>')
    return $sb.ToString()
}

function Send-AlertsEmail {
    param(
        [array]$Alerts,
        [string[]]$Summary,
        [string[]]$Errors,
        [int]$Crit,
        [int]$Warn,
        [int]$Hidden
    )

    if (-not $config.SendEmail) { Write-RunLog "E-Mail-Versand deaktiviert."; return }
    if (-not $config.SmtpServer -or -not $config.MailTo -or -not $config.MailFrom) {
        Write-RunLog "SMTP-Parameter unvollständig - keine E-Mail gesendet." -Level WARN
        return
    }
    # Only-on-error Modus
    if ($config.OnlyOnErrors -and $Crit -eq 0 -and $Warn -eq 0 -and $Errors.Count -eq 0) {
        Write-RunLog "OnlyOnErrors aktiv und keine Alerts/Fehler - keine E-Mail."
        return
    }

    $subjectPrefix = if ($config.SubjectPrefix) { $config.SubjectPrefix } else { '[OneView Alerts]' }
    $hostName = try { $env:COMPUTERNAME } catch { 'unknown' }
    $subject = "$subjectPrefix Critical=$Crit Warning=$Warn Errors=$($Errors.Count) ($hostName)"
    $htmlBody = Build-AlertsHtmlBody -Alerts $Alerts -Summary $Summary -Errors $Errors -Crit $Crit -Warn $Warn -Hidden $Hidden

    # Zertifikatsvalidierung weich setzen (interne CA / Self-Signed auf SMTP akzeptieren)
    try {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { param($s, $c, $ch, $e) return $true }
    }
    catch { }

    try {
        $smtpPort = if ($config.SmtpPort) { [int]$config.SmtpPort } else { 25 }
        $smtpClient = New-Object Net.Mail.SmtpClient($config.SmtpServer, $smtpPort)
        $smtpClient.EnableSsl = $true

        if ($config.SmtpUser -and $config.SmtpPasswordEncrypted) {
            try {
                $sp = ConvertTo-SecureString -String $config.SmtpPasswordEncrypted -Key $aesKey
                $smtpClient.Credentials = (New-Object System.Management.Automation.PSCredential($config.SmtpUser, $sp)).GetNetworkCredential()
            }
            catch {
                Write-RunLog "SMTP-Passwort konnte nicht entschlüsselt werden." -Level WARN
            }
        }

        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = New-Object System.Net.Mail.MailAddress($config.MailFrom)
        foreach ($rcpt in @($config.MailTo -split '\s*;\s*|\s*,\s*' | Where-Object { $_ })) {
            $mailMessage.To.Add($rcpt)
        }
        $mailMessage.Subject = $subject
        $mailMessage.Body = $htmlBody
        $mailMessage.IsBodyHtml = $true
        $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
        $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8
        if ($Crit -gt 0) { $mailMessage.Priority = [System.Net.Mail.MailPriority]::High }

        $smtpClient.Send($mailMessage)
        $mailMessage.Dispose()
        $smtpClient.Dispose()

        Write-RunLog "E-Mail versendet an: $(($config.MailTo -split '\s*;\s*|\s*,\s*' | Where-Object { $_ }) -join ', ')"
    }
    catch {
        Write-RunLog "E-Mail-Versand fehlgeschlagen: $($_.Exception.Message)" -Level ERROR
    }
}

Send-AlertsEmail -Alerts $ordered -Summary $summaryLines -Errors $errorDetails -Crit $critCount -Warn $warnCount -Hidden 0

Write-RunLog "Lauf beendet."
