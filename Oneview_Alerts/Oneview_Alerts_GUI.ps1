#Requires -Version 7.0
## Laufzeit-Prüfungen und Konsolenfenster ausblenden (nur Windows)
if (-not $IsWindows) { Write-Error "Dieses Tool unterstützt nur Windows (Server 2022)."; return }

try {
    # PS 7.5+ sicherstellen
    $minPs = [Version]"7.5.0"
    if ($PSVersionTable.PSVersion -lt $minPs) { Write-Error "PowerShell 7.5 oder neuer erforderlich."; return }
}
catch {}

# Konsolenfenster ausblenden (über Win32 API)
try {
    $hwnd = (Get-Process -Id $PID).MainWindowHandle
    $signature = @"
using System;
using System.Runtime.InteropServices;
public class NativeMethods {
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
    Add-Type -TypeDefinition $signature -ErrorAction Stop
    [NativeMethods]::ShowWindow($hwnd, 0) | Out-Null
}
catch {}

# Laden der benötigten Assemblies und Visual Styles aktivieren
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Skriptverzeichnis und Dateipfade definieren
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $MyInvocation.MyCommand.Path -Parent }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
$currentFile = Join-Path -Path $scriptDir -ChildPath "Alerts_current.txt"
$previousFile = Join-Path -Path $scriptDir -ChildPath "Alerts_previous.txt"
$errorFile = Join-Path -Path $scriptDir -ChildPath "Alerts_errors.txt"
$knownIssuesFile = Join-Path -Path $scriptDir -ChildPath "KnownIssues.txt"

# Globale Log-Puffer – hier werden alle Meldungen gesammelt
$global:logOutput = @()
$global:errorLog = @()
$script:isBusy = $false
$script:knownIssues = @()
# Zwischenspeicher für Anmeldeinformationen (SecureString-basiert, bleibt im Speicher
# verschlüsselt). Wird benötigt, damit der Auto-Refresh-Timer ohne erneute Eingabe
# arbeiten kann. Klartext-Passwort wird trotzdem nach der ersten Eingabe aus dem
# UI-Puffer entfernt.
$script:credential = $null

function Get-KnownIssues {
    try {
        if (Test-Path $knownIssuesFile) {
            $lines = Get-Content -Path $knownIssuesFile -ErrorAction Stop
            $script:knownIssues = @(
                $lines | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') }
            )
        }
        else {
            $script:knownIssues = @()
        }
    }
    catch {
        $script:knownIssues = @()
    }
}

function Set-KnownIssues {
    try {
        if ($script:knownIssues -and $script:knownIssues.Count -gt 0) {
            ($script:knownIssues -join [Environment]::NewLine) | Set-Content -Path $knownIssuesFile -Encoding UTF8
        }
        else {
            "# Eine Zeile pro Muster (case-insensitive, Teilstring)." | Set-Content -Path $knownIssuesFile -Encoding UTF8
        }
    }
    catch {}
}

# Prüfen, ob ein Alert zu den bekannten Issues gehört
function Test-HideAlert {
    param(
        [Parameter(Mandatory)]$Alert,
        [string]$Appliance
    )
    if (-not $script:knownIssues -or $script:knownIssues.Count -eq 0) { return $false }

    # Kandidatenfelder einsammeln (fehlertolerant)
    $vals = @()
    foreach ($p in 'description', 'message', 'alertTypeID', 'resourceName', 'name') {
        try {
            if ($Alert.PSObject.Properties.Match($p).Count -gt 0) {
                $v = $Alert.$p
                if ($v) { $vals += [string]$v }
            }
        }
        catch {}
    }
    # associatedResource.resourceName
    try {
        if ($Alert.PSObject.Properties.Match('associatedResource').Count -gt 0 -and $Alert.associatedResource) {
            if ($Alert.associatedResource.PSObject.Properties.Match('resourceName').Count -gt 0 -and $Alert.associatedResource.resourceName) {
                $vals += [string]$Alert.associatedResource.resourceName
            }
        }
    }
    catch {}
    # Kanonische Zeile wie in Detailausgabe, damit Copy/Paste funktioniert
    try {
        $sev = $Alert.Severity
        $st = $Alert.AlertState
        $ts = ''
        try { if ($Alert.PSObject.Properties.Match('created').Count -gt 0 -and $Alert.created) { $ts = ([datetime]$Alert.created).ToString('yyyy-MM-dd HH:mm:ss') } } catch {}
        $resName = $null
        try {
            if ($Alert.PSObject.Properties.Match('resourceName').Count -gt 0 -and $Alert.resourceName) { $resName = $Alert.resourceName }
            elseif ($Alert.PSObject.Properties.Match('associatedResource').Count -gt 0 -and $Alert.associatedResource -and $Alert.associatedResource.PSObject.Properties.Match('resourceName').Count -gt 0) { $resName = $Alert.associatedResource.resourceName }
        }
        catch {}
        $atype = $Alert.alertTypeID
        $msgf = $null
        try { $msgf = $Alert.message } catch {}
        if (-not $msgf) { try { $msgf = $Alert.description } catch {} }
        $canon = "[{0}][{1}] {2} {3} {4} ({5}): {6}" -f $sev, $st, $ts, $Appliance, $resName, $atype, $msgf
        if ($canon.Trim()) { $vals += $canon }
    }
    catch {}
    # Komplettes Objekt zusätzlich (immer anhängen)
    try { $vals += ($Alert | Out-String) } catch {}

    $aggregate = ($vals -join ' ') -as [string]
    if (-not $aggregate) { return $false }
    # Normalisieren: Mehrfache Whitespaces reduzieren, trimmen
    $norm = ($aggregate -replace '\s+', ' ').Trim()
    $normLower = $norm.ToLowerInvariant()

    foreach ($pattern in $script:knownIssues) {
        $p = $pattern.Trim()
        if (-not $p) { continue }
        $matched = $false
        if ($p.StartsWith('~')) {
            # explizite Regex (Case-insensitive)
            $rx = $p.Substring(1)
            try { if ($norm -imatch $rx) { $matched = $true } } catch {}
        }
        elseif ($p.StartsWith('^') -or $p.EndsWith('$')) {
            # Regex-Anker vermuten (direkt verwenden)
            try { if ($norm -imatch $p) { $matched = $true } } catch {}
        }
        else {
            $pl = ($p -replace '\s+', ' ').Trim().ToLowerInvariant()
            if ($normLower.Contains($pl)) { $matched = $true }
        }
        if ($matched) {
            if (-not $script:hidePatternStats) { $script:hidePatternStats = @{} }
            if ($script:hidePatternStats.ContainsKey($p)) { $script:hidePatternStats[$p]++ } else { $script:hidePatternStats[$p] = 1 }
            return $true
        }
    }
    return $false
}
# --------------------- GUI-Erstellung -----------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "© 2025 N.J. Airbus D&S - OneView Alerts"
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.ClientSize = New-Object System.Drawing.Size(940, 700)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(956, 739)

$grpCred = New-Object System.Windows.Forms.GroupBox
$grpCred.Text = "Anmeldeinformationen"
$grpCred.Location = New-Object System.Drawing.Point(10, 10)
$grpCred.Size = New-Object System.Drawing.Size(380, 110)
$grpCred.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($grpCred)

$lblUsername = New-Object System.Windows.Forms.Label
$lblUsername.Text = "Benutzername:"
$lblUsername.Location = New-Object System.Drawing.Point(10, 20)
$lblUsername.Size = New-Object System.Drawing.Size(100, 20)
$grpCred.Controls.Add($lblUsername)

$txtUsername = New-Object System.Windows.Forms.TextBox
$txtUsername.Location = New-Object System.Drawing.Point(120, 20)
$txtUsername.Size = New-Object System.Drawing.Size(250, 20)
$grpCred.Controls.Add($txtUsername)

$lblPassword = New-Object System.Windows.Forms.Label
$lblPassword.Text = "Kennwort:"
$lblPassword.Location = New-Object System.Drawing.Point(10, 50)
$lblPassword.Size = New-Object System.Drawing.Size(100, 20)
$grpCred.Controls.Add($lblPassword)

$txtPassword = New-Object System.Windows.Forms.TextBox
$txtPassword.Location = New-Object System.Drawing.Point(120, 50)
$txtPassword.Size = New-Object System.Drawing.Size(250, 20)
$txtPassword.UseSystemPasswordChar = $true
$grpCred.Controls.Add($txtPassword)

${grpFile} = New-Object System.Windows.Forms.GroupBox
${grpFile}.Text = "Appliance Datei auswählen"
${grpFile}.Location = New-Object System.Drawing.Point(460, 10)
${grpFile}.Size = New-Object System.Drawing.Size(470, 140)
${grpFile}.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add(${grpFile})

$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text = "Datei/Modus:"
$lblFile.Location = New-Object System.Drawing.Point(10, 20)
$lblFile.Size = New-Object System.Drawing.Size(100, 20)
${grpFile}.Controls.Add($lblFile)

$cmbFile = New-Object System.Windows.Forms.ComboBox
$cmbFile.Location = New-Object System.Drawing.Point(110, 20)
$cmbFile.Size = New-Object System.Drawing.Size(310, 20)
$cmbFile.DropDownStyle = 'DropDownList'
$cmbFile.Items.AddRange(@("Oneview_GOV.txt", "Oneview_DIV.txt", "Beide Dateien"))
$cmbFile.SelectedIndex = 0
${grpFile}.Controls.Add($cmbFile)

# Einstellungen (Zeitraum, Auto-Refresh, Filter)
$lblRange = New-Object System.Windows.Forms.Label
$lblRange.Text = "Zeitraum:"
$lblRange.Location = New-Object System.Drawing.Point(10, 50)
$lblRange.Size = New-Object System.Drawing.Size(100, 20)
${grpFile}.Controls.Add($lblRange)

$numRange = New-Object System.Windows.Forms.NumericUpDown
$numRange.Location = New-Object System.Drawing.Point(110, 50)
$numRange.Size = New-Object System.Drawing.Size(70, 20)
$numRange.Minimum = 1
$numRange.Maximum = 365
$numRange.Value = 30
${grpFile}.Controls.Add($numRange)

$cmbRangeUnit = New-Object System.Windows.Forms.ComboBox
$cmbRangeUnit.Location = New-Object System.Drawing.Point(190, 50)
$cmbRangeUnit.Size = New-Object System.Drawing.Size(110, 20)
$cmbRangeUnit.DropDownStyle = 'DropDownList'
$cmbRangeUnit.Items.AddRange(@('Tage', 'Stunden'))
$cmbRangeUnit.SelectedIndex = 0
${grpFile}.Controls.Add($cmbRangeUnit)

$lblInterval = New-Object System.Windows.Forms.Label
$lblInterval.Text = "Auto-Refresh (Min):"
$lblInterval.Location = New-Object System.Drawing.Point(305, 50)
$lblInterval.Size = New-Object System.Drawing.Size(100, 20)
${grpFile}.Controls.Add($lblInterval)

$numInterval = New-Object System.Windows.Forms.NumericUpDown
$numInterval.Location = New-Object System.Drawing.Point(410, 50)
$numInterval.Size = New-Object System.Drawing.Size(40, 20)
$numInterval.Minimum = 1
$numInterval.Maximum = 1440
$numInterval.Value = 30
${grpFile}.Controls.Add($numInterval)

$chkHideKnown = New-Object System.Windows.Forms.CheckBox
$chkHideKnown.Text = "Bekannte Issues ausblenden"
$chkHideKnown.Location = New-Object System.Drawing.Point(10, 80)
$chkHideKnown.AutoSize = $true
$chkHideKnown.Checked = $true
${grpFile}.Controls.Add($chkHideKnown)

$chkOwnerUnknown = New-Object System.Windows.Forms.CheckBox
$chkOwnerUnknown.Text = "Nur Owner = unknown"
$chkOwnerUnknown.Location = New-Object System.Drawing.Point(240, 110)
$chkOwnerUnknown.AutoSize = $true
$chkOwnerUnknown.Checked = $true
${grpFile}.Controls.Add($chkOwnerUnknown)

$btnManageKnown = New-Object System.Windows.Forms.Button
$btnManageKnown.Text = "Bekannte Issues verwalten"
$btnManageKnown.Location = New-Object System.Drawing.Point(10, 105)
$btnManageKnown.Size = New-Object System.Drawing.Size(210, 28)
${grpFile}.Controls.Add($btnManageKnown)


# Max. Details (Begrenzung für Detailausgabe)
$lblMaxDetails = New-Object System.Windows.Forms.Label
$lblMaxDetails.Text = "Max. Details:"
$lblMaxDetails.Location = New-Object System.Drawing.Point(240, 80)
$lblMaxDetails.Size = New-Object System.Drawing.Size(100, 20)
${grpFile}.Controls.Add($lblMaxDetails)

$numMaxDetails = New-Object System.Windows.Forms.NumericUpDown
$numMaxDetails.Location = New-Object System.Drawing.Point(350, 80)
$numMaxDetails.Size = New-Object System.Drawing.Size(70, 20)
$numMaxDetails.Minimum = 1
$numMaxDetails.Maximum = 5000
$numMaxDetails.Value = 100
${grpFile}.Controls.Add($numMaxDetails)

# (Ampel wird jetzt unten rechts platziert – Definition später nach Panel-Erstellung)

$pnlRichText = New-Object System.Windows.Forms.Panel
$pnlRichText.Location = New-Object System.Drawing.Point(10, 160)
$pnlRichText.Size = New-Object System.Drawing.Size(920, 370)
$pnlRichText.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor `
    [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$pnlRichText.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($pnlRichText)

$rtbOutput = New-Object System.Windows.Forms.RichTextBox
$rtbOutput.Dock = [System.Windows.Forms.DockStyle]::Fill
$rtbOutput.Font = New-Object System.Drawing.Font("Consolas", 10)
$rtbOutput.ReadOnly = $true
$rtbOutput.WordWrap = $false
$pnlRichText.Controls.Add($rtbOutput)

$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(10, 540)
$panel.Size = New-Object System.Drawing.Size(920, 90)
$panel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($panel)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Alerts abfragen"
$btnStart.Location = New-Object System.Drawing.Point(10, 10)
$btnStart.Size = New-Object System.Drawing.Size(120, 30)
$panel.Controls.Add($btnStart)

# Button: Aktuelle Datei anzeigen
$btnViewCurrent = New-Object System.Windows.Forms.Button
$btnViewCurrent.Text = "Log-Datei"
$btnViewCurrent.Location = New-Object System.Drawing.Point(140, 10)
$btnViewCurrent.Size = New-Object System.Drawing.Size(120, 30)
$panel.Controls.Add($btnViewCurrent)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object System.Drawing.Point(270, 10)
$btnClear.Size = New-Object System.Drawing.Size(120, 30)
$panel.Controls.Add($btnClear)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = "Exit"
$btnExit.Location = New-Object System.Drawing.Point(400, 10)
$btnExit.Size = New-Object System.Drawing.Size(120, 30)
$panel.Controls.Add($btnExit)

# Ampel Status jetzt horizontal mit drei Feldern (blass / aktiv)
$grpStatus = New-Object System.Windows.Forms.GroupBox
$grpStatus.Text = "Status"
$grpStatus.Size = New-Object System.Drawing.Size(330, 90)
$grpStatus.Location = New-Object System.Drawing.Point(($panel.Width - $grpStatus.Width - 10), 0)
$grpStatus.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$panel.Controls.Add($grpStatus)

$trafficContainer = New-Object System.Windows.Forms.Panel
$trafficContainer.Size = New-Object System.Drawing.Size(110, 30)
$trafficContainer.Location = New-Object System.Drawing.Point(8, 20)
$trafficContainer.BorderStyle = 'None'
$grpStatus.Controls.Add($trafficContainer)

function New-LightPanel([int]$x) {
    $p = New-Object System.Windows.Forms.Panel
    $p.Size = New-Object System.Drawing.Size(24, 24)
    $p.Location = New-Object System.Drawing.Point($x, 3)
    $p.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $p.BackColor = [System.Drawing.Color]::Gainsboro
    return $p
}

$pnlGreen = New-LightPanel -x 0
$pnlYellow = New-LightPanel -x 36
$pnlRed = New-LightPanel -x 72
$trafficContainer.Controls.AddRange(@($pnlGreen, $pnlYellow, $pnlRed))

$lblStatusText = New-Object System.Windows.Forms.Label
$lblStatusText.Location = New-Object System.Drawing.Point(8, 55)
$lblStatusText.Size = New-Object System.Drawing.Size(314, 28)
$lblStatusText.AutoEllipsis = $true
$lblStatusText.Text = "Noch keine Abfrage"
$grpStatus.Controls.Add($lblStatusText)

# Fortschrittsanzeige (etwas tiefer gesetzt für Platz der horizontalen Ampel)
$prgAppliances = New-Object System.Windows.Forms.ProgressBar
$prgAppliances.Location = New-Object System.Drawing.Point(10, 75)
$prgAppliances.Size = New-Object System.Drawing.Size(520, 10)
$prgAppliances.Style = 'Continuous'
$prgAppliances.Minimum = 0
$prgAppliances.Maximum = 100
$prgAppliances.Value = 0
$panel.Controls.Add($prgAppliances)

function Set-Lights([string[]]$activeColors) {
    # Noch blassere Grundfarben (geringere Alpha + etwas entsättigt)
    $pnlGreen.BackColor = [System.Drawing.Color]::FromArgb(40, 60, 180, 60)
    $pnlYellow.BackColor = [System.Drawing.Color]::FromArgb(40, 210, 190, 40)
    $pnlRed.BackColor = [System.Drawing.Color]::FromArgb(40, 210, 40, 40)
    if ($activeColors -contains 'Green') { $pnlGreen.BackColor = [System.Drawing.Color]::LimeGreen }
    if ($activeColors -contains 'Yellow') { $pnlYellow.BackColor = [System.Drawing.Color]::Gold }
    if ($activeColors -contains 'Red') { $pnlRed.BackColor = [System.Drawing.Color]::Red }
}

function Update-TrafficLight {
    param(
        [int]$CriticalCount,
        [int]$WarningCount
    )
    # Keine Probleme -> nur Grün kräftig
    # Nur Warnings -> Gelb kräftig
    # Nur Critical -> Rot kräftig
    # Critical und Warning -> Rot UND Gelb kräftig
    $active = @()
    if ($CriticalCount -gt 0 -and $WarningCount -gt 0) { $active = @('Red', 'Yellow') }
    elseif ($CriticalCount -gt 0) { $active = @('Red') }
    elseif ($WarningCount -gt 0) { $active = @('Yellow') }
    else { $active = @('Green') }
    Set-Lights -activeColors $active
    $script:trafficState = [PSCustomObject]@{
        Critical = $CriticalCount; Warning = $WarningCount; ActiveColors = $active; Blink = $false; Toggle = $false; BlinkPanels = @()
    }
    # Blink nur für Rot; wenn kein Rot aber Gelb aktiv dann Gelb blinken lassen.
    if ($active -contains 'Red') { $script:trafficState.Blink = $true; $script:trafficState.BlinkPanels += $pnlRed }
    elseif ($active -contains 'Yellow') { $script:trafficState.Blink = $true; $script:trafficState.BlinkPanels += $pnlYellow }
    if ($active -contains 'Red') { $lblStatusText.Text = "Critical: $CriticalCount" }
    elseif ($active -contains 'Yellow') { $lblStatusText.Text = "Warning: $WarningCount" }
    else { $lblStatusText.Text = "OK (keine Alerts)" }
}

# Blink-Timer
$trafficTimer = New-Object System.Windows.Forms.Timer
$trafficTimer.Interval = 800
$trafficTimer.Add_Tick({
        if (-not $script:trafficState) { return }
        if (-not $script:trafficState.Blink) { return }
        if (-not $script:trafficState.BlinkPanels -or $script:trafficState.BlinkPanels.Count -eq 0) { return }
        $script:trafficState.Toggle = -not $script:trafficState.Toggle
        foreach ($p in $script:trafficState.BlinkPanels) {
            if ($script:trafficState.Toggle) {
                $c = $p.BackColor
                $p.BackColor = [System.Drawing.Color]::FromArgb(255, [Math]::Min($c.R + 60, 255), [Math]::Min($c.G + 60, 255), [Math]::Min($c.B + 60, 255))
            }
            else {
                if ($p -eq $pnlRed) { $p.BackColor = [System.Drawing.Color]::Red }
                elseif ($p -eq $pnlYellow) { $p.BackColor = [System.Drawing.Color]::Gold }
                elseif ($p -eq $pnlGreen) { $p.BackColor = [System.Drawing.Color]::LimeGreen }
            }
        }
    })
$trafficTimer.Start()

# Re-Positionieren bei Größenänderung (rechts ausrichten)
$form.Add_SizeChanged({
        $grpStatus.Location = New-Object System.Drawing.Point(($panel.Width - $grpStatus.Width - 10), 5)
    })

# Modul- und Abhängigkeitsprüfung
function Test-Dependencies {
    # Wir laden die Module nicht im Hauptprozess (Konflikt 660 vs. 1000),
    # sondern prüfen nur, dass mindestens eines verfügbar ist.
    $have660 = [bool](Get-Module -ListAvailable -Name 'HPEOneView.660')
    $have1000 = [bool](Get-Module -ListAvailable -Name 'HPEOneView.1000')
    if (-not ($have660 -or $have1000)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Kein HPE OneView PowerShell Modul gefunden (HPEOneView.660 für OV 6.60 / HPEOneView.1000 für OV 11.10+).",
            "Modul fehlt", 0, 16) | Out-Null
        return $false
    }
    $script:haveModule660 = $have660
    $script:haveModule1000 = $have1000
    return $true
}

# Zertifikatsprüfung für HPE OneView Appliances deaktivieren
# (selbstsignierte / abgelaufene Zertifikate werden akzeptiert)
function Disable-CertificateValidation {
    try {
        # TLS 1.2/1.3 zulassen
        [System.Net.ServicePointManager]::SecurityProtocol = `
            [System.Net.SecurityProtocolType]::Tls12 -bor `
            [System.Net.SecurityProtocolType]::Tls13
    }
    catch {
        try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {}
    }
    try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch {}
    $Global:SetLibraryBypassCertificatePolicy = $true
}

# OneView-Version einer Appliance per REST /rest/version (unauthentifiziert) ermitteln.
# Rückgabe: Objekt mit Appliance, Module, SoftwareVersion.
# Mapping (gemäß HPE API Reference):
#   OV 6.60      -> API 3800  -> HPEOneView.660  (Legacy)
#   OV 7.x-11.x  -> API > 3800 -> HPEOneView.1000 (aktuell/neu)
# Fallback bei Fehler: HPEOneView.660 (sicherer gegen Header-Fehler).
function Get-ApplianceOVModule {
    param([Parameter(Mandatory)][string]$Appliance)

    $module = 'HPEOneView.660'
    $softwareVersion = $null
    try {
        $uri = "https://$Appliance/rest/version"
        # /rest/version ist in neueren OV-Versionen auch ohne X-API-Version aufrufbar,
        # einige Builds erwarten dennoch den Header. Wir senden einen konservativen Wert.
        $headers = @{ 'X-API-Version' = '3800' }
        $resp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -TimeoutSec 10 -SkipCertificateCheck -ErrorAction Stop
        if ($resp.softwareVersion) {
            $softwareVersion = [string]$resp.softwareVersion
            # Primär: Software-Hauptversion auswerten (z.B. "6.60.05", "11.10.00")
            if ($softwareVersion -match '^(\d+)') {
                $major = [int]$Matches[1]
                if ($major -ge 7) { $module = 'HPEOneView.1000' } else { $module = 'HPEOneView.660' }
            }
        }
        elseif ($resp.currentVersion) {
            # Fallback: API-Level. OV 6.60 = 3800 -> 660; alles darüber -> 1000.
            $apiVer = 0
            [int]::TryParse([string]$resp.currentVersion, [ref]$apiVer) | Out-Null
            if ($apiVer -gt 3800) { $module = 'HPEOneView.1000' } else { $module = 'HPEOneView.660' }
        }
    }
    catch {}
    return [PSCustomObject]@{
        Appliance       = $Appliance
        Module          = $module
        SoftwareVersion = $softwareVersion
    }
}

Get-KnownIssues

# Hilfsfunktion: Zeitstempel aus Alert robust extrahieren
function Get-AlertTimestamp {
    param([object]$Alert)
    foreach ($p in 'created', 'date', 'EventTimestamp') {
        if ($Alert.PSObject.Properties.Match($p).Count -gt 0 -and $Alert.$p) {
            try { return [datetime]$Alert.$p } catch {}
        }
    }
    return $null
}

# --------------------- Funktionaler Code für Alert-Abfrage ---------------------
function Invoke-Alerts {
    if ($script:isBusy) { return }
    $script:isBusy = $true
    try {
        # Known Issues bei jedem Lauf neu laden, falls Datei extern geändert wurde
        Get-KnownIssues
        $btnStart.Enabled = $false
        $btnExit.Enabled = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        if (-not (Test-Dependencies)) { return }

        # Zertifikatsprüfung deaktivieren (selbstsignierte OneView-Zertifikate zulassen)
        Disable-CertificateValidation

        $global:logOutput = @()
        $global:errorLog = @()
        $rtbOutput.Clear()
    
        $global:logOutput += "#####################################################"
        $global:logOutput += "Alerts Log - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $global:logOutput += "#####################################################"
        $global:logOutput += ""
    
        # Anmeldeinformationen ermitteln:
        #  - Wenn der Benutzer Username + Passwort eingegeben hat, wird ein neuer
        #    PSCredential erzeugt und für spätere Auto-Refresh-Läufe zwischengespeichert.
        #  - Andernfalls wird der zwischengespeicherte Credential (SecureString) verwendet.
        #  - Das Klartext-Passwort wird in jedem Fall sofort aus dem UI-Puffer entfernt,
        #    sodass es nicht dauerhaft als Klartext im Speicher liegt.
        if (-not [string]::IsNullOrWhiteSpace($txtPassword.Text)) {
            if ([string]::IsNullOrWhiteSpace($txtUsername.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Bitte Benutzername und Kennwort eingeben.", "Fehlende Anmeldeinformationen",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $secPassword = ConvertTo-SecureString $txtPassword.Text -AsPlainText -Force
            $script:credential = New-Object System.Management.Automation.PSCredential($txtUsername.Text, $secPassword)
            # Klartext-Passwort sofort aus dem UI-Puffer entfernen
            $txtPassword.Text = ''
            Remove-Variable secPassword -ErrorAction SilentlyContinue
        }
        elseif ($null -eq $script:credential) {
            [System.Windows.Forms.MessageBox]::Show("Bitte Benutzername und Kennwort eingeben.", "Fehlende Anmeldeinformationen",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $credential = $script:credential
    
        $applianceFiles = @()
        switch ($cmbFile.SelectedItem) {
            "Oneview_GOV.txt" { $applianceFiles = @(Join-Path $scriptDir "Oneview_GOV.txt") }
            "Oneview_DIV.txt" { $applianceFiles = @(Join-Path $scriptDir "Oneview_DIV.txt") }
            "Beide Dateien" { $applianceFiles = @( (Join-Path $scriptDir "Oneview_GOV.txt"), (Join-Path $scriptDir "Oneview_DIV.txt") ) }
            default {
                [System.Windows.Forms.MessageBox]::Show("Keine gültige Dateiauswahl getroffen.", "Auswahl ungültig",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }
        }
        $appliances = @()
        foreach ($file in $applianceFiles) {
            if (-Not (Test-Path $file)) {
                [System.Windows.Forms.MessageBox]::Show("Die Appliance-Datei wurde nicht gefunden: " + $file, "Datei nicht gefunden",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                continue
            }
            $content = @(Get-Content -Path $file -ErrorAction SilentlyContinue | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') })
            if ($content.Count -eq 0) {
                $global:logOutput += "Datei $file ist leer."
            }
            else {
                $appliances += $content
            }
        }
        if ($appliances.Count -eq 0) {
            $global:logOutput += "Keine Appliance-Einträge gefunden."
            $rtbOutput.Lines = $global:logOutput
            $rtbOutput.SelectionStart = $rtbOutput.Text.Length
            $rtbOutput.ScrollToCaret()
            return
        }
    
        # Zeitraum aus UI
        $timespan = if ($cmbRangeUnit.SelectedItem -eq 'Stunden') {
            New-TimeSpan -Hours ([int]$numRange.Value)
        }
        else { New-TimeSpan -Days ([int]$numRange.Value) }
        $errorDetails = @()
        $script:allAlerts = @()
        $now = Get-Date
        $cutoff = $now - $timespan

        $total = $appliances.Count
        $index = 0
        if ($total -gt 0) { $prgAppliances.Value = 0 }
        # ProgressBar Reset beim Start
        try { $prgAppliances.Value = 0 } catch {}

        # ----------------------------------------------------------
        # 1) OneView-Version pro Appliance ermitteln und gruppieren
        # ----------------------------------------------------------
        try { $lblStatusText.Text = "Ermittle OneView-Versionen ..." } catch {}
        [System.Windows.Forms.Application]::DoEvents()
        $applianceInfos = foreach ($a in $appliances) {
            $info = Get-ApplianceOVModule -Appliance $a
            $verTxt = if ($info.SoftwareVersion) { $info.SoftwareVersion } else { '(unbekannt, Fallback)' }
            $global:logOutput += "Version $($a): $verTxt -> $($info.Module)"
            $info
        }
        $byModule = $applianceInfos | Group-Object -Property Module

        # Prüfen, ob für jede benötigte Version das Modul installiert ist
        foreach ($grp in $byModule) {
            $mod = $grp.Name
            $installed = if ($mod -eq 'HPEOneView.660') { $script:haveModule660 } elseif ($mod -eq 'HPEOneView.1000') { $script:haveModule1000 } else { $false }
            if (-not $installed) {
                $names = ($grp.Group.Appliance -join ', ')
                $msg = "Modul $mod fehlt - Appliances werden übersprungen: $names"
                $global:logOutput += $msg
                $global:errorLog += $msg
                $errorDetails += $msg
            }
        }

        # ----------------------------------------------------------
        # 2) Batch-ScriptBlock: läuft als Start-Job in eigenem Prozess
        #    (eigene Prozesse = Modul-Isolation 660 vs. 1000)
        # ----------------------------------------------------------
        $batchScript = {
            param(
                [string]$ModuleName,
                [string[]]$ApplianceList,
                [System.Management.Automation.PSCredential]$Credential,
                [timespan]$Timespan,
                [datetime]$Cutoff,
                [bool]$OwnerUnknownOnly
            )
            # Zertifikats-Bypass im Job-Prozess
            try {
                [System.Net.ServicePointManager]::SecurityProtocol = `
                    [System.Net.SecurityProtocolType]::Tls12 -bor `
                    [System.Net.SecurityProtocolType]::Tls13
            } catch {
                try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {}
            }
            try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch {}
            $Global:SetLibraryBypassCertificatePolicy = $true

            try {
                Import-Module $ModuleName -Force -ErrorAction Stop
            }
            catch {
                foreach ($a in $ApplianceList) {
                    [PSCustomObject]@{ Type='RESULT'; Appliance=$a; Module=$ModuleName; Alerts=@(); Error="Modul $ModuleName konnte nicht geladen werden: $($_.Exception.Message)" }
                }
                return
            }

            function Test-OwnerUnknown2 {
                param([object]$Alert)
                $val = $null
                foreach ($p in 'assignedToUser','owner','assignedTo') {
                    if ($Alert.PSObject.Properties.Match($p).Count -gt 0 -and $null -ne $Alert.$p) {
                        $val = [string]$Alert.$p
                        break
                    }
                }
                if ([string]::IsNullOrWhiteSpace($val)) { return $true }
                return ($val -match '^(unknown|unassigned|none)$')
            }

            function Get-AlertTimestamp2 {
                param([object]$Alert)
                foreach ($p in 'created','date','EventTimestamp') {
                    if ($Alert.PSObject.Properties.Match($p).Count -gt 0 -and $Alert.$p) {
                        try { return [datetime]$Alert.$p } catch {}
                    }
                }
                return $null
            }

            foreach ($appliance in $ApplianceList) {
                [PSCustomObject]@{ Type='PROGRESS'; Appliance=$appliance; Module=$ModuleName }
                $alerts = @()
                $errorText = $null
                $connected = $false
                try {
                    Connect-OVMgmt -Hostname $appliance -Credential $Credential -ErrorAction Stop | Out-Null
                    $connected = $true
                    try {
                        $ac = @(Get-OVAlert -severity Critical -Timespan $Timespan -AlertState active)
                        $aw = @(Get-OVAlert -severity Warning  -Timespan $Timespan -AlertState active)
                        $lc = @(Get-OVAlert -severity Critical -Timespan $Timespan -AlertState locked)
                        $lw = @(Get-OVAlert -severity Warning  -Timespan $Timespan -AlertState locked)
                        $alerts = @($ac + $aw + $lc + $lw)
                    }
                    catch {
                        # Fallback ohne -Timespan
                        $base = @(Get-OVAlert -severity Critical, Warning -AlertState active, locked)
                        $alerts = @($base | Where-Object {
                            $ts = Get-AlertTimestamp2 -Alert $_
                            $ts -ge $Cutoff -and ($_.Severity -ieq 'Critical' -or $_.Severity -ieq 'Warning') -and ($_.AlertState -ieq 'Active' -or $_.AlertState -ieq 'Locked')
                        })
                    }
                    if ($OwnerUnknownOnly) {
                        $alerts = @($alerts | Where-Object { Test-OwnerUnknown2 -Alert $_ })
                    }
                }
                catch {
                    $errorText = "Verbindung/Abfrage fehlgeschlagen: $($_.Exception.Message)"
                }
                finally {
                    if ($connected) { try { Disconnect-OVMgmt -ErrorAction SilentlyContinue | Out-Null } catch {} }
                }
                [PSCustomObject]@{ Type='RESULT'; Appliance=$appliance; Module=$ModuleName; Alerts=$alerts; Error=$errorText }
            }
        }

        # ----------------------------------------------------------
        # 3) Jobs pro Modul starten
        # ----------------------------------------------------------
        $jobs = @()
        foreach ($grp in $byModule) {
            $mod = $grp.Name
            $installed = if ($mod -eq 'HPEOneView.660') { $script:haveModule660 } elseif ($mod -eq 'HPEOneView.1000') { $script:haveModule1000 } else { $false }
            if (-not $installed) { continue }
            $listForJob = @($grp.Group.Appliance)
            $jobs += Start-Job -Name $mod -ScriptBlock $batchScript -ArgumentList @(
                $mod, $listForJob, $credential, $timespan, $cutoff, [bool]$chkOwnerUnknown.Checked
            )
        }

        # ----------------------------------------------------------
        # 4) Jobs pollen, Fortschritt & Ergebnisse einsammeln
        # ----------------------------------------------------------
        $collected = @{}   # Appliance -> Ergebnis
        while ($jobs.Count -gt 0 -and ($jobs | Where-Object { $_.State -eq 'Running' -or $_.HasMoreData })) {
            foreach ($job in $jobs) {
                $out = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
                foreach ($o in $out) {
                    if ($null -eq $o) { continue }
                    if ($o.Type -eq 'PROGRESS') {
                        $index++
                        try { $lblStatusText.Text = ("Abfrage {0}/{1}: {2} ({3})" -f $index, $total, $o.Appliance, $o.Module) } catch {}
                        if ($total -gt 0) {
                            $pct = [int](($index / $total) * 100)
                            if ($pct -gt 100) { $pct = 100 }
                            try { $prgAppliances.Value = $pct } catch {}
                        }
                    }
                    elseif ($o.Type -eq 'RESULT') {
                        $collected[$o.Appliance] = $o
                    }
                }
            }
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 200
        }
        # Letzte Reste einsammeln
        foreach ($job in $jobs) {
            $out = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
            foreach ($o in $out) {
                if ($o -and $o.Type -eq 'RESULT') { $collected[$o.Appliance] = $o }
            }
        }
        $jobs | ForEach-Object { try { Remove-Job -Job $_ -Force } catch {} }

        # ----------------------------------------------------------
        # 5) Ergebnisse aufbereiten (wie bisher)
        # ----------------------------------------------------------
        foreach ($appliance in $appliances) {
            $global:logOutput += "============================================"
            $global:logOutput += "Appliance: $($appliance)"
            $res = $collected[$appliance]
            if (-not $res) {
                $errMsg = "Fehler: Kein Ergebnis für $appliance (Modul evtl. nicht installiert)."
                $global:logOutput += $errMsg
                $global:errorLog += $errMsg
                $errorDetails += $errMsg
                $global:logOutput += ""
                continue
            }
            if ($res.Error) {
                $errMsg = "Fehler: OneView Appliance $appliance - $($res.Error)"
                $global:logOutput += $errMsg
                $global:errorLog += $errMsg
                $errorDetails += $errMsg
                $global:logOutput += ""
                continue
            }
            $alerts = @($res.Alerts)
            if ($alerts -and $alerts.Count -gt 0) {
                $filtered = $alerts
                $hiddenCount = 0
                if ($chkHideKnown.Checked -and $script:knownIssues.Count -gt 0) {
                    $filtered = foreach ($al in $alerts) {
                        if (Test-HideAlert -Alert $al -Appliance $appliance) { $hiddenCount++; continue }
                        $al
                    }
                }

                $script:allAlerts += ($filtered | ForEach-Object {
                        $res2 = $null
                        if ($_.PSObject.Properties.Match('resourceName').Count -gt 0 -and $_.resourceName) { $res2 = $_.resourceName }
                        elseif ($_.PSObject.Properties.Match('associatedResource').Count -gt 0 -and $_.associatedResource -and $_.associatedResource.PSObject.Properties.Match('resourceName').Count -gt 0) { $res2 = $_.associatedResource.resourceName }
                        [PSCustomObject]@{
                            Appliance = $appliance
                            Severity  = $_.Severity
                            State     = $_.AlertState
                            Created   = $_.created
                            Resource  = $res2
                            Type      = $_.alertTypeID
                            Message   = if ($_.message) { $_.message } else { $_.description }
                        }
                    })

                $grouped = $filtered | Group-Object -Property { "$($_.Severity) - $($_.AlertState)" }
                $global:logOutput += "Alarm-Zusammenfassung für $($appliance):"
                foreach ($g in $grouped) { $global:logOutput += "[$($g.Name)] : $($g.Count)" }
                $global:logOutput += "Details (gefiltert) für diese Appliance: $($filtered.Count)"
                if ($hiddenCount -gt 0) { $global:logOutput += "($hiddenCount bekannte Issue(s) ausgeblendet)" }
            }
            else {
                $global:logOutput += "Keine Alarme (Critical/Warning, Active/Locked) für $($appliance) gefunden."
            }
            $global:logOutput += ""
        }

        $global:logOutput | Set-Content -Path $currentFile -Encoding UTF8
        if ($global:errorLog.Count -gt 0) {
            $global:errorLog | Set-Content -Path $errorFile -Encoding UTF8
        }

        # Vorherige Datei wird aktuell nicht für Delta genutzt

        # Ausgabe in der RichTextBox mit Formatierung:
        # 1. Zusammenfassung mit farblicher Differenzierung nach Severity
        # 2. Immer danach die Detail-Liste (nicht mehr nur wenn keine Summary vorhanden)
        $rtbOutput.Clear()
        $rtbOutput.SelectionColor = [System.Drawing.Color]::Black
        $rtbOutput.AppendText("=== Zusammenfassung Fehler und Warnings ===" + [Environment]::NewLine)

        $summaryLines = $global:logOutput | Where-Object {
            $_ -match '^Alarm-Zusammenfassung' -or $_ -match '^\[' -or $_ -match '^\(' -or $_ -match '^Keine Alarme' -or $_ -match '^Details \(gefiltert\)'
        }

        foreach ($line in $summaryLines) {
            if ($line -match '^\[') {
                # Muster: [Severity - State] : Count
                $severity = $null; $state = $null; $count = 0
                if ($line -match '^\[(?<sev>[^\]-]+)\s*-\s*(?<st>[^\]]+)\]\s*:\s*(?<cnt>\d+)') {
                    $severity = $matches['sev']
                    $state = $matches['st']
                    $count = [int]$matches['cnt']
                }
                switch -Regex ($severity) {
                    '^(?i)Critical$' { $rtbOutput.SelectionColor = if ($count -gt 0) { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Black }; break }
                    '^(?i)Warning$' { $rtbOutput.SelectionColor = if ($count -gt 0) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Black }; break }
                    default { $rtbOutput.SelectionColor = [System.Drawing.Color]::Black }
                }
                $rtbOutput.AppendText($line + [Environment]::NewLine)
                $rtbOutput.SelectionColor = [System.Drawing.Color]::Black
            }
            elseif ($line -match '^\(') {
                # Hinweiszeile (ausgeblendete bekannte Issues) in Grau
                $rtbOutput.SelectionColor = [System.Drawing.Color]::Gray
                $rtbOutput.AppendText($line + [Environment]::NewLine)
                $rtbOutput.SelectionColor = [System.Drawing.Color]::Black
            }
            else {
                # Kopf- & Infozeilen schwarz
                $rtbOutput.SelectionColor = [System.Drawing.Color]::Black
                $rtbOutput.AppendText($line + [Environment]::NewLine)
            }
        }

        # Leerzeile bevor Details
        $rtbOutput.AppendText([Environment]::NewLine)

        # Detailausgabe der Alerts mit Farbmarkierung nach Severity
        if ($script:allAlerts -and $script:allAlerts.Count -gt 0) {
            $rtbOutput.SelectionColor = [System.Drawing.Color]::Black
            $rtbOutput.AppendText("=== Alert-Details ===" + [Environment]::NewLine)

            $ordered = $script:allAlerts | Sort-Object `
            @{ Expression = { if ($_.Severity -ieq 'Critical') { 0 } elseif ($_.Severity -ieq 'Warning') { 1 } else { 2 } } }, `
            @{ Expression = { if ($_.State -ieq 'Active') { 0 } else { 1 } } }, `
            @{ Expression = { try { [datetime]$_.Created } catch { [datetime]::MinValue } }; Descending = $true }

            # Begrenzen auf Max. Details
            $limit = [int]$numMaxDetails.Value
            if ($ordered.Count -gt $limit) {
                $rtbOutput.SelectionColor = [System.Drawing.Color]::Black
                $rtbOutput.AppendText(("(auf {0} Details begrenzt)" -f $limit) + [Environment]::NewLine)
                $ordered = $ordered | Select-Object -First $limit
            }

            foreach ($a in $ordered) {
                $color = if ($a.Severity -ieq 'Critical') { [System.Drawing.Color]::Red } elseif ($a.Severity -ieq 'Warning') { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Black }
                $rtbOutput.SelectionColor = $color
                $ts = ''
                try { if ($a.Created) { $ts = ([datetime]$a.Created).ToString('yyyy-MM-dd HH:mm:ss') } } catch {}
                $res = if ($a.Resource) { $a.Resource } else { '' }
                $type = if ($a.Type) { $a.Type } else { '' }
                $msg = if ($a.Message) { $a.Message } else { '' }
                $line = "[{0}][{1}] {2} {3} {4} ({5}): {6}" -f $a.Severity, $a.State, $ts, $a.Appliance, $res, $type, $msg
                $rtbOutput.AppendText($line + [Environment]::NewLine)
            }
            $rtbOutput.SelectionColor = [System.Drawing.Color]::Black
        }
        else {
            $rtbOutput.SelectionColor = [System.Drawing.Color]::Black
            $rtbOutput.AppendText("Keine Fehler oder Warnings beim letzten Check gefunden." + [Environment]::NewLine)
        }
        $rtbOutput.AppendText("Erfasst am: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + [Environment]::NewLine)
        if ($errorDetails.Count -gt 0) {
            $rtbOutput.AppendText("=== Fehlerdetails ===" + [Environment]::NewLine)
            $rtbOutput.SelectionColor = [System.Drawing.Color]::Red
            $rtbOutput.AppendText((($errorDetails -join [Environment]::NewLine)) + [Environment]::NewLine)
            $rtbOutput.SelectionColor = [System.Drawing.Color]::Black
        }
        $rtbOutput.SelectionStart = $rtbOutput.Text.Length
        $rtbOutput.ScrollToCaret()

        # Ampel aktualisieren basierend auf gefilterten Alerts (bekannte ausgeblendete sind bereits entfernt)
        $critCount = ($script:allAlerts | Where-Object { $_.Severity -ieq 'Critical' }).Count
        $warnCount = ($script:allAlerts | Where-Object { $_.Severity -ieq 'Warning' }).Count
        Update-TrafficLight -CriticalCount $critCount -WarningCount $warnCount

        Copy-Item -Path $currentFile -Destination $previousFile -Force
    }
    finally {
        # Erst auf 100 setzen (falls regulär beendet), dann nach kurzer Verzögerung auf 0 zurück
        try { $prgAppliances.Value = 100 } catch {}
        Start-Sleep -Milliseconds 300
        try { $prgAppliances.Value = 0 } catch {}
        $btnStart.Enabled = $true
        $btnExit.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $script:isBusy = $false
        # Timer-Intervall nach UI anpassen
        $timer.Interval = [int]$numInterval.Value * 60000
    }
}

# --------------------- Button-Events -------------------------
$btnStart.Add_Click({ Invoke-Alerts })
$btnViewCurrent.Add_Click({
        try {
            if (-not (Test-Path $currentFile)) {
                [System.Windows.Forms.MessageBox]::Show("Datei nicht gefunden: $currentFile", "Aktuelle Datei", 0, 48) | Out-Null
                return
            }
            $dlg = New-Object System.Windows.Forms.Form
            $dlg.Text = 'Aktuelle Alerts-Datei'
            $dlg.StartPosition = 'CenterParent'
            $dlg.Size = New-Object System.Drawing.Size(800, 600)

            $rt = New-Object System.Windows.Forms.RichTextBox
            $rt.Dock = 'Fill'
            $rt.Font = New-Object System.Drawing.Font('Consolas', 10)
            $rt.ReadOnly = $true
            $rt.WordWrap = $false
            $rt.Text = (Get-Content -Path $currentFile -Raw)
            $dlg.Controls.Add($rt)

            $dlg.ShowDialog($form) | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Öffnen: $_", "Aktuelle Datei", 0, 16) | Out-Null
        }
    })
$btnClear.Add_Click({
        $rtbOutput.Clear()
        $rtbOutput.SelectionStart = 0
        $rtbOutput.ScrollToCaret()
    })
$btnExit.Add_Click({ $form.Close() })

# Bekannte Issues verwalten (einfacher Editor)
$btnManageKnown.Add_Click({
        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = 'Bekannte Issues verwalten'
        $dlg.StartPosition = 'CenterParent'
        $dlg.Size = New-Object System.Drawing.Size(600, 400)
        $dlg.MinimizeBox = $false
        $dlg.MaximizeBox = $false

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Multiline = $true
        $txt.ScrollBars = 'Vertical'
        $txt.Dock = 'Fill'
        $txt.Font = New-Object System.Drawing.Font('Consolas', 10)

        $panelButtons = New-Object System.Windows.Forms.Panel
        $panelButtons.Dock = 'Bottom'
        $panelButtons.Height = 45

        $btnSave = New-Object System.Windows.Forms.Button
        $btnSave.Text = 'Speichern'
        $btnSave.Width = 100
        $btnSave.Height = 28
        $btnSave.Location = New-Object System.Drawing.Point(370, 8)

        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Text = 'Schließen'
        $btnClose.Width = 100
        $btnClose.Height = 28
        $btnClose.Location = New-Object System.Drawing.Point(480, 8)

        # Inhalte laden/speichern
        try {
            if (Test-Path $knownIssuesFile) { $txt.Text = [IO.File]::ReadAllText($knownIssuesFile) } else { $txt.Text = "# Eine Zeile pro Muster (case-insensitive, Teilstring)." }
        }
        catch { $txt.Text = "" }

        $btnSave.Add_Click({
                try {
                    [IO.File]::WriteAllText($knownIssuesFile, $txt.Text)
                    Get-KnownIssues
                    [System.Windows.Forms.MessageBox]::Show('Gespeichert.', 'Info', 0, 64) | Out-Null
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show('Speichern fehlgeschlagen.', 'Fehler', 0, 16) | Out-Null
                }
            })
        $btnClose.Add_Click({ $dlg.Close() })

        $panelButtons.Controls.AddRange(@($btnSave, $btnClose))
        $dlg.Controls.Add($txt)
        $dlg.Controls.Add($panelButtons)
        $dlg.ShowDialog($form) | Out-Null
    })

# --------------------- Automatischer Timer -------------------------
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1800000  # Default 30 Min
$timer.Add_Tick({ if (-not $script:isBusy) { Invoke-Alerts } })
$timer.Start()

# Timer sauber beim Schließen des Forms freigeben
$form.Add_FormClosing({
        try { if ($timer) { $timer.Stop(); $timer.Dispose() } } catch {}
        try { if ($trafficTimer) { $trafficTimer.Stop(); $trafficTimer.Dispose() } } catch {}
        # Zwischengespeicherte Anmeldeinformationen beim Schließen aus dem Speicher entfernen
        try { $script:credential = $null } catch {}
    })

# Intervall dynamisch bei Änderung der NumericUpDown anpassen (sofort wirksam)
$numInterval.Add_ValueChanged({
    try {
        if ($null -ne $timer) {
            $newMs = [int]$numInterval.Value * 60000
            if ($newMs -lt 60000) { $newMs = 60000 }  # Sicherheitsnetz (>=1 Minute laut Minimum)
            if ($timer.Interval -ne $newMs) {
                $timer.Stop()
                $timer.Interval = $newMs
                $timer.Start()
            }
        }
    } catch {}
})

# --------------------- GUI starten ---------------------------
[System.Windows.Forms.Application]::Run($form)