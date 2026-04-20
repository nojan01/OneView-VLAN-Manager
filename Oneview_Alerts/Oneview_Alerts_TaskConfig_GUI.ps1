<#
.SYNOPSIS
    Konfigurations-GUI für den geplanten Task "OneView Alerts".

.DESCRIPTION
    Pflegt:
      - OneView-Credentials (verschlüsselt mit lokalem AES-Key)
      - Appliance-Datei-Auswahl, Zeitraum, Max. Details, Known-Issues-Filter
      - SMTP-Server, Absender, Empfänger, Auth, SSL
      - Geplanten Task (Scheduled Task) registrieren / löschen / testen

    Schreibt alerts_task_config.json im gleichen Ordner. Credentials werden
    in alerts_task_cred.xml abgelegt, verschlüsselt mit alerts_task_key.bin
    (AES) - damit ist die Entschlüsselung auch unter dem SYSTEM-Konto möglich.
    Die Schlüsseldatei sollte entsprechend per NTFS-ACL geschützt werden.
#>

if (-not $IsWindows) { Write-Error "Nur Windows"; return }

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $MyInvocation.MyCommand.Path -Parent }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
$configPath = Join-Path $scriptDir 'alerts_task_config.json'
$keyFile = Join-Path $scriptDir 'alerts_task_key.bin'
$credFile = Join-Path $scriptDir 'alerts_task_cred.xml'
$knownIssuesFile = Join-Path $scriptDir 'KnownIssues.txt'
$runnerScript = Join-Path $scriptDir 'Oneview_Alerts_Scheduled.ps1'

$taskName = 'OneView_Alerts_Daily'

# ---------------------------------------------------------------------------
# Auto-Detect: PowerShell 7 (pwsh.exe) bevorzugt
# ---------------------------------------------------------------------------
function Find-DefaultPwsh {
    $candidates = @(
        "$env:ProgramFiles\PowerShell\7\pwsh.exe",
        "$env:ProgramFiles\PowerShell\pwsh.exe",
        "${env:ProgramFiles(x86)}\PowerShell\7\pwsh.exe"
    )
    foreach ($p in $candidates) { if ($p -and (Test-Path $p)) { return $p } }
    $cmd = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
}

# ---------------------------------------------------------------------------
# Helpers: AES-Key, Config laden/speichern
# ---------------------------------------------------------------------------
function Get-OrCreate-AesKey {
    if (-not (Test-Path $keyFile)) {
        $key = New-Object byte[] 32
        [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($key)
        [IO.File]::WriteAllBytes($keyFile, $key)
        # ACL einschränken: nur aktueller Benutzer + SYSTEM + Administratoren
        try {
            $acl = Get-Acl $keyFile
            $acl.SetAccessRuleProtection($true, $false)
            foreach ($id in @('NT AUTHORITY\SYSTEM', 'BUILTIN\Administrators', "$env:USERDOMAIN\$env:USERNAME")) {
                try {
                    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($id, 'FullControl', 'Allow')
                    $acl.AddAccessRule($rule)
                }
                catch {}
            }
            Set-Acl $keyFile $acl
        }
        catch {}
    }
    return [IO.File]::ReadAllBytes($keyFile)
}

function Load-Config {
    if (Test-Path $configPath) {
        try { return Get-Content -Path $configPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
    }
    # Defaults
    return [PSCustomObject]@{
        ApplianceMode         = 'GOV'
        RangeValue            = 1
        RangeUnit             = 'Days'
        MaxDetails            = 100
        HideKnown             = $true
        OwnerUnknownOnly      = $false
        SendEmail             = $true
        OnlyOnErrors          = $false
        SmtpServer            = ''
        SmtpPort              = 25
        UseSsl                = $false
        SmtpUser              = ''
        SmtpPasswordEncrypted = ''
        MailFrom              = ''
        MailTo                = ''
        SubjectPrefix         = '[OneView Alerts]'
        ScheduleTimes         = @('07:00')
        ScheduleDays          = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')
        ScheduleAtStartup     = $false
        PowerShellExe         = (Find-DefaultPwsh)
        RunnerScript          = $runnerScript
        TaskUserMode          = 'Interactive'
    }
}

function Save-Config($cfg) {
    $cfg | ConvertTo-Json -Depth 5 | Set-Content -Path $configPath -Encoding UTF8
}

function Load-CredentialXml {
    if (Test-Path $credFile) {
        try { return Import-Clixml -Path $credFile } catch {}
    }
    return $null
}

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "OneView Alerts - Task-Konfiguration"
$form.Size = New-Object System.Drawing.Size(720, 1060)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'Sizable'
$form.MinimumSize = New-Object System.Drawing.Size(720, 1060)

# -------- Credentials --------
$grpCred = New-Object System.Windows.Forms.GroupBox
$grpCred.Text = "OneView Credentials (wird verschlüsselt gespeichert)"
$grpCred.Location = New-Object System.Drawing.Point(10, 10)
$grpCred.Size = New-Object System.Drawing.Size(685, 100)
$form.Controls.Add($grpCred)

$lblU = New-Object System.Windows.Forms.Label; $lblU.Text = "Benutzer:"; $lblU.Location = "10,28"; $lblU.Size = "90,20"
$grpCred.Controls.Add($lblU)
$txtUser = New-Object System.Windows.Forms.TextBox; $txtUser.Location = "105,25"; $txtUser.Size = "270,22"
$grpCred.Controls.Add($txtUser)

$lblP = New-Object System.Windows.Forms.Label; $lblP.Text = "Passwort:"; $lblP.Location = "10,58"; $lblP.Size = "90,20"
$grpCred.Controls.Add($lblP)
$txtPass = New-Object System.Windows.Forms.TextBox; $txtPass.Location = "105,55"; $txtPass.Size = "270,22"; $txtPass.UseSystemPasswordChar = $true
$grpCred.Controls.Add($txtPass)

$lblCredInfo = New-Object System.Windows.Forms.Label
$lblCredInfo.Location = "395,28"; $lblCredInfo.Size = "280,50"
$lblCredInfo.Text = "(bei leerem Feld wird der bestehende Wert beibehalten)"
$grpCred.Controls.Add($lblCredInfo)

# -------- Abfrage-Parameter --------
$grpQuery = New-Object System.Windows.Forms.GroupBox
$grpQuery.Text = "Abfrage-Parameter"
$grpQuery.Location = New-Object System.Drawing.Point(10, 120)
$grpQuery.Size = New-Object System.Drawing.Size(685, 140)
$form.Controls.Add($grpQuery)

$lblMode = New-Object System.Windows.Forms.Label; $lblMode.Text = "Appliance-Datei:"; $lblMode.Location = "10,28"; $lblMode.Size = "120,20"
$grpQuery.Controls.Add($lblMode)
$cmbMode = New-Object System.Windows.Forms.ComboBox; $cmbMode.Location = "135,25"; $cmbMode.Size = "180,22"; $cmbMode.DropDownStyle = 'DropDownList'
[void]$cmbMode.Items.AddRange(@('GOV', 'DIV', 'BOTH'))
$grpQuery.Controls.Add($cmbMode)

$lblRange = New-Object System.Windows.Forms.Label; $lblRange.Text = "Zeitraum:"; $lblRange.Location = "335,28"; $lblRange.Size = "70,20"
$grpQuery.Controls.Add($lblRange)
$numRange = New-Object System.Windows.Forms.NumericUpDown; $numRange.Location = "410,25"; $numRange.Size = "70,22"; $numRange.Minimum = 1; $numRange.Maximum = 365
$grpQuery.Controls.Add($numRange)
$cmbRangeUnit = New-Object System.Windows.Forms.ComboBox; $cmbRangeUnit.Location = "485,25"; $cmbRangeUnit.Size = "90,22"; $cmbRangeUnit.DropDownStyle = 'DropDownList'
[void]$cmbRangeUnit.Items.AddRange(@('Days', 'Hours'))
$grpQuery.Controls.Add($cmbRangeUnit)

$lblMax = New-Object System.Windows.Forms.Label; $lblMax.Text = "Max. Details:"; $lblMax.Location = "10,58"; $lblMax.Size = "120,20"
$grpQuery.Controls.Add($lblMax)
$numMax = New-Object System.Windows.Forms.NumericUpDown; $numMax.Location = "135,55"; $numMax.Size = "80,22"; $numMax.Minimum = 1; $numMax.Maximum = 5000
$grpQuery.Controls.Add($numMax)

$chkHide = New-Object System.Windows.Forms.CheckBox; $chkHide.Text = "Bekannte Issues ausblenden"; $chkHide.Location = "240,57"; $chkHide.Size = "220,20"
$grpQuery.Controls.Add($chkHide)

$btnIssues = New-Object System.Windows.Forms.Button; $btnIssues.Text = "Bekannte Issues verwalten"; $btnIssues.Location = "475,54"; $btnIssues.Size = "195,26"
$grpQuery.Controls.Add($btnIssues)

$chkOwnerUnknown = New-Object System.Windows.Forms.CheckBox
$chkOwnerUnknown.Text = "Nur Alerts mit Owner = unknown"
$chkOwnerUnknown.Location = "10,88"; $chkOwnerUnknown.Size = "260,22"
$grpQuery.Controls.Add($chkOwnerUnknown)

# -------- E-Mail --------
$grpMail = New-Object System.Windows.Forms.GroupBox
$grpMail.Text = "E-Mail (SMTP)"
$grpMail.Location = New-Object System.Drawing.Point(10, 270)
$grpMail.Size = New-Object System.Drawing.Size(685, 220)
$form.Controls.Add($grpMail)

$chkSend = New-Object System.Windows.Forms.CheckBox; $chkSend.Text = "E-Mail senden"; $chkSend.Location = "10,25"; $chkSend.Size = "130,22"
$grpMail.Controls.Add($chkSend)
$chkOnlyErr = New-Object System.Windows.Forms.CheckBox; $chkOnlyErr.Text = "Nur bei Alerts / Fehlern"; $chkOnlyErr.Location = "150,25"; $chkOnlyErr.Size = "200,22"
$grpMail.Controls.Add($chkOnlyErr)

$lblSrv = New-Object System.Windows.Forms.Label; $lblSrv.Text = "SMTP-Server:"; $lblSrv.Location = "10,55"; $lblSrv.Size = "120,20"
$grpMail.Controls.Add($lblSrv)
$txtSrv = New-Object System.Windows.Forms.TextBox; $txtSrv.Location = "135,52"; $txtSrv.Size = "280,22"
$grpMail.Controls.Add($txtSrv)

$lblPort = New-Object System.Windows.Forms.Label; $lblPort.Text = "Port:"; $lblPort.Location = "425,55"; $lblPort.Size = "40,20"
$grpMail.Controls.Add($lblPort)
$numPort = New-Object System.Windows.Forms.NumericUpDown; $numPort.Location = "465,52"; $numPort.Size = "70,22"; $numPort.Minimum = 1; $numPort.Maximum = 65535
$grpMail.Controls.Add($numPort)
$chkSsl = New-Object System.Windows.Forms.CheckBox; $chkSsl.Text = "SSL"; $chkSsl.Location = "545,53"; $chkSsl.Size = "60,22"
$grpMail.Controls.Add($chkSsl)

$lblFrom = New-Object System.Windows.Forms.Label; $lblFrom.Text = "Absender:"; $lblFrom.Location = "10,85"; $lblFrom.Size = "120,20"
$grpMail.Controls.Add($lblFrom)
$txtFrom = New-Object System.Windows.Forms.TextBox; $txtFrom.Location = "135,82"; $txtFrom.Size = "400,22"
$grpMail.Controls.Add($txtFrom)

$lblTo = New-Object System.Windows.Forms.Label; $lblTo.Text = "Empfänger:"; $lblTo.Location = "10,115"; $lblTo.Size = "120,20"
$grpMail.Controls.Add($lblTo)
$txtTo = New-Object System.Windows.Forms.TextBox; $txtTo.Location = "135,112"; $txtTo.Size = "520,22"
$grpMail.Controls.Add($txtTo)
$lblToHint = New-Object System.Windows.Forms.Label; $lblToHint.Text = "(Mehrere per ';' oder ',' trennen)"; $lblToHint.Location = "135,135"; $lblToHint.Size = "400,16"; $lblToHint.ForeColor = 'Gray'
$grpMail.Controls.Add($lblToHint)

$lblSU = New-Object System.Windows.Forms.Label; $lblSU.Text = "SMTP-User:"; $lblSU.Location = "10,155"; $lblSU.Size = "120,20"
$grpMail.Controls.Add($lblSU)
$txtSU = New-Object System.Windows.Forms.TextBox; $txtSU.Location = "135,152"; $txtSU.Size = "280,22"
$grpMail.Controls.Add($txtSU)

$lblSP = New-Object System.Windows.Forms.Label; $lblSP.Text = "SMTP-Pass:"; $lblSP.Location = "10,182"; $lblSP.Size = "120,20"
$grpMail.Controls.Add($lblSP)
$txtSP = New-Object System.Windows.Forms.TextBox; $txtSP.Location = "135,180"; $txtSP.Size = "280,22"; $txtSP.UseSystemPasswordChar = $true
$grpMail.Controls.Add($txtSP)
$lblSPhint = New-Object System.Windows.Forms.Label; $lblSPhint.Text = "(leer = unverändert)"; $lblSPhint.Location = "425,183"; $lblSPhint.Size = "115,20"; $lblSPhint.ForeColor = 'Gray'
$grpMail.Controls.Add($lblSPhint)

$lblSubj = New-Object System.Windows.Forms.Label; $lblSubj.Text = "Betreff-Prefix:"; $lblSubj.Location = "335,25"; $lblSubj.Size = "100,20"
$grpMail.Controls.Add($lblSubj)
$txtSubj = New-Object System.Windows.Forms.TextBox; $txtSubj.Location = "440,22"; $txtSubj.Size = "220,22"
$grpMail.Controls.Add($txtSubj)

$btnTestMail = New-Object System.Windows.Forms.Button; $btnTestMail.Text = "Test-Mail senden"; $btnTestMail.Location = "545,178"; $btnTestMail.Size = "130,26"
$grpMail.Controls.Add($btnTestMail)

# -------- Script-Pfade --------
$grpPaths = New-Object System.Windows.Forms.GroupBox
$grpPaths.Text = "Script-Pfade"
$grpPaths.Location = New-Object System.Drawing.Point(10, 500)
$grpPaths.Size = New-Object System.Drawing.Size(685, 90)
$form.Controls.Add($grpPaths)

$lblPwsh = New-Object System.Windows.Forms.Label; $lblPwsh.Text = "PowerShell:"; $lblPwsh.Location = "10,28"; $lblPwsh.Size = "120,20"
$grpPaths.Controls.Add($lblPwsh)
$txtPwsh = New-Object System.Windows.Forms.TextBox; $txtPwsh.Location = "135,25"; $txtPwsh.Size = "440,22"
$grpPaths.Controls.Add($txtPwsh)
$btnPwshBrowse = New-Object System.Windows.Forms.Button; $btnPwshBrowse.Text = "Durchsuchen..."; $btnPwshBrowse.Location = "580,23"; $btnPwshBrowse.Size = "95,26"
$grpPaths.Controls.Add($btnPwshBrowse)

$lblRunner = New-Object System.Windows.Forms.Label; $lblRunner.Text = "Runner-Script:"; $lblRunner.Location = "10,58"; $lblRunner.Size = "120,20"
$grpPaths.Controls.Add($lblRunner)
$txtRunner = New-Object System.Windows.Forms.TextBox; $txtRunner.Location = "135,55"; $txtRunner.Size = "440,22"
$grpPaths.Controls.Add($txtRunner)
$btnRunnerBrowse = New-Object System.Windows.Forms.Button; $btnRunnerBrowse.Text = "Durchsuchen..."; $btnRunnerBrowse.Location = "580,53"; $btnRunnerBrowse.Size = "95,26"
$grpPaths.Controls.Add($btnRunnerBrowse)

$btnPwshBrowse.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "PowerShell (pwsh.exe;powershell.exe)|pwsh.exe;powershell.exe|Alle Dateien (*.*)|*.*"
        $ofd.Title = "PowerShell-Executable auswählen"
        if ($txtPwsh.Text -and (Test-Path $txtPwsh.Text)) { $ofd.InitialDirectory = Split-Path $txtPwsh.Text -Parent }
        if ($ofd.ShowDialog() -eq 'OK') { $txtPwsh.Text = $ofd.FileName }
    })
$btnRunnerBrowse.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "PowerShell-Scripts (*.ps1)|*.ps1|Alle Dateien (*.*)|*.*"
        $ofd.Title = "Runner-Script auswählen"
        if ($txtRunner.Text -and (Test-Path $txtRunner.Text)) { $ofd.InitialDirectory = Split-Path $txtRunner.Text -Parent }
        elseif (Test-Path $scriptDir) { $ofd.InitialDirectory = $scriptDir }
        if ($ofd.ShowDialog() -eq 'OK') { $txtRunner.Text = $ofd.FileName }
    })

# -------- Geplanter Task --------
$grpTask = New-Object System.Windows.Forms.GroupBox
$grpTask.Text = "Geplanter Task"
$grpTask.Location = New-Object System.Drawing.Point(10, 600)
$grpTask.Size = New-Object System.Drawing.Size(685, 340)
$form.Controls.Add($grpTask)

# Startzeiten (Liste mit Hinzufügen/Entfernen)
$lblTimes = New-Object System.Windows.Forms.Label; $lblTimes.Text = "Startzeiten:"; $lblTimes.Location = "10,28"; $lblTimes.Size = "120,20"
$grpTask.Controls.Add($lblTimes)
$lstTimes = New-Object System.Windows.Forms.ListBox; $lstTimes.Location = "135,25"; $lstTimes.Size = "150,70"
$grpTask.Controls.Add($lstTimes)
$dtpTime = New-Object System.Windows.Forms.DateTimePicker; $dtpTime.Format = 'Custom'; $dtpTime.CustomFormat = 'HH:mm'; $dtpTime.ShowUpDown = $true
$dtpTime.Location = "295,25"; $dtpTime.Size = "80,22"; $dtpTime.Value = (Get-Date -Hour 7 -Minute 0 -Second 0)
$grpTask.Controls.Add($dtpTime)
$btnTimeAdd = New-Object System.Windows.Forms.Button; $btnTimeAdd.Text = "Hinzufügen"; $btnTimeAdd.Location = "385,24"; $btnTimeAdd.Size = "100,24"
$grpTask.Controls.Add($btnTimeAdd)
$btnTimeDel = New-Object System.Windows.Forms.Button; $btnTimeDel.Text = "Entfernen"; $btnTimeDel.Location = "385,52"; $btnTimeDel.Size = "100,24"
$grpTask.Controls.Add($btnTimeDel)
$btnTimeAdd.Add_Click({
        $t = $dtpTime.Value.ToString('HH:mm')
        if (-not ($lstTimes.Items -contains $t)) { [void]$lstTimes.Items.Add($t) }
    })
$btnTimeDel.Add_Click({
        if ($lstTimes.SelectedIndex -ge 0) { $lstTimes.Items.RemoveAt($lstTimes.SelectedIndex) }
    })

# Wochentage
$lblDays = New-Object System.Windows.Forms.Label; $lblDays.Text = "Wochentage:"; $lblDays.Location = "10,110"; $lblDays.Size = "120,20"
$grpTask.Controls.Add($lblDays)
$dayDefs = @(
    @{ Key = 'Monday'; Label = 'Mo' },
    @{ Key = 'Tuesday'; Label = 'Di' },
    @{ Key = 'Wednesday'; Label = 'Mi' },
    @{ Key = 'Thursday'; Label = 'Do' },
    @{ Key = 'Friday'; Label = 'Fr' },
    @{ Key = 'Saturday'; Label = 'Sa' },
    @{ Key = 'Sunday'; Label = 'So' }
)
$dayChecks = @{}
$x = 135
foreach ($d in $dayDefs) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $d.Label; $cb.Location = "$x,108"; $cb.Size = "50,22"
    $cb.Checked = $true
    $grpTask.Controls.Add($cb)
    $dayChecks[$d.Key] = $cb
    $x += 55
}

# Zusätzlicher Trigger bei Systemstart
$chkAtBoot = New-Object System.Windows.Forms.CheckBox
$chkAtBoot.Text = "Zusätzlich bei Systemstart ausführen"
$chkAtBoot.Location = "10,138"; $chkAtBoot.Size = "320,22"
$grpTask.Controls.Add($chkAtBoot)

# Task-Benutzer
$lblUser2 = New-Object System.Windows.Forms.Label; $lblUser2.Text = "Task-Benutzer:"; $lblUser2.Location = "10,168"; $lblUser2.Size = "120,20"
$grpTask.Controls.Add($lblUser2)
$cmbUser = New-Object System.Windows.Forms.ComboBox; $cmbUser.Location = "135,165"; $cmbUser.Size = "200,22"; $cmbUser.DropDownStyle = 'DropDownList'
[void]$cmbUser.Items.AddRange(@('Aktueller Benutzer (interaktiv)', 'Aktueller Benutzer (S4U, ohne Passwort)', 'SYSTEM', 'NETWORK SERVICE', 'Eigener Benutzer...'))
$cmbUser.SelectedIndex = 0
$grpTask.Controls.Add($cmbUser)
$txtTaskUser = New-Object System.Windows.Forms.TextBox; $txtTaskUser.Location = "345,165"; $txtTaskUser.Size = "180,22"; $txtTaskUser.Enabled = $false
$txtTaskUser.PlaceholderText = "DOMAIN\User"
$grpTask.Controls.Add($txtTaskUser)
$cmbUser.Add_SelectedIndexChanged({ $txtTaskUser.Enabled = ($cmbUser.SelectedItem -eq 'Eigener Benutzer...') })

# Hinweistext zum gewählten Modus
$lblUserInfo = New-Object System.Windows.Forms.Label
$lblUserInfo.Location = "135,190"; $lblUserInfo.Size = "535,50"
$lblUserInfo.ForeColor = 'DarkBlue'
$grpTask.Controls.Add($lblUserInfo)

$updateUserInfo = {
    switch ([string]$cmbUser.SelectedItem) {
        'Aktueller Benutzer (interaktiv)' {
            $lblUserInfo.Text = "Läuft NUR wenn du interaktiv eingeloggt bist. Kein Passwort nötig.`r`nNach Logout/Reboot wird der Task nicht ausgeführt."
        }
        'Aktueller Benutzer (S4U, ohne Passwort)' {
            $lblUserInfo.Text = "Läuft auch ohne Login, aber OHNE Netzwerk-Zugriff.`r`nFür SMTP/OneView-Zugriff NICHT geeignet."
        }
        'SYSTEM' {
            $lblUserInfo.Text = "Läuft immer, ohne Passwort. Netzwerk nur als Computer-Account (HOSTNAME$).`r`nSMTP-Relays akzeptieren das oft nicht."
        }
        'NETWORK SERVICE' {
            $lblUserInfo.Text = "Läuft immer, ohne Passwort. Eingeschränkte Rechte."
        }
        'Eigener Benutzer...' {
            $lblUserInfo.Text = "EMPFOHLEN: Domain-Account (z.B. CORP\nojan) + Passwort.`r`nLäuft immer, mit vollem Netzwerk-Zugriff. Bei Passwort-Wechsel (z.B. alle 90 Tage) hier neu eintragen."
        }
        default { $lblUserInfo.Text = '' }
    }
}
$cmbUser.Add_SelectedIndexChanged($updateUserInfo)
& $updateUserInfo

# Task-Passwort
$lblTaskPw = New-Object System.Windows.Forms.Label; $lblTaskPw.Text = "Task-Passwort:"; $lblTaskPw.Location = "10,250"; $lblTaskPw.Size = "120,20"
$grpTask.Controls.Add($lblTaskPw)
$txtTaskPw = New-Object System.Windows.Forms.TextBox; $txtTaskPw.Location = "135,247"; $txtTaskPw.Size = "200,22"; $txtTaskPw.UseSystemPasswordChar = $true
$grpTask.Controls.Add($txtTaskPw)
$lblTaskPwHint = New-Object System.Windows.Forms.Label; $lblTaskPwHint.Text = "(nur bei 'Eigener Benutzer' benötigt)"; $lblTaskPwHint.Location = "340,250"; $lblTaskPwHint.Size = "270,20"; $lblTaskPwHint.ForeColor = 'Gray'
$grpTask.Controls.Add($lblTaskPwHint)

$btnTaskInstall = New-Object System.Windows.Forms.Button; $btnTaskInstall.Text = "Task anlegen/aktualisieren"; $btnTaskInstall.Location = "10,290"; $btnTaskInstall.Size = "200,30"
$grpTask.Controls.Add($btnTaskInstall)
$btnTaskRemove = New-Object System.Windows.Forms.Button; $btnTaskRemove.Text = "Task entfernen"; $btnTaskRemove.Location = "220,290"; $btnTaskRemove.Size = "140,30"
$grpTask.Controls.Add($btnTaskRemove)
$btnTaskRun = New-Object System.Windows.Forms.Button; $btnTaskRun.Text = "Task jetzt ausführen"; $btnTaskRun.Location = "370,290"; $btnTaskRun.Size = "160,30"
$grpTask.Controls.Add($btnTaskRun)
$btnTaskStatus = New-Object System.Windows.Forms.Button; $btnTaskStatus.Text = "Status"; $btnTaskStatus.Location = "540,290"; $btnTaskStatus.Size = "130,30"
$grpTask.Controls.Add($btnTaskStatus)

# -------- Aktionen --------
$pnlBtn = New-Object System.Windows.Forms.Panel; $pnlBtn.Location = "10,950"; $pnlBtn.Size = "685,40"
$form.Controls.Add($pnlBtn)
$btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text = "Konfiguration speichern"; $btnSave.Location = "0,5"; $btnSave.Size = "200,30"
$pnlBtn.Controls.Add($btnSave)
$btnClose = New-Object System.Windows.Forms.Button; $btnClose.Text = "Schließen"; $btnClose.Location = "210,5"; $btnClose.Size = "120,30"
$pnlBtn.Controls.Add($btnClose)
$btnClose.Add_Click({ $form.Close() })

# ---------------------------------------------------------------------------
# Config in GUI laden
# ---------------------------------------------------------------------------
$config = Load-Config
$cmbMode.SelectedItem = $config.ApplianceMode
if (-not $cmbMode.SelectedItem) { $cmbMode.SelectedIndex = 0 }
$numRange.Value = [int]$config.RangeValue
$cmbRangeUnit.SelectedItem = $config.RangeUnit
if (-not $cmbRangeUnit.SelectedItem) { $cmbRangeUnit.SelectedIndex = 0 }
$numMax.Value = [int]$config.MaxDetails
$chkHide.Checked = [bool]$config.HideKnown
$chkOwnerUnknown.Checked = if ($null -ne $config.OwnerUnknownOnly) { [bool]$config.OwnerUnknownOnly } else { $false }
$chkSend.Checked = [bool]$config.SendEmail
$chkOnlyErr.Checked = [bool]$config.OnlyOnErrors
$txtSrv.Text = [string]$config.SmtpServer
$numPort.Value = if ($config.SmtpPort) { [int]$config.SmtpPort } else { 25 }
$chkSsl.Checked = [bool]$config.UseSsl
$txtFrom.Text = [string]$config.MailFrom
$txtTo.Text = [string]$config.MailTo
$txtSU.Text = [string]$config.SmtpUser
$txtSubj.Text = if ($config.SubjectPrefix) { $config.SubjectPrefix } else { '[OneView Alerts]' }

# Schedule-Teil laden
$lstTimes.Items.Clear()
$loadedTimes = @()
if ($config.ScheduleTimes) { $loadedTimes = @($config.ScheduleTimes) }
if ($loadedTimes.Count -eq 0) { $loadedTimes = @('07:00') }
foreach ($t in $loadedTimes) { [void]$lstTimes.Items.Add([string]$t) }

$loadedDays = @()
if ($config.ScheduleDays) { $loadedDays = @($config.ScheduleDays) }
if ($loadedDays.Count -eq 0) { $loadedDays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday') }
foreach ($k in $dayChecks.Keys) { $dayChecks[$k].Checked = ($loadedDays -contains $k) }

$chkAtBoot.Checked = [bool]$config.ScheduleAtStartup

# Pfade-Felder
$txtPwsh.Text = if ($config.PowerShellExe) { [string]$config.PowerShellExe } else { Find-DefaultPwsh }
$txtRunner.Text = if ($config.RunnerScript) { [string]$config.RunnerScript } else { $runnerScript }

# Task-User-Modus
switch ([string]$config.TaskUserMode) {
    'Interactive' { $cmbUser.SelectedIndex = 0 }
    'S4U' { $cmbUser.SelectedIndex = 1 }
    'SYSTEM' { $cmbUser.SelectedIndex = 2 }
    'NETWORK SERVICE' { $cmbUser.SelectedIndex = 3 }
    'Custom' {
        $cmbUser.SelectedIndex = 4
        if ($config.TaskUserName) { $txtTaskUser.Text = [string]$config.TaskUserName }
    }
    default { $cmbUser.SelectedIndex = 0 }
}

$credXml = Load-CredentialXml
if ($credXml) { $txtUser.Text = [string]$credXml.Username }

# ---------------------------------------------------------------------------
# Known Issues Editor
# ---------------------------------------------------------------------------
$btnIssues.Add_Click({
        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = 'Bekannte Issues verwalten'
        $dlg.Size = New-Object System.Drawing.Size(640, 460); $dlg.StartPosition = 'CenterParent'
        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Multiline = $true; $txt.ScrollBars = 'Vertical'; $txt.Dock = 'Fill'
        $txt.Font = New-Object System.Drawing.Font('Consolas', 10)
        $pnl = New-Object System.Windows.Forms.Panel; $pnl.Dock = 'Bottom'; $pnl.Height = 45
        $bs = New-Object System.Windows.Forms.Button; $bs.Text = 'Speichern'; $bs.Location = '400,8'; $bs.Size = '100,28'
        $bc = New-Object System.Windows.Forms.Button; $bc.Text = 'Schließen'; $bc.Location = '510,8'; $bc.Size = '100,28'
        $pnl.Controls.AddRange(@($bs, $bc))
        if (Test-Path $knownIssuesFile) { $txt.Text = [IO.File]::ReadAllText($knownIssuesFile) }
        else { $txt.Text = "# Eine Zeile pro Muster (case-insensitive, Teilstring)." }
        $bs.Add_Click({ [IO.File]::WriteAllText($knownIssuesFile, $txt.Text); [System.Windows.Forms.MessageBox]::Show('Gespeichert.', 'Info', 0, 64) | Out-Null })
        $bc.Add_Click({ $dlg.Close() })
        $dlg.Controls.Add($txt); $dlg.Controls.Add($pnl)
        $dlg.ShowDialog($form) | Out-Null
    })

# ---------------------------------------------------------------------------
# Konfiguration speichern (inkl. Credentials)
# ---------------------------------------------------------------------------
function Save-AllConfig {
    $aesKey = Get-OrCreate-AesKey

    # OneView Credentials
    if ($txtUser.Text.Trim() -and $txtPass.Text.Length -gt 0) {
        $sec = ConvertTo-SecureString $txtPass.Text -AsPlainText -Force
        $encPw = ConvertFrom-SecureString -SecureString $sec -Key $aesKey
        [PSCustomObject]@{ Username = $txtUser.Text.Trim(); EncryptedPassword = $encPw } |
        Export-Clixml -Path $credFile
        try {
            $acl = Get-Acl $credFile
            $acl.SetAccessRuleProtection($true, $false)
            foreach ($id in @('NT AUTHORITY\SYSTEM', 'BUILTIN\Administrators', "$env:USERDOMAIN\$env:USERNAME")) {
                try { $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($id, 'FullControl', 'Allow'))) } catch {}
            }
            Set-Acl $credFile $acl
        }
        catch {}
        $txtPass.Text = ''
    }
    elseif ($txtUser.Text.Trim() -and -not (Test-Path $credFile)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Passwort setzen.", 'Hinweis', 0, 48) | Out-Null
        return $false
    }

    # SMTP-Passwort ggf. neu verschlüsseln
    $smtpEnc = [string]$config.SmtpPasswordEncrypted
    if ($txtSP.Text.Length -gt 0) {
        $sec = ConvertTo-SecureString $txtSP.Text -AsPlainText -Force
        $smtpEnc = ConvertFrom-SecureString -SecureString $sec -Key $aesKey
        $txtSP.Text = ''
    }

    $newCfg = [PSCustomObject]@{
        ApplianceMode         = [string]$cmbMode.SelectedItem
        RangeValue            = [int]$numRange.Value
        RangeUnit             = [string]$cmbRangeUnit.SelectedItem
        MaxDetails            = [int]$numMax.Value
        HideKnown             = [bool]$chkHide.Checked
        OwnerUnknownOnly      = [bool]$chkOwnerUnknown.Checked
        SendEmail             = [bool]$chkSend.Checked
        OnlyOnErrors          = [bool]$chkOnlyErr.Checked
        SmtpServer            = $txtSrv.Text.Trim()
        SmtpPort              = [int]$numPort.Value
        UseSsl                = [bool]$chkSsl.Checked
        SmtpUser              = $txtSU.Text.Trim()
        SmtpPasswordEncrypted = $smtpEnc
        MailFrom              = $txtFrom.Text.Trim()
        MailTo                = $txtTo.Text.Trim()
        SubjectPrefix         = $txtSubj.Text.Trim()
        ScheduleTimes         = @($lstTimes.Items | ForEach-Object { [string]$_ })
        ScheduleDays          = @($dayChecks.Keys | Where-Object { $dayChecks[$_].Checked })
        ScheduleAtStartup     = [bool]$chkAtBoot.Checked
        PowerShellExe         = $txtPwsh.Text.Trim()
        RunnerScript          = $txtRunner.Text.Trim()
        TaskUserMode          = switch ([string]$cmbUser.SelectedItem) {
            'Aktueller Benutzer (interaktiv)' { 'Interactive' }
            'Aktueller Benutzer (S4U, ohne Passwort)' { 'S4U' }
            'SYSTEM' { 'SYSTEM' }
            'NETWORK SERVICE' { 'NETWORK SERVICE' }
            'Eigener Benutzer...' { 'Custom' }
            default { 'Interactive' }
        }
        TaskUserName          = if ([string]$cmbUser.SelectedItem -eq 'Eigener Benutzer...') { $txtTaskUser.Text.Trim() } else { '' }
    }
    Save-Config $newCfg
    $script:config = $newCfg
    return $true
}

$btnSave.Add_Click({
        if (Save-AllConfig) { [System.Windows.Forms.MessageBox]::Show('Konfiguration gespeichert.', 'OK', 0, 64) | Out-Null }
    })

# ---------------------------------------------------------------------------
# Test-Mail
# ---------------------------------------------------------------------------
$btnTestMail.Add_Click({
        if (-not (Save-AllConfig)) { return }
        $aesKey = Get-OrCreate-AesKey
        $cfg = Load-Config
        try {
            if (-not $cfg.SmtpServer -or -not $cfg.MailFrom -or -not $cfg.MailTo) {
                [System.Windows.Forms.MessageBox]::Show('SMTP-Server, Absender und Empfänger müssen gesetzt sein.', 'Test-Mail', 0, 48) | Out-Null
                return
            }

            # Zertifikatsvalidierung weich setzen (interne CA / Self-Signed)
            try {
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { param($s, $c, $ch, $e) return $true }
            }
            catch { }

            $smtpPort = if ($cfg.SmtpPort) { [int]$cfg.SmtpPort } else { 25 }
            $smtpClient = New-Object Net.Mail.SmtpClient($cfg.SmtpServer, $smtpPort)
            # STARTTLS erzwingen (SmtpClient nutzt bei Port 25/587 automatisch STARTTLS wenn EnableSsl=true)
            $smtpClient.EnableSsl = $true

            if ($cfg.SmtpUser -and $cfg.SmtpPasswordEncrypted) {
                $sp = ConvertTo-SecureString -String $cfg.SmtpPasswordEncrypted -Key $aesKey
                $smtpClient.Credentials = (New-Object System.Management.Automation.PSCredential($cfg.SmtpUser, $sp)).GetNetworkCredential()
            }

            $mailMessage = New-Object System.Net.Mail.MailMessage
            $mailMessage.From = New-Object System.Net.Mail.MailAddress($cfg.MailFrom)
            foreach ($rcpt in @($cfg.MailTo -split '\s*;\s*|\s*,\s*' | Where-Object { $_ })) {
                $mailMessage.To.Add($rcpt)
            }
            $mailMessage.Subject = "$($cfg.SubjectPrefix) TEST"
            $mailMessage.Body = "Dies ist eine Test-E-Mail von der OneView Alerts Task-Konfiguration.`r`nZeit: $(Get-Date)"
            $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
            $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8

            $smtpClient.Send($mailMessage)
            $mailMessage.Dispose()
            $smtpClient.Dispose()

            [System.Windows.Forms.MessageBox]::Show('Test-Mail gesendet.', 'OK', 0, 64) | Out-Null
        }
        catch {
            $errMsg = $_.Exception.Message
            if ($_.Exception.InnerException) {
                $errMsg += "`r`n`r`nInner: $($_.Exception.InnerException.Message)"
            }
            [System.Windows.Forms.MessageBox]::Show("Fehler: $errMsg", 'Test-Mail', 0, 16) | Out-Null
        }
    })

# ---------------------------------------------------------------------------
# Scheduled Task Verwaltung
# ---------------------------------------------------------------------------
function Get-TaskPrincipalAndCred {
    $userChoice = [string]$cmbUser.SelectedItem
    $me = "$env:USERDOMAIN\$env:USERNAME"
    switch ($userChoice) {
        'Aktueller Benutzer (interaktiv)' { return @{ User = $me; LogonType = 'Interactive'; Password = $null } }
        'Aktueller Benutzer (S4U, ohne Passwort)' { return @{ User = $me; LogonType = 'S4U'; Password = $null } }
        'SYSTEM' { return @{ User = 'SYSTEM'; LogonType = 'ServiceAccount'; Password = $null } }
        'NETWORK SERVICE' { return @{ User = 'NT AUTHORITY\NETWORK SERVICE'; LogonType = 'ServiceAccount'; Password = $null } }
        'Eigener Benutzer...' {
            $u = $txtTaskUser.Text.Trim()
            if (-not $u) { throw "Bitte Task-Benutzer angeben." }
            if (-not $txtTaskPw.Text) { throw "Bitte Task-Passwort angeben." }
            return @{ User = $u; LogonType = 'Password'; Password = $txtTaskPw.Text }
        }
    }
}

$btnTaskInstall.Add_Click({
        if (-not (Save-AllConfig)) { return }

        # Pfade aus GUI verwenden
        $pwshExe = $txtPwsh.Text.Trim()
        $runner  = $txtRunner.Text.Trim()
        if (-not $pwshExe -or -not (Test-Path $pwshExe)) {
            [System.Windows.Forms.MessageBox]::Show("PowerShell-Executable nicht gefunden: $pwshExe", 'Fehler', 0, 16) | Out-Null
            return
        }
        if (-not $runner -or -not (Test-Path $runner)) {
            [System.Windows.Forms.MessageBox]::Show("Runner-Script nicht gefunden: $runner", 'Fehler', 0, 16) | Out-Null
            return
        }
        $runnerDir = Split-Path $runner -Parent
        # Zeiten und Tage einsammeln
        $times = @($lstTimes.Items | ForEach-Object { [string]$_ })
        if ($times.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine Startzeit hinzufügen.", 'Hinweis', 0, 48) | Out-Null
            return
        }
        $selectedDays = @($dayChecks.Keys | Where-Object { $dayChecks[$_].Checked })
        if ($selectedDays.Count -eq 0 -and -not $chkAtBoot.Checked) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens einen Wochentag wählen (oder 'Bei Systemstart' aktivieren).", 'Hinweis', 0, 48) | Out-Null
            return
        }

        try {
            $p = Get-TaskPrincipalAndCred
            $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$runner`""
            $action = New-ScheduledTaskAction -Execute $pwshExe -Argument $arg -WorkingDirectory $runnerDir

            # Trigger: pro Startzeit ein Trigger (täglich oder wöchentlich je nach Tagesauswahl)
            $triggers = @()
            $allDays = ($selectedDays.Count -eq 7)
            foreach ($t in $times) {
                if ($allDays) {
                    $triggers += (New-ScheduledTaskTrigger -Daily -At $t)
                }
                elseif ($selectedDays.Count -gt 0) {
                    $triggers += (New-ScheduledTaskTrigger -Weekly -DaysOfWeek $selectedDays -At $t)
                }
            }
            if ($chkAtBoot.Checked) {
                $triggers += (New-ScheduledTaskTrigger -AtStartup)
            }

            $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)

            if ($p.LogonType -eq 'ServiceAccount') {
                $principal = New-ScheduledTaskPrincipal -UserId $p.User -LogonType ServiceAccount -RunLevel Highest
                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal -Force | Out-Null
            }
            elseif ($p.LogonType -eq 'Interactive') {
                # Läuft nur wenn der Benutzer eingeloggt ist - kein Passwort nötig
                $principal = New-ScheduledTaskPrincipal -UserId $p.User -LogonType Interactive -RunLevel Limited
                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal -Force | Out-Null
            }
            elseif ($p.LogonType -eq 'S4U') {
                # Service-for-User: läuft auch ohne aktive Sitzung, ohne Passwort,
                # aber ohne Netzwerk-Zugriff. Benötigt "SeBatchLogonRight" für den User.
                $principal = New-ScheduledTaskPrincipal -UserId $p.User -LogonType S4U -RunLevel Limited
                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal -Force | Out-Null
            }
            else {
                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Settings $settings -User $p.User -Password $p.Password -RunLevel Highest -Force | Out-Null
                $txtTaskPw.Text = ''
            }
            $msg = "Task '$taskName' registriert.`r`nStartzeiten: $($times -join ', ')"
            if ($selectedDays.Count -gt 0 -and -not $allDays) { $msg += "`r`nTage: $($selectedDays -join ', ')" }
            if ($chkAtBoot.Checked) { $msg += "`r`n+ Trigger bei Systemstart" }
            [System.Windows.Forms.MessageBox]::Show($msg, 'OK', 0, 64) | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)", 'Task-Registrierung', 0, 16) | Out-Null
        }
    })

$btnTaskRemove.Add_Click({
        try {
            if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                [System.Windows.Forms.MessageBox]::Show("Task entfernt.", 'OK', 0, 64) | Out-Null
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Task '$taskName' existiert nicht.", 'Hinweis', 0, 48) | Out-Null
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)", 'Fehler', 0, 16) | Out-Null
        }
    })

$btnTaskRun.Add_Click({
        try {
            Start-ScheduledTask -TaskName $taskName -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show("Task gestartet. Ergebnisse siehe Logs-Ordner / E-Mail.", 'OK', 0, 64) | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)", 'Task-Ausführung', 0, 16) | Out-Null
        }
    })

$btnTaskStatus.Add_Click({
        try {
            $t = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
            $info = Get-ScheduledTaskInfo -TaskName $taskName
            $msg = @(
                "Task:        $($t.TaskName)"
                "Status:      $($t.State)"
                "Letzter Lauf: $($info.LastRunTime)"
                "Ergebnis:    $($info.LastTaskResult)"
                "Nächster:    $($info.NextRunTime)"
            ) -join [Environment]::NewLine
            [System.Windows.Forms.MessageBox]::Show($msg, 'Task-Status', 0, 64) | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Task nicht gefunden oder Fehler: $($_.Exception.Message)", 'Status', 0, 48) | Out-Null
        }
    })

[System.Windows.Forms.Application]::Run($form)
