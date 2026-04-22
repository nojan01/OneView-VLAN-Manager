#Requires -Version 7.0
<#
.SYNOPSIS
    Konfigurations-GUI für den geplanten Task "OneView Backup".

.DESCRIPTION
    Pflegt:
      - OneView-Credentials (verschlüsselt mit lokalem AES-Key)
      - Backup-Passphrase (verschlüsselt)
      - IP-Dateien (OV 6.60 / OV 11.10), Backup-Zielordner, Retention
      - Optional: PSCP-Transfer (Host, User, Remote-Pfad, Pfad zu pscp/plink)
      - Optional: SMTP/E-Mail Versand
      - Geplanter Task (Scheduled Task) registrieren / löschen / testen

    Schreibt backup_task_config.json im gleichen Ordner. Credentials werden
    in backup_task_cred.xml abgelegt, verschlüsselt mit backup_task_key.bin
    (AES) - damit ist die Entschlüsselung auch unter dem SYSTEM-Konto möglich.
#>

if (-not $IsWindows) { Write-Error "Nur Windows"; return }

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $MyInvocation.MyCommand.Path -Parent }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
$configPath = Join-Path $scriptDir 'backup_task_config.json'
$keyFile = Join-Path $scriptDir 'backup_task_key.bin'
$credFile = Join-Path $scriptDir 'backup_task_cred.xml'
$runnerScript = Join-Path $scriptDir 'Backup_OneView_Scheduled.ps1'

$taskName = 'OneView_Backup_Daily'

# ---------------------------------------------------------------------------
# Helpers
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

function Get-OrCreate-AesKey {
    if (-not (Test-Path $keyFile)) {
        $key = New-Object byte[] 32
        [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($key)
        [IO.File]::WriteAllBytes($keyFile, $key)
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
    return [PSCustomObject]@{
        IPFile660             = (Join-Path $scriptDir '..\Oneview_660.txt')
        IPFile1110            = (Join-Path $scriptDir '..\Oneview.txt')
        BackupBaseDir         = (Join-Path $scriptDir 'OneView_Backup')
        LocalRetentionDays    = 5
        TransferEnabled       = $false
        TransferHost          = ''
        TransferUser          = ''
        TransferRemotePath    = ''
        PscpPath              = ''
        PlinkPath             = ''
        RemoteCleanupEnabled  = $false
        RemoteRetentionDays   = 30
        SendEmail             = $false
        OnlyOnErrors          = $false
        SmtpServer            = ''
        SmtpPort              = 25
        UseSsl                = $false
        SmtpUser              = ''
        SmtpPasswordEncrypted = ''
        MailFrom              = ''
        MailTo                = ''
        SubjectPrefix         = '[OneView Backup]'
        ScheduleTimes         = @('03:00')
        ScheduleDays          = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')
        ScheduleAtStartup     = $false
        PowerShellExe         = (Find-DefaultPwsh)
        RunnerScript          = $runnerScript
        TaskUserMode          = 'Interactive'
        TaskUserName          = ''
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
$form.Text = "OneView Backup - Task-Konfiguration"
$form.Size = New-Object System.Drawing.Size(740, 1120)
$form.StartPosition = 'CenterScreen'
$form.AutoScroll = $true
$form.FormBorderStyle = 'Sizable'
$form.MinimumSize = New-Object System.Drawing.Size(740, 700)

# -------- Credentials --------
$grpCred = New-Object System.Windows.Forms.GroupBox
$grpCred.Text = "OneView Credentials + Backup-Passphrase (verschlüsselt gespeichert)"
$grpCred.Location = New-Object System.Drawing.Point(10, 10)
$grpCred.Size = New-Object System.Drawing.Size(700, 140)
$form.Controls.Add($grpCred)

$lblU = New-Object System.Windows.Forms.Label; $lblU.Text = "Benutzer:"; $lblU.Location = "10,28"; $lblU.Size = "120,20"
$grpCred.Controls.Add($lblU)
$txtUser = New-Object System.Windows.Forms.TextBox; $txtUser.Location = "135,25"; $txtUser.Size = "260,22"
$grpCred.Controls.Add($txtUser)

$lblP = New-Object System.Windows.Forms.Label; $lblP.Text = "Passwort:"; $lblP.Location = "10,58"; $lblP.Size = "120,20"
$grpCred.Controls.Add($lblP)
$txtPass = New-Object System.Windows.Forms.TextBox; $txtPass.Location = "135,55"; $txtPass.Size = "260,22"; $txtPass.UseSystemPasswordChar = $true
$grpCred.Controls.Add($txtPass)

$lblPP = New-Object System.Windows.Forms.Label; $lblPP.Text = "Backup-Passphrase:"; $lblPP.Location = "10,88"; $lblPP.Size = "120,20"
$grpCred.Controls.Add($lblPP)
$txtPassphrase = New-Object System.Windows.Forms.TextBox; $txtPassphrase.Location = "135,85"; $txtPassphrase.Size = "260,22"; $txtPassphrase.UseSystemPasswordChar = $true
$grpCred.Controls.Add($txtPassphrase)

$lblCredInfo = New-Object System.Windows.Forms.Label
$lblCredInfo.Location = "410,28"; $lblCredInfo.Size = "280,90"
$lblCredInfo.Text = "Leere Felder lassen den bestehenden Wert unverändert.`r`n`r`nDie Backup-Passphrase wird benötigt, um das OneView-Backup zu verschlüsseln."
$grpCred.Controls.Add($lblCredInfo)

# -------- Pfade / IP-Dateien --------
$grpPathsIP = New-Object System.Windows.Forms.GroupBox
$grpPathsIP.Text = "IP-Dateien und Backup-Ziel"
$grpPathsIP.Location = New-Object System.Drawing.Point(10, 160)
$grpPathsIP.Size = New-Object System.Drawing.Size(700, 170)
$form.Controls.Add($grpPathsIP)

$lbl660 = New-Object System.Windows.Forms.Label; $lbl660.Text = "OV 6.60 IP-Datei:"; $lbl660.Location = "10,28"; $lbl660.Size = "130,20"
$grpPathsIP.Controls.Add($lbl660)
$txtIP660 = New-Object System.Windows.Forms.TextBox; $txtIP660.Location = "145,25"; $txtIP660.Size = "440,22"
$grpPathsIP.Controls.Add($txtIP660)
$btnIP660 = New-Object System.Windows.Forms.Button; $btnIP660.Text = "..."; $btnIP660.Location = "590,23"; $btnIP660.Size = "60,26"
$grpPathsIP.Controls.Add($btnIP660)

$lbl1110 = New-Object System.Windows.Forms.Label; $lbl1110.Text = "OV 11.10 IP-Datei:"; $lbl1110.Location = "10,58"; $lbl1110.Size = "130,20"
$grpPathsIP.Controls.Add($lbl1110)
$txtIP1110 = New-Object System.Windows.Forms.TextBox; $txtIP1110.Location = "145,55"; $txtIP1110.Size = "440,22"
$grpPathsIP.Controls.Add($txtIP1110)
$btnIP1110 = New-Object System.Windows.Forms.Button; $btnIP1110.Text = "..."; $btnIP1110.Location = "590,53"; $btnIP1110.Size = "60,26"
$grpPathsIP.Controls.Add($btnIP1110)

$lblDir = New-Object System.Windows.Forms.Label; $lblDir.Text = "Backup-Zielordner:"; $lblDir.Location = "10,88"; $lblDir.Size = "130,20"
$grpPathsIP.Controls.Add($lblDir)
$txtBaseDir = New-Object System.Windows.Forms.TextBox; $txtBaseDir.Location = "145,85"; $txtBaseDir.Size = "440,22"
$grpPathsIP.Controls.Add($txtBaseDir)
$btnBaseDir = New-Object System.Windows.Forms.Button; $btnBaseDir.Text = "..."; $btnBaseDir.Location = "590,83"; $btnBaseDir.Size = "60,26"
$grpPathsIP.Controls.Add($btnBaseDir)

$lblRet = New-Object System.Windows.Forms.Label; $lblRet.Text = "Lokale Retention (Tage):"; $lblRet.Location = "10,118"; $lblRet.Size = "150,20"
$grpPathsIP.Controls.Add($lblRet)
$numRet = New-Object System.Windows.Forms.NumericUpDown; $numRet.Location = "165,115"; $numRet.Size = "70,22"; $numRet.Minimum = 0; $numRet.Maximum = 3650
$grpPathsIP.Controls.Add($numRet)
$lblRetHint = New-Object System.Windows.Forms.Label; $lblRetHint.Text = "(0 = keine Bereinigung)"; $lblRetHint.Location = "240,118"; $lblRetHint.Size = "200,20"; $lblRetHint.ForeColor = 'Gray'
$grpPathsIP.Controls.Add($lblRetHint)

$btnIP660.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "Textdateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
        if ($txtIP660.Text -and (Test-Path (Split-Path $txtIP660.Text -Parent -ErrorAction SilentlyContinue))) {
            $ofd.InitialDirectory = Split-Path $txtIP660.Text -Parent
        }
        if ($ofd.ShowDialog() -eq 'OK') { $txtIP660.Text = $ofd.FileName }
    })
$btnIP1110.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "Textdateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
        if ($txtIP1110.Text -and (Test-Path (Split-Path $txtIP1110.Text -Parent -ErrorAction SilentlyContinue))) {
            $ofd.InitialDirectory = Split-Path $txtIP1110.Text -Parent
        }
        if ($ofd.ShowDialog() -eq 'OK') { $txtIP1110.Text = $ofd.FileName }
    })
$btnBaseDir.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($txtBaseDir.Text -and (Test-Path $txtBaseDir.Text)) { $fbd.SelectedPath = $txtBaseDir.Text }
        if ($fbd.ShowDialog() -eq 'OK') { $txtBaseDir.Text = $fbd.SelectedPath }
    })

# -------- Transfer --------
$grpTransfer = New-Object System.Windows.Forms.GroupBox
$grpTransfer.Text = "Übertragung (optional, PSCP / PLINK)"
$grpTransfer.Location = New-Object System.Drawing.Point(10, 340)
$grpTransfer.Size = New-Object System.Drawing.Size(700, 210)
$form.Controls.Add($grpTransfer)

$chkTransfer = New-Object System.Windows.Forms.CheckBox
$chkTransfer.Text = "Backups per PSCP übertragen"
$chkTransfer.Location = "10,25"; $chkTransfer.Size = "260,22"
$grpTransfer.Controls.Add($chkTransfer)

$lblTHost = New-Object System.Windows.Forms.Label; $lblTHost.Text = "Host:"; $lblTHost.Location = "10,55"; $lblTHost.Size = "80,20"
$grpTransfer.Controls.Add($lblTHost)
$txtTHost = New-Object System.Windows.Forms.TextBox; $txtTHost.Location = "95,52"; $txtTHost.Size = "220,22"
$grpTransfer.Controls.Add($txtTHost)

$lblTUser = New-Object System.Windows.Forms.Label; $lblTUser.Text = "User:"; $lblTUser.Location = "325,55"; $lblTUser.Size = "50,20"
$grpTransfer.Controls.Add($lblTUser)
$txtTUser = New-Object System.Windows.Forms.TextBox; $txtTUser.Location = "380,52"; $txtTUser.Size = "160,22"
$grpTransfer.Controls.Add($txtTUser)

$lblTPath = New-Object System.Windows.Forms.Label; $lblTPath.Text = "Remote-Pfad:"; $lblTPath.Location = "10,85"; $lblTPath.Size = "90,20"
$grpTransfer.Controls.Add($lblTPath)
$txtTPath = New-Object System.Windows.Forms.TextBox; $txtTPath.Location = "95,82"; $txtTPath.Size = "540,22"
$grpTransfer.Controls.Add($txtTPath)

$lblTPw = New-Object System.Windows.Forms.Label; $lblTPw.Text = "Transfer-Pass:"; $lblTPw.Location = "10,115"; $lblTPw.Size = "90,20"
$grpTransfer.Controls.Add($lblTPw)
$txtTPw = New-Object System.Windows.Forms.TextBox; $txtTPw.Location = "95,112"; $txtTPw.Size = "220,22"; $txtTPw.UseSystemPasswordChar = $true
$grpTransfer.Controls.Add($txtTPw)
$lblTPwHint = New-Object System.Windows.Forms.Label; $lblTPwHint.Text = "(leer = OneView-Passwort verwenden)"; $lblTPwHint.Location = "325,115"; $lblTPwHint.Size = "320,20"; $lblTPwHint.ForeColor = 'Gray'
$grpTransfer.Controls.Add($lblTPwHint)

$lblPscp = New-Object System.Windows.Forms.Label; $lblPscp.Text = "pscp.exe:"; $lblPscp.Location = "10,145"; $lblPscp.Size = "90,20"
$grpTransfer.Controls.Add($lblPscp)
$txtPscp = New-Object System.Windows.Forms.TextBox; $txtPscp.Location = "95,142"; $txtPscp.Size = "490,22"
$grpTransfer.Controls.Add($txtPscp)
$btnPscp = New-Object System.Windows.Forms.Button; $btnPscp.Text = "..."; $btnPscp.Location = "590,140"; $btnPscp.Size = "60,26"
$grpTransfer.Controls.Add($btnPscp)

$chkRemote = New-Object System.Windows.Forms.CheckBox
$chkRemote.Text = "Remote-Bereinigung >"; $chkRemote.Location = "10,175"; $chkRemote.Size = "160,22"
$grpTransfer.Controls.Add($chkRemote)
$numRemoteDays = New-Object System.Windows.Forms.NumericUpDown; $numRemoteDays.Location = "170,174"; $numRemoteDays.Size = "60,22"; $numRemoteDays.Minimum = 1; $numRemoteDays.Maximum = 3650
$grpTransfer.Controls.Add($numRemoteDays)
$lblRemoteHint = New-Object System.Windows.Forms.Label; $lblRemoteHint.Text = "Tage (via plink.exe):"; $lblRemoteHint.Location = "235,177"; $lblRemoteHint.Size = "130,20"
$grpTransfer.Controls.Add($lblRemoteHint)
$txtPlink = New-Object System.Windows.Forms.TextBox; $txtPlink.Location = "365,174"; $txtPlink.Size = "220,22"
$grpTransfer.Controls.Add($txtPlink)
$btnPlink = New-Object System.Windows.Forms.Button; $btnPlink.Text = "..."; $btnPlink.Location = "590,172"; $btnPlink.Size = "60,26"
$grpTransfer.Controls.Add($btnPlink)

$btnPscp.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "pscp.exe|pscp.exe|Alle Dateien (*.*)|*.*"
        if ($ofd.ShowDialog() -eq 'OK') { $txtPscp.Text = $ofd.FileName }
    })
$btnPlink.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "plink.exe|plink.exe|Alle Dateien (*.*)|*.*"
        if ($ofd.ShowDialog() -eq 'OK') { $txtPlink.Text = $ofd.FileName }
    })

# -------- E-Mail --------
$grpMail = New-Object System.Windows.Forms.GroupBox
$grpMail.Text = "E-Mail (SMTP, optional)"
$grpMail.Location = New-Object System.Drawing.Point(10, 560)
$grpMail.Size = New-Object System.Drawing.Size(700, 220)
$form.Controls.Add($grpMail)

$chkSend = New-Object System.Windows.Forms.CheckBox; $chkSend.Text = "E-Mail senden"; $chkSend.Location = "10,25"; $chkSend.Size = "130,22"
$grpMail.Controls.Add($chkSend)
$chkOnlyErr = New-Object System.Windows.Forms.CheckBox; $chkOnlyErr.Text = "Nur bei Fehlern"; $chkOnlyErr.Location = "150,25"; $chkOnlyErr.Size = "150,22"
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
$txtTo = New-Object System.Windows.Forms.TextBox; $txtTo.Location = "135,112"; $txtTo.Size = "540,22"
$grpMail.Controls.Add($txtTo)
$lblToHint = New-Object System.Windows.Forms.Label; $lblToHint.Text = "(mehrere per ';' oder ',' trennen)"; $lblToHint.Location = "135,135"; $lblToHint.Size = "400,16"; $lblToHint.ForeColor = 'Gray'
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
$txtSubj = New-Object System.Windows.Forms.TextBox; $txtSubj.Location = "440,22"; $txtSubj.Size = "240,22"
$grpMail.Controls.Add($txtSubj)

$btnTestMail = New-Object System.Windows.Forms.Button; $btnTestMail.Text = "Test-Mail senden"; $btnTestMail.Location = "545,178"; $btnTestMail.Size = "135,26"
$grpMail.Controls.Add($btnTestMail)

# -------- Script-Pfade --------
$grpPaths = New-Object System.Windows.Forms.GroupBox
$grpPaths.Text = "Script-Pfade"
$grpPaths.Location = New-Object System.Drawing.Point(10, 790)
$grpPaths.Size = New-Object System.Drawing.Size(700, 90)
$form.Controls.Add($grpPaths)

$lblPwsh = New-Object System.Windows.Forms.Label; $lblPwsh.Text = "PowerShell:"; $lblPwsh.Location = "10,28"; $lblPwsh.Size = "120,20"
$grpPaths.Controls.Add($lblPwsh)
$txtPwsh = New-Object System.Windows.Forms.TextBox; $txtPwsh.Location = "135,25"; $txtPwsh.Size = "440,22"
$grpPaths.Controls.Add($txtPwsh)
$btnPwshBrowse = New-Object System.Windows.Forms.Button; $btnPwshBrowse.Text = "..."; $btnPwshBrowse.Location = "580,23"; $btnPwshBrowse.Size = "95,26"
$grpPaths.Controls.Add($btnPwshBrowse)

$lblRunner = New-Object System.Windows.Forms.Label; $lblRunner.Text = "Runner-Script:"; $lblRunner.Location = "10,58"; $lblRunner.Size = "120,20"
$grpPaths.Controls.Add($lblRunner)
$txtRunner = New-Object System.Windows.Forms.TextBox; $txtRunner.Location = "135,55"; $txtRunner.Size = "440,22"
$grpPaths.Controls.Add($txtRunner)
$btnRunnerBrowse = New-Object System.Windows.Forms.Button; $btnRunnerBrowse.Text = "..."; $btnRunnerBrowse.Location = "580,53"; $btnRunnerBrowse.Size = "95,26"
$grpPaths.Controls.Add($btnRunnerBrowse)

$btnPwshBrowse.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "PowerShell (pwsh.exe;powershell.exe)|pwsh.exe;powershell.exe|Alle Dateien (*.*)|*.*"
        if ($ofd.ShowDialog() -eq 'OK') { $txtPwsh.Text = $ofd.FileName }
    })
$btnRunnerBrowse.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "PowerShell-Scripts (*.ps1)|*.ps1|Alle Dateien (*.*)|*.*"
        if ($ofd.ShowDialog() -eq 'OK') { $txtRunner.Text = $ofd.FileName }
    })

# -------- Geplanter Task --------
$grpTask = New-Object System.Windows.Forms.GroupBox
$grpTask.Text = "Geplanter Task"
$grpTask.Location = New-Object System.Drawing.Point(10, 890)
$grpTask.Size = New-Object System.Drawing.Size(700, 340)
$form.Controls.Add($grpTask)

$lblTimes = New-Object System.Windows.Forms.Label; $lblTimes.Text = "Startzeiten:"; $lblTimes.Location = "10,28"; $lblTimes.Size = "120,20"
$grpTask.Controls.Add($lblTimes)
$lstTimes = New-Object System.Windows.Forms.ListBox; $lstTimes.Location = "135,25"; $lstTimes.Size = "150,70"
$grpTask.Controls.Add($lstTimes)
$dtpTime = New-Object System.Windows.Forms.DateTimePicker; $dtpTime.Format = 'Custom'; $dtpTime.CustomFormat = 'HH:mm'; $dtpTime.ShowUpDown = $true
$dtpTime.Location = "295,25"; $dtpTime.Size = "80,22"; $dtpTime.Value = (Get-Date -Hour 3 -Minute 0 -Second 0)
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

$chkAtBoot = New-Object System.Windows.Forms.CheckBox
$chkAtBoot.Text = "Zusätzlich bei Systemstart ausführen"
$chkAtBoot.Location = "10,138"; $chkAtBoot.Size = "320,22"
$grpTask.Controls.Add($chkAtBoot)

$lblUser2 = New-Object System.Windows.Forms.Label; $lblUser2.Text = "Task-Benutzer:"; $lblUser2.Location = "10,168"; $lblUser2.Size = "120,20"
$grpTask.Controls.Add($lblUser2)
$cmbUser = New-Object System.Windows.Forms.ComboBox; $cmbUser.Location = "135,165"; $cmbUser.Size = "220,22"; $cmbUser.DropDownStyle = 'DropDownList'
[void]$cmbUser.Items.AddRange(@('Aktueller Benutzer (interaktiv)', 'Aktueller Benutzer (S4U, ohne Passwort)', 'SYSTEM', 'NETWORK SERVICE', 'Eigener Benutzer...'))
$cmbUser.SelectedIndex = 0
$grpTask.Controls.Add($cmbUser)
$txtTaskUser = New-Object System.Windows.Forms.TextBox; $txtTaskUser.Location = "365,165"; $txtTaskUser.Size = "200,22"; $txtTaskUser.Enabled = $false
$txtTaskUser.PlaceholderText = "DOMAIN\User"
$grpTask.Controls.Add($txtTaskUser)
$cmbUser.Add_SelectedIndexChanged({ $txtTaskUser.Enabled = ($cmbUser.SelectedItem -eq 'Eigener Benutzer...') })

$lblUserInfo = New-Object System.Windows.Forms.Label
$lblUserInfo.Location = "135,190"; $lblUserInfo.Size = "550,50"
$lblUserInfo.ForeColor = 'DarkBlue'
$grpTask.Controls.Add($lblUserInfo)

$updateUserInfo = {
    switch ([string]$cmbUser.SelectedItem) {
        'Aktueller Benutzer (interaktiv)' {
            $lblUserInfo.Text = "Läuft NUR wenn du interaktiv eingeloggt bist. Kein Passwort nötig.`r`nNach Logout/Reboot wird der Task nicht ausgeführt."
        }
        'Aktueller Benutzer (S4U, ohne Passwort)' {
            $lblUserInfo.Text = "Läuft auch ohne Login, aber OHNE Netzwerk-Zugriff.`r`nFür Backup auf Appliance/SMTP NICHT geeignet."
        }
        'SYSTEM' {
            $lblUserInfo.Text = "Läuft immer, ohne Passwort. Netzwerk nur als Computer-Account (HOSTNAME$)."
        }
        'NETWORK SERVICE' {
            $lblUserInfo.Text = "Läuft immer, ohne Passwort. Eingeschränkte Rechte."
        }
        'Eigener Benutzer...' {
            $lblUserInfo.Text = "EMPFOHLEN: Domain-Account + Passwort. Läuft immer, mit vollem Netzwerk-Zugriff.`r`nBei Passwort-Wechsel hier neu eintragen."
        }
        default { $lblUserInfo.Text = '' }
    }
}
$cmbUser.Add_SelectedIndexChanged($updateUserInfo)
& $updateUserInfo

$lblTaskPw = New-Object System.Windows.Forms.Label; $lblTaskPw.Text = "Task-Passwort:"; $lblTaskPw.Location = "10,250"; $lblTaskPw.Size = "120,20"
$grpTask.Controls.Add($lblTaskPw)
$txtTaskPw = New-Object System.Windows.Forms.TextBox; $txtTaskPw.Location = "135,247"; $txtTaskPw.Size = "220,22"; $txtTaskPw.UseSystemPasswordChar = $true
$grpTask.Controls.Add($txtTaskPw)
$lblTaskPwHint = New-Object System.Windows.Forms.Label; $lblTaskPwHint.Text = "(nur bei 'Eigener Benutzer' benötigt)"; $lblTaskPwHint.Location = "365,250"; $lblTaskPwHint.Size = "270,20"; $lblTaskPwHint.ForeColor = 'Gray'
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
$pnlBtn = New-Object System.Windows.Forms.Panel; $pnlBtn.Location = "10,1240"; $pnlBtn.Size = "700,40"
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
$txtIP660.Text = [string]$config.IPFile660
$txtIP1110.Text = [string]$config.IPFile1110
$txtBaseDir.Text = [string]$config.BackupBaseDir
$numRet.Value = if ($null -ne $config.LocalRetentionDays) { [int]$config.LocalRetentionDays } else { 5 }

$chkTransfer.Checked = [bool]$config.TransferEnabled
$txtTHost.Text = [string]$config.TransferHost
$txtTUser.Text = [string]$config.TransferUser
$txtTPath.Text = [string]$config.TransferRemotePath
$txtPscp.Text = [string]$config.PscpPath
$txtPlink.Text = [string]$config.PlinkPath
$chkRemote.Checked = [bool]$config.RemoteCleanupEnabled
$numRemoteDays.Value = if ($config.RemoteRetentionDays) { [int]$config.RemoteRetentionDays } else { 30 }

$chkSend.Checked = [bool]$config.SendEmail
$chkOnlyErr.Checked = [bool]$config.OnlyOnErrors
$txtSrv.Text = [string]$config.SmtpServer
$numPort.Value = if ($config.SmtpPort) { [int]$config.SmtpPort } else { 25 }
$chkSsl.Checked = [bool]$config.UseSsl
$txtFrom.Text = [string]$config.MailFrom
$txtTo.Text = [string]$config.MailTo
$txtSU.Text = [string]$config.SmtpUser
$txtSubj.Text = if ($config.SubjectPrefix) { $config.SubjectPrefix } else { '[OneView Backup]' }

$lstTimes.Items.Clear()
$loadedTimes = @()
if ($config.ScheduleTimes) { $loadedTimes = @($config.ScheduleTimes) }
if ($loadedTimes.Count -eq 0) { $loadedTimes = @('03:00') }
foreach ($t in $loadedTimes) { [void]$lstTimes.Items.Add([string]$t) }

$loadedDays = @()
if ($config.ScheduleDays) { $loadedDays = @($config.ScheduleDays) }
if ($loadedDays.Count -eq 0) { $loadedDays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday') }
foreach ($k in $dayChecks.Keys) { $dayChecks[$k].Checked = ($loadedDays -contains $k) }

$chkAtBoot.Checked = [bool]$config.ScheduleAtStartup

$txtPwsh.Text = if ($config.PowerShellExe) { [string]$config.PowerShellExe } else { Find-DefaultPwsh }
$txtRunner.Text = if ($config.RunnerScript) { [string]$config.RunnerScript } else { $runnerScript }

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
# Save
# ---------------------------------------------------------------------------
function Save-AllConfig {
    $aesKey = Get-OrCreate-AesKey

    # OneView Credentials + Passphrase + optional Transfer-Password
    $needSaveCred = $false
    $newUser = $txtUser.Text.Trim()
    $existing = Load-CredentialXml

    $encPw = if ($existing) { [string]$existing.EncryptedPassword } else { '' }
    $encPassphrase = if ($existing) { [string]$existing.EncryptedPassphrase } else { '' }
    $encTransfer = if ($existing) { [string]$existing.EncryptedTransferPassword } else { '' }

    if ($txtPass.Text.Length -gt 0) {
        $sec = ConvertTo-SecureString $txtPass.Text -AsPlainText -Force
        $encPw = ConvertFrom-SecureString -SecureString $sec -Key $aesKey
        $txtPass.Text = ''
        $needSaveCred = $true
    }
    if ($txtPassphrase.Text.Length -gt 0) {
        $sec = ConvertTo-SecureString $txtPassphrase.Text -AsPlainText -Force
        $encPassphrase = ConvertFrom-SecureString -SecureString $sec -Key $aesKey
        $txtPassphrase.Text = ''
        $needSaveCred = $true
    }
    if ($txtTPw.Text.Length -gt 0) {
        $sec = ConvertTo-SecureString $txtTPw.Text -AsPlainText -Force
        $encTransfer = ConvertFrom-SecureString -SecureString $sec -Key $aesKey
        $txtTPw.Text = ''
        $needSaveCred = $true
    }
    if ($existing -and $newUser -and ($newUser -ne [string]$existing.Username)) {
        $needSaveCred = $true
    }
    if (-not $existing -and $newUser) {
        $needSaveCred = $true
    }

    if ($newUser -and -not $encPw) {
        [System.Windows.Forms.MessageBox]::Show("Bitte OneView-Passwort setzen.", 'Hinweis', 0, 48) | Out-Null
        return $false
    }
    if ($newUser -and -not $encPassphrase) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Backup-Passphrase setzen.", 'Hinweis', 0, 48) | Out-Null
        return $false
    }

    if ($needSaveCred -and $newUser) {
        $credObj = [PSCustomObject]@{
            Username                  = $newUser
            EncryptedPassword         = $encPw
            EncryptedPassphrase       = $encPassphrase
            EncryptedTransferPassword = $encTransfer
        }
        $credObj | Export-Clixml -Path $credFile
        try {
            $acl = Get-Acl $credFile
            $acl.SetAccessRuleProtection($true, $false)
            foreach ($id in @('NT AUTHORITY\SYSTEM', 'BUILTIN\Administrators', "$env:USERDOMAIN\$env:USERNAME")) {
                try { $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($id, 'FullControl', 'Allow'))) } catch {}
            }
            Set-Acl $credFile $acl
        }
        catch {}
    }

    # SMTP-Passwort
    $smtpEnc = [string]$config.SmtpPasswordEncrypted
    if ($txtSP.Text.Length -gt 0) {
        $sec = ConvertTo-SecureString $txtSP.Text -AsPlainText -Force
        $smtpEnc = ConvertFrom-SecureString -SecureString $sec -Key $aesKey
        $txtSP.Text = ''
    }

    $newCfg = [PSCustomObject]@{
        IPFile660             = $txtIP660.Text.Trim()
        IPFile1110            = $txtIP1110.Text.Trim()
        BackupBaseDir         = $txtBaseDir.Text.Trim()
        LocalRetentionDays    = [int]$numRet.Value
        TransferEnabled       = [bool]$chkTransfer.Checked
        TransferHost          = $txtTHost.Text.Trim()
        TransferUser          = $txtTUser.Text.Trim()
        TransferRemotePath    = $txtTPath.Text.Trim()
        PscpPath              = $txtPscp.Text.Trim()
        PlinkPath             = $txtPlink.Text.Trim()
        RemoteCleanupEnabled  = [bool]$chkRemote.Checked
        RemoteRetentionDays   = [int]$numRemoteDays.Value
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
            try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { param($s, $c, $ch, $e) return $true } } catch {}
            $smtpPort = if ($cfg.SmtpPort) { [int]$cfg.SmtpPort } else { 25 }
            $smtpClient = New-Object Net.Mail.SmtpClient($cfg.SmtpServer, $smtpPort)
            $smtpClient.EnableSsl = [bool]$cfg.UseSsl
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
            $mailMessage.Body = "Dies ist eine Test-E-Mail der OneView Backup Task-Konfiguration.`r`nZeit: $(Get-Date)"
            $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
            $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8
            $smtpClient.Send($mailMessage)
            $mailMessage.Dispose()
            $smtpClient.Dispose()
            [System.Windows.Forms.MessageBox]::Show('Test-Mail gesendet.', 'OK', 0, 64) | Out-Null
        }
        catch {
            $errMsg = $_.Exception.Message
            if ($_.Exception.InnerException) { $errMsg += "`r`n`r`nInner: $($_.Exception.InnerException.Message)" }
            [System.Windows.Forms.MessageBox]::Show("Fehler: $errMsg", 'Test-Mail', 0, 16) | Out-Null
        }
    })

# ---------------------------------------------------------------------------
# Scheduled Task
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
        $pwshExe = $txtPwsh.Text.Trim()
        $runner = $txtRunner.Text.Trim()
        if (-not $pwshExe -or -not (Test-Path $pwshExe)) {
            [System.Windows.Forms.MessageBox]::Show("PowerShell-Executable nicht gefunden: $pwshExe", 'Fehler', 0, 16) | Out-Null
            return
        }
        if (-not $runner -or -not (Test-Path $runner)) {
            [System.Windows.Forms.MessageBox]::Show("Runner-Script nicht gefunden: $runner", 'Fehler', 0, 16) | Out-Null
            return
        }
        $runnerDir = Split-Path $runner -Parent
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

            $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 4)

            if ($p.LogonType -eq 'ServiceAccount') {
                $principal = New-ScheduledTaskPrincipal -UserId $p.User -LogonType ServiceAccount -RunLevel Highest
                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal -Force | Out-Null
            }
            elseif ($p.LogonType -eq 'Interactive') {
                $principal = New-ScheduledTaskPrincipal -UserId $p.User -LogonType Interactive -RunLevel Limited
                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal -Force | Out-Null
            }
            elseif ($p.LogonType -eq 'S4U') {
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
                "Task:         $($t.TaskName)"
                "Status:       $($t.State)"
                "Letzter Lauf: $($info.LastRunTime)"
                "Ergebnis:     $($info.LastTaskResult)"
                "Nächster:     $($info.NextRunTime)"
            ) -join [Environment]::NewLine
            [System.Windows.Forms.MessageBox]::Show($msg, 'Task-Status', 0, 64) | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Task nicht gefunden oder Fehler: $($_.Exception.Message)", 'Status', 0, 48) | Out-Null
        }
    })

[System.Windows.Forms.Application]::Run($form)
