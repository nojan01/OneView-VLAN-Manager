# ============================================================================
#  HPE OneView Config-Backup – Kombiniert (OV 6.60 + OV 11.10)
#  Parallel: OV 660 und OV 1000 laufen in eigenen Prozessen (Start-Job)
# ============================================================================

# Skriptordner ermitteln
$scriptFolder = $PSScriptRoot

# =============================
# Konsolenfenster ausblenden
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
    }
}
"@
}
$consolePtr = [Win32.NativeMethods]::GetConsoleWindow()
if ($consolePtr -ne [System.IntPtr]::Zero) {
    [Win32.NativeMethods]::ShowWindow($consolePtr, 0)
}

# =============================
# Erforderliche Assemblies laden
# =============================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -----------------------------
# GUI-Aufbau
# -----------------------------
$form = New-Object System.Windows.Forms.Form
$null = $form.Handle
$form.Text = "© 2025 N.J. Airbus D&S - HPE OneView Config-Backup (OV 6.60 + 11.10)"
$form.Size = New-Object System.Drawing.Size(800,980)
$form.StartPosition = "CenterScreen"

# --- Credentials ---
$labelUsername = New-Object System.Windows.Forms.Label
$labelUsername.Location = New-Object System.Drawing.Point(10,20)
$labelUsername.Size = New-Object System.Drawing.Size(80,20)
$labelUsername.Text = "Login Name:"
$form.Controls.Add($labelUsername)

$textBoxUsername = New-Object System.Windows.Forms.TextBox
$textBoxUsername.Location = New-Object System.Drawing.Point(150,20)
$textBoxUsername.Size = New-Object System.Drawing.Size(200,20)
$textBoxUsername.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxUsername)

$labelPassword = New-Object System.Windows.Forms.Label
$labelPassword.Location = New-Object System.Drawing.Point(10,60)
$labelPassword.Size = New-Object System.Drawing.Size(80,20)
$labelPassword.Text = "Passwort:"
$form.Controls.Add($labelPassword)

$textBoxPassword = New-Object System.Windows.Forms.TextBox
$textBoxPassword.Location = New-Object System.Drawing.Point(150,60)
$textBoxPassword.Size = New-Object System.Drawing.Size(200,20)
$textBoxPassword.UseSystemPasswordChar = $true
$textBoxPassword.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxPassword)

$labelPassphrase = New-Object System.Windows.Forms.Label
$labelPassphrase.Location = New-Object System.Drawing.Point(10,100)
$labelPassphrase.Size = New-Object System.Drawing.Size(130,20)
$labelPassphrase.Text = "Passwort Backupdatei:"
$form.Controls.Add($labelPassphrase)

$textBoxPassphrase = New-Object System.Windows.Forms.TextBox
$textBoxPassphrase.Location = New-Object System.Drawing.Point(150,100)
$textBoxPassphrase.Size = New-Object System.Drawing.Size(200,20)
$textBoxPassphrase.UseSystemPasswordChar = $true
$textBoxPassphrase.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxPassphrase)

# --- IP-Datei: OV 6.60 ---
$labelIPList660 = New-Object System.Windows.Forms.Label
$labelIPList660.Location = New-Object System.Drawing.Point(10,145)
$labelIPList660.Size = New-Object System.Drawing.Size(130,20)
$labelIPList660.Text = "OV 6.60 IP-Datei:"
$labelIPList660.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelIPList660)

$textBoxIPList660 = New-Object System.Windows.Forms.TextBox
$textBoxIPList660.Location = New-Object System.Drawing.Point(150,145)
$textBoxIPList660.Size = New-Object System.Drawing.Size(400,20)
$textBoxIPList660.Text = (Join-Path $scriptFolder "Oneview_660.txt")
$textBoxIPList660.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxIPList660)

$buttonBrowse660 = New-Object System.Windows.Forms.Button
$buttonBrowse660.Location = New-Object System.Drawing.Point(560,145)
$buttonBrowse660.Size = New-Object System.Drawing.Size(75,23)
$buttonBrowse660.Text = "Browse..."
$form.Controls.Add($buttonBrowse660)
$buttonBrowse660.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Textdateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxIPList660.Text = $ofd.FileName
        Load-Appliances
    }
})

# --- IP-Datei: OV 11.10 ---
$labelIPList1110 = New-Object System.Windows.Forms.Label
$labelIPList1110.Location = New-Object System.Drawing.Point(10,175)
$labelIPList1110.Size = New-Object System.Drawing.Size(130,20)
$labelIPList1110.Text = "OV 11.10 IP-Datei:"
$labelIPList1110.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelIPList1110)

$textBoxIPList1110 = New-Object System.Windows.Forms.TextBox
$textBoxIPList1110.Location = New-Object System.Drawing.Point(150,175)
$textBoxIPList1110.Size = New-Object System.Drawing.Size(400,20)
$textBoxIPList1110.Text = (Join-Path $scriptFolder "Oneview.txt")
$textBoxIPList1110.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxIPList1110)

$buttonBrowse1110 = New-Object System.Windows.Forms.Button
$buttonBrowse1110.Location = New-Object System.Drawing.Point(560,175)
$buttonBrowse1110.Size = New-Object System.Drawing.Size(75,23)
$buttonBrowse1110.Text = "Browse..."
$form.Controls.Add($buttonBrowse1110)
$buttonBrowse1110.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Textdateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxIPList1110.Text = $ofd.FileName
        Load-Appliances
    }
})

# --- Appliance-Auswahl ---
$labelApplianceSelect = New-Object System.Windows.Forms.Label
$labelApplianceSelect.Location = New-Object System.Drawing.Point(10,208)
$labelApplianceSelect.Size = New-Object System.Drawing.Size(140,20)
$labelApplianceSelect.Text = "Appliance-Auswahl:"
$labelApplianceSelect.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelApplianceSelect)

$buttonSelectAll = New-Object System.Windows.Forms.Button
$buttonSelectAll.Location = New-Object System.Drawing.Point(150,205)
$buttonSelectAll.Size = New-Object System.Drawing.Size(60,23)
$buttonSelectAll.Text = "Alle"
$form.Controls.Add($buttonSelectAll)

$buttonSelectNone = New-Object System.Windows.Forms.Button
$buttonSelectNone.Location = New-Object System.Drawing.Point(220,205)
$buttonSelectNone.Size = New-Object System.Drawing.Size(60,23)
$buttonSelectNone.Text = "Keine"
$form.Controls.Add($buttonSelectNone)

$applianceList = New-Object System.Windows.Forms.CheckedListBox
$applianceList.Location = New-Object System.Drawing.Point(10,232)
$applianceList.Size = New-Object System.Drawing.Size(760,200)
$applianceList.CheckOnClick = $true
$applianceList.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($applianceList)

# Funktion: Appliances aus IP-Dateien laden
function Load-Appliances {
    $applianceList.Items.Clear()
    $ipFile660 = $textBoxIPList660.Text
    $ipFile1110 = $textBoxIPList1110.Text
    if (-not [string]::IsNullOrWhiteSpace($ipFile660) -and (Test-Path $ipFile660)) {
        $ips = @(Get-Content $ipFile660 | Where-Object { $_.Trim() -ne '' })
        foreach ($ip in $ips) {
            $applianceList.Items.Add("$($ip.Trim())   (OV 6.60)", $true) | Out-Null
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($ipFile1110) -and (Test-Path $ipFile1110)) {
        $ips = @(Get-Content $ipFile1110 | Where-Object { $_.Trim() -ne '' })
        foreach ($ip in $ips) {
            $applianceList.Items.Add("$($ip.Trim())   (OV 11.10)", $true) | Out-Null
        }
    }
}

$buttonSelectAll.Add_Click({
    for ($i = 0; $i -lt $applianceList.Items.Count; $i++) {
        $applianceList.SetItemChecked($i, $true)
    }
})

$buttonSelectNone.Add_Click({
    for ($i = 0; $i -lt $applianceList.Items.Count; $i++) {
        $applianceList.SetItemChecked($i, $false)
    }
})

# --- Log-Bereich ---
$panelRichLog = New-Object System.Windows.Forms.Panel
$panelRichLog.Location = New-Object System.Drawing.Point(10,445)
$panelRichLog.Size = New-Object System.Drawing.Size(760,180)
$panelRichLog.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Dock = "Fill"
$logBox.Multiline = $true
$logBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$panelRichLog.Controls.Add($logBox)
$form.Controls.Add($panelRichLog)

# --- ListView ---
$detailedListView = New-Object System.Windows.Forms.ListView
$detailedListView.Location = New-Object System.Drawing.Point(10,635)
$detailedListView.Size = New-Object System.Drawing.Size(760,210)
$detailedListView.View = [System.Windows.Forms.View]::Details
$detailedListView.FullRowSelect = $true
$detailedListView.GridLines = $true
$detailedListView.Scrollable = $true
$detailedListView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$detailedListView.Columns.Add("Appliance Name",200) | Out-Null
$detailedListView.Columns.Add("OV Version",80) | Out-Null
$detailedListView.Columns.Add("Status",120) | Out-Null
$detailedListView.Columns.Add("Details",360) | Out-Null
$form.Controls.Add($detailedListView)

$detailedListView.Add_Resize({
    $fixedTotal = 200 + 80 + 120 + 360
    if ($detailedListView.ClientSize.Width -ge $fixedTotal) {
        $detailedListView.Columns[0].Width = 200
        $detailedListView.Columns[1].Width = 80
        $detailedListView.Columns[2].Width = 120
        $detailedListView.Columns[3].Width = 360
    }
    else {
        $ratio = $detailedListView.ClientSize.Width / $fixedTotal
        $detailedListView.Columns[0].Width = [Math]::Floor(200 * $ratio)
        $detailedListView.Columns[1].Width = [Math]::Floor(80 * $ratio)
        $detailedListView.Columns[2].Width = [Math]::Floor(120 * $ratio)
        $detailedListView.Columns[3].Width = $detailedListView.ClientSize.Width - $detailedListView.Columns[0].Width - $detailedListView.Columns[1].Width - $detailedListView.Columns[2].Width
    }
})

# --- StatusStrip ---
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.Dock = [System.Windows.Forms.DockStyle]::Bottom
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Bereit..."
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

# --- Buttons ---
$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Location = New-Object System.Drawing.Point(10,855)
$buttonStart.Size = New-Object System.Drawing.Size(120,24)
$buttonStart.Text = "Start OV Backup"
$form.Controls.Add($buttonStart)

$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Location = New-Object System.Drawing.Point(150,855)
$buttonExit.Size = New-Object System.Drawing.Size(100,24)
$buttonExit.Text = "Exit"
$form.Controls.Add($buttonExit)
$buttonExit.Add_Click({ $form.Close() })

# Hilfsfunktion zum Loggen
function Write-Log {
    param ([string]$message)
    $logBox.AppendText("$message`r`n")
    $logBox.ScrollToCaret()
}

# -----------------------------
# Start-Button Event
# -----------------------------
$buttonStart.Add_Click({
    $buttonStart.Enabled = $false

    $username = $textBoxUsername.Text
    $password = $textBoxPassword.Text
    $passphrase = $textBoxPassphrase.Text
    if ([string]::IsNullOrEmpty($username) -or [string]::IsNullOrEmpty($password) -or [string]::IsNullOrEmpty($passphrase)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Benutzername, Passwort und Passphrase ausfüllen.", "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $buttonStart.Enabled = $true
        return
    }

    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($username, $securePassword)

    $date = Get-Date -Format "yyyy-MM-dd"
    $baseBackupDir = Join-Path $scriptFolder "OneView_Backup"
    $logFilePath = Join-Path $baseBackupDir "Backup_Log_${date}.txt"
    $folderPath = Join-Path -Path $baseBackupDir -ChildPath $date
    if (-not (Test-Path -Path $folderPath)) {
        New-Item -ItemType Directory -Path $folderPath | Out-Null
        Write-Log "Ordner '$folderPath' wurde erstellt."
    }
    else {
        Write-Log "Der Ordner '$folderPath' existiert bereits."
    }

    # Ausgewählte Appliances aus CheckedListBox lesen
    $appliances660  = @()
    $appliances1110 = @()
    for ($i = 0; $i -lt $applianceList.Items.Count; $i++) {
        if ($applianceList.GetItemChecked($i)) {
            $itemText = $applianceList.Items[$i].ToString()
            if ($itemText -match '\(OV 6\.60\)$') {
                $appliances660 += ($itemText -replace '\s+\(OV 6\.60\)$','').Trim()
            }
            elseif ($itemText -match '\(OV 11\.10\)$') {
                $appliances1110 += ($itemText -replace '\s+\(OV 11\.10\)$','').Trim()
            }
        }
    }

    if ($appliances660.Count -eq 0 -and $appliances1110.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine Appliance auswählen.", "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $buttonStart.Enabled = $true
        return
    }

    $orchestratorBlock = {
        param(
            [string[]]$appliances660,
            [string[]]$appliances1110,
            [System.Management.Automation.PSCredential]$credential,
            [string]$baseBackupDir,
            [string]$folderPath,
            [string]$date,
            [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$uiQueue,
            [string]$passphrase,
            [string]$scriptFolder,
            [string]$logFilePath
        )

        try {
        $totalAll = $appliances660.Count + $appliances1110.Count
        $counter  = 0

        # -----------------------------------------------------------
        #  Batch-ScriptBlock: Läuft als Start-Job in eigenem Prozess
        #  (Eigener Prozess = eigenes Modul, KEIN Assembly-Konflikt)
        # -----------------------------------------------------------
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
                [string]$LogFilePath
            )

            # Appliance-Liste aus Trennzeichen-String rekonstruieren
            $ApplianceList = @($ApplianceListStr -split '\|')
            if ($ApplianceList.Count -eq 0) { return }

            # SSL-Zertifikatsprüfung deaktivieren (eigener Prozess erbt keine Einstellungen)
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            try {
                [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
            } catch {}
            # HPE OneView Modul-eigene Zertifikatsprüfung deaktivieren
            $Global:SetLibraryBypassCertificatePolicy = $true

            # Modul laden (eigener Prozess = kein Konflikt mit anderem Modul)
            try {
                Import-Module $ModuleName -Force -ErrorAction Stop
                [PSCustomObject]@{ Type='LOG'; Message="=== $ModuleName geladen – Starte Backup für $($ApplianceList.Count) Appliance(s) (OV $VersionLabel) ===" }
            }
            catch {
                [PSCustomObject]@{ Type='MODULE_FAIL'; Message="FEHLER: Konnte $ModuleName nicht laden: $($_.Exception.Message)"; Count=$ApplianceList.Count }
                return
            }

            foreach ($appliance in $ApplianceList) {
                [PSCustomObject]@{ Type='PROGRESS'; Appliance=$appliance; VersionLabel=$VersionLabel }

                $currentFolder = Join-Path -Path $FolderPath -ChildPath $appliance
                if (-not (Test-Path $currentFolder)) {
                    try {
                        New-Item -ItemType Directory -Path $currentFolder -ErrorAction Stop | Out-Null
                    }
                    catch {
                        [PSCustomObject]@{ Type='UPDATE'; Appliance=$appliance; Status='Fehler'; Detail='Ordner konnte nicht erstellt werden.' }
                        continue
                    }
                }
                Set-Location -Path $currentFolder

                [PSCustomObject]@{ Type='LOG'; Message="Verbinde zu Appliance: $appliance" }

                $maxRetries = 2
                $backupSuccess = $false

                for ($attempt = 1; $attempt -le ($maxRetries + 1); $attempt++) {
                    try {
                        # --- Backup als Job mit Timeout (300s / 5 Min) ---
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

                        $backupDone = $backupJob | Wait-Job -Timeout 300
                        if (-not $backupDone) {
                            $backupJob | Stop-Job -PassThru | Remove-Job -Force
                            throw "Timeout nach 300 Sekunden (Appliance antwortet nicht)"
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
                        # Eventuell Fehlermeldungen aus dem Job-Output einsammeln
                        $jobOutput = @(Receive-Job $backupJob -ErrorAction SilentlyContinue)
                        Remove-Job $backupJob -Force

                        [PSCustomObject]@{ Type='UPDATE'; Appliance=$appliance; Status='Erfolgreich'; Detail="Backup erstellt (Versuch $attempt)." }
                        break
                    }
                    catch {
                        $errMsg = $_.Exception.Message

                        if ($attempt -le $maxRetries) {
                            [PSCustomObject]@{ Type='LOG'; Message="WARNUNG: $appliance Versuch $attempt fehlgeschlagen: $errMsg – Retry in 15s..." }
                            Start-Sleep -Seconds 15
                        }
                        else {
                            [PSCustomObject]@{ Type='UPDATE'; Appliance=$appliance; Status='Fehler'; Detail="$errMsg (nach $attempt Versuchen)" }
                            ("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Fehler bei Appliance ${appliance}: ${errMsg} (nach $attempt Versuchen)") |
                                Out-File -Append -FilePath (Join-Path -Path $BaseBackupDir -ChildPath "Error_Log_${Date}.txt")
                        }
                    }
                    finally {
                        Remove-Item -Path "$currentFolder\*.log" -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }

        # -----------------------------------------------------------
        #  Jobs starten (eigene Prozesse = echte Modul-Isolation)
        # -----------------------------------------------------------
        $jobs = @()

        # --- OV 6.60 Job ---
        if ($appliances660.Count -gt 0) {
            $list660Str = $appliances660 -join '|'
            $jobs += Start-Job -Name 'OV 6.60' -ScriptBlock $batchScript -ArgumentList @(
                $list660Str, "HPEOneView.660", "6.60",
                $credential, $folderPath, $baseBackupDir, $date,
                $passphrase, $logFilePath
            )
        }

        # --- OV 11.10 Job ---
        if ($appliances1110.Count -gt 0) {
            $list1110Str = $appliances1110 -join '|'
            $jobs += Start-Job -Name 'OV 11.10' -ScriptBlock $batchScript -ArgumentList @(
                $list1110Str, "HPEOneView.1000", "11.10",
                $credential, $folderPath, $baseBackupDir, $date,
                $passphrase, $logFilePath
            )
        }

        # -----------------------------------------------------------
        #  Jobs pollen und GUI aktualisieren
        # -----------------------------------------------------------
        while ($jobs | Where-Object { $_.State -eq 'Running' }) {
            foreach ($job in $jobs) {
                # Job-Fehler anzeigen (z.B. bei Crash/Exception)
                if ($job.State -eq 'Failed') {
                    $jn = $job.Name
                    $je = $null
                    if ($job.ChildJobs.Count -gt 0 -and $job.ChildJobs[0].JobStateInfo.Reason) {
                        $je = $job.ChildJobs[0].JobStateInfo.Reason.Message
                    }
                    if (-not $je) { $je = "Unbekannter Fehler (Job State: Failed)" }
                    $uiQueue.Enqueue(@{ Type='LOG'; Text="FEHLER im Job ${jn}: $je" })
                }
                $messages = @(Receive-Job $job -ErrorAction SilentlyContinue)
                foreach ($msg in $messages) {
                    if ($null -eq $msg -or $null -eq $msg.Type) { continue }
                    switch ($msg.Type) {
                        'LOG' {
                            $uiQueue.Enqueue(@{ Type='LOG'; Text=$msg.Message })
                        }
                        'MODULE_FAIL' {
                            $totalAll -= $msg.Count
                            $uiQueue.Enqueue(@{ Type='LOG'; Text=$msg.Message })
                        }
                        'PROGRESS' {
                            $counter++
                            $uiQueue.Enqueue(@{ Type='PROGRESS'; Counter=$counter; Total=$totalAll; Appliance=$msg.Appliance; Version=$msg.VersionLabel })
                        }
                        'UPDATE' {
                            $uiQueue.Enqueue(@{ Type='UPDATE'; Appliance=$msg.Appliance; Status=$msg.Status; Detail=$msg.Detail })
                        }
                    }
                }
            }
            Start-Sleep -Milliseconds 250
        }

        # --- Restliche Ausgaben einsammeln ---
        foreach ($job in $jobs) {
            $messages = @(Receive-Job $job -ErrorAction SilentlyContinue)
            foreach ($msg in $messages) {
                if ($null -eq $msg -or $null -eq $msg.Type) { continue }
                switch ($msg.Type) {
                    'LOG' {
                        $uiQueue.Enqueue(@{ Type='LOG'; Text=$msg.Message })
                    }
                    'UPDATE' {
                        $uiQueue.Enqueue(@{ Type='UPDATE'; Appliance=$msg.Appliance; Status=$msg.Status; Detail=$msg.Detail })
                    }
                }
            }
            $lbl = $job.Name
            $je = $null
            if ($job.State -eq 'Failed' -and $job.ChildJobs.Count -gt 0 -and $job.ChildJobs[0].JobStateInfo.Reason) {
                $je = $job.ChildJobs[0].JobStateInfo.Reason.Message
            }
            if ($je) {
                $uiQueue.Enqueue(@{ Type='LOG'; Text="FEHLER im Job ${lbl}: $je" })
            }
            $uiQueue.Enqueue(@{ Type='LOG'; Text="--- $lbl Backup-Batch abgeschlossen ---" })
            Remove-Job $job -Force
        }

        # -----------------------------------------------------------
        #  Backup-Übertragung per PSCP
        # -----------------------------------------------------------
        $uiQueue.Enqueue(@{ Type='LOG'; Text="Backup zum Host sxwotn331n wird durchgeführt..." })
        $uiQueue.Enqueue(@{ Type='STATUS'; Text="Backup wird übertragen..." })

        $pscpExe = Join-Path $scriptFolder "tools\pscp.exe"
        if (-not (Test-Path $pscpExe)) {
            $msg = "$(Get-Date) - Warnung: pscp.exe nicht gefunden: $pscpExe. Übertragung wird übersprungen."
            $msg | Out-File -Append -FilePath $logFilePath
            $uiQueue.Enqueue(@{ Type='LOG'; Text=$msg })
        }
        else {
            $pscpWorkingDir = [System.IO.Path]::GetDirectoryName($pscpExe)
            try {
                $plainPassword = $credential.GetNetworkCredential().Password
                $source = "$baseBackupDir\*"
                $destination = "hpbackup@sxwotn331n:/home/hpbackup/data/OneView_Backup"
                $pscpArgs = @("-r", "-pw", "$plainPassword", $source, $destination)
                $pscpProcess = Start-Process -FilePath $pscpExe -WorkingDirectory $pscpWorkingDir -ArgumentList $pscpArgs -NoNewWindow -PassThru
                if (-not $pscpProcess.WaitForExit(300000)) {
                    $pscpProcess.Kill()
                    throw "PSCP hat nach 5 Minuten nicht geendet und wurde abgebrochen."
                }
                if ($pscpProcess.ExitCode -ne 0) {
                    throw "PSCP-Fehler: ExitCode $($pscpProcess.ExitCode)"
                }
                $msg = "$(Get-Date) - Backup erfolgreich übertragen."
                $msg | Out-File -Append -FilePath $logFilePath
                $uiQueue.Enqueue(@{ Type='LOG'; Text=$msg })
            }
            catch {
                $errorMsg = "$(Get-Date) - Fehler bei PSCP: $($_.Exception.Message)"
                $errorMsg | Out-File -Append -FilePath $logFilePath
                $uiQueue.Enqueue(@{ Type='LOG'; Text=$errorMsg })
            }
            finally {
                Remove-Variable plainPassword -ErrorAction SilentlyContinue
            }
        }

        # -----------------------------------------------------------
        #  Remote-Bereinigung per PLINK
        # -----------------------------------------------------------
        $plinkExe = Join-Path $scriptFolder "tools\plink.exe"
        if (-not (Test-Path $plinkExe)) {
            $msg = "$(Get-Date) - Warnung: plink.exe nicht gefunden: $plinkExe. Remote-Löschung wird übersprungen."
            $msg | Out-File -Append -FilePath $logFilePath
            $uiQueue.Enqueue(@{ Type='LOG'; Text=$msg })
        }
        else {
            $plinkWorkingDir = [System.IO.Path]::GetDirectoryName($plinkExe)
            try {
                $plainPassword = $credential.GetNetworkCredential().Password
                $remoteCmd = 'find /home/hpbackup/data/OneView_Backup -mindepth 1 -depth -mtime +30 -exec rm -rf {} \;'
                $plinkArgs = @("-batch", "-ssh", "-pw", "$plainPassword", "hpbackup@sxwotn331n", $remoteCmd)
                $plinkProcess = Start-Process -FilePath $plinkExe -WorkingDirectory $plinkWorkingDir -ArgumentList $plinkArgs -NoNewWindow -PassThru
                if (-not $plinkProcess.WaitForExit(60000)) {
                    $plinkProcess.Kill()
                    throw "PLINK hat nach 60 Sekunden nicht geendet und wurde abgebrochen."
                }
                if ($plinkProcess.ExitCode -ne 0) {
                    throw "PLINK-Fehler: ExitCode $($plinkProcess.ExitCode)"
                }
                $msg = "$(Get-Date) - Alte Backups erfolgreich gelöscht."
                $msg | Out-File -Append -FilePath $logFilePath
                $uiQueue.Enqueue(@{ Type='LOG'; Text=$msg })
            }
            catch {
                $errorMsg = "$(Get-Date) - Fehler bei PLINK: $($_.Exception.Message)"
                $errorMsg | Out-File -Append -FilePath $logFilePath
                $uiQueue.Enqueue(@{ Type='LOG'; Text=$errorMsg })
            }
            finally {
                Remove-Variable plainPassword -ErrorAction SilentlyContinue
            }
        }

        # -----------------------------------------------------------
        #  Lokale Bereinigung (>5 Tage)
        # -----------------------------------------------------------
        try {
            Get-ChildItem -Path $baseBackupDir -Recurse -Force |
                Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-5) } |
                Remove-Item -Recurse -Force
            $uiQueue.Enqueue(@{ Type='LOG'; Text="Lokale Backups älter als 5 Tage wurden gelöscht." })
        }
        catch {
            $uiQueue.Enqueue(@{ Type='LOG'; Text="Fehler beim Löschen lokaler Backups: $($_.Exception.Message)" })
        }

        # Abschluss
        $uiQueue.Enqueue(@{ Type='FINISHED' })
        }
        catch {
            $errMsg = $_.Exception.Message
            $uiQueue.Enqueue(@{ Type='CRITICAL_ERROR'; Error=$errMsg })
        }
    }

    # ConcurrentQueue für thread-sichere GUI-Kommunikation
    $script:uiQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()

    # GUI-Timer: Verarbeitet Nachrichten aus der Queue auf dem UI-Thread
    $script:guiTimer = New-Object System.Windows.Forms.Timer
    $script:guiTimer.Interval = 200
    $script:guiTimer.Add_Tick({
        $msg = $null
        while ($script:uiQueue.TryDequeue([ref]$msg)) {
            switch ($msg.Type) {
                'LOG' {
                    $logBox.AppendText("$($msg.Text)`r`n")
                    $logBox.ScrollToCaret()
                }
                'STATUS' {
                    $statusLabel.Text = $msg.Text
                }
                'PROGRESS' {
                    $logBox.AppendText("Verarbeite Appliance: $($msg.Appliance) ($($msg.Counter) von $($msg.Total))`r`n")
                    $logBox.ScrollToCaret()
                    $statusLabel.Text = "Bearbeite Appliance $($msg.Counter) von $($msg.Total)"
                    $listItem = New-Object System.Windows.Forms.ListViewItem($msg.Appliance)
                    $listItem.Name = $msg.Appliance
                    $listItem.SubItems.Add($msg.Version)
                    $listItem.SubItems.Add("Wird verarbeitet")
                    $listItem.SubItems.Add("Start...")
                    $detailedListView.Items.Add($listItem) | Out-Null
                    $listItem.EnsureVisible()
                }
                'UPDATE' {
                    $logBox.AppendText("Appliance $($msg.Appliance): $($msg.Status) – $($msg.Detail)`r`n")
                    $logBox.ScrollToCaret()
                    $item = $detailedListView.Items[$msg.Appliance]
                    if ($item -ne $null) {
                        $item.SubItems[2].Text = $msg.Status
                        $item.SubItems[3].Text = $msg.Detail
                        $item.EnsureVisible()
                    }
                }
                'FINISHED' {
                    $logBox.AppendText("`r`nBackup-Vorgang abgeschlossen.`r`n")
                    $logBox.ScrollToCaret()
                    $statusLabel.Text = "Fertig"
                    $buttonStart.Enabled = $true
                    $script:guiTimer.Stop()
                }
                'CRITICAL_ERROR' {
                    $logBox.AppendText("KRITISCHER FEHLER im Orchestrator: $($msg.Error)`r`n")
                    $logBox.ScrollToCaret()
                    $statusLabel.Text = "Fehler"
                    $buttonStart.Enabled = $true
                    $script:guiTimer.Stop()
                }
            }
        }
    })
    $script:guiTimer.Start()

    # Starte Orchestrator-Runspace
    $asyncPS = [powershell]::Create()
    $asyncPS.AddScript($orchestratorBlock) | Out-Null
    $asyncPS.AddArgument($appliances660) | Out-Null
    $asyncPS.AddArgument($appliances1110) | Out-Null
    $asyncPS.AddArgument($credential) | Out-Null
    $asyncPS.AddArgument($baseBackupDir) | Out-Null
    $asyncPS.AddArgument($folderPath) | Out-Null
    $asyncPS.AddArgument($date) | Out-Null
    $asyncPS.AddArgument($script:uiQueue) | Out-Null
    $asyncPS.AddArgument($passphrase) | Out-Null
    $asyncPS.AddArgument($scriptFolder) | Out-Null
    $asyncPS.AddArgument($logFilePath) | Out-Null
    $null = $asyncPS.BeginInvoke()
})

# Appliance-Liste beim Start automatisch laden
$form.Add_Shown({ Load-Appliances })

# Formular anzeigen
$form.ShowDialog()
