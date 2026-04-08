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
$form.Size = New-Object System.Drawing.Size(800,780)
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
    }
})

# --- Log-Bereich ---
$panelRichLog = New-Object System.Windows.Forms.Panel
$panelRichLog.Location = New-Object System.Drawing.Point(10,210)
$panelRichLog.Size = New-Object System.Drawing.Size(760,200)
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
$detailedListView.Location = New-Object System.Drawing.Point(10,420)
$detailedListView.Size = New-Object System.Drawing.Size(760,250)
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
$buttonStart.Location = New-Object System.Drawing.Point(10,690)
$buttonStart.Size = New-Object System.Drawing.Size(120,24)
$buttonStart.Text = "Start OV Backup"
$form.Controls.Add($buttonStart)

$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Location = New-Object System.Drawing.Point(150,690)
$buttonExit.Size = New-Object System.Drawing.Size(100,24)
$buttonExit.Text = "Exit"
$form.Controls.Add($buttonExit)
$buttonExit.Add_Click({ $form.Close() })

# Hilfsfunktion zum Loggen
function Write-Log {
    param ([string]$message)
    $form.BeginInvoke([action]{
        $logBox.AppendText("$message`r`n")
        $logBox.ScrollToCaret()
    }) | Out-Null
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

    # Appliance-Listen einlesen
    $ipFile660  = $textBoxIPList660.Text
    $ipFile1110 = $textBoxIPList1110.Text

    $appliances660  = @()
    $appliances1110 = @()

    if ([string]::IsNullOrWhiteSpace($ipFile660) -and [string]::IsNullOrWhiteSpace($ipFile1110)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine IP-Datei angeben.", "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $buttonStart.Enabled = $true
        return
    }

    if (-not [string]::IsNullOrWhiteSpace($ipFile660)) {
        if (Test-Path $ipFile660) {
            $appliances660 = @(Get-Content $ipFile660 | Where-Object { $_.Trim() -ne '' })
        } else {
            Write-Log "WARNUNG: OV 6.60 IP-Datei nicht gefunden: $ipFile660"
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ipFile1110)) {
        if (Test-Path $ipFile1110) {
            $appliances1110 = @(Get-Content $ipFile1110 | Where-Object { $_.Trim() -ne '' })
        } else {
            Write-Log "WARNUNG: OV 11.10 IP-Datei nicht gefunden: $ipFile1110"
        }
    }

    if ($appliances660.Count -eq 0 -and $appliances1110.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Appliances in den IP-Dateien gefunden.", "Fehler",
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
            [System.Windows.Forms.Form]$form,
            [System.Windows.Forms.RichTextBox]$logBox,
            [System.Windows.Forms.ListView]$detailedListView,
            [System.Windows.Forms.Button]$buttonStart,
            [System.Windows.Forms.ToolStripStatusLabel]$statusLabel,
            [string]$passphrase,
            [string]$scriptFolder,
            [string]$logFilePath
        )

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

                try {
                    Connect-OVMgmt -Hostname $appliance -Credential $Credential -ErrorAction Stop

                    try {
                        $passphraseSecure = ConvertTo-SecureString $Passphrase -AsPlainText -Force
                        New-OVBackup -Location $currentFolder -Force -Passphrase $passphraseSecure -ErrorAction Stop
                    }
                    finally {
                        Remove-Variable passphraseSecure -ErrorAction SilentlyContinue
                    }
                    Disconnect-OVMgmt

                    [PSCustomObject]@{ Type='UPDATE'; Appliance=$appliance; Status='Erfolgreich'; Detail='Backup erstellt.' }
                }
                catch {
                    $errMsg = $_.Exception.Message
                    [PSCustomObject]@{ Type='UPDATE'; Appliance=$appliance; Status='Fehler'; Detail=$errMsg }
                    ("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Fehler bei Appliance ${appliance}: ${errMsg}") |
                        Out-File -Append -FilePath (Join-Path -Path $BaseBackupDir -ChildPath "Error_Log_${Date}.txt")
                    continue
                }
                finally {
                    Remove-Item -Path "$currentFolder\*.log" -Force -ErrorAction SilentlyContinue
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
                if ($job.State -eq 'Failed' -and $job.Name -ne $null) {
                    $jn = $job.Name
                    $je = $job.ChildJobs[0].JobStateInfo.Reason.Message
                    if ($je) {
                        $form.BeginInvoke([action]{
                            $logBox.AppendText("FEHLER im Job $jn : $je`r`n")
                            $logBox.ScrollToCaret()
                        }) | Out-Null
                    }
                }
                $messages = @(Receive-Job $job -ErrorAction SilentlyContinue)
                foreach ($msg in $messages) {
                    if ($null -eq $msg -or $null -eq $msg.Type) { continue }
                    switch ($msg.Type) {
                        'LOG' {
                            $txt = $msg.Message
                            $form.BeginInvoke([action]{
                                $logBox.AppendText("$txt`r`n")
                                $logBox.ScrollToCaret()
                            }) | Out-Null
                        }
                        'MODULE_FAIL' {
                            $totalAll -= $msg.Count
                            $txt = $msg.Message
                            $form.BeginInvoke([action]{
                                $logBox.AppendText("$txt`r`n")
                                $logBox.ScrollToCaret()
                            }) | Out-Null
                        }
                        'PROGRESS' {
                            $counter++
                            $c = $counter
                            $t = $totalAll
                            $a = $msg.Appliance
                            $v = $msg.VersionLabel
                            $form.Invoke([action]{
                                $logBox.AppendText("Verarbeite Appliance: $a ($c von $t)`r`n")
                                $logBox.ScrollToCaret()
                                $statusLabel.Text = "Bearbeite Appliance $c von $t"
                                $listItem = New-Object System.Windows.Forms.ListViewItem($a)
                                $listItem.Name = $a
                                $listItem.SubItems.Add($v)
                                $listItem.SubItems.Add("Wird verarbeitet")
                                $listItem.SubItems.Add("Start...")
                                $detailedListView.Items.Add($listItem) | Out-Null
                                $listItem.EnsureVisible()
                            })
                        }
                        'UPDATE' {
                            $a = $msg.Appliance
                            $s = $msg.Status
                            $d = $msg.Detail
                            $form.Invoke([action]{
                                $logBox.AppendText("Appliance ${a}: $s – $d`r`n")
                                $logBox.ScrollToCaret()
                                $item = $detailedListView.Items[$a]
                                if ($item -ne $null) {
                                    $item.SubItems[2].Text = $s
                                    $item.SubItems[3].Text = $d
                                    $item.EnsureVisible()
                                }
                            })
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
                        $txt = $msg.Message
                        $form.BeginInvoke([action]{
                            $logBox.AppendText("$txt`r`n")
                            $logBox.ScrollToCaret()
                        }) | Out-Null
                    }
                    'UPDATE' {
                        $a = $msg.Appliance
                        $s = $msg.Status
                        $d = $msg.Detail
                        $form.Invoke([action]{
                            $logBox.AppendText("Appliance ${a}: $s – $d`r`n")
                            $logBox.ScrollToCaret()
                            $item = $detailedListView.Items[$a]
                            if ($item -ne $null) {
                                $item.SubItems[2].Text = $s
                                $item.SubItems[3].Text = $d
                                $item.EnsureVisible()
                            }
                        })
                    }
                }
            }
            $lbl = $job.Name
            $je = $null
            if ($job.State -eq 'Failed') {
                $je = $job.ChildJobs[0].JobStateInfo.Reason.Message
            }
            $form.BeginInvoke([action]{
                if ($je) {
                    $logBox.AppendText("FEHLER im Job $lbl : $je`r`n")
                }
                $logBox.AppendText("--- $lbl Backup-Batch abgeschlossen ---`r`n")
                $logBox.ScrollToCaret()
            }) | Out-Null
            Remove-Job $job -Force
        }

        # -----------------------------------------------------------
        #  Backup-Übertragung per PSCP
        # -----------------------------------------------------------
        $form.BeginInvoke([action]{
            $logBox.AppendText("Backup zum Host sxwotn331n wird durchgeführt...`r`n")
            $logBox.ScrollToCaret()
            $statusLabel.Text = "Backup wird übertragen..."
        }) | Out-Null

        $pscpExe = Join-Path $scriptFolder "tools\pscp.exe"
        if (-not (Test-Path $pscpExe)) {
            $msg = "$(Get-Date) - Warnung: pscp.exe nicht gefunden: $pscpExe. Übertragung wird übersprungen."
            $msg | Out-File -Append -FilePath $logFilePath
            $form.BeginInvoke([action]{
                $logBox.AppendText("$msg`r`n")
                $logBox.ScrollToCaret()
            }) | Out-Null
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
                $form.BeginInvoke([action]{
                    $logBox.AppendText("$msg`r`n")
                    $logBox.ScrollToCaret()
                }) | Out-Null
            }
            catch {
                $errorMsg = "$(Get-Date) - Fehler bei PSCP: $($_.Exception.Message)"
                $errorMsg | Out-File -Append -FilePath $logFilePath
                $form.BeginInvoke([action]{
                    $logBox.AppendText("$errorMsg`r`n")
                    $logBox.ScrollToCaret()
                }) | Out-Null
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
            $form.BeginInvoke([action]{
                $logBox.AppendText("$msg`r`n")
                $logBox.ScrollToCaret()
            }) | Out-Null
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
                $form.BeginInvoke([action]{
                    $logBox.AppendText("$msg`r`n")
                    $logBox.ScrollToCaret()
                }) | Out-Null
            }
            catch {
                $errorMsg = "$(Get-Date) - Fehler bei PLINK: $($_.Exception.Message)"
                $errorMsg | Out-File -Append -FilePath $logFilePath
                $form.BeginInvoke([action]{
                    $logBox.AppendText("$errorMsg`r`n")
                    $logBox.ScrollToCaret()
                }) | Out-Null
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
            $form.BeginInvoke([action]{
                $logBox.AppendText("Lokale Backups älter als 5 Tage wurden gelöscht.`r`n")
                $logBox.ScrollToCaret()
            }) | Out-Null
        }
        catch {
            $form.BeginInvoke([action]{
                $logBox.AppendText("Fehler beim Löschen lokaler Backups: $($_.Exception.Message)`r`n")
                $logBox.ScrollToCaret()
            }) | Out-Null
        }

        # Abschluss
        $form.BeginInvoke([action]{
            $logBox.AppendText("`r`nBackup-Vorgang abgeschlossen.`r`n")
            $logBox.ScrollToCaret()
            $statusLabel.Text = "Fertig"
            $buttonStart.Enabled = $true
        }) | Out-Null
    }

    # Starte Orchestrator-Runspace
    $asyncPS = [powershell]::Create()
    $asyncPS.AddScript($orchestratorBlock) | Out-Null
    $asyncPS.AddArgument($appliances660) | Out-Null
    $asyncPS.AddArgument($appliances1110) | Out-Null
    $asyncPS.AddArgument($credential) | Out-Null
    $asyncPS.AddArgument($baseBackupDir) | Out-Null
    $asyncPS.AddArgument($folderPath) | Out-Null
    $asyncPS.AddArgument($date) | Out-Null
    $asyncPS.AddArgument($form) | Out-Null
    $asyncPS.AddArgument($logBox) | Out-Null
    $asyncPS.AddArgument($detailedListView) | Out-Null
    $asyncPS.AddArgument($buttonStart) | Out-Null
    $asyncPS.AddArgument($statusLabel) | Out-Null
    $asyncPS.AddArgument($passphrase) | Out-Null
    $asyncPS.AddArgument($scriptFolder) | Out-Null
    $asyncPS.AddArgument($logFilePath) | Out-Null
    $null = $asyncPS.BeginInvoke()
})

# Formular anzeigen
$form.ShowDialog()
