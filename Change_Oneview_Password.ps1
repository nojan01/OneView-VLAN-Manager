# ============================================================================
#  HPE OneView Password Change – Kombiniert (OV 6.60 + OV 11.10)
#  Sequentiell: Jede Appliance wird einzeln per Start-Job verarbeitet
#  X-API-Version wird automatisch ermittelt
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
$form.Text = "© 2025 N.J. Airbus D&S - HPE OneView Password Change (OV 6.60 + 11.10)"
$form.Size = New-Object System.Drawing.Size(800,880)
$form.StartPosition = "CenterScreen"

# --- Admin-Credentials (zum Einloggen) ---
$labelAdminUser = New-Object System.Windows.Forms.Label
$labelAdminUser.Location = New-Object System.Drawing.Point(10,20)
$labelAdminUser.Size = New-Object System.Drawing.Size(130,20)
$labelAdminUser.Text = "Admin Login:"
$labelAdminUser.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelAdminUser)

$textBoxAdminUser = New-Object System.Windows.Forms.TextBox
$textBoxAdminUser.Location = New-Object System.Drawing.Point(150,20)
$textBoxAdminUser.Size = New-Object System.Drawing.Size(200,20)
$textBoxAdminUser.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxAdminUser)

$labelAdminPass = New-Object System.Windows.Forms.Label
$labelAdminPass.Location = New-Object System.Drawing.Point(10,50)
$labelAdminPass.Size = New-Object System.Drawing.Size(130,20)
$labelAdminPass.Text = "Admin Passwort:"
$form.Controls.Add($labelAdminPass)

$textBoxAdminPass = New-Object System.Windows.Forms.TextBox
$textBoxAdminPass.Location = New-Object System.Drawing.Point(150,50)
$textBoxAdminPass.Size = New-Object System.Drawing.Size(200,20)
$textBoxAdminPass.UseSystemPasswordChar = $true
$textBoxAdminPass.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxAdminPass)

# --- Ziel-User & neues Passwort ---
$labelTargetUser = New-Object System.Windows.Forms.Label
$labelTargetUser.Location = New-Object System.Drawing.Point(10,90)
$labelTargetUser.Size = New-Object System.Drawing.Size(130,20)
$labelTargetUser.Text = "Ziel-Username:"
$labelTargetUser.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelTargetUser)

$textBoxTargetUser = New-Object System.Windows.Forms.TextBox
$textBoxTargetUser.Location = New-Object System.Drawing.Point(150,90)
$textBoxTargetUser.Size = New-Object System.Drawing.Size(200,20)
$textBoxTargetUser.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxTargetUser)

$labelNewPass = New-Object System.Windows.Forms.Label
$labelNewPass.Location = New-Object System.Drawing.Point(10,120)
$labelNewPass.Size = New-Object System.Drawing.Size(130,20)
$labelNewPass.Text = "Neues Passwort:"
$labelNewPass.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelNewPass)

$textBoxNewPass = New-Object System.Windows.Forms.TextBox
$textBoxNewPass.Location = New-Object System.Drawing.Point(150,120)
$textBoxNewPass.Size = New-Object System.Drawing.Size(200,20)
$textBoxNewPass.UseSystemPasswordChar = $true
$textBoxNewPass.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxNewPass)

$labelConfirmPass = New-Object System.Windows.Forms.Label
$labelConfirmPass.Location = New-Object System.Drawing.Point(10,150)
$labelConfirmPass.Size = New-Object System.Drawing.Size(130,20)
$labelConfirmPass.Text = "Passwort bestätigen:"
$form.Controls.Add($labelConfirmPass)

$textBoxConfirmPass = New-Object System.Windows.Forms.TextBox
$textBoxConfirmPass.Location = New-Object System.Drawing.Point(150,150)
$textBoxConfirmPass.Size = New-Object System.Drawing.Size(200,20)
$textBoxConfirmPass.UseSystemPasswordChar = $true
$textBoxConfirmPass.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxConfirmPass)

# --- IP-Datei: OV 6.60 ---
$labelIPList660 = New-Object System.Windows.Forms.Label
$labelIPList660.Location = New-Object System.Drawing.Point(10,190)
$labelIPList660.Size = New-Object System.Drawing.Size(130,20)
$labelIPList660.Text = "OV 6.60 IP-Datei:"
$labelIPList660.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelIPList660)

$textBoxIPList660 = New-Object System.Windows.Forms.TextBox
$textBoxIPList660.Location = New-Object System.Drawing.Point(150,190)
$textBoxIPList660.Size = New-Object System.Drawing.Size(400,20)
$textBoxIPList660.Text = (Join-Path $scriptFolder "Oneview_660.txt")
$textBoxIPList660.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxIPList660)

$buttonBrowse660 = New-Object System.Windows.Forms.Button
$buttonBrowse660.Location = New-Object System.Drawing.Point(560,190)
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
$labelIPList1110.Location = New-Object System.Drawing.Point(10,220)
$labelIPList1110.Size = New-Object System.Drawing.Size(130,20)
$labelIPList1110.Text = "OV 11.10 IP-Datei:"
$labelIPList1110.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelIPList1110)

$textBoxIPList1110 = New-Object System.Windows.Forms.TextBox
$textBoxIPList1110.Location = New-Object System.Drawing.Point(150,220)
$textBoxIPList1110.Size = New-Object System.Drawing.Size(400,20)
$textBoxIPList1110.Text = (Join-Path $scriptFolder "Oneview.txt")
$textBoxIPList1110.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($textBoxIPList1110)

$buttonBrowse1110 = New-Object System.Windows.Forms.Button
$buttonBrowse1110.Location = New-Object System.Drawing.Point(560,220)
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
$labelApplianceSelect.Location = New-Object System.Drawing.Point(10,253)
$labelApplianceSelect.Size = New-Object System.Drawing.Size(140,20)
$labelApplianceSelect.Text = "Appliance-Auswahl:"
$labelApplianceSelect.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelApplianceSelect)

$buttonSelectAll = New-Object System.Windows.Forms.Button
$buttonSelectAll.Location = New-Object System.Drawing.Point(150,250)
$buttonSelectAll.Size = New-Object System.Drawing.Size(60,23)
$buttonSelectAll.Text = "Alle"
$form.Controls.Add($buttonSelectAll)

$buttonSelectNone = New-Object System.Windows.Forms.Button
$buttonSelectNone.Location = New-Object System.Drawing.Point(220,250)
$buttonSelectNone.Size = New-Object System.Drawing.Size(60,23)
$buttonSelectNone.Text = "Keine"
$form.Controls.Add($buttonSelectNone)

$applianceList = New-Object System.Windows.Forms.CheckedListBox
$applianceList.Location = New-Object System.Drawing.Point(10,277)
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
$panelRichLog.Location = New-Object System.Drawing.Point(10,485)
$panelRichLog.Size = New-Object System.Drawing.Size(760,160)
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
$detailedListView.Location = New-Object System.Drawing.Point(10,655)
$detailedListView.Size = New-Object System.Drawing.Size(760,140)
$detailedListView.View = [System.Windows.Forms.View]::Details
$detailedListView.FullRowSelect = $true
$detailedListView.GridLines = $true
$detailedListView.Scrollable = $true
$detailedListView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$detailedListView.Columns.Add("Appliance",200) | Out-Null
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
$buttonStart.Location = New-Object System.Drawing.Point(10,805)
$buttonStart.Size = New-Object System.Drawing.Size(160,24)
$buttonStart.Text = "Passwort ändern"
$form.Controls.Add($buttonStart)

$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Location = New-Object System.Drawing.Point(190,805)
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

    $adminUser = $textBoxAdminUser.Text
    $adminPass = $textBoxAdminPass.Text
    $targetUser = $textBoxTargetUser.Text
    $newPass = $textBoxNewPass.Text
    $confirmPass = $textBoxConfirmPass.Text

    # Validierung
    if ([string]::IsNullOrEmpty($adminUser) -or [string]::IsNullOrEmpty($adminPass)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin Login und Passwort ausfüllen.", "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $buttonStart.Enabled = $true
        return
    }
    if ([string]::IsNullOrEmpty($targetUser)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Ziel-Username angeben.", "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $buttonStart.Enabled = $true
        return
    }
    if ([string]::IsNullOrEmpty($newPass) -or [string]::IsNullOrEmpty($confirmPass)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte neues Passwort und Bestätigung ausfüllen.", "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $buttonStart.Enabled = $true
        return
    }
    if ($newPass -ne $confirmPass) {
        [System.Windows.Forms.MessageBox]::Show("Die Passwörter stimmen nicht überein.", "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $buttonStart.Enabled = $true
        return
    }

    # Sicherheitsabfrage
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Passwort für User '$targetUser' auf allen ausgewählten Appliances ändern?",
        "Bestätigung",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
        $buttonStart.Enabled = $true
        return
    }

    $secureAdminPass = ConvertTo-SecureString $adminPass -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($adminUser, $secureAdminPass)

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

    $logBox.Clear()
    $detailedListView.Items.Clear()

    $orchestratorBlock = {
        param(
            [string[]]$appliances660,
            [string[]]$appliances1110,
            [System.Management.Automation.PSCredential]$credential,
            [string]$targetUser,
            [string]$newPassword,
            [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$uiQueue
        )

        try {
        $totalAll = $appliances660.Count + $appliances1110.Count
        $counter  = 0

        # -----------------------------------------------------------
        #  Job-ScriptBlock: Läuft als Start-Job in eigenem Prozess
        #  (Eigener Prozess = eigenes Modul, KEIN Assembly-Konflikt)
        #  X-API-Version wird automatisch ermittelt
        # -----------------------------------------------------------
        $changePasswordJob = {
            param(
                [string]$Appliance,
                [string]$ModuleName,
                [string]$VersionLabel,
                [System.Management.Automation.PSCredential]$Credential,
                [string]$TargetUser,
                [string]$NewPassword
            )

            # SSL-Zertifikatsprüfung deaktivieren
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch {}
            $Global:SetLibraryBypassCertificatePolicy = $true

            # Modul laden
            try {
                Import-Module $ModuleName -Force -ErrorAction Stop
            }
            catch {
                [PSCustomObject]@{ Type='ERROR'; Message="Konnte $ModuleName nicht laden: $($_.Exception.Message)" }
                return
            }

            # X-API-Version automatisch ermitteln
            try {
                $apiResponse = Invoke-RestMethod -Uri "https://$Appliance/rest/version" -Method Get -SkipCertificateCheck -ErrorAction Stop
                $apiVersion = $apiResponse.currentVersion
                [PSCustomObject]@{ Type='LOG'; Message="X-API-Version für $Appliance ermittelt: $apiVersion" }
            }
            catch {
                [PSCustomObject]@{ Type='LOG'; Message="WARNUNG: X-API-Version konnte nicht ermittelt werden: $($_.Exception.Message) – Verwende Modul-Standard." }
                $apiVersion = $null
            }

            # Verbinden
            try {
                Connect-OVMgmt -Hostname $Appliance -Credential $Credential -ErrorAction Stop
                [PSCustomObject]@{ Type='LOG'; Message="Verbunden mit $Appliance (OV $VersionLabel)" }
            }
            catch {
                [PSCustomObject]@{ Type='ERROR'; Message="Verbindung zu $Appliance fehlgeschlagen: $($_.Exception.Message)" }
                return
            }

            # Passwort ändern
            try {
                Set-OVUser -UserName $TargetUser -Password $NewPassword -ErrorAction Stop
                [PSCustomObject]@{ Type='SUCCESS'; Message="Passwort für '$TargetUser' auf $Appliance erfolgreich geändert." }
            }
            catch {
                [PSCustomObject]@{ Type='ERROR'; Message="Passwort-Änderung auf $Appliance fehlgeschlagen: $($_.Exception.Message)" }
            }
            finally {
                try { Disconnect-OVMgmt -ErrorAction SilentlyContinue } catch {}
            }
        }

        # -----------------------------------------------------------
        #  Sequentiell alle Appliances verarbeiten
        # -----------------------------------------------------------
        $batches = @()
        if ($appliances660.Count -gt 0) {
            $batches += @{ List = $appliances660; Module = "HPEOneView.660"; Label = "6.60" }
        }
        if ($appliances1110.Count -gt 0) {
            $batches += @{ List = $appliances1110; Module = "HPEOneView.1000"; Label = "11.10" }
        }

        foreach ($batch in $batches) {
            $uiQueue.Enqueue(@{ Type='LOG'; Text="=== Starte Passwort-Änderung für OV $($batch.Label) ($($batch.List.Count) Appliance(s)) ===" })

            foreach ($appliance in $batch.List) {
                $counter++
                $uiQueue.Enqueue(@{ Type='PROGRESS'; Counter=$counter; Total=$totalAll; Appliance=$appliance; Version=$batch.Label })

                # Start-Job für Modul-Isolation
                $job = Start-Job -ScriptBlock $changePasswordJob -ArgumentList @(
                    $appliance, $batch.Module, $batch.Label,
                    $credential, $targetUser, $newPassword
                )

                # Timeout 120s pro Appliance
                $done = $job | Wait-Job -Timeout 120
                if (-not $done) {
                    $job | Stop-Job -PassThru | Remove-Job -Force
                    $uiQueue.Enqueue(@{ Type='UPDATE'; Appliance=$appliance; Status='Fehler'; Detail="Timeout nach 120 Sekunden" })
                    $uiQueue.Enqueue(@{ Type='LOG'; Text="FEHLER: $appliance – Timeout nach 120 Sekunden" })
                    continue
                }

                # Job-Ergebnisse auswerten
                $messages = @(Receive-Job $job -ErrorAction SilentlyContinue)
                Remove-Job $job -Force

                $hasSuccess = $false
                $hasError = $false
                foreach ($msg in $messages) {
                    if ($null -eq $msg -or $null -eq $msg.Type) { continue }
                    switch ($msg.Type) {
                        'LOG' {
                            $uiQueue.Enqueue(@{ Type='LOG'; Text=$msg.Message })
                        }
                        'SUCCESS' {
                            $hasSuccess = $true
                            $uiQueue.Enqueue(@{ Type='LOG'; Text=$msg.Message })
                            $uiQueue.Enqueue(@{ Type='UPDATE'; Appliance=$appliance; Status='Erfolgreich'; Detail="Passwort geändert" })
                        }
                        'ERROR' {
                            $hasError = $true
                            $uiQueue.Enqueue(@{ Type='LOG'; Text="FEHLER: $($msg.Message)" })
                            $uiQueue.Enqueue(@{ Type='UPDATE'; Appliance=$appliance; Status='Fehler'; Detail=$msg.Message })
                        }
                    }
                }

                if (-not $hasError -and -not $hasSuccess) {
                    $uiQueue.Enqueue(@{ Type='UPDATE'; Appliance=$appliance; Status='Fehler'; Detail="Unbekannter Fehler im Job" })
                    $uiQueue.Enqueue(@{ Type='LOG'; Text="FEHLER: $appliance – Unbekannter Fehler im Job" })
                }
            }

            $uiQueue.Enqueue(@{ Type='LOG'; Text="--- OV $($batch.Label) abgeschlossen ---" })
        }

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
                    $logBox.AppendText("`r`nPasswort-Änderung abgeschlossen.`r`n")
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
    $asyncPS.AddArgument($targetUser) | Out-Null
    $asyncPS.AddArgument($newPass) | Out-Null
    $asyncPS.AddArgument($script:uiQueue) | Out-Null
    $null = $asyncPS.BeginInvoke()
})

# Appliance-Liste beim Start automatisch laden
$form.Add_Shown({ Load-Appliances })

# Formular anzeigen
$form.ShowDialog()