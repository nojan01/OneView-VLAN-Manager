<#
.SYNOPSIS
    GUI für HPE OneView Ethernet Network Import/Export.

.DESCRIPTION
    Grafische Oberfläche zum Erstellen und Exportieren von Ethernet Networks
    in HPE OneView. Fragt Anmeldedaten ab und bietet ein Auswahlmenü für
    die OneView Appliance (aus Appliances.txt).

.NOTES
    Autor:   OneView VLAN Projekt
    Datum:   2026-02-06
    Benötigt: PowerShell 7.x (Windows), Modul "ImportExcel"
#>

# ============================================================================
#  PowerShell-Konsolenfenster verstecken
# ============================================================================
Add-Type -Name Win32 -Namespace Native -MemberDefinition @"
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
"@ -ErrorAction SilentlyContinue
$consoleWindow = [Native.Win32]::GetConsoleWindow()
if ($consoleWindow -ne [IntPtr]::Zero) {
    [Native.Win32]::ShowWindow($consoleWindow, 0) | Out-Null   # 0 = SW_HIDE
}

# ============================================================================
#  Prüfungen & Assemblies
# ============================================================================
if (-not $IsWindows) {
    Write-Error "Die GUI benötigt Windows (System.Windows.Forms)."
    return
}

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$appliancesFile = Join-Path $scriptDir "Appliances.txt"

# ============================================================================
#  Appliances aus Textdatei laden
# ============================================================================
function Get-AppliancesFromFile {
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Appliance-Datei nicht gefunden:`n$FilePath`n`nBitte erstellen Sie die Datei mit einer Appliance pro Zeile.",
            "Datei nicht gefunden",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return @()
    }

    $appliances = @()
    Get-Content -Path $FilePath |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith('#') } |
        ForEach-Object {
            $parts = $_ -split '\s*;\s*', 2
            $appliances += [PSCustomObject]@{
                Hostname = $parts[0].Trim()
                Type     = if ($parts.Count -gt 1) { $parts[1].Trim() } else { "" }
            }
        }

    return $appliances
}

# ============================================================================
#  Hauptformular erstellen
# ============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "HPE OneView – Network Manager"
$form.Size = New-Object System.Drawing.Size(820, 900)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(700, 750)
$form.MaximizeBox = $true
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Show
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# ============================================================================
#  GroupBox: Anmeldeinformationen
# ============================================================================
$grpCred = New-Object System.Windows.Forms.GroupBox
$grpCred.Text = "Anmeldeinformationen"
$grpCred.Location = New-Object System.Drawing.Point(15, 20)
$grpCred.Size = New-Object System.Drawing.Size(770, 110)
$grpCred.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($grpCred)

# Benutzername
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Benutzername:"
$lblUser.Location = New-Object System.Drawing.Point(15, 30)
$lblUser.Size = New-Object System.Drawing.Size(110, 23)
$grpCred.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(130, 27)
$txtUser.Size = New-Object System.Drawing.Size(620, 23)
$txtUser.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpCred.Controls.Add($txtUser)

# Kennwort
$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Kennwort:"
$lblPass.Location = New-Object System.Drawing.Point(15, 65)
$lblPass.Size = New-Object System.Drawing.Size(110, 23)
$grpCred.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(130, 62)
$txtPass.Size = New-Object System.Drawing.Size(620, 23)
$txtPass.UseSystemPasswordChar = $true
$txtPass.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpCred.Controls.Add($txtPass)

# Appliances laden (für Backup, Multi-Deploy und Import-Dialog)
$appliances = Get-AppliancesFromFile -FilePath $appliancesFile

# ============================================================================
#  GroupBox: Aktionen
# ============================================================================
$grpActions = New-Object System.Windows.Forms.GroupBox
$grpActions.Text = "Aktionen"
$grpActions.Location = New-Object System.Drawing.Point(15, 150)
$grpActions.Size = New-Object System.Drawing.Size(770, 130)
$grpActions.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($grpActions)

# TableLayoutPanel für gleichmässige Button-Verteilung (1x3)
$tblActions = New-Object System.Windows.Forms.TableLayoutPanel
$tblActions.Location = New-Object System.Drawing.Point(10, 22)
$tblActions.Size = New-Object System.Drawing.Size(748, 96)
$tblActions.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tblActions.ColumnCount = 3
$tblActions.RowCount = 2
$tblActions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
$tblActions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) | Out-Null
$tblActions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.34))) | Out-Null
$tblActions.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
$tblActions.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
$grpActions.Controls.Add($tblActions)

# Button: Import (Erstellen aus Excel)
$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Text = "Netzwerke erstellen (Import)"
$btnImport.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnImport.Margin = New-Object System.Windows.Forms.Padding(0, 0, 3, 0)
$btnImport.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnImport.ForeColor = [System.Drawing.Color]::White
$btnImport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnImport.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblActions.Controls.Add($btnImport, 0, 0)

# Button: VLAN Backup (Multi-Appliance)
$btnBackup = New-Object System.Windows.Forms.Button
$btnBackup.Text = "VLAN Backup (Multi)"
$btnBackup.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnBackup.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$btnBackup.BackColor = [System.Drawing.Color]::FromArgb(180, 120, 0)
$btnBackup.ForeColor = [System.Drawing.Color]::White
$btnBackup.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBackup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblActions.Controls.Add($btnBackup, 1, 0)

# Button: Netzwerk erstellen (Multi-Appliance)
$btnMultiDeploy = New-Object System.Windows.Forms.Button
$btnMultiDeploy.Text = "Netzwerk erstellen (Multi)"
$btnMultiDeploy.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnMultiDeploy.Margin = New-Object System.Windows.Forms.Padding(3, 0, 0, 0)
$btnMultiDeploy.BackColor = [System.Drawing.Color]::FromArgb(140, 0, 180)
$btnMultiDeploy.ForeColor = [System.Drawing.Color]::White
$btnMultiDeploy.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnMultiDeploy.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblActions.Controls.Add($btnMultiDeploy, 2, 0)

# Button: Network Set Import (Erstellen aus Excel)
$btnNSImport = New-Object System.Windows.Forms.Button
$btnNSImport.Text = "Network Sets importieren"
$btnNSImport.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnNSImport.Margin = New-Object System.Windows.Forms.Padding(0, 3, 3, 0)
$btnNSImport.BackColor = [System.Drawing.Color]::FromArgb(0, 128, 128)
$btnNSImport.ForeColor = [System.Drawing.Color]::White
$btnNSImport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnNSImport.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblActions.Controls.Add($btnNSImport, 0, 1)

# Button: Network Set Backup (Multi-Appliance)
$btnNSBackup = New-Object System.Windows.Forms.Button
$btnNSBackup.Text = "Network Set Backup (Multi)"
$btnNSBackup.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnNSBackup.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 0)
$btnNSBackup.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 80)
$btnNSBackup.ForeColor = [System.Drawing.Color]::White
$btnNSBackup.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnNSBackup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblActions.Controls.Add($btnNSBackup, 1, 1)

# ============================================================================
#  Ausgabebereich (Log)
# ============================================================================
$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = "Protokoll"
$grpLog.Location = New-Object System.Drawing.Point(15, 300)
$grpLog.Size = New-Object System.Drawing.Size(770, 530)
$grpLog.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($grpLog)

$rtbLog = New-Object System.Windows.Forms.RichTextBox
$rtbLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$rtbLog.Margin = New-Object System.Windows.Forms.Padding(10)
$rtbLog.ReadOnly = $true
$rtbLog.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$rtbLog.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$rtbLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$rtbLog.WordWrap = $false
$rtbLog.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
$grpLog.Controls.Add($rtbLog)

# ============================================================================
#  Hilfsfunktionen
# ============================================================================

function Write-GUILog {
    param(
        [string]$Message,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::FromArgb(200, 200, 200)
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    $rtbLog.SelectionStart = $rtbLog.TextLength
    $rtbLog.SelectionColor = $Color
    $rtbLog.AppendText("[$timestamp] $Message`n")
    $rtbLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Invoke-SubprocessWithLiveOutput {
    <# Startet pwsh-Subprozess und zeigt Ausgabe zeilenweise live im Protokollfenster. #>
    param([string]$Command)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "pwsh"
    $psi.Arguments = "-NoProfile -NoLogo -Command $Command"
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $proc = [System.Diagnostics.Process]::Start($psi)

    # Stderr asynchron lesen (verhindert Deadlock)
    $stderrTask = $proc.StandardError.ReadToEndAsync()

    # Stdout zeilenweise lesen – GUI bleibt reaktionsfähig
    while ($null -ne ($line = $proc.StandardOutput.ReadLine())) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $color = [System.Drawing.Color]::FromArgb(200, 200, 200)
        if ($line -match "\[ERROR\]")   { $color = [System.Drawing.Color]::FromArgb(255, 80, 80) }
        elseif ($line -match "\[WARN\]")    { $color = [System.Drawing.Color]::FromArgb(255, 200, 60) }
        elseif ($line -match "\[SUCCESS\]") { $color = [System.Drawing.Color]::FromArgb(80, 220, 80) }
        $rtbLog.SelectionStart = $rtbLog.TextLength
        $rtbLog.SelectionColor = $color
        $rtbLog.AppendText("$line`n")
        $rtbLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    $proc.WaitForExit()

    # Stderr-Ausgabe anzeigen
    $stderr = $stderrTask.GetAwaiter().GetResult()
    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        foreach ($errLine in ($stderr -split "`n")) {
            if ([string]::IsNullOrWhiteSpace($errLine)) { continue }
            $rtbLog.SelectionStart = $rtbLog.TextLength
            $rtbLog.SelectionColor = [System.Drawing.Color]::FromArgb(255, 80, 80)
            $rtbLog.AppendText("$errLine`n")
            $rtbLog.ScrollToCaret()
        }
    }

    return $proc.ExitCode
}

function Show-ApplianceSelectionDialog {
    <#
    .SYNOPSIS  Zeigt einen Dialog mit Checkbox-Liste zur Auswahl von Appliances.
    #>
    param(
        [object[]]$Appliances,
        [string]$Title = "Appliances auswählen"
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(550, 450)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Bitte wählen Sie die gewünschten Appliances aus:"
    $lblInfo.Location = New-Object System.Drawing.Point(15, 12)
    $lblInfo.Size = New-Object System.Drawing.Size(500, 20)
    $dlg.Controls.Add($lblInfo)

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(15, 38)
    $clb.Size = New-Object System.Drawing.Size(505, 290)
    $clb.CheckOnClick = $true
    foreach ($a in $Appliances) {
        $displayName = if ($a.Type) { "$($a.Hostname) ($($a.Type))" } else { $a.Hostname }
        $clb.Items.Add($displayName, $false) | Out-Null
    }
    $dlg.Controls.Add($clb)

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = "Alle auswählen"
    $btnAll.Location = New-Object System.Drawing.Point(15, 340)
    $btnAll.Size = New-Object System.Drawing.Size(120, 28)
    $btnAll.Add_Click({
        for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) }
    })
    $dlg.Controls.Add($btnAll)

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = "Keine auswählen"
    $btnNone.Location = New-Object System.Drawing.Point(145, 340)
    $btnNone.Size = New-Object System.Drawing.Size(120, 28)
    $btnNone.Add_Click({
        for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) }
    })
    $dlg.Controls.Add($btnNone)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(370, 340)
    $btnOK.Size = New-Object System.Drawing.Size(55, 28)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.AcceptButton = $btnOK
    $dlg.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Location = New-Object System.Drawing.Point(430, 340)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.CancelButton = $btnCancel
    $dlg.Controls.Add($btnCancel)

    $result = $dlg.ShowDialog($form)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    $selected = @()
    for ($i = 0; $i -lt $clb.Items.Count; $i++) {
        if ($clb.GetItemChecked($i)) {
            $selected += $Appliances[$i].Hostname
        }
    }
    $dlg.Dispose()
    return $selected
}

function Show-SingleApplianceSelectionDialog {
    <#
    .SYNOPSIS  Zeigt einen Dialog zur Auswahl genau einer Appliance.
    #>
    param(
        [object[]]$Appliances,
        [string]$Title = "Appliance auswählen"
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(550, 450)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Bitte wählen Sie eine Appliance aus:"
    $lblInfo.Location = New-Object System.Drawing.Point(15, 12)
    $lblInfo.Size = New-Object System.Drawing.Size(500, 20)
    $dlg.Controls.Add($lblInfo)

    $lb = New-Object System.Windows.Forms.ListBox
    $lb.Location = New-Object System.Drawing.Point(15, 38)
    $lb.Size = New-Object System.Drawing.Size(505, 290)
    $lb.SelectionMode = [System.Windows.Forms.SelectionMode]::One
    foreach ($a in $Appliances) {
        $displayName = if ($a.Type) { "$($a.Hostname) ($($a.Type))" } else { $a.Hostname }
        $lb.Items.Add($displayName) | Out-Null
    }
    if ($lb.Items.Count -gt 0) { $lb.SelectedIndex = 0 }
    $dlg.Controls.Add($lb)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(370, 340)
    $btnOK.Size = New-Object System.Drawing.Size(55, 28)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.AcceptButton = $btnOK
    $dlg.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Location = New-Object System.Drawing.Point(430, 340)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.CancelButton = $btnCancel
    $dlg.Controls.Add($btnCancel)

    $result = $dlg.ShowDialog($form)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    if ($null -eq $lb.SelectedItem) { return $null }

    $selected = $Appliances[$lb.SelectedIndex].Hostname
    $dlg.Dispose()
    return $selected
}

function Show-NetworkParameterDialog {
    <#
    .SYNOPSIS  Zeigt einen Dialog zur Eingabe von Netzwerk-Parametern.
    #>
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Netzwerk-Parameter eingeben"
    $dlg.Size = New-Object System.Drawing.Size(560, 580)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $yPos = 15
    $lblWidth = 170
    $ctrlX = 185
    $ctrlWidth = 340
    $rowHeight = 32

    # NetworkName
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Network Name:"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $txtNetName = New-Object System.Windows.Forms.TextBox
    $txtNetName.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $txtNetName.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    $dlg.Controls.Add($txtNetName)
    $yPos += $rowHeight

    # VlanId
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "VLAN ID:"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $nudVlanId = New-Object System.Windows.Forms.NumericUpDown
    $nudVlanId.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $nudVlanId.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    $nudVlanId.Minimum = 1; $nudVlanId.Maximum = 4094; $nudVlanId.Value = 100
    $dlg.Controls.Add($nudVlanId)
    $yPos += $rowHeight

    # EthernetNetworkType
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Ethernet Network Type:"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $cmbEthType = New-Object System.Windows.Forms.ComboBox
    $cmbEthType.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $cmbEthType.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    $cmbEthType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @("Tagged", "Untagged", "Tunnel") | ForEach-Object { $cmbEthType.Items.Add($_) | Out-Null }
    $cmbEthType.SelectedIndex = 0
    $dlg.Controls.Add($cmbEthType)
    $yPos += $rowHeight

    # Purpose
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Purpose:"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $cmbPurpose = New-Object System.Windows.Forms.ComboBox
    $cmbPurpose.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $cmbPurpose.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    $cmbPurpose.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @("General", "Management", "VMMigration", "FaultTolerance", "ISCSI") | ForEach-Object { $cmbPurpose.Items.Add($_) | Out-Null }
    $cmbPurpose.SelectedIndex = 0
    $dlg.Controls.Add($cmbPurpose)
    $yPos += $rowHeight

    # SmartLink
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "SmartLink:"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $chkSmartLink = New-Object System.Windows.Forms.CheckBox
    $chkSmartLink.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $chkSmartLink.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    $chkSmartLink.Checked = $true
    $dlg.Controls.Add($chkSmartLink)
    $yPos += $rowHeight

    # PrivateNetwork
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Private Network:"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $chkPrivate = New-Object System.Windows.Forms.CheckBox
    $chkPrivate.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $chkPrivate.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    $chkPrivate.Checked = $false
    $dlg.Controls.Add($chkPrivate)
    $yPos += $rowHeight

    # PreferredBandwidthGb
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Preferred Bandwidth (Gb):"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $nudBwPref = New-Object System.Windows.Forms.NumericUpDown
    $nudBwPref.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $nudBwPref.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    $nudBwPref.Minimum = 0.1; $nudBwPref.Maximum = 800; $nudBwPref.DecimalPlaces = 1; $nudBwPref.Increment = 0.5; $nudBwPref.Value = 2.5
    $dlg.Controls.Add($nudBwPref)
    $yPos += $rowHeight

    # MaximumBandwidthGb
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Maximum Bandwidth (Gb):"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $nudBwMax = New-Object System.Windows.Forms.NumericUpDown
    $nudBwMax.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $nudBwMax.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    $nudBwMax.Minimum = 0.1; $nudBwMax.Maximum = 800; $nudBwMax.DecimalPlaces = 1; $nudBwMax.Increment = 0.5; $nudBwMax.Value = 50
    $dlg.Controls.Add($nudBwMax)
    $yPos += $rowHeight

    # Scope
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Scope (optional):"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $txtScope = New-Object System.Windows.Forms.TextBox
    $txtScope.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $txtScope.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    $dlg.Controls.Add($txtScope)
    $yPos += $rowHeight

    # Hinweis: Network Set Zuweisung erfolgt im nächsten Schritt pro Appliance
    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = "Hinweis: Network Set Zuweisung erfolgt im nächsten Schritt pro Appliance."
    $lblHint.Location = New-Object System.Drawing.Point(15, ($yPos + 3))
    $lblHint.Size = New-Object System.Drawing.Size(520, 20)
    $lblHint.ForeColor = [System.Drawing.Color]::Gray
    $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $dlg.Controls.Add($lblHint)
    $yPos += $rowHeight + 15

    # Buttons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(300, $yPos)
    $btnOK.Size = New-Object System.Drawing.Size(80, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.AcceptButton = $btnOK
    $dlg.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Location = New-Object System.Drawing.Point(385, $yPos)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.CancelButton = $btnCancel
    $dlg.Controls.Add($btnCancel)

    $result = $dlg.ShowDialog($form)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        $dlg.Dispose()
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($txtNetName.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte einen Network Name eingeben.",
            "Pflichtfeld fehlt",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        $dlg.Dispose()
        return $null
    }

    $params = @{
        NetworkName          = $txtNetName.Text.Trim()
        VlanId               = [int]$nudVlanId.Value
        EthernetNetworkType  = $cmbEthType.SelectedItem.ToString()
        Purpose              = $cmbPurpose.SelectedItem.ToString()
        SmartLink            = $chkSmartLink.Checked
        PrivateNetwork       = $chkPrivate.Checked
        PreferredBandwidthGb = [double]$nudBwPref.Value
        MaximumBandwidthGb   = [double]$nudBwMax.Value
        Scope                = $txtScope.Text.Trim()
    }
    $dlg.Dispose()
    return $params
}

# ============================================================================
#  OneView API – Inline-Hilfsfunktionen (für Network Set Abfrage)
# ============================================================================

function Connect-OneViewAPIInline {
    param(
        [Parameter(Mandatory)][string]$Hostname,
        [Parameter(Mandatory)][string]$Username,
        [Parameter(Mandatory)][string]$Password,
        [int]$ApiVersion = 8000
    )

    $baseUri  = "https://$Hostname"
    $loginUri = "$baseUri/rest/login-sessions"

    $body = @{
        userName        = $Username
        password        = $Password
        authLoginDomain = "Local"
    } | ConvertTo-Json

    $headers = @{
        "Content-Type"  = "application/json"
        "X-API-Version" = $ApiVersion
    }

    $response = Invoke-RestMethod -Uri $loginUri -Method Post -Headers $headers -Body $body -SkipCertificateCheck
    $sessionId = $response.sessionID
    if ([string]::IsNullOrEmpty($sessionId)) {
        throw "Keine sessionID erhalten von $Hostname"
    }

    return @{
        BaseUri   = $baseUri
        SessionId = $sessionId
        Headers   = @{
            "Content-Type"  = "application/json"
            "X-API-Version" = $ApiVersion
            "Auth"          = $sessionId
        }
    }
}

function Disconnect-OneViewAPIInline {
    param([Parameter(Mandatory)][hashtable]$Session)
    try {
        Invoke-RestMethod -Uri "$($Session.BaseUri)/rest/login-sessions" `
            -Method Delete -Headers $Session.Headers -SkipCertificateCheck | Out-Null
    } catch { }
}

function Get-NetworkSetsInline {
    param([Parameter(Mandatory)][hashtable]$Session)

    $allMembers = [System.Collections.Generic.List[object]]::new()
    $start = 0
    $pageSize = 200

    do {
        $uri = "$($Session.BaseUri)/rest/network-sets?start=$start&count=$pageSize"
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck
        if ($response.members) {
            $allMembers.AddRange([object[]]$response.members)
        }
        $total = $response.total
        $start += $response.members.Count
    } while ($allMembers.Count -lt $total)

    return $allMembers
}

function Get-ApiVersionInline {
    <#
    .SYNOPSIS  Ermittelt die aktuelle API-Version via GET /rest/version (ohne Auth)
    #>
    param(
        [Parameter(Mandatory)][string]$Hostname,
        [int]$FallbackVersion = 8000
    )

    $uri = "https://$Hostname/rest/version"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Get -SkipCertificateCheck
        return [int]$response.currentVersion
    }
    catch {
        return $FallbackVersion
    }
}

function Show-NetworkSetMappingDialog {
    <#
    .SYNOPSIS  TreeView-Dialog zur Zuweisung von Network Sets pro Appliance.
    .DESCRIPTION
        Zeigt eine Baumansicht: Appliance → Network Sets (mit Checkboxen).
        Der Benutzer kann pro Appliance die Network Sets ankreuzen,
        denen das neue VLAN zugewiesen werden soll.
    .RETURNS
        Hashtable: ApplianceHostname → @("NetworkSet1", "NetworkSet2", ...)
        oder $null bei Abbruch.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$ApplianceNetworkSets  # Hostname → @(Name, Name, ...)
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Network Set Zuweisung pro Appliance"
    $dlg.Size = New-Object System.Drawing.Size(650, 550)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dlg.MinimumSize = New-Object System.Drawing.Size(550, 420)
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Wählen Sie pro Appliance die Network Sets, denen das VLAN zugewiesen werden soll:"
    $lblInfo.Location = New-Object System.Drawing.Point(15, 12)
    $lblInfo.Size = New-Object System.Drawing.Size(600, 20)
    $lblInfo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $dlg.Controls.Add($lblInfo)

    $tv = New-Object System.Windows.Forms.TreeView
    $tv.Location = New-Object System.Drawing.Point(15, 38)
    $tv.Size = New-Object System.Drawing.Size(605, 400)
    $tv.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $tv.CheckBoxes = $true
    $tv.FullRowSelect = $true
    $tv.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.Controls.Add($tv)

    # TreeView befüllen
    foreach ($hostname in ($ApplianceNetworkSets.Keys | Sort-Object)) {
        $parentNode = $tv.Nodes.Add($hostname, $hostname)
        $parentNode.NodeFont = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
        $sets = $ApplianceNetworkSets[$hostname]
        foreach ($setName in ($sets | Sort-Object)) {
            $parentNode.Nodes.Add($setName, $setName) | Out-Null
        }
        $parentNode.Expand()
    }

    # Parent-Check → alle Kinder ein/ausschalten
    $tv.Add_AfterCheck({
        param($sender, $e)
        $node = $e.Node
        if ($node.Nodes.Count -gt 0) {
            foreach ($child in $node.Nodes) {
                $child.Checked = $node.Checked
            }
        }
    })

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Size = New-Object System.Drawing.Size(80, 30)
    $btnOK.Location = New-Object System.Drawing.Point(455, 455)
    $btnOK.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.AcceptButton = $btnOK
    $dlg.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(540, 455)
    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.CancelButton = $btnCancel
    $dlg.Controls.Add($btnCancel)

    $result = $dlg.ShowDialog($form)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        $dlg.Dispose()
        return $null
    }

    # Ergebnis auslesen: pro Appliance die angehakten Network Sets
    $mapping = @{}
    foreach ($parentNode in $tv.Nodes) {
        $selectedSets = @()
        foreach ($child in $parentNode.Nodes) {
            if ($child.Checked) {
                $selectedSets += $child.Text
            }
        }
        $mapping[$parentNode.Text] = $selectedSets
    }
    $dlg.Dispose()
    return $mapping
}

function Test-InputValid {
    if ([string]::IsNullOrWhiteSpace($txtUser.Text) -or [string]::IsNullOrWhiteSpace($txtPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte Benutzername und Kennwort eingeben.",
            "Fehlende Anmeldeinformationen",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $false
    }
    return $true
}

function New-TempConfig {
    <#
    .SYNOPSIS
        Erstellt eine temporäre config.json mit der ausgewählten Appliance.
    #>
    param(
        [string]$Hostname,
        [string]$ExcelPath,
        [string]$TempDir
    )

    $configObj = @{
        OneViewAppliances = @(
            @{
                Name        = $Hostname
                Hostname    = $Hostname
                Description = "Ausgewählt über GUI"
            }
        )
        ApiVersion    = 8000
        ExcelFilePath = $ExcelPath
        ExcelSheetName = "VLANs"
        DefaultSettings = @{
            Purpose              = "General"
            SmartLink            = $true
            PrivateNetwork       = $false
            EthernetNetworkType  = "Tagged"
            PreferredBandwidthGb = 2.5
            MaximumBandwidthGb   = 50
        }
    }

    # Originale config.json lesen falls vorhanden (für ApiVersion & Defaults)
    $origConfig = Join-Path $scriptDir "config.json"
    if (Test-Path $origConfig) {
        try {
            $orig = Get-Content -Path $origConfig -Raw | ConvertFrom-Json
            $configObj.ApiVersion = $orig.ApiVersion
            if ($orig.DefaultSettings) {
                $configObj.DefaultSettings = @{
                    Purpose              = $orig.DefaultSettings.Purpose
                    SmartLink            = $orig.DefaultSettings.SmartLink
                    PrivateNetwork       = $orig.DefaultSettings.PrivateNetwork
                    EthernetNetworkType  = $orig.DefaultSettings.EthernetNetworkType
                    PreferredBandwidthGb = $orig.DefaultSettings.PreferredBandwidthGb
                    MaximumBandwidthGb   = $orig.DefaultSettings.MaximumBandwidthGb
                }
            }
            if ($orig.ExcelSheetName) {
                $configObj.ExcelSheetName = $orig.ExcelSheetName
            }
        }
        catch { }
    }

    $tempConfig = Join-Path $TempDir "config_temp.json"
    $configObj | ConvertTo-Json -Depth 5 | Set-Content -Path $tempConfig -Encoding UTF8
    return $tempConfig
}

function New-TempConfigNS {
    <#
    .SYNOPSIS
        Erstellt eine temporäre config.json für Network Set Operationen.
    #>
    param(
        [string]$Hostname,
        [string]$ExcelPath,
        [string]$TempDir
    )

    $configObj = @{
        OneViewAppliances = @(
            @{
                Name        = $Hostname
                Hostname    = $Hostname
                Description = "Ausgewählt über GUI"
            }
        )
        ApiVersion                = 8000
        NetworkSetExcelFilePath   = $ExcelPath
        NetworkSetExcelSheetName  = "NetworkSets"
        NetworkSetDefaultSettings = @{
            PreferredBandwidthGb = 2.5
            MaximumBandwidthGb   = 20
        }
    }

    # Originale config.json lesen falls vorhanden
    $origConfig = Join-Path $scriptDir "config.json"
    if (Test-Path $origConfig) {
        try {
            $orig = Get-Content -Path $origConfig -Raw | ConvertFrom-Json
            $configObj.ApiVersion = $orig.ApiVersion
            if ($orig.NetworkSetExcelSheetName) {
                $configObj.NetworkSetExcelSheetName = $orig.NetworkSetExcelSheetName
            }
            if ($orig.NetworkSetDefaultSettings) {
                $configObj.NetworkSetDefaultSettings = @{
                    PreferredBandwidthGb = $orig.NetworkSetDefaultSettings.PreferredBandwidthGb
                    MaximumBandwidthGb   = $orig.NetworkSetDefaultSettings.MaximumBandwidthGb
                }
            }
        }
        catch { }
    }

    $tempConfig = Join-Path $TempDir "config_ns_temp.json"
    $configObj | ConvertTo-Json -Depth 5 | Set-Content -Path $tempConfig -Encoding UTF8
    return $tempConfig
}

# ============================================================================
#  Button-Events
# ============================================================================

$btnImport.Add_Click({
    if (-not (Test-InputValid)) { return }

    # Schritt 1: Appliance auswählen (Einzelauswahl)
    if (-not $appliances -or $appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Keine Appliances in der Appliances.txt gefunden.",
            "Keine Appliances",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $hostname = Show-SingleApplianceSelectionDialog -Appliances $appliances -Title "Import – Appliance auswählen"
    if (-not $hostname) { return }

    # Schritt 2: Excel-Datei auswählen
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Excel-Dateien (*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*"
    $ofd.InitialDirectory = $scriptDir
    $ofd.Title = "Excel-Datei für Import auswählen"
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $excelPath = $ofd.FileName

    # Schritt 3: Bestätigung
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Ethernet Networks aus der Excel-Datei auf`n$hostname`nerstellen?`n`nDatei: $excelPath",
        "Import bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnImport.Enabled = $false
    $btnNSImport.Enabled = $false
    $btnNSBackup.Enabled = $false
    $rtbLog.Clear()

    try {
        Write-GUILog "Starte Import auf $hostname ..." -Color ([System.Drawing.Color]::Cyan)

        # Temporäre Config erstellen
        $tempConfig = New-TempConfig -Hostname $hostname `
            -ExcelPath $excelPath `
            -TempDir $scriptDir

        # Credentials als Umgebungsvariablen übergeben (sicher im Prozess)
        $env:OV_USERNAME = $txtUser.Text
        $env:OV_PASSWORD = $txtPass.Text

        # Script aufrufen (Konsolenfenster versteckt)
        $importScript = Join-Path $scriptDir "Create-EthernetNetworks.ps1"
        $psCommand = @"
`$env:OV_USERNAME = '$($txtUser.Text)'
`$env:OV_PASSWORD = '$($txtPass.Text -replace "'", "''")'
`$secPass = ConvertTo-SecureString `$env:OV_PASSWORD -AsPlainText -Force
`$global:guiCredential = New-Object System.Management.Automation.PSCredential(`$env:OV_USERNAME, `$secPass)
function Get-Credential { param([string]`$Message) return `$global:guiCredential }
function Read-Host { param([string]`$Prompt) return 'J' }
& '$importScript' -ConfigPath '$tempConfig'
"@
        Invoke-SubprocessWithLiveOutput -Command $psCommand | Out-Null

        Write-GUILog "Import abgeschlossen." -Color ([System.Drawing.Color]::Cyan)

        # Temp-Config aufräumen
        Remove-Item -Path $tempConfig -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
    }
    finally {
        $env:OV_USERNAME = $null
        $env:OV_PASSWORD = $null
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnImport.Enabled = $true
        $btnNSImport.Enabled = $true
        $btnNSBackup.Enabled = $true
    }
})

# ============================================================================
#  Button-Event: VLAN Backup (Multi-Appliance)
# ============================================================================
$btnBackup.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtUser.Text) -or [string]::IsNullOrWhiteSpace($txtPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte Benutzername und Kennwort eingeben.",
            "Fehlende Anmeldeinformationen",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if (-not $appliances -or $appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Keine Appliances in der Appliances.txt gefunden.",
            "Keine Appliances",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $selectedAppliances = Show-ApplianceSelectionDialog -Appliances $appliances -Title "VLAN Backup – Appliances auswählen"
    if (-not $selectedAppliances -or $selectedAppliances.Count -eq 0) { return }

    $applianceList = $selectedAppliances -join "`n"
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "VLAN Backup von $($selectedAppliances.Count) Appliance(s) erstellen?`n`n$applianceList",
        "Backup bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnImport.Enabled = $false
    $btnBackup.Enabled = $false
    $btnMultiDeploy.Enabled = $false
    $btnNSImport.Enabled = $false
    $btnNSBackup.Enabled = $false
    $rtbLog.Clear()

    try {
        $backupDir = Join-Path $scriptDir "Backups" (Get-Date -Format "yyyyMMdd_HHmmss")
        New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
        Write-GUILog "Backup-Verzeichnis: $backupDir" -Color ([System.Drawing.Color]::Cyan)

        $logsDir = Join-Path $scriptDir "Logs"
        if (-not (Test-Path $logsDir)) { New-Item -Path $logsDir -ItemType Directory -Force | Out-Null }
        $sharedLogPath = Join-Path $logsDir ("VLAN_Export_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

        $successCount = 0
        $errorCount = 0

        foreach ($applianceHost in $selectedAppliances) {
            Write-GUILog "Starte Backup von $applianceHost ..." -Color ([System.Drawing.Color]::Cyan)

            $safeName = $applianceHost -replace '[\\/:*?\"<>|\.]', '_'
            $exportPath = Join-Path $backupDir ("${safeName}.xlsx")

            $tempConfig = New-TempConfig -Hostname $applianceHost -ExcelPath "" -TempDir $scriptDir

            try {
                $exportScript = Join-Path $scriptDir "Export-EthernetNetworks.ps1"
                $psCommand = @"
`$env:OV_USERNAME = '$($txtUser.Text)'
`$env:OV_PASSWORD = '$($txtPass.Text -replace "'", "''")'
`$secPass = ConvertTo-SecureString `$env:OV_PASSWORD -AsPlainText -Force
`$global:guiCredential = New-Object System.Management.Automation.PSCredential(`$env:OV_USERNAME, `$secPass)
function Get-Credential { param([string]`$Message) return `$global:guiCredential }
& '$exportScript' -ConfigPath '$tempConfig' -OutputPath '$exportPath' -LogPath '$sharedLogPath'
"@
                $exitCode = Invoke-SubprocessWithLiveOutput -Command $psCommand

                if ($exitCode -eq 0 -and (Test-Path $exportPath)) {
                    Write-GUILog "Backup erfolgreich: $applianceHost" -Color ([System.Drawing.Color]::FromArgb(80, 220, 80))
                    $successCount++
                } else {
                    Write-GUILog "Backup fehlgeschlagen: $applianceHost" -Color ([System.Drawing.Color]::FromArgb(255, 80, 80))
                    $errorCount++
                }
            }
            catch {
                Write-GUILog "Fehler bei $applianceHost : $_" -Color ([System.Drawing.Color]::Red)
                $errorCount++
            }
            finally {
                Remove-Item -Path $tempConfig -Force -ErrorAction SilentlyContinue
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        Write-GUILog "" -Color ([System.Drawing.Color]::Cyan)
        Write-GUILog "Backup abgeschlossen: $successCount erfolgreich, $errorCount fehlgeschlagen" -Color ([System.Drawing.Color]::Cyan)
        Write-GUILog "Verzeichnis: $backupDir" -Color ([System.Drawing.Color]::Cyan)

        [System.Windows.Forms.MessageBox]::Show(
            "Backup abgeschlossen!`n`nErfolgreich: $successCount`nFehlgeschlagen: $errorCount`n`nVerzeichnis: $backupDir",
            "Backup Ergebnis",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnImport.Enabled = $true
        $btnBackup.Enabled = $true
        $btnMultiDeploy.Enabled = $true
        $btnNSImport.Enabled = $true
        $btnNSBackup.Enabled = $true
    }
})

# ============================================================================
#  Button-Event: Netzwerk erstellen (Multi-Appliance)
# ============================================================================
$btnMultiDeploy.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtUser.Text) -or [string]::IsNullOrWhiteSpace($txtPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte Benutzername und Kennwort eingeben.",
            "Fehlende Anmeldeinformationen",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if (-not $appliances -or $appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Keine Appliances in der Appliances.txt gefunden.",
            "Keine Appliances",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    # Schritt 1: Netzwerk-Parameter abfragen
    $netParams = Show-NetworkParameterDialog
    if (-not $netParams) { return }

    # Schritt 2: Appliances auswählen
    $selectedAppliances = Show-ApplianceSelectionDialog -Appliances $appliances -Title "Netzwerk erstellen – Ziel-Appliances auswählen"
    if (-not $selectedAppliances -or $selectedAppliances.Count -eq 0) { return }

    # Schritt 3: Network Sets von jeder Appliance abrufen
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $rtbLog.Clear()
    Write-GUILog "Lese Network Sets von den ausgewählten Appliances..." -Color ([System.Drawing.Color]::Cyan)
    [System.Windows.Forms.Application]::DoEvents()

    $applianceNetworkSets = @{}
    $fetchErrors = @()

    foreach ($applianceHost in $selectedAppliances) {
        Write-GUILog "  Verbinde zu $applianceHost ..." -Color ([System.Drawing.Color]::FromArgb(200, 200, 200))
        [System.Windows.Forms.Application]::DoEvents()

        $session = $null
        try {
            # API-Version automatisch pro Appliance erkennen
            $apiVersion = Get-ApiVersionInline -Hostname $applianceHost
            Write-GUILog "  API-Version: $apiVersion" -Color ([System.Drawing.Color]::FromArgb(200, 200, 200))

            $session = Connect-OneViewAPIInline -Hostname $applianceHost `
                -Username $txtUser.Text -Password $txtPass.Text -ApiVersion $apiVersion

            $networkSets = Get-NetworkSetsInline -Session $session
            $setNames = @($networkSets | ForEach-Object { $_.name } | Sort-Object)
            $applianceNetworkSets[$applianceHost] = $setNames

            Write-GUILog "  $applianceHost : $($setNames.Count) Network Sets gefunden" -Color ([System.Drawing.Color]::FromArgb(80, 220, 80))
        }
        catch {
            Write-GUILog "  Fehler bei $applianceHost : $_" -Color ([System.Drawing.Color]::FromArgb(255, 80, 80))
            $fetchErrors += $applianceHost
            $applianceNetworkSets[$applianceHost] = @()
        }
        finally {
            if ($session) { Disconnect-OneViewAPIInline -Session $session }
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default

    if ($fetchErrors.Count -gt 0) {
        $continueChoice = [System.Windows.Forms.MessageBox]::Show(
            "Verbindung fehlgeschlagen bei:`n$($fetchErrors -join "`n")`n`nTrotzdem fortfahren (ohne diese Appliances)?",
            "Verbindungsfehler",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($continueChoice -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    # Prüfen ob überhaupt Network Sets vorhanden
    $hasAnySets = $false
    foreach ($key in $applianceNetworkSets.Keys) {
        if ($applianceNetworkSets[$key].Count -gt 0) { $hasAnySets = $true; break }
    }

    $networkSetMapping = @{}

    if ($hasAnySets) {
        # Schritt 4: Network Set Mapping Dialog anzeigen
        $networkSetMapping = Show-NetworkSetMappingDialog -ApplianceNetworkSets $applianceNetworkSets
        if ($null -eq $networkSetMapping) { return }
    } else {
        Write-GUILog "Keine Network Sets auf den Appliances gefunden – VLAN wird ohne Network Set Zuweisung erstellt." -Color ([System.Drawing.Color]::FromArgb(255, 200, 60))
    }

    # Bestätigung
    $summaryLines = @("Netzwerk: $($netParams.NetworkName) (VLAN $($netParams.VlanId))", "")
    foreach ($applianceHost in $selectedAppliances) {
        $assignedSets = if ($networkSetMapping.ContainsKey($applianceHost) -and $networkSetMapping[$applianceHost].Count -gt 0) {
            ($networkSetMapping[$applianceHost]) -join ", "
        } else { "(keine Network Set Zuweisung)" }
        $summaryLines += "$applianceHost : $assignedSets"
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        ($summaryLines -join "`n"),
        "Multi-Deploy bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    # Schritt 5: Deployment
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnImport.Enabled = $false
    $btnBackup.Enabled = $false
    $btnMultiDeploy.Enabled = $false
    $btnNSImport.Enabled = $false
    $btnNSBackup.Enabled = $false
    $rtbLog.Clear()

    $tempExcelFiles = @()

    try {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force
        }
        Import-Module ImportExcel

        Write-GUILog "Netzwerk: $($netParams.NetworkName) (VLAN $($netParams.VlanId))" -Color ([System.Drawing.Color]::Cyan)
        Write-GUILog "Ziel-Appliances: $($selectedAppliances.Count)" -Color ([System.Drawing.Color]::Cyan)
        Write-GUILog "" -Color ([System.Drawing.Color]::Cyan)

        $logsDir = Join-Path $scriptDir "Logs"
        if (-not (Test-Path $logsDir)) { New-Item -Path $logsDir -ItemType Directory -Force | Out-Null }
        $sharedLogPath = Join-Path $logsDir ("VLAN_MultiDeploy_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

        $successCount = 0
        $errorCount = 0

        foreach ($applianceHost in $selectedAppliances) {
            Write-GUILog "Erstelle Netzwerk auf $applianceHost ..." -Color ([System.Drawing.Color]::Cyan)

            # Network Set Wert für diese Appliance (mit '; ' getrennt bei Mehrfachzuweisung)
            $netSetValue = ""
            if ($networkSetMapping.ContainsKey($applianceHost) -and $networkSetMapping[$applianceHost].Count -gt 0) {
                $netSetValue = ($networkSetMapping[$applianceHost]) -join "; "
                Write-GUILog "  Network Sets: $netSetValue" -Color ([System.Drawing.Color]::FromArgb(200, 200, 200))
            }

            # Pro Appliance eine eigene temporäre Excel-Datei
            $safeName = $applianceHost -replace '[\\/:*?\"<>|\.]', '_'
            $tempExcel = Join-Path $scriptDir ("temp_multideploy_{0}_{1}.xlsx" -f $safeName, (Get-Date -Format "yyyyMMdd_HHmmss"))
            $tempExcelFiles += $tempExcel

            [PSCustomObject]@{
                NetworkName          = $netParams.NetworkName
                VlanId               = $netParams.VlanId
                EthernetNetworkType  = $netParams.EthernetNetworkType
                Purpose              = $netParams.Purpose
                SmartLink            = $netParams.SmartLink
                PrivateNetwork       = $netParams.PrivateNetwork
                PreferredBandwidthGb = $netParams.PreferredBandwidthGb
                MaximumBandwidthGb   = $netParams.MaximumBandwidthGb
                Scope                = $netParams.Scope
                NetworkSet           = $netSetValue
            } | Export-Excel -Path $tempExcel -WorksheetName "VLANs"

            $tempConfig = New-TempConfig -Hostname $applianceHost -ExcelPath $tempExcel -TempDir $scriptDir

            try {
                $importScript = Join-Path $scriptDir "Create-EthernetNetworks.ps1"
                $psCommand = @"
`$env:OV_USERNAME = '$($txtUser.Text)'
`$env:OV_PASSWORD = '$($txtPass.Text -replace "'", "''")'
`$secPass = ConvertTo-SecureString `$env:OV_PASSWORD -AsPlainText -Force
`$global:guiCredential = New-Object System.Management.Automation.PSCredential(`$env:OV_USERNAME, `$secPass)
function Get-Credential { param([string]`$Message) return `$global:guiCredential }
function Read-Host { param([string]`$Prompt) return 'J' }
& '$importScript' -ConfigPath '$tempConfig' -LogPath '$sharedLogPath'
"@
                $exitCode = Invoke-SubprocessWithLiveOutput -Command $psCommand

                if ($exitCode -eq 0) {
                    Write-GUILog "Erfolgreich auf $applianceHost" -Color ([System.Drawing.Color]::FromArgb(80, 220, 80))
                    $successCount++
                } else {
                    Write-GUILog "Fehlgeschlagen auf $applianceHost" -Color ([System.Drawing.Color]::FromArgb(255, 80, 80))
                    $errorCount++
                }
            }
            catch {
                Write-GUILog "Fehler bei $applianceHost : $_" -Color ([System.Drawing.Color]::Red)
                $errorCount++
            }
            finally {
                Remove-Item -Path $tempConfig -Force -ErrorAction SilentlyContinue
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        Write-GUILog "" -Color ([System.Drawing.Color]::Cyan)
        Write-GUILog "Multi-Deploy abgeschlossen: $successCount erfolgreich, $errorCount fehlgeschlagen" -Color ([System.Drawing.Color]::Cyan)

        [System.Windows.Forms.MessageBox]::Show(
            "Netzwerk-Erstellung abgeschlossen!`n`nNetzwerk: $($netParams.NetworkName) (VLAN $($netParams.VlanId))`nErfolgreich: $successCount`nFehlgeschlagen: $errorCount",
            "Multi-Deploy Ergebnis",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
    }
    finally {
        foreach ($tf in $tempExcelFiles) {
            if (Test-Path $tf) { Remove-Item -Path $tf -Force -ErrorAction SilentlyContinue }
        }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnImport.Enabled = $true
        $btnBackup.Enabled = $true
        $btnMultiDeploy.Enabled = $true
        $btnNSImport.Enabled = $true
        $btnNSBackup.Enabled = $true
    }
})

# ============================================================================
#  Button-Event: Network Sets importieren (Einzelne Appliance)
# ============================================================================
$btnNSImport.Add_Click({
    if (-not (Test-InputValid)) { return }

    if (-not $appliances -or $appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Keine Appliances in der Appliances.txt gefunden.",
            "Keine Appliances",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $hostname = Show-SingleApplianceSelectionDialog -Appliances $appliances -Title "Network Set Import – Appliance auswählen"
    if (-not $hostname) { return }

    # Excel-Datei auswählen
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Excel-Dateien (*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*"
    $ofd.InitialDirectory = $scriptDir
    $ofd.Title = "Excel-Datei mit Network Sets auswählen"
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $excelPath = $ofd.FileName

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Network Sets aus der Excel-Datei auf`n$hostname`nerstellen/aktualisieren?`n`nDatei: $excelPath",
        "Network Set Import bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnNSImport.Enabled = $false
    $rtbLog.Clear()

    try {
        Write-GUILog "Starte Network Set Import auf $hostname ..." -Color ([System.Drawing.Color]::Cyan)

        $tempConfig = New-TempConfigNS -Hostname $hostname `
            -ExcelPath $excelPath `
            -TempDir $scriptDir

        $importScript = Join-Path $scriptDir "Create-NetworkSets.ps1"
        $psCommand = @"
`$env:OV_USERNAME = '$($txtUser.Text)'
`$env:OV_PASSWORD = '$($txtPass.Text -replace "'", "''")'
`$secPass = ConvertTo-SecureString `$env:OV_PASSWORD -AsPlainText -Force
`$global:guiCredential = New-Object System.Management.Automation.PSCredential(`$env:OV_USERNAME, `$secPass)
function Get-Credential { param([string]`$Message) return `$global:guiCredential }
function Read-Host { param([string]`$Prompt) return 'J' }
& '$importScript' -ConfigPath '$tempConfig'
"@
        Invoke-SubprocessWithLiveOutput -Command $psCommand | Out-Null

        Write-GUILog "Network Set Import abgeschlossen." -Color ([System.Drawing.Color]::Cyan)

        Remove-Item -Path $tempConfig -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnNSImport.Enabled = $true
    }
})

# ============================================================================
#  Button-Event: Network Set Backup (Multi-Appliance)
# ============================================================================
$btnNSBackup.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtUser.Text) -or [string]::IsNullOrWhiteSpace($txtPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte Benutzername und Kennwort eingeben.",
            "Fehlende Anmeldeinformationen",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if (-not $appliances -or $appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Keine Appliances in der Appliances.txt gefunden.",
            "Keine Appliances",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $selectedAppliances = Show-ApplianceSelectionDialog -Appliances $appliances -Title "Network Set Backup – Appliances auswählen"
    if (-not $selectedAppliances -or $selectedAppliances.Count -eq 0) { return }

    $applianceList = $selectedAppliances -join "`n"
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Network Set Backup von $($selectedAppliances.Count) Appliance(s) erstellen?`n`n$applianceList",
        "Network Set Backup bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnImport.Enabled = $false
    $btnBackup.Enabled = $false
    $btnMultiDeploy.Enabled = $false
    $btnNSImport.Enabled = $false
    $btnNSBackup.Enabled = $false
    $rtbLog.Clear()

    try {
        $backupDir = Join-Path $scriptDir "Backups" (Get-Date -Format "yyyyMMdd_HHmmss")
        New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
        Write-GUILog "Backup-Verzeichnis: $backupDir" -Color ([System.Drawing.Color]::Cyan)

        $logsDir = Join-Path $scriptDir "Logs"
        if (-not (Test-Path $logsDir)) { New-Item -Path $logsDir -ItemType Directory -Force | Out-Null }
        $sharedLogPath = Join-Path $logsDir ("NetworkSet_Export_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

        $successCount = 0
        $errorCount = 0

        foreach ($applianceHost in $selectedAppliances) {
            Write-GUILog "Starte Network Set Backup von $applianceHost ..." -Color ([System.Drawing.Color]::Cyan)

            $safeName = $applianceHost -replace '[\\/:*?\"<>|\.]', '_'
            $exportPath = Join-Path $backupDir ("NetworkSets_${safeName}.xlsx")

            $tempConfig = New-TempConfigNS -Hostname $applianceHost -ExcelPath "" -TempDir $scriptDir

            try {
                $exportScript = Join-Path $scriptDir "Export-NetworkSets.ps1"
                $psCommand = @"
`$env:OV_USERNAME = '$($txtUser.Text)'
`$env:OV_PASSWORD = '$($txtPass.Text -replace "'", "''")'
`$secPass = ConvertTo-SecureString `$env:OV_PASSWORD -AsPlainText -Force
`$global:guiCredential = New-Object System.Management.Automation.PSCredential(`$env:OV_USERNAME, `$secPass)
function Get-Credential { param([string]`$Message) return `$global:guiCredential }
& '$exportScript' -ConfigPath '$tempConfig' -OutputPath '$exportPath' -LogPath '$sharedLogPath'
"@
                $exitCode = Invoke-SubprocessWithLiveOutput -Command $psCommand

                if ($exitCode -eq 0 -and (Test-Path $exportPath)) {
                    Write-GUILog "Network Set Backup erfolgreich: $applianceHost" -Color ([System.Drawing.Color]::FromArgb(80, 220, 80))
                    $successCount++
                } else {
                    Write-GUILog "Network Set Backup fehlgeschlagen: $applianceHost" -Color ([System.Drawing.Color]::FromArgb(255, 80, 80))
                    $errorCount++
                }
            }
            catch {
                Write-GUILog "Fehler bei $applianceHost : $_" -Color ([System.Drawing.Color]::Red)
                $errorCount++
            }
            finally {
                Remove-Item -Path $tempConfig -Force -ErrorAction SilentlyContinue
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        Write-GUILog "" -Color ([System.Drawing.Color]::Cyan)
        Write-GUILog "Network Set Backup abgeschlossen: $successCount erfolgreich, $errorCount fehlgeschlagen" -Color ([System.Drawing.Color]::Cyan)
        Write-GUILog "Verzeichnis: $backupDir" -Color ([System.Drawing.Color]::Cyan)

        [System.Windows.Forms.MessageBox]::Show(
            "Network Set Backup abgeschlossen!`n`nErfolgreich: $successCount`nFehlgeschlagen: $errorCount`n`nVerzeichnis: $backupDir",
            "Network Set Backup Ergebnis",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnImport.Enabled = $true
        $btnBackup.Enabled = $true
        $btnMultiDeploy.Enabled = $true
        $btnNSImport.Enabled = $true
        $btnNSBackup.Enabled = $true
    }
})

# ============================================================================
#  Formular anzeigen
# ============================================================================
$form.Add_Shown({ $txtUser.Focus() })
[void]$form.ShowDialog()
