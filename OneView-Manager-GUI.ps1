<#
.SYNOPSIS
    GUI für HPE OneView Manager – Netzwerke, Network Sets & Server Profiles.

.DESCRIPTION
    Grafische Oberfläche zum Verwalten von Ethernet Networks, Network Sets
    und Server Profiles in HPE OneView. Fragt Anmeldedaten ab und bietet
    ein Auswahlmenü für die OneView Appliance (aus Appliances.txt).

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
$form.Text = "HPE OneView Manager"
$form.Size = New-Object System.Drawing.Size(820, 1080)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(700, 850)
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
#  GroupBox: Server Profiles
# ============================================================================
$grpSP = New-Object System.Windows.Forms.GroupBox
$grpSP.Text = "Server Profiles"
$grpSP.Location = New-Object System.Drawing.Point(15, 295)
$grpSP.Size = New-Object System.Drawing.Size(770, 80)
$grpSP.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($grpSP)

$tblSP = New-Object System.Windows.Forms.TableLayoutPanel
$tblSP.Location = New-Object System.Drawing.Point(10, 22)
$tblSP.Size = New-Object System.Drawing.Size(748, 46)
$tblSP.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tblSP.ColumnCount = 4
$tblSP.RowCount = 1
$tblSP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$tblSP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$tblSP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$tblSP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$tblSP.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$grpSP.Controls.Add($tblSP)

# Button: Server Profile exportieren
$btnSPExport = New-Object System.Windows.Forms.Button
$btnSPExport.Text = "SP exportieren (JSON)"
$btnSPExport.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnSPExport.Margin = New-Object System.Windows.Forms.Padding(0, 0, 3, 0)
$btnSPExport.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 160)
$btnSPExport.ForeColor = [System.Drawing.Color]::White
$btnSPExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSPExport.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblSP.Controls.Add($btnSPExport, 0, 0)

# Button: Server Profile importieren
$btnSPImport = New-Object System.Windows.Forms.Button
$btnSPImport.Text = "SP importieren (JSON)"
$btnSPImport.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnSPImport.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$btnSPImport.BackColor = [System.Drawing.Color]::FromArgb(160, 60, 60)
$btnSPImport.ForeColor = [System.Drawing.Color]::White
$btnSPImport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSPImport.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblSP.Controls.Add($btnSPImport, 1, 0)

# Button: Server Profile verwalten
$btnSPManage = New-Object System.Windows.Forms.Button
$btnSPManage.Text = "SP verwalten"
$btnSPManage.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnSPManage.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$btnSPManage.BackColor = [System.Drawing.Color]::FromArgb(100, 60, 160)
$btnSPManage.ForeColor = [System.Drawing.Color]::White
$btnSPManage.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSPManage.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblSP.Controls.Add($btnSPManage, 2, 0)

# Button: JSON Editor
$btnSPJsonEdit = New-Object System.Windows.Forms.Button
$btnSPJsonEdit.Text = "SP JSON Editor"
$btnSPJsonEdit.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnSPJsonEdit.Margin = New-Object System.Windows.Forms.Padding(3, 0, 0, 0)
$btnSPJsonEdit.BackColor = [System.Drawing.Color]::FromArgb(50, 120, 80)
$btnSPJsonEdit.ForeColor = [System.Drawing.Color]::White
$btnSPJsonEdit.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSPJsonEdit.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblSP.Controls.Add($btnSPJsonEdit, 3, 0)

# ============================================================================
#  GroupBox: Server Profile Templates
# ============================================================================
$grpSPT = New-Object System.Windows.Forms.GroupBox
$grpSPT.Text = "Server Profile Templates"
$grpSPT.Location = New-Object System.Drawing.Point(15, 385)
$grpSPT.Size = New-Object System.Drawing.Size(770, 80)
$grpSPT.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($grpSPT)

$tblSPT = New-Object System.Windows.Forms.TableLayoutPanel
$tblSPT.Location = New-Object System.Drawing.Point(10, 22)
$tblSPT.Size = New-Object System.Drawing.Size(748, 46)
$tblSPT.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tblSPT.ColumnCount = 4
$tblSPT.RowCount = 1
$tblSPT.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$tblSPT.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$tblSPT.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$tblSPT.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$tblSPT.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$grpSPT.Controls.Add($tblSPT)

# Button: SPT exportieren
$btnSPTExport = New-Object System.Windows.Forms.Button
$btnSPTExport.Text = "SPT exportieren (JSON)"
$btnSPTExport.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnSPTExport.Margin = New-Object System.Windows.Forms.Padding(0, 0, 3, 0)
$btnSPTExport.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 160)
$btnSPTExport.ForeColor = [System.Drawing.Color]::White
$btnSPTExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSPTExport.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblSPT.Controls.Add($btnSPTExport, 0, 0)

# Button: SPT importieren
$btnSPTImport = New-Object System.Windows.Forms.Button
$btnSPTImport.Text = "SPT importieren (JSON)"
$btnSPTImport.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnSPTImport.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$btnSPTImport.BackColor = [System.Drawing.Color]::FromArgb(160, 60, 60)
$btnSPTImport.ForeColor = [System.Drawing.Color]::White
$btnSPTImport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSPTImport.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblSPT.Controls.Add($btnSPTImport, 1, 0)

# Button: SPT verwalten
$btnSPTManage = New-Object System.Windows.Forms.Button
$btnSPTManage.Text = "SPT verwalten"
$btnSPTManage.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnSPTManage.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$btnSPTManage.BackColor = [System.Drawing.Color]::FromArgb(100, 60, 160)
$btnSPTManage.ForeColor = [System.Drawing.Color]::White
$btnSPTManage.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSPTManage.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblSPT.Controls.Add($btnSPTManage, 2, 0)

# Button: SPT JSON Editor
$btnSPTJsonEdit = New-Object System.Windows.Forms.Button
$btnSPTJsonEdit.Text = "SPT JSON Editor"
$btnSPTJsonEdit.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnSPTJsonEdit.Margin = New-Object System.Windows.Forms.Padding(3, 0, 0, 0)
$btnSPTJsonEdit.BackColor = [System.Drawing.Color]::FromArgb(50, 120, 80)
$btnSPTJsonEdit.ForeColor = [System.Drawing.Color]::White
$btnSPTJsonEdit.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSPTJsonEdit.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tblSPT.Controls.Add($btnSPTJsonEdit, 3, 0)

# ============================================================================
#  Ausgabebereich (Log)
# ============================================================================
$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = "Protokoll"
$grpLog.Location = New-Object System.Drawing.Point(15, 480)
$grpLog.Size = New-Object System.Drawing.Size(770, 440)
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
    $dlg.Size = New-Object System.Drawing.Size(550, 490)
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

    # Typ-Filter-Buttons dynamisch erzeugen (z.B. "Alle ESXi", "Alle VDI")
    $types = @($Appliances | Where-Object { $_.Type } | ForEach-Object { $_.Type } | Select-Object -Unique | Sort-Object)
    $typeButtonX = 275
    foreach ($typeName in $types) {
        $btnType = New-Object System.Windows.Forms.Button
        $btnType.Text = "Alle $typeName"
        $btnType.Location = New-Object System.Drawing.Point($typeButtonX, 340)
        $btnType.Size = New-Object System.Drawing.Size(85, 28)
        $btnType.Tag = $typeName
        $btnType.Add_Click({
            $filterType = $this.Tag
            for ($i = 0; $i -lt $clb.Items.Count; $i++) {
                $clb.SetItemChecked($i, ($Appliances[$i].Type -eq $filterType))
            }
        })
        $dlg.Controls.Add($btnType)
        $typeButtonX += 95
    }

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(370, 380)
    $btnOK.Size = New-Object System.Drawing.Size(55, 28)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.AcceptButton = $btnOK
    $dlg.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Location = New-Object System.Drawing.Point(430, 380)
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

# ============================================================================
#  OneView API – Server Profile Funktionen (Inline)
# ============================================================================

function Get-ServerProfilesInline {
    param([Parameter(Mandatory)][hashtable]$Session)
    $allMembers = [System.Collections.Generic.List[object]]::new()
    $start = 0; $pageSize = 100
    do {
        $uri = "$($Session.BaseUri)/rest/server-profiles?start=$start&count=$pageSize"
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck
        if ($response.members) { $allMembers.AddRange([object[]]$response.members) }
        $total = $response.total
        $start += $response.members.Count
    } while ($allMembers.Count -lt $total)
    return $allMembers
}

function Get-ServerProfileTemplatesInline {
    param([Parameter(Mandatory)][hashtable]$Session)
    $allMembers = [System.Collections.Generic.List[object]]::new()
    $start = 0; $pageSize = 100
    do {
        $uri = "$($Session.BaseUri)/rest/server-profile-templates?start=$start&count=$pageSize"
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck
        if ($response.members) { $allMembers.AddRange([object[]]$response.members) }
        $total = $response.total
        $start += $response.members.Count
    } while ($allMembers.Count -lt $total)
    return $allMembers
}

function Get-ServerHardwareInline {
    param([Parameter(Mandatory)][hashtable]$Session)
    $allMembers = [System.Collections.Generic.List[object]]::new()
    $start = 0; $pageSize = 100
    do {
        $uri = "$($Session.BaseUri)/rest/server-hardware?start=$start&count=$pageSize"
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck
        if ($response.members) { $allMembers.AddRange([object[]]$response.members) }
        $total = $response.total
        $start += $response.members.Count
    } while ($allMembers.Count -lt $total)
    return $allMembers
}

function New-ServerProfileFromTemplateInline {
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][string]$TemplateUri,
        [Parameter(Mandatory)][string]$ProfileName,
        [string]$Description = "",
        [string]$ServerHardwareUri = ""
    )

    # Neues Profil aus Template erzeugen lassen
    $newProfileUri = "$($Session.BaseUri)$TemplateUri/new-profile"
    $newProfile = Invoke-RestMethod -Uri $newProfileUri -Method Get -Headers $Session.Headers -SkipCertificateCheck

    $newProfile.name = $ProfileName
    $newProfile.description = $Description
    if ($ServerHardwareUri) {
        $newProfile.serverHardwareUri = $ServerHardwareUri
    }

    $body = $newProfile | ConvertTo-Json -Depth 20
    $response = Invoke-RestMethod -Uri "$($Session.BaseUri)/rest/server-profiles" `
        -Method Post -Headers $Session.Headers -Body $body -SkipCertificateCheck
    return $response
}

function Update-ServerProfileInline {
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][object]$Profile
    )
    $body = $Profile | ConvertTo-Json -Depth 20
    $response = Invoke-RestMethod -Uri "$($Session.BaseUri)$($Profile.uri)" `
        -Method Put -Headers $Session.Headers -Body $body -SkipCertificateCheck
    return $response
}

function Remove-ServerProfileInline {
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][string]$ProfileUri
    )
    $response = Invoke-RestMethod -Uri "$($Session.BaseUri)$ProfileUri" `
        -Method Delete -Headers $Session.Headers -SkipCertificateCheck
    return $response
}

# ============================================================================
#  OneView API – Server Profile Template Funktionen (Inline)
# ============================================================================

function Update-ServerProfileTemplateInline {
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][object]$Template
    )
    $body = $Template | ConvertTo-Json -Depth 20
    $response = Invoke-RestMethod -Uri "$($Session.BaseUri)$($Template.uri)" `
        -Method Put -Headers $Session.Headers -Body $body -SkipCertificateCheck
    return $response
}

function Remove-ServerProfileTemplateInline {
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][string]$TemplateUri
    )
    $response = Invoke-RestMethod -Uri "$($Session.BaseUri)$TemplateUri" `
        -Method Delete -Headers $Session.Headers -SkipCertificateCheck
    return $response
}

function New-ServerProfileTemplateFromProfileInline {
    <#
    .SYNOPSIS  Erstellt ein neues Server Profile Template aus einem existierenden Server Profile.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][string]$ServerProfileUri,
        [string]$TemplateName = "",
        [string]$Description = ""
    )

    # Profil laden
    $profile = Invoke-RestMethod -Uri "$($Session.BaseUri)$ServerProfileUri" `
        -Method Get -Headers $Session.Headers -SkipCertificateCheck

    # Template-Body erstellen: POST /rest/server-profile-templates erwartet die Template-Felder
    $templateBody = @{
        type                    = "ServerProfileTemplateV8"
        name                    = if ($TemplateName) { $TemplateName } else { "$($profile.name)_Template" }
        description             = if ($Description) { $Description } else { "Erstellt aus Profil: $($profile.name)" }
        serverHardwareTypeUri   = $profile.serverHardwareTypeUri
        enclosureGroupUri       = $profile.enclosureGroupUri
        connectionSettings      = $profile.connectionSettings
        boot                    = $profile.boot
        bootMode                = $profile.bootMode
        bios                    = $profile.bios
        firmware                = $profile.firmware
        localStorage            = $profile.localStorage
        sanStorage              = $profile.sanStorage
        managementProcessor     = $profile.managementProcessor
    }

    $body = $templateBody | ConvertTo-Json -Depth 20
    $response = Invoke-RestMethod -Uri "$($Session.BaseUri)/rest/server-profile-templates" `
        -Method Post -Headers $Session.Headers -Body $body -SkipCertificateCheck
    return $response
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

# ============================================================================
#  Server Profile Management Dialog
# ============================================================================

function Show-ServerProfileManageDialog {
    <#
    .SYNOPSIS  Zeigt einen Dialog zum Verwalten von Server Profiles (Anzeigen, Bearbeiten, Neu, Löschen).
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][string]$Hostname
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Server Profile verwalten – $Hostname"
    $dlg.Size = New-Object System.Drawing.Size(950, 700)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dlg.MinimumSize = New-Object System.Drawing.Size(800, 550)
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # ── Toolbar ──
    $pnlToolbar = New-Object System.Windows.Forms.Panel
    $pnlToolbar.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlToolbar.Height = 40
    $dlg.Controls.Add($pnlToolbar)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Aktualisieren"
    $btnRefresh.Location = New-Object System.Drawing.Point(10, 6)
    $btnRefresh.Size = New-Object System.Drawing.Size(110, 28)
    $pnlToolbar.Controls.Add($btnRefresh)

    $btnNewSP = New-Object System.Windows.Forms.Button
    $btnNewSP.Text = "Neu erstellen"
    $btnNewSP.Location = New-Object System.Drawing.Point(130, 6)
    $btnNewSP.Size = New-Object System.Drawing.Size(110, 28)
    $btnNewSP.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnNewSP.ForeColor = [System.Drawing.Color]::White
    $btnNewSP.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlToolbar.Controls.Add($btnNewSP)

    $btnEditSP = New-Object System.Windows.Forms.Button
    $btnEditSP.Text = "Bearbeiten"
    $btnEditSP.Location = New-Object System.Drawing.Point(250, 6)
    $btnEditSP.Size = New-Object System.Drawing.Size(100, 28)
    $pnlToolbar.Controls.Add($btnEditSP)

    $btnDeleteSP = New-Object System.Windows.Forms.Button
    $btnDeleteSP.Text = "Löschen"
    $btnDeleteSP.Location = New-Object System.Drawing.Point(360, 6)
    $btnDeleteSP.Size = New-Object System.Drawing.Size(80, 28)
    $btnDeleteSP.BackColor = [System.Drawing.Color]::FromArgb(200, 50, 50)
    $btnDeleteSP.ForeColor = [System.Drawing.Color]::White
    $btnDeleteSP.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlToolbar.Controls.Add($btnDeleteSP)

    $btnExportOne = New-Object System.Windows.Forms.Button
    $btnExportOne.Text = "Als JSON exportieren"
    $btnExportOne.Location = New-Object System.Drawing.Point(450, 6)
    $btnExportOne.Size = New-Object System.Drawing.Size(140, 28)
    $pnlToolbar.Controls.Add($btnExportOne)

    $btnJsonEditor = New-Object System.Windows.Forms.Button
    $btnJsonEditor.Text = "JSON Editor"
    $btnJsonEditor.Location = New-Object System.Drawing.Point(600, 6)
    $btnJsonEditor.Size = New-Object System.Drawing.Size(110, 28)
    $btnJsonEditor.BackColor = [System.Drawing.Color]::FromArgb(50, 120, 80)
    $btnJsonEditor.ForeColor = [System.Drawing.Color]::White
    $btnJsonEditor.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlToolbar.Controls.Add($btnJsonEditor)

    # ── SplitContainer ──
    $splitMain = New-Object System.Windows.Forms.SplitContainer
    $splitMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitMain.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $splitMain.SplitterDistance = 350
    $dlg.Controls.Add($splitMain)
    $splitMain.BringToFront()

    # ── Liste (links) ──
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Dock = [System.Windows.Forms.DockStyle]::Fill
    $dgv.ReadOnly = $true
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $dgv.MultiSelect = $false
    $dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $dgv.BackgroundColor = [System.Drawing.Color]::White
    $dgv.RowHeadersVisible = $false
    $splitMain.Panel1.Controls.Add($dgv)

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = "Name"; $colName.Name = "Name"; $colName.FillWeight = 40
    $colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStatus.HeaderText = "Status"; $colStatus.Name = "Status"; $colStatus.FillWeight = 15
    $colHW = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colHW.HeaderText = "Server Hardware"; $colHW.Name = "Hardware"; $colHW.FillWeight = 30
    $colTemplate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colTemplate.HeaderText = "Template"; $colTemplate.Name = "Template"; $colTemplate.FillWeight = 25
    $dgv.Columns.AddRange(@($colName, $colStatus, $colHW, $colTemplate))

    # ── Details (rechts) ──
    $txtDetails = New-Object System.Windows.Forms.RichTextBox
    $txtDetails.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtDetails.ReadOnly = $true
    $txtDetails.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $txtDetails.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $txtDetails.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtDetails.WordWrap = $false
    $splitMain.Panel2.Controls.Add($txtDetails)

    # ── Daten ──
    $script:spProfiles      = @()
    $script:spTemplates     = @()
    $script:spHardware      = @()
    $script:spTemplateMap   = @{}
    $script:spHardwareMap   = @{}

    $loadProfiles = {
        $dgv.Rows.Clear()
        $txtDetails.Clear()
        try {
            $script:spProfiles  = @(Get-ServerProfilesInline -Session $Session)
            $script:spTemplates = @(Get-ServerProfileTemplatesInline -Session $Session)
            $script:spHardware  = @(Get-ServerHardwareInline -Session $Session)

            $script:spTemplateMap = @{}
            foreach ($t in $script:spTemplates) { $script:spTemplateMap[$t.uri] = $t.name }

            $script:spHardwareMap = @{}
            foreach ($h in $script:spHardware) { $script:spHardwareMap[$h.uri] = $h.name }

            foreach ($p in $script:spProfiles) {
                $hwName   = if ($p.serverHardwareUri -and $script:spHardwareMap.ContainsKey($p.serverHardwareUri)) { $script:spHardwareMap[$p.serverHardwareUri] } else { "(ohne)" }
                $tplName  = if ($p.serverProfileTemplateUri -and $script:spTemplateMap.ContainsKey($p.serverProfileTemplateUri)) { $script:spTemplateMap[$p.serverProfileTemplateUri] } else { "(ohne)" }
                $dgv.Rows.Add($p.name, $p.status, $hwName, $tplName) | Out-Null
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Laden: $_", "Fehler",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    }

    & $loadProfiles

    # ── Selection Changed ──
    $dgv.Add_SelectionChanged({
        if ($dgv.SelectedRows.Count -eq 0) { return }
        $idx = $dgv.SelectedRows[0].Index
        if ($idx -ge 0 -and $idx -lt $script:spProfiles.Count) {
            $selected = $script:spProfiles[$idx]
            $json = $selected | ConvertTo-Json -Depth 10
            $txtDetails.Clear()
            $txtDetails.Text = $json
        }
    })

    # ── Refresh ──
    $btnRefresh.Add_Click({ & $loadProfiles })

    # ── Neu erstellen ──
    $btnNewSP.Add_Click({
        $result = Show-ServerProfileEditDialog -Session $Session -Mode "Create" `
            -Templates $script:spTemplates -Hardware $script:spHardware `
            -HardwareMap $script:spHardwareMap
        if ($result) { & $loadProfiles }
    })

    # ── Bearbeiten ──
    $btnEditSP.Add_Click({
        if ($dgv.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte ein Profil auswählen.", "Hinweis",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $idx = $dgv.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:spProfiles.Count) { return }
        $profile = $script:spProfiles[$idx]

        $result = Show-ServerProfileEditDialog -Session $Session -Mode "Edit" `
            -ExistingProfile $profile -Templates $script:spTemplates `
            -Hardware $script:spHardware -HardwareMap $script:spHardwareMap
        if ($result) { & $loadProfiles }
    })

    # ── Löschen ──
    $btnDeleteSP.Add_Click({
        if ($dgv.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte ein Profil auswählen.", "Hinweis",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $idx = $dgv.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:spProfiles.Count) { return }
        $profile = $script:spProfiles[$idx]

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Server Profile '$($profile.name)' wirklich löschen?`n`nDiese Aktion kann nicht rückgängig gemacht werden!",
            "Profil löschen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        try {
            Remove-ServerProfileInline -Session $Session -ProfileUri $profile.uri
            [System.Windows.Forms.MessageBox]::Show("Profil '$($profile.name)' wurde gelöscht.", "Erfolgreich",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            & $loadProfiles
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Löschen: $_", "Fehler",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    # ── Einzelexport ──
    $btnExportOne.Add_Click({
        if ($dgv.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte ein Profil auswählen.", "Hinweis",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $idx = $dgv.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:spProfiles.Count) { return }
        $profile = $script:spProfiles[$idx]

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "JSON Dateien (*.json)|*.json"
        $sfd.FileName = ($profile.name -replace '[\\/:*?\"<>|\s]', '_') + ".json"
        $sfd.Title = "Profil als JSON exportieren"
        if ($sfd.ShowDialog($dlg) -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $profile | ConvertTo-Json -Depth 20 | Set-Content -Path $sfd.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exportiert nach:`n$($sfd.FileName)", "Export erfolgreich",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    })

    # ── JSON Editor ──
    $btnJsonEditor.Add_Click({
        $initialJson = ""
        if ($dgv.SelectedRows.Count -gt 0) {
            $idx = $dgv.SelectedRows[0].Index
            if ($idx -ge 0 -and $idx -lt $script:spProfiles.Count) {
                $initialJson = $script:spProfiles[$idx] | ConvertTo-Json -Depth 20
            }
        }
        Show-ServerProfileJsonEditor -Session $Session -Hostname $Hostname -InitialJson $initialJson
        & $loadProfiles
    })

    $dlg.ShowDialog($form) | Out-Null
    $dlg.Dispose()
}

function Show-ServerProfileEditDialog {
    <#
    .SYNOPSIS  Dialog zum Erstellen oder Bearbeiten eines Server Profiles.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [ValidateSet("Create","Edit")]
        [string]$Mode = "Create",
        [object]$ExistingProfile = $null,
        [object[]]$Templates = @(),
        [object[]]$Hardware = @(),
        [hashtable]$HardwareMap = @{}
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = if ($Mode -eq "Create") { "Neues Server Profile erstellen" } else { "Server Profile bearbeiten: $($ExistingProfile.name)" }
    $dlg.Size = New-Object System.Drawing.Size(600, 500)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $yPos = 15
    $lblWidth = 170
    $ctrlX = 185
    $ctrlWidth = 380
    $rowHeight = 32

    # Name
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Profil-Name:"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $txtProfileName = New-Object System.Windows.Forms.TextBox
    $txtProfileName.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $txtProfileName.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    if ($ExistingProfile) { $txtProfileName.Text = $ExistingProfile.name }
    $dlg.Controls.Add($txtProfileName)
    $yPos += $rowHeight

    # Description
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Beschreibung:"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $txtProfileDesc = New-Object System.Windows.Forms.TextBox
    $txtProfileDesc.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $txtProfileDesc.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    if ($ExistingProfile) { $txtProfileDesc.Text = $ExistingProfile.description }
    $dlg.Controls.Add($txtProfileDesc)
    $yPos += $rowHeight

    # Server Profile Template (nur bei Create)
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Profile Template:"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $cmbTemplate = New-Object System.Windows.Forms.ComboBox
    $cmbTemplate.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $cmbTemplate.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    $cmbTemplate.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbTemplate.Items.Add("(Kein Template)") | Out-Null
    foreach ($t in ($Templates | Sort-Object -Property name)) {
        $cmbTemplate.Items.Add($t.name) | Out-Null
    }
    $cmbTemplate.SelectedIndex = 0
    if ($Mode -eq "Edit") {
        $cmbTemplate.Enabled = $false
        if ($ExistingProfile.serverProfileTemplateUri) {
            foreach ($t in $Templates) {
                if ($t.uri -eq $ExistingProfile.serverProfileTemplateUri) {
                    $idx = [Array]::IndexOf($cmbTemplate.Items.Cast([string]).ToArray(), $t.name)
                    if ($idx -ge 0) { $cmbTemplate.SelectedIndex = $idx }
                    break
                }
            }
        }
    }
    $dlg.Controls.Add($cmbTemplate)
    $yPos += $rowHeight

    # Server Hardware
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Server Hardware:"
    $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
    $dlg.Controls.Add($lbl)
    $cmbHardware = New-Object System.Windows.Forms.ComboBox
    $cmbHardware.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $cmbHardware.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
    $cmbHardware.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbHardware.Items.Add("(Nicht zugewiesen)") | Out-Null
    foreach ($h in ($Hardware | Sort-Object -Property name)) {
        $statusTag = if ($h.state -eq "NoProfileApplied") { " [verfügbar]" } else { " [belegt]" }
        $cmbHardware.Items.Add("$($h.name)$statusTag") | Out-Null
    }
    $cmbHardware.SelectedIndex = 0
    if ($ExistingProfile -and $ExistingProfile.serverHardwareUri) {
        for ($i = 0; $i -lt $Hardware.Count; $i++) {
            if ($Hardware[$i].uri -eq $ExistingProfile.serverHardwareUri) {
                $cmbHardware.SelectedIndex = $i + 1
                break
            }
        }
    }
    $dlg.Controls.Add($cmbHardware)
    $yPos += $rowHeight

    # Firmware Baseline (Anzeige, nur Edit)
    if ($Mode -eq "Edit") {
        $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Firmware Baseline:"
        $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
        $dlg.Controls.Add($lbl)
        $txtFW = New-Object System.Windows.Forms.TextBox
        $txtFW.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $txtFW.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
        $txtFW.ReadOnly = $true
        $fwUri = if ($ExistingProfile.firmware -and $ExistingProfile.firmware.firmwareBaselineUri) { $ExistingProfile.firmware.firmwareBaselineUri } else { "" }
        $txtFW.Text = if ($fwUri) { $fwUri } else { "(nicht gesetzt)" }
        $dlg.Controls.Add($txtFW)
        $yPos += $rowHeight
    }

    # Boot Order (Anzeige/Bearbeitung, für Edit)
    if ($Mode -eq "Edit" -and $ExistingProfile.boot -and $ExistingProfile.boot.order) {
        $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Boot Order:"
        $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
        $dlg.Controls.Add($lbl)
        $txtBoot = New-Object System.Windows.Forms.TextBox
        $txtBoot.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $txtBoot.Size = New-Object System.Drawing.Size($ctrlWidth, 23)
        $txtBoot.Text = ($ExistingProfile.boot.order -join ", ")
        $dlg.Controls.Add($txtBoot)
        $yPos += $rowHeight
    }

    # Connections Anzeige (nur Edit)
    if ($Mode -eq "Edit" -and $ExistingProfile.connectionSettings -and $ExistingProfile.connectionSettings.connections) {
        $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Connections:"
        $lbl.Location = New-Object System.Drawing.Point(15, ($yPos + 3)); $lbl.Size = New-Object System.Drawing.Size($lblWidth, 20)
        $dlg.Controls.Add($lbl)
        $txtConn = New-Object System.Windows.Forms.TextBox
        $txtConn.Location = New-Object System.Drawing.Point($ctrlX, $yPos); $txtConn.Size = New-Object System.Drawing.Size($ctrlWidth, 80)
        $txtConn.Multiline = $true; $txtConn.ReadOnly = $true; $txtConn.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $connLines = @()
        foreach ($conn in $ExistingProfile.connectionSettings.connections) {
            $connLines += "ID:$($conn.id) $($conn.name) – $($conn.functionType) (Port: $($conn.portId))"
        }
        $txtConn.Text = ($connLines -join "`r`n")
        $dlg.Controls.Add($txtConn)
        $yPos += 86
    }

    $yPos += 20

    # Buttons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = if ($Mode -eq "Create") { "Erstellen" } else { "Speichern" }
    $btnOK.Location = New-Object System.Drawing.Point(370, $yPos)
    $btnOK.Size = New-Object System.Drawing.Size(90, 30)
    $btnOK.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnOK.ForeColor = [System.Drawing.Color]::White
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $dlg.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Location = New-Object System.Drawing.Point(465, $yPos)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.CancelButton = $btnCancel
    $dlg.Controls.Add($btnCancel)

    $script:spEditResult = $false

    $btnOK.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtProfileName.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte einen Profil-Namen eingeben.", "Pflichtfeld",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnOK.Enabled = $false

        try {
            if ($Mode -eq "Create") {
                # Prüfen ob Template ausgewählt
                if ($cmbTemplate.SelectedIndex -gt 0) {
                    $selectedTemplateName = $cmbTemplate.SelectedItem.ToString()
                    $templateObj = $Templates | Where-Object { $_.name -eq $selectedTemplateName } | Select-Object -First 1

                    $hwUri = ""
                    if ($cmbHardware.SelectedIndex -gt 0) {
                        $sortedHW = $Hardware | Sort-Object -Property name
                        $hwUri = $sortedHW[$cmbHardware.SelectedIndex - 1].uri
                    }

                    New-ServerProfileFromTemplateInline -Session $Session `
                        -TemplateUri $templateObj.uri `
                        -ProfileName $txtProfileName.Text.Trim() `
                        -Description $txtProfileDesc.Text.Trim() `
                        -ServerHardwareUri $hwUri

                    [System.Windows.Forms.MessageBox]::Show(
                        "Server Profile '$($txtProfileName.Text.Trim())' wurde erstellt.",
                        "Erfolgreich",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                } else {
                    # Ohne Template – minimales Profil
                    $newProfile = @{
                        type        = "ServerProfileV12"
                        name        = $txtProfileName.Text.Trim()
                        description = $txtProfileDesc.Text.Trim()
                    }

                    if ($cmbHardware.SelectedIndex -gt 0) {
                        $sortedHW = $Hardware | Sort-Object -Property name
                        $newProfile.serverHardwareUri = $sortedHW[$cmbHardware.SelectedIndex - 1].uri
                    }

                    $body = $newProfile | ConvertTo-Json -Depth 20
                    Invoke-RestMethod -Uri "$($Session.BaseUri)/rest/server-profiles" `
                        -Method Post -Headers $Session.Headers -Body $body -SkipCertificateCheck | Out-Null

                    [System.Windows.Forms.MessageBox]::Show(
                        "Server Profile '$($txtProfileName.Text.Trim())' wurde erstellt.",
                        "Erfolgreich",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                }
                $script:spEditResult = $true
                $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $dlg.Close()
            } else {
                # ── Edit-Modus ──
                $ExistingProfile.name        = $txtProfileName.Text.Trim()
                $ExistingProfile.description = $txtProfileDesc.Text.Trim()

                # Server Hardware Zuweisung
                if ($cmbHardware.SelectedIndex -gt 0) {
                    $sortedHW = $Hardware | Sort-Object -Property name
                    $ExistingProfile.serverHardwareUri = $sortedHW[$cmbHardware.SelectedIndex - 1].uri
                } else {
                    $ExistingProfile.serverHardwareUri = $null
                }

                # Boot Order Update
                if ($txtBoot -and -not [string]::IsNullOrWhiteSpace($txtBoot.Text)) {
                    $ExistingProfile.boot.order = @($txtBoot.Text -split ',\s*' | ForEach-Object { $_.Trim() })
                }

                Update-ServerProfileInline -Session $Session -Profile $ExistingProfile | Out-Null

                [System.Windows.Forms.MessageBox]::Show(
                    "Server Profile '$($txtProfileName.Text.Trim())' wurde aktualisiert.",
                    "Erfolgreich",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

                $script:spEditResult = $true
                $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $dlg.Close()
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler: $_", "Fehler",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnOK.Enabled = $true
        }
    })

    $dlg.ShowDialog($form) | Out-Null
    $dlg.Dispose()
    return $script:spEditResult
}

# ============================================================================
#  JSON Editor Dialog – Vollzugriff auf alle Server Profile Felder
# ============================================================================
function Show-ServerProfileJsonEditor {
    <#
    .SYNOPSIS  JSON-Editor zum vollständigen Bearbeiten, Erstellen und Updaten von Server Profiles.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [string]$Hostname = "",
        [string]$InitialJson = "",
        [string]$InitialFilePath = ""
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Server Profile – JSON Editor$(if ($Hostname) { " – $Hostname" })"
    $dlg.Size = New-Object System.Drawing.Size(950, 750)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dlg.MinimumSize = New-Object System.Drawing.Size(700, 500)
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # ── Toolbar oben ──
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 44
    $dlg.Controls.Add($pnlTop)

    $btnLoadFile = New-Object System.Windows.Forms.Button
    $btnLoadFile.Text = "JSON laden"
    $btnLoadFile.Location = New-Object System.Drawing.Point(10, 8)
    $btnLoadFile.Size = New-Object System.Drawing.Size(100, 28)
    $pnlTop.Controls.Add($btnLoadFile)

    $btnSaveFile = New-Object System.Windows.Forms.Button
    $btnSaveFile.Text = "JSON speichern"
    $btnSaveFile.Location = New-Object System.Drawing.Point(118, 8)
    $btnSaveFile.Size = New-Object System.Drawing.Size(110, 28)
    $pnlTop.Controls.Add($btnSaveFile)

    $btnValidate = New-Object System.Windows.Forms.Button
    $btnValidate.Text = "Validieren"
    $btnValidate.Location = New-Object System.Drawing.Point(236, 8)
    $btnValidate.Size = New-Object System.Drawing.Size(90, 28)
    $pnlTop.Controls.Add($btnValidate)

    $btnFormat = New-Object System.Windows.Forms.Button
    $btnFormat.Text = "Formatieren"
    $btnFormat.Location = New-Object System.Drawing.Point(334, 8)
    $btnFormat.Size = New-Object System.Drawing.Size(95, 28)
    $pnlTop.Controls.Add($btnFormat)

    # Separator
    $sep = New-Object System.Windows.Forms.Label
    $sep.Text = "|"
    $sep.Location = New-Object System.Drawing.Point(440, 12)
    $sep.Size = New-Object System.Drawing.Size(10, 20)
    $sep.ForeColor = [System.Drawing.Color]::Gray
    $pnlTop.Controls.Add($sep)

    $btnCreateOV = New-Object System.Windows.Forms.Button
    $btnCreateOV.Text = "Neu anlegen (OneView)"
    $btnCreateOV.Location = New-Object System.Drawing.Point(458, 8)
    $btnCreateOV.Size = New-Object System.Drawing.Size(155, 28)
    $btnCreateOV.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnCreateOV.ForeColor = [System.Drawing.Color]::White
    $btnCreateOV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlTop.Controls.Add($btnCreateOV)

    $btnUpdateOV = New-Object System.Windows.Forms.Button
    $btnUpdateOV.Text = "Update (OneView)"
    $btnUpdateOV.Location = New-Object System.Drawing.Point(621, 8)
    $btnUpdateOV.Size = New-Object System.Drawing.Size(135, 28)
    $btnUpdateOV.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 80)
    $btnUpdateOV.ForeColor = [System.Drawing.Color]::White
    $btnUpdateOV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlTop.Controls.Add($btnUpdateOV)

    $btnLoadOV = New-Object System.Windows.Forms.Button
    $btnLoadOV.Text = "Von OneView laden"
    $btnLoadOV.Location = New-Object System.Drawing.Point(764, 8)
    $btnLoadOV.Size = New-Object System.Drawing.Size(135, 28)
    $btnLoadOV.BackColor = [System.Drawing.Color]::FromArgb(100, 60, 160)
    $btnLoadOV.ForeColor = [System.Drawing.Color]::White
    $btnLoadOV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlTop.Controls.Add($btnLoadOV)

    # ── Statusleiste ──
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $lblStatus.Height = 24
    $lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblStatus.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $lblStatus.Text = "Bereit"
    $dlg.Controls.Add($lblStatus)

    # ── JSON Text Editor ──
    $txtJson = New-Object System.Windows.Forms.RichTextBox
    $txtJson.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtJson.Font = New-Object System.Drawing.Font("Consolas", 10)
    $txtJson.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $txtJson.ForeColor = [System.Drawing.Color]::FromArgb(212, 212, 212)
    $txtJson.WordWrap = $false
    $txtJson.AcceptsTab = $true
    $txtJson.ShortcutsEnabled = $true
    $txtJson.DetectUrls = $false
    $dlg.Controls.Add($txtJson)
    $txtJson.BringToFront()

    # Aktuelle Datei merken
    $script:jsonEditorFilePath = $InitialFilePath

    # Initialer Inhalt
    if ($InitialJson) {
        $txtJson.Text = $InitialJson
        $lblStatus.Text = if ($InitialFilePath) { "Geladen: $InitialFilePath" } else { "Profil aus OneView geladen" }
    }

    # ── Hilfsfunktion: JSON validieren ──
    $validateJson = {
        try {
            $null = $txtJson.Text | ConvertFrom-Json -ErrorAction Stop
            return $true
        } catch {
            return $false
        }
    }

    # ── Hilfsfunktion: Read-Only-Felder entfernen (für Create) ──
    $removeReadOnlyFields = {
        param([hashtable]$ht)
        $readOnlyFields = @(
            "uri", "eTag", "created", "modified", "uuid", "serialNumber",
            "serialNumberType", "taskUri", "stateReason", "refreshState",
            "associatedServer", "inProgress", "scopesUri"
        )
        foreach ($f in $readOnlyFields) { $ht.Remove($f) }
        return $ht
    }

    # ── JSON laden ──
    $btnLoadFile.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "JSON Dateien (*.json)|*.json|Alle Dateien (*.*)|*.*"
        $ofd.Title = "Server Profile JSON laden"
        if ($ofd.ShowDialog($dlg) -ne [System.Windows.Forms.DialogResult]::OK) { return }

        try {
            $content = Get-Content -Path $ofd.FileName -Raw -Encoding UTF8
            # Validieren
            $null = $content | ConvertFrom-Json -ErrorAction Stop
            # Formatiert anzeigen
            $obj = $content | ConvertFrom-Json
            $txtJson.Text = $obj | ConvertTo-Json -Depth 20
            $script:jsonEditorFilePath = $ofd.FileName
            $lblStatus.Text = "Geladen: $($ofd.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Laden der Datei:`n$_",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    # ── JSON speichern ──
    $btnSaveFile.Add_Click({
        if (-not (& $validateJson)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Das JSON ist ungültig. Bitte korrigieren Sie die Syntax.",
                "JSON ungültig", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "JSON Dateien (*.json)|*.json"
        $sfd.Title = "Server Profile JSON speichern"
        if ($script:jsonEditorFilePath) {
            $sfd.FileName = [System.IO.Path]::GetFileName($script:jsonEditorFilePath)
            $sfd.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:jsonEditorFilePath)
        } else {
            try {
                $obj = $txtJson.Text | ConvertFrom-Json
                if ($obj.name) { $sfd.FileName = ($obj.name -replace '[\\/:*?\"<>|\s]', '_') + ".json" }
            } catch {}
        }
        if ($sfd.ShowDialog($dlg) -ne [System.Windows.Forms.DialogResult]::OK) { return }

        try {
            # Formatiert speichern
            $obj = $txtJson.Text | ConvertFrom-Json
            $obj | ConvertTo-Json -Depth 20 | Set-Content -Path $sfd.FileName -Encoding UTF8
            $script:jsonEditorFilePath = $sfd.FileName
            $lblStatus.Text = "Gespeichert: $($sfd.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Speichern:`n$_",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    # ── Validieren ──
    $btnValidate.Add_Click({
        if (& $validateJson) {
            $obj = $txtJson.Text | ConvertFrom-Json
            $fields = ($obj | Get-Member -MemberType NoteProperty).Count
            $lblStatus.Text = "JSON ist gültig ($fields Felder auf Root-Ebene)"
            [System.Windows.Forms.MessageBox]::Show(
                "JSON ist gültig.`n`nRoot-Felder: $fields`nName: $($obj.name)",
                "Validierung OK", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } else {
            $lblStatus.Text = "JSON ist UNGÜLTIG!"
            try {
                $null = $txtJson.Text | ConvertFrom-Json -ErrorAction Stop
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "JSON ist ungültig:`n`n$($_.Exception.Message)",
                    "Validierung fehlgeschlagen", [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            }
        }
    })

    # ── Formatieren ──
    $btnFormat.Add_Click({
        if (-not (& $validateJson)) {
            [System.Windows.Forms.MessageBox]::Show(
                "JSON kann nicht formatiert werden – Syntax ungültig.",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        $obj = $txtJson.Text | ConvertFrom-Json
        $txtJson.Text = $obj | ConvertTo-Json -Depth 20
        $lblStatus.Text = "JSON formatiert"
    })

    # ── Neu anlegen in OneView ──
    $btnCreateOV.Add_Click({
        if (-not (& $validateJson)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Das JSON ist ungültig. Bitte zuerst korrigieren.",
                "JSON ungültig", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $obj = $txtJson.Text | ConvertFrom-Json
        $profileName = if ($obj.name) { $obj.name } else { "(unbenannt)" }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Server Profile '$profileName' als neues Profil in OneView anlegen?`n`nRead-Only-Felder (uri, eTag, uuid, etc.) werden automatisch entfernt.",
            "Profil erstellen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnCreateOV.Enabled = $false
        $btnUpdateOV.Enabled = $false
        $lblStatus.Text = "Erstelle Profil '$profileName'..."

        try {
            $ht = $txtJson.Text | ConvertFrom-Json -AsHashtable
            $cleanHt = & $removeReadOnlyFields $ht

            $body = $cleanHt | ConvertTo-Json -Depth 20
            $response = Invoke-RestMethod -Uri "$($Session.BaseUri)/rest/server-profiles" `
                -Method Post -Headers $Session.Headers -Body $body `
                -ContentType "application/json" -SkipCertificateCheck

            $taskMsg = ""
            if ($response.uri -and $response.uri -match "/rest/tasks/") {
                $taskMsg = "`n`nTask-URI: $($response.uri)`nDen Task-Status können Sie im OneView prüfen."
            }

            $lblStatus.Text = "Profil '$profileName' erfolgreich erstellt"
            [System.Windows.Forms.MessageBox]::Show(
                "Server Profile '$profileName' wurde erfolgreich erstellt.$taskMsg",
                "Erfolgreich", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } catch {
            $lblStatus.Text = "Fehler beim Erstellen"
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Erstellen:`n`n$($_.Exception.Message)",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnCreateOV.Enabled = $true
            $btnUpdateOV.Enabled = $true
        }
    })

    # ── Update in OneView ──
    $btnUpdateOV.Add_Click({
        if (-not (& $validateJson)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Das JSON ist ungültig. Bitte zuerst korrigieren.",
                "JSON ungültig", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $obj = $txtJson.Text | ConvertFrom-Json
        $profileName = if ($obj.name) { $obj.name } else { "(unbenannt)" }

        if (-not $obj.uri) {
            [System.Windows.Forms.MessageBox]::Show(
                "Das Profil enthält kein 'uri'-Feld.`n`nEin Update ist nur für bestehende Profile möglich, die ein 'uri'-Feld besitzen.`nLaden Sie das Profil direkt von OneView oder verwenden Sie 'Neu anlegen'.",
                "URI fehlt", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Server Profile '$profileName' in OneView aktualisieren?`n`nURI: $($obj.uri)`n`nAlle Änderungen im JSON werden auf das bestehende Profil angewendet.",
            "Profil aktualisieren",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnCreateOV.Enabled = $false
        $btnUpdateOV.Enabled = $false
        $lblStatus.Text = "Aktualisiere Profil '$profileName'..."

        try {
            $body = $txtJson.Text
            $response = Invoke-RestMethod -Uri "$($Session.BaseUri)$($obj.uri)" `
                -Method Put -Headers $Session.Headers -Body $body `
                -ContentType "application/json" -SkipCertificateCheck

            $taskMsg = ""
            if ($response.uri -and $response.uri -match "/rest/tasks/") {
                $taskMsg = "`n`nTask-URI: $($response.uri)`nDen Task-Status können Sie im OneView prüfen."
            }

            $lblStatus.Text = "Profil '$profileName' erfolgreich aktualisiert"
            [System.Windows.Forms.MessageBox]::Show(
                "Server Profile '$profileName' wurde erfolgreich aktualisiert.$taskMsg",
                "Erfolgreich", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } catch {
            $lblStatus.Text = "Fehler beim Aktualisieren"
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Aktualisieren:`n`n$($_.Exception.Message)",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnCreateOV.Enabled = $true
            $btnUpdateOV.Enabled = $true
        }
    })

    # ── Von OneView laden (Profil auswählen) ──
    $btnLoadOV.Add_Click({
        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $lblStatus.Text = "Lade Profile von OneView..."

        try {
            $profiles = @(Get-ServerProfilesInline -Session $Session)

            if ($profiles.Count -eq 0) {
                $lblStatus.Text = "Keine Profile gefunden"
                [System.Windows.Forms.MessageBox]::Show(
                    "Keine Server Profile auf dieser Appliance gefunden.",
                    "Keine Profile", [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                return
            }

            # Auswahldialog
            $pickDlg = New-Object System.Windows.Forms.Form
            $pickDlg.Text = "Server Profile auswählen"
            $pickDlg.Size = New-Object System.Drawing.Size(500, 450)
            $pickDlg.StartPosition = "CenterParent"
            $pickDlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $pickDlg.MaximizeBox = $false; $pickDlg.MinimizeBox = $false
            $pickDlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

            $lblPick = New-Object System.Windows.Forms.Label
            $lblPick.Text = "Wählen Sie ein Profil zum Laden in den Editor:"
            $lblPick.Location = New-Object System.Drawing.Point(15, 12)
            $lblPick.Size = New-Object System.Drawing.Size(460, 20)
            $pickDlg.Controls.Add($lblPick)

            $lstProfiles = New-Object System.Windows.Forms.ListBox
            $lstProfiles.Location = New-Object System.Drawing.Point(15, 38)
            $lstProfiles.Size = New-Object System.Drawing.Size(455, 320)
            $lstProfiles.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            foreach ($p in ($profiles | Sort-Object -Property name)) {
                $lstProfiles.Items.Add("$($p.name)  [$($p.status)]") | Out-Null
            }
            if ($lstProfiles.Items.Count -gt 0) { $lstProfiles.SelectedIndex = 0 }
            $pickDlg.Controls.Add($lstProfiles)

            $btnPickOK = New-Object System.Windows.Forms.Button
            $btnPickOK.Text = "Laden"
            $btnPickOK.Size = New-Object System.Drawing.Size(80, 30)
            $btnPickOK.Location = New-Object System.Drawing.Point(305, 370)
            $btnPickOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $pickDlg.AcceptButton = $btnPickOK
            $pickDlg.Controls.Add($btnPickOK)

            $btnPickCancel = New-Object System.Windows.Forms.Button
            $btnPickCancel.Text = "Abbrechen"
            $btnPickCancel.Size = New-Object System.Drawing.Size(80, 30)
            $btnPickCancel.Location = New-Object System.Drawing.Point(390, 370)
            $btnPickCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $pickDlg.CancelButton = $btnPickCancel
            $pickDlg.Controls.Add($btnPickCancel)

            # Doppelklick = sofort laden
            $lstProfiles.Add_DoubleClick({
                $pickDlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $pickDlg.Close()
            })

            $pickResult = $pickDlg.ShowDialog($dlg)

            if ($pickResult -eq [System.Windows.Forms.DialogResult]::OK -and $lstProfiles.SelectedIndex -ge 0) {
                $sortedProfiles = $profiles | Sort-Object -Property name
                $selectedProfile = $sortedProfiles[$lstProfiles.SelectedIndex]

                # Profil vollständig per GET /rest/server-profiles/{id} laden
                $fullProfile = Invoke-RestMethod -Uri "$($Session.BaseUri)$($selectedProfile.uri)" `
                    -Method Get -Headers $Session.Headers -SkipCertificateCheck

                $txtJson.Text = $fullProfile | ConvertTo-Json -Depth 20
                $script:jsonEditorFilePath = ""
                $lblStatus.Text = "Geladen von OneView: $($fullProfile.name)"
            }

            $pickDlg.Dispose()

        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Laden von OneView:`n$_",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
            $lblStatus.Text = if ($txtJson.Text) { $lblStatus.Text } else { "Bereit" }
        }
    })

    $dlg.ShowDialog($form) | Out-Null
    $dlg.Dispose()
}

# ============================================================================
#  Server Profile Template Management Dialog
# ============================================================================

function Show-ServerProfileTemplateManageDialog {
    <#
    .SYNOPSIS  Zeigt einen Dialog zum Verwalten von Server Profile Templates (Anzeigen, Bearbeiten, Neu, Löschen).
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [Parameter(Mandatory)][string]$Hostname
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Server Profile Templates verwalten – $Hostname"
    $dlg.Size = New-Object System.Drawing.Size(950, 700)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dlg.MinimumSize = New-Object System.Drawing.Size(800, 550)
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # ── Toolbar ──
    $pnlToolbar = New-Object System.Windows.Forms.Panel
    $pnlToolbar.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlToolbar.Height = 40
    $dlg.Controls.Add($pnlToolbar)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Aktualisieren"
    $btnRefresh.Location = New-Object System.Drawing.Point(10, 6)
    $btnRefresh.Size = New-Object System.Drawing.Size(110, 28)
    $pnlToolbar.Controls.Add($btnRefresh)

    $btnNewSPT = New-Object System.Windows.Forms.Button
    $btnNewSPT.Text = "Aus Profil erstellen"
    $btnNewSPT.Location = New-Object System.Drawing.Point(130, 6)
    $btnNewSPT.Size = New-Object System.Drawing.Size(140, 28)
    $btnNewSPT.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnNewSPT.ForeColor = [System.Drawing.Color]::White
    $btnNewSPT.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlToolbar.Controls.Add($btnNewSPT)

    $btnDeleteSPT = New-Object System.Windows.Forms.Button
    $btnDeleteSPT.Text = "Löschen"
    $btnDeleteSPT.Location = New-Object System.Drawing.Point(280, 6)
    $btnDeleteSPT.Size = New-Object System.Drawing.Size(80, 28)
    $btnDeleteSPT.BackColor = [System.Drawing.Color]::FromArgb(200, 50, 50)
    $btnDeleteSPT.ForeColor = [System.Drawing.Color]::White
    $btnDeleteSPT.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlToolbar.Controls.Add($btnDeleteSPT)

    $btnExportOne = New-Object System.Windows.Forms.Button
    $btnExportOne.Text = "Als JSON exportieren"
    $btnExportOne.Location = New-Object System.Drawing.Point(370, 6)
    $btnExportOne.Size = New-Object System.Drawing.Size(140, 28)
    $pnlToolbar.Controls.Add($btnExportOne)

    $btnJsonEditor = New-Object System.Windows.Forms.Button
    $btnJsonEditor.Text = "JSON Editor"
    $btnJsonEditor.Location = New-Object System.Drawing.Point(520, 6)
    $btnJsonEditor.Size = New-Object System.Drawing.Size(110, 28)
    $btnJsonEditor.BackColor = [System.Drawing.Color]::FromArgb(50, 120, 80)
    $btnJsonEditor.ForeColor = [System.Drawing.Color]::White
    $btnJsonEditor.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlToolbar.Controls.Add($btnJsonEditor)

    $btnNewProfile = New-Object System.Windows.Forms.Button
    $btnNewProfile.Text = "Profil erzeugen"
    $btnNewProfile.Location = New-Object System.Drawing.Point(640, 6)
    $btnNewProfile.Size = New-Object System.Drawing.Size(120, 28)
    $btnNewProfile.BackColor = [System.Drawing.Color]::FromArgb(100, 60, 160)
    $btnNewProfile.ForeColor = [System.Drawing.Color]::White
    $btnNewProfile.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlToolbar.Controls.Add($btnNewProfile)

    # ── SplitContainer ──
    $splitMain = New-Object System.Windows.Forms.SplitContainer
    $splitMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitMain.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $splitMain.SplitterDistance = 350
    $dlg.Controls.Add($splitMain)
    $splitMain.BringToFront()

    # ── Liste (links) ──
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Dock = [System.Windows.Forms.DockStyle]::Fill
    $dgv.ReadOnly = $true
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $dgv.MultiSelect = $false
    $dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $dgv.BackgroundColor = [System.Drawing.Color]::White
    $dgv.RowHeadersVisible = $false
    $splitMain.Panel1.Controls.Add($dgv)

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = "Name"; $colName.Name = "Name"; $colName.FillWeight = 40
    $colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStatus.HeaderText = "Status"; $colStatus.Name = "Status"; $colStatus.FillWeight = 15
    $colHWType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colHWType.HeaderText = "Server HW Type"; $colHWType.Name = "HWType"; $colHWType.FillWeight = 30
    $colEncGroup = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colEncGroup.HeaderText = "Enclosure Group"; $colEncGroup.Name = "EncGroup"; $colEncGroup.FillWeight = 25
    $dgv.Columns.AddRange(@($colName, $colStatus, $colHWType, $colEncGroup))

    # ── Details (rechts) ──
    $txtDetails = New-Object System.Windows.Forms.RichTextBox
    $txtDetails.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtDetails.ReadOnly = $true
    $txtDetails.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $txtDetails.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $txtDetails.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtDetails.WordWrap = $false
    $splitMain.Panel2.Controls.Add($txtDetails)

    # ── Daten ──
    $script:sptTemplates    = @()
    $script:sptHWTypeMap    = @{}
    $script:sptEncGroupMap  = @{}

    $loadTemplates = {
        $dgv.Rows.Clear()
        $txtDetails.Clear()
        try {
            $script:sptTemplates = @(Get-ServerProfileTemplatesInline -Session $Session)

            # Server Hardware Types laden für Anzeige
            $script:sptHWTypeMap = @{}
            try {
                $start = 0; $pageSize = 100
                do {
                    $uri = "$($Session.BaseUri)/rest/server-hardware-types?start=$start&count=$pageSize"
                    $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck
                    if ($response.members) {
                        foreach ($hwt in $response.members) { $script:sptHWTypeMap[$hwt.uri] = $hwt.name }
                    }
                    $total = $response.total
                    $start += $response.members.Count
                } while ($script:sptHWTypeMap.Count -lt $total)
            } catch { }

            # Enclosure Groups laden für Anzeige
            $script:sptEncGroupMap = @{}
            try {
                $uri = "$($Session.BaseUri)/rest/enclosure-groups?start=0&count=256"
                $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $Session.Headers -SkipCertificateCheck
                if ($response.members) {
                    foreach ($eg in $response.members) { $script:sptEncGroupMap[$eg.uri] = $eg.name }
                }
            } catch { }

            foreach ($t in $script:sptTemplates) {
                $hwTypeName  = if ($t.serverHardwareTypeUri -and $script:sptHWTypeMap.ContainsKey($t.serverHardwareTypeUri)) { $script:sptHWTypeMap[$t.serverHardwareTypeUri] } else { "(ohne)" }
                $encGrpName  = if ($t.enclosureGroupUri -and $script:sptEncGroupMap.ContainsKey($t.enclosureGroupUri)) { $script:sptEncGroupMap[$t.enclosureGroupUri] } else { "(ohne)" }
                $dgv.Rows.Add($t.name, $t.status, $hwTypeName, $encGrpName) | Out-Null
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Laden: $_", "Fehler",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    }

    & $loadTemplates

    # ── Selection Changed ──
    $dgv.Add_SelectionChanged({
        if ($dgv.SelectedRows.Count -eq 0) { return }
        $idx = $dgv.SelectedRows[0].Index
        if ($idx -ge 0 -and $idx -lt $script:sptTemplates.Count) {
            $selected = $script:sptTemplates[$idx]
            $json = $selected | ConvertTo-Json -Depth 10
            $txtDetails.Clear()
            $txtDetails.Text = $json
        }
    })

    # ── Refresh ──
    $btnRefresh.Add_Click({ & $loadTemplates })

    # ── Aus Profil erstellen ──
    $btnNewSPT.Add_Click({
        # Profile laden und Auswahldialog zeigen
        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $profiles = @(Get-ServerProfilesInline -Session $Session)
            if ($profiles.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Keine Server Profile gefunden.", "Keine Profile",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                return
            }

            $pickDlg = New-Object System.Windows.Forms.Form
            $pickDlg.Text = "Server Profile als Vorlage auswählen"
            $pickDlg.Size = New-Object System.Drawing.Size(500, 520)
            $pickDlg.StartPosition = "CenterParent"
            $pickDlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $pickDlg.MaximizeBox = $false; $pickDlg.MinimizeBox = $false
            $pickDlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

            $lblPick = New-Object System.Windows.Forms.Label
            $lblPick.Text = "Profil auswählen, aus dem ein Template erstellt werden soll:"
            $lblPick.Location = New-Object System.Drawing.Point(15, 12)
            $lblPick.Size = New-Object System.Drawing.Size(460, 20)
            $pickDlg.Controls.Add($lblPick)

            $lstProfiles = New-Object System.Windows.Forms.ListBox
            $lstProfiles.Location = New-Object System.Drawing.Point(15, 38)
            $lstProfiles.Size = New-Object System.Drawing.Size(455, 280)
            $lstProfiles.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            foreach ($p in ($profiles | Sort-Object -Property name)) {
                $lstProfiles.Items.Add("$($p.name)  [$($p.status)]") | Out-Null
            }
            if ($lstProfiles.Items.Count -gt 0) { $lstProfiles.SelectedIndex = 0 }
            $pickDlg.Controls.Add($lstProfiles)

            $lblName = New-Object System.Windows.Forms.Label
            $lblName.Text = "Template-Name:"
            $lblName.Location = New-Object System.Drawing.Point(15, 330)
            $lblName.Size = New-Object System.Drawing.Size(110, 20)
            $pickDlg.Controls.Add($lblName)

            $txtTplName = New-Object System.Windows.Forms.TextBox
            $txtTplName.Location = New-Object System.Drawing.Point(130, 328)
            $txtTplName.Size = New-Object System.Drawing.Size(340, 23)
            $pickDlg.Controls.Add($txtTplName)

            $lblDesc = New-Object System.Windows.Forms.Label
            $lblDesc.Text = "Beschreibung:"
            $lblDesc.Location = New-Object System.Drawing.Point(15, 362)
            $lblDesc.Size = New-Object System.Drawing.Size(110, 20)
            $pickDlg.Controls.Add($lblDesc)

            $txtTplDesc = New-Object System.Windows.Forms.TextBox
            $txtTplDesc.Location = New-Object System.Drawing.Point(130, 360)
            $txtTplDesc.Size = New-Object System.Drawing.Size(340, 23)
            $pickDlg.Controls.Add($txtTplDesc)

            # Name automatisch setzen wenn Profil gewählt wird
            $lstProfiles.Add_SelectedIndexChanged({
                if ($lstProfiles.SelectedIndex -ge 0) {
                    $sortedP = $profiles | Sort-Object -Property name
                    $selP = $sortedP[$lstProfiles.SelectedIndex]
                    if ([string]::IsNullOrWhiteSpace($txtTplName.Text) -or $txtTplName.Text -match '_Template$') {
                        $txtTplName.Text = "$($selP.name)_Template"
                    }
                }
            })
            # Initial setzen
            if ($lstProfiles.SelectedIndex -ge 0) {
                $sortedP = $profiles | Sort-Object -Property name
                $txtTplName.Text = "$($sortedP[0].name)_Template"
            }

            $btnPickOK = New-Object System.Windows.Forms.Button
            $btnPickOK.Text = "Erstellen"
            $btnPickOK.Size = New-Object System.Drawing.Size(90, 30)
            $btnPickOK.Location = New-Object System.Drawing.Point(295, 400)
            $btnPickOK.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
            $btnPickOK.ForeColor = [System.Drawing.Color]::White
            $btnPickOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $btnPickOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $pickDlg.AcceptButton = $btnPickOK
            $pickDlg.Controls.Add($btnPickOK)

            $btnPickCancel = New-Object System.Windows.Forms.Button
            $btnPickCancel.Text = "Abbrechen"
            $btnPickCancel.Size = New-Object System.Drawing.Size(80, 30)
            $btnPickCancel.Location = New-Object System.Drawing.Point(390, 400)
            $btnPickCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $pickDlg.CancelButton = $btnPickCancel
            $pickDlg.Controls.Add($btnPickCancel)

            $pickResult = $pickDlg.ShowDialog($dlg)

            if ($pickResult -eq [System.Windows.Forms.DialogResult]::OK -and $lstProfiles.SelectedIndex -ge 0) {
                $sortedP = $profiles | Sort-Object -Property name
                $selectedProfile = $sortedP[$lstProfiles.SelectedIndex]
                $tplName = $txtTplName.Text.Trim()
                $tplDesc = $txtTplDesc.Text.Trim()

                if ([string]::IsNullOrWhiteSpace($tplName)) {
                    [System.Windows.Forms.MessageBox]::Show("Bitte einen Template-Namen eingeben.", "Pflichtfeld",
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                    $pickDlg.Dispose()
                    return
                }

                try {
                    New-ServerProfileTemplateFromProfileInline -Session $Session `
                        -ServerProfileUri $selectedProfile.uri `
                        -TemplateName $tplName -Description $tplDesc

                    [System.Windows.Forms.MessageBox]::Show(
                        "Template '$tplName' wurde aus Profil '$($selectedProfile.name)' erstellt.",
                        "Erfolgreich",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                    & $loadTemplates
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Fehler beim Erstellen: $_", "Fehler",
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                }
            }
            $pickDlg.Dispose()

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler: $_", "Fehler",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    # ── Löschen ──
    $btnDeleteSPT.Add_Click({
        if ($dgv.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte ein Template auswählen.", "Hinweis",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $idx = $dgv.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:sptTemplates.Count) { return }
        $template = $script:sptTemplates[$idx]

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Server Profile Template '$($template.name)' wirklich löschen?`n`nDiese Aktion kann nicht rückgängig gemacht werden!",
            "Template löschen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        try {
            Remove-ServerProfileTemplateInline -Session $Session -TemplateUri $template.uri
            [System.Windows.Forms.MessageBox]::Show("Template '$($template.name)' wurde gelöscht.", "Erfolgreich",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            & $loadTemplates
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Löschen: $_", "Fehler",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    # ── Einzelexport ──
    $btnExportOne.Add_Click({
        if ($dgv.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte ein Template auswählen.", "Hinweis",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $idx = $dgv.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:sptTemplates.Count) { return }
        $template = $script:sptTemplates[$idx]

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "JSON Dateien (*.json)|*.json"
        $sfd.FileName = ($template.name -replace '[\\/:*?\"<>|\s]', '_') + ".json"
        $sfd.Title = "Template als JSON exportieren"
        if ($sfd.ShowDialog($dlg) -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $template | ConvertTo-Json -Depth 20 | Set-Content -Path $sfd.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exportiert nach:`n$($sfd.FileName)", "Export erfolgreich",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    })

    # ── JSON Editor ──
    $btnJsonEditor.Add_Click({
        $initialJson = ""
        if ($dgv.SelectedRows.Count -gt 0) {
            $idx = $dgv.SelectedRows[0].Index
            if ($idx -ge 0 -and $idx -lt $script:sptTemplates.Count) {
                $initialJson = $script:sptTemplates[$idx] | ConvertTo-Json -Depth 20
            }
        }
        Show-ServerProfileTemplateJsonEditor -Session $Session -Hostname $Hostname -InitialJson $initialJson
        & $loadTemplates
    })

    # ── Profil aus Template erzeugen ──
    $btnNewProfile.Add_Click({
        if ($dgv.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte ein Template auswählen.", "Hinweis",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $idx = $dgv.SelectedRows[0].Index
        if ($idx -lt 0 -or $idx -ge $script:sptTemplates.Count) { return }
        $template = $script:sptTemplates[$idx]

        # Profilname-Dialog
        $nameDlg = New-Object System.Windows.Forms.Form
        $nameDlg.Text = "Neues Server Profile aus Template"
        $nameDlg.Size = New-Object System.Drawing.Size(450, 250)
        $nameDlg.StartPosition = "CenterParent"
        $nameDlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $nameDlg.MaximizeBox = $false; $nameDlg.MinimizeBox = $false
        $nameDlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

        $lblTpl = New-Object System.Windows.Forms.Label
        $lblTpl.Text = "Template: $($template.name)"
        $lblTpl.Location = New-Object System.Drawing.Point(15, 15)
        $lblTpl.Size = New-Object System.Drawing.Size(400, 20)
        $lblTpl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $nameDlg.Controls.Add($lblTpl)

        $lblN = New-Object System.Windows.Forms.Label
        $lblN.Text = "Profil-Name:"
        $lblN.Location = New-Object System.Drawing.Point(15, 50)
        $lblN.Size = New-Object System.Drawing.Size(100, 20)
        $nameDlg.Controls.Add($lblN)
        $txtProfName = New-Object System.Windows.Forms.TextBox
        $txtProfName.Location = New-Object System.Drawing.Point(120, 48)
        $txtProfName.Size = New-Object System.Drawing.Size(290, 23)
        $nameDlg.Controls.Add($txtProfName)

        $lblD = New-Object System.Windows.Forms.Label
        $lblD.Text = "Beschreibung:"
        $lblD.Location = New-Object System.Drawing.Point(15, 82)
        $lblD.Size = New-Object System.Drawing.Size(100, 20)
        $nameDlg.Controls.Add($lblD)
        $txtProfDesc = New-Object System.Windows.Forms.TextBox
        $txtProfDesc.Location = New-Object System.Drawing.Point(120, 80)
        $txtProfDesc.Size = New-Object System.Drawing.Size(290, 23)
        $nameDlg.Controls.Add($txtProfDesc)

        # Server Hardware auswählen
        $lblHW = New-Object System.Windows.Forms.Label
        $lblHW.Text = "Server Hardware:"
        $lblHW.Location = New-Object System.Drawing.Point(15, 114)
        $lblHW.Size = New-Object System.Drawing.Size(100, 20)
        $nameDlg.Controls.Add($lblHW)
        $cmbHW = New-Object System.Windows.Forms.ComboBox
        $cmbHW.Location = New-Object System.Drawing.Point(120, 112)
        $cmbHW.Size = New-Object System.Drawing.Size(290, 23)
        $cmbHW.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        $cmbHW.Items.Add("(Nicht zugewiesen)") | Out-Null
        try {
            $hwList = @(Get-ServerHardwareInline -Session $Session)
            foreach ($h in ($hwList | Sort-Object -Property name)) {
                $tag = if ($h.state -eq "NoProfileApplied") { " [verfügbar]" } else { " [belegt]" }
                $cmbHW.Items.Add("$($h.name)$tag") | Out-Null
            }
        } catch { }
        $cmbHW.SelectedIndex = 0
        $nameDlg.Controls.Add($cmbHW)

        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "Erstellen"
        $btnOK.Size = New-Object System.Drawing.Size(90, 30)
        $btnOK.Location = New-Object System.Drawing.Point(230, 160)
        $btnOK.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        $btnOK.ForeColor = [System.Drawing.Color]::White
        $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $nameDlg.AcceptButton = $btnOK
        $nameDlg.Controls.Add($btnOK)

        $btnCn = New-Object System.Windows.Forms.Button
        $btnCn.Text = "Abbrechen"
        $btnCn.Size = New-Object System.Drawing.Size(80, 30)
        $btnCn.Location = New-Object System.Drawing.Point(325, 160)
        $btnCn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $nameDlg.CancelButton = $btnCn
        $nameDlg.Controls.Add($btnCn)

        $nameResult = $nameDlg.ShowDialog($dlg)

        if ($nameResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $profName = $txtProfName.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($profName)) {
                [System.Windows.Forms.MessageBox]::Show("Bitte einen Profil-Namen eingeben.", "Pflichtfeld",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                $nameDlg.Dispose()
                return
            }

            $hwUri = ""
            if ($cmbHW.SelectedIndex -gt 0) {
                $sortedHW = $hwList | Sort-Object -Property name
                $hwUri = $sortedHW[$cmbHW.SelectedIndex - 1].uri
            }

            try {
                New-ServerProfileFromTemplateInline -Session $Session `
                    -TemplateUri $template.uri -ProfileName $profName `
                    -Description $txtProfDesc.Text.Trim() -ServerHardwareUri $hwUri

                [System.Windows.Forms.MessageBox]::Show(
                    "Server Profile '$profName' wurde aus Template '$($template.name)' erstellt.",
                    "Erfolgreich",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Erstellen: $_", "Fehler",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            }
        }
        $nameDlg.Dispose()
    })

    $dlg.ShowDialog($form) | Out-Null
    $dlg.Dispose()
}

# ============================================================================
#  JSON Editor Dialog – Server Profile Templates
# ============================================================================
function Show-ServerProfileTemplateJsonEditor {
    <#
    .SYNOPSIS  JSON-Editor zum vollständigen Bearbeiten, Erstellen und Updaten von Server Profile Templates.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Session,
        [string]$Hostname = "",
        [string]$InitialJson = "",
        [string]$InitialFilePath = ""
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Server Profile Template – JSON Editor$(if ($Hostname) { " – $Hostname" })"
    $dlg.Size = New-Object System.Drawing.Size(950, 750)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dlg.MinimumSize = New-Object System.Drawing.Size(700, 500)
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # ── Toolbar oben ──
    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 44
    $dlg.Controls.Add($pnlTop)

    $btnLoadFile = New-Object System.Windows.Forms.Button
    $btnLoadFile.Text = "JSON laden"
    $btnLoadFile.Location = New-Object System.Drawing.Point(10, 8)
    $btnLoadFile.Size = New-Object System.Drawing.Size(100, 28)
    $pnlTop.Controls.Add($btnLoadFile)

    $btnSaveFile = New-Object System.Windows.Forms.Button
    $btnSaveFile.Text = "JSON speichern"
    $btnSaveFile.Location = New-Object System.Drawing.Point(118, 8)
    $btnSaveFile.Size = New-Object System.Drawing.Size(110, 28)
    $pnlTop.Controls.Add($btnSaveFile)

    $btnValidate = New-Object System.Windows.Forms.Button
    $btnValidate.Text = "Validieren"
    $btnValidate.Location = New-Object System.Drawing.Point(236, 8)
    $btnValidate.Size = New-Object System.Drawing.Size(90, 28)
    $pnlTop.Controls.Add($btnValidate)

    $btnFormat = New-Object System.Windows.Forms.Button
    $btnFormat.Text = "Formatieren"
    $btnFormat.Location = New-Object System.Drawing.Point(334, 8)
    $btnFormat.Size = New-Object System.Drawing.Size(95, 28)
    $pnlTop.Controls.Add($btnFormat)

    $sep = New-Object System.Windows.Forms.Label
    $sep.Text = "|"
    $sep.Location = New-Object System.Drawing.Point(440, 12)
    $sep.Size = New-Object System.Drawing.Size(10, 20)
    $sep.ForeColor = [System.Drawing.Color]::Gray
    $pnlTop.Controls.Add($sep)

    $btnCreateOV = New-Object System.Windows.Forms.Button
    $btnCreateOV.Text = "Neu anlegen (OneView)"
    $btnCreateOV.Location = New-Object System.Drawing.Point(458, 8)
    $btnCreateOV.Size = New-Object System.Drawing.Size(155, 28)
    $btnCreateOV.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnCreateOV.ForeColor = [System.Drawing.Color]::White
    $btnCreateOV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlTop.Controls.Add($btnCreateOV)

    $btnUpdateOV = New-Object System.Windows.Forms.Button
    $btnUpdateOV.Text = "Update (OneView)"
    $btnUpdateOV.Location = New-Object System.Drawing.Point(621, 8)
    $btnUpdateOV.Size = New-Object System.Drawing.Size(135, 28)
    $btnUpdateOV.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 80)
    $btnUpdateOV.ForeColor = [System.Drawing.Color]::White
    $btnUpdateOV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlTop.Controls.Add($btnUpdateOV)

    $btnLoadOV = New-Object System.Windows.Forms.Button
    $btnLoadOV.Text = "Von OneView laden"
    $btnLoadOV.Location = New-Object System.Drawing.Point(764, 8)
    $btnLoadOV.Size = New-Object System.Drawing.Size(135, 28)
    $btnLoadOV.BackColor = [System.Drawing.Color]::FromArgb(100, 60, 160)
    $btnLoadOV.ForeColor = [System.Drawing.Color]::White
    $btnLoadOV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlTop.Controls.Add($btnLoadOV)

    # ── Statusleiste ──
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $lblStatus.Height = 24
    $lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblStatus.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $lblStatus.Text = "Bereit"
    $dlg.Controls.Add($lblStatus)

    # ── JSON Text Editor ──
    $txtJson = New-Object System.Windows.Forms.RichTextBox
    $txtJson.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtJson.Font = New-Object System.Drawing.Font("Consolas", 10)
    $txtJson.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $txtJson.ForeColor = [System.Drawing.Color]::FromArgb(212, 212, 212)
    $txtJson.WordWrap = $false
    $txtJson.AcceptsTab = $true
    $txtJson.ShortcutsEnabled = $true
    $txtJson.DetectUrls = $false
    $dlg.Controls.Add($txtJson)
    $txtJson.BringToFront()

    $script:jsonEditorFilePath = $InitialFilePath

    if ($InitialJson) {
        $txtJson.Text = $InitialJson
        $lblStatus.Text = if ($InitialFilePath) { "Geladen: $InitialFilePath" } else { "Template aus OneView geladen" }
    }

    # ── Hilfsfunktion: JSON validieren ──
    $validateJson = {
        try {
            $null = $txtJson.Text | ConvertFrom-Json -ErrorAction Stop
            return $true
        } catch {
            return $false
        }
    }

    # ── Hilfsfunktion: Read-Only-Felder entfernen (für Create) ──
    $removeReadOnlyFields = {
        param([hashtable]$ht)
        $readOnlyFields = @(
            "uri", "eTag", "created", "modified", "uuid",
            "taskUri", "stateReason", "refreshState",
            "inProgress", "scopesUri"
        )
        foreach ($f in $readOnlyFields) { $ht.Remove($f) }
        return $ht
    }

    # ── JSON laden ──
    $btnLoadFile.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "JSON Dateien (*.json)|*.json|Alle Dateien (*.*)|*.*"
        $ofd.Title = "Server Profile Template JSON laden"
        if ($ofd.ShowDialog($dlg) -ne [System.Windows.Forms.DialogResult]::OK) { return }

        try {
            $content = Get-Content -Path $ofd.FileName -Raw -Encoding UTF8
            $null = $content | ConvertFrom-Json -ErrorAction Stop
            $obj = $content | ConvertFrom-Json
            $txtJson.Text = $obj | ConvertTo-Json -Depth 20
            $script:jsonEditorFilePath = $ofd.FileName
            $lblStatus.Text = "Geladen: $($ofd.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Laden der Datei:`n$_",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    # ── JSON speichern ──
    $btnSaveFile.Add_Click({
        if (-not (& $validateJson)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Das JSON ist ungültig. Bitte korrigieren Sie die Syntax.",
                "JSON ungültig", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "JSON Dateien (*.json)|*.json"
        $sfd.Title = "Server Profile Template JSON speichern"
        if ($script:jsonEditorFilePath) {
            $sfd.FileName = [System.IO.Path]::GetFileName($script:jsonEditorFilePath)
            $sfd.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:jsonEditorFilePath)
        } else {
            try {
                $obj = $txtJson.Text | ConvertFrom-Json
                if ($obj.name) { $sfd.FileName = ($obj.name -replace '[\\/:*?\"<>|\s]', '_') + ".json" }
            } catch {}
        }
        if ($sfd.ShowDialog($dlg) -ne [System.Windows.Forms.DialogResult]::OK) { return }

        try {
            $obj = $txtJson.Text | ConvertFrom-Json
            $obj | ConvertTo-Json -Depth 20 | Set-Content -Path $sfd.FileName -Encoding UTF8
            $script:jsonEditorFilePath = $sfd.FileName
            $lblStatus.Text = "Gespeichert: $($sfd.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Speichern:`n$_",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    # ── Validieren ──
    $btnValidate.Add_Click({
        if (& $validateJson) {
            $obj = $txtJson.Text | ConvertFrom-Json
            $fields = ($obj | Get-Member -MemberType NoteProperty).Count
            $lblStatus.Text = "JSON ist gültig ($fields Felder auf Root-Ebene)"
            [System.Windows.Forms.MessageBox]::Show(
                "JSON ist gültig.`n`nRoot-Felder: $fields`nName: $($obj.name)",
                "Validierung OK", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } else {
            $lblStatus.Text = "JSON ist UNGÜLTIG!"
            try {
                $null = $txtJson.Text | ConvertFrom-Json -ErrorAction Stop
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "JSON ist ungültig:`n`n$($_.Exception.Message)",
                    "Validierung fehlgeschlagen", [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            }
        }
    })

    # ── Formatieren ──
    $btnFormat.Add_Click({
        if (-not (& $validateJson)) {
            [System.Windows.Forms.MessageBox]::Show(
                "JSON kann nicht formatiert werden – Syntax ungültig.",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        $obj = $txtJson.Text | ConvertFrom-Json
        $txtJson.Text = $obj | ConvertTo-Json -Depth 20
        $lblStatus.Text = "JSON formatiert"
    })

    # ── Neu anlegen in OneView ──
    $btnCreateOV.Add_Click({
        if (-not (& $validateJson)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Das JSON ist ungültig. Bitte zuerst korrigieren.",
                "JSON ungültig", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $obj = $txtJson.Text | ConvertFrom-Json
        $templateName = if ($obj.name) { $obj.name } else { "(unbenannt)" }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Server Profile Template '$templateName' als neues Template in OneView anlegen?`n`nRead-Only-Felder (uri, eTag, etc.) werden automatisch entfernt.",
            "Template erstellen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnCreateOV.Enabled = $false
        $btnUpdateOV.Enabled = $false
        $lblStatus.Text = "Erstelle Template '$templateName'..."

        try {
            $ht = $txtJson.Text | ConvertFrom-Json -AsHashtable
            $cleanHt = & $removeReadOnlyFields $ht

            $body = $cleanHt | ConvertTo-Json -Depth 20
            $response = Invoke-RestMethod -Uri "$($Session.BaseUri)/rest/server-profile-templates" `
                -Method Post -Headers $Session.Headers -Body $body `
                -ContentType "application/json" -SkipCertificateCheck

            $lblStatus.Text = "Template '$templateName' erfolgreich erstellt"
            [System.Windows.Forms.MessageBox]::Show(
                "Server Profile Template '$templateName' wurde erfolgreich erstellt.",
                "Erfolgreich", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } catch {
            $lblStatus.Text = "Fehler beim Erstellen"
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Erstellen:`n`n$($_.Exception.Message)",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnCreateOV.Enabled = $true
            $btnUpdateOV.Enabled = $true
        }
    })

    # ── Update in OneView ──
    $btnUpdateOV.Add_Click({
        if (-not (& $validateJson)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Das JSON ist ungültig. Bitte zuerst korrigieren.",
                "JSON ungültig", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $obj = $txtJson.Text | ConvertFrom-Json
        $templateName = if ($obj.name) { $obj.name } else { "(unbenannt)" }

        if (-not $obj.uri) {
            [System.Windows.Forms.MessageBox]::Show(
                "Das Template enthält kein 'uri'-Feld.`n`nEin Update ist nur für bestehende Templates möglich, die ein 'uri'-Feld besitzen.`nLaden Sie das Template direkt von OneView oder verwenden Sie 'Neu anlegen'.",
                "URI fehlt", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Server Profile Template '$templateName' in OneView aktualisieren?`n`nURI: $($obj.uri)`n`nAlle Änderungen im JSON werden auf das bestehende Template angewendet.",
            "Template aktualisieren",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnCreateOV.Enabled = $false
        $btnUpdateOV.Enabled = $false
        $lblStatus.Text = "Aktualisiere Template '$templateName'..."

        try {
            $body = $txtJson.Text
            $response = Invoke-RestMethod -Uri "$($Session.BaseUri)$($obj.uri)" `
                -Method Put -Headers $Session.Headers -Body $body `
                -ContentType "application/json" -SkipCertificateCheck

            $lblStatus.Text = "Template '$templateName' erfolgreich aktualisiert"
            [System.Windows.Forms.MessageBox]::Show(
                "Server Profile Template '$templateName' wurde erfolgreich aktualisiert.",
                "Erfolgreich", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } catch {
            $lblStatus.Text = "Fehler beim Aktualisieren"
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Aktualisieren:`n`n$($_.Exception.Message)",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnCreateOV.Enabled = $true
            $btnUpdateOV.Enabled = $true
        }
    })

    # ── Von OneView laden (Template auswählen) ──
    $btnLoadOV.Add_Click({
        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $lblStatus.Text = "Lade Templates von OneView..."

        try {
            $templates = @(Get-ServerProfileTemplatesInline -Session $Session)

            if ($templates.Count -eq 0) {
                $lblStatus.Text = "Keine Templates gefunden"
                [System.Windows.Forms.MessageBox]::Show(
                    "Keine Server Profile Templates auf dieser Appliance gefunden.",
                    "Keine Templates", [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                return
            }

            $pickDlg = New-Object System.Windows.Forms.Form
            $pickDlg.Text = "Server Profile Template auswählen"
            $pickDlg.Size = New-Object System.Drawing.Size(500, 450)
            $pickDlg.StartPosition = "CenterParent"
            $pickDlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $pickDlg.MaximizeBox = $false; $pickDlg.MinimizeBox = $false
            $pickDlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

            $lblPick = New-Object System.Windows.Forms.Label
            $lblPick.Text = "Wählen Sie ein Template zum Laden in den Editor:"
            $lblPick.Location = New-Object System.Drawing.Point(15, 12)
            $lblPick.Size = New-Object System.Drawing.Size(460, 20)
            $pickDlg.Controls.Add($lblPick)

            $lstTemplates = New-Object System.Windows.Forms.ListBox
            $lstTemplates.Location = New-Object System.Drawing.Point(15, 38)
            $lstTemplates.Size = New-Object System.Drawing.Size(455, 320)
            $lstTemplates.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            foreach ($t in ($templates | Sort-Object -Property name)) {
                $lstTemplates.Items.Add("$($t.name)  [$($t.status)]") | Out-Null
            }
            if ($lstTemplates.Items.Count -gt 0) { $lstTemplates.SelectedIndex = 0 }
            $pickDlg.Controls.Add($lstTemplates)

            $btnPickOK = New-Object System.Windows.Forms.Button
            $btnPickOK.Text = "Laden"
            $btnPickOK.Size = New-Object System.Drawing.Size(80, 30)
            $btnPickOK.Location = New-Object System.Drawing.Point(305, 370)
            $btnPickOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $pickDlg.AcceptButton = $btnPickOK
            $pickDlg.Controls.Add($btnPickOK)

            $btnPickCancel = New-Object System.Windows.Forms.Button
            $btnPickCancel.Text = "Abbrechen"
            $btnPickCancel.Size = New-Object System.Drawing.Size(80, 30)
            $btnPickCancel.Location = New-Object System.Drawing.Point(390, 370)
            $btnPickCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $pickDlg.CancelButton = $btnPickCancel
            $pickDlg.Controls.Add($btnPickCancel)

            $lstTemplates.Add_DoubleClick({
                $pickDlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $pickDlg.Close()
            })

            $pickResult = $pickDlg.ShowDialog($dlg)

            if ($pickResult -eq [System.Windows.Forms.DialogResult]::OK -and $lstTemplates.SelectedIndex -ge 0) {
                $sortedTemplates = $templates | Sort-Object -Property name
                $selectedTemplate = $sortedTemplates[$lstTemplates.SelectedIndex]

                $fullTemplate = Invoke-RestMethod -Uri "$($Session.BaseUri)$($selectedTemplate.uri)" `
                    -Method Get -Headers $Session.Headers -SkipCertificateCheck

                $txtJson.Text = $fullTemplate | ConvertTo-Json -Depth 20
                $script:jsonEditorFilePath = ""
                $lblStatus.Text = "Geladen von OneView: $($fullTemplate.name)"
            }

            $pickDlg.Dispose()

        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Laden von OneView:`n$_",
                "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
            $lblStatus.Text = if ($txtJson.Text) { $lblStatus.Text } else { "Bereit" }
        }
    })

    $dlg.ShowDialog($form) | Out-Null
    $dlg.Dispose()
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
    $btnSPExport.Enabled = $false
    $btnSPImport.Enabled = $false
    $btnSPManage.Enabled = $false
    $btnSPJsonEdit.Enabled = $false
    $btnSPTExport.Enabled = $false
    $btnSPTImport.Enabled = $false
    $btnSPTManage.Enabled = $false
    $btnSPTJsonEdit.Enabled = $false
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
        $btnSPExport.Enabled = $true
        $btnSPImport.Enabled = $true
        $btnSPManage.Enabled = $true
        $btnSPJsonEdit.Enabled = $true
        $btnSPTExport.Enabled = $true
        $btnSPTImport.Enabled = $true
        $btnSPTManage.Enabled = $true
        $btnSPTJsonEdit.Enabled = $true
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
        $btnSPExport.Enabled = $true
        $btnSPImport.Enabled = $true
        $btnSPManage.Enabled = $true
        $btnSPJsonEdit.Enabled = $true
        $btnSPTExport.Enabled = $true
        $btnSPTImport.Enabled = $true
        $btnSPTManage.Enabled = $true
        $btnSPTJsonEdit.Enabled = $true
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
    $btnSPExport.Enabled = $false
    $btnSPImport.Enabled = $false
    $btnSPManage.Enabled = $false
    $btnSPJsonEdit.Enabled = $false
    $btnSPTExport.Enabled = $false
    $btnSPTImport.Enabled = $false
    $btnSPTManage.Enabled = $false
    $btnSPTJsonEdit.Enabled = $false
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
        $btnSPExport.Enabled = $true
        $btnSPImport.Enabled = $true
        $btnSPManage.Enabled = $true
        $btnSPJsonEdit.Enabled = $true
        $btnSPTExport.Enabled = $true
        $btnSPTImport.Enabled = $true
        $btnSPTManage.Enabled = $true
        $btnSPTJsonEdit.Enabled = $true
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
    $btnSPExport.Enabled = $false
    $btnSPImport.Enabled = $false
    $btnSPManage.Enabled = $false
    $btnSPJsonEdit.Enabled = $false
    $btnSPTExport.Enabled = $false
    $btnSPTImport.Enabled = $false
    $btnSPTManage.Enabled = $false
    $btnSPTJsonEdit.Enabled = $false
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
        $btnSPExport.Enabled = $true
        $btnSPImport.Enabled = $true
        $btnSPManage.Enabled = $true
        $btnSPJsonEdit.Enabled = $true
        $btnSPTExport.Enabled = $true
        $btnSPTImport.Enabled = $true
        $btnSPTManage.Enabled = $true
        $btnSPTJsonEdit.Enabled = $true
    }
})

# ============================================================================
#  Button-Event: Server Profile exportieren (Multi-Appliance)
# ============================================================================
$btnSPExport.Add_Click({
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

    $selectedAppliances = Show-ApplianceSelectionDialog -Appliances $appliances -Title "SP Export – Appliances auswählen"
    if (-not $selectedAppliances -or $selectedAppliances.Count -eq 0) { return }

    $applianceList = $selectedAppliances -join "`n"
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Server Profiles von $($selectedAppliances.Count) Appliance(s) exportieren?`n`n$applianceList",
        "SP Export bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnSPExport.Enabled = $false
    $btnSPImport.Enabled = $false
    $btnSPManage.Enabled = $false
    $btnSPJsonEdit.Enabled = $false
    $btnSPTExport.Enabled = $false
    $btnSPTImport.Enabled = $false
    $btnSPTManage.Enabled = $false
    $btnSPTJsonEdit.Enabled = $false
    $rtbLog.Clear()

    try {
        $backupDir = Join-Path $scriptDir "Backups" ("SP_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
        New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
        Write-GUILog "Backup-Verzeichnis: $backupDir" -Color ([System.Drawing.Color]::Cyan)

        $logsDir = Join-Path $scriptDir "Logs"
        if (-not (Test-Path $logsDir)) { New-Item -Path $logsDir -ItemType Directory -Force | Out-Null }
        $sharedLogPath = Join-Path $logsDir ("SP_Export_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

        $successCount = 0
        $errorCount = 0

        foreach ($applianceHost in $selectedAppliances) {
            Write-GUILog "Starte SP Export von $applianceHost ..." -Color ([System.Drawing.Color]::Cyan)

            $safeName = $applianceHost -replace '[\\/:*?\"<>|\.]', '_'
            $exportPath = Join-Path $backupDir $safeName
            New-Item -Path $exportPath -ItemType Directory -Force | Out-Null

            $tempConfig = New-TempConfig -Hostname $applianceHost -ExcelPath "" -TempDir $scriptDir

            try {
                $exportScript = Join-Path $scriptDir "Export-ServerProfiles.ps1"
                $psCommand = @"
`$env:OV_USERNAME = '$($txtUser.Text)'
`$env:OV_PASSWORD = '$($txtPass.Text -replace "'", "''")'
`$secPass = ConvertTo-SecureString `$env:OV_PASSWORD -AsPlainText -Force
`$global:guiCredential = New-Object System.Management.Automation.PSCredential(`$env:OV_USERNAME, `$secPass)
function Get-Credential { param([string]`$Message) return `$global:guiCredential }
& '$exportScript' -ConfigPath '$tempConfig' -OutputPath '$exportPath' -LogPath '$sharedLogPath'
"@
                $exitCode = Invoke-SubprocessWithLiveOutput -Command $psCommand

                if ($exitCode -eq 0) {
                    Write-GUILog "SP Export erfolgreich: $applianceHost" -Color ([System.Drawing.Color]::FromArgb(80, 220, 80))
                    $successCount++
                } else {
                    Write-GUILog "SP Export fehlgeschlagen: $applianceHost" -Color ([System.Drawing.Color]::FromArgb(255, 80, 80))
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
        Write-GUILog "SP Export abgeschlossen: $successCount erfolgreich, $errorCount fehlgeschlagen" -Color ([System.Drawing.Color]::Cyan)
        Write-GUILog "Verzeichnis: $backupDir" -Color ([System.Drawing.Color]::Cyan)

        [System.Windows.Forms.MessageBox]::Show(
            "Server Profile Export abgeschlossen!`n`nErfolgreich: $successCount`nFehlgeschlagen: $errorCount`n`nVerzeichnis: $backupDir",
            "SP Export Ergebnis",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnSPExport.Enabled = $true
        $btnSPImport.Enabled = $true
        $btnSPManage.Enabled = $true
        $btnSPJsonEdit.Enabled = $true
        $btnSPTExport.Enabled = $true
        $btnSPTImport.Enabled = $true
        $btnSPTManage.Enabled = $true
        $btnSPTJsonEdit.Enabled = $true
    }
})

# ============================================================================
#  Button-Event: Server Profile importieren (Einzelne Appliance)
# ============================================================================
$btnSPImport.Add_Click({
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

    $hostname = Show-SingleApplianceSelectionDialog -Appliances $appliances -Title "SP Import – Appliance auswählen"
    if (-not $hostname) { return }

    # Modus auswählen
    $modeForm = New-Object System.Windows.Forms.Form
    $modeForm.Text = "Import-Modus wählen"
    $modeForm.Size = New-Object System.Drawing.Size(400, 200)
    $modeForm.StartPosition = "CenterParent"
    $modeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $modeForm.MaximizeBox = $false; $modeForm.MinimizeBox = $false
    $modeForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblMode = New-Object System.Windows.Forms.Label
    $lblMode.Text = "Wählen Sie den Import-Modus:"
    $lblMode.Location = New-Object System.Drawing.Point(15, 15)
    $lblMode.Size = New-Object System.Drawing.Size(360, 20)
    $modeForm.Controls.Add($lblMode)

    $cmbMode = New-Object System.Windows.Forms.ComboBox
    $cmbMode.Location = New-Object System.Drawing.Point(15, 45)
    $cmbMode.Size = New-Object System.Drawing.Size(350, 23)
    $cmbMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @("Auto (Erstellen oder Aktualisieren)", "Create (Nur neue erstellen)", "Update (Nur bestehende aktualisieren)") |
        ForEach-Object { $cmbMode.Items.Add($_) | Out-Null }
    $cmbMode.SelectedIndex = 0
    $modeForm.Controls.Add($cmbMode)

    $btnModeOK = New-Object System.Windows.Forms.Button
    $btnModeOK.Text = "OK"; $btnModeOK.Location = New-Object System.Drawing.Point(200, 100)
    $btnModeOK.Size = New-Object System.Drawing.Size(75, 28)
    $btnModeOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $modeForm.AcceptButton = $btnModeOK
    $modeForm.Controls.Add($btnModeOK)

    $btnModeCancel = New-Object System.Windows.Forms.Button
    $btnModeCancel.Text = "Abbrechen"; $btnModeCancel.Location = New-Object System.Drawing.Point(280, 100)
    $btnModeCancel.Size = New-Object System.Drawing.Size(85, 28)
    $btnModeCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $modeForm.CancelButton = $btnModeCancel
    $modeForm.Controls.Add($btnModeCancel)

    if ($modeForm.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) {
        $modeForm.Dispose(); return
    }
    $importMode = @("Auto", "Create", "Update")[$cmbMode.SelectedIndex]
    $modeForm.Dispose()

    # JSON-Datei oder Verzeichnis auswählen
    $msgResult = [System.Windows.Forms.MessageBox]::Show(
        "Einzelne JSON-Datei importieren?`n`nJa = Datei auswählen`nNein = Verzeichnis auswählen (mehrere Profile)",
        "Eingabe wählen",
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($msgResult -eq [System.Windows.Forms.DialogResult]::Cancel) { return }

    $inputPath = ""
    if ($msgResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "JSON Dateien (*.json)|*.json"
        $ofd.InitialDirectory = $scriptDir
        $ofd.Title = "Server Profile JSON-Datei auswählen"
        if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $inputPath = $ofd.FileName
    } else {
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "Verzeichnis mit Server Profile JSON-Dateien auswählen"
        $fbd.SelectedPath = $scriptDir
        if ($fbd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $inputPath = $fbd.SelectedPath
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Server Profiles importieren auf`n$hostname`n`nModus: $importMode`nQuelle: $inputPath",
        "SP Import bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnSPImport.Enabled = $false
    $rtbLog.Clear()

    try {
        Write-GUILog "Starte SP Import auf $hostname (Modus: $importMode) ..." -Color ([System.Drawing.Color]::Cyan)

        $tempConfig = New-TempConfig -Hostname $hostname -ExcelPath "" -TempDir $scriptDir

        $importScript = Join-Path $scriptDir "Import-ServerProfiles.ps1"
        $escapedInputPath = $inputPath -replace "'", "''"
        $psCommand = @"
`$env:OV_USERNAME = '$($txtUser.Text)'
`$env:OV_PASSWORD = '$($txtPass.Text -replace "'", "''")'
`$secPass = ConvertTo-SecureString `$env:OV_PASSWORD -AsPlainText -Force
`$global:guiCredential = New-Object System.Management.Automation.PSCredential(`$env:OV_USERNAME, `$secPass)
function Get-Credential { param([string]`$Message) return `$global:guiCredential }
& '$importScript' -ConfigPath '$tempConfig' -InputPath '$escapedInputPath' -Mode '$importMode'
"@
        Invoke-SubprocessWithLiveOutput -Command $psCommand | Out-Null

        Write-GUILog "SP Import abgeschlossen." -Color ([System.Drawing.Color]::Cyan)

        Remove-Item -Path $tempConfig -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnSPImport.Enabled = $true
        $btnSPJsonEdit.Enabled = $true
        $btnSPTExport.Enabled = $true
        $btnSPTImport.Enabled = $true
        $btnSPTManage.Enabled = $true
        $btnSPTJsonEdit.Enabled = $true
    }
})

# ============================================================================
#  Button-Event: Server Profile JSON Editor
# ============================================================================
$btnSPJsonEdit.Add_Click({
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

    $hostname = Show-SingleApplianceSelectionDialog -Appliances $appliances -Title "SP JSON Editor – Appliance auswählen"
    if (-not $hostname) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Write-GUILog "Verbinde zu $hostname für JSON Editor..." -Color ([System.Drawing.Color]::Cyan)

    $session = $null
    try {
        $apiVersion = Get-ApiVersionInline -Hostname $hostname
        $session = Connect-OneViewAPIInline -Hostname $hostname `
            -Username $txtUser.Text -Password $txtPass.Text -ApiVersion $apiVersion

        Write-GUILog "Verbunden. Öffne JSON Editor..." -Color ([System.Drawing.Color]::FromArgb(80, 220, 80))
        $form.Cursor = [System.Windows.Forms.Cursors]::Default

        Show-ServerProfileJsonEditor -Session $session -Hostname $hostname
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
        [System.Windows.Forms.MessageBox]::Show(
            "Verbindung fehlgeschlagen:`n$_",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        if ($session) { Disconnect-OneViewAPIInline -Session $session }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# ============================================================================
#  Button-Event: Server Profile verwalten (Dialog)
# ============================================================================
$btnSPManage.Add_Click({
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

    $hostname = Show-SingleApplianceSelectionDialog -Appliances $appliances -Title "SP verwalten – Appliance auswählen"
    if (-not $hostname) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Write-GUILog "Verbinde zu $hostname für Server Profile Verwaltung..." -Color ([System.Drawing.Color]::Cyan)

    $session = $null
    try {
        $apiVersion = Get-ApiVersionInline -Hostname $hostname
        $session = Connect-OneViewAPIInline -Hostname $hostname `
            -Username $txtUser.Text -Password $txtPass.Text -ApiVersion $apiVersion

        Write-GUILog "Verbunden. Öffne Verwaltungsdialog..." -Color ([System.Drawing.Color]::FromArgb(80, 220, 80))
        $form.Cursor = [System.Windows.Forms.Cursors]::Default

        Show-ServerProfileManageDialog -Session $session -Hostname $hostname
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
        [System.Windows.Forms.MessageBox]::Show(
            "Verbindung fehlgeschlagen:`n$_",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        if ($session) { Disconnect-OneViewAPIInline -Session $session }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# ============================================================================
#  Button-Event: Server Profile Template exportieren (Multi-Appliance)
# ============================================================================
$btnSPTExport.Add_Click({
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

    $selectedAppliances = Show-ApplianceSelectionDialog -Appliances $appliances -Title "SPT Export – Appliances auswählen"
    if (-not $selectedAppliances -or $selectedAppliances.Count -eq 0) { return }

    $applianceList = $selectedAppliances -join "`n"
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Server Profile Templates von $($selectedAppliances.Count) Appliance(s) exportieren?`n`n$applianceList",
        "SPT Export bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnSPTExport.Enabled = $false
    $btnSPTImport.Enabled = $false
    $btnSPTManage.Enabled = $false
    $btnSPTJsonEdit.Enabled = $false
    $rtbLog.Clear()

    try {
        $backupDir = Join-Path $scriptDir "Backups" ("SPT_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
        New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
        Write-GUILog "Backup-Verzeichnis: $backupDir" -Color ([System.Drawing.Color]::Cyan)

        $logsDir = Join-Path $scriptDir "Logs"
        if (-not (Test-Path $logsDir)) { New-Item -Path $logsDir -ItemType Directory -Force | Out-Null }
        $sharedLogPath = Join-Path $logsDir ("SPT_Export_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

        $successCount = 0
        $errorCount = 0

        foreach ($applianceHost in $selectedAppliances) {
            Write-GUILog "Starte SPT Export von $applianceHost ..." -Color ([System.Drawing.Color]::Cyan)

            $safeName = $applianceHost -replace '[\\/:*?\"<>|\.]', '_'
            $exportPath = Join-Path $backupDir $safeName
            New-Item -Path $exportPath -ItemType Directory -Force | Out-Null

            $tempConfig = New-TempConfig -Hostname $applianceHost -ExcelPath "" -TempDir $scriptDir

            try {
                $exportScript = Join-Path $scriptDir "Export-ServerProfileTemplates.ps1"
                $psCommand = @"
`$env:OV_USERNAME = '$($txtUser.Text)'
`$env:OV_PASSWORD = '$($txtPass.Text -replace "'", "''")'
`$secPass = ConvertTo-SecureString `$env:OV_PASSWORD -AsPlainText -Force
`$global:guiCredential = New-Object System.Management.Automation.PSCredential(`$env:OV_USERNAME, `$secPass)
function Get-Credential { param([string]`$Message) return `$global:guiCredential }
& '$exportScript' -ConfigPath '$tempConfig' -OutputPath '$exportPath' -LogPath '$sharedLogPath'
"@
                $exitCode = Invoke-SubprocessWithLiveOutput -Command $psCommand

                if ($exitCode -eq 0) {
                    Write-GUILog "SPT Export erfolgreich: $applianceHost" -Color ([System.Drawing.Color]::FromArgb(80, 220, 80))
                    $successCount++
                } else {
                    Write-GUILog "SPT Export fehlgeschlagen: $applianceHost" -Color ([System.Drawing.Color]::FromArgb(255, 80, 80))
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
        Write-GUILog "SPT Export abgeschlossen: $successCount erfolgreich, $errorCount fehlgeschlagen" -Color ([System.Drawing.Color]::Cyan)
        Write-GUILog "Verzeichnis: $backupDir" -Color ([System.Drawing.Color]::Cyan)

        [System.Windows.Forms.MessageBox]::Show(
            "Server Profile Template Export abgeschlossen!`n`nErfolgreich: $successCount`nFehlgeschlagen: $errorCount`n`nVerzeichnis: $backupDir",
            "SPT Export Ergebnis",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnSPTExport.Enabled = $true
        $btnSPTImport.Enabled = $true
        $btnSPTManage.Enabled = $true
        $btnSPTJsonEdit.Enabled = $true
    }
})

# ============================================================================
#  Button-Event: Server Profile Template importieren (Einzelne Appliance)
# ============================================================================
$btnSPTImport.Add_Click({
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

    $hostname = Show-SingleApplianceSelectionDialog -Appliances $appliances -Title "SPT Import – Appliance auswählen"
    if (-not $hostname) { return }

    # Modus auswählen
    $modeForm = New-Object System.Windows.Forms.Form
    $modeForm.Text = "Import-Modus wählen"
    $modeForm.Size = New-Object System.Drawing.Size(400, 200)
    $modeForm.StartPosition = "CenterParent"
    $modeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $modeForm.MaximizeBox = $false; $modeForm.MinimizeBox = $false
    $modeForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblMode = New-Object System.Windows.Forms.Label
    $lblMode.Text = "Wählen Sie den Import-Modus:"
    $lblMode.Location = New-Object System.Drawing.Point(15, 15)
    $lblMode.Size = New-Object System.Drawing.Size(360, 20)
    $modeForm.Controls.Add($lblMode)

    $cmbMode = New-Object System.Windows.Forms.ComboBox
    $cmbMode.Location = New-Object System.Drawing.Point(15, 45)
    $cmbMode.Size = New-Object System.Drawing.Size(350, 23)
    $cmbMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @("Auto (Erstellen oder Aktualisieren)", "Create (Nur neue erstellen)", "Update (Nur bestehende aktualisieren)") |
        ForEach-Object { $cmbMode.Items.Add($_) | Out-Null }
    $cmbMode.SelectedIndex = 0
    $modeForm.Controls.Add($cmbMode)

    $btnModeOK = New-Object System.Windows.Forms.Button
    $btnModeOK.Text = "OK"; $btnModeOK.Location = New-Object System.Drawing.Point(200, 100)
    $btnModeOK.Size = New-Object System.Drawing.Size(75, 28)
    $btnModeOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $modeForm.AcceptButton = $btnModeOK
    $modeForm.Controls.Add($btnModeOK)

    $btnModeCancel = New-Object System.Windows.Forms.Button
    $btnModeCancel.Text = "Abbrechen"; $btnModeCancel.Location = New-Object System.Drawing.Point(280, 100)
    $btnModeCancel.Size = New-Object System.Drawing.Size(85, 28)
    $btnModeCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $modeForm.CancelButton = $btnModeCancel
    $modeForm.Controls.Add($btnModeCancel)

    if ($modeForm.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) {
        $modeForm.Dispose(); return
    }
    $importMode = @("Auto", "Create", "Update")[$cmbMode.SelectedIndex]
    $modeForm.Dispose()

    # JSON-Datei oder Verzeichnis auswählen
    $msgResult = [System.Windows.Forms.MessageBox]::Show(
        "Einzelne JSON-Datei importieren?`n`nJa = Datei auswählen`nNein = Verzeichnis auswählen (mehrere Templates)",
        "Eingabe wählen",
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($msgResult -eq [System.Windows.Forms.DialogResult]::Cancel) { return }

    $inputPath = ""
    if ($msgResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "JSON Dateien (*.json)|*.json"
        $ofd.InitialDirectory = $scriptDir
        $ofd.Title = "Server Profile Template JSON-Datei auswählen"
        if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $inputPath = $ofd.FileName
    } else {
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "Verzeichnis mit Server Profile Template JSON-Dateien auswählen"
        $fbd.SelectedPath = $scriptDir
        if ($fbd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $inputPath = $fbd.SelectedPath
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Server Profile Templates importieren auf`n$hostname`n`nModus: $importMode`nQuelle: $inputPath",
        "SPT Import bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnSPTImport.Enabled = $false
    $rtbLog.Clear()

    try {
        Write-GUILog "Starte SPT Import auf $hostname (Modus: $importMode) ..." -Color ([System.Drawing.Color]::Cyan)

        $tempConfig = New-TempConfig -Hostname $hostname -ExcelPath "" -TempDir $scriptDir

        $importScript = Join-Path $scriptDir "Import-ServerProfileTemplates.ps1"
        $escapedInputPath = $inputPath -replace "'", "''"
        $psCommand = @"
`$env:OV_USERNAME = '$($txtUser.Text)'
`$env:OV_PASSWORD = '$($txtPass.Text -replace "'", "''")'
`$secPass = ConvertTo-SecureString `$env:OV_PASSWORD -AsPlainText -Force
`$global:guiCredential = New-Object System.Management.Automation.PSCredential(`$env:OV_USERNAME, `$secPass)
function Get-Credential { param([string]`$Message) return `$global:guiCredential }
& '$importScript' -ConfigPath '$tempConfig' -InputPath '$escapedInputPath' -Mode '$importMode'
"@
        Invoke-SubprocessWithLiveOutput -Command $psCommand | Out-Null

        Write-GUILog "SPT Import abgeschlossen." -Color ([System.Drawing.Color]::Cyan)

        Remove-Item -Path $tempConfig -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnSPTImport.Enabled = $true
        $btnSPTJsonEdit.Enabled = $true
    }
})

# ============================================================================
#  Button-Event: Server Profile Template JSON Editor
# ============================================================================
$btnSPTJsonEdit.Add_Click({
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

    $hostname = Show-SingleApplianceSelectionDialog -Appliances $appliances -Title "SPT JSON Editor – Appliance auswählen"
    if (-not $hostname) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Write-GUILog "Verbinde zu $hostname für SPT JSON Editor..." -Color ([System.Drawing.Color]::Cyan)

    $session = $null
    try {
        $apiVersion = Get-ApiVersionInline -Hostname $hostname
        $session = Connect-OneViewAPIInline -Hostname $hostname `
            -Username $txtUser.Text -Password $txtPass.Text -ApiVersion $apiVersion

        Write-GUILog "Verbunden. Öffne JSON Editor..." -Color ([System.Drawing.Color]::FromArgb(80, 220, 80))
        $form.Cursor = [System.Windows.Forms.Cursors]::Default

        Show-ServerProfileTemplateJsonEditor -Session $session -Hostname $hostname
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
        [System.Windows.Forms.MessageBox]::Show(
            "Verbindung fehlgeschlagen:`n$_",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        if ($session) { Disconnect-OneViewAPIInline -Session $session }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# ============================================================================
#  Button-Event: Server Profile Template verwalten (Dialog)
# ============================================================================
$btnSPTManage.Add_Click({
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

    $hostname = Show-SingleApplianceSelectionDialog -Appliances $appliances -Title "SPT verwalten – Appliance auswählen"
    if (-not $hostname) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Write-GUILog "Verbinde zu $hostname für Server Profile Template Verwaltung..." -Color ([System.Drawing.Color]::Cyan)

    $session = $null
    try {
        $apiVersion = Get-ApiVersionInline -Hostname $hostname
        $session = Connect-OneViewAPIInline -Hostname $hostname `
            -Username $txtUser.Text -Password $txtPass.Text -ApiVersion $apiVersion

        Write-GUILog "Verbunden. Öffne Verwaltungsdialog..." -Color ([System.Drawing.Color]::FromArgb(80, 220, 80))
        $form.Cursor = [System.Windows.Forms.Cursors]::Default

        Show-ServerProfileTemplateManageDialog -Session $session -Hostname $hostname
    }
    catch {
        Write-GUILog "Fehler: $_" -Color ([System.Drawing.Color]::Red)
        [System.Windows.Forms.MessageBox]::Show(
            "Verbindung fehlgeschlagen:`n$_",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        if ($session) { Disconnect-OneViewAPIInline -Session $session }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# ============================================================================
#  Formular anzeigen
# ============================================================================
$form.Add_Shown({ $txtUser.Focus() })
[void]$form.ShowDialog()
