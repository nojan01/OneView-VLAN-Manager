#Requires -Version 7.0
<#
.SYNOPSIS
    HPE Synergy eFuse Tool – GUI

.DESCRIPTION
    Grafische Oberfläche zum Auslösen eines eFuse (Reset-OVEnclosureDevice -eFuse)
    auf HPE Synergy Komponenten (Device / ICM / Appliance / FLM).
    Im Stil der übrigen GUIs des OneView_VLAN_Projekt:
    - Login-Bereich
    - Appliance-Auswahl (aus Oneview.txt)
    - Verbinden mit OneView
    - Auswahl von Enclosure / Komponente / Slot
    - Bestätigungsabfrage
    - Protokoll-Bereich

.NOTES
    Autor:   N.J. Airbus D&S
    Benötigt: HPE OneView PowerShell Module (HPEOneView.*) und Oneview.txt
              (eine Appliance pro Zeile) im Skript-Ordner.
#>

# ============================================================================
#  PowerShell-Konsolenfenster verstecken (nur Windows)
# ============================================================================
try {
    if (-not ([System.Management.Automation.PSTypeName]'Native.Win32eFuse').Type) {
        Add-Type -Name 'Win32eFuse' -Namespace 'Native' -MemberDefinition @"
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
"@ -ErrorAction SilentlyContinue
    }
    $consoleWindow = [Native.Win32eFuse]::GetConsoleWindow()
    if ($consoleWindow -ne [IntPtr]::Zero) {
        [Native.Win32eFuse]::ShowWindow($consoleWindow, 0) | Out-Null
    }
} catch { }

# ============================================================================
#  Assemblies & Voraussetzungen
# ============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir     = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$applianceFile = Join-Path $scriptDir "Oneview.txt"

# Mapping: Anzeige-Name -> interner OneView-Komponentenname
$componentMapping = [ordered]@{
    "Device (BladeServer)"        = "Device"
    "ICM (VirtualConnectModule)"  = "ICM"
    "Appliance (Composer)"        = "Appliance"
    "FLM (FrameLinkModule)"       = "FLM"
}

# Globaler Verbindungsstatus
$script:Connection = $null

# ============================================================================
#  Hauptformular
# ============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text            = "© 2025 N.J. Airbus D&S - HPE Synergy eFuse Tool"
$form.Size            = New-Object System.Drawing.Size(760, 760)
$form.StartPosition   = "CenterScreen"
$form.MinimumSize     = New-Object System.Drawing.Size(720, 700)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

$boldFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

# ----------------------------------------------------------------------------
#  GroupBox: Anmeldeinformationen
# ----------------------------------------------------------------------------
$grpCred = New-Object System.Windows.Forms.GroupBox
$grpCred.Text     = "Anmeldeinformationen"
$grpCred.Location = New-Object System.Drawing.Point(15, 15)
$grpCred.Size     = New-Object System.Drawing.Size(710, 110)
$grpCred.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($grpCred)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text     = "Benutzername:"
$lblUser.Location = New-Object System.Drawing.Point(15, 30)
$lblUser.Size     = New-Object System.Drawing.Size(110, 23)
$grpCred.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(130, 27)
$txtUser.Size     = New-Object System.Drawing.Size(560, 23)
$txtUser.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpCred.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text     = "Kennwort:"
$lblPass.Location = New-Object System.Drawing.Point(15, 65)
$lblPass.Size     = New-Object System.Drawing.Size(110, 23)
$grpCred.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location             = New-Object System.Drawing.Point(130, 62)
$txtPass.Size                 = New-Object System.Drawing.Size(560, 23)
$txtPass.UseSystemPasswordChar = $true
$txtPass.Anchor               = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpCred.Controls.Add($txtPass)

# ----------------------------------------------------------------------------
#  GroupBox: Appliance-Verbindung
# ----------------------------------------------------------------------------
$grpAppliance = New-Object System.Windows.Forms.GroupBox
$grpAppliance.Text     = "OneView Appliance"
$grpAppliance.Location = New-Object System.Drawing.Point(15, 135)
$grpAppliance.Size     = New-Object System.Drawing.Size(710, 90)
$grpAppliance.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($grpAppliance)

$lblAppliance = New-Object System.Windows.Forms.Label
$lblAppliance.Text     = "Appliance:"
$lblAppliance.Location = New-Object System.Drawing.Point(15, 30)
$lblAppliance.Size     = New-Object System.Drawing.Size(110, 23)
$grpAppliance.Controls.Add($lblAppliance)

$cmbAppliance = New-Object System.Windows.Forms.ComboBox
$cmbAppliance.DropDownStyle = "DropDownList"
$cmbAppliance.Location      = New-Object System.Drawing.Point(130, 27)
$cmbAppliance.Size          = New-Object System.Drawing.Size(440, 23)
$cmbAppliance.Anchor        = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpAppliance.Controls.Add($cmbAppliance)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text      = "Verbinden"
$btnConnect.Location  = New-Object System.Drawing.Point(580, 26)
$btnConnect.Size      = New-Object System.Drawing.Size(110, 26)
$btnConnect.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$btnConnect.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnConnect.ForeColor = [System.Drawing.Color]::White
$btnConnect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnConnect.Font      = $boldFont
$grpAppliance.Controls.Add($btnConnect)

$lblConnState = New-Object System.Windows.Forms.Label
$lblConnState.Text      = "Nicht verbunden"
$lblConnState.Location  = New-Object System.Drawing.Point(130, 58)
$lblConnState.Size      = New-Object System.Drawing.Size(440, 20)
$lblConnState.ForeColor = [System.Drawing.Color]::Firebrick
$lblConnState.Font      = $boldFont
$grpAppliance.Controls.Add($lblConnState)

$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Text      = "Trennen"
$btnDisconnect.Location  = New-Object System.Drawing.Point(580, 56)
$btnDisconnect.Size      = New-Object System.Drawing.Size(110, 24)
$btnDisconnect.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$btnDisconnect.Enabled   = $false
$grpAppliance.Controls.Add($btnDisconnect)

# ----------------------------------------------------------------------------
#  GroupBox: eFuse Parameter
# ----------------------------------------------------------------------------
$grpParams = New-Object System.Windows.Forms.GroupBox
$grpParams.Text     = "eFuse Parameter"
$grpParams.Location = New-Object System.Drawing.Point(15, 235)
$grpParams.Size     = New-Object System.Drawing.Size(710, 170)
$grpParams.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($grpParams)

# Enclosure
$lblEnclosure = New-Object System.Windows.Forms.Label
$lblEnclosure.Text     = "Enclosure:"
$lblEnclosure.Location = New-Object System.Drawing.Point(15, 30)
$lblEnclosure.Size     = New-Object System.Drawing.Size(110, 23)
$grpParams.Controls.Add($lblEnclosure)

$cmbEnclosure = New-Object System.Windows.Forms.ComboBox
$cmbEnclosure.DropDownStyle = "DropDownList"
$cmbEnclosure.Location      = New-Object System.Drawing.Point(130, 27)
$cmbEnclosure.Size          = New-Object System.Drawing.Size(560, 23)
$cmbEnclosure.Anchor        = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$cmbEnclosure.Enabled       = $false
$grpParams.Controls.Add($cmbEnclosure)

# Komponente
$lblComponent = New-Object System.Windows.Forms.Label
$lblComponent.Text     = "Komponente:"
$lblComponent.Location = New-Object System.Drawing.Point(15, 65)
$lblComponent.Size     = New-Object System.Drawing.Size(110, 23)
$grpParams.Controls.Add($lblComponent)

$cmbComponent = New-Object System.Windows.Forms.ComboBox
$cmbComponent.DropDownStyle = "DropDownList"
$cmbComponent.Location      = New-Object System.Drawing.Point(130, 62)
$cmbComponent.Size          = New-Object System.Drawing.Size(560, 23)
$cmbComponent.Anchor        = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
foreach ($k in $componentMapping.Keys) { $cmbComponent.Items.Add($k) | Out-Null }
if ($cmbComponent.Items.Count -gt 0) { $cmbComponent.SelectedIndex = 0 }
$grpParams.Controls.Add($cmbComponent)

# Slot
$lblSlot = New-Object System.Windows.Forms.Label
$lblSlot.Text     = "Slot-Nummer:"
$lblSlot.Location = New-Object System.Drawing.Point(15, 100)
$lblSlot.Size     = New-Object System.Drawing.Size(110, 23)
$grpParams.Controls.Add($lblSlot)

$numSlot = New-Object System.Windows.Forms.NumericUpDown
$numSlot.Location = New-Object System.Drawing.Point(130, 98)
$numSlot.Size     = New-Object System.Drawing.Size(80, 23)
$numSlot.Minimum  = 1
$numSlot.Maximum  = 999
$numSlot.Value    = 1
$grpParams.Controls.Add($numSlot)

# Ausführen-Button
$btnExecute = New-Object System.Windows.Forms.Button
$btnExecute.Text      = "eFuse ausführen..."
$btnExecute.Location  = New-Object System.Drawing.Point(130, 130)
$btnExecute.Size      = New-Object System.Drawing.Size(200, 30)
$btnExecute.BackColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
$btnExecute.ForeColor = [System.Drawing.Color]::White
$btnExecute.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExecute.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnExecute.Enabled   = $false
$grpParams.Controls.Add($btnExecute)

# ----------------------------------------------------------------------------
#  GroupBox: Protokoll
# ----------------------------------------------------------------------------
$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text     = "Protokoll"
$grpLog.Location = New-Object System.Drawing.Point(15, 415)
$grpLog.Size     = New-Object System.Drawing.Size(710, 290)
$grpLog.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($grpLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline   = $true
$txtLog.ReadOnly    = $true
$txtLog.ScrollBars  = "Vertical"
$txtLog.Location    = New-Object System.Drawing.Point(10, 22)
$txtLog.Size        = New-Object System.Drawing.Size(690, 258)
$txtLog.Anchor      = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$txtLog.Font        = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor   = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtLog.ForeColor   = [System.Drawing.Color]::Gainsboro
$grpLog.Controls.Add($txtLog)

# ============================================================================
#  Hilfsfunktionen
# ============================================================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error')]
        [string]$Level = 'Info'
    )
    $stamp = (Get-Date).ToString("HH:mm:ss")
    $tag = switch ($Level) {
        'Success' { '[OK]   ' }
        'Warning' { '[WARN] ' }
        'Error'   { '[FEHL] ' }
        default   { '[INFO] ' }
    }
    $line = "$stamp $tag$Message"
    $txtLog.AppendText($line + [Environment]::NewLine)
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Import-AppliancesFromFile {
    $cmbAppliance.Items.Clear()
    if (-not (Test-Path $applianceFile)) {
        Write-Log "Datei nicht gefunden: $applianceFile" -Level Error
        Write-Log "Bitte 'Oneview.txt' (eine Appliance pro Zeile) im Ordner '$scriptDir' ablegen." -Level Warning
        return
    }
    $list = Get-Content $applianceFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') }
    foreach ($a in $list) { $cmbAppliance.Items.Add($a) | Out-Null }
    if ($cmbAppliance.Items.Count -gt 0) {
        $cmbAppliance.SelectedIndex = 0
        Write-Log "$($cmbAppliance.Items.Count) Appliance(s) aus Oneview.txt geladen." -Level Info
    } else {
        Write-Log "Oneview.txt enthält keine Einträge." -Level Warning
    }
}

function Set-ConnectedState {
    param([bool]$Connected, [string]$ApplianceName = '')

    if ($Connected) {
        $lblConnState.Text      = "Verbunden mit: $ApplianceName"
        $lblConnState.ForeColor = [System.Drawing.Color]::ForestGreen
        $btnConnect.Enabled     = $false
        $btnDisconnect.Enabled  = $true
        $cmbAppliance.Enabled   = $false
        $cmbEnclosure.Enabled   = $true
        $btnExecute.Enabled     = $true
        $txtUser.Enabled        = $false
        $txtPass.Enabled        = $false
    } else {
        $lblConnState.Text      = "Nicht verbunden"
        $lblConnState.ForeColor = [System.Drawing.Color]::Firebrick
        $btnConnect.Enabled     = $true
        $btnDisconnect.Enabled  = $false
        $cmbAppliance.Enabled   = $true
        $cmbEnclosure.Enabled   = $false
        $btnExecute.Enabled     = $false
        $cmbEnclosure.Items.Clear()
        $txtUser.Enabled        = $true
        $txtPass.Enabled        = $true
    }
}

function Test-OVModuleLoaded {
    if (-not (Get-Command -Name Connect-OVMgmt -ErrorAction SilentlyContinue)) {
        Write-Log "Cmdlet 'Connect-OVMgmt' nicht gefunden. Versuche HPEOneView-Modul zu laden..." -Level Warning
        $mod = Get-Module -ListAvailable -Name 'HPEOneView.*' | Sort-Object Version -Descending | Select-Object -First 1
        if ($mod) {
            try {
                Import-Module $mod.Name -ErrorAction Stop
                Write-Log "Modul $($mod.Name) (v$($mod.Version)) geladen." -Level Success
                return $true
            } catch {
                Write-Log "Fehler beim Laden von $($mod.Name): $($_.Exception.Message)" -Level Error
                return $false
            }
        } else {
            Write-Log "Kein HPEOneView-Modul gefunden. Bitte installieren (Install-Module HPEOneView.*)." -Level Error
            return $false
        }
    }
    return $true
}

# ============================================================================
#  Event-Handler
# ============================================================================
$btnConnect.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtUser.Text) -or [string]::IsNullOrWhiteSpace($txtPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Benutzername und Kennwort eingeben.", "Anmeldung",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if (-not $cmbAppliance.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Bitte eine Appliance auswählen.", "Appliance",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if (-not (Test-OVModuleLoaded)) { return }

    $appliance = [string]$cmbAppliance.SelectedItem
    $cred = New-Object System.Management.Automation.PSCredential(
        $txtUser.Text,
        (ConvertTo-SecureString $txtPass.Text -AsPlainText -Force)
    )

    Write-Log "Verbinde mit $appliance ..." -Level Info
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $script:Connection = Connect-OVMgmt -Appliance $appliance -Credential $cred -ErrorAction Stop
        Write-Log "Erfolgreich verbunden mit $appliance." -Level Success
        Set-ConnectedState -Connected $true -ApplianceName $appliance

        # Enclosures laden
        Write-Log "Lade Enclosures ..." -Level Info
        $cmbEnclosure.Items.Clear()
        $encs = Get-OVEnclosure -ErrorAction Stop
        foreach ($e in $encs) { $cmbEnclosure.Items.Add($e.Name) | Out-Null }
        if ($cmbEnclosure.Items.Count -gt 0) {
            $cmbEnclosure.SelectedIndex = 0
            Write-Log "$($cmbEnclosure.Items.Count) Enclosure(s) geladen." -Level Success
        } else {
            Write-Log "Keine Enclosures gefunden." -Level Warning
        }
    }
    catch {
        Write-Log "Verbindung fehlgeschlagen: $($_.Exception.Message)" -Level Error
        try { Disconnect-OVMgmt -ErrorAction SilentlyContinue } catch { }
        $script:Connection = $null
        Set-ConnectedState -Connected $false
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnDisconnect.Add_Click({
    try {
        Disconnect-OVMgmt -ErrorAction SilentlyContinue
        Write-Log "Verbindung getrennt." -Level Info
    } catch {
        Write-Log "Fehler beim Trennen: $($_.Exception.Message)" -Level Warning
    }
    $script:Connection = $null
    Set-ConnectedState -Connected $false
})

$btnExecute.Add_Click({
    if (-not $script:Connection) {
        [System.Windows.Forms.MessageBox]::Show("Keine aktive OneView-Verbindung.", "Nicht verbunden",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if (-not $cmbEnclosure.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Bitte ein Enclosure auswählen.", "Enclosure",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if (-not $cmbComponent.SelectedItem) { return }

    $enclosureName     = [string]$cmbEnclosure.SelectedItem
    $componentDisplay  = [string]$cmbComponent.SelectedItem
    $componentInternal = $componentMapping[$componentDisplay]
    $slot              = [int]$numSlot.Value
    $appliance         = [string]$cmbAppliance.SelectedItem

    # ---- Bestätigungsabfrage ----
    $msg = @"
Sie sind dabei, einen eFuse auszulösen!

Appliance:   $appliance
Enclosure:   $enclosureName
Komponente:  $componentDisplay  (intern: $componentInternal)
Slot:        $slot

Diese Aktion schaltet das Gerät hart ab. Möchten Sie wirklich fortfahren?
"@
    $result = [System.Windows.Forms.MessageBox]::Show(
        $msg,
        "eFuse ausführen - bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Log "Abbruch durch Benutzer (Bestätigung verweigert)." -Level Warning
        return
    }

    Write-Log "Starte eFuse: Enclosure='$enclosureName', Component='$componentInternal', Slot=$slot ..." -Level Info
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnExecute.Enabled = $false
    try {
        $task = Get-OVEnclosure -Name $enclosureName -ErrorAction Stop |
                Reset-OVEnclosureDevice -Component $componentInternal -DeviceID $slot -eFuse -Confirm:$false -ErrorAction Stop
        Write-Log "eFuse-Befehl an OneView gesendet." -Level Success
        if ($task) {
            Write-Log "Task: $($task.name) – Status: $($task.taskState)" -Level Info
        }
        [System.Windows.Forms.MessageBox]::Show(
            "eFuse-Befehl wurde erfolgreich ausgeführt.",
            "Erfolg",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Write-Log "Fehler beim Ausführen des eFuse-Befehls: $($_.Exception.Message)" -Level Error
        [System.Windows.Forms.MessageBox]::Show(
            "Fehler beim Ausführen des eFuse-Befehls:`n$($_.Exception.Message)",
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnExecute.Enabled = $true
    }
})

# Beim Schließen ggf. trennen
$form.Add_FormClosing({
    if ($script:Connection) {
        try { Disconnect-OVMgmt -ErrorAction SilentlyContinue } catch { }
    }
})

# ============================================================================
#  Initialisierung
# ============================================================================
$form.Add_Shown({
    $form.Activate()
    Write-Log "HPE Synergy eFuse Tool gestartet." -Level Info
    Write-Log "Skript-Ordner: $scriptDir" -Level Info
    Import-AppliancesFromFile
    Set-ConnectedState -Connected $false
})

[void]$form.ShowDialog()
