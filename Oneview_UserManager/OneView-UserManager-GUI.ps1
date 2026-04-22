# ============================================================================
#  HPE OneView User Manager GUI
#  Direkte REST-API – keine HPE PowerShell-Module erforderlich
#  X-API-Version wird automatisch pro Appliance ermittelt
#  Unterstützt OV 6.60 + OV 11.10 (und beliebige andere Versionen)
# ============================================================================

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
# Assemblies
# =============================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =============================
# REST-Helper-Code (wird im Runspace per Invoke-Expression geladen)
# =============================
$script:restCode = @'
function OV-GetApiVersion {
    param([string]$A)
    $r = Invoke-RestMethod -Uri "https://$A/rest/version" -Method Get -SkipCertificateCheck -ErrorAction Stop
    [int]$r.currentVersion
}
function OV-Login {
    param([string]$A,[string]$U,[string]$P,[int]$V)
    $b = @{userName=$U;password=$P;authLoginDomain="Local"} | ConvertTo-Json
    $h = @{"Content-Type"="application/json";"X-API-Version"="$V"}
    $r = Invoke-RestMethod -Uri "https://$A/rest/login-sessions" -Method Post -Body $b -Headers $h -SkipCertificateCheck -ErrorAction Stop
    if ([string]::IsNullOrEmpty($r.sessionID)) { throw "Keine sessionID erhalten von $A" }
    $r.sessionID
}
function OV-Logout {
    param([string]$A,[string]$S,[int]$V)
    $h = @{Auth=$S;"X-API-Version"="$V"}
    try { Invoke-RestMethod -Uri "https://$A/rest/login-sessions" -Method Delete -Headers $h -SkipCertificateCheck -EA SilentlyContinue } catch {}
}
function OV-Rest {
    param([string]$A,[string]$S,[int]$V,[string]$M,[string]$E,$Body)
    $h = @{Auth=$S;"X-API-Version"="$V"}
    $p = @{Uri="https://$A$E";Method=$M;Headers=$h;ContentType="application/json";SkipCertificateCheck=$true;ErrorAction="Stop"}
    if ($Body) { $p.Body = (ConvertTo-Json -InputObject $Body -Depth 10) }
    Invoke-RestMethod @p
}
'@

# =============================
# Haupt-Formular
# =============================
$form = New-Object System.Windows.Forms.Form
$null = $form.Handle
$form.Text = "© 2025 N.J. Airbus D&S - HPE OneView User Manager"
$form.Size = New-Object System.Drawing.Size(960,1060)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI",9)

$boldFont = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)

# ─────────────────────────────────────────
# Oberer Bereich: Admin-Login
# ─────────────────────────────────────────
$lblAdminUser = New-Object System.Windows.Forms.Label
$lblAdminUser.Location = '10,15'; $lblAdminUser.Size = '100,20'; $lblAdminUser.Text = "Admin Login:"; $lblAdminUser.Font = $boldFont
$form.Controls.Add($lblAdminUser)

$txtAdminUser = New-Object System.Windows.Forms.TextBox
$txtAdminUser.Location = '115,13'; $txtAdminUser.Size = '150,22'; $txtAdminUser.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtAdminUser)

$lblAdminPass = New-Object System.Windows.Forms.Label
$lblAdminPass.Location = '280,15'; $lblAdminPass.Size = '70,20'; $lblAdminPass.Text = "Passwort:"
$form.Controls.Add($lblAdminPass)

$txtAdminPass = New-Object System.Windows.Forms.TextBox
$txtAdminPass.Location = '355,13'; $txtAdminPass.Size = '150,22'; $txtAdminPass.UseSystemPasswordChar = $true; $txtAdminPass.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtAdminPass)

# ─────────────────────────────────────────
# IP-Dateien
# ─────────────────────────────────────────
$lblIP660 = New-Object System.Windows.Forms.Label
$lblIP660.Location = '10,47'; $lblIP660.Size = '110,20'; $lblIP660.Text = "OV 6.60 IP-Datei:"; $lblIP660.Font = $boldFont
$form.Controls.Add($lblIP660)

$txtIP660 = New-Object System.Windows.Forms.TextBox
$txtIP660.Location = '125,45'; $txtIP660.Size = '580,22'; $txtIP660.Text = (Join-Path $scriptFolder "Oneview_660.txt"); $txtIP660.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtIP660)

$btnBrowse660 = New-Object System.Windows.Forms.Button
$btnBrowse660.Location = '715,44'; $btnBrowse660.Size = '75,24'; $btnBrowse660.Text = "Browse..."
$form.Controls.Add($btnBrowse660)

$lblIP1110 = New-Object System.Windows.Forms.Label
$lblIP1110.Location = '10,75'; $lblIP1110.Size = '110,20'; $lblIP1110.Text = "OV 11.10 IP-Datei:"; $lblIP1110.Font = $boldFont
$form.Controls.Add($lblIP1110)

$txtIP1110 = New-Object System.Windows.Forms.TextBox
$txtIP1110.Location = '125,73'; $txtIP1110.Size = '580,22'; $txtIP1110.Text = (Join-Path $scriptFolder "Oneview.txt"); $txtIP1110.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtIP1110)

$btnBrowse1110 = New-Object System.Windows.Forms.Button
$btnBrowse1110.Location = '715,72'; $btnBrowse1110.Size = '75,24'; $btnBrowse1110.Text = "Browse..."
$form.Controls.Add($btnBrowse1110)

# ─────────────────────────────────────────
# Appliance-Auswahl (CheckedListBox)
# ─────────────────────────────────────────
$lblApplSel = New-Object System.Windows.Forms.Label
$lblApplSel.Location = '10,105'; $lblApplSel.Size = '140,20'; $lblApplSel.Text = "Appliance-Auswahl:"; $lblApplSel.Font = $boldFont
$form.Controls.Add($lblApplSel)

$btnSelAll = New-Object System.Windows.Forms.Button
$btnSelAll.Location = '155,102'; $btnSelAll.Size = '60,24'; $btnSelAll.Text = "Alle"
$form.Controls.Add($btnSelAll)

$btnSelNone = New-Object System.Windows.Forms.Button
$btnSelNone.Location = '222,102'; $btnSelNone.Size = '60,24'; $btnSelNone.Text = "Keine"
$form.Controls.Add($btnSelNone)

$chkAppliances = New-Object System.Windows.Forms.CheckedListBox
$chkAppliances.Location = '10,130'; $chkAppliances.Size = '920,195'; $chkAppliances.CheckOnClick = $true; $chkAppliances.BorderStyle = 'FixedSingle'
$form.Controls.Add($chkAppliances)

# Appliance-Lade-Funktion
function Load-Appliances {
    $chkAppliances.Items.Clear()
    if (-not [string]::IsNullOrWhiteSpace($txtIP660.Text) -and (Test-Path $txtIP660.Text)) {
        @(Get-Content $txtIP660.Text | Where-Object { $_.Trim() -ne '' }) | ForEach-Object {
            $chkAppliances.Items.Add("$($_.Trim())   (OV 6.60)", $false) | Out-Null
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($txtIP1110.Text) -and (Test-Path $txtIP1110.Text)) {
        @(Get-Content $txtIP1110.Text | Where-Object { $_.Trim() -ne '' }) | ForEach-Object {
            $chkAppliances.Items.Add("$($_.Trim())   (OV 11.10)", $false) | Out-Null
        }
    }
    Update-ApplianceComboBoxes
}

$btnSelAll.Add_Click({
    for ($i = 0; $i -lt $chkAppliances.Items.Count; $i++) { $chkAppliances.SetItemChecked($i, $true) }
})
$btnSelNone.Add_Click({
    for ($i = 0; $i -lt $chkAppliances.Items.Count; $i++) { $chkAppliances.SetItemChecked($i, $false) }
})
$btnBrowse660.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "Textdateien (*.txt)|*.txt|Alle (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') { $txtIP660.Text = $ofd.FileName; Load-Appliances }
})
$btnBrowse1110.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "Textdateien (*.txt)|*.txt|Alle (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') { $txtIP1110.Text = $ofd.FileName; Load-Appliances }
})

# Hilfsfunktion: IP aus ComboBox-Eintrag extrahieren
function Get-IPFromCombo { param([string]$t); if ($t -match '^\s*(.+?)\s+\(OV') { $Matches[1] } else { $t.Trim() } }

# Hilfsfunktion: Ausgewählte Appliances als Objekte
function Get-CheckedAppliances {
    $result = @()
    for ($i = 0; $i -lt $chkAppliances.Items.Count; $i++) {
        if ($chkAppliances.GetItemChecked($i)) {
            $t = $chkAppliances.Items[$i].ToString()
            if ($t -match '^\s*(.+?)\s+\(OV (.+?)\)\s*$') {
                $result += @{ IP = $Matches[1]; Version = $Matches[2] }
            }
        }
    }
    ,$result
}

# Hilfsfunktion: ComboBoxen in Tabs aktualisieren
function Update-ApplianceComboBoxes {
    $items = @()
    for ($i = 0; $i -lt $chkAppliances.Items.Count; $i++) {
        if ($chkAppliances.GetItemChecked($i)) {
            $items += $chkAppliances.Items[$i].ToString()
        }
    }
    foreach ($cb in @($cboT1Appl, $cboT3Appl, $cboT4Appl)) {
        if ($null -ne $cb) {
            $prev = $cb.Text
            $cb.Items.Clear()
            foreach ($item in $items) { $cb.Items.Add($item) | Out-Null }
            if ($cb.Items.Count -gt 0) {
                $idx = $cb.Items.IndexOf($prev)
                $cb.SelectedIndex = if ($idx -ge 0) { $idx } else { 0 }
            }
        }
    }
}

# ─────────────────────────────────────────
# TabControl
# ─────────────────────────────────────────
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = '10,333'; $tabControl.Size = '920,465'
$form.Controls.Add($tabControl)

# ═══════════════════════════════════════════
# TAB 1: Benutzer-Übersicht
# ═══════════════════════════════════════════
$tab1 = New-Object System.Windows.Forms.TabPage; $tab1.Text = "Benutzer-Übersicht"
$tabControl.TabPages.Add($tab1)

$lblT1Appl = New-Object System.Windows.Forms.Label
$lblT1Appl.Location = '10,12'; $lblT1Appl.Size = '70,20'; $lblT1Appl.Text = "Appliance:"
$tab1.Controls.Add($lblT1Appl)

$cboT1Appl = New-Object System.Windows.Forms.ComboBox
$cboT1Appl.Location = '85,10'; $cboT1Appl.Size = '350,23'; $cboT1Appl.DropDownStyle = 'DropDownList'
$tab1.Controls.Add($cboT1Appl)

$btnT1Load = New-Object System.Windows.Forms.Button
$btnT1Load.Location = '445,9'; $btnT1Load.Size = '100,25'; $btnT1Load.Text = "User laden"
$tab1.Controls.Add($btnT1Load)

$dgvT1 = New-Object System.Windows.Forms.DataGridView
$dgvT1.Location = '10,42'; $dgvT1.Size = '890,385'
$dgvT1.AllowUserToAddRows = $false; $dgvT1.AllowUserToDeleteRows = $false
$dgvT1.ReadOnly = $true; $dgvT1.SelectionMode = 'FullRowSelect'
$dgvT1.AutoSizeColumnsMode = 'Fill'; $dgvT1.RowHeadersVisible = $false
$dgvT1.BorderStyle = 'FixedSingle'
$dgvT1.Columns.Add("userName","Benutzername") | Out-Null
$dgvT1.Columns.Add("fullName","Vollständiger Name") | Out-Null
$dgvT1.Columns.Add("emailAddress","E-Mail") | Out-Null
$dgvT1.Columns.Add("enabled","Aktiv") | Out-Null
$dgvT1.Columns.Add("roles","Rollen") | Out-Null
$dgvT1.Columns["userName"].FillWeight = 15
$dgvT1.Columns["fullName"].FillWeight = 20
$dgvT1.Columns["emailAddress"].FillWeight = 25
$dgvT1.Columns["enabled"].FillWeight = 8
$dgvT1.Columns["roles"].FillWeight = 32
$tab1.Controls.Add($dgvT1)

# ═══════════════════════════════════════════
# TAB 2: Passwort ändern (Multi-Appliance)
# ═══════════════════════════════════════════
$tab2 = New-Object System.Windows.Forms.TabPage; $tab2.Text = "Passwort ändern"
$tabControl.TabPages.Add($tab2)

$lblT2User = New-Object System.Windows.Forms.Label
$lblT2User.Location = '10,12'; $lblT2User.Size = '120,20'; $lblT2User.Text = "Ziel-Username:"; $lblT2User.Font = $boldFont
$tab2.Controls.Add($lblT2User)

$txtT2User = New-Object System.Windows.Forms.TextBox
$txtT2User.Location = '140,10'; $txtT2User.Size = '200,22'; $txtT2User.BorderStyle = 'FixedSingle'
$tab2.Controls.Add($txtT2User)

$lblT2NewPw = New-Object System.Windows.Forms.Label
$lblT2NewPw.Location = '10,42'; $lblT2NewPw.Size = '120,20'; $lblT2NewPw.Text = "Neues Passwort:"
$tab2.Controls.Add($lblT2NewPw)

$txtT2NewPw = New-Object System.Windows.Forms.TextBox
$txtT2NewPw.Location = '140,40'; $txtT2NewPw.Size = '200,22'; $txtT2NewPw.UseSystemPasswordChar = $true; $txtT2NewPw.BorderStyle = 'FixedSingle'
$tab2.Controls.Add($txtT2NewPw)

$lblT2ConfPw = New-Object System.Windows.Forms.Label
$lblT2ConfPw.Location = '10,72'; $lblT2ConfPw.Size = '120,20'; $lblT2ConfPw.Text = "PW bestätigen:"
$tab2.Controls.Add($lblT2ConfPw)

$txtT2ConfPw = New-Object System.Windows.Forms.TextBox
$txtT2ConfPw.Location = '140,70'; $txtT2ConfPw.Size = '200,22'; $txtT2ConfPw.UseSystemPasswordChar = $true; $txtT2ConfPw.BorderStyle = 'FixedSingle'
$tab2.Controls.Add($txtT2ConfPw)

$chkT2Self = New-Object System.Windows.Forms.CheckBox
$chkT2Self.Location = '10,102'; $chkT2Self.Size = '250,20'; $chkT2Self.Text = "Eigenes Passwort ändern (Self-Edit)"
$tab2.Controls.Add($chkT2Self)

$lblT2CurPw = New-Object System.Windows.Forms.Label
$lblT2CurPw.Location = '10,130'; $lblT2CurPw.Size = '120,20'; $lblT2CurPw.Text = "Aktuelles PW:"; $lblT2CurPw.Enabled = $false
$tab2.Controls.Add($lblT2CurPw)

$txtT2CurPw = New-Object System.Windows.Forms.TextBox
$txtT2CurPw.Location = '140,128'; $txtT2CurPw.Size = '200,22'; $txtT2CurPw.UseSystemPasswordChar = $true; $txtT2CurPw.BorderStyle = 'FixedSingle'; $txtT2CurPw.Enabled = $false
$tab2.Controls.Add($txtT2CurPw)

$chkT2Self.Add_CheckedChanged({
    $txtT2CurPw.Enabled = $chkT2Self.Checked
    $lblT2CurPw.Enabled = $chkT2Self.Checked
})

$btnT2Start = New-Object System.Windows.Forms.Button
$btnT2Start.Location = '10,160'; $btnT2Start.Size = '160,26'; $btnT2Start.Text = "Passwort ändern"; $btnT2Start.Font = $boldFont
$tab2.Controls.Add($btnT2Start)

$lvT2 = New-Object System.Windows.Forms.ListView
$lvT2.Location = '10,195'; $lvT2.Size = '890,230'; $lvT2.View = 'Details'; $lvT2.FullRowSelect = $true; $lvT2.GridLines = $true; $lvT2.BorderStyle = 'FixedSingle'
$lvT2.Columns.Add("Appliance",200) | Out-Null
$lvT2.Columns.Add("Version",80) | Out-Null
$lvT2.Columns.Add("Status",100) | Out-Null
$lvT2.Columns.Add("Details",500) | Out-Null
$tab2.Controls.Add($lvT2)

# ═══════════════════════════════════════════
# TAB 3: Benutzer verwalten
# ═══════════════════════════════════════════
$tab3 = New-Object System.Windows.Forms.TabPage; $tab3.Text = "Benutzer verwalten"
$tab3.AutoScroll = $true
$tabControl.TabPages.Add($tab3)

# --- GroupBox: Neuen Benutzer anlegen ---
$grpT3New = New-Object System.Windows.Forms.GroupBox
$grpT3New.Location = '10,5'; $grpT3New.Size = '890,280'; $grpT3New.Text = "Neuen Benutzer anlegen (auf allen ausgewählten Appliances)"
$tab3.Controls.Add($grpT3New)

$lblT3nUser = New-Object System.Windows.Forms.Label; $lblT3nUser.Location = '10,25'; $lblT3nUser.Size = '80,20'; $lblT3nUser.Text = "Username:"
$grpT3New.Controls.Add($lblT3nUser)
$txtT3nUser = New-Object System.Windows.Forms.TextBox; $txtT3nUser.Location = '95,23'; $txtT3nUser.Size = '150,22'; $txtT3nUser.BorderStyle = 'FixedSingle'
$grpT3New.Controls.Add($txtT3nUser)

$lblT3nFull = New-Object System.Windows.Forms.Label; $lblT3nFull.Location = '260,25'; $lblT3nFull.Size = '80,20'; $lblT3nFull.Text = "Full Name:"
$grpT3New.Controls.Add($lblT3nFull)
$txtT3nFull = New-Object System.Windows.Forms.TextBox; $txtT3nFull.Location = '345,23'; $txtT3nFull.Size = '180,22'; $txtT3nFull.BorderStyle = 'FixedSingle'
$grpT3New.Controls.Add($txtT3nFull)

$lblT3nMail = New-Object System.Windows.Forms.Label; $lblT3nMail.Location = '10,55'; $lblT3nMail.Size = '80,20'; $lblT3nMail.Text = "E-Mail:"
$grpT3New.Controls.Add($lblT3nMail)
$txtT3nMail = New-Object System.Windows.Forms.TextBox; $txtT3nMail.Location = '95,53'; $txtT3nMail.Size = '200,22'; $txtT3nMail.BorderStyle = 'FixedSingle'
$grpT3New.Controls.Add($txtT3nMail)

$lblT3nPw = New-Object System.Windows.Forms.Label; $lblT3nPw.Location = '310,55'; $lblT3nPw.Size = '80,20'; $lblT3nPw.Text = "Passwort:"
$grpT3New.Controls.Add($lblT3nPw)
$txtT3nPw = New-Object System.Windows.Forms.TextBox; $txtT3nPw.Location = '395,53'; $txtT3nPw.Size = '150,22'; $txtT3nPw.UseSystemPasswordChar = $true; $txtT3nPw.BorderStyle = 'FixedSingle'
$grpT3New.Controls.Add($txtT3nPw)

$chkT3nEnabled = New-Object System.Windows.Forms.CheckBox
$chkT3nEnabled.Location = '560,54'; $chkT3nEnabled.Size = '80,22'; $chkT3nEnabled.Text = "Aktiv"; $chkT3nEnabled.Checked = $true
$grpT3New.Controls.Add($chkT3nEnabled)

$lblT3nRoles = New-Object System.Windows.Forms.Label; $lblT3nRoles.Location = '10,88'; $lblT3nRoles.Size = '80,20'; $lblT3nRoles.Text = "Rollen:"; $lblT3nRoles.Font = $boldFont
$grpT3New.Controls.Add($lblT3nRoles)
$btnT3nLoadRoles = New-Object System.Windows.Forms.Button
$btnT3nLoadRoles.Location = '95,85'; $btnT3nLoadRoles.Size = '200,25'; $btnT3nLoadRoles.Text = "Verfügbare Rollen laden"
$grpT3New.Controls.Add($btnT3nLoadRoles)
$chkT3nRoles = New-Object System.Windows.Forms.CheckedListBox
$chkT3nRoles.Location = '10,115'; $chkT3nRoles.Size = '870,120'; $chkT3nRoles.CheckOnClick = $true; $chkT3nRoles.MultiColumn = $true; $chkT3nRoles.ColumnWidth = 280
$grpT3New.Controls.Add($chkT3nRoles)

$btnT3Create = New-Object System.Windows.Forms.Button
$btnT3Create.Location = '10,242'; $btnT3Create.Size = '200,26'; $btnT3Create.Text = "Benutzer anlegen"; $btnT3Create.Font = $boldFont
$grpT3New.Controls.Add($btnT3Create)

# --- GroupBox: Benutzer bearbeiten / löschen ---
$grpT3Edit = New-Object System.Windows.Forms.GroupBox
$grpT3Edit.Location = '10,290'; $grpT3Edit.Size = '890,195'; $grpT3Edit.Text = "Benutzer bearbeiten / löschen"
$tab3.Controls.Add($grpT3Edit)

$lblT3eAppl = New-Object System.Windows.Forms.Label; $lblT3eAppl.Location = '10,25'; $lblT3eAppl.Size = '70,20'; $lblT3eAppl.Text = "Appliance:"
$grpT3Edit.Controls.Add($lblT3eAppl)

$cboT3Appl = New-Object System.Windows.Forms.ComboBox
$cboT3Appl.Location = '85,23'; $cboT3Appl.Size = '300,23'; $cboT3Appl.DropDownStyle = 'DropDownList'
$grpT3Edit.Controls.Add($cboT3Appl)

$btnT3LoadUsers = New-Object System.Windows.Forms.Button
$btnT3LoadUsers.Location = '395,22'; $btnT3LoadUsers.Size = '100,25'; $btnT3LoadUsers.Text = "User laden"
$grpT3Edit.Controls.Add($btnT3LoadUsers)

$lblT3eUser = New-Object System.Windows.Forms.Label; $lblT3eUser.Location = '10,55'; $lblT3eUser.Size = '70,20'; $lblT3eUser.Text = "User:"
$grpT3Edit.Controls.Add($lblT3eUser)

$cboT3User = New-Object System.Windows.Forms.ComboBox
$cboT3User.Location = '85,53'; $cboT3User.Size = '200,23'; $cboT3User.DropDownStyle = 'DropDownList'
$grpT3Edit.Controls.Add($cboT3User)

$btnT3LoadDetail = New-Object System.Windows.Forms.Button
$btnT3LoadDetail.Location = '295,52'; $btnT3LoadDetail.Size = '100,25'; $btnT3LoadDetail.Text = "Details laden"
$grpT3Edit.Controls.Add($btnT3LoadDetail)

$lblT3eFull = New-Object System.Windows.Forms.Label; $lblT3eFull.Location = '10,85'; $lblT3eFull.Size = '70,20'; $lblT3eFull.Text = "Full Name:"
$grpT3Edit.Controls.Add($lblT3eFull)
$txtT3eFull = New-Object System.Windows.Forms.TextBox; $txtT3eFull.Location = '85,83'; $txtT3eFull.Size = '200,22'; $txtT3eFull.BorderStyle = 'FixedSingle'
$grpT3Edit.Controls.Add($txtT3eFull)

$lblT3eMail = New-Object System.Windows.Forms.Label; $lblT3eMail.Location = '300,85'; $lblT3eMail.Size = '50,20'; $lblT3eMail.Text = "E-Mail:"
$grpT3Edit.Controls.Add($lblT3eMail)
$txtT3eMail = New-Object System.Windows.Forms.TextBox; $txtT3eMail.Location = '355,83'; $txtT3eMail.Size = '200,22'; $txtT3eMail.BorderStyle = 'FixedSingle'
$grpT3Edit.Controls.Add($txtT3eMail)

$chkT3eEnabled = New-Object System.Windows.Forms.CheckBox
$chkT3eEnabled.Location = '570,84'; $chkT3eEnabled.Size = '80,22'; $chkT3eEnabled.Text = "Aktiv"
$grpT3Edit.Controls.Add($chkT3eEnabled)

$btnT3Save = New-Object System.Windows.Forms.Button
$btnT3Save.Location = '10,118'; $btnT3Save.Size = '180,26'; $btnT3Save.Text = "Änderungen speichern"
$grpT3Edit.Controls.Add($btnT3Save)

$btnT3Delete = New-Object System.Windows.Forms.Button
$btnT3Delete.Location = '200,118'; $btnT3Delete.Size = '260,26'; $btnT3Delete.Text = "User auf ALLEN Appliances löschen"
$btnT3Delete.ForeColor = [System.Drawing.Color]::DarkRed
$grpT3Edit.Controls.Add($btnT3Delete)

$lblT3Hint = New-Object System.Windows.Forms.Label
$lblT3Hint.Location = '10,152'; $lblT3Hint.Size = '860,40'
$lblT3Hint.Text = "Hinweis: 'Benutzer anlegen' erstellt den User auf allen oben ausgewählten Appliances. 'Änderungen speichern' wirkt nur auf die einzelne Appliance. 'User löschen' löscht auf allen ausgewählten Appliances."
$lblT3Hint.ForeColor = [System.Drawing.Color]::Gray
$grpT3Edit.Controls.Add($lblT3Hint)

# ═══════════════════════════════════════════
# TAB 4: Rollen verwalten
# ═══════════════════════════════════════════
$tab4 = New-Object System.Windows.Forms.TabPage; $tab4.Text = "Rollen verwalten"
$tabControl.TabPages.Add($tab4)

$lblT4Appl = New-Object System.Windows.Forms.Label; $lblT4Appl.Location = '10,12'; $lblT4Appl.Size = '70,20'; $lblT4Appl.Text = "Appliance:"
$tab4.Controls.Add($lblT4Appl)

$cboT4Appl = New-Object System.Windows.Forms.ComboBox
$cboT4Appl.Location = '85,10'; $cboT4Appl.Size = '300,23'; $cboT4Appl.DropDownStyle = 'DropDownList'
$tab4.Controls.Add($cboT4Appl)

$btnT4LoadUsers = New-Object System.Windows.Forms.Button
$btnT4LoadUsers.Location = '395,9'; $btnT4LoadUsers.Size = '100,25'; $btnT4LoadUsers.Text = "User laden"
$tab4.Controls.Add($btnT4LoadUsers)

$lblT4User = New-Object System.Windows.Forms.Label; $lblT4User.Location = '10,42'; $lblT4User.Size = '70,20'; $lblT4User.Text = "User:"
$tab4.Controls.Add($lblT4User)

$cboT4User = New-Object System.Windows.Forms.ComboBox
$cboT4User.Location = '85,40'; $cboT4User.Size = '200,23'; $cboT4User.DropDownStyle = 'DropDownList'
$tab4.Controls.Add($cboT4User)

$btnT4LoadRoles = New-Object System.Windows.Forms.Button
$btnT4LoadRoles.Location = '295,39'; $btnT4LoadRoles.Size = '120,25'; $btnT4LoadRoles.Text = "Rollen laden"
$tab4.Controls.Add($btnT4LoadRoles)

$lblT4Roles = New-Object System.Windows.Forms.Label
$lblT4Roles.Location = '10,72'; $lblT4Roles.Size = '400,20'; $lblT4Roles.Text = "Verfügbare Rollen (Haken = dem User zugewiesen):"; $lblT4Roles.Font = $boldFont
$tab4.Controls.Add($lblT4Roles)

$chkT4Roles = New-Object System.Windows.Forms.CheckedListBox
$chkT4Roles.Location = '10,95'; $chkT4Roles.Size = '500,200'; $chkT4Roles.CheckOnClick = $true; $chkT4Roles.BorderStyle = 'FixedSingle'
$tab4.Controls.Add($chkT4Roles)

$btnT4Replace = New-Object System.Windows.Forms.Button
$btnT4Replace.Location = '10,305'; $btnT4Replace.Size = '200,26'; $btnT4Replace.Text = "Rollen komplett ersetzen"; $btnT4Replace.Font = $boldFont
$tab4.Controls.Add($btnT4Replace)

$btnT4Add = New-Object System.Windows.Forms.Button
$btnT4Add.Location = '220,305'; $btnT4Add.Size = '200,26'; $btnT4Add.Text = "Markierte hinzufügen"
$tab4.Controls.Add($btnT4Add)

$btnT4Remove = New-Object System.Windows.Forms.Button
$btnT4Remove.Location = '430,305'; $btnT4Remove.Size = '200,26'; $btnT4Remove.Text = "Markierte entfernen"
$tab4.Controls.Add($btnT4Remove)

$chkT4Multi = New-Object System.Windows.Forms.CheckBox
$chkT4Multi.Location = '10,340'; $chkT4Multi.Size = '500,22'
$chkT4Multi.Text = "Änderungen auf ALLE ausgewählten Appliances anwenden (nicht nur die oben gewählte)"
$tab4.Controls.Add($chkT4Multi)

# ─────────────────────────────────────────
# Log-Bereich
# ─────────────────────────────────────────
$panelLog = New-Object System.Windows.Forms.Panel
$panelLog.Location = '10,805'; $panelLog.Size = '920,150'; $panelLog.BorderStyle = 'FixedSingle'
$form.Controls.Add($panelLog)

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Dock = 'Fill'; $logBox.ReadOnly = $true; $logBox.BorderStyle = 'None'
$logBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$panelLog.Controls.Add($logBox)

# StatusStrip
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.Dock = 'Bottom'
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel; $statusLabel.Text = "Bereit"
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

# Exit-Button
$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Location = '10,965'; $btnExit.Size = '100,26'; $btnExit.Text = "Exit"
$form.Controls.Add($btnExit)
$btnExit.Add_Click({ $form.Close() })

# ─────────────────────────────────────────
# Async-Engine: ConcurrentQueue + Timer
# ─────────────────────────────────────────
$script:uiQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
$script:guiTimer = New-Object System.Windows.Forms.Timer
$script:guiTimer.Interval = 200
$script:guiTimer.Add_Tick({
    $msg = $null
    while ($script:uiQueue.TryDequeue([ref]$msg)) {
        switch ($msg.Type) {
            'LOG' {
                $logBox.AppendText("$($msg.Text)`r`n"); $logBox.ScrollToCaret()
            }
            'STATUS' {
                $statusLabel.Text = $msg.Text
            }
            'PROGRESS' {
                $logBox.AppendText("[$($msg.Counter)/$($msg.Total)] $($msg.Appliance)...`r`n"); $logBox.ScrollToCaret()
                $statusLabel.Text = "Appliance $($msg.Counter) von $($msg.Total)"
                $li = New-Object System.Windows.Forms.ListViewItem($msg.Appliance)
                $li.Name = $msg.Appliance
                $li.SubItems.Add($msg.Version) | Out-Null
                $li.SubItems.Add("Wird verarbeitet...") | Out-Null
                $li.SubItems.Add("") | Out-Null
                $lvT2.Items.Add($li) | Out-Null; $li.EnsureVisible()
            }
            'UPDATE' {
                $logBox.AppendText("  → $($msg.Appliance): $($msg.Status) – $($msg.Detail)`r`n"); $logBox.ScrollToCaret()
                $li = $lvT2.Items[$msg.Appliance]
                if ($li) { $li.SubItems[2].Text = $msg.Status; $li.SubItems[3].Text = $msg.Detail; $li.EnsureVisible() }
            }
            'T1_USERLIST' {
                $dgvT1.Rows.Clear()
                foreach ($u in $msg.Data) {
                    $roles = ''
                    if ($u.permissions) { $roles = ($u.permissions | ForEach-Object { $_.roleName }) -join ', ' }
                    $dgvT1.Rows.Add($u.userName, $u.fullName, $u.emailAddress, $u.enabled, $roles) | Out-Null
                }
                $logBox.AppendText("$($msg.Data.Count) Benutzer geladen.`r`n"); $logBox.ScrollToCaret()
            }
            'T3_USERLIST' {
                $cboT3User.Items.Clear()
                foreach ($u in $msg.Data) { $cboT3User.Items.Add($u.userName) | Out-Null }
                if ($cboT3User.Items.Count -gt 0) { $cboT3User.SelectedIndex = 0 }
                $logBox.AppendText("$($msg.Data.Count) User für Bearbeitung geladen.`r`n"); $logBox.ScrollToCaret()
            }
            'T3_USERDETAIL' {
                $txtT3eFull.Text = $msg.Data.fullName
                $txtT3eMail.Text = $msg.Data.emailAddress
                $chkT3eEnabled.Checked = [bool]$msg.Data.enabled
                $logBox.AppendText("Details für '$($msg.Data.userName)' geladen.`r`n"); $logBox.ScrollToCaret()
            }
            'T4_USERLIST' {
                $cboT4User.Items.Clear()
                foreach ($u in $msg.Data) { $cboT4User.Items.Add($u.userName) | Out-Null }
                if ($cboT4User.Items.Count -gt 0) { $cboT4User.SelectedIndex = 0 }
            }
            'T3_ROLES' {
                $chkT3nRoles.Items.Clear()
                foreach ($role in $msg.Available) {
                    $chkT3nRoles.Items.Add($role, $false) | Out-Null
                }
                $logBox.AppendText("$($msg.Available.Count) verfügbare Rollen geladen.`r`n"); $logBox.ScrollToCaret()
            }
            'T4_ROLES' {
                $chkT4Roles.Items.Clear()
                foreach ($role in $msg.Available) {
                    $assigned = $msg.Assigned -contains $role
                    $chkT4Roles.Items.Add($role, $assigned) | Out-Null
                }
                $script:t4OriginalAssignedCount = $msg.Assigned.Count
                $logBox.AppendText("Rollen für '$($msg.UserName)' geladen ($($msg.Assigned.Count) zugewiesen).`r`n"); $logBox.ScrollToCaret()
            }
            'FINISHED' {
                $logBox.AppendText("`r`nVorgang abgeschlossen.`r`n"); $logBox.ScrollToCaret()
                $statusLabel.Text = "Fertig"
                $btnT2Start.Enabled = $true; $btnT3Create.Enabled = $true; $btnT3Save.Enabled = $true
                $btnT3Delete.Enabled = $true; $btnT4Replace.Enabled = $true; $btnT4Add.Enabled = $true
                $btnT4Remove.Enabled = $true
                $btnT1Load.Enabled = $true; $btnT3LoadUsers.Enabled = $true; $btnT3LoadDetail.Enabled = $true
                $btnT4LoadUsers.Enabled = $true; $btnT4LoadRoles.Enabled = $true; $btnT3nLoadRoles.Enabled = $true
            }
            'ERROR' {
                $logBox.SelectionColor = [System.Drawing.Color]::Red
                $logBox.AppendText("FEHLER: $($msg.Text)`r`n"); $logBox.ScrollToCaret()
                $logBox.SelectionColor = $logBox.ForeColor
            }
            'CRITICAL_ERROR' {
                $logBox.SelectionColor = [System.Drawing.Color]::Red
                $logBox.AppendText("KRITISCHER FEHLER: $($msg.Error)`r`n"); $logBox.ScrollToCaret()
                $logBox.SelectionColor = $logBox.ForeColor
                $statusLabel.Text = "Fehler"
                $btnT2Start.Enabled = $true; $btnT3Create.Enabled = $true; $btnT3Save.Enabled = $true
                $btnT3Delete.Enabled = $true; $btnT4Replace.Enabled = $true; $btnT4Add.Enabled = $true
                $btnT4Remove.Enabled = $true
                $btnT1Load.Enabled = $true; $btnT3LoadUsers.Enabled = $true; $btnT3LoadDetail.Enabled = $true
                $btnT4LoadUsers.Enabled = $true; $btnT4LoadRoles.Enabled = $true; $btnT3nLoadRoles.Enabled = $true
            }
        }
    }
})
$script:guiTimer.Start()

# Hilfsfunktion: Async-Operation starten
function Start-AsyncOp {
    param([scriptblock]$Block, [object[]]$Arguments, [hashtable]$Params)
    $ps = [powershell]::Create()
    $ps.AddScript($Block) | Out-Null
    if ($Params) {
        $ps.AddArgument($Params) | Out-Null
    }
    elseif ($Arguments) {
        foreach ($a in $Arguments) { $ps.AddArgument($a) | Out-Null }
    }
    $null = $ps.BeginInvoke()
}

# ─────────────────────────────────────────
# ComboBox-Update bei Tab-Wechsel
# ─────────────────────────────────────────
$tabControl.Add_SelectedIndexChanged({ Update-ApplianceComboBoxes })
$chkAppliances.Add_ItemCheck({
    # ItemCheck feuert VOR der Änderung, daher verzögert aktualisieren
    $form.BeginInvoke([Action]{ Update-ApplianceComboBoxes })
})

# ═══════════════════════════════════════════════════════════════════
#  EVENT-HANDLER: Tab 1 – Benutzer laden
# ═══════════════════════════════════════════════════════════════════
$btnT1Load.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ($cboT1Appl.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Appliance auswählen.", "Hinweis", 'OK', 'Warning'); return
    }
    $btnT1Load.Enabled = $false
    $appliance = Get-IPFromCombo $cboT1Appl.SelectedItem.ToString()
    $logBox.AppendText("Lade Benutzer von $appliance...`r`n"); $logBox.ScrollToCaret()

    Start-AsyncOp -Block {
        param($restCode, $appliance, $adminUser, $adminPass, $uiQueue)
        Invoke-Expression $restCode
        try {
            $v = OV-GetApiVersion -A $appliance
            $uiQueue.Enqueue(@{ Type='LOG'; Text="API-Version $appliance = $v" })
            $uiQueue.Enqueue(@{ Type='LOG'; Text="Login auf $appliance..." })
            $s = OV-Login -A $appliance -U $adminUser -P $adminPass -V $v
            $uiQueue.Enqueue(@{ Type='LOG'; Text="Login erfolgreich. Lade Benutzerliste..." })
            $resp = OV-Rest -A $appliance -S $s -V $v -M Get -E "/rest/users?start=0&count=1000"
            OV-Logout -A $appliance -S $s -V $v
            $users = if ($resp.members) { $resp.members } elseif ($resp -is [array]) { $resp } else { @($resp) }
            $users = @($users | Where-Object { $_.userName -ne 'administrator' })
            $uiQueue.Enqueue(@{ Type='T1_USERLIST'; Data=$users })
        }
        catch {
            $uiQueue.Enqueue(@{ Type='ERROR'; Text="$appliance – $($_.Exception.Message)" })
        }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Arguments @($script:restCode, $appliance, $txtAdminUser.Text, $txtAdminPass.Text, $script:uiQueue)
})

# ═══════════════════════════════════════════════════════════════════
#  EVENT-HANDLER: Tab 2 – Passwort ändern (Multi-Appliance)
# ═══════════════════════════════════════════════════════════════════
$btnT2Start.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    $target = $txtT2User.Text; $newPw = $txtT2NewPw.Text; $confPw = $txtT2ConfPw.Text
    $isSelf = $chkT2Self.Checked; $curPw = $txtT2CurPw.Text

    if ([string]::IsNullOrEmpty($target)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Ziel-Username angeben.", "Fehler", 'OK', 'Error'); return
    }
    if ($target -eq 'administrator') {
        [System.Windows.Forms.MessageBox]::Show("Der Administrator-Account kann nicht per Script geändert werden.", "Geschützt", 'OK', 'Warning'); return
    }
    if ([string]::IsNullOrEmpty($newPw) -or $newPw -ne $confPw) {
        [System.Windows.Forms.MessageBox]::Show("Passwörter fehlen oder stimmen nicht überein.", "Fehler", 'OK', 'Error'); return
    }
    if ($newPw -match '[<>;,"''&\\/|+:=\s]') {
        [System.Windows.Forms.MessageBox]::Show("Passwort enthält ungültige Zeichen.`nVerboten: < > ; , `" ' & \ / | + : = und Leerzeichen", "Fehler", 'OK', 'Error'); return
    }
    if ($isSelf -and [string]::IsNullOrEmpty($curPw)) {
        [System.Windows.Forms.MessageBox]::Show("Für Self-Edit ist das aktuelle Passwort erforderlich.", "Fehler", 'OK', 'Error'); return
    }
    $appliances = Get-CheckedAppliances
    if ($appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Appliances ausgewählt.", "Fehler", 'OK', 'Warning'); return
    }
    $res = [System.Windows.Forms.MessageBox]::Show(
        "Passwort für '$target' auf $($appliances.Count) Appliance(s) ändern?", "Bestätigung", 'YesNo', 'Warning')
    if ($res -ne 'Yes') { return }

    $btnT2Start.Enabled = $false; $lvT2.Items.Clear()
    $logBox.AppendText("=== Passwort-Änderung gestartet ===" + "`r`n"); $logBox.ScrollToCaret()

    Start-AsyncOp -Block {
        param($p)
        $restCode=$p.restCode; $appliances=$p.appliances; $adminUser=$p.adminUser; $adminPass=$p.adminPass
        $target=$p.target; $newPw=$p.newPw; $isSelf=$p.isSelf; $curPw=$p.curPw; $uiQueue=$p.uiQueue
        Invoke-Expression $restCode
        try {
            $total = $appliances.Count; $c = 0
            foreach ($entry in $appliances) {
                $c++; $ip = $entry.IP; $ver = $entry.Version
                $uiQueue.Enqueue(@{ Type='PROGRESS'; Counter=$c; Total=$total; Appliance=$ip; Version=$ver })
                try {
                    $v = OV-GetApiVersion -A $ip
                    $s = OV-Login -A $ip -U $adminUser -P $adminPass -V $v
                    # GET vollständiges User-Objekt, dann password setzen und PUT
                    $user = OV-Rest -A $ip -S $s -V $v -M Get -E "/rest/users/$target"
                    $user | Add-Member -NotePropertyName password -NotePropertyValue $newPw -Force
                    if ($isSelf) { $user | Add-Member -NotePropertyName currentPassword -NotePropertyValue $curPw -Force }
                    OV-Rest -A $ip -S $s -V $v -M Put -E "/rest/users" -Body $user
                    OV-Logout -A $ip -S $s -V $v
                    $uiQueue.Enqueue(@{ Type='UPDATE'; Appliance=$ip; Status='Erfolgreich'; Detail="Passwort geändert" })
                }
                catch {
                    $uiQueue.Enqueue(@{ Type='UPDATE'; Appliance=$ip; Status='Fehler'; Detail=$_.Exception.Message })
                }
            }
        }
        catch { $uiQueue.Enqueue(@{ Type='CRITICAL_ERROR'; Error=$_.Exception.Message }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Params @{
        restCode=$script:restCode; appliances=$appliances; adminUser=$txtAdminUser.Text
        adminPass=$txtAdminPass.Text; target=$target; newPw=$newPw; isSelf=$isSelf; curPw=$curPw
        uiQueue=$script:uiQueue
    }
})

# ═══════════════════════════════════════════════════════════════════
#  EVENT-HANDLER: Tab 3 – User laden (Einzel-Appliance)
# ═══════════════════════════════════════════════════════════════════
$btnT3LoadUsers.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ($cboT3Appl.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Appliance auswählen.", "Hinweis", 'OK', 'Warning'); return
    }
    $btnT3LoadUsers.Enabled = $false
    $appliance = Get-IPFromCombo $cboT3Appl.SelectedItem.ToString()

    Start-AsyncOp -Block {
        param($restCode, $appliance, $adminUser, $adminPass, $uiQueue)
        Invoke-Expression $restCode
        try {
            $v = OV-GetApiVersion -A $appliance
            $s = OV-Login -A $appliance -U $adminUser -P $adminPass -V $v
            $resp = OV-Rest -A $appliance -S $s -V $v -M Get -E "/rest/users?count=1000"
            OV-Logout -A $appliance -S $s -V $v
            $members = @($resp.members | Where-Object { $_.userName -ne 'administrator' })
            $uiQueue.Enqueue(@{ Type='T3_USERLIST'; Data=$members })
        }
        catch { $uiQueue.Enqueue(@{ Type='ERROR'; Text="User laden: $($_.Exception.Message)" }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Arguments @($script:restCode, $appliance, $txtAdminUser.Text, $txtAdminPass.Text, $script:uiQueue)
})

# Tab 3 – User-Details laden
$btnT3LoadDetail.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ($cboT3Appl.SelectedItem -eq $null -or $cboT3User.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Appliance und User auswählen.", "Hinweis", 'OK', 'Warning'); return
    }
    $btnT3LoadDetail.Enabled = $false
    $appliance = Get-IPFromCombo $cboT3Appl.SelectedItem.ToString()
    $userName = $cboT3User.SelectedItem.ToString()

    Start-AsyncOp -Block {
        param($restCode, $appliance, $adminUser, $adminPass, $userName, $uiQueue)
        Invoke-Expression $restCode
        try {
            $v = OV-GetApiVersion -A $appliance
            $s = OV-Login -A $appliance -U $adminUser -P $adminPass -V $v
            $resp = OV-Rest -A $appliance -S $s -V $v -M Get -E "/rest/users/$userName"
            OV-Logout -A $appliance -S $s -V $v
            $uiQueue.Enqueue(@{ Type='T3_USERDETAIL'; Data=$resp })
        }
        catch { $uiQueue.Enqueue(@{ Type='ERROR'; Text="Details laden: $($_.Exception.Message)" }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Arguments @($script:restCode, $appliance, $txtAdminUser.Text, $txtAdminPass.Text, $userName, $script:uiQueue)
})

# Tab 3 – Verfügbare Rollen laden (von erster ausgewählter Appliance)
$btnT3nLoadRoles.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    $appliances = Get-CheckedAppliances
    if ($appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Appliances ausgewählt.", "Fehler", 'OK', 'Warning'); return
    }
    $btnT3nLoadRoles.Enabled = $false
    $firstIP = $appliances[0].IP
    Start-AsyncOp -Block {
        param($restCode, $ip, $adminUser, $adminPass, $uiQueue)
        Invoke-Expression $restCode
        try {
            $v = OV-GetApiVersion -A $ip
            $s = OV-Login -A $ip -U $adminUser -P $adminPass -V $v
            $rolesResp = OV-Rest -A $ip -S $s -V $v -M Get -E "/rest/roles?count=1000"
            $available = @($rolesResp.members | ForEach-Object { $_.roleName } | Sort-Object)
            OV-Logout -A $ip -S $s -V $v
            $uiQueue.Enqueue(@{ Type='T3_ROLES'; Available=$available })
        }
        catch { $uiQueue.Enqueue(@{ Type='ERROR'; Text="Rollen laden: $($_.Exception.Message)" }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Arguments @($script:restCode, $firstIP, $txtAdminUser.Text, $txtAdminPass.Text, $script:uiQueue)
})

# Tab 3 – Benutzer anlegen (Multi-Appliance)
$btnT3Create.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    $un = $txtT3nUser.Text; $fn = $txtT3nFull.Text; $em = $txtT3nMail.Text; $pw = $txtT3nPw.Text; $en = $chkT3nEnabled.Checked
    if ([string]::IsNullOrEmpty($un) -or [string]::IsNullOrEmpty($pw)) {
        [System.Windows.Forms.MessageBox]::Show("Username und Passwort sind Pflichtfelder.", "Fehler", 'OK', 'Error'); return
    }
    if ($pw -match '[<>;,"''&\\/|+:=\s]') {
        [System.Windows.Forms.MessageBox]::Show("Passwort enthält ungültige Zeichen.`nVerboten: < > ; , `" ' & \ / | + : = und Leerzeichen", "Fehler", 'OK', 'Error'); return
    }
    # Gewählte Rollen aus CheckedListBox sammeln
    $selectedRoles = @()
    for ($i = 0; $i -lt $chkT3nRoles.Items.Count; $i++) {
        if ($chkT3nRoles.GetItemChecked($i)) { $selectedRoles += $chkT3nRoles.Items[$i].ToString() }
    }
    if ($selectedRoles.Count -eq 0) {
        # Fallback: "Read only" wenn keine Rollen geladen/gewählt wurden
        $selectedRoles = @("Read only")
    }
    $appliances = Get-CheckedAppliances
    if ($appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Appliances ausgewählt.", "Fehler", 'OK', 'Warning'); return
    }
    $res = [System.Windows.Forms.MessageBox]::Show(
        "Benutzer '$un' auf $($appliances.Count) Appliance(s) anlegen?`nRollen: $($selectedRoles -join ', ')", "Bestätigung", 'YesNo', 'Question')
    if ($res -ne 'Yes') { return }

    $btnT3Create.Enabled = $false
    $logBox.AppendText("=== Benutzer '$un' wird angelegt ===" + "`r`n"); $logBox.ScrollToCaret()

    Start-AsyncOp -Block {
        param($p)
        $restCode=$p.restCode; $appliances=$p.appliances; $adminUser=$p.adminUser; $adminPass=$p.adminPass
        $un=$p.un; $fn=$p.fn; $em=$p.em; $pw=$p.pw; $en=$p.en; $roles=$p.roles; $uiQueue=$p.uiQueue
        Invoke-Expression $restCode
        try {
            foreach ($entry in $appliances) {
                $ip = $entry.IP
                try {
                    $v = OV-GetApiVersion -A $ip
                    $s = OV-Login -A $ip -U $adminUser -P $adminPass -V $v
                    $body = @{
                        type = "UserAndPermissions"
                        userName = $un; fullName = $fn; emailAddress = $em
                        password = $pw; enabled = $en
                        permissions = @($roles | ForEach-Object { @{ roleName = $_ } })
                    }
                    OV-Rest -A $ip -S $s -V $v -M Post -E "/rest/users" -Body $body
                    OV-Logout -A $ip -S $s -V $v
                    $uiQueue.Enqueue(@{ Type='LOG'; Text="$ip – User '$un' erfolgreich angelegt mit Rollen: $($roles -join ', ')" })
                }
                catch {
                    $uiQueue.Enqueue(@{ Type='ERROR'; Text="$ip – User anlegen fehlgeschlagen: $($_.Exception.Message)" })
                }
            }
        }
        catch { $uiQueue.Enqueue(@{ Type='CRITICAL_ERROR'; Error=$_.Exception.Message }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Params @{
        restCode=$script:restCode; appliances=$appliances; adminUser=$txtAdminUser.Text
        adminPass=$txtAdminPass.Text; un=$un; fn=$fn; em=$em; pw=$pw; en=$en
        roles=$selectedRoles; uiQueue=$script:uiQueue
    }
})

# Tab 3 – Benutzer bearbeiten (Einzel-Appliance)
$btnT3Save.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ($cboT3Appl.SelectedItem -eq $null -or $cboT3User.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Appliance und User auswählen.", "Hinweis", 'OK', 'Warning'); return
    }
    $btnT3Save.Enabled = $false
    $appliance = Get-IPFromCombo $cboT3Appl.SelectedItem.ToString()
    $userName = $cboT3User.SelectedItem.ToString()
    $fn = $txtT3eFull.Text; $em = $txtT3eMail.Text; $en = $chkT3eEnabled.Checked

    Start-AsyncOp -Block {
        param($restCode, $appliance, $adminUser, $adminPass, $userName, $fn, $em, $en, $uiQueue)
        Invoke-Expression $restCode
        try {
            $v = OV-GetApiVersion -A $appliance
            $s = OV-Login -A $appliance -U $adminUser -P $adminPass -V $v
            # GET vollständiges User-Objekt, Felder ändern, dann PUT
            $user = OV-Rest -A $appliance -S $s -V $v -M Get -E "/rest/users/$userName"
            $user.fullName = $fn
            $user.emailAddress = $em
            $user.enabled = $en
            OV-Rest -A $appliance -S $s -V $v -M Put -E "/rest/users" -Body $user
            OV-Logout -A $appliance -S $s -V $v
            $uiQueue.Enqueue(@{ Type='LOG'; Text="User '$userName' auf $appliance aktualisiert." })
        }
        catch { $uiQueue.Enqueue(@{ Type='ERROR'; Text="Speichern: $($_.Exception.Message)" }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Arguments @($script:restCode, $appliance, $txtAdminUser.Text, $txtAdminPass.Text, $userName, $fn, $em, $en, $script:uiQueue)
})

# Tab 3 – Benutzer löschen (Multi-Appliance)
$btnT3Delete.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ($cboT3User.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte User zum Löschen auswählen.", "Hinweis", 'OK', 'Warning'); return
    }
    $userName = $cboT3User.SelectedItem.ToString()
    $appliances = Get-CheckedAppliances
    if ($appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Appliances ausgewählt.", "Fehler", 'OK', 'Warning'); return
    }
    $res = [System.Windows.Forms.MessageBox]::Show(
        "⚠ Benutzer '$userName' auf $($appliances.Count) Appliance(s) UNWIDERRUFLICH löschen?",
        "Warnung", 'YesNo', 'Warning')
    if ($res -ne 'Yes') { return }

    $btnT3Delete.Enabled = $false
    $logBox.AppendText("=== Lösche User '$userName' ===" + "`r`n"); $logBox.ScrollToCaret()

    Start-AsyncOp -Block {
        param($p)
        $restCode=$p.restCode; $appliances=$p.appliances; $adminUser=$p.adminUser; $adminPass=$p.adminPass
        $userName=$p.userName; $uiQueue=$p.uiQueue
        Invoke-Expression $restCode
        try {
            foreach ($entry in $appliances) {
                $ip = $entry.IP
                try {
                    $v = OV-GetApiVersion -A $ip
                    $s = OV-Login -A $ip -U $adminUser -P $adminPass -V $v
                    OV-Rest -A $ip -S $s -V $v -M Delete -E "/rest/users/$userName"
                    OV-Logout -A $ip -S $s -V $v
                    $uiQueue.Enqueue(@{ Type='LOG'; Text="$ip – User '$userName' gelöscht." })
                }
                catch {
                    $uiQueue.Enqueue(@{ Type='ERROR'; Text="$ip – Löschen fehlgeschlagen: $($_.Exception.Message)" })
                }
            }
        }
        catch { $uiQueue.Enqueue(@{ Type='CRITICAL_ERROR'; Error=$_.Exception.Message }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Params @{
        restCode=$script:restCode; appliances=$appliances; adminUser=$txtAdminUser.Text
        adminPass=$txtAdminPass.Text; userName=$userName; uiQueue=$script:uiQueue
    }
})

# ═══════════════════════════════════════════════════════════════════
#  EVENT-HANDLER: Tab 4 – Rollen
# ═══════════════════════════════════════════════════════════════════
$btnT4LoadUsers.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ($cboT4Appl.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Appliance auswählen.", "Hinweis", 'OK', 'Warning'); return
    }
    $btnT4LoadUsers.Enabled = $false
    $appliance = Get-IPFromCombo $cboT4Appl.SelectedItem.ToString()

    Start-AsyncOp -Block {
        param($restCode, $appliance, $adminUser, $adminPass, $uiQueue)
        Invoke-Expression $restCode
        try {
            $v = OV-GetApiVersion -A $appliance
            $s = OV-Login -A $appliance -U $adminUser -P $adminPass -V $v
            $resp = OV-Rest -A $appliance -S $s -V $v -M Get -E "/rest/users?count=1000"
            OV-Logout -A $appliance -S $s -V $v
            $members = @($resp.members | Where-Object { $_.userName -ne 'administrator' })
            $uiQueue.Enqueue(@{ Type='T4_USERLIST'; Data=$members })
        }
        catch { $uiQueue.Enqueue(@{ Type='ERROR'; Text="User laden: $($_.Exception.Message)" }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Arguments @($script:restCode, $appliance, $txtAdminUser.Text, $txtAdminPass.Text, $script:uiQueue)
})

# Tab 4 – Rollen laden (inkl. verfügbare Rollen von der Appliance)
$btnT4LoadRoles.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ($cboT4Appl.SelectedItem -eq $null -or $cboT4User.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Appliance und User auswählen.", "Hinweis", 'OK', 'Warning'); return
    }
    $btnT4LoadRoles.Enabled = $false
    $appliance = Get-IPFromCombo $cboT4Appl.SelectedItem.ToString()
    $userName = $cboT4User.SelectedItem.ToString()

    Start-AsyncOp -Block {
        param($restCode, $appliance, $adminUser, $adminPass, $userName, $uiQueue)
        Invoke-Expression $restCode
        try {
            $v = OV-GetApiVersion -A $appliance
            $s = OV-Login -A $appliance -U $adminUser -P $adminPass -V $v
            # Verfügbare Rollen laden
            $rolesResp = OV-Rest -A $appliance -S $s -V $v -M Get -E "/rest/roles?count=1000"
            $available = @($rolesResp.members | ForEach-Object { $_.roleName } | Sort-Object)
            # Zugewiesene Rollen des Users laden
            $userRoles = OV-Rest -A $appliance -S $s -V $v -M Get -E "/rest/users/role/$userName"
            $assigned = @($userRoles.members | ForEach-Object { $_.roleName })
            OV-Logout -A $appliance -S $s -V $v
            $uiQueue.Enqueue(@{ Type='T4_ROLES'; Available=$available; Assigned=$assigned; UserName=$userName })
        }
        catch { $uiQueue.Enqueue(@{ Type='ERROR'; Text="Rollen laden: $($_.Exception.Message)" }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Arguments @($script:restCode, $appliance, $txtAdminUser.Text, $txtAdminPass.Text, $userName, $script:uiQueue)
})

# Tab 4 – Rollen komplett ersetzen
$btnT4Replace.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ($cboT4User.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte User auswählen.", "Hinweis", 'OK', 'Warning'); return
    }
    $userName = $cboT4User.SelectedItem.ToString()
    $selectedRoles = @()
    for ($i = 0; $i -lt $chkT4Roles.Items.Count; $i++) {
        if ($chkT4Roles.GetItemChecked($i)) { $selectedRoles += $chkT4Roles.Items[$i].ToString() }
    }
    if ($selectedRoles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine Rolle markieren.", "Hinweis", 'OK', 'Warning'); return
    }

    $targetAppliances = @()
    if ($chkT4Multi.Checked) {
        $targetAppliances = Get-CheckedAppliances
    } else {
        if ($cboT4Appl.SelectedItem) {
            $ip = Get-IPFromCombo $cboT4Appl.SelectedItem.ToString()
            $ver = if ($cboT4Appl.SelectedItem.ToString() -match '\(OV (.+?)\)') { $Matches[1] } else { "?" }
            $targetAppliances = @(@{ IP=$ip; Version=$ver })
        }
    }
    if ($targetAppliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Appliances ausgewählt.", "Fehler", 'OK', 'Warning'); return
    }

    $res = [System.Windows.Forms.MessageBox]::Show(
        "Rollen von '$userName' auf $($targetAppliances.Count) Appliance(s) ersetzen durch:`n$($selectedRoles -join ', ')?",
        "Bestätigung", 'YesNo', 'Warning')
    if ($res -ne 'Yes') { return }

    $btnT4Replace.Enabled = $false
    Start-AsyncOp -Block {
        param($p)
        $restCode=$p.restCode; $appliances=$p.appliances; $adminUser=$p.adminUser; $adminPass=$p.adminPass
        $userName=$p.userName; $roles=$p.roles; $uiQueue=$p.uiQueue
        Invoke-Expression $restCode
        try {
            foreach ($entry in $appliances) {
                $ip = $entry.IP
                try {
                    $v = OV-GetApiVersion -A $ip
                    $s = OV-Login -A $ip -U $adminUser -P $adminPass -V $v
                    $body = @($roles | ForEach-Object { @{ roleName = $_ } })
                    OV-Rest -A $ip -S $s -V $v -M Put -E "/rest/users/$userName/roles?multiResource=true" -Body $body
                    OV-Logout -A $ip -S $s -V $v
                    $uiQueue.Enqueue(@{ Type='LOG'; Text="$ip – Rollen für '$userName' ersetzt: $($roles -join ', ')" })
                }
                catch {
                    $uiQueue.Enqueue(@{ Type='ERROR'; Text="$ip – Rollen ersetzen: $($_.Exception.Message)" })
                }
            }
        }
        catch { $uiQueue.Enqueue(@{ Type='CRITICAL_ERROR'; Error=$_.Exception.Message }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Params @{
        restCode=$script:restCode; appliances=$targetAppliances; adminUser=$txtAdminUser.Text
        adminPass=$txtAdminPass.Text; userName=$userName; roles=$selectedRoles; uiQueue=$script:uiQueue
    }
})

# Tab 4 – Rollen hinzufügen
$btnT4Add.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ($cboT4User.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte User auswählen.", "Hinweis", 'OK', 'Warning'); return
    }
    $userName = $cboT4User.SelectedItem.ToString()
    $selectedRoles = @()
    for ($i = 0; $i -lt $chkT4Roles.Items.Count; $i++) {
        if ($chkT4Roles.GetItemChecked($i)) { $selectedRoles += $chkT4Roles.Items[$i].ToString() }
    }
    if ($selectedRoles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine Rolle markieren.", "Hinweis", 'OK', 'Warning'); return
    }

    $targetAppliances = @()
    if ($chkT4Multi.Checked) {
        $targetAppliances = Get-CheckedAppliances
    } else {
        if ($cboT4Appl.SelectedItem) {
            $ip = Get-IPFromCombo $cboT4Appl.SelectedItem.ToString()
            $ver = if ($cboT4Appl.SelectedItem.ToString() -match '\(OV (.+?)\)') { $Matches[1] } else { "?" }
            $targetAppliances = @(@{ IP=$ip; Version=$ver })
        }
    }
    if ($targetAppliances.Count -eq 0) { return }

    $btnT4Add.Enabled = $false
    Start-AsyncOp -Block {
        param($p)
        $restCode=$p.restCode; $appliances=$p.appliances; $adminUser=$p.adminUser; $adminPass=$p.adminPass
        $userName=$p.userName; $roles=$p.roles; $uiQueue=$p.uiQueue
        Invoke-Expression $restCode
        try {
            foreach ($entry in $appliances) {
                $ip = $entry.IP
                try {
                    $v = OV-GetApiVersion -A $ip
                    $s = OV-Login -A $ip -U $adminUser -P $adminPass -V $v
                    $body = @($roles | ForEach-Object { @{ roleName = $_ } })
                    OV-Rest -A $ip -S $s -V $v -M Post -E "/rest/users/$userName/roles?multiResource=true" -Body $body
                    OV-Logout -A $ip -S $s -V $v
                    $uiQueue.Enqueue(@{ Type='LOG'; Text="$ip – Rollen für '$userName' hinzugefügt: $($roles -join ', ')" })
                }
                catch {
                    $uiQueue.Enqueue(@{ Type='ERROR'; Text="$ip – Rollen hinzufügen: $($_.Exception.Message)" })
                }
            }
        }
        catch { $uiQueue.Enqueue(@{ Type='CRITICAL_ERROR'; Error=$_.Exception.Message }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Params @{
        restCode=$script:restCode; appliances=$targetAppliances; adminUser=$txtAdminUser.Text
        adminPass=$txtAdminPass.Text; userName=$userName; roles=$selectedRoles; uiQueue=$script:uiQueue
    }
})

# Tab 4 – Markierte Rollen entfernen
$btnT4Remove.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ($cboT4User.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte User auswählen.", "Hinweis", 'OK', 'Warning'); return
    }
    $userName = $cboT4User.SelectedItem.ToString()
    $selectedRolesToRemove = @()
    for ($i = 0; $i -lt $chkT4Roles.Items.Count; $i++) {
        if ($chkT4Roles.GetItemChecked($i)) { $selectedRolesToRemove += $chkT4Roles.Items[$i].ToString() }
    }
    if ($selectedRolesToRemove.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Rollen markiert – bitte Rollen zum Entfernen auswählen.", "Hinweis", 'OK', 'Information'); return
    }
    # Prüfen ob mindestens eine Rolle übrig bleibt
    if ($script:t4OriginalAssignedCount -gt 0 -and $selectedRolesToRemove.Count -ge $script:t4OriginalAssignedCount) {
        [System.Windows.Forms.MessageBox]::Show("Es muss mindestens eine Rolle zugewiesen bleiben. Bitte nicht alle Rollen zum Entfernen markieren.", "Fehler", 'OK', 'Error'); return
    }

    $res = [System.Windows.Forms.MessageBox]::Show(
        "Folgende Rollen von '$userName' entfernen?`n$($selectedRolesToRemove -join ', ')",
        "Bestätigung", 'YesNo', 'Warning')
    if ($res -ne 'Yes') { return }

    $targetAppliances = @()
    if ($chkT4Multi.Checked) {
        $targetAppliances = Get-CheckedAppliances
    } else {
        if ($cboT4Appl.SelectedItem) {
            $ip = Get-IPFromCombo $cboT4Appl.SelectedItem.ToString()
            $ver = if ($cboT4Appl.SelectedItem.ToString() -match '\(OV (.+?)\)') { $Matches[1] } else { "?" }
            $targetAppliances = @(@{ IP=$ip; Version=$ver })
        }
    }
    if ($targetAppliances.Count -eq 0) { return }

    $btnT4Remove.Enabled = $false
    Start-AsyncOp -Block {
        param($p)
        $restCode=$p.restCode; $appliances=$p.appliances; $adminUser=$p.adminUser; $adminPass=$p.adminPass
        $userName=$p.userName; $roles=$p.roles; $uiQueue=$p.uiQueue
        Invoke-Expression $restCode
        try {
            foreach ($entry in $appliances) {
                $ip = $entry.IP
                try {
                    $v = OV-GetApiVersion -A $ip
                    $s = OV-Login -A $ip -U $adminUser -P $adminPass -V $v
                    # Jede zu entfernende Rolle einzeln per DELETE
                    foreach ($r in $roles) {
                        $ep = "/rest/users/roles?filter=`"userName='$userName'`"&filter=`"roleName='$r'`""
                        OV-Rest -A $ip -S $s -V $v -M Delete -E $ep
                    }
                    OV-Logout -A $ip -S $s -V $v
                    $uiQueue.Enqueue(@{ Type='LOG'; Text="$ip – Rollen für '$userName' entfernt: $($roles -join ', ')" })
                }
                catch {
                    $uiQueue.Enqueue(@{ Type='ERROR'; Text="$ip – Rollen entfernen: $($_.Exception.Message)" })
                }
            }
        }
        catch { $uiQueue.Enqueue(@{ Type='CRITICAL_ERROR'; Error=$_.Exception.Message }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Params @{
        restCode=$script:restCode; appliances=$targetAppliances; adminUser=$txtAdminUser.Text
        adminPass=$txtAdminPass.Text; userName=$userName; roles=$selectedRolesToRemove; uiQueue=$script:uiQueue
    }
})

# ─────────────────────────────────────────
# Initial laden + Formular anzeigen
# ─────────────────────────────────────────
$form.Add_Shown({
    Load-Appliances
    Update-ApplianceComboBoxes
})

$form.ShowDialog()
