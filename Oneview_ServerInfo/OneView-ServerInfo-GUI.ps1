#Requires -Version 7.0
# ============================================================================
#  HPE OneView Server-Info GUI
#  Zeigt alle verfuegbaren Hardware-Informationen zu einem Server (anhand
#  Servername oder Seriennummer) versionsuebergreifend an.
#
#  Direkte REST-API - keine HPE PowerShell-Module erforderlich.
#  X-API-Version wird pro Appliance automatisch ermittelt
#  (/rest/version -> currentVersion). Damit funktioniert das Tool fuer
#  OV 6.60 (X-API ~3800/4600), OV 11.10 (~8400), OV 11.20 (~8600) und
#  jede kuenftige Version, ohne Code-Aenderung.
#
#  Datenquellen pro Server:
#    /rest/server-hardware            -> Liste / Detail (Hardware, CPU, RAM,
#                                         iLO, Power, Health, mpFirmware,
#                                         processor*, memory*, locationUri)
#    /rest/server-hardware/{id}/firmware -> Firmware-Inventory aller
#                                         Komponenten (BIOS, iLO, NICs, etc.)
#    /rest/server-profiles            -> zugewiesenes Server-Profile (falls)
#    /rest/enclosures/{id}            -> Frame/Chassis-Name + Bay (Blades)
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
# REST-Helper
# =============================
# Zentraler Timeout (Sekunden) fuer ALLE OneView-REST-Aufrufe.
# Ohne Timeout wartet Invoke-RestMethod in PowerShell 7 unendlich -> die GUI
# friert beim Detail-Laden (Klick auf Server) komplett ein, sobald nur ein
# Endpunkt (iLO/Enclosure/Subressource) nicht antwortet.
$script:OvRestTimeoutSec = 30

function OV-GetApiVersion {
    param([string]$A)
    $r = Invoke-RestMethod -Uri "https://$A/rest/version" -Method Get -SkipCertificateCheck -TimeoutSec $script:OvRestTimeoutSec -ErrorAction Stop
    [int]$r.currentVersion
}
function OV-Login {
    param([string]$A,[string]$U,[string]$P,[int]$V)
    $b = @{ userName = $U; password = $P; authLoginDomain = "Local" } | ConvertTo-Json
    $h = @{ "Content-Type" = "application/json"; "X-API-Version" = "$V" }
    $r = Invoke-RestMethod -Uri "https://$A/rest/login-sessions" -Method Post -Body $b -Headers $h -SkipCertificateCheck -TimeoutSec $script:OvRestTimeoutSec -ErrorAction Stop
    if ([string]::IsNullOrEmpty($r.sessionID)) { throw "Keine sessionID erhalten von $A" }
    $r.sessionID
}
function OV-Logout {
    param([string]$A,[string]$S,[int]$V)
    $h = @{ Auth = $S; "X-API-Version" = "$V" }
    try { Invoke-RestMethod -Uri "https://$A/rest/login-sessions" -Method Delete -Headers $h -SkipCertificateCheck -TimeoutSec 10 -EA SilentlyContinue } catch {}
}
function OV-Rest {
    param([string]$A,[string]$S,[int]$V,[string]$M,[string]$E,[int]$T = 0)
    $h = @{ Auth = $S; "X-API-Version" = "$V" }
    $to = if ($T -gt 0) { $T } else { $script:OvRestTimeoutSec }
    Invoke-RestMethod -Uri "https://$A$E" -Method $M -Headers $h -ContentType "application/json" -SkipCertificateCheck -TimeoutSec $to -ErrorAction Stop
}

# Holt alle Mitglieder einer pageable Collection (members + nextPageUri)
function OV-RestAll {
    param([string]$A,[string]$S,[int]$V,[string]$E)
    $items = @()
    $endpoint = $E
    while ($endpoint) {
        $page = OV-Rest -A $A -S $S -V $V -M Get -E $endpoint
        if ($page.members) { $items += $page.members }
        if ($page.nextPageUri) { $endpoint = $page.nextPageUri } else { $endpoint = $null }
    }
    ,$items
}

# =============================
# Haupt-Formular
# =============================
$form = New-Object System.Windows.Forms.Form
$null = $form.Handle
$form.Text = "© 2025 N.J. Airbus D&S - HPE OneView Server-Info (Auto-Version)"
# An kleine Notebook-Bildschirme anpassen
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$wWidth  = [Math]::Min(1100, $screen.Width  - 40)
$wHeight = [Math]::Min(900,  $screen.Height - 40)
$form.Size = New-Object System.Drawing.Size($wWidth, $wHeight)
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.AutoScroll = $true   # falls Fenster doch zu klein wird, scrollen
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$boldFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

# ---- Login ----
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Location = '10,15'; $lblUser.Size = '90,20'; $lblUser.Text = "Login Name:"; $lblUser.Font = $boldFont
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = '105,13'; $txtUser.Size = '160,22'; $txtUser.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Location = '280,15'; $lblPass.Size = '70,20'; $lblPass.Text = "Passwort:"
$form.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = '355,13'; $txtPass.Size = '160,22'; $txtPass.UseSystemPasswordChar = $true; $txtPass.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtPass)

# ---- iLO-Credentials (optional, nur fuer monitored Server) ----
$lblIloUser = New-Object System.Windows.Forms.Label
$lblIloUser.Location = '535,15'; $lblIloUser.Size = '70,20'; $lblIloUser.Text = "iLO User:"
$form.Controls.Add($lblIloUser)

$txtIloUser = New-Object System.Windows.Forms.TextBox
$txtIloUser.Location = '605,13'; $txtIloUser.Size = '140,22'; $txtIloUser.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtIloUser)

$lblIloPass = New-Object System.Windows.Forms.Label
$lblIloPass.Location = '755,15'; $lblIloPass.Size = '70,20'; $lblIloPass.Text = "iLO Pwd:"
$form.Controls.Add($lblIloPass)

$txtIloPass = New-Object System.Windows.Forms.TextBox
$txtIloPass.Location = '825,13'; $txtIloPass.Size = '140,22'; $txtIloPass.UseSystemPasswordChar = $true; $txtIloPass.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtIloPass)

$tipIlo = New-Object System.Windows.Forms.ToolTip
$tipIlo.AutoPopDelay = 15000
$tipIlo.SetToolTip($txtIloUser, "Nur erforderlich fuer MONITORED Server (OneView hat dort kein SSO-Vertrauen am iLO).`r`nFuer MANAGED Server wird automatisch OneView-SSO genutzt - diese Felder bleiben dann ungenutzt.`r`nIdealerweise ein read-only iLO-Account (per iLO-Federation einheitlich ausgerollt).")
$tipIlo.SetToolTip($txtIloPass, "Nur erforderlich fuer MONITORED Server. Bei MANAGED Servern ignoriert.")

# ---- IP-Datei ----
$lblIP = New-Object System.Windows.Forms.Label
$lblIP.Location = '10,47'; $lblIP.Size = '110,20'; $lblIP.Text = "OneView IP-Datei:"; $lblIP.Font = $boldFont
$form.Controls.Add($lblIP)

$txtIP = New-Object System.Windows.Forms.TextBox
$txtIP.Location = '125,45'; $txtIP.Size = '850,22'; $txtIP.Text = (Join-Path $scriptFolder "Oneview.txt"); $txtIP.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtIP)

$btnBrowseIP = New-Object System.Windows.Forms.Button
$btnBrowseIP.Location = '985,44'; $btnBrowseIP.Size = '75,24'; $btnBrowseIP.Text = "Browse..."
$form.Controls.Add($btnBrowseIP)

# ---- Appliance-Auswahl ----
$lblApp = New-Object System.Windows.Forms.Label
$lblApp.Location = '10,79'; $lblApp.Size = '140,20'; $lblApp.Text = "Appliance(s):"; $lblApp.Font = $boldFont
$form.Controls.Add($lblApp)

$btnSelAll = New-Object System.Windows.Forms.Button
$btnSelAll.Location = '155,76'; $btnSelAll.Size = '60,24'; $btnSelAll.Text = "Alle"
$form.Controls.Add($btnSelAll)

$btnSelNone = New-Object System.Windows.Forms.Button
$btnSelNone.Location = '222,76'; $btnSelNone.Size = '60,24'; $btnSelNone.Text = "Keine"
$form.Controls.Add($btnSelNone)

$chkAppliances = New-Object System.Windows.Forms.CheckedListBox
$chkAppliances.Location = '10,104'; $chkAppliances.Size = '1050,110'; $chkAppliances.CheckOnClick = $true; $chkAppliances.BorderStyle = 'FixedSingle'
$form.Controls.Add($chkAppliances)

# ---- Suche ----
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Location = '10,228'; $lblSearch.Size = '180,20'; $lblSearch.Text = "Servername oder Seriennr.:"; $lblSearch.Font = $boldFont
$form.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = '195,226'; $txtSearch.Size = '250,22'; $txtSearch.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtSearch)

# Tooltip mit Wildcard-Hinweis
$tipSearch = New-Object System.Windows.Forms.ToolTip
$tipSearch.AutoPopDelay = 15000
$tipSearch.SetToolTip($txtSearch, "Sucht in Name, Seriennummer und Modell (Gross-/Kleinschreibung egal).`r`nTeilstrings funktionieren direkt: 'srv01' findet 'srv01-prod-frontend'.`r`nWildcards: *  = beliebig viele Zeichen,  ?  = ein Zeichen.`r`nBeispiele:  srv01*   |   *prod*   |   ?bl0?-*   |   CZ12345*")

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Location = '455,225'; $btnSearch.Size = '110,25'; $btnSearch.Text = "Suchen"
$form.Controls.Add($btnSearch)

$btnExportTxt = New-Object System.Windows.Forms.Button
$btnExportTxt.Location = '575,225'; $btnExportTxt.Size = '110,25'; $btnExportTxt.Text = "TXT-Bericht..."
$btnExportTxt.Enabled = $false
$form.Controls.Add($btnExportTxt)

$btnExportHtml = New-Object System.Windows.Forms.Button
$btnExportHtml.Location = '690,225'; $btnExportHtml.Size = '125,25'; $btnExportHtml.Text = "HTML-Bericht..."
$btnExportHtml.Enabled = $false
$form.Controls.Add($btnExportHtml)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Location = '985,225'; $btnExit.Size = '75,25'; $btnExit.Text = "Exit"
$form.Controls.Add($btnExit)
$btnExit.Add_Click({ $form.Close() })

# ---- Treffer-Liste (falls mehrere Treffer) ----
$lblHits = New-Object System.Windows.Forms.Label
$lblHits.Location = '10,260'; $lblHits.Size = '300,20'; $lblHits.Text = "Treffer:"; $lblHits.Font = $boldFont
$form.Controls.Add($lblHits)

$dgvHits = New-Object System.Windows.Forms.DataGridView
$dgvHits.Location = '10,282'; $dgvHits.Size = '1050,140'
$dgvHits.AllowUserToAddRows = $false; $dgvHits.AllowUserToDeleteRows = $false
$dgvHits.ReadOnly = $true; $dgvHits.SelectionMode = 'FullRowSelect'; $dgvHits.MultiSelect = $false
$dgvHits.AutoSizeColumnsMode = 'Fill'; $dgvHits.RowHeadersVisible = $false
$dgvHits.BackgroundColor = [System.Drawing.SystemColors]::Window
$dgvHits.BorderStyle = 'FixedSingle'
$dgvHits.Columns.Add("appliance", "Appliance") | Out-Null
$dgvHits.Columns.Add("name", "Server") | Out-Null
$dgvHits.Columns.Add("profile", "Profil / Hostname") | Out-Null
$dgvHits.Columns.Add("serial", "Seriennr.") | Out-Null
$dgvHits.Columns.Add("model", "Modell") | Out-Null
$dgvHits.Columns.Add("formFactor", "Formfaktor") | Out-Null
$dgvHits.Columns.Add("location", "Frame / Bay") | Out-Null
$dgvHits.Columns.Add("status", "Status") | Out-Null
$dgvHits.Columns.Add("power", "Power") | Out-Null
$dgvHits.Columns["appliance"].FillWeight  = 12
$dgvHits.Columns["name"].FillWeight       = 14
$dgvHits.Columns["profile"].FillWeight    = 16
$dgvHits.Columns["serial"].FillWeight     = 10
$dgvHits.Columns["model"].FillWeight      = 16
$dgvHits.Columns["formFactor"].FillWeight = 8
$dgvHits.Columns["location"].FillWeight   = 14
$dgvHits.Columns["status"].FillWeight     = 7
$dgvHits.Columns["power"].FillWeight      = 7
$form.Controls.Add($dgvHits)

# Speichert pro Zeile das vollständige Server-Hardware-Objekt + Appliance/Session-Hinweis
$script:hitObjects = @()  # Array von Hashtables: @{Appliance, ApiVersion, ServerHw, EnclosureUri}
$script:isLoadingDetails = $false  # Re-Entrancy-Guard fuer Show-Details
$script:lastDetailRow = -1         # zuletzt geladene Treffer-Zeile (Doppel-Load-Sperre)

# ---- Detail-Bereich (TabControl) ----
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = '10,432'; $tabControl.Size = '1050,380'
$form.Controls.Add($tabControl)

# Tab "Übersicht"
$tabOverview = New-Object System.Windows.Forms.TabPage; $tabOverview.Text = "Übersicht"
$tabControl.TabPages.Add($tabOverview)

$txtOverview = New-Object System.Windows.Forms.TextBox
$txtOverview.Multiline = $true; $txtOverview.ReadOnly = $true; $txtOverview.ScrollBars = 'Vertical'
$txtOverview.Dock = 'Fill'; $txtOverview.BorderStyle = 'FixedSingle'
$txtOverview.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabOverview.Controls.Add($txtOverview)

# Tab "CPU & RAM"
$tabCpu = New-Object System.Windows.Forms.TabPage; $tabCpu.Text = "CPU & RAM"
$tabControl.TabPages.Add($tabCpu)
$txtCpu = New-Object System.Windows.Forms.TextBox
$txtCpu.Multiline = $true; $txtCpu.ReadOnly = $true; $txtCpu.ScrollBars = 'Vertical'
$txtCpu.Dock = 'Fill'; $txtCpu.BorderStyle = 'FixedSingle'
$txtCpu.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabCpu.Controls.Add($txtCpu)

# Tab "Firmware"
$tabFw = New-Object System.Windows.Forms.TabPage; $tabFw.Text = "Firmware-Inventory"
$tabControl.TabPages.Add($tabFw)
$dgvFw = New-Object System.Windows.Forms.DataGridView
$dgvFw.Dock = 'Fill'
$dgvFw.AllowUserToAddRows = $false; $dgvFw.AllowUserToDeleteRows = $false
$dgvFw.ReadOnly = $true; $dgvFw.SelectionMode = 'FullRowSelect'; $dgvFw.RowHeadersVisible = $false
$dgvFw.AutoSizeColumnsMode = 'Fill'; $dgvFw.BorderStyle = 'FixedSingle'
$dgvFw.BackgroundColor = [System.Drawing.SystemColors]::Window
$dgvFw.Columns.Add("componentName", "Komponente") | Out-Null
$dgvFw.Columns.Add("componentLocation", "Position") | Out-Null
$dgvFw.Columns.Add("componentVersion", "Version") | Out-Null
$dgvFw.Columns.Add("componentKey", "Key") | Out-Null
$dgvFw.Columns["componentName"].FillWeight = 35
$dgvFw.Columns["componentLocation"].FillWeight = 20
$dgvFw.Columns["componentVersion"].FillWeight = 20
$dgvFw.Columns["componentKey"].FillWeight = 25
$tabFw.Controls.Add($dgvFw)

# Tab "Netzwerk / Adapter"
$tabNet = New-Object System.Windows.Forms.TabPage; $tabNet.Text = "Adapter / Ports"
$tabControl.TabPages.Add($tabNet)
# SplitContainer: oben Tabelle (Hardware-Ports), unten Detailtext (Adapter, FlexNICs, Profil-Connections)
$splitNet = New-Object System.Windows.Forms.SplitContainer
$splitNet.Dock = 'Fill'
$splitNet.Orientation = 'Horizontal'
$splitNet.SplitterDistance = 180
$splitNet.Panel1MinSize = 100
$splitNet.Panel2MinSize = 100
$tabNet.Controls.Add($splitNet)

$dgvNet = New-Object System.Windows.Forms.DataGridView
$dgvNet.Dock = 'Fill'
$dgvNet.AllowUserToAddRows = $false; $dgvNet.AllowUserToDeleteRows = $false
$dgvNet.ReadOnly = $true; $dgvNet.SelectionMode = 'FullRowSelect'; $dgvNet.RowHeadersVisible = $false
$dgvNet.AutoSizeColumnsMode = 'Fill'; $dgvNet.BorderStyle = 'FixedSingle'
$dgvNet.BackgroundColor = [System.Drawing.SystemColors]::Window
$dgvNet.Columns.Add("slot",     "Slot")           | Out-Null
$dgvNet.Columns.Add("adapter",  "Adapter")        | Out-Null
$dgvNet.Columns.Add("model",    "Modell")         | Out-Null
$dgvNet.Columns.Add("fw",       "Firmware")       | Out-Null
$dgvNet.Columns.Add("port",     "Port")           | Out-Null
$dgvNet.Columns.Add("type",     "Typ")            | Out-Null
$dgvNet.Columns.Add("speedCur", "Speed akt.")     | Out-Null
$dgvNet.Columns.Add("speedMax", "Speed max.")     | Out-Null
$dgvNet.Columns.Add("mac",      "MAC")            | Out-Null
$dgvNet.Columns.Add("wwpn",     "WWPN / WWNN")    | Out-Null
$dgvNet.Columns.Add("status",   "Link/Status")    | Out-Null
$dgvNet.Columns["slot"].FillWeight     = 5
$dgvNet.Columns["adapter"].FillWeight  = 15
$dgvNet.Columns["model"].FillWeight    = 18
$dgvNet.Columns["fw"].FillWeight       = 8
$dgvNet.Columns["port"].FillWeight     = 5
$dgvNet.Columns["type"].FillWeight     = 7
$dgvNet.Columns["speedCur"].FillWeight = 7
$dgvNet.Columns["speedMax"].FillWeight = 7
$dgvNet.Columns["mac"].FillWeight      = 11
$dgvNet.Columns["wwpn"].FillWeight     = 11
$dgvNet.Columns["status"].FillWeight   = 8
$splitNet.Panel1.Controls.Add($dgvNet)

$txtNet = New-Object System.Windows.Forms.TextBox
$txtNet.Multiline = $true; $txtNet.ReadOnly = $true; $txtNet.ScrollBars = 'Both'; $txtNet.WordWrap = $false
$txtNet.Dock = 'Fill'; $txtNet.BorderStyle = 'FixedSingle'
$txtNet.Font = New-Object System.Drawing.Font("Consolas", 9)
$splitNet.Panel2.Controls.Add($txtNet)

# Tab "Server-Profil"
$tabProfile = New-Object System.Windows.Forms.TabPage; $tabProfile.Text = "Server-Profil"
$tabControl.TabPages.Add($tabProfile)
$txtProfile = New-Object System.Windows.Forms.TextBox
$txtProfile.Multiline = $true; $txtProfile.ReadOnly = $true; $txtProfile.ScrollBars = 'Vertical'
$txtProfile.Dock = 'Fill'; $txtProfile.BorderStyle = 'FixedSingle'
$txtProfile.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabProfile.Controls.Add($txtProfile)

# Tab "Storage / Laufwerke"
$tabStorage = New-Object System.Windows.Forms.TabPage; $tabStorage.Text = "Storage"
$tabControl.TabPages.Add($tabStorage)
$txtStorage = New-Object System.Windows.Forms.TextBox
$txtStorage.Multiline = $true; $txtStorage.ReadOnly = $true; $txtStorage.ScrollBars = 'Both'
$txtStorage.WordWrap = $false
$txtStorage.Dock = 'Fill'; $txtStorage.BorderStyle = 'FixedSingle'
$txtStorage.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabStorage.Controls.Add($txtStorage)

# Tab "Power / Thermal"
$tabPower = New-Object System.Windows.Forms.TabPage; $tabPower.Text = "Power / Thermal"
$tabControl.TabPages.Add($tabPower)
$txtPower = New-Object System.Windows.Forms.TextBox
$txtPower.Multiline = $true; $txtPower.ReadOnly = $true; $txtPower.ScrollBars = 'Both'
$txtPower.WordWrap = $false
$txtPower.Dock = 'Fill'; $txtPower.BorderStyle = 'FixedSingle'
$txtPower.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabPower.Controls.Add($txtPower)

# Tab "GPU / Grafik"
$tabGpu = New-Object System.Windows.Forms.TabPage; $tabGpu.Text = "GPU / Grafik"
$tabControl.TabPages.Add($tabGpu)
$txtGpu = New-Object System.Windows.Forms.TextBox
$txtGpu.Multiline = $true; $txtGpu.ReadOnly = $true; $txtGpu.ScrollBars = 'Both'
$txtGpu.WordWrap = $false
$txtGpu.Dock = 'Fill'; $txtGpu.BorderStyle = 'FixedSingle'
$txtGpu.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabGpu.Controls.Add($txtGpu)

# Tab "BIOS" (iLO Redfish via OneView SSO + Profil-BIOS)
$tabBios = New-Object System.Windows.Forms.TabPage; $tabBios.Text = "BIOS"
$tabControl.TabPages.Add($tabBios)
$txtBios = New-Object System.Windows.Forms.TextBox
$txtBios.Multiline = $true; $txtBios.ReadOnly = $true; $txtBios.ScrollBars = 'Both'
$txtBios.WordWrap = $false
$txtBios.Dock = 'Fill'; $txtBios.BorderStyle = 'FixedSingle'
$txtBios.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabBios.Controls.Add($txtBios)

# Tab "Alle Felder" (rekursiver Flach-Dump - hier steht WIRKLICH alles)
$tabAll = New-Object System.Windows.Forms.TabPage; $tabAll.Text = "Alle Felder"
$tabControl.TabPages.Add($tabAll)
$txtAll = New-Object System.Windows.Forms.TextBox
$txtAll.Multiline = $true; $txtAll.ReadOnly = $true; $txtAll.ScrollBars = 'Both'
$txtAll.WordWrap = $false
$txtAll.Dock = 'Fill'; $txtAll.BorderStyle = 'FixedSingle'
$txtAll.Font = New-Object System.Drawing.Font("Consolas", 8)
$tabAll.Controls.Add($txtAll)

# Tab "Raw JSON" (zur Diagnose)
$tabRaw = New-Object System.Windows.Forms.TabPage; $tabRaw.Text = "Raw JSON"
$tabControl.TabPages.Add($tabRaw)
$txtRaw = New-Object System.Windows.Forms.TextBox
$txtRaw.Multiline = $true; $txtRaw.ReadOnly = $true; $txtRaw.ScrollBars = 'Both'
$txtRaw.WordWrap = $false
$txtRaw.Dock = 'Fill'; $txtRaw.BorderStyle = 'FixedSingle'
$txtRaw.Font = New-Object System.Drawing.Font("Consolas", 8)
$tabRaw.Controls.Add($txtRaw)

# ---- StatusStrip ----
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.Dock = 'Bottom'
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Bereit..."
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

# =============================
# Anchors fuer kleine Bildschirme (Notebook): Controls wachsen mit
# =============================
$AnchorTLR  = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$AnchorTR   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$AnchorAll  = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

# Reihe IP-Datei
$txtIP.Anchor          = $AnchorTLR
$btnBrowseIP.Anchor    = $AnchorTR
# Appliance-Liste
$chkAppliances.Anchor  = $AnchorTLR
# Such-/Export-Reihe
$btnExportTxt.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$btnExportHtml.Anchor  = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$btnExit.Anchor        = $AnchorTR
# Treffer-Liste
$dgvHits.Anchor        = $AnchorTLR
# Detail-TabControl - faellt unten/seitlich mit der Form
$tabControl.Anchor     = $AnchorAll

# =============================
# Appliance-Liste laden
# =============================
function Load-Appliances {
    $chkAppliances.Items.Clear()
    if (-not [string]::IsNullOrWhiteSpace($txtIP.Text) -and (Test-Path $txtIP.Text)) {
        @(Get-Content $txtIP.Text | Where-Object { $_.Trim() -ne '' -and -not $_.Trim().StartsWith('#') }) | ForEach-Object {
            $chkAppliances.Items.Add("$($_.Trim())   (OV ?)", $true) | Out-Null
        }
    }
}
$btnSelAll.Add_Click({
    for ($i = 0; $i -lt $chkAppliances.Items.Count; $i++) { $chkAppliances.SetItemChecked($i, $true) }
})
$btnSelNone.Add_Click({
    for ($i = 0; $i -lt $chkAppliances.Items.Count; $i++) { $chkAppliances.SetItemChecked($i, $false) }
})
$btnBrowseIP.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Textdateien (*.txt)|*.txt|Alle (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') { $txtIP.Text = $ofd.FileName; Load-Appliances }
})

# Hilfsfunktion: IP aus CheckedList-Eintrag extrahieren
function Get-IPFromItem { param([string]$t); if ($t -match '^\s*(.+?)\s+\(OV') { $Matches[1] } else { $t.Trim() } }

# =============================
# Hilfsfunktionen Anzeige
# =============================
function Format-Bytes {
    param([Parameter(Mandatory = $false)]$Value)
    if ($null -eq $Value) { return '' }
    try {
        $b = [double]$Value
        if ($b -ge 1TB) { return ("{0:N2} TB" -f ($b / 1TB)) }
        if ($b -ge 1GB) { return ("{0:N2} GB" -f ($b / 1GB)) }
        if ($b -ge 1MB) { return ("{0:N2} MB" -f ($b / 1MB)) }
        return "$b B"
    } catch { return "$Value" }
}

function Get-Prop {
    param($obj, [string[]]$names, $default = '')
    foreach ($n in $names) {
        if ($null -ne $obj -and $obj.PSObject.Properties.Name -contains $n -and $null -ne $obj.$n -and "$($obj.$n)" -ne '') {
            return $obj.$n
        }
    }
    return $default
}

# Liefert "FrameName, Bay X" fuer Blades, "Rack-Server" fuer DL/Standalone
function Resolve-Location {
    param($A, $S, $V, $sh)
    $form = Get-Prop $sh @('formFactor', 'serverHardwareTypeUri') ''
    $bay  = Get-Prop $sh @('position', 'serverBay', 'locationBayNumber') ''
    $locUri = Get-Prop $sh @('locationUri') ''

    if ($locUri) {
        try {
            $enc = OV-Rest -A $A -S $S -V $V -M Get -E $locUri
            $encName = Get-Prop $enc @('name', 'enclosureName') $locUri
            $encSerial = Get-Prop $enc @('serialNumber') ''
            if ($bay) {
                if ($encSerial) { return "$encName (SN $encSerial), Bay $bay" }
                return "$encName, Bay $bay"
            } else {
                return "$encName"
            }
        } catch {
            if ($bay) { return "Enclosure (?), Bay $bay" } else { return "Enclosure (?)" }
        }
    }
    if ("$form" -match 'Blade') {
        if ($bay) { return "Blade, Bay $bay (Frame unbekannt)" } else { return "Blade (Frame/Bay unbekannt)" }
    }
    return "Rack-/Standalone-Server"
}

function Build-Overview {
    param($A, $verLabel, $sh, $location, $profileName)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("=== Server-Übersicht ===")
    [void]$sb.AppendLine(("Appliance       : {0}  (OV {1})" -f $A, $verLabel))
    [void]$sb.AppendLine(("Servername      : {0}" -f (Get-Prop $sh @('name'))))
    [void]$sb.AppendLine(("Hostname        : {0}" -f (Get-Prop $sh @('serverName'))))
    [void]$sb.AppendLine(("Modell          : {0}" -f (Get-Prop $sh @('model','shortModel'))))
    [void]$sb.AppendLine(("Seriennummer    : {0}" -f (Get-Prop $sh @('serialNumber'))))
    [void]$sb.AppendLine(("Server-UUID     : {0}" -f (Get-Prop $sh @('uuid','virtualSerialNumber'))))
    [void]$sb.AppendLine(("Asset Tag       : {0}" -f (Get-Prop $sh @('assetTag'))))
    [void]$sb.AppendLine(("Part-Number     : {0}" -f (Get-Prop $sh @('partNumber'))))
    [void]$sb.AppendLine(("Formfaktor      : {0}" -f (Get-Prop $sh @('formFactor'))))
    [void]$sb.AppendLine(("Position        : {0}" -f $location))
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("=== Status ===")
    [void]$sb.AppendLine(("Power-State     : {0}" -f (Get-Prop $sh @('powerState'))))
    [void]$sb.AppendLine(("Health-Status   : {0}" -f (Get-Prop $sh @('status'))))
    [void]$sb.AppendLine(("State           : {0}" -f (Get-Prop $sh @('state'))))
    [void]$sb.AppendLine(("State-Reason    : {0}" -f (Get-Prop $sh @('stateReason'))))
    [void]$sb.AppendLine(("Refresh-State   : {0}" -f (Get-Prop $sh @('refreshState'))))
    [void]$sb.AppendLine(("UID-Light       : {0}" -f (Get-Prop $sh @('uidState'))))
    [void]$sb.AppendLine(("Intrusion       : {0}" -f (Get-Prop $sh @('intrusionFlag'))))
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("=== iLO / Management ===")
    [void]$sb.AppendLine(("iLO Hostname/IP : {0}" -f (Get-Prop $sh @('mpHostInfo.mpHostName','mpIpAddresses','mpDnsName','mpHostName'))))
    if ($sh.PSObject.Properties.Name -contains 'mpHostInfo' -and $sh.mpHostInfo) {
        $mh = $sh.mpHostInfo
        [void]$sb.AppendLine(("iLO Hostname    : {0}" -f (Get-Prop $mh @('mpHostName'))))
        if ($mh.PSObject.Properties.Name -contains 'mpIpAddresses' -and $mh.mpIpAddresses) {
            foreach ($ip in $mh.mpIpAddresses) {
                [void]$sb.AppendLine(("iLO IP          : {0} ({1})" -f (Get-Prop $ip @('address')), (Get-Prop $ip @('type'))))
            }
        }
    }
    [void]$sb.AppendLine(("iLO-Modell      : {0}" -f (Get-Prop $sh @('mpModel'))))
    [void]$sb.AppendLine(("iLO-Firmware    : {0}" -f (Get-Prop $sh @('mpFirmwareVersion'))))
    [void]$sb.AppendLine(("ROM/BIOS        : {0}" -f (Get-Prop $sh @('romVersion'))))
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("=== Server-Profil ===")
    if ($profileName) {
        [void]$sb.AppendLine(("Zugewiesen      : {0}" -f $profileName))
    } else {
        [void]$sb.AppendLine("Zugewiesen      : (keines)")
    }
    [void]$sb.AppendLine(("Lizenz-Intent   : {0}" -f (Get-Prop $sh @('licensingIntent'))))
    return $sb.ToString()
}

function Build-CpuRam {
    param($A, $S, $V, $sh)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("=== Prozessor (Übersicht) ===")
    [void]$sb.AppendLine(("CPU-Typ         : {0}" -f (Get-Prop $sh @('processorType'))))
    [void]$sb.AppendLine(("CPU-Hersteller  : {0}" -f (Get-Prop $sh @('processorManufacturer','manufacturer'))))
    [void]$sb.AppendLine(("CPU-Anzahl      : {0}" -f (Get-Prop $sh @('processorCount'))))
    [void]$sb.AppendLine(("Cores/CPU       : {0}" -f (Get-Prop $sh @('processorCoreCount'))))
    [void]$sb.AppendLine(("Speed           : {0} MHz" -f (Get-Prop $sh @('processorSpeedMhz'))))
    [void]$sb.AppendLine("")

    # ---- Per-CPU Detail ----
    $cpus = Get-Container $sh @('processors','processorList','cpus','processor')
    if (-not $cpus) {
        $uri = Get-Prop $sh @('uri')
        if ($uri) {
            foreach ($ep in @('/processor','/processors','/cpu')) {
                $sub = Try-Rest -A $A -S $S -V $V -E "$uri$ep"
                if ($sub) {
                    $cpus = Get-Container $sub @('members','data','processors','processorList')
                    if (-not $cpus) { $cpus = $sub }
                    if ($cpus) { break }
                }
            }
        }
    }
    if ($cpus) {
        [void]$sb.AppendLine("=== Prozessoren (Detail) ===")
        $i = 0
        foreach ($c in $cpus) {
            $i++
            [void]$sb.AppendLine(("[CPU{0}] {1}" -f $i, (Get-Prop $c @('model','name','productName','processorModel','description'))))
            [void]$sb.AppendLine(("        Sockel        : {0}" -f (Get-Prop $c @('socket','location','slot','socketDesignation'))))
            [void]$sb.AppendLine(("        Hersteller    : {0}" -f (Get-Prop $c @('manufacturer','vendor'))))
            [void]$sb.AppendLine(("        Familie       : {0}" -f (Get-Prop $c @('family','processorFamily'))))
            [void]$sb.AppendLine(("        Architektur   : {0}" -f (Get-Prop $c @('architecture','instructionSet'))))
            [void]$sb.AppendLine(("        Cores/Threads : {0} / {1}" -f (Get-Prop $c @('totalCores','coreCount','cores')), (Get-Prop $c @('totalThreads','threadCount','threads','logicalProcessors'))))
            [void]$sb.AppendLine(("        Max Speed     : {0} MHz" -f (Get-Prop $c @('maxSpeedMHz','maxSpeedMhz','maxSpeed'))))
            [void]$sb.AppendLine(("        Aktuell       : {0} MHz" -f (Get-Prop $c @('currentSpeedMhz','currentSpeedMHz','speed','currentSpeed'))))
            [void]$sb.AppendLine(("        Cache L1/L2/L3: {0} / {1} / {2}" -f (Get-Prop $c @('l1CacheKiB','cacheL1Kb')), (Get-Prop $c @('l2CacheKiB','cacheL2Kb')), (Get-Prop $c @('l3CacheKiB','cacheL3Kb'))))
            [void]$sb.AppendLine(("        Stepping/Rev. : {0}" -f (Get-Prop $c @('stepping','revision'))))
            [void]$sb.AppendLine(("        Microcode     : {0}" -f (Get-Prop $c @('microcode','microcodeVersion'))))
            [void]$sb.AppendLine(("        Serial / PN   : {0}  /  {1}" -f (Get-Prop $c @('serialNumber','sn')), (Get-Prop $c @('partNumber'))))
            [void]$sb.AppendLine(("        Status/State  : {0} / {1}" -f (Get-Prop $c @('status')), (Get-Prop $c @('state','health'))))
        }
        [void]$sb.AppendLine("")
    } else {
        [void]$sb.AppendLine("(Keine Per-CPU-Details verfügbar — typisch für OV 6.60 / Gen7-9)")
        [void]$sb.AppendLine("")
    }

    # ---- Memory ----
    [void]$sb.AppendLine("=== Arbeitsspeicher (Übersicht) ===")
    $memMB = Get-Prop $sh @('memoryMb')
    if ($memMB) {
        $memGB = [math]::Round([double]$memMB / 1024, 2)
        [void]$sb.AppendLine(("Gesamt-RAM      : {0} MB ({1} GB)" -f $memMB, $memGB))
    } else {
        [void]$sb.AppendLine("Gesamt-RAM      : (unbekannt)")
    }
    [void]$sb.AppendLine(("Memory-Speed    : {0}" -f (Get-Prop $sh @('memorySpeedMhz','memoryOperatingSpeedMhz'))))
    [void]$sb.AppendLine(("Memory-Modus    : {0}" -f (Get-Prop $sh @('memoryOperatingMode','memoryMode'))))
    [void]$sb.AppendLine(("Slots gesamt    : {0}" -f (Get-Prop $sh @('memorySlotCount','totalSystemMemorySlots'))))
    [void]$sb.AppendLine("")

    # ---- Per-DIMM Detail ----
    $dimms = $null
    $memRoot = Get-Prop $sh @('memory') $null
    if ($memRoot) {
        $dimms = Get-Container $memRoot @('deviceList','memoryList','dimms','modules')
    }
    if (-not $dimms) {
        $dimms = Get-Container $sh @('memoryModules','dimms','memoryDevices')
    }
    if (-not $dimms) {
        $uri = Get-Prop $sh @('uri')
        if ($uri) {
            foreach ($ep in @('/memory','/memoryModules','/dimms')) {
                $sub = Try-Rest -A $A -S $S -V $V -E "$uri$ep"
                if ($sub) {
                    $dimms = Get-Container $sub @('members','data','deviceList','memoryList','dimms','modules')
                    if (-not $dimms) { $dimms = $sub }
                    if ($dimms) { break }
                }
            }
        }
    }
    if ($dimms) {
        [void]$sb.AppendLine("=== Memory-Module (Detail) ===")
        [void]$sb.AppendLine(("{0,-18} {1,-8} {2,-8} {3,-10} {4,-10} {5,-14} {6,-14} {7,-16} {8}" -f `
            'Locator','Size','Speed','Type','Tech','Manufacturer','PartNumber','Serial','State'))
        [void]$sb.AppendLine(('-' * 130))
        $populated = 0; $total = 0
        foreach ($d in $dimms) {
            $total++
            $loc   = Get-Prop $d @('deviceLocator','locator','location','slot','dimmLocator','memoryLocation')
            $size  = Get-Prop $d @('capacityMiB','sizeMB','sizeMb','capacityMB','capacity','memoryDeviceSize')
            $speed = Get-Prop $d @('operatingSpeedMhz','currentSpeedMhz','speedMhz','speed','memorySpeedMhz')
            $type  = Get-Prop $d @('memoryDeviceType','dimmType','type','memoryType')
            $tech  = Get-Prop $d @('memoryTechnology','technology','formFactor')
            $vend  = Get-Prop $d @('manufacturer','vendor','memoryManufacturer')
            $pn    = Get-Prop $d @('partNumber','memoryPartNumber')
            $sn    = Get-Prop $d @('serialNumber','memorySerialNumber','sn')
            $st    = Get-Prop $d @('status','state','memoryDeviceStatus')
            $sizeStr = if ($size) {
                $sz = [double]$size
                # Heuristik MiB -> GB
                if ($sz -gt 1024) { ("{0} GB" -f [math]::Round($sz/1024,1)) } else { "$size MB" }
            } else { '' }
            $isPop = $size -and "$size" -ne '0'
            if ($isPop) { $populated++ }
            [void]$sb.AppendLine(("{0,-18} {1,-8} {2,-8} {3,-10} {4,-10} {5,-14} {6,-14} {7,-16} {8}" -f `
                $loc, $sizeStr, $speed, $type, $tech, $vend, $pn, $sn, $st))
        }
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine(("Belegung        : {0} von {1} Slots belegt" -f $populated, $total))
    } else {
        [void]$sb.AppendLine("(Keine Per-DIMM-Details verfügbar — typisch für OV 6.60 / Gen7-9)")
    }

    return $sb.ToString()
}

function Fill-FirmwareGrid {
    param($A, $S, $V, $sh, [System.Windows.Forms.DataGridView]$grid)
    $grid.Rows.Clear()
    $uri = Get-Prop $sh @('uri')
    if (-not $uri) { return }
    try {
        $fw = OV-Rest -A $A -S $S -V $V -M Get -E "$uri/firmware"
    } catch {
        $grid.Rows.Add('(Firmware-Inventory nicht verfügbar)', '', $_.Exception.Message, '') | Out-Null
        return
    }
    $components = $null
    foreach ($p in @('components', 'firmwareComponents')) {
        if ($fw.PSObject.Properties.Name -contains $p -and $fw.$p) { $components = $fw.$p; break }
    }
    if (-not $components) {
        $grid.Rows.Add('(keine Komponenten gemeldet)', '', '', '') | Out-Null
        return
    }
    foreach ($c in $components) {
        $grid.Rows.Add(
            (Get-Prop $c @('componentName','name')),
            (Get-Prop $c @('componentLocation','location')),
            (Get-Prop $c @('componentVersion','version')),
            (Get-Prop $c @('componentKey','key'))
        ) | Out-Null
    }
}

function Fill-NetGrid {
    param(
        $sh,
        [System.Windows.Forms.DataGridView]$grid,
        [System.Windows.Forms.TextBox]$detailBox = $null,
        $A = $null, $S = $null, $V = $null,
        $prof = $null
    )
    $grid.Rows.Clear()
    $sb = New-Object System.Text.StringBuilder

    # --- Profil-Connections + Netzwerk-Mapping (fuer pro-Port Anreicherung) ---
    $connByPort = @{}     # Key: portId (z.B. "Mezz 3:1-a"), Value: connection-Objekt
    $connList   = @()
    $netNameCache = @{}   # uri -> name
    $resolveNetName = {
        param($u)
        if (-not $u) { return '' }
        if ($netNameCache.ContainsKey($u)) { return $netNameCache[$u] }
        $nm = ''
        if ($A -and $S -and $V) {
            try {
                $n = OV-Rest -A $A -S $S -V $V -M Get -E $u
                $nm = [string](Get-Prop $n @('name'))
            } catch { }
        }
        $netNameCache[$u] = $nm
        return $nm
    }

    if ($prof) {
        $connections = $null
        if ($prof.PSObject.Properties.Name -contains 'connectionSettings' -and $prof.connectionSettings -and $prof.connectionSettings.connections) {
            $connections = $prof.connectionSettings.connections
        } elseif ($prof.PSObject.Properties.Name -contains 'connections' -and $prof.connections) {
            $connections = $prof.connections
        }
        if ($connections) {
            foreach ($c in $connections) {
                $portKey = [string](Get-Prop $c @('portId'))
                if ($portKey) { $connByPort[$portKey] = $c }
                $connList += $c
            }
        }
    }

    # --- Adapter / Slots ---
    # Bei Synergy/Blade-Servern liefert /rest/server-hardware/{id} ein portMap-Objekt mit deviceSlots/devices.
    # Bei Rackmount-Servern (HPE ProLiant DL360/DL380/...) ist portMap NICHT vorhanden -
    # die Adapter/Port-Informationen liegen unter /rest/server-hardware/{id}/networkAdapters.
    # Quelle: HPE OneView API Reference (dp00007759en_us) - "GET /rest/server-hardware/{id}/networkAdapters"
    $devices = $null
    $shUri   = [string](Get-Prop $sh @('uri'))
    $usedRackmountEndpoint = $false
    $rawNetAdapters = $null

    if (($sh.PSObject.Properties.Name -contains 'portMap') -and $sh.portMap) {
        foreach ($p in @('deviceSlots', 'devices')) {
            if ($sh.portMap.PSObject.Properties.Name -contains $p -and $sh.portMap.$p) { $devices = $sh.portMap.$p; break }
        }
    }

    if (-not $devices) {
        # Fallback: dedizierter Rackmount-Endpunkt
        if ($shUri -and $A -and $S -and $V) {
            try { $rawNetAdapters = OV-Rest -A $A -S $S -V $V -M Get -E ("$shUri/networkAdapters") } catch { $rawNetAdapters = $null }
        }
        if ($rawNetAdapters) {
            $usedRackmountEndpoint = $true
            # Das Antwortobjekt kann verschiedene Container-Felder haben (Members/members/data/networkAdapters/items)
            # oder direkt ein Array sein. Get-Container loest dies generisch auf.
            $devices = Get-Container $rawNetAdapters
            if (-not $devices) {
                # Manchmal liegt die eigentliche Liste eine Ebene tiefer
                # Redfish nutzt 'Members', OneView u.a. 'networkAdapters'/'adapters'
                foreach ($p in @('Members','members','networkAdapters','adapters','items','data','deviceSlots','devices')) {
                    if ($rawNetAdapters.PSObject.Properties.Name -contains $p -and $rawNetAdapters.$p) {
                        $devices = $rawNetAdapters.$p; break
                    }
                }
            }
            # Falls die Antwort selbst schon ein Einzelobjekt (Adapter) ist
            if (-not $devices -and (
                  $rawNetAdapters.PSObject.Properties.Name -contains 'physicalPorts' -or
                  $rawNetAdapters.PSObject.Properties.Name -contains 'ports' -or
                  $rawNetAdapters.PSObject.Properties.Name -contains 'Controllers' -or
                  $rawNetAdapters.PSObject.Properties.Name -contains 'Ports' -or
                  $rawNetAdapters.PSObject.Properties.Name -contains 'NetworkPorts')) {
                $devices = @($rawNetAdapters)
            }
            # Wenn Members nur Links (@odata.id / href) enthaelt -> jedes Element nachladen
            if ($devices) {
                $expanded = @()
                foreach ($d in @($devices)) {
                    $hasInline = $false
                    foreach ($k in @('physicalPorts','ports','Ports','NetworkPorts','Controllers','deviceName','Name','Model','model')) {
                        if ($d.PSObject.Properties.Name -contains $k -and $d.$k) { $hasInline = $true; break }
                    }
                    if ($hasInline) { $expanded += $d; continue }
                    $link = ''
                    foreach ($k in @('uri','href','@odata.id','target')) {
                        if ($d.PSObject.Properties.Name -contains $k -and $d.$k) { $link = [string]$d.$k; break }
                    }
                    if ($link) {
                        try {
                            $full = OV-Rest -A $A -S $S -V $V -M Get -E $link
                            if ($full) { $expanded += $full } else { $expanded += $d }
                        } catch { $expanded += $d }
                    } else {
                        $expanded += $d
                    }
                }
                $devices = $expanded
            }
        }
    }

    if (-not $devices -or @($devices).Count -eq 0) {
        $msg = @()
        $msg += "(Keine Adapter-/Port-Informationen vom Server gemeldet.)"
        if ($usedRackmountEndpoint) {
            $msg += ""
            $msg += "Endpunkt /rest/server-hardware/{id}/networkAdapters wurde abgefragt,"
            $msg += "lieferte aber keine erkennbare Adapterliste."
            if ($rawNetAdapters) {
                $msg += ""
                $msg += "--- Rohantwort ---"
                try { $msg += ($rawNetAdapters | ConvertTo-Json -Depth 10) } catch { $msg += [string]$rawNetAdapters }
            }
        } else {
            $msg += "(portMap fehlt und /networkAdapters konnte nicht abgerufen werden.)"
        }
        if ($detailBox) { $detailBox.Text = ($msg -join [Environment]::NewLine) }
        return
    }

    # Firmware-Inventory laden (fuer Adapter-Firmware-Anreicherung)
    $fwMap = @{}     # location/component-name -> version
    try {
        $fwUri = (Get-Prop $sh @('uri'))
        if ($fwUri -and $A -and $S -and $V) {
            $fw = OV-Rest -A $A -S $S -V $V -M Get -E ("$fwUri/firmware")
            if ($fw -and $fw.components) {
                foreach ($c in $fw.components) {
                    $cn = [string](Get-Prop $c @('componentName','name'))
                    $cv = [string](Get-Prop $c @('componentVersion','version'))
                    $cl = [string](Get-Prop $c @('componentLocation','location'))
                    if ($cn) { $fwMap[$cn.ToLowerInvariant()] = $cv }
                    if ($cl) { $fwMap[$cl.ToLowerInvariant()] = $cv }
                }
            }
        }
    } catch { }

    [void]$sb.AppendLine("=== Netzwerk-Adapter (Detail) ===")
    [void]$sb.AppendLine(("Anzahl Adapter-Slots: {0}" -f @($devices).Count))
    [void]$sb.AppendLine("")

    foreach ($dev in $devices) {
        $slot     = Get-Prop $dev @('slotNumber','location','deviceNumber','slot','deviceSlot','Id','SlotNumber')
        $devName  = Get-Prop $dev @('deviceName','name','Name','adapterName','productName','Model')
        $model    = Get-Prop $dev @('model','Model','partNumber','PartNumber','productName','SKU')
        $part     = Get-Prop $dev @('partNumber','PartNumber','sparePartNumber','SKU')
        $sn       = Get-Prop $dev @('serialNumber','SerialNumber')
        $mfr      = Get-Prop $dev @('manufacturer','Manufacturer','vendor')
        $devLoc   = Get-Prop $dev @('location','Location','locationName','physicalLocation')
        $devType  = Get-Prop $dev @('deviceType','type','adapterType')
        $devClass = Get-Prop $dev @('class','interconnectClass')
        $devFwDirect = [string](Get-Prop $dev @('firmwareVersion','FirmwareVersion','firmware'))

        # Redfish: Controllers[] enthaelt FirmwarePackageVersion und ControllerCapabilities
        $ctrlList = $null
        if ($dev.PSObject.Properties.Name -contains 'Controllers' -and $dev.Controllers) {
            $ctrlList = @($dev.Controllers)
            if (-not $devFwDirect) {
                foreach ($ctrl in $ctrlList) {
                    $cfw = [string](Get-Prop $ctrl @('FirmwarePackageVersion','firmwarePackageVersion','firmwareVersion'))
                    if ($cfw) { $devFwDirect = $cfw; break }
                }
            }
        }

        $physPorts = $null
        foreach ($p in @('physicalPorts','ports','Ports','NetworkPorts')) {
            if ($dev.PSObject.Properties.Name -contains $p -and $dev.$p) { $physPorts = $dev.$p; break }
        }
        # Redfish-Ports kommen oft als Link-Collection mit Members[].@odata.id
        if ($physPorts) {
            $ppCollection = $null
            if ($physPorts.PSObject -and ($physPorts.PSObject.Properties.Name -contains 'Members' -or $physPorts.PSObject.Properties.Name -contains 'members')) {
                $ppCollection = if ($physPorts.PSObject.Properties.Name -contains 'Members') { $physPorts.Members } else { $physPorts.members }
            }
            if ($ppCollection) {
                $expandedPorts = @()
                foreach ($pm in @($ppCollection)) {
                    $pmInline = $false
                    foreach ($k in @('portNumber','PortNumber','PortId','LinkStatus','mac','MAC','AssociatedNetworkAddresses','CurrentLinkSpeedMbps','currentLinkSpeedMbps')) {
                        if ($pm.PSObject.Properties.Name -contains $k -and $pm.$k) { $pmInline = $true; break }
                    }
                    if ($pmInline) { $expandedPorts += $pm; continue }
                    $plink = ''
                    foreach ($k in @('uri','href','@odata.id')) {
                        if ($pm.PSObject.Properties.Name -contains $k -and $pm.$k) { $plink = [string]$pm.$k; break }
                    }
                    if ($plink) {
                        try { $pfull = OV-Rest -A $A -S $S -V $V -M Get -E $plink } catch { $pfull = $null }
                        if ($pfull) { $expandedPorts += $pfull } else { $expandedPorts += $pm }
                    } else {
                        $expandedPorts += $pm
                    }
                }
                $physPorts = $expandedPorts
            }
        }

        # Leere Bays (kein Adapter-Name UND keine Ports) ueberspringen,
        # damit die Trefferliste nicht von leeren Zeilen vollgemuellt wird.
        if (-not $devName -and (-not $physPorts -or @($physPorts).Count -eq 0)) {
            [void]$sb.AppendLine(("--- Slot {0} ---" -f $slot))
            [void]$sb.AppendLine("  (leerer Slot / kein Adapter installiert)")
            [void]$sb.AppendLine("")
            continue
        }

        # Adapter-Firmware aus Inventory raten (rackmount /networkAdapters liefert oft firmwareVersion direkt)
        $devFw = $devFwDirect
        $needles = @($devName, $model, "Slot $slot", "Mezz $slot", "Embedded LOM") | Where-Object { $_ }
        if (-not $devFw) { foreach ($n in $needles) {
            $k = ([string]$n).ToLowerInvariant()
            if ($fwMap.ContainsKey($k)) { $devFw = $fwMap[$k]; break }
            # Fuzzy: irgendein FW-Eintrag, der den Adapternamen enthaelt
            foreach ($kk in $fwMap.Keys) {
                if ($kk -like "*$k*") { $devFw = $fwMap[$kk]; break }
            }
            if ($devFw) { break }
        } }

        [void]$sb.AppendLine(("--- Slot {0} ---" -f $slot))
        if ($devName) { [void]$sb.AppendLine(("  Adapter      : {0}" -f $devName)) }
        if ($mfr)     { [void]$sb.AppendLine(("  Hersteller   : {0}" -f $mfr)) }
        if ($model)   { [void]$sb.AppendLine(("  Modell       : {0}" -f $model)) }
        if ($part)    { [void]$sb.AppendLine(("  PartNumber   : {0}" -f $part)) }
        if ($sn)      { [void]$sb.AppendLine(("  Seriennummer : {0}" -f $sn)) }
        if ($devLoc -and -not ($devLoc -is [pscustomobject])) { [void]$sb.AppendLine(("  Position     : {0}" -f $devLoc)) }
        if ($devType) { [void]$sb.AppendLine(("  DeviceType   : {0}" -f $devType)) }
        if ($devClass){ [void]$sb.AppendLine(("  Class        : {0}" -f $devClass)) }
        if ($devFw)   { [void]$sb.AppendLine(("  Firmware     : {0}" -f $devFw)) }
        if ($ctrlList) {
            $cidx = 0
            foreach ($ctrl in $ctrlList) {
                $cfw  = [string](Get-Prop $ctrl @('FirmwarePackageVersion','firmwarePackageVersion'))
                $cpc  = ''
                if ($ctrl.PSObject.Properties.Name -contains 'ControllerCapabilities' -and $ctrl.ControllerCapabilities) {
                    $cpc = [string](Get-Prop $ctrl.ControllerCapabilities @('NetworkPortCount'))
                }
                $line = ("  Controller[{0}]: FW {1}" -f $cidx, $cfw)
                if ($cpc) { $line += ("  Ports {0}" -f $cpc) }
                [void]$sb.AppendLine($line)
                $cidx++
            }
        }

        if (-not $physPorts) {
            $grid.Rows.Add($slot, $devName, $model, $devFw, '', '', '', '', '', '', '') | Out-Null
            [void]$sb.AppendLine("  (keine Ports gemeldet)")
            [void]$sb.AppendLine("")
            continue
        }

        $portIndex = 0
        foreach ($pp in $physPorts) {
            $portNo   = Get-Prop $pp @('portNumber','PortNumber','number','physicalPortNumber','port','portId','PortId','Id')
            $type     = Get-Prop $pp @('type','interconnectPortType','portType','portFunction','PortType','LinkNetworkTechnology','ActiveLinkTechnology')
            $mac      = Get-Prop $pp @('mac','MAC','macAddress','physicalMac','permanentMacAddress','PermanentMACAddress')
            $wwpn     = Get-Prop $pp @('wwpn','physicalWwpn','permanentWwpn')
            $wwnn     = Get-Prop $pp @('wwnn','physicalWwnn','permanentWwnn')
            $linkSt   = Get-Prop $pp @('linkStatus','LinkStatus','operationalStatus','status','state','linkState')
            $spdCur   = Get-Prop $pp @('currentLinkSpeedMbps','CurrentLinkSpeedMbps','linkSpeedMbps','currentSpeedMbps','currentSpeed','speedMbps','CurrentSpeedGbps')
            $spdMax   = Get-Prop $pp @('maxLinkSpeedMbps','MaxLinkSpeedMbps','maximumLinkSpeedMbps','maxSpeedMbps','maxSpeed','MaxSpeedGbps')
            $cnxStat  = Get-Prop $pp @('connectionStatus','adminStatus','InterfaceEnabled')
            $portUid  = Get-Prop $pp @('uid','portId','PortId','interconnectPortLabel')

            # Redfish: AssociatedNetworkAddresses -> Array von MACs/WWNs am Port
            if (-not $mac -and ($pp.PSObject.Properties.Name -contains 'AssociatedNetworkAddresses') -and $pp.AssociatedNetworkAddresses) {
                $assoc = @($pp.AssociatedNetworkAddresses)
                $macCand = $assoc | Where-Object { $_ -match '^([0-9a-fA-F]{2}[:\-]){5}[0-9a-fA-F]{2}$' } | Select-Object -First 1
                if ($macCand) { $mac = [string]$macCand }
            }
            # Redfish nutzt Gbps statt Mbps - umrechnen
            if ($spdCur -and $pp.PSObject.Properties.Name -contains 'CurrentSpeedGbps' -and $pp.CurrentSpeedGbps -eq $spdCur) {
                $spdCur = [int]$spdCur * 1000
            }
            if ($spdMax -and $pp.PSObject.Properties.Name -contains 'MaxSpeedGbps' -and $pp.MaxSpeedGbps -eq $spdMax) {
                $spdMax = [int]$spdMax * 1000
            }

            $spdCurStr = if ($spdCur) { "$spdCur Mbps" } else { '' }
            $spdMaxStr = if ($spdMax) { "$spdMax Mbps" } else { '' }
            $wwn       = ''
            if ($wwpn -and $wwnn) { $wwn = "$wwpn / $wwnn" }
            elseif ($wwpn)        { $wwn = $wwpn }
            elseif ($wwnn)        { $wwn = $wwnn }

            # Falls am physicalPort kein MAC/WWN -> aus erstem virtualPort uebernehmen
            $vps = $null
            if ($pp.PSObject.Properties.Name -contains 'virtualPorts' -and $pp.virtualPorts) { $vps = @($pp.virtualPorts) }
            if (-not $mac -and $vps) {
                $vp = $vps | Where-Object { (Get-Prop $_ @('mac')) } | Select-Object -First 1
                if ($vp) { $mac = Get-Prop $vp @('mac') }
            }
            if (-not $wwn -and $vps) {
                $vp = $vps | Where-Object { (Get-Prop $_ @('wwpn','wwnn')) } | Select-Object -First 1
                if ($vp) {
                    $vw = Get-Prop $vp @('wwpn'); $vn = Get-Prop $vp @('wwnn')
                    if ($vw -and $vn) { $wwn = "$vw / $vn" } elseif ($vw) { $wwn = $vw } elseif ($vn) { $wwn = $vn }
                }
            }

            # Bei mehreren Ports am selben Adapter Slot/Adapter/Modell/Firmware
            # nur in der ersten Zeile anzeigen - die folgenden Zeilen bekommen leere
            # Felder, damit der Adapter visuell gruppiert wirkt.
            if ($portIndex -eq 0) {
                $grid.Rows.Add($slot, $devName, $model, $devFw, $portNo, $type, $spdCurStr, $spdMaxStr, $mac, $wwn, $linkSt) | Out-Null
            } else {
                $grid.Rows.Add('', '', '', '', $portNo, $type, $spdCurStr, $spdMaxStr, $mac, $wwn, $linkSt) | Out-Null
            }
            $portIndex++

            [void]$sb.AppendLine(("  Port {0}" -f $portNo))
            if ($type)     { [void]$sb.AppendLine(("    Typ         : {0}" -f $type)) }
            if ($portUid)  { [void]$sb.AppendLine(("    Port-UID    : {0}" -f $portUid)) }
            if ($spdCurStr -or $spdMaxStr) { [void]$sb.AppendLine(("    Speed       : akt. {0,-12}  max. {1}" -f $spdCurStr,$spdMaxStr)) }
            if ($linkSt)   { [void]$sb.AppendLine(("    LinkStatus  : {0}" -f $linkSt)) }
            if ($cnxStat)  { [void]$sb.AppendLine(("    AdminStatus : {0}" -f $cnxStat)) }
            if ($mac)      { [void]$sb.AppendLine(("    MAC (phys)  : {0}" -f $mac)) }
            if ($wwpn)     { [void]$sb.AppendLine(("    WWPN (phys) : {0}" -f $wwpn)) }
            if ($wwnn)     { [void]$sb.AppendLine(("    WWNN (phys) : {0}" -f $wwnn)) }

            # Interconnect-Verbindung (Frame-Switch + Port am Switch)
            $icUri  = Get-Prop $pp @('interconnectUri')
            $icPort = Get-Prop $pp @('interconnectPort','interconnectPortLabel','interconnectPortName')
            if ($icUri -or $icPort) {
                $icName = ''
                if ($icUri -and $A -and $S -and $V) {
                    try {
                        $ic = OV-Rest -A $A -S $S -V $V -M Get -E $icUri
                        $icName = [string](Get-Prop $ic @('name'))
                    } catch { }
                }
                $icDisp = ($icName, $icPort) -join ' Port '
                [void]$sb.AppendLine(("    Interconnect: {0}" -f ($icDisp.Trim(' Port'))))
            }

            # FlexNICs / FlexHBAs (virtualPorts)
            if ($vps -and @($vps).Count -gt 0) {
                [void]$sb.AppendLine(("    FlexNICs ({0}):" -f @($vps).Count))
                foreach ($vp in $vps) {
                    $vpFn   = Get-Prop $vp @('portFunction','portNumber','functionName','function')
                    $vpType = Get-Prop $vp @('portType','functionType','type')
                    $vpMac  = Get-Prop $vp @('mac')
                    $vpWwpn = Get-Prop $vp @('wwpn')
                    $vpWwnn = Get-Prop $vp @('wwnn')
                    $vpSp   = Get-Prop $vp @('currentLinkSpeedMbps','speedMbps','currentSpeedMbps')
                    $vpStat = Get-Prop $vp @('status','linkStatus','state')
                    $line = ("      [{0}] {1,-10} MAC {2,-19} WWPN {3,-25}" -f $vpFn, $vpType, $vpMac, $vpWwpn)
                    if ($vpWwnn) { $line += " WWNN $vpWwnn" }
                    if ($vpSp)   { $line += " Speed ${vpSp}Mbps" }
                    if ($vpStat) { $line += " ($vpStat)" }
                    [void]$sb.AppendLine($line)
                }
            }

            # Profil-Connection-Match per portId
            # OneView nutzt portId-Strings wie "Flb 1:1-a", "Mezz 3:1-a", "Slot 1:1-a"
            if ($connByPort.Count -gt 0) {
                $matched = @()
                foreach ($portKey in $connByPort.Keys) {
                    $pidLow = $portKey.ToLowerInvariant()
                    if ($pidLow -match (":\s*$portNo-")) {
                        # Slot-Praefix grob abgleichen
                        if ($pidLow -match "(?:slot|mezz|flb|lom)\s*$slot\b" -or [string]::IsNullOrEmpty($slot)) {
                            $matched += $connByPort[$portKey]
                        }
                    }
                }
                if ($matched.Count -gt 0) {
                    [void]$sb.AppendLine(("    Profil-Connections ({0}):" -f $matched.Count))
                    foreach ($c in $matched) {
                        $cId   = Get-Prop $c @('id')
                        $cNm   = Get-Prop $c @('name')
                        $cFt   = Get-Prop $c @('functionType')
                        $cPid  = Get-Prop $c @('portId')
                        $cReq  = Get-Prop $c @('requestedMbps')
                        $cAlloc= Get-Prop $c @('allocatedMbps')
                        $cMac  = Get-Prop $c @('mac')
                        $cWwpn = Get-Prop $c @('wwpn')
                        $cBoot = Get-Prop $c @('boot.priority','boot')
                        $cNetU = Get-Prop $c @('networkUri')
                        $cNetT = Get-Prop $c @('networkName','networkType')
                        $netNm = $cNetT
                        if (-not $netNm -and $cNetU) { $netNm = & $resolveNetName $cNetU }
                        [void]$sb.AppendLine(("      ID {0,-3} {1,-22} {2,-8} portId {3,-14} Net '{4}'  req {5} / alloc {6} MAC {7} WWPN {8}" -f `
                            $cId, $cNm, $cFt, $cPid, $netNm, $cReq, $cAlloc, $cMac, $cWwpn))
                        if ($cBoot) { [void]$sb.AppendLine(("        Boot: {0}" -f $cBoot)) }
                    }
                }
            }
            [void]$sb.AppendLine("")
        }
        [void]$sb.AppendLine("")
    }

    # Profile-Connections, die KEINEM physischen Port zuordenbar waren -> separat anzeigen
    if ($connList.Count -gt 0) {
        [void]$sb.AppendLine("=== Server-Profil: Alle Connections ===")
        foreach ($c in $connList) {
            $cId   = Get-Prop $c @('id')
            $cNm   = Get-Prop $c @('name')
            $cFt   = Get-Prop $c @('functionType')
            $cPid  = Get-Prop $c @('portId')
            $cReq  = Get-Prop $c @('requestedMbps')
            $cAlloc= Get-Prop $c @('allocatedMbps')
            $cMac  = Get-Prop $c @('mac')
            $cWwpn = Get-Prop $c @('wwpn')
            $cWwnn = Get-Prop $c @('wwnn')
            $cNetU = Get-Prop $c @('networkUri')
            $cNetT = Get-Prop $c @('networkName')
            $netNm = $cNetT
            if (-not $netNm -and $cNetU) { $netNm = & $resolveNetName $cNetU }
            [void]$sb.AppendLine(("ID {0,-3} {1,-22} {2,-8} portId {3,-14} Net '{4}'  req {5} / alloc {6} Mbps  MAC {7} WWPN {8} WWNN {9}" -f `
                $cId, $cNm, $cFt, $cPid, $netNm, $cReq, $cAlloc, $cMac, $cWwpn, $cWwnn))
        }
    }

    if ($detailBox) { $detailBox.Text = $sb.ToString() }

    # Splitter dynamisch an die Zeilenzahl anpassen, damit oben kein leerer
    # grauer Bereich uebrig bleibt. Header (~24 px) + Zeilen * Hoehe + Puffer.
    try {
        $sc = $grid.Parent
        while ($sc -and -not ($sc -is [System.Windows.Forms.SplitContainer])) { $sc = $sc.Parent }
        if ($sc -and $sc.Orientation -eq [System.Windows.Forms.Orientation]::Horizontal) {
            $rowH    = if ($grid.RowTemplate.Height -gt 0) { $grid.RowTemplate.Height } else { 22 }
            $headerH = if ($grid.ColumnHeadersHeight -gt 0) { $grid.ColumnHeadersHeight } else { 24 }
            $rows    = [Math]::Max($grid.Rows.Count, 1)
            $needed  = $headerH + ($rows * $rowH) + 6
            $maxAllowed = [Math]::Max($sc.Panel1MinSize, $sc.Height - $sc.Panel2MinSize - $sc.SplitterWidth - 1)
            $newDist = [Math]::Max($sc.Panel1MinSize, [Math]::Min($needed, $maxAllowed))
            $sc.SplitterDistance = [int]$newDist
        }
    } catch { }
}

function Build-ProfileText {
    param($A, $S, $V, $sh)
    $sb = New-Object System.Text.StringBuilder
    $profUri = Get-Prop $sh @('serverProfileUri')
    if (-not $profUri) {
        [void]$sb.AppendLine("(Kein Server-Profil zugewiesen)")
        return $sb.ToString()
    }
    try {
        $prof = OV-Rest -A $A -S $S -V $V -M Get -E $profUri
        [void]$sb.AppendLine("=== Server-Profil ===")
        [void]$sb.AppendLine(("Name            : {0}" -f (Get-Prop $prof @('name'))))
        [void]$sb.AppendLine(("Beschreibung    : {0}" -f (Get-Prop $prof @('description'))))
        [void]$sb.AppendLine(("Status          : {0}" -f (Get-Prop $prof @('status'))))
        [void]$sb.AppendLine(("State           : {0}" -f (Get-Prop $prof @('state'))))
        [void]$sb.AppendLine(("Template        : {0}" -f (Get-Prop $prof @('serverProfileTemplateUri'))))
        [void]$sb.AppendLine(("MAC-Type        : {0}" -f (Get-Prop $prof @('macType'))))
        [void]$sb.AppendLine(("WWN-Type        : {0}" -f (Get-Prop $prof @('wwnType'))))
        [void]$sb.AppendLine(("Serial-Type     : {0}" -f (Get-Prop $prof @('serialNumberType'))))
        [void]$sb.AppendLine(("Affinity        : {0}" -f (Get-Prop $prof @('affinity'))))
        [void]$sb.AppendLine(("Hide Unused FM  : {0}" -f (Get-Prop $prof @('hideUnusedFlexNics'))))
        if ($prof.PSObject.Properties.Name -contains 'connectionSettings' -and $prof.connectionSettings -and $prof.connectionSettings.connections) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Connections ---")
            foreach ($c in $prof.connectionSettings.connections) {
                [void]$sb.AppendLine(("ID {0,-3} {1,-20} Port {2,-10} Type {3,-10} MAC {4} WWN {5}" -f `
                    (Get-Prop $c @('id')),
                    (Get-Prop $c @('name')),
                    (Get-Prop $c @('portId')),
                    (Get-Prop $c @('functionType')),
                    (Get-Prop $c @('mac')),
                    (Get-Prop $c @('wwpn'))))
            }
        } elseif ($prof.PSObject.Properties.Name -contains 'connections' -and $prof.connections) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Connections ---")
            foreach ($c in $prof.connections) {
                [void]$sb.AppendLine(("ID {0,-3} {1,-20} Port {2,-10} Type {3,-10} MAC {4} WWN {5}" -f `
                    (Get-Prop $c @('id')),
                    (Get-Prop $c @('name')),
                    (Get-Prop $c @('portId')),
                    (Get-Prop $c @('functionType')),
                    (Get-Prop $c @('mac')),
                    (Get-Prop $c @('wwpn'))))
            }
        }
        return $sb.ToString()
    } catch {
        [void]$sb.AppendLine("Fehler beim Laden des Profils: $($_.Exception.Message)")
        return $sb.ToString()
    }
}

# Liefert den ersten nicht-leeren Property-Wert egal ob als Top-Level-Feld
# oder unter einem typischen Container ('data', 'memberOfCollection') liegt.
function Get-Container {
    param($obj, [string[]]$names)
    if ($null -eq $obj) { return $null }
    foreach ($n in $names) {
        if ($obj.PSObject.Properties.Name -contains $n -and $obj.$n) { return $obj.$n }
    }
    return $null
}

# Versucht, einen Sub-Endpunkt zu laden und gibt $null zurueck, wenn nicht
# verfuegbar (typisch fuer OV 6.60 / Gen7-9). Setzt -ErrorVariable und kann
# optional einen Statuscode-/Fehler-Tracker zurueckliefern.
function Try-Rest {
    param([string]$A, [string]$S, [int]$V, [string]$E, $Status = $null)
    try {
        $r = OV-Rest -A $A -S $S -V $V -M Get -E $E
        if ($Status -is [ref]) { $Status.Value = 'OK' }
        return $r
    } catch {
        $code = $null
        try {
            if ($_.Exception -and $_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $code = [int]$_.Exception.Response.StatusCode
            }
        } catch {}
        if ($Status -is [ref]) {
            if ($code) { $Status.Value = "HTTP $code" }
            else { $Status.Value = "Fehler: $($_.Exception.Message)" }
        }
        return $null
    }
}

function Build-Storage {
    param($A, $S, $V, $sh)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("=== Storage / Lokale Laufwerke ===")
    [void]$sb.AppendLine("(Hinweis: in OV 6.60 / Gen7-9 sind Storage-Details häufig stark eingeschränkt)")
    [void]$sb.AppendLine("")

    $uri = Get-Prop $sh @('uri')
    if ($uri) {
        # Normalisieren: keinen Trailing-Slash, immer mit /rest/ beginnend
        $uri = ([string]$uri).TrimEnd('/')
        [void]$sb.AppendLine(("(Hardware-URI: {0})" -f $uri))
    }

    # Generation aus Model/ShortModel/Generation ableiten
    $modelStr = [string](Get-Prop $sh @('model','Model','shortModel','ShortModel','generation','Generation'))
    $gen = 0
    $isPlus = $false
    # Gen10 Plus / Gen10+ separat erkennen (anderer Endpoint als Gen10)
    if ($modelStr -match 'Gen\s*(\d+)\s*(Plus|\+)') {
        $gen = [int]$Matches[1]
        $isPlus = $true
    } elseif ($modelStr -match 'Gen\s*(\d+)') {
        $gen = [int]$Matches[1]
    }
    $genLabel = if ($gen) { ("Gen{0}{1}" -f $gen, $(if ($isPlus) { '+' } else { '' })) } else { 'unbekannt' }

    # Endpoint-Auswahl gemaess Generation:
    #   Gen <= 9              -> nur /localStorage   (alter HPE-Smart-Storage-Endpoint)
    #   Gen 10 (ohne Plus)    -> nur /localStorage   (alter Endpoint)
    #   Gen 10+ / Gen 10 Plus -> nur /localStorageV2 (neuer Redfish-Endpoint)
    #   Gen >= 11             -> nur /localStorageV2
    #   Gen unbekannt         -> beide
    $useV1 = $true
    $useV2 = $true
    if     ($gen -ge 11)                          { $useV1 = $false; $useV2 = $true }
    elseif ($gen -eq 10 -and $isPlus)             { $useV1 = $false; $useV2 = $true }
    elseif ($gen -eq 10 -and -not $isPlus)        { $useV1 = $true;  $useV2 = $false }
    elseif ($gen -ge 1 -and $gen -le 9)           { $useV1 = $true;  $useV2 = $false }

    $epList = @()
    if ($useV1) { $epList += "GET $uri/localStorage" }
    if ($useV2) { $epList += "GET $uri/localStorageV2" }
    if ($uri) {
        [void]$sb.AppendLine(("(Modell: {0}  |  Generation: {1}  ->  {2})" -f $modelStr, $genLabel, ($epList -join '  +  ')))
        [void]$sb.AppendLine("")
    }

    # 1) Felder direkt aus dem Hardware-Detail (Synergy/Blades liefern oft hier)
    $local   = Get-Container $sh @('localStorage', 'localStorageInfo', 'storage')
    $rawV1   = $null
    $rawV2   = $null

    # 2) Sub-Endpoints je nach Generation aufrufen
    if ($uri) {
        $st1 = "skip ($genLabel)"; $st2 = "skip ($genLabel)"
        if ($useV1) {
            $st1 = 'n/a'
            $rawV1 = Try-Rest -A $A -S $S -V $V -E "$uri/localStorage"   -Status ([ref]$st1)
        }
        if ($useV2) {
            $st2 = 'n/a'
            $rawV2 = Try-Rest -A $A -S $S -V $V -E "$uri/localStorageV2" -Status ([ref]$st2)
        }
        [void]$sb.AppendLine(("(Status localStorage  : {0})" -f $st1))
        [void]$sb.AppendLine(("(Status localStorageV2: {0})" -f $st2))
        [void]$sb.AppendLine("")
    }

    if (-not $local -and -not $rawV1 -and -not $rawV2) {
        [void]$sb.AppendLine("(Keine Storage-Daten verfügbar)")
        return $sb.ToString()
    }

    # Hilfsfunktion: Status/State aus Redfish-Status-Objekt
    function _Format-Status($obj) {
        $h = ''; $st = ''
        if ($obj -and $obj.PSObject.Properties.Name -contains 'Status' -and $obj.Status) {
            $h  = [string](Get-Prop $obj.Status @('Health','HealthRollup'))
            $st = [string](Get-Prop $obj.Status @('State'))
        }
        if (-not $h)  { $h  = [string](Get-Prop $obj @('status','Health')) }
        if (-not $st) { $st = [string](Get-Prop $obj @('state','State')) }
        if ($h -or $st) { return ("{0} / {1}" -f $h, $st) }
        return ''
    }

    # Hilfsfunktion: FirmwareVersion.Current.VersionString -> string (alte Schema)
    function _Get-FwVersion($obj) {
        $v = [string](Get-Prop $obj @('firmwareVersion','FirmwareVersion','firmware'))
        if ($v) { return $v }
        if ($obj -and $obj.PSObject.Properties.Name -contains 'FirmwareVersion' -and $obj.FirmwareVersion) {
            $fw = $obj.FirmwareVersion
            if ($fw.PSObject.Properties.Name -contains 'Current' -and $fw.Current) {
                $v = [string](Get-Prop $fw.Current @('VersionString','versionString'))
                if ($v) { return $v }
            }
        }
        return ''
    }

    # Hilfsfunktion: Location-Objekte (Redfish) zu String
    function _Format-Location($obj) {
        if ($null -eq $obj) { return '' }
        if ($obj -is [string]) { return $obj }
        # Array von Location-Objekten?
        if ($obj -is [System.Collections.IEnumerable] -and -not ($obj -is [string])) {
            $items = @()
            foreach ($it in $obj) { $items += _Format-Location $it }
            return ($items | Where-Object { $_ } | Select-Object -Unique) -join '; '
        }
        $info = [string](Get-Prop $obj @('Info','info'))
        if ($info) { return $info }
        if ($obj.PSObject.Properties.Name -contains 'PartLocation' -and $obj.PartLocation) {
            $lt = [string](Get-Prop $obj.PartLocation @('LocationType'))
            $lv = [string](Get-Prop $obj.PartLocation @('LocationOrdinalValue'))
            $sl = [string](Get-Prop $obj.PartLocation @('ServiceLabel'))
            $parts = @()
            if ($sl) { $parts += $sl }
            elseif ($lt -or $lv) { $parts += ("{0} {1}" -f $lt, $lv).Trim() }
            return ($parts -join ' ')
        }
        return ''
    }

    # === Variante A: altes Schema /localStorage (Single Controller) ===
    $v1Controllers = @()
    if ($rawV1) {
        # Antwort ist meist EIN Controller-Objekt; manchmal eine Liste
        if ($rawV1.PSObject.Properties.Name -contains 'Model' -or
            $rawV1.PSObject.Properties.Name -contains 'AdapterType' -or
            $rawV1.PSObject.Properties.Name -contains 'LogicalDrives' -or
            $rawV1.PSObject.Properties.Name -contains 'PhysicalDrives') {
            $v1Controllers = @($rawV1)
        } else {
            $c1 = Get-Container $rawV1 @('Members','members','controllers','data','items')
            if ($c1) { $v1Controllers = @($c1) }
        }
    }
    if (-not $v1Controllers -and $local) {
        $c1 = Get-Container $local @('controllers','storageControllers','data')
        if ($c1) { $v1Controllers = @($c1) }
    }

    if ($v1Controllers -and @($v1Controllers).Count -gt 0) {
        [void]$sb.AppendLine("--- Controller (localStorage) ---")
        $i = 0
        foreach ($c in $v1Controllers) {
            $i++
            $name = Get-Prop $c @('Name','name','model','Model','controllerName','adapterType','AdapterType')
            [void]$sb.AppendLine(("[{0}] {1}" -f $i, $name))
            $loc = [string](Get-Prop $c @('Location','location','slotNumber','controllerLocation'))
            if ($loc) { [void]$sb.AppendLine(("    Slot/Location : {0}" -f $loc)) }
            $mdl = [string](Get-Prop $c @('Model','model'))
            if ($mdl -and $mdl -ne $name) { [void]$sb.AppendLine(("    Modell        : {0}" -f $mdl)) }
            $sn = [string](Get-Prop $c @('SerialNumber','serialNumber','sn'))
            if ($sn) { [void]$sb.AppendLine(("    Serial        : {0}" -f $sn)) }
            $fw = _Get-FwVersion $c
            if ($fw) { [void]$sb.AppendLine(("    Firmware      : {0}" -f $fw)) }
            $mode = [string](Get-Prop $c @('CurrentOperatingMode','currentOperatingMode','mode'))
            if ($mode) { [void]$sb.AppendLine(("    Mode          : {0}" -f $mode)) }
            $cache = [string](Get-Prop $c @('CacheMemorySizeMiB','cacheMemorySizeMB','cacheMemorySize'))
            if ($cache) { [void]$sb.AppendLine(("    Cache (MiB)   : {0}" -f $cache)) }
            $cacheSn = [string](Get-Prop $c @('CacheModuleSerialNumber'))
            if ($cacheSn) { [void]$sb.AppendLine(("    Cache-Modul SN: {0}" -f $cacheSn)) }
            $bps = [string](Get-Prop $c @('BackupPowerSourceStatus'))
            if ($bps) { [void]$sb.AppendLine(("    Backup Power  : {0}" -f $bps)) }
            $dwc = [string](Get-Prop $c @('DriveWriteCache'))
            if ($dwc) { [void]$sb.AppendLine(("    DriveWriteCache: {0}" -f $dwc)) }
            $enc = [string](Get-Prop $c @('EncryptionEnabled'))
            if ($enc -ne '') { [void]$sb.AppendLine(("    Encryption    : {0}" -f $enc)) }
            $intP = [string](Get-Prop $c @('InternalPortCount'))
            $extP = [string](Get-Prop $c @('ExternalPortCount'))
            if ($intP -or $extP) { [void]$sb.AppendLine(("    Ports int/ext : {0} / {1}" -f $intP, $extP)) }
            $stStr = _Format-Status $c
            if ($stStr) { [void]$sb.AppendLine(("    Status/State  : {0}" -f $stStr)) }

            $drives = Get-Container $c @('PhysicalDrives','physicalDrives','drives','disks')
            if ($drives) {
                [void]$sb.AppendLine("    -- Physische Laufwerke --")
                foreach ($d in $drives) {
                    $loc2  = [string](Get-Prop $d @('Location','location','driveLocation','bay','slot'))
                    $mdl2  = [string](Get-Prop $d @('Model','model','driveModel','name'))
                    $sn2   = [string](Get-Prop $d @('SerialNumber','serialNumber','sn'))
                    $capGB = [string](Get-Prop $d @('CapacityGB','capacityGB','capacity','sizeGB'))
                    $fw2   = _Get-FwVersion $d
                    $media = [string](Get-Prop $d @('MediaType','mediaType'))
                    $intf  = [string](Get-Prop $d @('InterfaceType','interfaceType'))
                    $rpm   = [string](Get-Prop $d @('RotationalSpeedRpm','rotationalSpeedRpm'))
                    $use   = [string](Get-Prop $d @('DiskDriveUse'))
                    $life  = [string](Get-Prop $d @('SSDEnduranceUtilizationPercentage'))
                    $poh   = [string](Get-Prop $d @('PowerOnHours'))
                    $stStr2= _Format-Status $d
                    [void]$sb.AppendLine(("       {0,-12} {1,-22} SN {2,-18} {3,-6} GB  FW {4,-10} {5}" -f `
                        $loc2, $mdl2, $sn2, $capGB, $fw2, $stStr2))
                    $extra = @()
                    if ($media) { $extra += ("Type=$media") }
                    if ($intf)  { $extra += ("IF=$intf") }
                    if ($rpm)   { $extra += ("RPM=$rpm") }
                    if ($use)   { $extra += ("Use=$use") }
                    if ($life)  { $extra += ("Endurance=$life%") }
                    if ($poh)   { $extra += ("PoH=$poh") }
                    if ($extra.Count) { [void]$sb.AppendLine(("           " + ($extra -join '  '))) }
                }
            }

            $logical = Get-Container $c @('LogicalDrives','logicalDrives','volumes','arrays')
            if ($logical) {
                [void]$sb.AppendLine("    -- Logische Laufwerke / Volumes --")
                foreach ($l in $logical) {
                    $lname = [string](Get-Prop $l @('LogicalDriveName','name','volumeName','logicalDriveName'))
                    $lraid = [string](Get-Prop $l @('Raid','raid','raidLevel'))
                    $lcap  = [string](Get-Prop $l @('CapacityMiB','capacityMiB','capacityGB','capacity','sizeGB'))
                    $lcapU = if ($l.PSObject.Properties.Name -contains 'CapacityMiB' -and $l.CapacityMiB) { 'MiB' } else { 'GB' }
                    $ltype = [string](Get-Prop $l @('LogicalDriveType','MediaType'))
                    $lstr  = _Format-Status $l
                    [void]$sb.AppendLine(("       {0,-20} RAID {1,-6} {2,-10} {3,-4}  {4,-10} {5}" -f `
                        $lname, $lraid, $lcap, $lcapU, $ltype, $lstr))
                    $reasons = Get-Container $l @('LogicalDriveStatusReasons')
                    if ($reasons) {
                        [void]$sb.AppendLine(("           Reasons     : {0}" -f (($reasons | ForEach-Object { [string]$_ }) -join ', ')))
                    }
                }
            }

            $enclosures = Get-Container $c @('StorageEnclosures','storageEnclosures')
            if ($enclosures) {
                [void]$sb.AppendLine("    -- Storage-Enclosures --")
                foreach ($e in $enclosures) {
                    $eid   = [string](Get-Prop $e @('Id','id'))
                    $emdl  = [string](Get-Prop $e @('Model','model'))
                    $esn   = [string](Get-Prop $e @('SerialNumber','serialNumber'))
                    $eloc  = [string](Get-Prop $e @('Location','location'))
                    $ebays = [string](Get-Prop $e @('DriveBayCount'))
                    $efw   = _Get-FwVersion $e
                    $estr  = _Format-Status $e
                    [void]$sb.AppendLine(("       Id {0,-6} {1,-22} SN {2,-18} Loc {3,-10} Bays {4,-3} FW {5,-10} {6}" -f `
                        $eid, $emdl, $esn, $eloc, $ebays, $efw, $estr))
                }
            }
            [void]$sb.AppendLine("")
        }
    }

    # === Variante B: neues Redfish-Schema /localStorageV2 (Gen10+/Gen11) ===
    # WICHTIG: Es wird ausschliesslich der OneView-Endpunkt
    #   GET /rest/server-hardware/{id}/localStorageV2
    # verwendet. Eingebettete @odata.id-Referenzen werden NUR dann verfolgt,
    # wenn sie wieder ein OneView-relativer Pfad sind (beginnt mit "/rest/").
    # Redfish-iLO-URIs (z.B. /redfish/...) oder absolute https-URLs werden
    # bewusst NICHT aufgerufen, da diese ueber OneView nicht erreichbar sind.
    function _Expand-Refs($val) {
        if ($null -eq $val) { return @() }
        # Collection-Wrapper { Members: [...] }
        if ($val.PSObject -and $val.PSObject.Properties.Name -contains 'Members' -and $val.Members) {
            $out = @()
            foreach ($m in $val.Members) { $out += (_Expand-Refs $m) }
            return $out
        }
        # Array / Liste
        if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string]) -and -not ($val -is [pscustomobject])) {
            $out = @()
            foreach ($m in $val) { $out += (_Expand-Refs $m) }
            return $out
        }
        # Einzel-Objekt: pruefen, ob es nur eine Referenz ist
        $linkUri = $null
        foreach ($lp in @('@odata.id','uri','href')) {
            if ($val.PSObject -and $val.PSObject.Properties.Name -contains $lp -and $val.$lp) {
                $linkUri = [string]$val.$lp; break
            }
        }
        $realPropCount = 0
        if ($val.PSObject) {
            foreach ($p in $val.PSObject.Properties.Name) {
                if ($p -notmatch '^@' -and $p -notin @('uri','href','Id','id','Name','name')) { $realPropCount++ }
            }
        }
        if ($linkUri -and $realPropCount -lt 2) {
            # Erlaubt sind appliance-relative Pfade: /rest/... (OneView) ODER
            # /redfish/... (von OneView gespiegelter Redfish-Endpunkt auf
            # derselben Appliance). Absolute https-URLs werden nicht verfolgt.
            if ($linkUri -match '^/(rest|redfish)/') {
                $sub = Try-Rest -A $A -S $S -V $V -E $linkUri
                if ($sub) { return (_Expand-Refs $sub) }
            }
            # Wenn nicht aufloesbar: Objekt selbst zurueckgeben (enthaelt
            # mindestens Id/Name als Hinweis)
            return @($val)
        }
        return @($val)
    }

    # Hilfsfunktion: liest aus einem Storage-Objekt ein Sub-Feld unter mehreren
    # möglichen Namen, expandiert Members/Links.
    function _Get-V2List($obj, [string[]]$names) {
        foreach ($n in $names) {
            if ($obj.PSObject -and $obj.PSObject.Properties.Name -contains $n -and $obj.$n) {
                return (_Expand-Refs $obj.$n)
            }
        }
        return @()
    }

    if ($rawV2) {
        [void]$sb.AppendLine("--- Storage System (localStorageV2) ---")

        # rawV2 kann selbst eine Collection ODER ein einzelnes Storage-System sein.
        $systems = @()
        foreach ($wf in @('Members','members','data','items','storageSystems','StorageSystems','Systems','systems')) {
            if ($rawV2.PSObject.Properties.Name -contains $wf -and $rawV2.$wf) {
                $systems = _Expand-Refs $rawV2.$wf
                break
            }
        }
        if (-not $systems -or @($systems).Count -eq 0) { $systems = @($rawV2) }

        [void]$sb.AppendLine("")

        $sysIdx = 0
        foreach ($sys in $systems) {
            $sysIdx++
            $sysName = [string](Get-Prop $sys @('Name','name'))
            $sysId   = [string](Get-Prop $sys @('Id','id'))
            $sysLoc  = [string](Get-Prop $sys @('Location','location'))
            $sysSt   = _Format-Status $sys
            [void]$sb.AppendLine(("[System {0}] Name/Id : {1} / {2}" -f $sysIdx, $sysName, $sysId))
            if ($sysLoc) { [void]$sb.AppendLine(("  Location       : {0}" -f $sysLoc)) }
            if ($sysSt)  { [void]$sb.AppendLine(("  Status/State   : {0}" -f $sysSt)) }
            [void]$sb.AppendLine("")

            # Controllers (bevorzugt) sonst StorageControllers (deprecated)
            $ctrls = _Get-V2List $sys @('Controllers','controllers')
            if (-not $ctrls -or @($ctrls).Count -eq 0) {
                $ctrls = _Get-V2List $sys @('StorageControllers','storageControllers')
            }
        if ($ctrls -and @($ctrls).Count -gt 0) {
            [void]$sb.AppendLine(("  -- Controller ({0}) --" -f @($ctrls).Count))
            $ci = 0
            foreach ($c in $ctrls) {
                $ci++
                $cname = [string](Get-Prop $c @('Name','name','Model','model'))
                $cmfr  = [string](Get-Prop $c @('Manufacturer'))
                $cmdl  = [string](Get-Prop $c @('Model'))
                $cpn   = [string](Get-Prop $c @('PartNumber'))
                $csn   = [string](Get-Prop $c @('SerialNumber'))
                $csku  = [string](Get-Prop $c @('SKU'))
                $cfw   = [string](Get-Prop $c @('FirmwareVersion'))
                $cspd  = [string](Get-Prop $c @('SpeedGbps'))
                $cstr  = _Format-Status $c
                [void]$sb.AppendLine(("  [{0}] {1}" -f $ci, $cname))
                if ($cmfr) { [void]$sb.AppendLine(("      Hersteller   : {0}" -f $cmfr)) }
                if ($cmdl -and $cmdl -ne $cname) { [void]$sb.AppendLine(("      Modell       : {0}" -f $cmdl)) }
                if ($cpn)  { [void]$sb.AppendLine(("      PartNumber   : {0}" -f $cpn)) }
                if ($csn)  { [void]$sb.AppendLine(("      Serial       : {0}" -f $csn)) }
                if ($csku) { [void]$sb.AppendLine(("      SKU          : {0}" -f $csku)) }
                if ($cfw)  { [void]$sb.AppendLine(("      Firmware     : {0}" -f $cfw)) }
                if ($cspd) { [void]$sb.AppendLine(("      SpeedGbps    : {0}" -f $cspd)) }
                $cloc = _Format-Location (Get-Prop $c @('Location'))
                if ($cloc) { [void]$sb.AppendLine(("      Location     : {0}" -f $cloc)) }
                if ($c.PSObject.Properties.Name -contains 'PCIeInterface' -and $c.PCIeInterface) {
                    $pe = $c.PCIeInterface
                    [void]$sb.AppendLine(("      PCIe         : {0} (max {1}) Lanes {2}/{3}" -f `
                        (Get-Prop $pe @('PCIeType')), (Get-Prop $pe @('MaxPCIeType')),
                        (Get-Prop $pe @('LanesInUse')), (Get-Prop $pe @('MaxLanes'))))
                }
                if ($c.PSObject.Properties.Name -contains 'CacheSummary' -and $c.CacheSummary) {
                    $cs = $c.CacheSummary
                    $tot = [string](Get-Prop $cs @('TotalCacheSizeMiB'))
                    $per = [string](Get-Prop $cs @('PersistentCacheSizeMiB'))
                    if ($tot -or $per) { [void]$sb.AppendLine(("      Cache (MiB)  : total {0}, persistent {1}" -f $tot, $per)) }
                }
                $raidTypes = Get-Container $c @('SupportedRAIDTypes')
                if ($raidTypes) { [void]$sb.AppendLine(("      RAID Support : {0}" -f (($raidTypes | ForEach-Object { [string]$_ }) -join ', '))) }
                if ($cstr) { [void]$sb.AppendLine(("      Status/State : {0}" -f $cstr)) }

                $cports = Get-Container $c @('Ports','ports')
                if ($cports) {
                    foreach ($cp in $cports) {
                        $cpid   = [string](Get-Prop $cp @('PortId','Id','Name'))
                        $cpType = [string](Get-Prop $cp @('PortType'))
                        $cpProt = [string](Get-Prop $cp @('PortProtocol'))
                        $cpCur  = [string](Get-Prop $cp @('CurrentSpeedGbps'))
                        $cpMax  = [string](Get-Prop $cp @('MaxSpeedGbps'))
                        $cpW    = [string](Get-Prop $cp @('Width'))
                        $cpStr  = _Format-Status $cp
                        [void]$sb.AppendLine(("        Port {0,-6} {1,-12} {2,-6} {3}/{4} Gbps  W{5}  {6}" -f `
                            $cpid, $cpType, $cpProt, $cpCur, $cpMax, $cpW, $cpStr))
                    }
                }
            }
            [void]$sb.AppendLine("")
        }

        # Drives (Top-Level oder per Storage-System; Refs werden aufgelöst)
        $drives2 = _Get-V2List $sys @('Drives','drives')
        if ($drives2 -and @($drives2).Count -gt 0) {
            [void]$sb.AppendLine(("  -- Physische Laufwerke ({0}) --" -f @($drives2).Count))
            foreach ($d in $drives2) {
                $dloc   = _Format-Location (Get-Prop $d @('Location'))
                if (-not $dloc) { $dloc = _Format-Location (Get-Prop $d @('PhysicalLocation')) }
                $dname  = [string](Get-Prop $d @('Name','name'))
                $dmdl   = [string](Get-Prop $d @('Model','model'))
                $dmfr   = [string](Get-Prop $d @('Manufacturer'))
                $dsn    = [string](Get-Prop $d @('SerialNumber'))
                $dpn    = [string](Get-Prop $d @('PartNumber'))
                $drev   = [string](Get-Prop $d @('Revision'))
                $dmedia = [string](Get-Prop $d @('MediaType'))
                $dproto = [string](Get-Prop $d @('Protocol'))
                $dcapB  = Get-Prop $d @('CapacityBytes')
                $dcap   = ''
                if ($dcapB) { try { $dcap = ("{0:N0} GB" -f ([double]$dcapB / 1GB)) } catch { $dcap = [string]$dcapB } }
                $dcurS  = [string](Get-Prop $d @('NegotiatedSpeedGbs'))
                $dmaxS  = [string](Get-Prop $d @('CapableSpeedGbs'))
                $drpm   = [string](Get-Prop $d @('RotationSpeedRPM'))
                $dlife  = [string](Get-Prop $d @('PredictedMediaLifeLeftPercent'))
                $dpred  = [string](Get-Prop $d @('FailurePredicted'))
                $dled   = [string](Get-Prop $d @('IndicatorLED'))
                $dstr   = _Format-Status $d
                [void]$sb.AppendLine(("    {0,-14} {1,-22} SN {2,-18} {3,-10}  FW {4,-10} {5}" -f `
                    $dloc, ($dmdl ? $dmdl : $dname), $dsn, $dcap, $drev, $dstr))
                $extra = @()
                if ($dmfr)   { $extra += ("Mfr=$dmfr") }
                if ($dpn)    { $extra += ("PN=$dpn") }
                if ($dmedia) { $extra += ("Type=$dmedia") }
                if ($dproto) { $extra += ("Proto=$dproto") }
                if ($dcurS)  { $extra += ("CurSpd=$dcurS Gbps") }
                if ($dmaxS)  { $extra += ("MaxSpd=$dmaxS Gbps") }
                if ($drpm)   { $extra += ("RPM=$drpm") }
                if ($dlife)  { $extra += ("Life=$dlife%") }
                if ($dpred -eq 'True') { $extra += "FailurePredicted!" }
                if ($dled -and $dled -ne 'Off') { $extra += ("LED=$dled") }
                if ($extra.Count) { [void]$sb.AppendLine(("        " + ($extra -join '  '))) }
            }
            [void]$sb.AppendLine("")
        }

        # Volumes
        $vols = _Get-V2List $sys @('Volumes','volumes')
        if ($vols -and @($vols).Count -gt 0) {
            [void]$sb.AppendLine(("  -- Volumes ({0}) --" -f @($vols).Count))
            foreach ($vv in $vols) {
                $vname = [string](Get-Prop $vv @('DisplayName','Name'))
                $vid   = [string](Get-Prop $vv @('Id'))
                $vraid = [string](Get-Prop $vv @('RAIDType'))
                $vtype = [string](Get-Prop $vv @('VolumeType'))
                $vuse  = [string](Get-Prop $vv @('VolumeUsage'))
                $vcapB = Get-Prop $vv @('CapacityBytes')
                $vcap  = ''
                if ($vcapB) { try { $vcap = ("{0:N0} GB" -f ([double]$vcapB / 1GB)) } catch { $vcap = [string]$vcapB } }
                $venc  = [string](Get-Prop $vv @('Encrypted'))
                $vlun  = [string](Get-Prop $vv @('LogicalUnitNumber'))
                $vstr  = _Format-Status $vv
                [void]$sb.AppendLine(("    {0,-22} Id {1,-6} {2,-8} {3,-10} {4,-10} Enc={5,-5} LUN={6,-4} {7}" -f `
                    $vname, $vid, $vraid, $vtype, $vcap, $venc, $vlun, $vstr))
                if ($vuse) { [void]$sb.AppendLine(("        Usage: {0}" -f $vuse)) }
            }
            [void]$sb.AppendLine("")
        }

        # Enclosures
        $encl2 = _Get-V2List $sys @('Enclosures','enclosures')
        if ($encl2 -and @($encl2).Count -gt 0) {
            [void]$sb.AppendLine(("  -- Enclosures ({0}) --" -f @($encl2).Count))
            foreach ($e in $encl2) {
                $eid  = [string](Get-Prop $e @('Id'))
                $en   = [string](Get-Prop $e @('Name'))
                $emfr = [string](Get-Prop $e @('Manufacturer'))
                $emdl = [string](Get-Prop $e @('Model'))
                $epn  = [string](Get-Prop $e @('PartNumber'))
                $esn  = [string](Get-Prop $e @('SerialNumber'))
                $esku = [string](Get-Prop $e @('SKU'))
                $etyp = [string](Get-Prop $e @('ChassisType'))
                $epwr = [string](Get-Prop $e @('PowerState'))
                $estr = _Format-Status $e
                [void]$sb.AppendLine(("    Id {0,-6} {1,-22} {2,-14} SN {3,-18} PN {4,-12} Pwr {5,-8} {6}" -f `
                    $eid, ($emdl ? $emdl : $en), $etyp, $esn, $epn, $epwr, $estr))
                if ($emfr -or $esku) { [void]$sb.AppendLine(("        Mfr=$emfr  SKU=$esku")) }
            }
            [void]$sb.AppendLine("")
        }

            # Diagnose: wenn aus diesem System nichts gefunden wurde, Properties zeigen
            if ((-not $ctrls -or @($ctrls).Count -eq 0) -and `
                (-not $drives2 -or @($drives2).Count -eq 0) -and `
                (-not $vols -or @($vols).Count -eq 0) -and `
                (-not $encl2 -or @($encl2).Count -eq 0)) {
                [void]$sb.AppendLine("  (Keine Sub-Ressourcen im Storage-System gefunden)")
                if ($sys.PSObject -and $sys.PSObject.Properties) {
                    $propList = @($sys.PSObject.Properties.Name) -join ', '
                    [void]$sb.AppendLine(("  Vorhandene Felder: {0}" -f $propList))
                }
                # Roh-JSON-Auszug zur Diagnose
                try {
                    $rawJson = $sys | ConvertTo-Json -Depth 4 -Compress
                    if ($rawJson.Length -gt 2000) { $rawJson = $rawJson.Substring(0,2000) + ' ...[truncated]' }
                    [void]$sb.AppendLine("  Roh-JSON (gekürzt):")
                    [void]$sb.AppendLine(("    {0}" -f $rawJson))
                } catch { }
            }
        } # end foreach $sys
    }

    # Synergy/Blade: sasLogicalJBODs aus dem Hardware-Objekt
    if ($local) {
        $sas = Get-Container $local @('sasLogicalJBODs','enclosures')
        if ($sas) {
            [void]$sb.AppendLine("--- Externe / SAS / JBOD ---")
            foreach ($e in $sas) {
                [void]$sb.AppendLine(("  {0}  {1}" -f (Get-Prop $e @('name','model')), (Get-Prop $e @('serialNumber'))))
            }
        }
    }

    return $sb.ToString()
}

function Build-Power {
    param($A, $S, $V, $sh)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("=== Power & Thermal ===")
    [void]$sb.AppendLine("(Hinweis: in OV 6.60 / Gen7-9 sind viele Power-Felder nicht verfügbar)")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine(("Power-State        : {0}" -f (Get-Prop $sh @('powerState'))))
    [void]$sb.AppendLine(("Power-Lock         : {0}" -f (Get-Prop $sh @('powerLock'))))
    [void]$sb.AppendLine(("Power-Capacity (W) : {0}" -f (Get-Prop $sh @('powerCapacity','maxPowerCapacity'))))
    [void]$sb.AppendLine(("Power-Allocated    : {0}" -f (Get-Prop $sh @('powerAllocatedWatts','allocatedPowerWatts'))))
    [void]$sb.AppendLine(("Hot-Plug Drives    : {0}" -f (Get-Prop $sh @('hotPlugDrivesAllowed'))))
    [void]$sb.AppendLine("")

    # Power Supplies
    # Quelle (Rackmount/Redfish): GET /rest/server-hardware/{id}/powerSupplies
    $ps = Get-Container $sh @('powerSupplies','powerSupply')
    $rawPs = $null
    if (-not $ps) {
        $uri = Get-Prop $sh @('uri')
        if ($uri) {
            $rawPs = Try-Rest -A $A -S $S -V $V -E "$uri/powerSupplies"
            if ($rawPs) {
                # Redfish-Hülle ist Members[]; ältere OneView-APIs nutzen data/powerSupplies/items
                $ps = Get-Container $rawPs @('Members','members','data','powerSupplies','items')
                if (-not $ps) { $ps = $rawPs }
            }
        }
    }
    # Wenn Members nur Links enthält, jedes Element nachladen
    if ($ps) {
        $expandedPs = @()
        foreach ($u in @($ps)) {
            $inline = $false
            foreach ($k in @('Model','model','Name','name','PartNumber','partNumber','SerialNumber','serialNumber','PowerCapacityWatts','Oem')) {
                if ($u.PSObject.Properties.Name -contains $k -and $u.$k) { $inline = $true; break }
            }
            if ($inline) { $expandedPs += $u; continue }
            $link = ''
            foreach ($k in @('uri','href','@odata.id')) {
                if ($u.PSObject.Properties.Name -contains $k -and $u.$k) { $link = [string]$u.$k; break }
            }
            if ($link) {
                $full = Try-Rest -A $A -S $S -V $V -E $link
                if ($full) { $expandedPs += $full } else { $expandedPs += $u }
            } else {
                $expandedPs += $u
            }
        }
        $ps = $expandedPs
    }
    if ($ps -and @($ps).Count -gt 0) {
        [void]$sb.AppendLine("--- Netzteile ---")
        $i = 0
        foreach ($u in $ps) {
            $i++
            $hpe = $null
            if ($u.PSObject.Properties.Name -contains 'Oem' -and $u.Oem -and $u.Oem.PSObject.Properties.Name -contains 'Hpe') {
                $hpe = $u.Oem.Hpe
            }
            $bay = Get-Prop $u @('bayNumber','slot','position','MemberId')
            if (-not $bay -and $hpe) { $bay = Get-Prop $hpe @('BayNumber','Id') }
            $model    = Get-Prop $u @('Model','model','Name','name','PartNumber','partNumber')
            $mfr      = Get-Prop $u @('Manufacturer','manufacturer')
            $serial   = Get-Prop $u @('SerialNumber','serialNumber','sparePartNumber')
            $part     = Get-Prop $u @('PartNumber','partNumber')
            $spare    = Get-Prop $u @('SparePartNumber','sparePartNumber')
            $cap      = Get-Prop $u @('PowerCapacityWatts','outputCapacityWatts','capacityWatts','outputWatts')
            $fw       = Get-Prop $u @('FirmwareVersion','firmwareVersion')
            $psuType  = Get-Prop $u @('PowerSupplyType','powerSupplyType')
            $lineV    = Get-Prop $u @('LineInputVoltage','lineInputVoltage')
            $lineVT   = Get-Prop $u @('LineInputVoltageType','lineInputVoltageType')
            $lastW    = Get-Prop $u @('LastPowerOutputWatts','lastPowerOutputWatts')
            $st       = Get-Prop $u @('status')
            $stState  = Get-Prop $u @('state')
            if (-not $st -and ($u.PSObject.Properties.Name -contains 'Status') -and $u.Status) {
                $st = Get-Prop $u.Status @('Health','HealthRollup')
                if (-not $stState) { $stState = Get-Prop $u.Status @('State') }
            }
            $avgW = ''; $maxW = ''; $hotplug = ''; $mismatched = ''; $ipduCap = ''; $psuState2 = ''
            if ($hpe) {
                $avgW       = Get-Prop $hpe @('AveragePowerOutputWatts')
                $maxW       = Get-Prop $hpe @('MaxPowerOutputWatts')
                $hotplug    = Get-Prop $hpe @('HotplugCapable')
                $mismatched = Get-Prop $hpe @('Mismatched')
                $ipduCap    = Get-Prop $hpe @('iPDUCapable')
                if ($hpe.PSObject.Properties.Name -contains 'PowerSupplyStatus' -and $hpe.PowerSupplyStatus) {
                    $psuState2 = Get-Prop $hpe.PowerSupplyStatus @('State')
                }
            }

            [void]$sb.AppendLine(("[PS{0}] {1}" -f $i, $model))
            if ($mfr)     { [void]$sb.AppendLine(("       Hersteller   : {0}" -f $mfr)) }
            if ($bay)     { [void]$sb.AppendLine(("       Bay/Slot     : {0}" -f $bay)) }
            if ($serial)  { [void]$sb.AppendLine(("       Serial       : {0}" -f $serial)) }
            if ($part)    { [void]$sb.AppendLine(("       PartNumber   : {0}" -f $part)) }
            if ($spare)   { [void]$sb.AppendLine(("       SparePart    : {0}" -f $spare)) }
            if ($cap)     { [void]$sb.AppendLine(("       Capacity (W) : {0}" -f $cap)) }
            if ($lastW -ne '' -and $null -ne $lastW) { [void]$sb.AppendLine(("       Last Out (W) : {0}" -f $lastW)) }
            if ($avgW -ne '' -and $null -ne $avgW)   { [void]$sb.AppendLine(("       Avg Out (W)  : {0}" -f $avgW)) }
            if ($maxW -ne '' -and $null -ne $maxW)   { [void]$sb.AppendLine(("       Max Out (W)  : {0}" -f $maxW)) }
            if ($lineV)   { [void]$sb.AppendLine(("       Line Voltage : {0} V {1}" -f $lineV, $lineVT)) }
            elseif ($lineVT) { [void]$sb.AppendLine(("       Line Type    : {0}" -f $lineVT)) }
            if ($psuType) { [void]$sb.AppendLine(("       Typ          : {0}" -f $psuType)) }
            if ($fw)      { [void]$sb.AppendLine(("       Firmware     : {0}" -f $fw)) }
            if ($hotplug -ne '' -and $null -ne $hotplug)       { [void]$sb.AppendLine(("       HotPlug      : {0}" -f $hotplug)) }
            if ($mismatched -ne '' -and $null -ne $mismatched) { [void]$sb.AppendLine(("       Mismatched   : {0}" -f $mismatched)) }
            if ($ipduCap -ne '' -and $null -ne $ipduCap)       { [void]$sb.AppendLine(("       iPDU-fähig   : {0}" -f $ipduCap)) }
            if ($psuState2) { [void]$sb.AppendLine(("       PSU-Status   : {0}" -f $psuState2)) }
            if ($st -or $stState) { [void]$sb.AppendLine(("       Status/State : {0} / {1}" -f $st, $stState)) }

            # iPDU-Details (HPE Oem)
            if ($hpe -and ($hpe.PSObject.Properties.Name -contains 'iPDU') -and $hpe.iPDU) {
                $ipdu = $hpe.iPDU
                $ipduId  = Get-Prop $ipdu @('Id')
                $ipduIp  = Get-Prop $ipdu @('IPAddress')
                $ipduMac = Get-Prop $ipdu @('MacAddress')
                $ipduMdl = Get-Prop $ipdu @('Model')
                $ipduSn  = Get-Prop $ipdu @('SerialNumber')
                $ipduSt  = ''
                if ($ipdu.PSObject.Properties.Name -contains 'iPDUStatus' -and $ipdu.iPDUStatus) {
                    $ipduSt = Get-Prop $ipdu.iPDUStatus @('State','Health')
                }
                if ($ipduId -or $ipduIp -or $ipduMdl) {
                    [void]$sb.AppendLine( "       iPDU:")
                    if ($ipduId)  { [void]$sb.AppendLine(("         Id         : {0}" -f $ipduId)) }
                    if ($ipduMdl) { [void]$sb.AppendLine(("         Modell     : {0}" -f $ipduMdl)) }
                    if ($ipduSn)  { [void]$sb.AppendLine(("         Serial     : {0}" -f $ipduSn)) }
                    if ($ipduIp)  { [void]$sb.AppendLine(("         IP         : {0}" -f $ipduIp)) }
                    if ($ipduMac) { [void]$sb.AppendLine(("         MAC        : {0}" -f $ipduMac)) }
                    if ($ipduSt)  { [void]$sb.AppendLine(("         Status     : {0}" -f $ipduSt)) }
                }
            }
        }
        [void]$sb.AppendLine("")
    } elseif ($rawPs) {
        # Diagnose: Endpunkt antwortet, aber unbekanntes Schema
        [void]$sb.AppendLine("--- Netzteile (Rohantwort) ---")
        try { [void]$sb.AppendLine(($rawPs | ConvertTo-Json -Depth 8)) } catch { [void]$sb.AppendLine([string]$rawPs) }
        [void]$sb.AppendLine("")
    }

    # Fans
    $fans = Get-Container $sh @('fans')
    if (-not $fans) {
        $uri = Get-Prop $sh @('uri')
        if ($uri) {
            $sub = Try-Rest -A $A -S $S -V $V -E "$uri/fans"
            if ($sub) { $fans = Get-Container $sub @('members','data'); if (-not $fans) { $fans = $sub } }
        }
    }
    if ($fans) {
        [void]$sb.AppendLine("--- Lüfter ---")
        foreach ($f in $fans) {
            [void]$sb.AppendLine(("  Bay {0,-3} {1,-20} Speed {2,-6} Status {3} / {4}" -f `
                (Get-Prop $f @('bayNumber','slot','position')),
                (Get-Prop $f @('name','model')),
                (Get-Prop $f @('speedRpm','speedPercentage','speed')),
                (Get-Prop $f @('status')),
                (Get-Prop $f @('state'))))
        }
        [void]$sb.AppendLine("")
    }

    # Thermal sensors / Temperatur
    $temps = Get-Container $sh @('temperatureSensors','thermalSensors','temperatures')
    if (-not $temps) {
        $uri = Get-Prop $sh @('uri')
        if ($uri) {
            $sub = Try-Rest -A $A -S $S -V $V -E "$uri/temperature"
            if ($sub) { $temps = Get-Container $sub @('members','data'); if (-not $temps) { $temps = $sub } }
        }
    }
    if ($temps) {
        [void]$sb.AppendLine("--- Temperatur-Sensoren ---")
        foreach ($t in $temps) {
            [void]$sb.AppendLine(("  {0,-30} {1,5} °C   crit {2,5}  status {3}" -f `
                (Get-Prop $t @('name','sensorName','location')),
                (Get-Prop $t @('reading','currentReading','temperature')),
                (Get-Prop $t @('criticalThreshold','upperCriticalThreshold')),
                (Get-Prop $t @('status'))))
        }
        [void]$sb.AppendLine("")
    }

    # Utilization (Strom-/Temperatur-Live)
    $uri = Get-Prop $sh @('uri')
    if ($uri) {
        $util = Try-Rest -A $A -S $S -V $V -E "$uri/utilization?fields=AveragePower,PeakPower,AmbientTemperature"
        if ($util -and $util.metricList) {
            [void]$sb.AppendLine("--- Utilization (Live) ---")
            foreach ($m in $util.metricList) {
                $vals = $m.metricSamples | Select-Object -First 1
                $val  = if ($vals) { $vals[1] } else { '' }
                [void]$sb.AppendLine(("  {0,-25} = {1}" -f (Get-Prop $m @('metricName')), $val))
            }
        }
    }

    return $sb.ToString()
}

function Build-Graphics {
    param($A, $S, $V, $sh)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("=== Grafikkarten / GPU ===")
    [void]$sb.AppendLine("(Hinweis: GPU-Inventar ist erst ab neueren OV-Versionen / Gen10+ verlässlich verfügbar.")
    [void]$sb.AppendLine(" In OV 6.60 / Gen7-9 fehlen diese Daten häufig komplett — dann nur iLO-Embedded-VGA.)")
    [void]$sb.AppendLine("")

    $found = $false

    # 1) Direkter Container im Hardware-Detail
    $gpus = Get-Container $sh @('gpuList','graphicsCards','gpus','accelerators')
    if (-not $gpus) {
        # 2) Sub-Endpunkt versuchen (existiert nicht überall)
        $uri = Get-Prop $sh @('uri')
        if ($uri) {
            foreach ($ep in @('/gpu','/graphics','/accelerators')) {
                $sub = Try-Rest -A $A -S $S -V $V -E "$uri$ep"
                if ($sub) {
                    $gpus = Get-Container $sub @('members','data','gpuList')
                    if (-not $gpus) { $gpus = $sub }
                    if ($gpus) { break }
                }
            }
        }
    }
    if ($gpus) {
        $found = $true
        [void]$sb.AppendLine("--- Diskrete / Add-In GPUs ---")
        $i = 0
        foreach ($g in $gpus) {
            $i++
            [void]$sb.AppendLine(("[GPU{0}] {1}" -f $i, (Get-Prop $g @('model','name','productName','deviceName'))))
            [void]$sb.AppendLine(("       Hersteller   : {0}" -f (Get-Prop $g @('manufacturer','vendor'))))
            [void]$sb.AppendLine(("       Slot         : {0}" -f (Get-Prop $g @('slot','location','slotNumber','pciSlot'))))
            [void]$sb.AppendLine(("       Memory       : {0}" -f (Get-Prop $g @('memorySize','memoryMb','memoryGB','frameBufferSize'))))
            [void]$sb.AppendLine(("       Serial       : {0}" -f (Get-Prop $g @('serialNumber','sn'))))
            [void]$sb.AppendLine(("       Part-Number  : {0}" -f (Get-Prop $g @('partNumber','sparePartNumber'))))
            [void]$sb.AppendLine(("       Firmware     : {0}" -f (Get-Prop $g @('firmwareVersion','version'))))
            [void]$sb.AppendLine(("       Status/State : {0} / {1}" -f (Get-Prop $g @('status')), (Get-Prop $g @('state'))))
        }
        [void]$sb.AppendLine("")
    }

    # 3) PCI-Devices durchsuchen (manche Versionen liefern nur das)
    $pci = Get-Container $sh @('pciDevices','pciCards')
    if ($pci) {
        $gpuPci = @()
        foreach ($p in $pci) {
            $cls = "$((Get-Prop $p @('deviceClass','class','classCode')))"
            $nam = "$((Get-Prop $p @('name','deviceName','productName','model')))"
            if ($cls -match 'VGA|Display|3D|Graphic' -or $nam -match 'NVIDIA|AMD|Radeon|Quadro|Tesla|Matrox|GeForce|MI[0-9]|GPU') {
                $gpuPci += $p
            }
        }
        if ($gpuPci.Count -gt 0) {
            $found = $true
            [void]$sb.AppendLine("--- PCI-Geräte (gefiltert: VGA / GPU) ---")
            foreach ($p in $gpuPci) {
                [void]$sb.AppendLine(("  Slot {0,-4} {1,-40} Vendor {2,-8} Device {3,-8} {4}" -f `
                    (Get-Prop $p @('slot','location','slotNumber')),
                    (Get-Prop $p @('name','deviceName','productName','model')),
                    (Get-Prop $p @('vendorId','vendor')),
                    (Get-Prop $p @('deviceId')),
                    (Get-Prop $p @('deviceClass','class'))))
            }
            [void]$sb.AppendLine("")
        }
    }

    # 4) Firmware-Inventory durchsuchen
    $uri = Get-Prop $sh @('uri')
    if ($uri) {
        $fw = Try-Rest -A $A -S $S -V $V -E "$uri/firmware"
        if ($fw) {
            $comps = Get-Container $fw @('components','firmwareComponents')
            if ($comps) {
                $gpuFw = @()
                foreach ($c in $comps) {
                    $n = "$((Get-Prop $c @('componentName','name')))"
                    if ($n -match 'NVIDIA|AMD|Radeon|Quadro|Tesla|Matrox|GeForce|GPU|VGA|Graphics|MI[0-9]') { $gpuFw += $c }
                }
                if ($gpuFw.Count -gt 0) {
                    $found = $true
                    [void]$sb.AppendLine("--- Firmware-Inventory (Treffer für GPU/VGA) ---")
                    foreach ($c in $gpuFw) {
                        [void]$sb.AppendLine(("  {0,-40} @ {1,-20} Version {2}" -f `
                            (Get-Prop $c @('componentName','name')),
                            (Get-Prop $c @('componentLocation','location')),
                            (Get-Prop $c @('componentVersion','version'))))
                    }
                    [void]$sb.AppendLine("")
                }
            }
        }
    }

    # 5) iLO-Embedded VGA als Fallback
    [void]$sb.AppendLine("--- Embedded / iLO Onboard-VGA ---")
    [void]$sb.AppendLine(("  iLO-Modell    : {0}" -f (Get-Prop $sh @('mpModel'))))
    [void]$sb.AppendLine(("  iLO-Firmware  : {0}" -f (Get-Prop $sh @('mpFirmwareVersion'))))
    [void]$sb.AppendLine("  (HPE iLO stellt einen Matrox/ASPEED-kompatiblen VGA-Controller bereit)")

    if (-not $found) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("(Keine diskrete GPU im Inventar gefunden — Server hat vermutlich nur Onboard-VGA,")
        [void]$sb.AppendLine(" oder OV-Version / Hardware-Generation liefert dieses Inventar nicht.)")
    }

    return $sb.ToString()
}

# =============================
# BIOS via iLO Redfish (Single-Sign-On ueber OneView)
# =============================
# Holt vom OneView Appliance ein iLO-SSO-Ticket und parst IP + Session-Token.
# Versucht /iloSsoUrl (GET und POST) und faellt auf /remoteConsoleUrl zurueck.
# Rueckgabe: @{ IloIp; Token; SsoUrl } oder $null.
function Get-IloRedfishSession {
    param([string]$A, [string]$S, [int]$V, $sh)
    $uri = Get-Prop $sh @('uri')
    if (-not $uri) { return $null }
    $hdr = @{ Auth = $S; "X-API-Version" = "$V"; 'Accept' = 'application/json' }

    $ssoUrl = $null
    $attempts = @(
        @{ M = 'Get';  E = "$uri/iloSsoUrl" },
        @{ M = 'Post'; E = "$uri/iloSsoUrl" },
        @{ M = 'Get';  E = "$uri/remoteConsoleUrl" },
        @{ M = 'Post'; E = "$uri/remoteConsoleUrl" }
    )
    foreach ($att in $attempts) {
        try {
            $r = Invoke-RestMethod -Uri "https://$A$($att.E)" -Method $att.M -Headers $hdr `
                -ContentType 'application/json' -SkipCertificateCheck -TimeoutSec 30 -ErrorAction Stop
            foreach ($p in @('iloSsoUrl','remoteConsoleUrl','url','ssoUrl')) {
                if ($r -and $r.PSObject.Properties.Name -contains $p -and $r.$p) { $ssoUrl = [string]$r.$p; break }
            }
            if ($ssoUrl) { break }
        } catch { continue }
    }
    if (-not $ssoUrl) { return $null }

    # SSO-URL kann sein:
    #   https://<iLO-IP>/sso?sessionKey=<token>...
    #   hplocons://addr=<iLO-IP>&sessionkey=<token>
    $iloIp = $null; $token = $null
    if ($ssoUrl -match '(?i)addr=([^&/?]+)')                  { $iloIp = $matches[1] }
    elseif ($ssoUrl -match '(?i)^https?://([^/]+)')           { $iloIp = $matches[1] }
    if ($ssoUrl -match '(?i)session(?:key|id)=([^&]+)')        { $token = $matches[1] }
    if (-not $iloIp -or -not $token) { return $null }
    return @{ IloIp = $iloIp; Token = $token; SsoUrl = $ssoUrl }
}

function Invoke-IloRedfish {
    param($IloIp, [hashtable]$AuthHeader, [string]$Path, [int]$TimeoutSec = 30)
    # Defensive: $IloIp kann (durch fehlerhafte OV-Antworten) auch ein Array sein
    if ($IloIp -is [System.Collections.IEnumerable] -and -not ($IloIp -is [string])) {
        $IloIp = @($IloIp)[0]
    }
    $IloIp = ("$IloIp").Trim().Trim('[',']')
    if ([string]::IsNullOrWhiteSpace($IloIp) -or ($IloIp -match '\s')) {
        throw "Ungueltige iLO-Adresse: '$IloIp'"
    }
    # IPv6 in URI muss in eckigen Klammern stehen
    $host_ = if ($IloIp -match ':' -and $IloIp -notmatch '^\[') { "[$IloIp]" } else { $IloIp }
    $h = @{ 'Accept' = 'application/json' } + $AuthHeader
    Invoke-RestMethod -Uri "https://$host_$Path" -Method Get -Headers $h `
        -SkipCertificateCheck -TimeoutSec $TimeoutSec -ErrorAction Stop
}

# Schneller TCP-Erreichbarkeitstest (Port 443) mit kurzem Timeout.
# Ein nicht erreichbares iLO wuerde sonst bei jedem HTTPS-Aufruf den vollen
# Invoke-RestMethod-Timeout (30 s) abwarten - bei mehreren Kandidat-IPs und
# mehreren Pfaden summiert sich das im UI-Thread zu Minuten ("haengt").
function Test-IloReachable {
    param([string]$IloIp, [int]$Port = 443, [int]$TimeoutMs = 2000)
    $addr = ("$IloIp").Trim().Trim('[',']')
    if ($addr -match '^(.+)%[^%]+$') { $addr = $matches[1] }  # IPv6 Zone-Index entfernen
    if ([string]::IsNullOrWhiteSpace($addr) -or ($addr -match '\s')) { return $false }
    $client = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect($addr, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            return $false   # Timeout -> nicht erreichbar
        }
        $client.EndConnect($iar)
        return $client.Connected
    } catch {
        return $false
    } finally {
        if ($client) { try { $client.Close() } catch { } }
    }
}

# Prueft eine iLO-Auth (SSO-Token ODER Basic-Auth-Header) mit einem schnellen GET.
# Es werden mehrere Pfade probiert; sobald einer 2xx liefert, gilt die Auth als ok.
# Ein 401/403 auf z.B. /redfish/v1/ ist kein endgueltiger Beweis fuer falsche
# Credentials (manche iLOs verweigern den Service-Root anonym, akzeptieren aber
# /Systems/1 mit gueltigem Token / Basic-Auth).
# Wenn -LastErrorRef uebergeben wird, landet die zuletzt gesehene Fehlermeldung
# darin (per Reference) - hilfreich fuer das GUI-Logging.
function Test-IloRedfishAuth {
    param(
        $IloIp,
        [hashtable]$AuthHeader,
        [ref]$LastErrorRef
    )
    # Defensive: $IloIp muss ein Einzel-String sein
    if ($IloIp -is [System.Collections.IEnumerable] -and -not ($IloIp -is [string])) {
        $IloIp = @($IloIp)[0]
    }
    $IloIp = ("$IloIp").Trim()
    $lastErr = $null
    # Fail-fast: ist Port 443 ueberhaupt offen? Spart pro nicht erreichbarem
    # iLO bis zu 3x den vollen HTTPS-Timeout.
    if (-not (Test-IloReachable -IloIp $IloIp -Port 443 -TimeoutMs 2000)) {
        $msg = "iLO $IloIp nicht erreichbar (TCP 443 Timeout)"
        if ($LastErrorRef) { $LastErrorRef.Value = $msg }
        return $false
    }
    foreach ($p in @('/redfish/v1/Systems/1','/redfish/v1/','/redfish/v1/Managers/1')) {
        try {
            $null = Invoke-IloRedfish -IloIp $IloIp -AuthHeader $AuthHeader -Path $p -TimeoutSec 10
            if ($LastErrorRef) { $LastErrorRef.Value = $null }
            return $true
        } catch {
            $lastErr = "[$p] $($_.Exception.Message)"
            continue
        }
    }
    if ($LastErrorRef) { $LastErrorRef.Value = $lastErr }
    return $false
}

# Hilfsfunktion: zerlegt einen beliebigen Wert (String, String[],
# Objekt mit .address, Liste solcher Objekte, ...) in einzelne
# Adress-Strings. Schreibt jeden Treffer einzeln in die Pipeline.
function Expand-IloAddressValue {
    param($v)
    if ($null -eq $v) { return }
    if ($v -is [string]) {
        foreach ($t in ($v -split '[\s,;]+')) {
            $t2 = ($t + '').Trim().Trim('[',']')
            if ($t2) { Write-Output $t2 }
        }
        return
    }
    if ($v -is [System.Collections.IDictionary]) {
        if ($v.Contains('address')) { Expand-IloAddressValue -v $v['address'] }
        return
    }
    if ($v -is [System.Collections.IEnumerable]) {
        foreach ($e in $v) { Expand-IloAddressValue -v $e }
        return
    }
    if ($v.PSObject -and ($v.PSObject.Properties.Name -contains 'address')) {
        Expand-IloAddressValue -v $v.address
        return
    }
    $s = "$v".Trim()
    if ($s) { Write-Output $s }
}

# Liefert eine priorisierte Liste moeglicher iLO-Adressen aus dem
# Server-Hardware-Objekt. Wichtig bei Synergy: mpIpAddresses enthaelt
# oft mehrere Eintraege (IPv6 LinkLocal + IPv4 Static/DHCP). LinkLocal
# (fe80::/169.254.) ist von extern nicht per Basic-Auth nutzbar und
# wird hinten angestellt.
function Get-IloIpCandidatesFromHardware {
    param($sh)
    if (-not $sh) { return @() }

    $seen   = New-Object System.Collections.Generic.HashSet[string]
    $scored = New-Object System.Collections.Generic.List[object]

    $addOne = {
        param([string]$addr, [string]$type)
        if ([string]::IsNullOrWhiteSpace($addr)) { return }
        if ($addr -match '\s') { return }
        # IPv6 Zone-Index entfernen (fe80::1%eth0 -> fe80::1)
        if ($addr -match '^(.+)%[^%]+$') { $addr = $matches[1] }
        if ($seen.Contains($addr)) { return }
        [void]$seen.Add($addr)
        $isV6      = $addr -match ':'
        $isV4Local = $addr -match '^169\.254\.'
        $isV6Local = $addr -match '^(?i)fe80:'
        $t = ($type + '').ToLowerInvariant()
        $prio = 0
        if ($t -eq 'linklocal' -or $isV6Local) { $prio = 3 }
        elseif ($isV4Local)                     { $prio = 2 }
        elseif ($isV6)                          { $prio = 1 }
        [void]$scored.Add([pscustomobject]@{ Addr = $addr; Prio = $prio })
    }

    if ($sh.PSObject.Properties.Name -contains 'mpHostInfo' -and $sh.mpHostInfo) {
        $mh = $sh.mpHostInfo
        if ($mh.PSObject.Properties.Name -contains 'mpIpAddresses' -and $mh.mpIpAddresses) {
            foreach ($ip in @($mh.mpIpAddresses)) {
                $type = ''
                if ($null -ne $ip -and -not ($ip -is [string]) -and $ip.PSObject -and ($ip.PSObject.Properties.Name -contains 'type')) {
                    $type = "$($ip.type)"
                }
                foreach ($a in @(Expand-IloAddressValue -v $ip)) {
                    & $addOne $a $type
                }
            }
        }
        if ($mh.PSObject.Properties.Name -contains 'mpHostName' -and $mh.mpHostName) {
            foreach ($a in @(Expand-IloAddressValue -v $mh.mpHostName)) { & $addOne $a '' }
        }
    }
    foreach ($f in @('mpDnsName','mpHostName','mpIpAddress')) {
        if ($sh.PSObject.Properties.Name -contains $f -and $sh.$f) {
            foreach ($a in @(Expand-IloAddressValue -v $sh.$f)) { & $addOne $a '' }
        }
    }

    $out = New-Object System.Collections.Generic.List[string]
    foreach ($x in ($scored | Sort-Object Prio)) { [void]$out.Add([string]$x.Addr) }
    # Einzelne Strings in die Pipeline schreiben (kein Array-Wrap-Trick)
    foreach ($s in $out) { Write-Output $s }
}

# Backwards-compat: erster Kandidat (beste Wahl).
function Get-IloIpFromHardware {
    param($sh)
    $c = Get-IloIpCandidatesFromHardware -sh $sh
    if ($c -and $c.Count -gt 0) { return $c[0] }
    return $null
}

function Build-Bios {
    param([string]$A, [string]$S, [int]$V, $sh, [string]$IloUser = '', [string]$IloPass = '')
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("=== BIOS Settings ===")
    [void]$sb.AppendLine("(Quelle 1: Server-Profile.bios.overriddenSettings - nur von OneView gesetzte Werte)")
    [void]$sb.AppendLine("(Quelle 2: iLO Redfish /redfish/v1/Systems/1/Bios - alle aktuellen BIOS-Attribute)")
    [void]$sb.AppendLine("")

    # --- Quelle 1: Server-Profile BIOS ---
    [void]$sb.AppendLine("--- BIOS Settings im Server-Profile (OneView-managed) ---")
    $profUri = Get-Prop $sh @('serverProfileUri')
    if (-not $profUri) {
        [void]$sb.AppendLine("(Kein Server-Profil zugewiesen)")
    } else {
        try {
            $prof = OV-Rest -A $A -S $S -V $V -M Get -E $profUri
            if ($prof.PSObject.Properties.Name -contains 'bios' -and $prof.bios) {
                [void]$sb.AppendLine(("manageBios          : {0}" -f (Get-Prop $prof.bios @('manageBios'))))
                $overr = $null
                if ($prof.bios.PSObject.Properties.Name -contains 'overriddenSettings') { $overr = $prof.bios.overriddenSettings }
                if ($overr -and @($overr).Count -gt 0) {
                    [void]$sb.AppendLine(("Overridden Settings : {0}" -f @($overr).Count))
                    foreach ($o in $overr) {
                        [void]$sb.AppendLine(("  {0,-44} = {1}" -f (Get-Prop $o @('id')), (Get-Prop $o @('value'))))
                    }
                } else {
                    [void]$sb.AppendLine("Overridden Settings : (keine)")
                }
            } else {
                [void]$sb.AppendLine("(Profil enthaelt keinen 'bios'-Block)")
            }
        } catch {
            [void]$sb.AppendLine("(Profil-BIOS nicht lesbar: $($_.Exception.Message))")
        }
    }
    [void]$sb.AppendLine("")

    # --- Quelle 2: iLO Redfish per OneView SSO ---
    [void]$sb.AppendLine("--- iLO Redfish (vollstaendige BIOS-Attribute) ---")

    $iloIp     = $null
    $authHdr   = $null
    $modeLabel = $null
    $ssoToken  = $null   # nur fuer spaeteres DELETE

    $haveCreds = (-not [string]::IsNullOrEmpty($IloUser)) -and (-not [string]::IsNullOrEmpty($IloPass))

    # 1) Wenn iLO-Credentials in der GUI eingetragen sind: Basic-Auth zuerst
    #    (entspricht dem Verhalten des Referenz-Skripts Get-ProLiantHardwareInfo-iLO6.ps1).
    #    SSO ist nur dann sinnvoll, wenn keine Credentials vorliegen.
    if ($haveCreds) {
        $candidates = @(Get-IloIpCandidatesFromHardware -sh $sh)
        # Sicherheits-Flatten: falls doch ein nested array entsteht, alles auf Einzel-Strings reduzieren
        $candidates = @($candidates | ForEach-Object {
            if ($_ -is [System.Collections.IEnumerable] -and -not ($_ -is [string])) { $_ } else { ,$_ }
        } | ForEach-Object { "$_" } | Where-Object { $_ -and ($_ -notmatch '\s') })
        if (-not $candidates -or $candidates.Count -eq 0) {
            [void]$sb.AppendLine("(Keine iLO-IP im Server-Hardware-Objekt gefunden)")
        } else {
            $b64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$IloUser`:$IloPass"))
            $candHdr = @{ 'Authorization' = "Basic $b64" }
            $errors = New-Object System.Collections.Generic.List[string]
            foreach ($ip in $candidates) {
                $err = $null
                if (Test-IloRedfishAuth -IloIp $ip -AuthHeader $candHdr -LastErrorRef ([ref]$err)) {
                    $iloIp     = $ip
                    $authHdr   = $candHdr
                    $modeLabel = 'iLO Basic-Auth'
                    break
                }
                $errors.Add(("  - {0}: {1}" -f $ip, $err)) | Out-Null
            }
            if (-not $authHdr) {
                [void]$sb.AppendLine("(iLO Basic-Auth mit den eingegebenen iLO-Credentials fehlgeschlagen. Details:)")
                foreach ($e in $errors) { [void]$sb.AppendLine($e) }
            }
        }
    }

    # 2) Fallback: OneView-SSO probieren (Managed Server ohne iLO-Credentials)
    if (-not $authHdr) {
        $sess = Get-IloRedfishSession -A $A -S $S -V $V -sh $sh
        if ($sess) {
            $candHdr = @{ 'X-Auth-Token' = $sess.Token }
            $ssoErr = $null
            if (Test-IloRedfishAuth -IloIp $sess.IloIp -AuthHeader $candHdr -LastErrorRef ([ref]$ssoErr)) {
                $iloIp     = $sess.IloIp
                $authHdr   = $candHdr
                $modeLabel = 'OneView-SSO (Managed)'
                $ssoToken  = $sess.Token
            } else {
                [void]$sb.AppendLine(("(OneView-SSO-Token wurde vom iLO {0} nicht akzeptiert: {1})" -f $sess.IloIp, $ssoErr))
            }
        }
    }

    if (-not $authHdr) {
        if (-not $sess -and (-not $IloUser -or -not $IloPass)) {
            [void]$sb.AppendLine("(Server ist vermutlich MONITORED: OneView-SSO nicht moeglich.")
            [void]$sb.AppendLine(" Bitte iLO User + iLO Pwd oben in der GUI eintragen und Suche neu ausloesen.)")
        } elseif (-not $sess) {
            [void]$sb.AppendLine("(Weder SSO noch Basic-Auth erfolgreich - BIOS-Vollabzug nicht moeglich.)")
        } else {
            [void]$sb.AppendLine("(SSO lieferte ein Token, das vom iLO abgelehnt wurde - Mode unklar.)")
        }
        return $sb.ToString()
    }

    [void]$sb.AppendLine(("Mode             : {0}" -f $modeLabel))
    [void]$sb.AppendLine(("iLO              : {0}" -f $iloIp))
    [void]$sb.AppendLine("")

    # --- Secure Boot (eigener Redfish-Endpunkt, NICHT in /Bios/Attributes!) ---
    [void]$sb.AppendLine("--- Secure Boot ---")
    try {
        $sb2 = Invoke-IloRedfish -IloIp $iloIp -AuthHeader $authHdr -Path "/redfish/v1/Systems/1/SecureBoot"
        if ($sb2) {
            $sbEnable  = Get-Prop $sb2 @('SecureBootEnable')
            $sbCurrent = Get-Prop $sb2 @('SecureBootCurrentBoot')
            $sbMode    = Get-Prop $sb2 @('SecureBootMode')
            $sbReset   = Get-Prop $sb2 @('ResetKeysType')
            [void]$sb.AppendLine(("  SecureBootEnable       : {0}" -f $sbEnable))
            [void]$sb.AppendLine(("  SecureBootCurrentBoot  : {0}" -f $sbCurrent))
            if ($sbMode)  { [void]$sb.AppendLine(("  SecureBootMode         : {0}" -f $sbMode)) }
            if ($sbReset) { [void]$sb.AppendLine(("  ResetKeysType          : {0}" -f $sbReset)) }
            $status = '?'
            if ("$sbEnable" -in @('True','true','1') -and "$sbCurrent" -in @('Enabled','enabled','True','true')) {
                $status = 'AKTIV (enabled & enforcing)'
            } elseif ("$sbEnable" -in @('True','true','1')) {
                $status = 'eingeschaltet, aber CurrentBoot != Enabled (Reboot erforderlich?)'
            } elseif ("$sbEnable" -in @('False','false','0')) {
                $status = 'DEAKTIVIERT'
            }
            [void]$sb.AppendLine(("  => Status              : {0}" -f $status))
        } else {
            [void]$sb.AppendLine("  (keine SecureBoot-Antwort)")
        }
    } catch {
        [void]$sb.AppendLine("  FEHLER beim Lesen von /redfish/v1/Systems/1/SecureBoot: $($_.Exception.Message)")
    }
    [void]$sb.AppendLine("")

    # /Bios -> Attributes
    try {
        $bios = Invoke-IloRedfish -IloIp $iloIp -AuthHeader $authHdr -Path "/redfish/v1/Systems/1/Bios"
        if ($bios.PSObject.Properties.Name -contains 'AttributeRegistry' -and $bios.AttributeRegistry) {
            [void]$sb.AppendLine(("AttributeRegistry: {0}" -f $bios.AttributeRegistry))
        }
        if ($bios.PSObject.Properties.Name -contains 'Attributes' -and $bios.Attributes) {
            $attrs = $bios.Attributes
            $names = @($attrs.PSObject.Properties.Name) | Sort-Object
            [void]$sb.AppendLine(("Anzahl Settings  : {0}" -f $names.Count))
            [void]$sb.AppendLine("")
            foreach ($n in $names) {
                $v = $attrs.$n
                $vs = ''
                if ($null -ne $v) {
                    if ($v -is [string]) {
                        $vs = $v
                    } elseif ($v.GetType().Name -eq 'PSCustomObject' -or ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string]))) {
                        try { $vs = ($v | ConvertTo-Json -Compress -Depth 4) } catch { $vs = [string]$v }
                    } else {
                        $vs = [string]$v
                    }
                }
                [void]$sb.AppendLine(("  {0,-46} = {1}" -f $n, $vs))
            }
        } else {
            [void]$sb.AppendLine("(keine 'Attributes' im /Bios-Objekt)")
        }
    } catch {
        if ($_.Exception.Message -notmatch 'Cannot convert value.*System\.Int32') {
            [void]$sb.AppendLine("FEHLER beim Lesen von /redfish/v1/Systems/1/Bios: $($_.Exception.Message)")
        }
    }

    # /Bios/Settings -> pending changes
    try {
        $pend = Invoke-IloRedfish -IloIp $iloIp -AuthHeader $authHdr -Path "/redfish/v1/Systems/1/Bios/Settings"
        if ($pend -and $pend.PSObject.Properties.Name -contains 'Attributes' -and $pend.Attributes) {
            $pendNames = @($pend.Attributes.PSObject.Properties.Name)
            if ($pendNames.Count -gt 0) {
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine(("--- Pending BIOS Changes ({0}) ---" -f $pendNames.Count))
                foreach ($n in ($pendNames | Sort-Object)) {
                    $pv = $pend.Attributes.$n
                    $pvs = ''
                    if ($null -ne $pv) {
                        if ($pv -is [string]) {
                            $pvs = $pv
                        } elseif ($pv.GetType().Name -eq 'PSCustomObject' -or ($pv -is [System.Collections.IEnumerable] -and -not ($pv -is [string]))) {
                            try { $pvs = ($pv | ConvertTo-Json -Compress -Depth 4) } catch { $pvs = [string]$pv }
                        } else {
                            $pvs = [string]$pv
                        }
                    }
                    [void]$sb.AppendLine(("  {0,-46} -> {1}" -f $n, $pvs))
                }
            }
        }
    } catch { }

    # /Systems/1 -> Boot (BIOS-relevant)
    try {
        $sys = Invoke-IloRedfish -IloIp $iloIp -AuthHeader $authHdr -Path "/redfish/v1/Systems/1"
        if ($sys -and $sys.PSObject.Properties.Name -contains 'Boot' -and $sys.Boot) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Boot-Konfiguration ---")
            [void]$sb.AppendLine(("BootSourceOverrideEnabled : {0}" -f (Get-Prop $sys.Boot @('BootSourceOverrideEnabled'))))
            [void]$sb.AppendLine(("BootSourceOverrideTarget  : {0}" -f (Get-Prop $sys.Boot @('BootSourceOverrideTarget'))))
            [void]$sb.AppendLine(("BootSourceOverrideMode    : {0}" -f (Get-Prop $sys.Boot @('BootSourceOverrideMode'))))
            $bo = $null
            if ($sys.Boot.PSObject.Properties.Name -contains 'BootOrder') { $bo = $sys.Boot.BootOrder }
            if ($bo) {
                [void]$sb.AppendLine("BootOrder:")
                $i = 0
                foreach ($b in @($bo)) { $i++; [void]$sb.AppendLine(("  {0,2}. {1}" -f $i, $b)) }
            }
        }
    } catch { }

    # iLO-Session abmelden (nur wenn wir per SSO einen Token bekommen haben)
    if ($ssoToken) {
        try {
            Invoke-RestMethod -Uri "https://$iloIp/redfish/v1/SessionService/Sessions/$ssoToken" `
                -Method Delete -Headers @{ 'X-Auth-Token' = $ssoToken } -SkipCertificateCheck -TimeoutSec 10 -EA SilentlyContinue | Out-Null
        } catch { }
    }

    return $sb.ToString()
}

# Rekursiver Flach-Dump aller Felder eines Objekts.
# Gibt Zeilen wie "key.subkey[0].field = value" zurück. Begrenzt Tiefe und
# Listen-Länge, um die Anzeige nicht zu sprengen.
function Build-AllFields {
    param($obj, [int]$maxDepth = 8, [int]$maxItems = 200)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("=== Alle Felder (rekursiver Dump) ===")
    [void]$sb.AppendLine("")
    if ($null -eq $obj) {
        [void]$sb.AppendLine("(null)")
        return $sb.ToString()
    }
    # Stack: @(prefix, value, depth)
    $stack = New-Object System.Collections.Stack
    $stack.Push(@('', $obj, 0))
    $count = 0
    while ($stack.Count -gt 0 -and $count -lt 5000) {
        $entry = $stack.Pop()
        $prefix = $entry[0]; $val = $entry[1]; $depth = $entry[2]
        if ($null -eq $val) {
            [void]$sb.AppendLine(("{0} = (null)" -f $prefix)); $count++; continue
        }
        if ($depth -ge $maxDepth) {
            [void]$sb.AppendLine(("{0} = ... (max-depth)" -f $prefix)); $count++; continue
        }
        $t = $val.GetType()
        # Primitive / String / Date / Enum -> direkt ausgeben
        if ($t.IsPrimitive -or $val -is [string] -or $val -is [datetime] -or $val -is [decimal] -or $t.IsEnum) {
            $s = "$val"
            if ($s.Length -gt 400) { $s = $s.Substring(0,400) + '...' }
            [void]$sb.AppendLine(("{0} = {1}" -f $prefix, $s))
            $count++
            continue
        }
        # Array / List
        if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string]) -and -not ($val -is [System.Collections.IDictionary])) {
            $idx = 0
            $items = @($val)
            if ($items.Count -eq 0) {
                [void]$sb.AppendLine(("{0} = (leeres Array)" -f $prefix)); $count++; continue
            }
            $limit = [Math]::Min($items.Count, $maxItems)
            # In umgekehrter Reihenfolge auf Stack pushen damit Ausgabe sortiert bleibt
            for ($i = $limit - 1; $i -ge 0; $i--) {
                $stack.Push(@(("{0}[{1}]" -f $prefix, $i), $items[$i], $depth + 1))
            }
            if ($items.Count -gt $limit) {
                [void]$sb.AppendLine(("{0} = ... ({1} weitere Einträge)" -f $prefix, ($items.Count - $limit)))
            }
            continue
        }
        # Hashtable / IDictionary
        if ($val -is [System.Collections.IDictionary]) {
            $keys = @($val.Keys)
            for ($i = $keys.Count - 1; $i -ge 0; $i--) {
                $k = $keys[$i]
                $sub = if ($prefix) { "$prefix.$k" } else { "$k" }
                $stack.Push(@($sub, $val[$k], $depth + 1))
            }
            continue
        }
        # PSObject / sonstige Objekte mit Properties
        if ($val.PSObject -and $val.PSObject.Properties.Count -gt 0) {
            $props = @($val.PSObject.Properties)
            for ($i = $props.Count - 1; $i -ge 0; $i--) {
                $p = $props[$i]
                $sub = if ($prefix) { "$prefix.$($p.Name)" } else { "$($p.Name)" }
                try {
                    $stack.Push(@($sub, $p.Value, $depth + 1))
                } catch {
                    [void]$sb.AppendLine(("{0} = (n/a)" -f $sub)); $count++
                }
            }
            continue
        }
        # Fallback
        [void]$sb.AppendLine(("{0} = {1}" -f $prefix, "$val"))
        $count++
    }
    if ($count -ge 5000) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("(Ausgabe nach 5000 Zeilen abgeschnitten)")
    }
    return $sb.ToString()
}

# =============================
# Suche
# =============================
function Run-Search {
    $script:hitObjects = @()
    $script:lastDetailRow = -1
    $dgvHits.Rows.Clear()
    $txtOverview.Clear(); $txtCpu.Clear(); $txtRaw.Clear(); $txtProfile.Clear()
    $txtStorage.Clear(); $txtPower.Clear(); $txtGpu.Clear(); $txtAll.Clear()
    $txtNet.Clear(); $txtBios.Clear()
    $dgvFw.Rows.Clear(); $dgvNet.Rows.Clear()
    $btnExportTxt.Enabled = $false
    $btnExportHtml.Enabled = $false

    $u = $txtUser.Text; $p = $txtPass.Text; $q = $txtSearch.Text.Trim()
    if ([string]::IsNullOrEmpty($u) -or [string]::IsNullOrEmpty($p)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Benutzername und Passwort eingeben.", "Fehler",
            'OK', 'Error') | Out-Null; return
    }
    if ([string]::IsNullOrEmpty($q)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Servername oder Seriennummer eingeben.", "Fehler",
            'OK', 'Error') | Out-Null; return
    }

    # Ausgewählte Appliances einsammeln
    $appliances = @()
    $tagCounter = 0
    for ($i = 0; $i -lt $chkAppliances.Items.Count; $i++) {
        if ($chkAppliances.GetItemChecked($i)) {
            $ip = Get-IPFromItem $chkAppliances.Items[$i].ToString()
            if ($ip) {
                $tagCounter++
                $tag = "APP-{0:D2}" -f $tagCounter
                $appliances += [pscustomobject]@{ Index = $i; IP = $ip; Tag = $tag }
            }
        }
    }
    if ($appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Bitte mindestens eine Appliance auswählen.", "Fehler",
            'OK', 'Warning') | Out-Null; return
    }

    $btnSearch.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $statusLabel.Text = "Suche '$q' auf $($appliances.Count) Appliance(s)..."
    $form.Refresh()
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

    # Debug-Log neben dem Skript (Pattern aus OneView_Update: Datei-basierte Sichtbarkeit)
    $logPath = Join-Path $PSScriptRoot 'ServerInfo_Search.log'
    function Write-SearchLog {
        param([string]$Msg)
        try {
            $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
            Add-Content -Path $logPath -Value "[$ts] $Msg" -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch { }
    }
    Write-SearchLog ("--- Suche gestartet: '$q' auf {0} Appliance(s) ---" -f $appliances.Count)

    # Mapping IP <-> Tag NUR in einem separaten lokalen Mapping-File
    # (nicht im Haupt-Log, damit das anonymisierte Hauptlog gefahrlos
    # geteilt werden kann).
    $mapPath = Join-Path $PSScriptRoot 'ServerInfo_Search.map.log'
    try {
        $mapHeader = "[{0}] === Mapping fuer Suche '$q' ===" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Add-Content -Path $mapPath -Value $mapHeader -Encoding UTF8 -ErrorAction SilentlyContinue
        foreach ($ap in $appliances) {
            Add-Content -Path $mapPath -Value ("  {0}  =  {1}" -f $ap.Tag, $ap.IP) -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    } catch { }

    # Bestehende Sessions schliessen
    foreach ($oldH in $script:hitObjects) {
        try { OV-Logout -A $oldH.Appliance -S $oldH.Session -V $oldH.ApiVersion } catch { }
    }
    $script:hitObjects = @()

    # =====================================================
    # PARALLEL pro Appliance via ForEach-Object -Parallel
    # Bei 40+ Appliances bringt das den Faktor ~ThrottleLimit Speedup.
    # Jeder Worker schreibt ins gleiche Log (Add-Content mit Retry, threadsicher).
    # Es werden NUR primitive Werte ($using:) reingegeben - keine Form-Objekte.
    # =====================================================
    $throttle = [Math]::Min($appliances.Count, 12)
    $statusLabel.Text = "Suche '$q' parallel ($throttle gleichzeitig) auf $($appliances.Count) Appliance(s)..."
    $form.Refresh()

    $allResults = $appliances | ForEach-Object -ThrottleLimit $throttle -Parallel {
        $A   = $_.IP
        $idx = $_.Index
        $Tag = $_.Tag       # anonymisierter Bezeichner fuer Logs
        $U   = $using:u
        $P   = $using:p
        $Q   = $using:q
        $LP  = $using:logPath

        # Threadsicher anhaengen mit kleinem Retry-Loop (Add-Content kann bei
        # gleichzeitigem Zugriff gelegentlich kollidieren).
        function Write-Log {
            param([string]$M)
            $line = "[{0}] {1}" -f (Get-Date).ToString('HH:mm:ss.fff'), $M
            for ($try = 0; $try -lt 5; $try++) {
                try {
                    Add-Content -Path $using:logPath -Value $line -Encoding UTF8 -ErrorAction Stop
                    return
                } catch {
                    Start-Sleep -Milliseconds 20
                }
            }
        }

        $result = [pscustomobject]@{
            Appliance    = $A
            Tag          = $Tag
            Index        = $idx
            ApiVersion   = $null
            VerLabel     = '?'
            Session      = $null
            Found        = @()
            Error        = $null
            Stage        = 'start'
            TotalServers = 0
        }

        try {
            # 1) API-Version
            $result.Stage = 'version'
            $verResp = Invoke-RestMethod -Uri "https://$A/rest/version" -Method Get `
                -SkipCertificateCheck -TimeoutSec 30
            $V = $verResp.currentVersion
            if (-not $V) { throw "Konnte currentVersion nicht ermitteln." }
            $result.ApiVersion = $V
            $result.VerLabel   = [string]$V
            Write-Log "[$Tag] API-Version: $V"

            # 2) Login
            $result.Stage = 'login'
            $loginBody = @{ userName = $U; password = $P; authLoginDomain = 'Local' } | ConvertTo-Json -Compress
            $loginHdr  = @{ 'X-API-Version' = "$V"; 'Accept' = 'application/json' }
            $loginResp = Invoke-RestMethod -Uri "https://$A/rest/login-sessions" -Method Post `
                -Body $loginBody -ContentType 'application/json' -Headers $loginHdr `
                -SkipCertificateCheck -TimeoutSec 30
            $S = $loginResp.sessionID
            if (-not $S) { throw "Login lieferte keine sessionID." }
            $result.Session = $S
            Write-Log "[$Tag] Login OK"

            $hdr = @{ 'X-API-Version' = "$V"; 'Auth' = $S; 'Accept' = 'application/json' }

            # 3) Server-seitige Filter (name, serverName UND serialNumber)
            $result.Stage = 'filter-search'
            $qEsc = ([string]$Q).Replace("'", "''")
            # ArrayList ist in Parallel-Runspaces robuster als Generic List<T>
            # bzgl. Typkonvertierung und @()-Enumeration.
            $found = New-Object System.Collections.ArrayList
            foreach ($field in @('name','serverName','serialNumber')) {
                $ep = "/rest/server-hardware?filter=$field='$qEsc'&start=0&count=50"
                try {
                    $r = Invoke-RestMethod -Uri "https://$A$ep" -Method Get -Headers $hdr `
                        -SkipCertificateCheck -TimeoutSec 60
                    $cnt = 0
                    if ($null -ne $r -and $r.PSObject.Properties.Name -contains 'members' -and $null -ne $r.members) {
                        foreach ($m in @($r.members)) {
                            if ($null -ne $m) { [void]$found.Add($m); $cnt++ }
                        }
                    }
                    Write-Log "[$Tag] Filter $field='$Q' -> $cnt Treffer"
                } catch {
                    Write-Log "[$Tag] Filter $field FEHLER: $($_.Exception.Message)"
                }
            }

            # 3b) Auch im Server-Profil-Namen suchen und die zugehoerige
            #     Server-Hardware ueber serverHardwareUri aufloesen.
            try {
                $profEp = "/rest/server-profiles?start=0&count=200"
                $profHits = New-Object System.Collections.ArrayList
                $nextProf = $profEp
                while (-not [string]::IsNullOrEmpty([string]$nextProf)) {
                    $rp = $null
                    try {
                        $rp = Invoke-RestMethod -Uri "https://$A$nextProf" -Method Get -Headers $hdr `
                            -SkipCertificateCheck -TimeoutSec 60
                    } catch {
                        Write-Log "[$Tag] Profilliste FEHLER: $($_.Exception.Message)"; break
                    }
                    if ($null -ne $rp -and $rp.PSObject.Properties.Name -contains 'members' -and $null -ne $rp.members) {
                        foreach ($pm in @($rp.members)) {
                            if ($null -ne $pm) { [void]$profHits.Add($pm) }
                        }
                    }
                    $nextProf = $null
                    if ($null -ne $rp -and $rp.PSObject.Properties.Name -contains 'nextPageUri' -and $null -ne $rp.nextPageUri) {
                        $nextProf = [string]$rp.nextPageUri
                        if ([string]::IsNullOrEmpty($nextProf)) { $nextProf = $null }
                    }
                }
                $needleP = ''
                try { $needleP = ([string]$Q).ToLowerInvariant() } catch { $needleP = '' }
                $useRxP = $false; $rxP = $null
                if (-not [string]::IsNullOrEmpty($needleP) -and ($needleP.Contains('*') -or $needleP.Contains('?'))) {
                    try {
                        $escP = [regex]::Escape($needleP)
                        $patP = $escP -replace '\\\*', '.*' -replace '\\\?', '.'
                        $rxP = New-Object System.Text.RegularExpressions.Regex("^.*$patP.*$", 'IgnoreCase,CultureInvariant')
                        $useRxP = $true
                    } catch { $useRxP = $false }
                }
                $profMatched = 0
                $profSkippedUnassigned = 0
                # bereits gefundene Server-Hardware-URIs vormerken (Doppel-GETs vermeiden)
                $foundUris = @{}
                foreach ($fx in $found) {
                    if ($null -ne $fx -and $fx.PSObject.Properties.Name -contains 'uri' -and $null -ne $fx.uri) {
                        $foundUris[[string]$fx.uri] = $true
                    }
                }
                foreach ($pm in $profHits) {
                    $pn = ''
                    if ($null -ne $pm -and $pm.PSObject.Properties.Name -contains 'name' -and $null -ne $pm.name) {
                        $pn = ([string]$pm.name).ToLowerInvariant()
                    }
                    $isHitP = $false
                    if ($useRxP) { if ($rxP.IsMatch($pn)) { $isHitP = $true } }
                    else { if (-not [string]::IsNullOrEmpty($needleP) -and $pn.Contains($needleP)) { $isHitP = $true } }
                    if (-not $isHitP) { continue }
                    $shUri = $null
                    if ($pm.PSObject.Properties.Name -contains 'serverHardwareUri' -and $null -ne $pm.serverHardwareUri) {
                        $shUri = [string]$pm.serverHardwareUri
                    }
                    # Profil ohne zugewiesene Server-Hardware (z.B. unassigned/Template-aehnlich) -> ueberspringen
                    if ([string]::IsNullOrEmpty($shUri)) {
                        $profSkippedUnassigned++
                        continue
                    }
                    # Server-Hardware bereits in Trefferliste -> kein zusaetzlicher GET
                    if ($foundUris.ContainsKey($shUri)) { $profMatched++; continue }
                    try {
                        $shObj = Invoke-RestMethod -Uri "https://$A$shUri" -Method Get -Headers $hdr `
                            -SkipCertificateCheck -TimeoutSec 60
                        if ($null -ne $shObj) {
                            [void]$found.Add($shObj)
                            $foundUris[$shUri] = $true
                            $profMatched++
                        }
                    } catch {
                        Write-Log "[$Tag] Profil->Server-Hardware FEHLER ($shUri): $($_.Exception.Message)"
                    }
                }
                Write-Log "[$Tag] Profil-Name-Match: $profMatched Server (Profile geprueft: $($profHits.Count), unassigned uebersprungen: $profSkippedUnassigned)"
            } catch {
                Write-Log "[$Tag] Profil-Suche FEHLER: $($_.Exception.Message)"
            }

            # 4) IMMER Fallback: ganze Liste laden + Substring-Match
            $result.Stage = 'fallback-list'
            $fallbackErr = $null
            try {
                $all = New-Object System.Collections.ArrayList
                $next = "/rest/server-hardware?start=0&count=1000"
                $pageNo = 0
                while (-not [string]::IsNullOrEmpty([string]$next)) {
                    $pageNo++
                    $r = $null
                    try {
                        $r = Invoke-RestMethod -Uri "https://$A$next" -Method Get -Headers $hdr `
                            -SkipCertificateCheck -TimeoutSec 120
                    } catch {
                        Write-Log "[$Tag] Vollliste Seite $pageNo FEHLER: $($_.Exception.Message)"
                        break
                    }
                    if ($null -ne $r) {
                        $hasMembers = ($r.PSObject.Properties.Name -contains 'members')
                        if ($hasMembers -and $null -ne $r.members) {
                            foreach ($m in @($r.members)) {
                                if ($null -ne $m) { [void]$all.Add($m) }
                            }
                        }
                        $hasNext = ($r.PSObject.Properties.Name -contains 'nextPageUri')
                        if ($hasNext -and $null -ne $r.nextPageUri) {
                            $next = [string]$r.nextPageUri
                            if ([string]::IsNullOrEmpty($next)) { $next = $null }
                        } else {
                            $next = $null
                        }
                    } else {
                        $next = $null
                    }
                }
                $result.TotalServers = $all.Count
                Write-Log "[$Tag] Vollliste: $($all.Count) Server"

                $needle = ''
                try { $needle = ([string]$Q).ToLowerInvariant() } catch { $needle = '' }

                # Wildcards (* und ?) -> Regex; sonst Substring-Match.
                $useRegex = $false
                $rx = $null
                if (-not [string]::IsNullOrEmpty($needle) -and ($needle.Contains('*') -or $needle.Contains('?'))) {
                    try {
                        $escaped = [regex]::Escape($needle)
                        # Escape() macht * -> \* und ? -> \?  - zurueck zu .* bzw. .
                        $pattern = $escaped -replace '\\\*', '.*' -replace '\\\?', '.'
                        $rx = New-Object System.Text.RegularExpressions.Regex("^.*$pattern.*$", 'IgnoreCase,CultureInvariant')
                        $useRegex = $true
                        Write-Log "[$Tag] Wildcard-Suche aktiv, Pattern: $pattern"
                    } catch {
                        Write-Log "[$Tag] Wildcard-Pattern ungueltig, falle auf Substring zurueck: $($_.Exception.Message)"
                        $useRegex = $false
                    }
                }

                $clientHitCount = 0
                if (-not [string]::IsNullOrEmpty($needle)) {
                    foreach ($m in $all) {
                        try {
                            $nm = ''; $sn = ''; $mo = ''; $svn = ''
                            if ($null -ne $m -and $m.PSObject.Properties.Name -contains 'name'         -and $null -ne $m.name)         { $nm  = ([string]$m.name).ToLowerInvariant() }
                            if ($null -ne $m -and $m.PSObject.Properties.Name -contains 'serialNumber' -and $null -ne $m.serialNumber) { $sn  = ([string]$m.serialNumber).ToLowerInvariant() }
                            if ($null -ne $m -and $m.PSObject.Properties.Name -contains 'model'        -and $null -ne $m.model)        { $mo  = ([string]$m.model).ToLowerInvariant() }
                            if ($null -ne $m -and $m.PSObject.Properties.Name -contains 'serverName'   -and $null -ne $m.serverName)   { $svn = ([string]$m.serverName).ToLowerInvariant() }
                            $isHit = $false
                            if ($useRegex) {
                                if ($rx.IsMatch($nm) -or $rx.IsMatch($sn) -or $rx.IsMatch($mo) -or $rx.IsMatch($svn)) { $isHit = $true }
                            } else {
                                if ($nm.Contains($needle) -or $sn.Contains($needle) -or $mo.Contains($needle) -or $svn.Contains($needle)) { $isHit = $true }
                            }
                            if ($isHit) {
                                [void]$found.Add($m)
                                $clientHitCount++
                            }
                        } catch {
                            Write-Log "[$Tag] Match FEHLER fuer Eintrag: $($_.Exception.Message)"
                        }
                    }
                }
                Write-Log "[$Tag] Client-Match: $clientHitCount Treffer (gesamt-found: $($found.Count), regex=$useRegex)"
            } catch {
                $fallbackErr = $_
                $line = 0
                try { $line = [int]$_.InvocationInfo.ScriptLineNumber } catch { }
                Write-Log "[$Tag] Vollliste FEHLER (aussen) Zeile=$line : $($_.Exception.Message) | $($_.ScriptStackTrace)"
            }
            if ($null -ne $fallbackErr -and $found.Count -eq 0) {
                throw "Server-Hardware-Liste nicht abrufbar: $($fallbackErr.Exception.Message)"
            }

            # 5) Dedup ueber uri (hashtable-basiert)
            $result.Stage = 'dedup'
            $seen = @{}
            $deduped = New-Object System.Collections.ArrayList
            foreach ($sh in $found) {
                if ($null -eq $sh) { continue }
                $key = $null
                try {
                    if ($sh.PSObject.Properties.Name -contains 'uri' -and $null -ne $sh.uri) {
                        $key = [string]$sh.uri
                    }
                } catch { $key = $null }
                if ([string]::IsNullOrEmpty($key)) {
                    $sn = ''; $nm = ''
                    if ($sh.PSObject.Properties.Name -contains 'serialNumber' -and $null -ne $sh.serialNumber) { $sn = [string]$sh.serialNumber }
                    if ($sh.PSObject.Properties.Name -contains 'name'         -and $null -ne $sh.name)         { $nm = [string]$sh.name }
                    $key = "noUri::${sn}::${nm}"
                }
                if (-not $seen.ContainsKey($key)) {
                    $seen[$key] = $true
                    [void]$deduped.Add($sh)
                }
            }
            $found = $deduped
            Write-Log "[$Tag] Nach Dedup: $($found.Count) eindeutige Treffer"

            # 6) Enclosure-Location aufloesen - PRO SERVER try/catch,
            #    damit ein einzelner kaputter Eintrag nicht die ganze Trefferliste killt
            $result.Stage = 'enclosure-lookup'
            $encCache = @{}
            $hits = New-Object System.Collections.ArrayList
            # Shortcut: keine Treffer -> Stage komplett ueberspringen.
            # Generic.List[object] kann in Parallel-Runspaces "Argument types do not match" werfen.
            if ($null -eq $found -or @($found).Count -eq 0) {
                $result.Found = @()
                $result.Stage = 'done'
                Write-Log "[$Tag] FERTIG - 0 Treffer (keine Enclosure-Aufloesung noetig)"
                return $result
            }
            foreach ($sh in $found) {
                try {
                    # Sichere Property-Zugriffe (PSCustomObject kann Property fehlen)
                    $locUri = $null
                    if ($sh.PSObject.Properties.Name -contains 'locationUri') {
                        $locUri = [string]$sh.locationUri
                    }
                    $formFactor = ''
                    if ($sh.PSObject.Properties.Name -contains 'formFactor') {
                        $formFactor = [string]$sh.formFactor
                    }
                    $bay = ''
                    if ($sh.PSObject.Properties.Name -contains 'position' -and $null -ne $sh.position) {
                        $bay = [string]$sh.position
                    } elseif ($sh.PSObject.Properties.Name -contains 'serverBay' -and $null -ne $sh.serverBay) {
                        $bay = [string]$sh.serverBay
                    }

                    $loc = ''
                    if (-not [string]::IsNullOrEmpty($locUri)) {
                        if (-not $encCache.ContainsKey($locUri)) {
                            try {
                                $encCache[$locUri] = Invoke-RestMethod -Uri "https://$A$locUri" `
                                    -Method Get -Headers $hdr -SkipCertificateCheck -TimeoutSec 30
                            } catch {
                                $encCache[$locUri] = $null
                                Write-Log "[$Tag] Enclosure-GET '<uri>' FEHLER: $($_.Exception.Message)"
                            }
                        }
                        $enc = $encCache[$locUri]
                        if ($enc) {
                            $encName = if ($enc.PSObject.Properties.Name -contains 'name') { [string]$enc.name } else { '' }
                            $encSn   = if ($enc.PSObject.Properties.Name -contains 'serialNumber') { [string]$enc.serialNumber } else { '' }
                            if ($encSn) {
                                $loc = "$encName (SN $encSn), Bay $bay"
                            } else {
                                $loc = "$encName, Bay $bay"
                            }
                        } else {
                            $loc = "Enclosure (?), Bay $bay"
                        }
                    }
                    if (-not $loc) {
                        if ($formFactor -match 'Blade|HalfHeight|FullHeight') {
                            $loc = "Blade, Bay $bay (Frame unbekannt)"
                        } else {
                            $loc = "Rack-/Standalone-Server"
                        }
                    }

                    # Server-Profil-Name aufloesen (falls vorhanden). Server ohne
                    # zugewiesenes Profil bleiben mit leerem Namen stehen.
                    $profileName = ''
                    $profUri2 = $null
                    if ($sh.PSObject.Properties.Name -contains 'serverProfileUri' -and $null -ne $sh.serverProfileUri) {
                        $profUri2 = [string]$sh.serverProfileUri
                    }
                    if (-not [string]::IsNullOrEmpty($profUri2)) {
                        try {
                            $profObj2 = Invoke-RestMethod -Uri "https://$A$profUri2" -Method Get -Headers $hdr `
                                -SkipCertificateCheck -TimeoutSec 30
                            if ($null -ne $profObj2 -and $profObj2.PSObject.Properties.Name -contains 'name' -and $null -ne $profObj2.name) {
                                $profileName = [string]$profObj2.name
                            }
                        } catch {
                            Write-Log "[$Tag] Profil-Name-Lookup FEHLER: $($_.Exception.Message)"
                        }
                    }
                    # Fallback: serverName (z.B. iLO-/OS-Hostname) wenn kein Profil zugewiesen
                    if ([string]::IsNullOrEmpty($profileName) -and $sh.PSObject.Properties.Name -contains 'serverName' -and $null -ne $sh.serverName) {
                        $sn2 = [string]$sh.serverName
                        if (-not [string]::IsNullOrEmpty($sn2)) { $profileName = $sn2 }
                    }
                    [void]$hits.Add([pscustomobject]@{ ServerHw = $sh; Location = $loc; ProfileName = $profileName })
                } catch {
                    # Einzelner Server kaputt: trotzdem ohne Location uebernehmen
                    Write-Log ("[$Tag] Location-Lookup FEHLER fuer Server: {0}" -f $_.Exception.Message)
                    [void]$hits.Add([pscustomobject]@{ ServerHw = $sh; Location = '(Location unbekannt)'; ProfileName = '' })
                }
            }
            $result.Found = @($hits.ToArray())
            $result.Stage = 'done'
            Write-Log "[$Tag] FERTIG - $($hits.Count) Treffer"
        } catch {
            $result.Error = "[$($result.Stage)] $($_.Exception.Message)"
            Write-Log "[$Tag] ABBRUCH in Stage '$($result.Stage)': $($_.Exception.Message)"
        }

        return $result
    }

    $parallelResults = $allResults
    Write-SearchLog "--- Suche beendet ---"

    $totalHits = 0
    $errors = @()
    $perApplianceSummary = @()
    foreach ($res in $parallelResults) {
        try {
            if ($null -eq $res) {
                Write-SearchLog "Aggregator: NULL-Result uebersprungen"
                continue
            }
            $resIp  = [string]$res.Appliance
            $resTag = [string]$res.Tag
            if ([string]::IsNullOrEmpty($resTag)) { $resTag = '???' }
            if ($res.Error) {
                $errors += "$resIp`: $($res.Error)"
                Write-SearchLog "Aggregator [$resTag]: Worker-Error: $($res.Error)"
                continue
            }
            $foundCount = 0
            if ($null -ne $res.Found) { $foundCount = @($res.Found).Count }
            $perApplianceSummary += "$resIp`: $($res.TotalServers) Server gescannt, $foundCount Treffer"
            Write-SearchLog "Aggregator [$resTag]: $($res.TotalServers) gescannt, $foundCount Treffer"

            # Appliance-Label mit API-Version aktualisieren - defensiv per foreach
            # ($appliances | Where-Object IP -eq ...) kann bei deserialisierten
            # Strings aus Parallel-Runspaces "Argument Types do not match" werfen.
            $apIdx = $null
            foreach ($ap in $appliances) {
                if ([string]$ap.IP -eq $resIp) { $apIdx = [int]$ap.Index; break }
            }
            if ($null -ne $apIdx) {
                try {
                    $chkAppliances.Items[$apIdx] = "$resIp   (OV API $($res.VerLabel))"
                    $chkAppliances.SetItemChecked($apIdx, $true)
                } catch {
                    Write-SearchLog "Aggregator [$resTag]: Label-Update FEHLER: $($_.Exception.Message)"
                }
            }

            if ($foundCount -eq 0) { continue }

            foreach ($hit in @($res.Found)) {
                try {
                    if ($null -eq $hit) { continue }
                    $sh  = $hit.ServerHw
                    $loc = [string]$hit.Location
                    $hitProfileName = ''
                    if ($hit.PSObject.Properties.Name -contains 'ProfileName' -and $null -ne $hit.ProfileName) {
                        $hitProfileName = [string]$hit.ProfileName
                    }
                    if ($null -eq $sh) {
                        Write-SearchLog "Aggregator [$resTag]: Hit ohne ServerHw uebersprungen"
                        continue
                    }
                    $script:hitObjects += @{
                        Appliance  = $resIp
                        ApiVersion = $res.ApiVersion
                        Session    = $res.Session
                        VerLabel   = $res.VerLabel
                        ServerHw   = $sh
                        Location   = $loc
                        ProfileName= $hitProfileName
                    }
                    # Alle Werte als String erzwingen - DGV.Rows.Add verlangt
                    # konsistente Typen, sonst "Argument Types do not match".
                    $vName = [string](Get-Prop $sh @('name'))
                    $vSn   = [string](Get-Prop $sh @('serialNumber'))
                    $vMod  = [string](Get-Prop $sh @('model','shortModel'))
                    $vFf   = [string](Get-Prop $sh @('formFactor'))
                    $vSt   = [string](Get-Prop $sh @('status'))
                    $vPw   = [string](Get-Prop $sh @('powerState'))
                    [void]$dgvHits.Rows.Add($resIp, $vName, $hitProfileName, $vSn, $vMod, $vFf, $loc, $vSt, $vPw)
                    $totalHits++
                } catch {
                    Write-SearchLog "Aggregator [$resTag]: Hit-Add FEHLER: $($_.Exception.Message) | $($_.ScriptStackTrace)"
                    $errors += "$resIp`: Hit-Anzeige fehlgeschlagen: $($_.Exception.Message)"
                }
            }
        } catch {
            $ipForLog  = if ($res) { [string]$res.Appliance } else { '?' }
            $tagForLog = if ($res) { [string]$res.Tag } else { '???' }
            Write-SearchLog "Aggregator [$tagForLog]: AUSSERE FEHLER: $($_.Exception.Message) | $($_.ScriptStackTrace)"
            $errors += "$ipForLog`: $($_.Exception.Message)"
        }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    $btnSearch.Enabled = $true

    if ($errors.Count -gt 0) {
        # Fehler sind kritisch - sie wuerden sonst durch die "Kein Treffer"-Box uebermalt
        [System.Windows.Forms.MessageBox]::Show(
            ("Fehler bei einer oder mehreren Appliances:`r`n`r`n" + ($errors -join "`r`n")),
            "Such-Fehler", 'OK', 'Warning') | Out-Null
    }
    if ($perApplianceSummary.Count -gt 0) {
        $statusLabel.Text = ($perApplianceSummary -join '  |  ')
    }

    if ($totalHits -eq 0) {
        if ($errors.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                ("Kein Server fuer '$q' gefunden.`r`n`r`n" + ($perApplianceSummary -join "`r`n")),
                "Suche", 'OK', 'Information') | Out-Null
        }
    } elseif ($totalHits -eq 1) {
        $statusLabel.Text = "1 Treffer."
        $dgvHits.Rows[0].Selected = $true
        Show-Details 0
    } else {
        $statusLabel.Text = "$totalHits Treffer - bitte einen auswählen."
    }
}

function Show-Details {
    param([int]$rowIndex)
    if ($rowIndex -lt 0 -or $rowIndex -ge $script:hitObjects.Count) { return }
    # Re-Entrancy-Schutz: SelectionChanged feuert pro Klick mehrfach und MessageBoxen
    # pumpen die Message-Loop. Ohne Guard wird der schwere Detail-Load mehrfach
    # gleichzeitig angestossen -> GUI scheint zu haengen.
    if ($script:isLoadingDetails) { return }
    # Gleiche Zeile nicht erneut laden (vermeidet Doppel-Load bei 1 Treffer und
    # bei mehrfachem SelectionChanged fuer denselben Klick).
    if ($script:lastDetailRow -eq $rowIndex) { return }
    $script:isLoadingDetails = $true
    $script:lastDetailRow = $rowIndex
    $h = $script:hitObjects[$rowIndex]
    $A = $h.Appliance; $S = $h.Session; $V = $h.ApiVersion
    $sh = $h.ServerHw
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # Detail erneut frisch laden (mehr Felder als in der List-Version)
        $shFull = $sh
        try {
            $shFull = OV-Rest -A $A -S $S -V $V -M Get -E (Get-Prop $sh @('uri'))
        } catch { }

        $profName = ''
        if ($h -is [hashtable] -and $h.ContainsKey('ProfileName') -and $h.ProfileName) { $profName = [string]$h.ProfileName }
        $profUri = Get-Prop $shFull @('serverProfileUri')
        if ($profUri -and -not $profName) {
            try {
                $prof = OV-Rest -A $A -S $S -V $V -M Get -E $profUri
                $profName = Get-Prop $prof @('name')
            } catch { }
        }

        $txtOverview.Text = Build-Overview -A $A -verLabel $h.VerLabel -sh $shFull -location $h.Location -profileName $profName
        $txtCpu.Text = Build-CpuRam -A $A -S $S -V $V -sh $shFull
        Fill-FirmwareGrid -A $A -S $S -V $V -sh $shFull -grid $dgvFw
        # Vollstaendiges Profil fuer Adapter-Anreicherung holen (Connections, Netzwerke)
        $profObj = $null
        if ($profUri) { try { $profObj = OV-Rest -A $A -S $S -V $V -M Get -E $profUri } catch { } }
        Fill-NetGrid -sh $shFull -grid $dgvNet -detailBox $txtNet -A $A -S $S -V $V -prof $profObj
        $txtProfile.Text = Build-ProfileText -A $A -S $S -V $V -sh $shFull
        $txtStorage.Text = Build-Storage -A $A -S $S -V $V -sh $shFull
        $txtPower.Text   = Build-Power   -A $A -S $S -V $V -sh $shFull
        $txtGpu.Text     = Build-Graphics -A $A -S $S -V $V -sh $shFull
        $txtBios.Text    = Build-Bios -A $A -S $S -V $V -sh $shFull -IloUser $txtIloUser.Text -IloPass $txtIloPass.Text
        $txtAll.Text     = Build-AllFields $shFull
        $txtRaw.Text = ($shFull | ConvertTo-Json -Depth 12)
        $script:currentReport = @{
            Appliance  = $A
            ApiVersion = $h.ApiVersion
            VerLabel   = $h.VerLabel
            Location   = $h.Location
            ProfileName= $profName
            ServerHw   = $shFull
        }
        $btnExportTxt.Enabled = $true
        $btnExportHtml.Enabled = $true
        $statusLabel.Text = "Details geladen: $(Get-Prop $shFull @('name')) ($A)"
    } catch {
        $statusLabel.Text = "Fehler beim Laden der Details: $($_.Exception.Message)"
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $script:isLoadingDetails = $false
    }
}

# =============================
# Events
# =============================
$btnSearch.Add_Click({ Run-Search })
$txtSearch.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        Run-Search
    }
})

$dgvHits.Add_SelectionChanged({
    if ($dgvHits.SelectedRows.Count -gt 0) {
        Show-Details $dgvHits.SelectedRows[0].Index
    }
})

# =============================
# HTML-Report-Generator
# =============================
function HtmlEnc { param($v) if ($null -eq $v) { return "" } return [System.Net.WebUtility]::HtmlEncode([string]$v) }

function ConvertTo-KeyValueRows {
    param([string]$Text)
    $rows = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Text)) { return $rows }
    $currentSection = $null
    foreach ($rawLine in ($Text -split "`r?`n")) {
        $line = $rawLine.TrimEnd()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line -match '^\s*=+\s*(.+?)\s*=+\s*$') {
            $currentSection = $matches[1]; continue
        }
        $idx = $line.IndexOf(':')
        if ($idx -gt 0) {
            $k = $line.Substring(0, $idx).Trim()
            $v = $line.Substring($idx + 1).Trim()
            $rows.Add([pscustomobject]@{ Section = $currentSection; Key = $k; Value = $v })
        } else {
            $rows.Add([pscustomobject]@{ Section = $currentSection; Key = ''; Value = $line })
        }
    }
    return ,$rows
}

function Build-KvTableHtml {
    param([string]$Title, [string]$SourceText)
    $rows = ConvertTo-KeyValueRows -Text $SourceText
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("<section><h2>$(HtmlEnc $Title)</h2>")
    $lastSection = [object]'__init__'
    $tableOpen = $false
    foreach ($r in $rows) {
        if ($r.Section -ne $lastSection) {
            if ($tableOpen) { [void]$sb.AppendLine("</tbody></table>"); $tableOpen = $false }
            if ($r.Section) { [void]$sb.AppendLine("<h3>$(HtmlEnc $r.Section)</h3>") }
            $lastSection = $r.Section
        }
        if (-not $tableOpen) {
            [void]$sb.AppendLine("<table class='kv'><tbody>")
            $tableOpen = $true
        }
        if ($r.Key) {
            [void]$sb.AppendLine("<tr><th>$(HtmlEnc $r.Key)</th><td>$(HtmlEnc $r.Value)</td></tr>")
        } else {
            [void]$sb.AppendLine("<tr><td colspan='2'>$(HtmlEnc $r.Value)</td></tr>")
        }
    }
    if ($tableOpen) { [void]$sb.AppendLine("</tbody></table>") }
    [void]$sb.AppendLine("</section>")
    return $sb.ToString()
}

function Build-DgvTableHtml {
    param([string]$Title, [System.Windows.Forms.DataGridView]$Grid, [string]$EmptyText = "(keine Daten)")
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("<section><h2>$(HtmlEnc $Title)</h2>")
    $hasRows = $false
    foreach ($r in $Grid.Rows) { if (-not $r.IsNewRow) { $hasRows = $true; break } }
    if (-not $hasRows) {
        [void]$sb.AppendLine("<p class='muted'>$(HtmlEnc $EmptyText)</p></section>")
        return $sb.ToString()
    }
    [void]$sb.AppendLine("<table class='grid'><thead><tr>")
    foreach ($c in $Grid.Columns) {
        [void]$sb.AppendLine("<th>$(HtmlEnc $c.HeaderText)</th>")
    }
    [void]$sb.AppendLine("</tr></thead><tbody>")
    foreach ($r in $Grid.Rows) {
        if ($r.IsNewRow) { continue }
        [void]$sb.Append("<tr>")
        foreach ($c in $Grid.Columns) {
            $val = $r.Cells[$c.Name].Value
            [void]$sb.Append("<td>$(HtmlEnc $val)</td>")
        }
        [void]$sb.AppendLine("</tr>")
    }
    [void]$sb.AppendLine("</tbody></table></section>")
    return $sb.ToString()
}

function Build-PreSectionHtml {
    param([string]$Title, [string]$Body)
    if ([string]::IsNullOrWhiteSpace($Body)) {
        return "<section><h2>$(HtmlEnc $Title)</h2><p class='muted'>(keine Daten)</p></section>"
    }
    return "<section><h2>$(HtmlEnc $Title)</h2><pre>$(HtmlEnc $Body)</pre></section>"
}

function Build-HtmlReport {
    if (-not $script:currentReport) { return $null }
    $rep = $script:currentReport
    $sh  = $rep.ServerHw
    $name   = (Get-Prop $sh @('name'))
    $serial = (Get-Prop $sh @('serialNumber'))
    $model  = (Get-Prop $sh @('model','shortModel'))
    $power  = (Get-Prop $sh @('powerState'))
    $status = (Get-Prop $sh @('status'))
    $generated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

    $css = @"
:root { color-scheme: light dark; }
* { box-sizing: border-box; }
body { font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
       margin: 0; padding: 0; background: #f4f6f8; color: #1f2933; }
header.page { background: linear-gradient(135deg,#0b3d91,#1565c0); color: #fff;
              padding: 24px 32px; box-shadow: 0 2px 6px rgba(0,0,0,.15); }
header.page h1 { margin: 0 0 6px 0; font-size: 1.6rem; font-weight: 600; }
header.page .meta { font-size: .9rem; opacity: .9; }
header.page .badges { margin-top: 12px; }
.badge { display: inline-block; padding: 3px 10px; border-radius: 999px;
         background: rgba(255,255,255,.18); font-size: .8rem; margin-right: 6px; }
main { padding: 24px 32px; max-width: 1400px; margin: 0 auto; }
section { background: #fff; border: 1px solid #e1e4e8; border-radius: 8px;
          padding: 18px 22px; margin-bottom: 18px; box-shadow: 0 1px 2px rgba(0,0,0,.04); }
section h2 { margin: 0 0 12px 0; font-size: 1.15rem; color: #0b3d91;
             border-bottom: 2px solid #e1e4e8; padding-bottom: 6px; }
section h3 { margin: 14px 0 6px 0; font-size: .95rem; color: #334155; }
table { width: 100%; border-collapse: collapse; font-size: .88rem; }
table.kv th { text-align: left; width: 220px; vertical-align: top;
              padding: 4px 12px 4px 0; color: #475569; font-weight: 600; }
table.kv td { padding: 4px 0; vertical-align: top; word-break: break-word; }
table.grid { border: 1px solid #e1e4e8; }
table.grid th { background: #eef2f7; text-align: left; padding: 6px 8px;
                border-bottom: 1px solid #e1e4e8; font-weight: 600; color: #334155; }
table.grid td { padding: 5px 8px; border-bottom: 1px solid #f1f3f5; }
table.grid tr:hover td { background: #fafbfc; }
pre { background: #0f172a; color: #e2e8f0; padding: 14px; border-radius: 6px;
      overflow-x: auto; font-size: .8rem; line-height: 1.4;
      font-family: ui-monospace, "SF Mono", Menlo, Consolas, monospace; }
.muted { color: #6b7280; font-style: italic; }
footer { text-align: center; color: #6b7280; font-size: .8rem;
         padding: 16px; border-top: 1px solid #e1e4e8; margin-top: 8px; }
@media print {
  body { background: #fff; }
  header.page { background: #0b3d91 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  section { break-inside: avoid; box-shadow: none; }
}
"@

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("<!doctype html>")
    [void]$sb.AppendLine("<html lang='de'><head><meta charset='utf-8'>")
    [void]$sb.AppendLine("<title>OneView ServerInfo - $(HtmlEnc $name)</title>")
    [void]$sb.AppendLine("<style>$css</style></head><body>")
    [void]$sb.AppendLine("<header class='page'>")
    [void]$sb.AppendLine("<h1>$(HtmlEnc $name)</h1>")
    [void]$sb.AppendLine("<div class='meta'>$(HtmlEnc $model) &middot; SN $(HtmlEnc $serial) &middot; Appliance $(HtmlEnc $rep.Appliance) &middot; OneView $(HtmlEnc $rep.VerLabel)</div>")
    [void]$sb.AppendLine("<div class='meta'>Standort: $(HtmlEnc $rep.Location)</div>")
    [void]$sb.AppendLine("<div class='badges'>")
    if ($power)  { [void]$sb.AppendLine("<span class='badge'>Power: $(HtmlEnc $power)</span>") }
    if ($status) { [void]$sb.AppendLine("<span class='badge'>Status: $(HtmlEnc $status)</span>") }
    if ($rep.ProfileName) { [void]$sb.AppendLine("<span class='badge'>Profil: $(HtmlEnc $rep.ProfileName)</span>") }
    [void]$sb.AppendLine("</div></header><main>")

    [void]$sb.AppendLine((Build-KvTableHtml -Title "Übersicht" -SourceText $txtOverview.Text))
    [void]$sb.AppendLine((Build-PreSectionHtml -Title "CPU & RAM (Detail)" -Body $txtCpu.Text))
    [void]$sb.AppendLine((Build-DgvTableHtml -Title "Firmware-Inventory" -Grid $dgvFw -EmptyText "Keine Firmware-Komponenten gemeldet."))
    [void]$sb.AppendLine((Build-DgvTableHtml -Title "Adapter / Ports" -Grid $dgvNet -EmptyText "Keine Adapter-/Port-Informationen gemeldet."))
    [void]$sb.AppendLine((Build-PreSectionHtml -Title "Adapter / Ports - Details" -Body $txtNet.Text))
    [void]$sb.AppendLine((Build-PreSectionHtml -Title "Storage / Laufwerke" -Body $txtStorage.Text))
    [void]$sb.AppendLine((Build-PreSectionHtml -Title "Power / Thermal" -Body $txtPower.Text))
    [void]$sb.AppendLine((Build-PreSectionHtml -Title "GPU / Grafik" -Body $txtGpu.Text))
    [void]$sb.AppendLine((Build-PreSectionHtml -Title "BIOS Settings" -Body $txtBios.Text))
    [void]$sb.AppendLine((Build-PreSectionHtml -Title "Server-Profil" -Body $txtProfile.Text))
    [void]$sb.AppendLine((Build-PreSectionHtml -Title "Alle Felder (Rekursiver Dump)" -Body $txtAll.Text))

    [void]$sb.AppendLine("<footer>Erzeugt am $(HtmlEnc $generated) durch OneView-ServerInfo-GUI.ps1</footer>")
    [void]$sb.AppendLine("</main></body></html>")
    return $sb.ToString()
}

$btnExportTxt.Add_Click({
    if ([string]::IsNullOrEmpty($txtOverview.Text)) { return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "Textdatei (*.txt)|*.txt|Alle (*.*)|*.*"
    $defName = "ServerInfo_$((Get-Date).ToString('yyyyMMdd_HHmmss')).txt"
    $sfd.FileName = $defName
    if ($sfd.ShowDialog() -eq 'OK') {
        try {
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine($txtOverview.Text)
            [void]$sb.AppendLine()
            [void]$sb.AppendLine($txtCpu.Text)
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("=== Firmware-Inventory ===")
            foreach ($r in $dgvFw.Rows) {
                if ($r.IsNewRow) { continue }
                [void]$sb.AppendLine(("{0,-40} {1,-20} {2,-20} {3}" -f `
                    $r.Cells['componentName'].Value,
                    $r.Cells['componentLocation'].Value,
                    $r.Cells['componentVersion'].Value,
                    $r.Cells['componentKey'].Value))
            }
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("=== Adapter / Ports ===")
            foreach ($r in $dgvNet.Rows) {
                if ($r.IsNewRow) { continue }
                [void]$sb.AppendLine((("Slot {0,-4} {1,-22} {2,-20} FW {3,-12} Port {4,-4} {5,-8} " + `
                    "{6,-10}/{7,-10} MAC {8,-19} WWN {9,-25} {10}") -f `
                    $r.Cells['slot'].Value,
                    $r.Cells['adapter'].Value,
                    $r.Cells['model'].Value,
                    $r.Cells['fw'].Value,
                    $r.Cells['port'].Value,
                    $r.Cells['type'].Value,
                    $r.Cells['speedCur'].Value,
                    $r.Cells['speedMax'].Value,
                    $r.Cells['mac'].Value,
                    $r.Cells['wwpn'].Value,
                    $r.Cells['status'].Value))
            }
            [void]$sb.AppendLine()
            [void]$sb.AppendLine($txtNet.Text)
            [void]$sb.AppendLine()
            [void]$sb.AppendLine($txtStorage.Text)
            [void]$sb.AppendLine()
            [void]$sb.AppendLine($txtPower.Text)
            [void]$sb.AppendLine()
            [void]$sb.AppendLine($txtGpu.Text)
            [void]$sb.AppendLine()
            [void]$sb.AppendLine($txtBios.Text)
            [void]$sb.AppendLine()
            [void]$sb.AppendLine($txtProfile.Text)
            [void]$sb.AppendLine()
            [void]$sb.AppendLine($txtAll.Text)
            [System.IO.File]::WriteAllText($sfd.FileName, $sb.ToString(), [System.Text.Encoding]::UTF8)
            $statusLabel.Text = "Bericht gespeichert: $($sfd.FileName)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern: $($_.Exception.Message)", "Fehler",
                'OK', 'Error') | Out-Null
        }
    }
})

$btnExportHtml.Add_Click({
    if ([string]::IsNullOrEmpty($txtOverview.Text)) { return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "HTML-Datei (*.html)|*.html|Alle (*.*)|*.*"
    $svrName = if ($script:currentReport) { (Get-Prop $script:currentReport.ServerHw @('name')) } else { "Server" }
    $safeName = ($svrName -replace '[^A-Za-z0-9_.-]', '_')
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "Server" }
    $sfd.FileName = "ServerInfo_${safeName}_$((Get-Date).ToString('yyyyMMdd_HHmmss')).html"
    if ($sfd.ShowDialog() -eq 'OK') {
        try {
            $html = Build-HtmlReport
            if (-not $html) { throw "Keine Daten geladen." }
            [System.IO.File]::WriteAllText($sfd.FileName, $html, (New-Object System.Text.UTF8Encoding($true)))
            $statusLabel.Text = "HTML-Bericht gespeichert: $($sfd.FileName)"
            try { Start-Process $sfd.FileName | Out-Null } catch {}
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern: $($_.Exception.Message)", "Fehler",
                'OK', 'Error') | Out-Null
        }
    }
})

# Sessions beim Schließen sauber abmelden
$form.Add_FormClosing({
    foreach ($h in $script:hitObjects) {
        try { OV-Logout -A $h.Appliance -S $h.Session -V $h.ApiVersion } catch {}
    }
})

# Appliance-Liste beim Start automatisch laden
$form.Add_Shown({ Load-Appliances })

# Formular anzeigen
$form.ShowDialog() | Out-Null
