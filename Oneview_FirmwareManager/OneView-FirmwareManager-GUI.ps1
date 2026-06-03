#Requires -Version 7.0
# ============================================================================
#  HPE OneView Firmware Manager GUI
#  Direkte REST-API – keine HPE PowerShell-Module erforderlich
#  X-API-Version wird automatisch pro Appliance ermittelt
#  Unterstützt OV 6.60 + OV 11.10 (und beliebige andere Versionen)
#
#  Funktionen:
#   - Firmware Bundle (SPP-ISO / Hotfix .fwpkg / .exe) hochladen (multi-appliance)
#   - Firmware Bundles je Appliance auflisten
#   - Firmware Bundles löschen (multi-appliance)
#
#  REST-Endpunkte (laut HPE OneView API-Doku):
#   POST   /rest/firmware-bundles         (Upload SPP/Hotfix, multipart)
#   POST   /rest/firmware-bundles/resumable (große Dateien, > 2 GB ISO)
#   GET    /rest/firmware-drivers         (Liste vorhandener Bundles)
#   DELETE /rest/firmware-drivers/{name},{version}
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
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();
    }
}
"@
}
# DPI-Awareness aktivieren (vor Form-Erstellung), damit High-DPI-Notebooks das Layout korrekt skalieren
try { [Win32.NativeMethods]::SetProcessDPIAware() | Out-Null } catch {}
$consolePtr = [Win32.NativeMethods]::GetConsoleWindow()
if ($consolePtr -ne [System.IntPtr]::Zero) {
    [Win32.NativeMethods]::ShowWindow($consolePtr, 0)
}

# =============================
# Assemblies
# =============================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net.Http

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

# Firmware Bundle Upload via Multipart (HttpClient – unterstützt große Dateien)
# Verwendet -Resumable wenn Datei groß ist (>= 1 GB), sonst direkter POST
function OV-UploadFirmware {
    param(
        [string]$A,        # Appliance Hostname/IP
        [string]$S,        # Session-ID
        [int]$V,           # API-Version
        [string]$FilePath, # Lokaler Pfad zur ISO/fwpkg/exe
        [scriptblock]$ProgressCb = $null   # optional: param($percent,$bytesSent,$total)
    )
    if (-not (Test-Path -LiteralPath $FilePath)) { throw "Datei nicht gefunden: $FilePath" }
    $fi = Get-Item -LiteralPath $FilePath
    $fileName = $fi.Name
    $totalSize = $fi.Length

    # OneView akzeptiert auch große Bundles via /rest/firmware-bundles als Multipart-POST.
    # Der /resumable-Endpoint hat ein eigenes Chunk-Protokoll (Session-ID + PUT pro Offset)
    # und ist NICHT mit einem normalen multipart/form-data kompatibel - daher hier nicht verwenden.
    $endpoint = "/rest/firmware-bundles"
    $uri = "https://$A$endpoint"

    # HttpClient mit komplettem Zertifikats-Bypass (analog zu -SkipCertificateCheck der anderen Skripte).
    # DangerousAcceptAnyServerCertificateValidator ist die offizielle .NET-Variante, die garantiert greift -
    # ein Scriptblock-Callback wird unter PowerShell 7 manchmal nicht korrekt zugewiesen.
    $handler = New-Object System.Net.Http.HttpClientHandler
    try {
        $handler.ServerCertificateCustomValidationCallback = [System.Net.Http.HttpClientHandler]::DangerousAcceptAnyServerCertificateValidator
    } catch {
        $handler.ServerCertificateCustomValidationCallback = { param($a,$b,$c,$d) $true }
    }
    try { $handler.CheckCertificateRevocationList = $false } catch {}
    try { $handler.SslProtocols = [System.Security.Authentication.SslProtocols]::Tls12 } catch {}
    # Globaler Bypass zusätzlich (für interne Aufrufe über ServicePointManager)
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { param($s,$c,$ch,$e) $true }
    try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {}
    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromHours(6)
    $client.DefaultRequestHeaders.Add("Auth", $S)
    $client.DefaultRequestHeaders.Add("X-API-Version", "$V")
    $client.DefaultRequestHeaders.Add("uploadfilename", $fileName)
    # Accept JSON
    $client.DefaultRequestHeaders.Accept.Add(
        (New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json")))

    $fs = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        $streamContent = New-Object System.Net.Http.StreamContent($fs, 1MB)
        $streamContent.Headers.ContentType = New-Object System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream")
        # Wichtig: Name/FileName OHNE manuell gesetzte Anfuehrungszeichen - .NET quotet bereits selbst.
        # Doppelt-quoten erzeugt name=""file"" -> OneView meldet "Invalid parameter".
        $cd = New-Object System.Net.Http.Headers.ContentDispositionHeaderValue("form-data")
        $cd.Name = "file"
        $cd.FileName = $fileName
        $streamContent.Headers.ContentDisposition = $cd

        $multi = New-Object System.Net.Http.MultipartFormDataContent
        $multi.Add($streamContent)

        $task = $client.PostAsync($uri, $multi)
        # Optional: Fortschritt via Polling der Stream-Position
        if ($ProgressCb) {
            while (-not $task.IsCompleted) {
                Start-Sleep -Milliseconds 750
                try {
                    $sent = $fs.Position
                    if ($totalSize -gt 0) {
                        $pct = [int](($sent / $totalSize) * 100)
                        & $ProgressCb $pct $sent $totalSize
                    }
                } catch {}
            }
        }
        $resp = $task.GetAwaiter().GetResult()
        $body = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        if (-not $resp.IsSuccessStatusCode) {
            throw "Upload fehlgeschlagen ($([int]$resp.StatusCode) $($resp.ReasonPhrase)): $body"
        }
        if ($ProgressCb) { & $ProgressCb 100 $totalSize $totalSize }
        # Task-URI aus Location-Header extrahieren (für asynchrone Bundle-Imports)
        $taskUri = $null
        if ($resp.Headers.Location) { $taskUri = $resp.Headers.Location.ToString() }
        $parsed = $null
        try { $parsed = $body | ConvertFrom-Json } catch { $parsed = $null }
        return [PSCustomObject]@{
            TaskUri  = $taskUri
            Response = $parsed
            RawBody  = $body
        }
    }
    finally {
        $fs.Dispose()
        $client.Dispose()
        $handler.Dispose()
    }
}

# Wartet bis Task abgeschlossen / Fehler – gibt Task-Objekt zurück
function OV-WaitForTask {
    param(
        [string]$A, [string]$S, [int]$V, [string]$TaskUri,
        [int]$TimeoutSec = 7200,
        [scriptblock]$ProgressCb = $null
    )
    if ([string]::IsNullOrWhiteSpace($TaskUri)) { return $null }
    if ($TaskUri -notmatch '^/') { $TaskUri = "/$TaskUri" }
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $lastPct = -1
    while ((Get-Date) -lt $deadline) {
        try {
            $t = OV-Rest -A $A -S $S -V $V -M Get -E $TaskUri
            if ($ProgressCb -and $t.percentComplete -ne $null -and $t.percentComplete -ne $lastPct) {
                $lastPct = [int]$t.percentComplete
                & $ProgressCb $lastPct $t.taskState
            }
            switch -Regex ($t.taskState) {
                '^(Completed|Warning)$' { return $t }
                '^(Error|Killed|Terminated)$' {
                    $errMsg = ($t.taskErrors | ForEach-Object { $_.message }) -join '; '
                    if (-not $errMsg) { $errMsg = $t.taskStatus }
                    throw "Task-Fehler ($($t.taskState)): $errMsg"
                }
                default { Start-Sleep -Seconds 5 }
            }
        }
        catch {
            if ($_.Exception.Message -like 'Task-Fehler*') { throw }
            Start-Sleep -Seconds 5
        }
    }
    throw "Task-Timeout nach $TimeoutSec s ($TaskUri)"
}

# Final-Check: Bundle muss in /rest/firmware-drivers vorhanden + resourceState=Created sein
function OV-VerifyBundle {
    param(
        [string]$A, [string]$S, [int]$V,
        [string]$FileName,   # z.B. SPP2024050.2024_0510.10.iso
        [string]$BundleName, # optional: aus Task-Response
        [string]$BundleVersion
    )
    # Filter über Filename oder Name
    $resp = OV-Rest -A $A -S $S -V $V -M Get -E "/rest/firmware-drivers?start=0&count=2000"
    $list = if ($resp.members) { @($resp.members) } else { @() }
    $match = $null
    if ($BundleName) {
        $match = $list | Where-Object { $_.name -eq $BundleName -and (-not $BundleVersion -or $_.version -eq $BundleVersion) } | Select-Object -First 1
    }
    if (-not $match -and $FileName) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
        $match = $list | Where-Object { $_.name -eq $base -or $_.fileName -eq $FileName -or $_.resourceFilename -eq $FileName } | Select-Object -First 1
    }
    if (-not $match) {
        # Letzter Fallback: neuestes Bundle (höchste modified-Zeit)
        $match = $list | Sort-Object modified -Descending | Select-Object -First 1
    }
    if (-not $match) { throw "Bundle nach Upload nicht im Repository gefunden" }

    $state = "$($match.resourceState)"
    if ($state -notmatch '^(Created|AddingPackage)$') {
        throw "Bundle hat unerwarteten resourceState='$state' (Name: $($match.name))"
    }
    return $match
}
'@

# =============================
# Haupt-Formular
# =============================
$form = New-Object System.Windows.Forms.Form
$null = $form.Handle
$form.Text = "© 2025 N.J. Airbus D&S - HPE OneView Firmware Manager"
$form.Size = New-Object System.Drawing.Size(1100,1100)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI",9)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96,96)

$boldFont = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)

# ─────────────────────────────────────────
# Oberer Bereich: Admin-Login
# ─────────────────────────────────────────
$lblAdminUser = New-Object System.Windows.Forms.Label
$lblAdminUser.AutoSize = $true; $lblAdminUser.Location = '10,15'; $lblAdminUser.Text = "Admin Login:"; $lblAdminUser.Font = $boldFont
$form.Controls.Add($lblAdminUser)

$txtAdminUser = New-Object System.Windows.Forms.TextBox
$txtAdminUser.Location = '130,13'; $txtAdminUser.Size = '170,22'; $txtAdminUser.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtAdminUser)

$lblAdminPass = New-Object System.Windows.Forms.Label
$lblAdminPass.AutoSize = $true; $lblAdminPass.Location = '315,15'; $lblAdminPass.Text = "Passwort:"
$form.Controls.Add($lblAdminPass)

$txtAdminPass = New-Object System.Windows.Forms.TextBox
$txtAdminPass.Location = '390,13'; $txtAdminPass.Size = '180,22'; $txtAdminPass.UseSystemPasswordChar = $true; $txtAdminPass.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtAdminPass)

# ─────────────────────────────────────────
# IP-Dateien
# ─────────────────────────────────────────
$lblIP660 = New-Object System.Windows.Forms.Label
$lblIP660.AutoSize = $true; $lblIP660.Location = '10,47'; $lblIP660.Text = "OV 6.60 IP-Datei:"; $lblIP660.Font = $boldFont
$form.Controls.Add($lblIP660)

$txtIP660 = New-Object System.Windows.Forms.TextBox
$txtIP660.Location = '160,45'; $txtIP660.Size = '720,22'; $txtIP660.Text = (Join-Path (Split-Path $scriptFolder -Parent) "Oneview_660.txt"); $txtIP660.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtIP660)

$btnBrowse660 = New-Object System.Windows.Forms.Button
$btnBrowse660.Location = '890,44'; $btnBrowse660.Size = '90,24'; $btnBrowse660.Text = "Browse..."
$form.Controls.Add($btnBrowse660)

$lblIP1110 = New-Object System.Windows.Forms.Label
$lblIP1110.AutoSize = $true; $lblIP1110.Location = '10,75'; $lblIP1110.Text = "OV 11.10 IP-Datei:"; $lblIP1110.Font = $boldFont
$form.Controls.Add($lblIP1110)

$txtIP1110 = New-Object System.Windows.Forms.TextBox
$txtIP1110.Location = '160,73'; $txtIP1110.Size = '720,22'; $txtIP1110.Text = (Join-Path (Split-Path $scriptFolder -Parent) "Oneview.txt"); $txtIP1110.BorderStyle = 'FixedSingle'
$form.Controls.Add($txtIP1110)

$btnBrowse1110 = New-Object System.Windows.Forms.Button
$btnBrowse1110.Location = '890,72'; $btnBrowse1110.Size = '90,24'; $btnBrowse1110.Text = "Browse..."
$form.Controls.Add($btnBrowse1110)

# ─────────────────────────────────────────
# Appliance-Auswahl (CheckedListBox)
# ─────────────────────────────────────────
$lblApplSel = New-Object System.Windows.Forms.Label
$lblApplSel.AutoSize = $true; $lblApplSel.Location = '10,105'; $lblApplSel.Text = "Appliance-Auswahl:"; $lblApplSel.Font = $boldFont
$form.Controls.Add($lblApplSel)

$btnSelAll = New-Object System.Windows.Forms.Button
$btnSelAll.Location = '180,102'; $btnSelAll.Size = '70,24'; $btnSelAll.Text = "Alle"
$form.Controls.Add($btnSelAll)

$btnSelNone = New-Object System.Windows.Forms.Button
$btnSelNone.Location = '255,102'; $btnSelNone.Size = '70,24'; $btnSelNone.Text = "Keine"
$form.Controls.Add($btnSelNone)

$chkAppliances = New-Object System.Windows.Forms.CheckedListBox
$chkAppliances.Location = '10,130'; $chkAppliances.Size = '1060,195'; $chkAppliances.CheckOnClick = $true; $chkAppliances.BorderStyle = 'FixedSingle'
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

# Hilfsfunktion: SHA-256 aus einer Checksummendatei extrahieren.
# Unterstützt:
#   - HPE-Format mit Schluesselzeilen ("SHA-256 : <hash>", "SHA256:<hash>", ...)
#   - Standard-Format ("<hash>  filename" oder nur "<hash>")
function Read-HpeChecksumFile {
    param([Parameter(Mandatory)] [string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        # 1) "SHA-256 : <hash>" / "SHA256: <hash>" / "sha256 = <hash>"
        if ($line -match '(?i)\bsha[-_]?256\b\s*[:=]\s*([0-9a-f]{64})\b') {
            return $Matches[1].ToLower()
        }
        # 2) Standardformat: "<hash>  filename" oder nur "<hash>"
        if ($line -match '^\s*([0-9a-fA-F]{64})\b') {
            return $Matches[1].ToLower()
        }
    }
    return $null
}

# Hilfsfunktion: ComboBoxen in Tabs aktualisieren
function Update-ApplianceComboBoxes {
    $items = @()
    for ($i = 0; $i -lt $chkAppliances.Items.Count; $i++) {
        if ($chkAppliances.GetItemChecked($i)) {
            $items += $chkAppliances.Items[$i].ToString()
        }
    }
    foreach ($cb in @($cboT2Appl)) {
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
$tabControl.Location = '10,333'; $tabControl.Size = '1060,465'
$form.Controls.Add($tabControl)

# ═══════════════════════════════════════════
# TAB 1: Firmware Bundle hochladen (Multi-Appliance)
# ═══════════════════════════════════════════
$tab1 = New-Object System.Windows.Forms.TabPage; $tab1.Text = "Firmware hochladen"
$tabControl.TabPages.Add($tab1)

$lblT1File = New-Object System.Windows.Forms.Label
$lblT1File.Location = '10,15'; $lblT1File.AutoSize = $true; $lblT1File.Text = "Firmware-Datei:"; $lblT1File.Font = $boldFont
$tab1.Controls.Add($lblT1File)

$txtT1File = New-Object System.Windows.Forms.TextBox
$txtT1File.Location = '155,13'; $txtT1File.Size = '770,22'; $txtT1File.BorderStyle = 'FixedSingle'
$tab1.Controls.Add($txtT1File)

$btnT1Browse = New-Object System.Windows.Forms.Button
$btnT1Browse.Location = '935,12'; $btnT1Browse.Size = '100,24'; $btnT1Browse.Text = "Browse..."
$tab1.Controls.Add($btnT1Browse)

$lblT1Hint = New-Object System.Windows.Forms.Label
$lblT1Hint.Location = '155,40'; $lblT1Hint.AutoSize = $true
$lblT1Hint.Text = "Unterstützte Dateitypen: .iso (Service Pack), .fwpkg / .exe (Hotfix). Originalname/-endung NICHT ändern."
$lblT1Hint.ForeColor = [System.Drawing.Color]::Gray
$tab1.Controls.Add($lblT1Hint)

$lblT1Size = New-Object System.Windows.Forms.Label
$lblT1Size.Location = '155,60'; $lblT1Size.AutoSize = $true; $lblT1Size.Text = ""
$lblT1Size.ForeColor = [System.Drawing.Color]::DarkSlateGray
$tab1.Controls.Add($lblT1Size)

# Optionale SHA-256 Prüfung der lokalen Datei (vor Upload)
$lblT1Sha = New-Object System.Windows.Forms.Label
$lblT1Sha.Location = '10,85'; $lblT1Sha.AutoSize = $true; $lblT1Sha.Text = "SHA-256 (optional):"
$tab1.Controls.Add($lblT1Sha)

$txtT1Sha = New-Object System.Windows.Forms.TextBox
$txtT1Sha.Location = '155,83'; $txtT1Sha.Size = '655,22'; $txtT1Sha.BorderStyle = 'FixedSingle'
$txtT1Sha.PlaceholderText = "HPE-Hash zur Verifikation einfügen (leer = überspringen)"
$tab1.Controls.Add($txtT1Sha)

$btnT1LoadHash = New-Object System.Windows.Forms.Button
$btnT1LoadHash.Location = '815,82'; $btnT1LoadHash.Size = '115,24'; $btnT1LoadHash.Text = "Hash-Datei..."
$tab1.Controls.Add($btnT1LoadHash)

$btnT1Verify = New-Object System.Windows.Forms.Button
$btnT1Verify.Location = '935,82'; $btnT1Verify.Size = '100,24'; $btnT1Verify.Text = "Verifizieren"
$tab1.Controls.Add($btnT1Verify)

$lblT1ShaResult = New-Object System.Windows.Forms.Label
$lblT1ShaResult.Location = '155,108'; $lblT1ShaResult.AutoSize = $true; $lblT1ShaResult.Text = ""
$tab1.Controls.Add($lblT1ShaResult)

# Parallelität
$lblT1Par = New-Object System.Windows.Forms.Label
$lblT1Par.Location = '10,135'; $lblT1Par.AutoSize = $true; $lblT1Par.Text = "Parallel-Uploads:"
$tab1.Controls.Add($lblT1Par)

$numT1Par = New-Object System.Windows.Forms.NumericUpDown
$numT1Par.Location = '155,133'; $numT1Par.Size = '70,22'; $numT1Par.Minimum = 1; $numT1Par.Maximum = 50; $numT1Par.Value = 5
$tab1.Controls.Add($numT1Par)

$lblT1ParHint = New-Object System.Windows.Forms.Label
$lblT1ParHint.Location = '230,135'; $lblT1ParHint.AutoSize = $true
$lblT1ParHint.Text = "(empfohlen 5–10; höher belastet Netzwerk/Lokal-IO. Pro Appliance: Upload + Task-Polling + Final-Check)"
$lblT1ParHint.ForeColor = [System.Drawing.Color]::Gray
$tab1.Controls.Add($lblT1ParHint)

$btnT1Upload = New-Object System.Windows.Forms.Button
$btnT1Upload.Location = '10,162'; $btnT1Upload.Size = '340,28'
$btnT1Upload.Text = "Upload auf ALLE ausgewählten Appliances"
$btnT1Upload.Font = $boldFont
$tab1.Controls.Add($btnT1Upload)

$progressT1 = New-Object System.Windows.Forms.ProgressBar
$progressT1.Location = '360,165'; $progressT1.Size = '675,22'
$progressT1.Minimum = 0; $progressT1.Maximum = 100
$tab1.Controls.Add($progressT1)

$lvT1 = New-Object System.Windows.Forms.ListView
$lvT1.Location = '10,200'; $lvT1.Size = '1025,228'
$lvT1.View = 'Details'; $lvT1.FullRowSelect = $true; $lvT1.GridLines = $true; $lvT1.BorderStyle = 'FixedSingle'
$lvT1.Columns.Add("Appliance",170) | Out-Null
$lvT1.Columns.Add("Version",90) | Out-Null
$lvT1.Columns.Add("Phase",130) | Out-Null
$lvT1.Columns.Add("Fortschritt",110) | Out-Null
$lvT1.Columns.Add("Status",100) | Out-Null
$lvT1.Columns.Add("Details",405) | Out-Null
$tab1.Controls.Add($lvT1)

# SHA-256 Verifikation (lokal, vor Upload) – läuft im Hintergrund-Runspace
$btnT1Verify.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtT1File.Text) -or -not (Test-Path -LiteralPath $txtT1File.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte gültige Datei auswählen.", "Fehler", 'OK', 'Error'); return
    }
    $expected = $txtT1Sha.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($expected)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte erwarteten SHA-256 Hash eingeben (von HPE-Download-Seite).", "Hinweis", 'OK', 'Information'); return
    }
    $btnT1Verify.Enabled = $false
    $lblT1ShaResult.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $lblT1ShaResult.Text = "Berechne SHA-256... (kann bei mehreren GB einige Minuten dauern)"
    Start-AsyncOp -Block {
        param($filePath, $expected, $uiQueue)
        try {
            $h = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash
            $ok = ($h -ieq $expected.Trim())
            $uiQueue.Enqueue(@{ Type='SHA_RESULT'; Ok=$ok; Computed=$h; Expected=$expected })
        } catch {
            $uiQueue.Enqueue(@{ Type='SHA_RESULT'; Ok=$false; Computed=""; Expected=$expected; ErrorMsg=$_.Exception.Message })
        }
    } -Arguments @($txtT1File.Text, $expected, $script:uiQueue)
})

$btnT1Browse.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Firmware-Dateien (*.iso;*.fwpkg;*.exe;*.zip)|*.iso;*.fwpkg;*.exe;*.zip|Alle Dateien (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') {
        $txtT1File.Text = $ofd.FileName
        try {
            $fi = Get-Item -LiteralPath $ofd.FileName
            $mb = [math]::Round($fi.Length / 1MB, 1)
            $lblT1Size.Text = "Größe: $mb MB"
        } catch { $lblT1Size.Text = "" }
        # Auto-Detect: HPE-Checksummendatei neben der ISO suchen (gleicher Basename, .sha256sum / .sha256 / .sum)
        try {
            $dir  = Split-Path -LiteralPath $ofd.FileName -Parent
            $base = [System.IO.Path]::GetFileNameWithoutExtension($ofd.FileName)
            $candidates = @(
                (Join-Path $dir "$base.sha256sum"),
                (Join-Path $dir "$base.sha256"),
                (Join-Path $dir "$($ofd.FileName).sha256sum"),
                (Join-Path $dir "$($ofd.FileName).sha256")
            )
            $found = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
            if ($found) {
                $h = Read-HpeChecksumFile -Path $found
                if ($h) {
                    $txtT1Sha.Text = $h
                    $lblT1ShaResult.ForeColor = [System.Drawing.Color]::DarkSlateGray
                    $lblT1ShaResult.Text = "Hash automatisch geladen aus: $(Split-Path -Leaf $found)"
                }
            }
        } catch {}
    }
})

# Manuelles Laden einer Checksummendatei (HPE-Format oder Standard-Format)
$btnT1LoadHash.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Checksummendateien (*.sha256sum;*.sha256;*.sum;*.txt)|*.sha256sum;*.sha256;*.sum;*.txt|Alle Dateien (*.*)|*.*"
    if (-not [string]::IsNullOrWhiteSpace($txtT1File.Text)) {
        try { $ofd.InitialDirectory = Split-Path -LiteralPath $txtT1File.Text -Parent } catch {}
    }
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $h = Read-HpeChecksumFile -Path $ofd.FileName
            if ($h) {
                $txtT1Sha.Text = $h
                $lblT1ShaResult.ForeColor = [System.Drawing.Color]::DarkSlateGray
                $lblT1ShaResult.Text = "Hash geladen aus: $(Split-Path -Leaf $ofd.FileName)"
            } else {
                [System.Windows.Forms.MessageBox]::Show("In der Datei wurde kein gültiger SHA-256 Hash gefunden.", "Hinweis", 'OK', 'Warning') | Out-Null
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Lesen der Hash-Datei: $($_.Exception.Message)", "Fehler", 'OK', 'Error') | Out-Null
        }
    }
})

# ═══════════════════════════════════════════
# TAB 2: Firmware Bundles auflisten (Einzel-Appliance)
# ═══════════════════════════════════════════
$tab2 = New-Object System.Windows.Forms.TabPage; $tab2.Text = "Firmware auflisten"
$tabControl.TabPages.Add($tab2)

$lblT2Appl = New-Object System.Windows.Forms.Label
$lblT2Appl.Location = '10,12'; $lblT2Appl.AutoSize = $true; $lblT2Appl.Text = "Appliance:"
$tab2.Controls.Add($lblT2Appl)

$cboT2Appl = New-Object System.Windows.Forms.ComboBox
$cboT2Appl.Location = '95,10'; $cboT2Appl.Size = '420,23'; $cboT2Appl.DropDownStyle = 'DropDownList'
$tab2.Controls.Add($cboT2Appl)

$btnT2Load = New-Object System.Windows.Forms.Button
$btnT2Load.Location = '525,9'; $btnT2Load.Size = '160,25'; $btnT2Load.Text = "Bundles laden"
$tab2.Controls.Add($btnT2Load)

$btnT2Export = New-Object System.Windows.Forms.Button
$btnT2Export.Location = '695,9'; $btnT2Export.Size = '180,25'; $btnT2Export.Text = "Als CSV exportieren"
$tab2.Controls.Add($btnT2Export)

$dgvT2 = New-Object System.Windows.Forms.DataGridView
$dgvT2.Location = '10,42'; $dgvT2.Size = '1030,385'
$dgvT2.AllowUserToAddRows = $false; $dgvT2.AllowUserToDeleteRows = $false
$dgvT2.ReadOnly = $true; $dgvT2.SelectionMode = 'FullRowSelect'
$dgvT2.AutoSizeColumnsMode = 'Fill'; $dgvT2.RowHeadersVisible = $false
$dgvT2.BorderStyle = 'FixedSingle'
$dgvT2.Columns.Add("name","Name") | Out-Null
$dgvT2.Columns.Add("version","Version") | Out-Null
$dgvT2.Columns.Add("bundleType","Typ") | Out-Null
$dgvT2.Columns.Add("releaseDate","Release") | Out-Null
$dgvT2.Columns.Add("resourceState","Status") | Out-Null
$dgvT2.Columns.Add("uri","URI") | Out-Null
$dgvT2.Columns["name"].FillWeight = 30
$dgvT2.Columns["version"].FillWeight = 12
$dgvT2.Columns["bundleType"].FillWeight = 12
$dgvT2.Columns["releaseDate"].FillWeight = 12
$dgvT2.Columns["resourceState"].FillWeight = 10
$dgvT2.Columns["uri"].FillWeight = 24
$tab2.Controls.Add($dgvT2)

# ═══════════════════════════════════════════
# TAB 3: Firmware Bundle löschen (Multi-Appliance)
# ═══════════════════════════════════════════
$tab3 = New-Object System.Windows.Forms.TabPage; $tab3.Text = "Firmware löschen"
$tabControl.TabPages.Add($tab3)

$lblT3Hint = New-Object System.Windows.Forms.Label
$lblT3Hint.Location = '10,10'; $lblT3Hint.AutoSize = $true
$lblT3Hint.Text = "Bundles werden zuerst von der ersten ausgewählten Appliance geladen, dann auf ALLEN ausgewählten Appliances gelöscht."
$lblT3Hint.ForeColor = [System.Drawing.Color]::Gray
$tab3.Controls.Add($lblT3Hint)

$btnT3Refresh = New-Object System.Windows.Forms.Button
$btnT3Refresh.Location = '10,35'; $btnT3Refresh.Size = '240,26'; $btnT3Refresh.Text = "Bundles laden (1. Appliance)"
$tab3.Controls.Add($btnT3Refresh)

$btnT3Delete = New-Object System.Windows.Forms.Button
$btnT3Delete.Location = '260,35'; $btnT3Delete.Size = '360,26'
$btnT3Delete.Text = "Markierte Bundles auf ALLEN löschen"
$btnT3Delete.ForeColor = [System.Drawing.Color]::DarkRed
$btnT3Delete.Font = $boldFont
$tab3.Controls.Add($btnT3Delete)

$chkT3SelectAll = New-Object System.Windows.Forms.CheckBox
$chkT3SelectAll.Location = '630,38'; $chkT3SelectAll.AutoSize = $true; $chkT3SelectAll.Text = "Alle markieren / abwählen"
$tab3.Controls.Add($chkT3SelectAll)

$lvT3 = New-Object System.Windows.Forms.ListView
$lvT3.Location = '10,70'; $lvT3.Size = '1030,360'
$lvT3.View = 'Details'; $lvT3.FullRowSelect = $true; $lvT3.GridLines = $true
$lvT3.CheckBoxes = $true; $lvT3.BorderStyle = 'FixedSingle'
$lvT3.Columns.Add("Name",360) | Out-Null
$lvT3.Columns.Add("Version",130) | Out-Null
$lvT3.Columns.Add("Typ",120) | Out-Null
$lvT3.Columns.Add("Status",110) | Out-Null
$lvT3.Columns.Add("URI",290) | Out-Null
$tab3.Controls.Add($lvT3)

$chkT3SelectAll.Add_CheckedChanged({
    foreach ($it in $lvT3.Items) { $it.Checked = $chkT3SelectAll.Checked }
})

# ─────────────────────────────────────────
# Log-Bereich
# ─────────────────────────────────────────
$panelLog = New-Object System.Windows.Forms.Panel
$panelLog.Location = '10,805'; $panelLog.Size = '1060,150'; $panelLog.BorderStyle = 'FixedSingle'
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
$btnExit.Location = '10,965'; $btnExit.Size = '110,28'; $btnExit.Text = "Exit"
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
            'UPLOAD_INIT' {
                $lvT1.Items.Clear()
                foreach ($ip in $msg.Appliances) {
                    $li = New-Object System.Windows.Forms.ListViewItem($ip.IP)
                    $li.Name = $ip.IP
                    $li.SubItems.Add($ip.Version) | Out-Null
                    $li.SubItems.Add("Wartet...") | Out-Null   # Phase
                    $li.SubItems.Add("0%") | Out-Null            # Fortschritt
                    $li.SubItems.Add("-") | Out-Null             # Status
                    $li.SubItems.Add("") | Out-Null              # Details
                    $lvT1.Items.Add($li) | Out-Null
                }
                $progressT1.Value = 0
                $script:uploadDoneCount = 0
                $script:uploadTotalCount = $msg.Appliances.Count
            }
            'UPLOAD_PHASE' {
                $li = $lvT1.Items[$msg.Appliance]
                if ($li) { $li.SubItems[2].Text = $msg.Phase; $li.EnsureVisible() }
            }
            'UPLOAD_PROGRESS' {
                $li = $lvT1.Items[$msg.Appliance]
                if ($li) { $li.SubItems[3].Text = "$($msg.Percent)%" }
            }
            'UPLOAD_DONE' {
                $li = $lvT1.Items[$msg.Appliance]
                if ($li) {
                    $li.SubItems[2].Text = $msg.Phase
                    $li.SubItems[3].Text = if ($msg.Success) { "100%" } else { "-" }
                    $li.SubItems[4].Text = if ($msg.Success) { "OK" } else { "Fehler" }
                    $li.SubItems[5].Text = $msg.Detail
                    $li.ForeColor = if ($msg.Success) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkRed }
                    $li.EnsureVisible()
                }
                $script:uploadDoneCount++
                if ($script:uploadTotalCount -gt 0) {
                    $progressT1.Value = [Math]::Min(100, [int](($script:uploadDoneCount / $script:uploadTotalCount) * 100))
                }
                $statusLabel.Text = "Fertig: $($script:uploadDoneCount)/$($script:uploadTotalCount)"
                $logBox.AppendText("$($msg.Appliance): $($msg.Detail)`r`n"); $logBox.ScrollToCaret()
            }
            'SHA_RESULT' {
                $btnT1Verify.Enabled = $true
                if ($msg.Ok) {
                    $lblT1ShaResult.ForeColor = [System.Drawing.Color]::DarkGreen
                    $lblT1ShaResult.Text = "✓ SHA-256 OK – Datei stimmt mit erwartetem Hash überein."
                } elseif ($msg.ErrorMsg) {
                    $lblT1ShaResult.ForeColor = [System.Drawing.Color]::DarkRed
                    $lblT1ShaResult.Text = "Fehler: $($msg.ErrorMsg)"
                } else {
                    $lblT1ShaResult.ForeColor = [System.Drawing.Color]::DarkRed
                    $lblT1ShaResult.Text = "✗ SHA-256 Mismatch! Berechnet: $($msg.Computed)"
                }
            }
            'BUNDLE_LIST_T2' {
                $dgvT2.Rows.Clear()
                foreach ($b in $msg.Data) {
                    $dgvT2.Rows.Add($b.name, $b.version, $b.bundleType, $b.releaseDate, $b.resourceState, $b.uri) | Out-Null
                }
                $logBox.AppendText("$($msg.Data.Count) Firmware-Bundles geladen.`r`n"); $logBox.ScrollToCaret()
            }
            'BUNDLE_LIST_T3' {
                $lvT3.Items.Clear()
                foreach ($b in $msg.Data) {
                    $li = New-Object System.Windows.Forms.ListViewItem($b.name)
                    $li.SubItems.Add([string]$b.version) | Out-Null
                    $li.SubItems.Add([string]$b.bundleType) | Out-Null
                    $li.SubItems.Add([string]$b.resourceState) | Out-Null
                    $li.SubItems.Add([string]$b.uri) | Out-Null
                    $li.Tag = @{ name=$b.name; version=$b.version; uri=$b.uri }
                    $lvT3.Items.Add($li) | Out-Null
                }
                $logBox.AppendText("$($msg.Data.Count) Firmware-Bundles geladen (von $($msg.Source)).`r`n"); $logBox.ScrollToCaret()
            }
            'FINISHED' {
                $logBox.AppendText("Vorgang abgeschlossen.`r`n"); $logBox.ScrollToCaret()
                $statusLabel.Text = "Fertig"
                $btnT1Upload.Enabled = $true
                $btnT2Load.Enabled = $true
                $btnT3Refresh.Enabled = $true
                $btnT3Delete.Enabled = $true
                $progressT1.Value = 0
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
                $btnT1Upload.Enabled = $true; $btnT2Load.Enabled = $true
                $btnT3Refresh.Enabled = $true; $btnT3Delete.Enabled = $true
                $progressT1.Value = 0
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
    if ($Params) { $ps.AddArgument($Params) | Out-Null }
    elseif ($Arguments) { foreach ($a in $Arguments) { $ps.AddArgument($a) | Out-Null } }
    $null = $ps.BeginInvoke()
}

# ─────────────────────────────────────────
# ComboBox-Update bei Tab-Wechsel
# ─────────────────────────────────────────
$tabControl.Add_SelectedIndexChanged({ Update-ApplianceComboBoxes })
$chkAppliances.Add_ItemCheck({
    $form.BeginInvoke([Action]{ Update-ApplianceComboBoxes })
})

# ═══════════════════════════════════════════════════════════════════
#  EVENT-HANDLER: Tab 1 – Firmware hochladen
# ═══════════════════════════════════════════════════════════════════
$btnT1Upload.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ([string]::IsNullOrWhiteSpace($txtT1File.Text) -or -not (Test-Path -LiteralPath $txtT1File.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte gültige Firmware-Datei auswählen.", "Fehler", 'OK', 'Error'); return
    }
    $appliances = Get-CheckedAppliances
    if ($appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Appliances ausgewählt.", "Fehler", 'OK', 'Warning'); return
    }
    $fi = Get-Item -LiteralPath $txtT1File.Text
    $mb = [math]::Round($fi.Length / 1MB, 1)
    $expectedSha = $txtT1Sha.Text.Trim()
    $shaInfo = if ([string]::IsNullOrWhiteSpace($expectedSha)) { "OHNE SHA-256 Verifikation" } else { "MIT SHA-256 Verifikation" }
    $maxParallel = [int]$numT1Par.Value

    $res = [System.Windows.Forms.MessageBox]::Show(
        "Datei '$($fi.Name)' ($mb MB) auf $($appliances.Count) Appliance(s) parallel ($maxParallel gleichzeitig) hochladen?`n$shaInfo`n`nAchtung: Großer Upload kann lange dauern und belastet die Netzwerk-Anbindung deutlich!",
        "Bestätigung", 'YesNo', 'Question')
    if ($res -ne 'Yes') { return }

    # Optionale lokale SHA-256 Verifikation VOR dem Upload (einmalig)
    if (-not [string]::IsNullOrWhiteSpace($expectedSha)) {
        $logBox.AppendText("Berechne lokalen SHA-256 (kann mehrere Minuten dauern)...`r`n"); $logBox.ScrollToCaret()
        $statusLabel.Text = "SHA-256 wird berechnet..."
        try {
            $localHash = (Get-FileHash -LiteralPath $txtT1File.Text -Algorithm SHA256).Hash
        } catch {
            [System.Windows.Forms.MessageBox]::Show("SHA-256 Berechnung fehlgeschlagen: $($_.Exception.Message)", "Fehler", 'OK', 'Error'); return
        }
        if ($localHash -ine $expectedSha) {
            $msg = "SHA-256 stimmt NICHT überein!`nErwartet:  $expectedSha`nBerechnet: $localHash`n`nUpload abgebrochen."
            [System.Windows.Forms.MessageBox]::Show($msg, "SHA-256 Mismatch", 'OK', 'Error')
            $logBox.SelectionColor = [System.Drawing.Color]::Red
            $logBox.AppendText("$msg`r`n"); $logBox.ScrollToCaret()
            $logBox.SelectionColor = $logBox.ForeColor
            return
        }
        $logBox.AppendText("✓ SHA-256 OK ($localHash)`r`n"); $logBox.ScrollToCaret()
    }

    $btnT1Upload.Enabled = $false
    $logBox.AppendText("=== Upload '$($fi.Name)' auf $($appliances.Count) Appliances (max. $maxParallel parallel) gestartet ===`r`n"); $logBox.ScrollToCaret()
    $script:uiQueue.Enqueue(@{ Type='UPLOAD_INIT'; Appliances=$appliances })

    # Worker-Scriptblock pro Appliance (läuft parallel im Runspace-Pool)
    $worker = {
        param($p)
        $restCode=$p.restCode; $ip=$p.ip; $ver=$p.ver; $adminUser=$p.adminUser; $adminPass=$p.adminPass
        $filePath=$p.filePath; $uiQueue=$p.uiQueue
        Invoke-Expression $restCode
        $sess = $null; $apiV = $null
        try {
            $apiV = OV-GetApiVersion -A $ip
            $sess = OV-Login -A $ip -U $adminUser -P $adminPass -V $apiV

            # Phase 1: Upload mit Fortschritt
            $uiQueue.Enqueue(@{ Type='UPLOAD_PHASE'; Appliance=$ip; Phase='Upload' })
            $cb = {
                param($pct, $sent, $tot)
                $uiQueue.Enqueue(@{ Type='UPLOAD_PROGRESS'; Appliance=$ip; Percent=$pct })
            }.GetNewClosure()
            $up = OV-UploadFirmware -A $ip -S $sess -V $apiV -FilePath $filePath -ProgressCb $cb

            # Phase 2: Task-Polling (falls Task-URI zurückgegeben)
            $taskUri = $up.TaskUri
            if (-not $taskUri -and $up.Response -and $up.Response.taskUri) { $taskUri = $up.Response.taskUri }
            if (-not $taskUri -and $up.Response -and $up.Response.uri -and ($up.Response.uri -like '/rest/tasks/*')) { $taskUri = $up.Response.uri }

            if ($taskUri) {
                $uiQueue.Enqueue(@{ Type='UPLOAD_PHASE'; Appliance=$ip; Phase='Task wartet' })
                $taskCb = {
                    param($pct, $state)
                    $uiQueue.Enqueue(@{ Type='UPLOAD_PROGRESS'; Appliance=$ip; Percent=$pct })
                }.GetNewClosure()
                $taskRes = OV-WaitForTask -A $ip -S $sess -V $apiV -TaskUri $taskUri -TimeoutSec 7200 -ProgressCb $taskCb
            }

            # Phase 3: Final-Verifikation im Repository
            $uiQueue.Enqueue(@{ Type='UPLOAD_PHASE'; Appliance=$ip; Phase='Verifizieren' })
            $bundleName = $null; $bundleVer = $null
            if ($up.Response) {
                if ($up.Response.name)    { $bundleName = $up.Response.name }
                if ($up.Response.version) { $bundleVer  = $up.Response.version }
            }
            $fileName = [System.IO.Path]::GetFileName($filePath)
            $verified = OV-VerifyBundle -A $ip -S $sess -V $apiV -FileName $fileName -BundleName $bundleName -BundleVersion $bundleVer

            $detail = "OK: $($verified.name) $($verified.version) [$($verified.resourceState)]"
            $uiQueue.Enqueue(@{ Type='UPLOAD_DONE'; Appliance=$ip; Success=$true; Phase='Fertig'; Detail=$detail })
        }
        catch {
            $uiQueue.Enqueue(@{ Type='UPLOAD_DONE'; Appliance=$ip; Success=$false; Phase='Fehler'; Detail=$_.Exception.Message })
        }
        finally {
            if ($sess -and $apiV) { try { OV-Logout -A $ip -S $sess -V $apiV } catch {} }
        }
    }

    # Runspace-Pool für echte Parallelität
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $pool = [runspacefactory]::CreateRunspacePool(1, $maxParallel, $iss, $Host)
    $pool.Open()
    $script:uploadPool = $pool
    $script:uploadJobs = @()

    foreach ($entry in $appliances) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        $null = $ps.AddScript($worker).AddArgument(@{
            restCode  = $script:restCode
            ip        = $entry.IP
            ver       = $entry.Version
            adminUser = $txtAdminUser.Text
            adminPass = $txtAdminPass.Text
            filePath  = $txtT1File.Text
            uiQueue   = $script:uiQueue
        })
        $handle = $ps.BeginInvoke()
        $script:uploadJobs += [PSCustomObject]@{ PS=$ps; Handle=$handle; IP=$entry.IP }
    }

    # Watchdog-Timer: prüft ob alle Jobs fertig sind, dann FINISHED + Pool-Cleanup
    if ($script:uploadWatch) { $script:uploadWatch.Stop(); $script:uploadWatch.Dispose() }
    $script:uploadWatch = New-Object System.Windows.Forms.Timer
    $script:uploadWatch.Interval = 1000
    $script:uploadWatch.Add_Tick({
        $allDone = $true
        foreach ($j in $script:uploadJobs) { if (-not $j.Handle.IsCompleted) { $allDone = $false; break } }
        if ($allDone) {
            $script:uploadWatch.Stop()
            foreach ($j in $script:uploadJobs) {
                try { $null = $j.PS.EndInvoke($j.Handle) } catch {}
                try { $j.PS.Dispose() } catch {}
            }
            try { $script:uploadPool.Close(); $script:uploadPool.Dispose() } catch {}
            $script:uploadJobs = @()
            $script:uiQueue.Enqueue(@{ Type='FINISHED' })
        }
    })
    $script:uploadWatch.Start()
})

# ═══════════════════════════════════════════════════════════════════
#  EVENT-HANDLER: Tab 2 – Bundles auflisten
# ═══════════════════════════════════════════════════════════════════
$btnT2Load.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    if ($cboT2Appl.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Appliance auswählen.", "Hinweis", 'OK', 'Warning'); return
    }
    $btnT2Load.Enabled = $false
    $appliance = Get-IPFromCombo $cboT2Appl.SelectedItem.ToString()
    $logBox.AppendText("Lade Firmware-Bundles von $appliance...`r`n"); $logBox.ScrollToCaret()

    Start-AsyncOp -Block {
        param($restCode, $appliance, $adminUser, $adminPass, $uiQueue)
        Invoke-Expression $restCode
        try {
            $v = OV-GetApiVersion -A $appliance
            $s = OV-Login -A $appliance -U $adminUser -P $adminPass -V $v
            $resp = OV-Rest -A $appliance -S $s -V $v -M Get -E "/rest/firmware-drivers?start=0&count=2000"
            OV-Logout -A $appliance -S $s -V $v
            $bundles = if ($resp.members) { $resp.members } elseif ($resp -is [array]) { $resp } else { @($resp) }
            $uiQueue.Enqueue(@{ Type='BUNDLE_LIST_T2'; Data=$bundles })
        }
        catch { $uiQueue.Enqueue(@{ Type='ERROR'; Text="$appliance – $($_.Exception.Message)" }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Arguments @($script:restCode, $appliance, $txtAdminUser.Text, $txtAdminPass.Text, $script:uiQueue)
})

$btnT2Export.Add_Click({
    if ($dgvT2.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren.", "Hinweis", 'OK', 'Information'); return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
    $sfd.FileName = "FirmwareBundles_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    if ($sfd.ShowDialog() -ne 'OK') { return }
    $rows = @()
    foreach ($row in $dgvT2.Rows) {
        $rows += [PSCustomObject]@{
            Name = $row.Cells["name"].Value
            Version = $row.Cells["version"].Value
            BundleType = $row.Cells["bundleType"].Value
            ReleaseDate = $row.Cells["releaseDate"].Value
            ResourceState = $row.Cells["resourceState"].Value
            URI = $row.Cells["uri"].Value
        }
    }
    $rows | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $logBox.AppendText("CSV gespeichert: $($sfd.FileName)`r`n"); $logBox.ScrollToCaret()
})

# ═══════════════════════════════════════════════════════════════════
#  EVENT-HANDLER: Tab 3 – Bundles laden + löschen
# ═══════════════════════════════════════════════════════════════════
$btnT3Refresh.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    $appliances = Get-CheckedAppliances
    if ($appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Appliances ausgewählt.", "Fehler", 'OK', 'Warning'); return
    }
    $btnT3Refresh.Enabled = $false
    $first = $appliances[0].IP
    $logBox.AppendText("Lade Firmware-Bundles von $first...`r`n"); $logBox.ScrollToCaret()

    Start-AsyncOp -Block {
        param($restCode, $appliance, $adminUser, $adminPass, $uiQueue)
        Invoke-Expression $restCode
        try {
            $v = OV-GetApiVersion -A $appliance
            $s = OV-Login -A $appliance -U $adminUser -P $adminPass -V $v
            $resp = OV-Rest -A $appliance -S $s -V $v -M Get -E "/rest/firmware-drivers?start=0&count=2000"
            OV-Logout -A $appliance -S $s -V $v
            $bundles = if ($resp.members) { $resp.members } elseif ($resp -is [array]) { $resp } else { @($resp) }
            $uiQueue.Enqueue(@{ Type='BUNDLE_LIST_T3'; Data=$bundles; Source=$appliance })
        }
        catch { $uiQueue.Enqueue(@{ Type='ERROR'; Text="$appliance – $($_.Exception.Message)" }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Arguments @($script:restCode, $first, $txtAdminUser.Text, $txtAdminPass.Text, $script:uiQueue)
})

$btnT3Delete.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtAdminUser.Text) -or [string]::IsNullOrWhiteSpace($txtAdminPass.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Admin-Benutzername und Passwort eingeben.", "Credentials fehlen", 'OK', 'Warning'); return
    }
    $toDelete = @()
    foreach ($it in $lvT3.Items) {
        if ($it.Checked) {
            $toDelete += [PSCustomObject]@{
                Name    = $it.Tag.name
                Version = $it.Tag.version
                Uri     = $it.Tag.uri
            }
        }
    }
    if ($toDelete.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Bundles markiert.", "Hinweis", 'OK', 'Information'); return
    }
    $appliances = Get-CheckedAppliances
    if ($appliances.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Appliances ausgewählt.", "Fehler", 'OK', 'Warning'); return
    }
    $namesPreview = ($toDelete | Select-Object -First 5 | ForEach-Object { "  • $($_.Name) ($($_.Version))" }) -join "`n"
    if ($toDelete.Count -gt 5) { $namesPreview += "`n  • ... und $($toDelete.Count - 5) weitere" }

    $res = [System.Windows.Forms.MessageBox]::Show(
        "⚠ $($toDelete.Count) Bundle(s) auf $($appliances.Count) Appliance(s) UNWIDERRUFLICH löschen?`n`n$namesPreview",
        "Warnung", 'YesNo', 'Warning')
    if ($res -ne 'Yes') { return }

    $btnT3Delete.Enabled = $false
    $logBox.AppendText("=== Lösche $($toDelete.Count) Bundle(s) auf $($appliances.Count) Appliance(s) ===`r`n"); $logBox.ScrollToCaret()

    Start-AsyncOp -Block {
        param($p)
        $restCode=$p.restCode; $appliances=$p.appliances; $adminUser=$p.adminUser; $adminPass=$p.adminPass
        $items=$p.items; $uiQueue=$p.uiQueue
        Invoke-Expression $restCode
        try {
            foreach ($entry in $appliances) {
                $ip = $entry.IP
                try {
                    $v = OV-GetApiVersion -A $ip
                    $s = OV-Login -A $ip -U $adminUser -P $adminPass -V $v
                    foreach ($b in $items) {
                        try {
                            # Bevorzugt: URI, falls vorhanden
                            $endpoint = $null
                            if ($b.Uri) { $endpoint = $b.Uri }
                            else { $endpoint = "/rest/firmware-drivers/$($b.Name),$($b.Version)" }
                            OV-Rest -A $ip -S $s -V $v -M Delete -E $endpoint | Out-Null
                            $uiQueue.Enqueue(@{ Type='LOG'; Text="$ip – Gelöscht: $($b.Name) ($($b.Version))" })
                        }
                        catch {
                            # Fallback: per Suche das Bundle auf dieser Appliance finden und über dessen URI löschen
                            try {
                                $list = OV-Rest -A $ip -S $s -V $v -M Get -E "/rest/firmware-drivers?filter=name='$([uri]::EscapeDataString($b.Name))'&count=50"
                                $match = $null
                                if ($list.members) {
                                    $match = $list.members | Where-Object { $_.name -eq $b.Name -and (-not $b.Version -or $_.version -eq $b.Version) } | Select-Object -First 1
                                }
                                if ($match -and $match.uri) {
                                    OV-Rest -A $ip -S $s -V $v -M Delete -E $match.uri | Out-Null
                                    $uiQueue.Enqueue(@{ Type='LOG'; Text="$ip – Gelöscht (Fallback): $($b.Name) ($($b.Version))" })
                                } else {
                                    $uiQueue.Enqueue(@{ Type='ERROR'; Text="$ip – Bundle nicht gefunden: $($b.Name) ($($b.Version)) – $($_.Exception.Message)" })
                                }
                            }
                            catch {
                                $uiQueue.Enqueue(@{ Type='ERROR'; Text="$ip – Löschen '$($b.Name)' fehlgeschlagen: $($_.Exception.Message)" })
                            }
                        }
                    }
                    OV-Logout -A $ip -S $s -V $v
                }
                catch {
                    $uiQueue.Enqueue(@{ Type='ERROR'; Text="$ip – Login/Verbindung: $($_.Exception.Message)" })
                }
            }
        }
        catch { $uiQueue.Enqueue(@{ Type='CRITICAL_ERROR'; Error=$_.Exception.Message }) }
        $uiQueue.Enqueue(@{ Type='FINISHED' })
    } -Params @{
        restCode=$script:restCode; appliances=$appliances; adminUser=$txtAdminUser.Text
        adminPass=$txtAdminPass.Text; items=$toDelete; uiQueue=$script:uiQueue
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
