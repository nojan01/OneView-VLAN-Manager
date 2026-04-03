# Global: Skriptordner ermitteln – alle relativen Pfade basieren darauf
$global:scriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path


# HPE OneView Module Management
Get-Module -Name "HPEOneView*" | Remove-Module -Force -ErrorAction SilentlyContinue
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

try {
    Import-Module hpeoneview.1000 -Force -ErrorAction Stop
    Write-Host "✅ HPEOneView.1000 erfolgreich geladen" -ForegroundColor Green
} catch {
    Write-Host "❌ FEHLER: Konnte hpeoneview.1000 nicht laden: $_" -ForegroundColor Red
    exit 1
}

# ============================================
# SSL/TLS-Zertifikatsvalidierung deaktivieren
# (für selbstsignierte OneView-Zertifikate)
# ============================================
Write-Host "Deaktiviere SSL-Zertifikatsvalidierung für selbstsignierte Zertifikate..." -ForegroundColor Yellow

# Für PowerShell 7+: SkipCertificateCheck als Standard für alle REST-Aufrufe setzen
if ($PSVersionTable.PSVersion.Major -ge 7) {
    $PSDefaultParameterValues['Invoke-RestMethod:SkipCertificateCheck'] = $true
    $PSDefaultParameterValues['Invoke-WebRequest:SkipCertificateCheck'] = $true
}

# .NET-Level: Zertifikatsvalidierung global deaktivieren (betrifft auch HPE OneView Modul)
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

Write-Host "✅ SSL-Zertifikatsvalidierung deaktiviert" -ForegroundColor Green

# Prüfe PowerShell Version und zeige Warnung bei Bedarf
function Test-PowerShellVersion {
    $psVersion = $PSVersionTable.PSVersion
    Write-Host "Erkannte PowerShell Version: $psVersion" -ForegroundColor Cyan
    
    if ($psVersion.Major -lt 7) {
        Write-Host "" -ForegroundColor Yellow
        Write-Host "⚠️  WARNUNG: PowerShell Version $psVersion erkannt!" -ForegroundColor Yellow
        Write-Host "HPE OneView empfiehlt PowerShell 7 oder höher für optimale Kompatibilität." -ForegroundColor Yellow
        Write-Host "" -ForegroundColor Yellow
        Write-Host "Download PowerShell 7: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Cyan
        Write-Host "Installation: winget install Microsoft.PowerShell" -ForegroundColor Cyan
        Write-Host "" -ForegroundColor Yellow
        
        $continue = Read-Host "Möchten Sie trotzdem fortfahren? (j/N)"
        if ($continue -notmatch '^[jJyY]') {
            Write-Host "Script beendet. Bitte installieren Sie PowerShell 7." -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "✅ PowerShell 7+ erkannt - Optimal für HPE OneView!" -ForegroundColor Green
    }
}

# Ändere Hintergrund- und Schriftfarbe sowie den Fenstertitel
$host.UI.RawUI.BackgroundColor = "DarkBlue"
$host.UI.RawUI.ForegroundColor = "White"
Clear-Host
$host.UI.RawUI.WindowTitle = "© 2025 N.J. Airbus D&S - OneView/ Synergy Update Tool"

Write-Host "=================================================" -ForegroundColor Yellow
Write-Host "        HPE OneView/ Synergy Update Tool         " -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Yellow
Write-Host " "

# ============================================
# Assemblies für Windows Forms und Drawing laden
# ============================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Modul: Gemeinsame Hilfsfunktionen und API-Typen laden
function Initialize-NativeAPIs {
    if (-not ([System.Management.Automation.PSTypeName]"NativeMethods").Type) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential)]
public struct COORD {
    public short X;
    public short Y;
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct CONSOLE_FONT_INFOEX {
    public uint cbSize;
    public uint nFont;
    public COORD dwFontSize;
    public int FontFamily;
    public int FontWeight;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
    public string FaceName;
}

public class NativeMethods {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool AttachConsole(uint dwProcessId);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool FreeConsole();

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool SetCurrentConsoleFontEx(IntPtr hConsoleOutput, bool bMaximumWindow, ref CONSOLE_FONT_INFOEX lpConsoleCurrentFontEx);

    public const int GWL_STYLE = -16;
    public const int WS_VISIBLE  = 0x10000000;
    public const int WS_CHILD    = 0x40000000;
    public const int WS_CAPTION  = 0x00C00000;
    public const int WS_THICKFRAME = 0x00040000;
    public const int WS_BORDER   = 0x00800000;
    public const int WS_SYSMENU  = 0x00080000;
    public const int SW_SHOW = 5;
    public const int STD_OUTPUT_HANDLE = -11;
    public const int WS_DISABLED = 0x08000000;
    public const int GWL_EXSTYLE = -20;
    public const int WS_EX_CLIENTEDGE = 0x00000200;
    public const int WS_EX_WINDOWEDGE = 0x00000100;
    public const int WS_EX_DLGMODALFRAME = 0x00000001;
    public const int WS_EX_STATICEDGE = 0x00020000;
}
"@
    }
    if (-not ([System.Management.Automation.PSTypeName]"Win32").Type) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
}
[StructLayout(LayoutKind.Sequential)]
public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}
"@
    }
}

# Funktion zum Zentrieren des Konsolenfensters
function Set-ConsoleWindowPosition {
    param(
        [int]$WindowPixelWidth = 800,
        [int]$WindowPixelHeight = 600
    )
    Initialize-NativeAPIs

    $consoleHandle = [NativeMethods]::GetConsoleWindow()
    if ($consoleHandle -eq [IntPtr]::Zero) {
        Write-Error "Konsolenfenster wurde nicht gefunden."
        return
    }

    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $screenWidth = $screen.Width
    $screenHeight = $screen.Height

    $windowLeft = [int](($screenWidth - $WindowPixelWidth) / 2)
    $windowTop  = [int](($screenHeight - $WindowPixelHeight) / 2)
    $SWP_NOSIZE = 0x0001
    $result = [NativeMethods]::SetWindowPos($consoleHandle, [IntPtr]::Zero, $windowLeft, $windowTop, 0, 0, $SWP_NOSIZE)
    if (-not $result) {
        Write-Error "SetWindowPos konnte die Position nicht ändern."
    }
}

# Hilfsfunktion: Formular-Größe an Bildschirm anpassen (für kleine Monitore)
function Get-ScaledFormSize {
    param(
        [int]$DesiredWidth,
        [int]$DesiredHeight,
        [int]$MinWidth = 350,
        [int]$MinHeight = 250
    )
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $maxW = [int]($screen.Width * 0.92)
    $maxH = [int]($screen.Height * 0.92)
    $w = [Math]::Min($DesiredWidth, $maxW)
    $h = [Math]::Min($DesiredHeight, $maxH)
    $w = [Math]::Max($w, $MinWidth)
    $h = [Math]::Max($h, $MinHeight)
    return @{ Width = $w; Height = $h; ScreenWidth = $screen.Width; ScreenHeight = $screen.Height }
}

# Progress-Anzeige Funktionen
function Show-Progress {
    param(
        [int]$Current,
        [int]$Total,
        [string]$Activity = "Verarbeitung",
        [string]$Status = ""
    )

    if ($Total -gt 0) {
        $percentComplete = [Math]::Round(($Current / $Total) * 100)
        Write-Progress -Activity $Activity -Status "$Status ($Current von $Total)" -PercentComplete $percentComplete
    }
}

function ConvertTo-NormalizedVersion {
    param(
        [string]$version
    )
    try {
        $cleanVersion = $version -replace "(-.*)$", ""
        $segments = $cleanVersion -split '\.'
        
        while ($segments.Count -lt 3) {
            $segments += "0"
        }
        
        $versionObject = [PSCustomObject]@{
            Major = [int]$segments[0]
            Minor = [int]$segments[1]
            Build = [int]$segments[2]
            OriginalString = $cleanVersion
            MajorString = $segments[0]
            MinorString = $segments[1]
            BuildString = $segments[2]
        }
        
        $versionObject | Add-Member -MemberType ScriptMethod -Name "CompareTo" -Value {
            param($other)
            
            if ($other -is [System.Version]) {
                $otherMajor = $other.Major; $otherMinor = $other.Minor; $otherBuild = $other.Build
            }
            elseif ($null -ne $other.Major) {
                $otherMajor = $other.Major; $otherMinor = $other.Minor; $otherBuild = $other.Build
            }
            else {
                throw "Unbekannter Versionstyp für Vergleich"
            }
            
            if ($this.Major -lt $otherMajor) { return -1 }
            if ($this.Major -gt $otherMajor) { return 1 }
            if ($this.Minor -lt $otherMinor) { return -1 }
            if ($this.Minor -gt $otherMinor) { return 1 }
            if ($this.Build -lt $otherBuild) { return -1 }
            if ($this.Build -gt $otherBuild) { return 1 }
            
            return 0
        }
        
        $versionObject | Add-Member -MemberType ScriptMethod -Name "ToString" -Force -Value {
            return $this.OriginalString
        }
        
        return $versionObject
    }
    catch {
        Write-Warning "Fehler beim Parsen der Version '$version': $($_.Exception.Message)"
        try {
            return [System.Version]$version
        }
        catch {
            return ConvertTo-NormalizedVersion "0.0.0"
        }
    }
}

# Funktion zur automatischen Ermittlung der API-Version via GET /rest/version (ohne Auth)
function Get-ApiVersionInline {
    param(
        [Parameter(Mandatory)][string]$Hostname,
        [int]$FallbackVersion = 8200
    )
    $uri = "https://$Hostname/rest/version"
    try {
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            $response = Invoke-RestMethod -Uri $uri -Method Get -SkipCertificateCheck -TimeoutSec 10 -ErrorAction Stop
        } else {
            $response = Invoke-RestMethod -Uri $uri -Method Get -TimeoutSec 10 -ErrorAction Stop
        }
        return [int]$response.currentVersion
    }
    catch {
        return $FallbackVersion
    }
}

# Funktion zum Erstellen eines Worker-Skripts für parallele Ausführung
function New-WorkerScript {
    param(
        [string]$WorkerScriptPath,
        [string[]]$ApplianceGroup,
        $DesiredVersion,
        [string]$UpdateFilePath,
        [string]$LogDir,
        [string]$WorkerName,
        [string]$UserName,
        [string]$Password,
        [string]$Passphrase
    )
    
    $workerLines = @()
    $workerLines += "# Worker-Skript für parallele OneView Updates - $WorkerName"
    $workerLines += "# Automatisch generiert am $(Get-Date)"
    $workerLines += ""
    $workerLines += "# Global: Skriptordner ermitteln"
    $workerLines += '$global:scriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path'
    $workerLines += '$workerDisplayName = "' + $WorkerName + '"'
    $workerLines += ""
    $workerLines += "# Ändere Hintergrund- und Schriftfarbe für bessere Unterscheidung"
    $workerLines += '$host.UI.RawUI.BackgroundColor = "Black"'
    $workerLines += '$host.UI.RawUI.ForegroundColor = "White"'
    $workerLines += 'Clear-Host'
    $workerLines += '$host.UI.RawUI.WindowTitle = "© 2025 N.J. OneView Update Tool - ' + $WorkerName + '"'
    $workerLines += ""
    $workerLines += "# NativeMethods fuer Console-Handle und Font"
    $workerLines += 'if (-not ([System.Management.Automation.PSTypeName]"WorkerNative").Type) {'
    $workerLines += '    Add-Type @"'
    $workerLines += 'using System;'
    $workerLines += 'using System.Runtime.InteropServices;'
    $workerLines += '[StructLayout(LayoutKind.Sequential)]'
    $workerLines += 'public struct W_COORD {'
    $workerLines += '    public short X;'
    $workerLines += '    public short Y;'
    $workerLines += '}'
    $workerLines += '[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]'
    $workerLines += 'public struct W_CONSOLE_FONT_INFOEX {'
    $workerLines += '    public uint cbSize;'
    $workerLines += '    public uint nFont;'
    $workerLines += '    public W_COORD dwFontSize;'
    $workerLines += '    public int FontFamily;'
    $workerLines += '    public int FontWeight;'
    $workerLines += '    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]'
    $workerLines += '    public string FaceName;'
    $workerLines += '}'
    $workerLines += 'public class WorkerNative {'
    $workerLines += '    [DllImport("kernel32.dll")]'
    $workerLines += '    public static extern IntPtr GetConsoleWindow();'
    $workerLines += '    [DllImport("kernel32.dll")]'
    $workerLines += '    public static extern IntPtr GetStdHandle(int nStdHandle);'
    $workerLines += '    [DllImport("kernel32.dll", CharSet=CharSet.Unicode)]'
    $workerLines += '    public static extern bool SetCurrentConsoleFontEx(IntPtr hConsoleOutput, bool bMaximumWindow, ref W_CONSOLE_FONT_INFOEX lpConsoleCurrentFontEx);'
    $workerLines += '    [DllImport("kernel32.dll", SetLastError=true)]'
    $workerLines += '    public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);'
    $workerLines += '    [DllImport("kernel32.dll", SetLastError=true)]'
    $workerLines += '    public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);'
    $workerLines += '    public const int STD_OUTPUT_HANDLE = -11;'
    $workerLines += '    public const int STD_INPUT_HANDLE = -10;'
    $workerLines += '}'
    $workerLines += '"@'
    $workerLines += '}'
    $workerLines += ""
    $workerLines += 'Write-Host "==========================================================" -ForegroundColor Yellow'
    $workerLines += 'Write-Host "        HPE OneView Update Tool - ' + $WorkerName + '           " -ForegroundColor Cyan'
    $workerLines += 'Write-Host "==========================================================" -ForegroundColor Yellow'
    $workerLines += 'Write-Host " "'
    $workerLines += ""
    $workerLines += "# Prüfe PowerShell Version"
    $workerLines += 'Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan'
    $workerLines += 'if ($PSVersionTable.PSVersion.Major -lt 7) {'
    $workerLines += '    Write-Host "WARNUNG: PowerShell 7+ empfohlen für HPE OneView!" -ForegroundColor Yellow'
    $workerLines += '}'
    $workerLines += ""
    $workerLines += "# HPE OneView Modul laden"
    $workerLines += "try {"
    $workerLines += "    Import-Module hpeoneview.1000 -Force"
    $workerLines += '    Write-Host "HPE OneView Modul erfolgreich geladen." -ForegroundColor Green'
    $workerLines += "} catch {"
    $workerLines += '    Write-Host "Fehler beim Laden des HPE OneView Moduls: $_" -ForegroundColor Red'
    $workerLines += '    Read-Host "Drücke Enter zum Beenden..."'
    $workerLines += "    exit"
    $workerLines += "}"
    $workerLines += ""
    $workerLines += "# SSL/TLS-Zertifikatsvalidierung deaktivieren (selbstsignierte OneView-Zertifikate)"
    $workerLines += 'Write-Host "Deaktiviere SSL-Zertifikatsvalidierung..." -ForegroundColor Yellow'
    $workerLines += 'if ($PSVersionTable.PSVersion.Major -ge 7) {'
    $workerLines += "    `$PSDefaultParameterValues['Invoke-RestMethod:SkipCertificateCheck'] = `$true"
    $workerLines += "    `$PSDefaultParameterValues['Invoke-WebRequest:SkipCertificateCheck'] = `$true"
    $workerLines += '}'
    $workerLines += '[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }'
    $workerLines += 'Write-Host "SSL-Zertifikatsvalidierung deaktiviert" -ForegroundColor Green'
    $workerLines += ""
    # Anmeldedaten als Klartext einbetten und im Worker als SecureString/PSCredential erstellen
    $escapedPassword = $Password -replace "'", "''"
    $escapedPassphrase = $Passphrase -replace "'", "''"
    $workerLines += "# Anmeldedaten"
    $workerLines += "`$securePasswordForConnect = ConvertTo-SecureString '$escapedPassword' -AsPlainText -Force"
    $workerLines += "`$securePassphrase = ConvertTo-SecureString '$escapedPassphrase' -AsPlainText -Force"
    $workerLines += "`$credentials = New-Object System.Management.Automation.PSCredential ('$($UserName)', `$securePasswordForConnect)"
    $workerLines += ""
    $workerLines += "# Funktion zur präzisen Versionsnormalisierung (behält führende Nullen)"
    $workerLines += "function ConvertTo-NormalizedVersion {"
    $workerLines += "    param([string]`$version)"
    $workerLines += "    try {"
    $workerLines += '        $cleanVersion = $version -replace "(-.*)$", ""'
    $workerLines += '        $segments = $cleanVersion -split "\."'
    $workerLines += '        while ($segments.Count -lt 3) { $segments += "0" }'
    $workerLines += ""
    $workerLines += "        `$versionObject = [PSCustomObject]@{"
    $workerLines += "            Major = [int]`$segments[0]"
    $workerLines += "            Minor = [int]`$segments[1]"
    $workerLines += "            Build = [int]`$segments[2]"
    $workerLines += "            OriginalString = `$cleanVersion"
    $workerLines += "            MajorString = `$segments[0]"
    $workerLines += "            MinorString = `$segments[1]"
    $workerLines += "            BuildString = `$segments[2]"
    $workerLines += "        }"
    $workerLines += ""
    $workerLines += "        `$versionObject | Add-Member -MemberType ScriptMethod -Name 'CompareTo' -Value {"
    $workerLines += "            param(`$other)"
    $workerLines += "            if (`$other -is [System.Version]) {"
    $workerLines += "                `$otherMajor = `$other.Major; `$otherMinor = `$other.Minor; `$otherBuild = `$other.Build"
    $workerLines += "            } elseif (`$null -ne `$other.Major) {"
    $workerLines += "                `$otherMajor = `$other.Major; `$otherMinor = `$other.Minor; `$otherBuild = `$other.Build"
    $workerLines += "            } else { throw 'Unbekannter Versionstyp für Vergleich' }"
    $workerLines += "            if (`$this.Major -lt `$otherMajor) { return -1 }"
    $workerLines += "            if (`$this.Major -gt `$otherMajor) { return 1 }"
    $workerLines += "            if (`$this.Minor -lt `$otherMinor) { return -1 }"
    $workerLines += "            if (`$this.Minor -gt `$otherMinor) { return 1 }"
    $workerLines += "            if (`$this.Build -lt `$otherBuild) { return -1 }"
    $workerLines += "            if (`$this.Build -gt `$otherBuild) { return 1 }"
    $workerLines += "            return 0"
    $workerLines += "        }"
    $workerLines += ""
    $workerLines += "        `$versionObject | Add-Member -MemberType ScriptMethod -Name 'ToString' -Force -Value {"
    $workerLines += "            return `$this.OriginalString"
    $workerLines += "        }"
    $workerLines += ""
    $workerLines += "        return `$versionObject"
    $workerLines += "    }"
    $workerLines += "    catch {"
    $workerLines += "        try { return [System.Version]`$version }"
    $workerLines += "        catch { return ConvertTo-NormalizedVersion '0.0.0' }"
    $workerLines += "    }"
    $workerLines += "}"
    $workerLines += ""
    $workerLines += "# Definiere Appliances für diesen Worker"
    $workerLines += '$ApplianceIPs = @("' + ($ApplianceGroup -join '", "') + '")'
    $workerLines += '$desiredVersion = "' + $DesiredVersion.ToString() + '"'
    $workerLines += '$updateFilePath = "' + $UpdateFilePath + '"'
    $workerLines += '$logDir = "' + $LogDir + '"'
    $workerLines += 'if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }'
    $workerLines += ""
    $workerLines += "# Console-Handle in Datei schreiben fuer GUI-Embedding"
    $workerLines += '$hwnd = [WorkerNative]::GetConsoleWindow()'
    $workerLines += "`$hwnd.ToInt64() | Out-File -FilePath (Join-Path `$logDir '${WorkerName}_Handle.txt') -Force"
    $workerLines += ""
    $workerLines += "# QuickEdit deaktivieren (verhindert Crash beim Klicken im eingebetteten Fenster)"
    $workerLines += '$hIn = [WorkerNative]::GetStdHandle([WorkerNative]::STD_INPUT_HANDLE)'
    $workerLines += 'if ($hIn -ne [IntPtr]::Zero -and $hIn -ne ([IntPtr]::new(-1))) {'
    $workerLines += '    [uint32]$mode = 0'
    $workerLines += '    [WorkerNative]::GetConsoleMode($hIn, [ref]$mode) | Out-Null'
    $workerLines += '    $ENABLE_QUICK_EDIT = 0x0040'
    $workerLines += '    $ENABLE_MOUSE_INPUT = 0x0010'
    $workerLines += '    $mode = $mode -band (-bnot $ENABLE_QUICK_EDIT)'
    $workerLines += '    $mode = $mode -band (-bnot $ENABLE_MOUSE_INPUT)'
    $workerLines += '    [WorkerNative]::SetConsoleMode($hIn, $mode) | Out-Null'
    $workerLines += '}'
    $workerLines += ""
    $workerLines += "# Console-Font setzen"
    $workerLines += '$fi = New-Object W_CONSOLE_FONT_INFOEX'
    $workerLines += '$fi.cbSize = [uint32]84'
    $workerLines += '$fi.nFont = [uint32]0'
    $workerLines += '$c = New-Object W_COORD'
    $workerLines += '$c.X = [short]0'
    $workerLines += '$c.Y = [short]14'
    $workerLines += '$fi.dwFontSize = $c'
    $workerLines += '$fi.FontFamily = 54'
    $workerLines += '$fi.FontWeight = 400'
    $workerLines += '$fi.FaceName = "Consolas"'
    $workerLines += '$hOut = [WorkerNative]::GetStdHandle([WorkerNative]::STD_OUTPUT_HANDLE)'
    $workerLines += 'if ($hOut -ne [IntPtr]::Zero -and $hOut -ne ([IntPtr]::new(-1))) {'
    $workerLines += '    [WorkerNative]::SetCurrentConsoleFontEx($hOut, $false, [ref]$fi) | Out-Null'
    $workerLines += '}'
    $workerLines += ""
    $workerLines += "# FileSystemWatcher fuer dynamische Schriftgroesse (GUI schreibt gewuenschte Groesse)"
    $workerLines += "`$fontWatcher = New-Object System.IO.FileSystemWatcher"
    $workerLines += "`$fontWatcher.Path = `$logDir"
    $workerLines += "`$fontWatcher.Filter = '${WorkerName}_FontSize.txt'"
    $workerLines += "`$fontWatcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite"
    $workerLines += 'Register-ObjectEvent -InputObject $fontWatcher -EventName Changed -Action {'
    $workerLines += '    try {'
    $workerLines += '        Start-Sleep -Milliseconds 50'
    $workerLines += '        $fontSize = [int](Get-Content $Event.SourceEventArgs.FullPath -Raw).Trim()'
    $workerLines += '        if ($fontSize -ge 8 -and $fontSize -le 22) {'
    $workerLines += '            $fi2 = New-Object W_CONSOLE_FONT_INFOEX'
    $workerLines += '            $fi2.cbSize = [uint32]84'
    $workerLines += '            $fi2.nFont = [uint32]0'
    $workerLines += '            $c2 = New-Object W_COORD'
    $workerLines += '            $c2.X = [short]0'
    $workerLines += '            $c2.Y = [short]$fontSize'
    $workerLines += '            $fi2.dwFontSize = $c2'
    $workerLines += '            $fi2.FontFamily = 54'
    $workerLines += '            $fi2.FontWeight = 400'
    $workerLines += '            $fi2.FaceName = "Consolas"'
    $workerLines += '            $hOut2 = [WorkerNative]::GetStdHandle([WorkerNative]::STD_OUTPUT_HANDLE)'
    $workerLines += '            if ($hOut2 -ne [IntPtr]::Zero -and $hOut2 -ne ([IntPtr]::new(-1))) {'
    $workerLines += '                [WorkerNative]::SetCurrentConsoleFontEx($hOut2, $false, [ref]$fi2) | Out-Null'
    $workerLines += '            }'
    $workerLines += '        }'
    $workerLines += '    } catch { }'
    $workerLines += '} | Out-Null'
    $workerLines += '$fontWatcher.EnableRaisingEvents = $true'
    $workerLines += ""
    $workerLines += 'Write-Host "' + $WorkerName + ' startet Updates fuer folgende Appliances" -ForegroundColor Cyan'
    $workerLines += '$ApplianceIPs | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }'
    $workerLines += 'Write-Host " "'
    $workerLines += ""
    $workerLines += "# Log-Datei fuer diesen Worker"
    $workerLines += '$logFile = Join-Path $logDir "' + $WorkerName + '-OneView_Update_Log.txt"'
    $workerLines += 'if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }'
    $workerLines += 'if (Test-Path $logFile) {'
    $workerLines += '    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"'
    $workerLines += '    $archivedLogFile = Join-Path $logDir ("' + $WorkerName + '-OneView_Update_Log_" + $timestamp + ".txt")'
    $workerLines += "    Rename-Item -Path `$logFile -NewName `$archivedLogFile -Force"
    $workerLines += "}"
    $workerLines += "New-Item -ItemType File -Path `$logFile -Force | Out-Null"
    $workerLines += ""
    $workerLines += "foreach (`$Appliance in `$ApplianceIPs) {"
    $workerLines += '    Write-Host "========================================" -ForegroundColor Cyan'
    $workerLines += "    Write-Host (`$workerDisplayName + ': Bearbeite ' + `$Appliance) -ForegroundColor Yellow"
    $workerLines += '    Write-Host "========================================" -ForegroundColor Cyan'
    $workerLines += "    Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] ======== Bearbeite ' + `$Appliance + ' ========')"
    $workerLines += "    "
    $workerLines += "    # API-Version automatisch ermitteln"
    $workerLines += '    $apiVersionUri = "https://" + $Appliance + "/rest/version"'
    $workerLines += '    $apiVersion = 8200'
    $workerLines += '    try {'
    $workerLines += '        if ($PSVersionTable.PSVersion.Major -ge 7) {'
    $workerLines += '            $apiResp = Invoke-RestMethod -Uri $apiVersionUri -Method Get -SkipCertificateCheck -TimeoutSec 10 -ErrorAction Stop'
    $workerLines += '        } else {'
    $workerLines += '            $apiResp = Invoke-RestMethod -Uri $apiVersionUri -Method Get -TimeoutSec 10 -ErrorAction Stop'
    $workerLines += '        }'
    $workerLines += '        $apiVersion = [int]$apiResp.currentVersion'
    $workerLines += '    } catch {'
    $workerLines += '        Write-Host ("' + $WorkerName + ': API-Version konnte nicht ermittelt werden, verwende Fallback 8200") -ForegroundColor Yellow'
    $workerLines += '    }'
    $workerLines += '    Write-Host ("' + $WorkerName + ': API-Version: " + $apiVersion) -ForegroundColor Gray'
    $workerLines += "    "
    $workerLines += "    `$headers = @{"
    $workerLines += '        "X-API-Version" = $apiVersion.ToString()'
    $workerLines += '        "Accept" = "application/json"'
    $workerLines += '        "Authorization" = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credentials.UserName + ":" + $credentials.GetNetworkCredential().Password))'
    $workerLines += "    }"
    $workerLines += ""
    $workerLines += '    $uri = "https://" + $Appliance + "/rest/appliance/nodeinfo/version"'
    $workerLines += "    try {"
    $workerLines += "        if (`$PSVersionTable.PSVersion.Major -ge 7) {"
    $workerLines += "            `$response = Invoke-RestMethod -Uri `$uri -Headers `$headers -Method Get -TimeoutSec 600 -SkipCertificateCheck -ErrorAction Stop"
    $workerLines += "        } else {"
    $workerLines += "            `$response = Invoke-RestMethod -Uri `$uri -Headers `$headers -Method Get -TimeoutSec 600 -ErrorAction Stop"
    $workerLines += "        }"
    $workerLines += "    } catch {"
    $workerLines += "        Write-Host (`$workerDisplayName + ': Fehler beim Abrufen der Version von ' + `$Appliance + ' : ' + `$_) -ForegroundColor Red"
    $workerLines += "        Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date) + '] Fehler bei API-Abfrage: ' + `$_)"
    $workerLines += "        continue"
    $workerLines += "    }"
    $workerLines += ""
    $workerLines += "    `$currentVersionRaw = `$response.softwareversion"
    $workerLines += "    `$currentVersion = ConvertTo-NormalizedVersion `$currentVersionRaw"
    $workerLines += "    Write-Host (`$workerDisplayName + ': Ermittelte Version fuer ' + `$Appliance + ' : ' + `$currentVersion) -ForegroundColor Yellow"
    $workerLines += "    Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date) + '] HPE OneView Version: ' + `$currentVersion)"
    $workerLines += ""
    $workerLines += "    try {"
    $workerLines += "        `$desiredVersionObj = ConvertTo-NormalizedVersion `$desiredVersion"
    $workerLines += "        `$comparison = `$currentVersion.CompareTo(`$desiredVersionObj)"
    $workerLines += "        "
    $workerLines += "        if (`$comparison -lt 0) {"
    $workerLines += "            Write-Host (`$workerDisplayName + ': Update erforderlich fuer ' + `$Appliance + ' : Aktuelle Version (' + `$currentVersion + ') ist aelter als gewuenschte Version (' + `$desiredVersionObj + ').') -ForegroundColor Yellow"
    $workerLines += "            Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] Update erforderlich: ' + `$currentVersion + ' -> ' + `$desiredVersionObj)"
    $workerLines += "        } elseif (`$comparison -eq 0) {"
    $workerLines += "            Write-Host (`$workerDisplayName + ': Kein Update erforderlich fuer ' + `$Appliance + ' : Versionen sind identisch (' + `$currentVersion + ').') -ForegroundColor Green"
    $workerLines += "            Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] Kein Update erforderlich - Version identisch: ' + `$currentVersion)"
    $workerLines += "            continue"
    $workerLines += "        } else {"
    $workerLines += "            Write-Host (`$workerDisplayName + ': Keine Aktion erforderlich: Aktuelle Version (' + `$currentVersion + ') ist neuer als die gewuenschte Version (' + `$desiredVersionObj + ').') -ForegroundColor Green"
    $workerLines += "            Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] Keine Aktion - Version neuer: ' + `$currentVersion + ' > ' + `$desiredVersionObj)"
    $workerLines += "            continue"
    $workerLines += "        }"
    $workerLines += "    } catch {"
    $workerLines += "        Write-Host (`$workerDisplayName + ': Fehler beim Versionsvergleich: ' + `$_) -ForegroundColor Red"
    $workerLines += "        Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date) + '] Fehler beim Versionsvergleich: ' + `$_)"
    $workerLines += "        continue"
    $workerLines += "    }"
    $workerLines += ""
    $workerLines += "    Write-Host (`$workerDisplayName + ': Verbindung zu ' + `$Appliance + ' wird hergestellt...') -ForegroundColor Yellow"
    $workerLines += "    Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] Verbindung zu ' + `$Appliance + ' wird hergestellt...')"
    $workerLines += "    `$Connection = Connect-OVMgmt -Hostname `$Appliance -Credential `$credentials"
    $workerLines += "    if (-not `$Connection) {"
    $workerLines += "        Disconnect-OVMgmt"
    $workerLines += "        Write-Host (`$workerDisplayName + ': Fehler: Verbindung zu ' + `$Appliance + ' konnte nicht hergestellt werden!') -ForegroundColor Red"
    $workerLines += "        Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date) + '] Verbindung zu ' + `$Appliance + ' fehlgeschlagen!')"
    $workerLines += "        continue"
    $workerLines += "    }"
    $workerLines += ""
    $workerLines += "    try {"
    $workerLines += "        try {"
    $workerLines += "            `$pendingUpdate = Get-OVPendingUpdate -ErrorAction Stop"
    $workerLines += "        } catch {"
    $workerLines += "            Write-Host (`$workerDisplayName + ': Fehler beim Abrufen des Pending Updates: ' + `$_.Exception.Message) -ForegroundColor Red"
    $workerLines += "            `$pendingUpdate = `$null"
    $workerLines += "        }"
    $workerLines += "        "
    $workerLines += "        if (`$pendingUpdate) {"
    $workerLines += "            Write-Host '$WorkerName`: Ein Pending Update wurde gefunden:' -ForegroundColor Red"
    $workerLines += "            Write-Host ('$WorkerName`:   Dateiname    : ' + `$pendingUpdate.FileName) -ForegroundColor Yellow"
    $workerLines += "            Write-Host ('$WorkerName`:   Update-Version: ' + `$pendingUpdate.Version) -ForegroundColor Yellow"
    $workerLines += "            try {"
    $workerLines += "                Remove-OVPendingUpdate -Confirm:`$false -ErrorAction Stop"
    $workerLines += "                Write-Host (`$workerDisplayName + ': Das Pending Update ' + `$pendingUpdate.FileName + ' wurde erfolgreich entfernt.') -ForegroundColor Yellow"
    $workerLines += "            } catch {"
    $workerLines += "                Write-Host (`$workerDisplayName + ': Fehler beim Entfernen des Pending Updates: ' + `$_.Exception.Message) -ForegroundColor Red"
    $workerLines += "            }"
    $workerLines += "        } else {"
    $workerLines += "            Write-Host '$WorkerName`: Kein Pending Update gefunden. Das Update kann fortgesetzt werden.' -ForegroundColor Yellow"
    $workerLines += "        }"
    $workerLines += ""
    $workerLines += "        Write-Host (`$workerDisplayName + ': Update wird mit Datei ' + `$updateFilePath + ' durchgefuehrt....') -ForegroundColor Yellow"
    $workerLines += "        Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] SCHRITT 1: Starte Config Backup vor dem Update...')"
    $workerLines += "        Write-Host '$WorkerName`: SCHRITT 1: Starte Config Backup vor dem Update...' -ForegroundColor Yellow"
    $workerLines += '        $backupDir = Join-Path $logDir "OneView_Backup"'
    $workerLines += '        if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }'
    $workerLines += '        new-OVBackup -Location $backupDir -Force -Passphrase $securePassphrase'
    $workerLines += "        Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] SCHRITT 1: Config Backup abgeschlossen')"
    $workerLines += "        Write-Host '$WorkerName`: SCHRITT 1: Config Backup abgeschlossen' -ForegroundColor Green"
    $workerLines += "        "
    $workerLines += "        Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] SCHRITT 2: Starte Install-OVUpdate...')"
    $workerLines += "        Write-Host '$WorkerName`: SCHRITT 2: Starte Install-OVUpdate...' -ForegroundColor Yellow"
    $workerLines += "        "
    $workerLines += "        `$installOVUpdateCompleted = `$false"
    $workerLines += "        try {"
    $workerLines += "            Install-OVUpdate -file `$updateFilePath -Eula Accept -Confirm:`$false"
    $workerLines += "            Write-Host '$WorkerName`: Install-OVUpdate Befehl abgeschlossen' -ForegroundColor Green"
    $workerLines += "            `$installOVUpdateCompleted = `$true"
    $workerLines += "        } catch {"
    $workerLines += "            Write-Host '$WorkerName`: Install-OVUpdate unterbrochen - Update laeuft im Hintergrund weiter' -ForegroundColor Yellow"
    $workerLines += "            Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] Install-OVUpdate unterbrochen - Update laeuft im Hintergrund weiter')"
    $workerLines += "        }"
    $workerLines += "        "
    $workerLines += "        Disconnect-OVMgmt -ErrorAction SilentlyContinue"
    $workerLines += "        "
    $workerLines += "        Write-Host ' '"
    $workerLines += "        Write-Host ('$WorkerName`: Warte auf Update-Abschluss (Zielversion: ' + `$desiredVersionObj.ToString() + ')...') -ForegroundColor Cyan"
    $workerLines += "        Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] Warte auf Update-Abschluss - Zielversion: ' + `$desiredVersionObj.ToString())"
    $workerLines += "        "
    $workerLines += "        `$maxWaitMinutes = 120"
    $workerLines += "        `$pollInterval = 60"
    $workerLines += "        `$startTime = Get-Date"
    $workerLines += "        `$updateCompleted = `$false"
    $workerLines += "        `$versionUri = 'https://' + `$Appliance + '/rest/appliance/nodeinfo/version'"
    $workerLines += "        "
    $workerLines += "        Write-Host ('$WorkerName`: Warte 2 Minuten bevor erster Versions-Check...') -ForegroundColor Cyan"
    $workerLines += "        Start-Sleep -Seconds 120"
    $workerLines += "        "
    $workerLines += "        while (-not `$updateCompleted -and ((Get-Date) - `$startTime).TotalMinutes -lt `$maxWaitMinutes) {"
    $workerLines += "            `$elapsedMinutes = [math]::Round(((Get-Date) - `$startTime).TotalMinutes, 1)"
    $workerLines += "            try {"
    $workerLines += "                if (`$PSVersionTable.PSVersion.Major -ge 7) {"
    $workerLines += "                    `$versionResponse = Invoke-RestMethod -Uri `$versionUri -Headers `$headers -Method Get -TimeoutSec 15 -SkipCertificateCheck -ErrorAction Stop"
    $workerLines += "                } else {"
    $workerLines += "                    `$versionResponse = Invoke-RestMethod -Uri `$versionUri -Headers `$headers -Method Get -TimeoutSec 15 -ErrorAction Stop"
    $workerLines += "                }"
    $workerLines += "                `$reportedVersion = ConvertTo-NormalizedVersion `$versionResponse.softwareversion"
    $workerLines += "                `$versionComparison = `$reportedVersion.CompareTo(`$desiredVersionObj)"
    $workerLines += "                "
    $workerLines += "                if (`$versionComparison -ge 0) {"
    $workerLines += "                    `$updateCompleted = `$true"
    $workerLines += "                    Write-Host ('$WorkerName`: Appliance meldet Version ' + `$reportedVersion.ToString() + ' - Update abgeschlossen! (nach ' + `$elapsedMinutes + ' Min)') -ForegroundColor Green"
    $workerLines += "                    Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] Update abgeschlossen - Neue Version: ' + `$reportedVersion.ToString() + ' (nach ' + `$elapsedMinutes + ' Min)')"
    $workerLines += "                } else {"
    $workerLines += "                    Write-Host ('$WorkerName`: Warte seit ' + `$elapsedMinutes + ' Min... Aktuelle Version: ' + `$reportedVersion.ToString() + ' (Ziel: ' + `$desiredVersionObj.ToString() + ')') -ForegroundColor Cyan"
    $workerLines += "                    Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] Warte seit ' + `$elapsedMinutes + ' Min... Version: ' + `$reportedVersion.ToString() + ' (Ziel: ' + `$desiredVersionObj.ToString() + ')')"
    $workerLines += "                }"
    $workerLines += "            } catch {"
    $workerLines += "                Write-Host ('$WorkerName`: Warte seit ' + `$elapsedMinutes + ' Min... Appliance nicht erreichbar (Reboot laeuft)') -ForegroundColor Yellow"
    $workerLines += "                Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] Warte seit ' + `$elapsedMinutes + ' Min... Appliance nicht erreichbar (Reboot)')"
    $workerLines += "            }"
    $workerLines += "            "
    $workerLines += "            if (-not `$updateCompleted) {"
    $workerLines += "                Start-Sleep -Seconds `$pollInterval"
    $workerLines += "            }"
    $workerLines += "        }"
    $workerLines += "        "
    $workerLines += "        if (`$updateCompleted) {"
    $workerLines += "            Write-Host ('$WorkerName`: Starte Reconnect fuer Backup...') -ForegroundColor Cyan"
    $workerLines += "            `$connected = `$false"
    $workerLines += "            "
    $workerLines += "            for (`$attempt = 1; `$attempt -le 5; `$attempt++) {"
    $workerLines += "                try {"
    $workerLines += "                    Disconnect-OVMgmt -ErrorAction SilentlyContinue"
    $workerLines += "                    `$Connection = Connect-OVMgmt -Hostname `$Appliance -Credential `$credentials -ErrorAction Stop"
    $workerLines += "                    Write-Host '$WorkerName`: Reconnect erfolgreich!' -ForegroundColor Green"
    $workerLines += "                    `$connected = `$true"
    $workerLines += "                    break"
    $workerLines += "                } catch {"
    $workerLines += "                    Write-Host ('$WorkerName`: Reconnect-Versuch ' + `$attempt + '/5 fehlgeschlagen: ' + `$_.Exception.Message) -ForegroundColor Yellow"
    $workerLines += "                    Start-Sleep -Seconds 30"
    $workerLines += "                }"
    $workerLines += "            }"
    $workerLines += "            "
    $workerLines += "            if (`$connected) {"
    $workerLines += "                for (`$bAttempt = 1; `$bAttempt -le 3; `$bAttempt++) {"
    $workerLines += "                    try {"
    $workerLines += "                        Write-Host ('$WorkerName`: SCHRITT 3: Config Backup nach Update (Versuch ' + `$bAttempt + '/3)...') -ForegroundColor Yellow"
    $workerLines += '                        $backupDir = Join-Path $logDir "OneView_Backup"'
    $workerLines += '                        if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }'
    $workerLines += '                        new-OVBackup -Location $backupDir -Force -Passphrase $securePassphrase -ErrorAction Stop'
    $workerLines += "                        Write-Host '$WorkerName`: SCHRITT 3: Config Backup erfolgreich!' -ForegroundColor Green"
    $workerLines += "                        Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] SCHRITT 3: Config Backup nach Update erfolgreich')"
    $workerLines += "                        break"
    $workerLines += "                    } catch {"
    $workerLines += "                        Write-Host ('$WorkerName`: Backup fehlgeschlagen: ' + `$_.Exception.Message) -ForegroundColor Yellow"
    $workerLines += "                        Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] Backup-Versuch ' + `$bAttempt + ' fehlgeschlagen: ' + `$_.Exception.Message)"
    $workerLines += "                        if (`$bAttempt -lt 3) { Start-Sleep -Seconds 60 }"
    $workerLines += "                    }"
    $workerLines += "                }"
    $workerLines += "            } else {"
    $workerLines += "                Write-Host '$WorkerName`: Reconnect fuer Backup fehlgeschlagen - Backup uebersprungen' -ForegroundColor Red"
    $workerLines += "                Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] WARNUNG: Reconnect fuer Backup fehlgeschlagen')"
    $workerLines += "            }"
    $workerLines += "            "
    $workerLines += "            Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] SCHRITT 2: Install-OVUpdate abgeschlossen')"
    $workerLines += "            Write-Host '$WorkerName`: SCHRITT 2: Install-OVUpdate abgeschlossen' -ForegroundColor Green"
    $workerLines += "            Write-Host (`$workerDisplayName + ': Update erfolgreich fuer ' + `$Appliance) -ForegroundColor Green"
    $workerLines += "            Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date) + '] Update erfolgreich: ' + `$updateFilePath)"
    $workerLines += "        } else {"
    $workerLines += "            Write-Host ('$WorkerName`: Maximale Wartezeit (' + `$maxWaitMinutes + ' Min) erreicht!') -ForegroundColor Red"
    $workerLines += "            Write-Host '$WorkerName`: Bitte pruefen Sie den Update-Status in der OneView-Oberflaeche.' -ForegroundColor Yellow"
    $workerLines += "            Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] WARNUNG: Maximale Wartezeit erreicht - manueller Check erforderlich')"
    $workerLines += "        }"
    $workerLines += "        "
    $workerLines += "        Disconnect-OVMgmt"
    $workerLines += "        Write-Host ('$WorkerName`: Verbindung zu ' + `$Appliance + ' beendet.') -ForegroundColor Yellow"
    $workerLines += "    } catch {"
    $workerLines += "        Write-Host (`$workerDisplayName + ': Fehler waehrend des Updates fuer ' + `$Appliance + ' : ' + `$_) -ForegroundColor Red"
    $workerLines += "        Add-Content -Path `$logFile -Value ('[$WorkerName ' + (Get-Date) + '] Fehler waehrend des Updates: ' + `$_)"
    $workerLines += "        Disconnect-OVMgmt"
    $workerLines += "    }"
    $workerLines += "}"
    $workerLines += ""
    $workerLines += "Write-Host ' '"
    $workerLines += "Write-Host (`$workerDisplayName + ': Alle Updates abgeschlossen. Logs verfuegbar unter: ' + `$logFile) -ForegroundColor Green"
    $workerLines += "Write-Host ' '"
    $workerLines += "'DONE' | Out-File -FilePath (Join-Path `$logDir '${WorkerName}_Complete.txt') -Force"

    try {
        $workerLines | Out-File -FilePath $WorkerScriptPath -Encoding UTF8 -ErrorAction Stop
        
        if (-not (Test-Path $WorkerScriptPath)) {
            throw "Worker-Skript wurde nicht erstellt: $WorkerScriptPath"
        }
        
        $null = Get-Content $WorkerScriptPath -TotalCount 1 -ErrorAction Stop
        
        Write-Host "Worker-Skript erfolgreich erstellt: $WorkerScriptPath" -ForegroundColor Green
        Write-Host "  Skriptgröße: $((Get-Item $WorkerScriptPath).Length) Bytes" -ForegroundColor Gray
        
    } catch {
        throw "Fehler beim Erstellen des Worker-Skripts '$WorkerScriptPath': $($_.Exception.Message)"
    }
}

# Funktion, die den Update-Prozess für jede Appliance durchführt (sequenziell).
function Invoke-ApplianceUpdate {
    param(
        [string[]]$ApplianceIPs,
        $desiredVersion,
        [pscredential]$credentials,
        [string]$updateFilePath,
        [string]$logDir = (Join-Path $global:scriptFolder "OneView_Update")
    )
    
    $logFile = Join-Path $logDir "OneView_Update_Log.txt"
    if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    if (Test-Path $logFile) {
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $archivedLogFile = Join-Path $logDir ("OneView_Update_Log_$timestamp.txt")
        Rename-Item -Path $logFile -NewName $archivedLogFile -Force
        Write-Host "Alte Log-Datei archiviert unter: $archivedLogFile" -ForegroundColor Yellow
    }
    New-Item -ItemType File -Path $logFile -Force | Out-Null
    Write-Host "Neue Log-Datei erstellt: $logFile" -ForegroundColor Yellow

    foreach ($Appliance in $ApplianceIPs) {
        # API-Version automatisch ermitteln
        $apiVersion = Get-ApiVersionInline -Hostname $Appliance
        Write-Host "API-Version fuer ${Appliance}: $apiVersion" -ForegroundColor Gray

        $headers = @{
            "X-API-Version" = $apiVersion.ToString()
            "Accept" = "application/json"
            "Authorization" = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($credentials.UserName)`:$($credentials.GetNetworkCredential().Password)"))
        }
    
        $uri = "https://$Appliance/rest/appliance/nodeinfo/version"
        try {
            if ($PSVersionTable.PSVersion.Major -ge 7) {
                $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -TimeoutSec 600 -SkipCertificateCheck -ErrorAction Stop
            } else {
                $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -TimeoutSec 600 -ErrorAction Stop
            }
        } catch {
            Write-Host "Fehler beim Abrufen der Version von ${Appliance}: $_" -ForegroundColor Red
            Add-Content -Path $logFile -Value "[$(Get-Date)] Fehler bei API-Abfrage: $_"
            continue
        }

        $currentVersionRaw = $response.softwareversion
        $currentVersion = ConvertTo-NormalizedVersion $currentVersionRaw
        Write-Host "Ermittelte Version für ${Appliance}: $currentVersion" -ForegroundColor Yellow
        Add-Content -Path $logFile -Value ("[{0}] HPE OneView Version: {1}" -f (Get-Date), $currentVersion)

        try {
            $desiredVersionObj = ConvertTo-NormalizedVersion $desiredVersion.ToString()
            $comparison = $currentVersion.CompareTo($desiredVersionObj)
            
            if ($comparison -lt 0) {
                Write-Host "Update erforderlich für ${Appliance}: Aktuelle Version ($currentVersion) ist älter als gewünschte Version ($desiredVersionObj)." -ForegroundColor Yellow
            } elseif ($comparison -eq 0) {
                Write-Host "Kein Update erforderlich für ${Appliance}: Versionen sind identisch ($currentVersion)." -ForegroundColor Green
                continue
            } else {
                Write-Host "Keine Aktion erforderlich: Aktuelle Version ($currentVersion) ist neuer als die gewünschte Version ($desiredVersionObj)." -ForegroundColor Green
                continue
            }
        } catch {
            Write-Host "Fehler beim Versionsvergleich: $_" -ForegroundColor Red
            Add-Content -Path $logFile -Value ("[{0}] Fehler beim Versionsvergleich: {1}" -f (Get-Date), $_)
            continue
        }

        Write-Host "Verbindung zu ${Appliance} wird hergestellt..." -ForegroundColor Yellow
        $Connection = Connect-OVMgmt -Hostname $Appliance -Credential $credentials
        if (-not $Connection) {
            Disconnect-OVMgmt
            Write-Host "Fehler: Verbindung zu ${Appliance} konnte nicht hergestellt werden!" -ForegroundColor Red
            Write-Host "Verbindung zu ${Appliance} beendet." 
            Add-Content -Path $logFile -Value "[$(Get-Date)] Verbindung zu ${Appliance} fehlgeschlagen!"
            continue
        }
    
        try {
            try {
                $pendingUpdate = Get-OVPendingUpdate -ErrorAction Stop
            } catch {
                Write-Host "Fehler beim Abrufen des Pending Updates: $($_.Exception.Message)" -ForegroundColor Red
                $pendingUpdate = $null
            }
            if ($pendingUpdate) {
                Write-Host "Ein Pending Update wurde gefunden:" -ForegroundColor Red
                Write-Host "  Dateiname    : $($pendingUpdate.FileName)" -ForegroundColor Yellow
                Write-Host "  Update-Version: $($pendingUpdate.Version)" -ForegroundColor Yellow
                try {
                    Remove-OVPendingUpdate -Confirm:$false -ErrorAction Stop
                    Write-Host "Das Pending Update '$($pendingUpdate.FileName)' wurde erfolgreich entfernt." -ForegroundColor Yellow
                } catch {
                    Write-Host "Fehler beim Entfernen des Pending Updates '$($pendingUpdate.FileName)': $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "Kein Pending Update gefunden. Das Update kann fortgesetzt werden." -ForegroundColor Yellow
            }
    
            Write-Host "Update wird mit Datei $updateFilePath durchgeführt...." -ForegroundColor Yellow
            Write-Host " "
            Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] SCHRITT 1: Starte Config Backup vor dem Update für ${Appliance}..."
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] SCHRITT 1: Starte Config Backup vor dem Update..." -ForegroundColor Yellow
            Write-Host " "
            $backupDir = Join-Path $logDir "OneView_Backup"
            if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
            new-OVBackup -Location $backupDir -Force -Passphrase $securePassphrase
            Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] SCHRITT 1: Config Backup abgeschlossen"
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] SCHRITT 1: Config Backup abgeschlossen" -ForegroundColor Green
            $host.UI.RawUI.BackgroundColor = "DarkBlue"
            Write-Host " "
            Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] SCHRITT 2: Starte Install-OVUpdate für ${Appliance}..."
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] SCHRITT 2: Starte Install-OVUpdate..." -ForegroundColor Yellow
            Write-Host " "
            
            $installOVUpdateCompleted = $false
            try {
                Install-OVUpdate -file $updateFilePath -Eula Accept -Confirm:$false
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Install-OVUpdate Befehl abgeschlossen" -ForegroundColor Green
                $installOVUpdateCompleted = $true
            } catch {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Install-OVUpdate unterbrochen: $($_.Exception.Message)" -ForegroundColor Yellow
                Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Install-OVUpdate unterbrochen - Update läuft im Hintergrund weiter"
            }
            
            Disconnect-OVMgmt -ErrorAction SilentlyContinue
            
            Write-Host " "
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Warte auf Update-Abschluss (Zielversion: $($desiredVersionObj.ToString()))..." -ForegroundColor Cyan
            Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Warte auf Update-Abschluss - Zielversion: $($desiredVersionObj.ToString())"
            
            $maxWaitMinutes = 120
            $pollInterval = 60
            $startTime = Get-Date
            $updateCompleted = $false
            $versionHeaders = @{
                "X-API-Version" = $apiVersion.ToString()
                "Accept" = "application/json"
                "Authorization" = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($credentials.UserName)`:$($credentials.GetNetworkCredential().Password)"))
            }
            $versionUri = "https://$Appliance/rest/appliance/nodeinfo/version"
            
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Warte 2 Minuten bevor erster Versions-Check..." -ForegroundColor Cyan
            Start-Sleep -Seconds 120
            
            while (-not $updateCompleted -and ((Get-Date) - $startTime).TotalMinutes -lt $maxWaitMinutes) {
                $elapsedMinutes = [math]::Round(((Get-Date) - $startTime).TotalMinutes, 1)
                try {
                    if ($PSVersionTable.PSVersion.Major -ge 7) {
                        $versionResponse = Invoke-RestMethod -Uri $versionUri -Headers $versionHeaders -Method Get -TimeoutSec 15 -SkipCertificateCheck -ErrorAction Stop
                    } else {
                        $versionResponse = Invoke-RestMethod -Uri $versionUri -Headers $versionHeaders -Method Get -TimeoutSec 15 -ErrorAction Stop
                    }
                    $reportedVersion = ConvertTo-NormalizedVersion $versionResponse.softwareversion
                    $versionComparison = $reportedVersion.CompareTo($desiredVersionObj)
                    
                    if ($versionComparison -ge 0) {
                        $updateCompleted = $true
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Appliance meldet Version $($reportedVersion.ToString()) - Update abgeschlossen! (nach $elapsedMinutes Min)" -ForegroundColor Green
                        Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Update abgeschlossen - Neue Version: $($reportedVersion.ToString()) (nach $elapsedMinutes Min)"
                    } else {
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Warte seit $elapsedMinutes Min... Aktuelle Version: $($reportedVersion.ToString()) (Ziel: $($desiredVersionObj.ToString()))" -ForegroundColor Cyan
                    }
                } catch {
                    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Warte seit $elapsedMinutes Min... Appliance nicht erreichbar (Reboot läuft)" -ForegroundColor Yellow
                }
                
                if (-not $updateCompleted) {
                    Start-Sleep -Seconds $pollInterval
                }
            }
            
            if ($updateCompleted) {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Starte Reconnect für Backup..." -ForegroundColor Cyan
                $connected = $false
                
                for ($attempt = 1; $attempt -le 5; $attempt++) {
                    try {
                        Disconnect-OVMgmt -ErrorAction SilentlyContinue
                        $Connection = Connect-OVMgmt -Hostname $Appliance -Credential $credentials -ErrorAction Stop
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Reconnect erfolgreich!" -ForegroundColor Green
                        $connected = $true
                        break
                    } catch {
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Reconnect-Versuch $attempt/5 fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor Yellow
                        Start-Sleep -Seconds 30
                    }
                }
                
                if ($connected) {
                    for ($bAttempt = 1; $bAttempt -le 3; $bAttempt++) {
                        try {
                            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] SCHRITT 3: Config Backup nach Update (Versuch $bAttempt/3)..." -ForegroundColor Yellow
                            $backupDir = Join-Path $logDir "OneView_Backup"
                            if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
                            new-OVBackup -Location $backupDir -Force -Passphrase $securePassphrase -ErrorAction Stop
                            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] SCHRITT 3: Config Backup erfolgreich!" -ForegroundColor Green
                            Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] SCHRITT 3: Config Backup nach Update erfolgreich"
                            break
                        } catch {
                            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Backup fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor Yellow
                            Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Backup-Versuch $bAttempt fehlgeschlagen: $($_.Exception.Message)"
                            if ($bAttempt -lt 3) { Start-Sleep -Seconds 60 }
                        }
                    }
                } else {
                    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Reconnect für Backup fehlgeschlagen - Backup übersprungen" -ForegroundColor Red
                    Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARNUNG: Reconnect für Backup fehlgeschlagen"
                }
                
                Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] SCHRITT 2: Install-OVUpdate abgeschlossen"
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] SCHRITT 2: Install-OVUpdate abgeschlossen" -ForegroundColor Green
                Write-Host " "
                Write-Host "Update erfolgreich für ${Appliance} mit Datei $updateFilePath" -ForegroundColor Green
                Add-Content -Path $logFile -Value "[$(Get-Date)] Update erfolgreich: ${updateFilePath}"
            } else {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Maximale Wartezeit ($maxWaitMinutes Min) erreicht!" -ForegroundColor Red
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Bitte prüfen Sie den Update-Status in der OneView-Oberfläche." -ForegroundColor Yellow
                Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARNUNG: Maximale Wartezeit erreicht - manueller Check erforderlich"
            }
            
            $host.UI.RawUI.BackgroundColor = "DarkBlue"
            Disconnect-OVMgmt -ErrorAction SilentlyContinue
            Write-Host " "
            Write-Host "Verbindung zu ${Appliance} beendet." -ForegroundColor Yellow
        } catch {
            Write-Host "Fehler während des Updates für ${Appliance}: $_" -ForegroundColor Red
            Add-Content -Path $logFile -Value "[$(Get-Date)] Fehler während des Updates: $_"
            Write-Host "Verbindung zu ${Appliance} beendet." -ForegroundColor Yellow
            Disconnect-OVMgmt -ErrorAction SilentlyContinue
        }
    }
    Write-Host " "
    Write-Host "Skript abgeschlossen. Logs verfügbar unter: $logFile" -ForegroundColor Yellow
    Write-Host " "
    $host.UI.RawUI.ForegroundColor = "Green"
    Read-Host -Prompt "Drücke eine beliebige Taste, um fortzufahren..."
}

# ============================================
# Cleanup-Funktion
# ============================================
function Cleanup-Resources {
    Write-Host "Räume Ressourcen auf..." -ForegroundColor Yellow
    
    if ($global:runspacePool) {
        try {
            $global:runspacePool.Close()
            $global:runspacePool.Dispose()
            Write-Host "✅ Runspace Pool geschlossen" -ForegroundColor Green
        } catch {
            Write-Host "⚠️ Fehler beim Schließen des Runspace Pools: $_" -ForegroundColor Yellow
        }
    }
    
    Get-Module -Name "HPEOneView*" | Remove-Module -Force -ErrorAction SilentlyContinue
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    
    Write-Host "✅ Cleanup abgeschlossen" -ForegroundColor Green
}

Register-EngineEvent PowerShell.Exiting -Action { 
    if (Get-Command -Name Cleanup-Resources -ErrorAction SilentlyContinue) { 
        Cleanup-Resources 
    } 
} -SupportEvent

trap {
    Write-Host "`n⚠️ Script wurde unterbrochen!" -ForegroundColor Yellow
    Write-Host "Fehler: $_" -ForegroundColor Red
    Write-Host "Zeile: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    Cleanup-Resources
    Read-Host "Drücke Enter zum Beenden..."
    exit 1
}

# ============================================
# UNIFIED GUI - Alles in einem Fenster
# ============================================
function Show-UnifiedGUI {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Initialize-NativeAPIs

    # --- Farben und Fonts ---
    $colorHeader     = [System.Drawing.Color]::FromArgb(0, 120, 212)   # Blau
    $colorDarkBg     = [System.Drawing.Color]::FromArgb(30, 30, 30)    # Dunkler Hintergrund
    $colorPanelBg    = [System.Drawing.Color]::FromArgb(45, 45, 45)    # Panel-Hintergrund
    $colorWorker1    = [System.Drawing.Color]::FromArgb(76, 175, 80)   # Grün
    $colorWorker2    = [System.Drawing.Color]::FromArgb(206, 147, 216) # Magenta/Lila
    $colorWorker3    = [System.Drawing.Color]::FromArgb(77, 208, 225)  # Cyan
    $colorWorker4    = [System.Drawing.Color]::FromArgb(255, 213, 79)  # Gelb
    $colorWhite      = [System.Drawing.Color]::White
    $colorLightGray  = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $colorWarning    = [System.Drawing.Color]::FromArgb(255, 152, 0)   # Orange
    $colorSuccess    = [System.Drawing.Color]::FromArgb(76, 175, 80)   # Grün
    $colorError      = [System.Drawing.Color]::FromArgb(244, 67, 54)   # Rot

    $fontTitle       = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $fontSection     = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $fontNormal      = New-Object System.Drawing.Font("Segoe UI", 9)
    $fontSmall       = New-Object System.Drawing.Font("Segoe UI", 8)
    $fontConsole     = New-Object System.Drawing.Font("Consolas", 9)
    $fontWorkerTitle = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    $workerColors = @($colorWorker1, $colorWorker2, $colorWorker3, $colorWorker4)
    $workerConsoleColors = @("Green", "Magenta", "Cyan", "Yellow")

    # --- Hauptformular ---
    $scaled = Get-ScaledFormSize -DesiredWidth 1400 -DesiredHeight 900 -MinWidth 1000 -MinHeight 700
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "© 2025 N.J. Airbus D&S - HPE Synergy Update Tool"
    $form.Size = New-Object System.Drawing.Size($scaled.Width, $scaled.Height)
    # MinimumSize an Bildschirmgröße anpassen (kleine Notebooks berücksichtigen)
    $minFormW = [Math]::Min(800, [int]($scaled.ScreenWidth * 0.7))
    $minFormH = [Math]::Min(550, [int]($scaled.ScreenHeight * 0.7))
    $minFormW = [Math]::Max($minFormW, 650)
    $minFormH = [Math]::Max($minFormH, 450)
    $form.MinimumSize = New-Object System.Drawing.Size($minFormW, $minFormH)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.BackColor = $colorDarkBg
    $form.ForeColor = $colorWhite
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable

    # ========================================
    # HEADER
    # ========================================
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $headerPanel.Height = 55
    $headerPanel.BackColor = $colorHeader

    $headerLabel = New-Object System.Windows.Forms.Label
    $headerLabel.Text = "HPE OneView / Synergy Update Tool"
    $headerLabel.Font = $fontTitle
    $headerLabel.ForeColor = $colorWhite
    $headerLabel.AutoSize = $true
    $headerLabel.Location = New-Object System.Drawing.Point(20, 15)
    $headerPanel.Controls.Add($headerLabel)

    $headerVersion = New-Object System.Windows.Forms.Label
    $headerVersion.Text = "v2.0 - Unified Dashboard"
    $headerVersion.Font = $fontSmall
    $headerVersion.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 255)
    $headerVersion.AutoSize = $true
    $headerVersion.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $headerVersion.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 180), 20)
    $headerPanel.Controls.Add($headerVersion)

    # ========================================
    # STATUS-BAR (unten)
    # ========================================
    $statusBar = New-Object System.Windows.Forms.StatusStrip
    $statusBar.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
    $statusItem = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusItem.Text = "Bereit - Bitte Einstellungen konfigurieren und Update starten"
    $statusItem.ForeColor = $colorLightGray
    $statusBar.Items.Add($statusItem) | Out-Null

    # ========================================
    # HAUPT-SPLITCONTAINER (Links: Settings / Rechts: Worker-Output)
    # ========================================
    $mainSplit = New-Object System.Windows.Forms.SplitContainer
    $mainSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainSplit.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $mainSplit.SplitterWidth = 5
    $mainSplit.FixedPanel = [System.Windows.Forms.FixedPanel]::Panel1
    $mainSplit.BackColor = $colorDarkBg

    # ========================================
    # LINKE SEITE: SETTINGS PANEL
    # ========================================
    $settingsPanel = New-Object System.Windows.Forms.Panel
    $settingsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $settingsPanel.AutoScroll = $true
    $settingsPanel.BackColor = $colorPanelBg
    $settingsPanel.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 10)
    # Realistische Anfangsgröße setzen, damit Anchor-Berechnungen korrekt sind
    $settingsPanel.Size = New-Object System.Drawing.Size(380, 800)

    $yPos = 10

    # --- Anmeldedaten ---
    $loginBox = New-Object System.Windows.Forms.GroupBox
    $loginBox.Text = "🔑 Anmeldedaten"
    $loginBox.ForeColor = $colorWhite
    $loginBox.Font = $fontSection
    $loginBox.Location = New-Object System.Drawing.Point(10, $yPos)
    $loginBox.Size = New-Object System.Drawing.Size(($settingsPanel.ClientSize.Width - 30), 135)
    $loginBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    $userNameLbl = New-Object System.Windows.Forms.Label
    $userNameLbl.Text = "Benutzername:"
    $userNameLbl.Font = $fontNormal
    $userNameLbl.ForeColor = $colorLightGray
    $userNameLbl.Location = New-Object System.Drawing.Point(10, 25)
    $userNameLbl.AutoSize = $true
    $loginBox.Controls.Add($userNameLbl)

    $userNameTextBox = New-Object System.Windows.Forms.TextBox
    $userNameTextBox.Font = $fontNormal
    $userNameTextBox.Location = New-Object System.Drawing.Point(120, 22)
    $userNameTextBox.Size = New-Object System.Drawing.Size(($loginBox.ClientSize.Width - 135), 24)
    $userNameTextBox.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $userNameTextBox.ForeColor = $colorWhite
    $userNameTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $loginBox.Controls.Add($userNameTextBox)

    $passwordLbl = New-Object System.Windows.Forms.Label
    $passwordLbl.Text = "Passwort:"
    $passwordLbl.Font = $fontNormal
    $passwordLbl.ForeColor = $colorLightGray
    $passwordLbl.Location = New-Object System.Drawing.Point(10, 60)
    $passwordLbl.AutoSize = $true
    $loginBox.Controls.Add($passwordLbl)

    $passwordTextBox = New-Object System.Windows.Forms.TextBox
    $passwordTextBox.Font = $fontNormal
    $passwordTextBox.Location = New-Object System.Drawing.Point(120, 57)
    $passwordTextBox.Size = New-Object System.Drawing.Size(($loginBox.ClientSize.Width - 135), 24)
    $passwordTextBox.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $passwordTextBox.ForeColor = $colorWhite
    $passwordTextBox.UseSystemPasswordChar = $true
    $passwordTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $loginBox.Controls.Add($passwordTextBox)

    $backupPassLbl = New-Object System.Windows.Forms.Label
    $backupPassLbl.Text = "Backup-Passw.:"
    $backupPassLbl.Font = $fontNormal
    $backupPassLbl.ForeColor = $colorLightGray
    $backupPassLbl.Location = New-Object System.Drawing.Point(10, 95)
    $backupPassLbl.AutoSize = $true
    $loginBox.Controls.Add($backupPassLbl)

    $backupPassTextBox = New-Object System.Windows.Forms.TextBox
    $backupPassTextBox.Font = $fontNormal
    $backupPassTextBox.Location = New-Object System.Drawing.Point(120, 92)
    $backupPassTextBox.Size = New-Object System.Drawing.Size(($loginBox.ClientSize.Width - 135), 24)
    $backupPassTextBox.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $backupPassTextBox.ForeColor = $colorWhite
    $backupPassTextBox.UseSystemPasswordChar = $true
    $backupPassTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $loginBox.Controls.Add($backupPassTextBox)

    $settingsPanel.Controls.Add($loginBox)
    $yPos += 145

    # --- Update-Datei & Version ---
    $fileBox = New-Object System.Windows.Forms.GroupBox
    $fileBox.Text = "📁 Update-Datei & Version"
    $fileBox.ForeColor = $colorWhite
    $fileBox.Font = $fontSection
    $fileBox.Location = New-Object System.Drawing.Point(10, $yPos)
    $fileBox.Size = New-Object System.Drawing.Size(($settingsPanel.ClientSize.Width - 30), 145)
    $fileBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    $selectFileBtn = New-Object System.Windows.Forms.Button
    $selectFileBtn.Text = "Datei auswählen..."
    $selectFileBtn.Font = $fontNormal
    $selectFileBtn.Location = New-Object System.Drawing.Point(10, 25)
    $selectFileBtn.Size = New-Object System.Drawing.Size(150, 28)
    $selectFileBtn.BackColor = $colorHeader
    $selectFileBtn.ForeColor = $colorWhite
    $selectFileBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $fileBox.Controls.Add($selectFileBtn)

    $filePathLabel = New-Object System.Windows.Forms.TextBox
    $filePathLabel.Font = $fontSmall
    $filePathLabel.Location = New-Object System.Drawing.Point(10, 58)
    $filePathLabel.Size = New-Object System.Drawing.Size(($fileBox.ClientSize.Width - 20), 22)
    $filePathLabel.ReadOnly = $true
    $filePathLabel.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $filePathLabel.ForeColor = $colorLightGray
    $filePathLabel.Text = "(Keine Datei ausgewählt)"
    $filePathLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $fileBox.Controls.Add($filePathLabel)

    $versionLbl = New-Object System.Windows.Forms.Label
    $versionLbl.Text = "Zielversion (xX.YY.ZZ):"
    $versionLbl.Font = $fontNormal
    $versionLbl.ForeColor = $colorLightGray
    $versionLbl.Location = New-Object System.Drawing.Point(10, 90)
    $versionLbl.AutoSize = $true
    $fileBox.Controls.Add($versionLbl)

    $versionTextBox = New-Object System.Windows.Forms.TextBox
    $versionTextBox.Font = New-Object System.Drawing.Font("Consolas", 11)
    $versionTextBox.Location = New-Object System.Drawing.Point(175, 87)
    $versionTextBox.Size = New-Object System.Drawing.Size(153, 26)
    $versionTextBox.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $versionTextBox.ForeColor = $colorWhite
    $versionTextBox.MaxLength = 8
    $versionTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $fileBox.Controls.Add($versionTextBox)

    $fileValidationLabel = New-Object System.Windows.Forms.Label
    $fileValidationLabel.Text = ""
    $fileValidationLabel.Font = $fontSmall
    $fileValidationLabel.ForeColor = $colorError
    $fileValidationLabel.Location = New-Object System.Drawing.Point(10, 118)
    $fileValidationLabel.Size = New-Object System.Drawing.Size(($fileBox.ClientSize.Width - 20), 20)
    $fileValidationLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $fileBox.Controls.Add($fileValidationLabel)

    $settingsPanel.Controls.Add($fileBox)
    $yPos += 155

    # --- Appliance-Auswahl ---
    $applianceBox = New-Object System.Windows.Forms.GroupBox
    $applianceBox.Text = "🖥 Appliance-Auswahl"
    $applianceBox.ForeColor = $colorWhite
    $applianceBox.Font = $fontSection
    $applianceBox.Location = New-Object System.Drawing.Point(10, $yPos)
    $applianceBox.Size = New-Object System.Drawing.Size(($settingsPanel.ClientSize.Width - 30), 230)
    $applianceBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    $selectAllChk = New-Object System.Windows.Forms.CheckBox
    $selectAllChk.Text = "Alle auswählen"
    $selectAllChk.Font = $fontSmall
    $selectAllChk.ForeColor = $colorLightGray
    $selectAllChk.Location = New-Object System.Drawing.Point(10, 22)
    $selectAllChk.AutoSize = $true
    $applianceBox.Controls.Add($selectAllChk)

    $applianceListBox = New-Object System.Windows.Forms.CheckedListBox
    $applianceListBox.Font = $fontNormal
    $applianceListBox.Location = New-Object System.Drawing.Point(10, 45)
    $applianceListBox.Size = New-Object System.Drawing.Size(($applianceBox.ClientSize.Width - 20), 175)
    $applianceListBox.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $applianceListBox.ForeColor = $colorWhite
    $applianceListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $applianceListBox.CheckOnClick = $true
    $applianceListBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    # Appliances aus Datei laden
    $oneviewFile = Join-Path $global:scriptFolder "oneview_upd.txt"
    if (Test-Path $oneviewFile) {
        $instances = Get-Content -Path $oneviewFile -ErrorAction SilentlyContinue | Where-Object { $_.Trim() -ne "" }
        foreach ($inst in $instances) {
            $applianceListBox.Items.Add($inst.Trim()) | Out-Null
        }
    }

    $applianceBox.Controls.Add($applianceListBox)
    $settingsPanel.Controls.Add($applianceBox)
    $yPos += 240

    # "Alle auswählen" Checkbox Event
    $selectAllChk.Add_CheckedChanged({
        for ($i = 0; $i -lt $applianceListBox.Items.Count; $i++) {
            $applianceListBox.SetItemChecked($i, $selectAllChk.Checked)
        }
    })

    # --- Update-Modus ---
    $modeBox = New-Object System.Windows.Forms.GroupBox
    $modeBox.Text = "⚙ Update-Modus"
    $modeBox.ForeColor = $colorWhite
    $modeBox.Font = $fontSection
    $modeBox.Location = New-Object System.Drawing.Point(10, $yPos)
    $modeBox.Size = New-Object System.Drawing.Size(($settingsPanel.ClientSize.Width - 30), 110)
    $modeBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    $sequentialRadio = New-Object System.Windows.Forms.RadioButton
    $sequentialRadio.Text = "Sequenziell (nacheinander)"
    $sequentialRadio.Font = $fontNormal
    $sequentialRadio.ForeColor = $colorLightGray
    $sequentialRadio.Location = New-Object System.Drawing.Point(10, 25)
    $sequentialRadio.AutoSize = $true
    $sequentialRadio.Checked = $true
    $modeBox.Controls.Add($sequentialRadio)

    $parallelRadio = New-Object System.Windows.Forms.RadioButton
    $parallelRadio.Text = "Parallel (gleichzeitig)"
    $parallelRadio.Font = $fontNormal
    $parallelRadio.ForeColor = $colorLightGray
    $parallelRadio.Location = New-Object System.Drawing.Point(10, 50)
    $parallelRadio.AutoSize = $true
    $modeBox.Controls.Add($parallelRadio)

    $workerLbl = New-Object System.Windows.Forms.Label
    $workerLbl.Text = "Worker:"
    $workerLbl.Font = $fontNormal
    $workerLbl.ForeColor = $colorLightGray
    $workerLbl.Location = New-Object System.Drawing.Point(30, 78)
    $workerLbl.AutoSize = $true
    $modeBox.Controls.Add($workerLbl)

    $workerCombo = New-Object System.Windows.Forms.ComboBox
    $workerCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $workerCombo.Items.AddRange(@("2", "3", "4"))
    $workerCombo.SelectedIndex = 0
    $workerCombo.Font = $fontNormal
    $workerCombo.Location = New-Object System.Drawing.Point(85, 75)
    $workerCombo.Size = New-Object System.Drawing.Size(50, 24)
    $workerCombo.Enabled = $false
    $modeBox.Controls.Add($workerCombo)

    $workerHint = New-Object System.Windows.Forms.Label
    $workerHint.Text = "(gleichmäßig verteilt)"
    $workerHint.Font = $fontSmall
    $workerHint.ForeColor = [System.Drawing.Color]::Gray
    $workerHint.Location = New-Object System.Drawing.Point(142, 79)
    $workerHint.AutoSize = $true
    $modeBox.Controls.Add($workerHint)

    $parallelRadio.Add_CheckedChanged({
        $workerCombo.Enabled = $parallelRadio.Checked
    })

    $settingsPanel.Controls.Add($modeBox)
    $yPos += 120

    # --- Start Button ---
    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Text = "▶  Update starten"
    $startButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $startButton.Location = New-Object System.Drawing.Point(10, $yPos)
    $startButton.Size = New-Object System.Drawing.Size(($settingsPanel.ClientSize.Width - 30), 45)
    $startButton.BackColor = $colorSuccess
    $startButton.ForeColor = $colorWhite
    $startButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $startButton.FlatAppearance.BorderSize = 0
    $startButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $settingsPanel.Controls.Add($startButton)
    $yPos += 55

    # --- Abbruch Button ---
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "✕  Beenden"
    $cancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cancelButton.Location = New-Object System.Drawing.Point(10, $yPos)
    $cancelButton.Size = New-Object System.Drawing.Size(($settingsPanel.ClientSize.Width - 30), 35)
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $cancelButton.ForeColor = $colorLightGray
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelButton.FlatAppearance.BorderSize = 0
    $cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $settingsPanel.Controls.Add($cancelButton)

    $mainSplit.Panel1.Controls.Add($settingsPanel)

    # ========================================
    # RECHTE SEITE: WORKER OUTPUT PANELS
    # ========================================
    $rightPanel = New-Object System.Windows.Forms.Panel
    $rightPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $rightPanel.BackColor = $colorDarkBg
    $rightPanel.Padding = New-Object System.Windows.Forms.Padding(5)

    # TableLayoutPanel für Worker-Panels (2x2 Grid)
    $workerGrid = New-Object System.Windows.Forms.TableLayoutPanel
    $workerGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $workerGrid.ColumnCount = 2
    $workerGrid.RowCount = 2
    $workerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $workerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $workerGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $workerGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $workerGrid.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::Single
    $workerGrid.BackColor = $colorDarkBg
    $workerGrid.Padding = New-Object System.Windows.Forms.Padding(2)

    # 4 Worker-Panels erstellen
    $workerPanels = @()
    $workerHostPanels = @()
    $workerTitleLabels = @()
    $workerStatusLabels = @()
    
    for ($w = 0; $w -lt 4; $w++) {
        $wPanel = New-Object System.Windows.Forms.Panel
        $wPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
        $wPanel.BackColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
        $wPanel.Margin = New-Object System.Windows.Forms.Padding(3)

        # Worker-Titel
        $wTitlePanel = New-Object System.Windows.Forms.Panel
        $wTitlePanel.Dock = [System.Windows.Forms.DockStyle]::Top
        $wTitlePanel.Height = 28
        $wTitlePanel.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)

        $wTitle = New-Object System.Windows.Forms.Label
        $wTitle.Text = "Worker $($w + 1)"
        $wTitle.Font = $fontWorkerTitle
        $wTitle.ForeColor = $workerColors[$w]
        $wTitle.AutoSize = $true
        $wTitle.Location = New-Object System.Drawing.Point(8, 4)
        $wTitlePanel.Controls.Add($wTitle)

        $wStatus = New-Object System.Windows.Forms.Label
        $wStatus.Text = "Wartend"
        $wStatus.Font = $fontSmall
        $wStatus.ForeColor = [System.Drawing.Color]::Gray
        $wStatus.AutoSize = $true
        $wStatus.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        $wStatus.Location = New-Object System.Drawing.Point(200, 7)
        $wTitlePanel.Controls.Add($wStatus)

        # Host-Panel für eingebettetes Konsolenfenster
        $wHostPanel = New-Object System.Windows.Forms.Panel
        $wHostPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
        $wHostPanel.BackColor = [System.Drawing.Color]::FromArgb(12, 12, 12)

        # Fill zuerst, dann Top (WinForms dockt zuletzt hinzugefügtes Control zuerst)
        $wPanel.Controls.Add($wHostPanel)
        $wPanel.Controls.Add($wTitlePanel)

        $col = $w % 2
        $row = [math]::Floor($w / 2)
        $workerGrid.Controls.Add($wPanel, $col, $row)

        $workerPanels += $wPanel
        $workerHostPanels += $wHostPanel
        $workerTitleLabels += $wTitle
        $workerStatusLabels += $wStatus
    }

    $rightPanel.Controls.Add($workerGrid)
    $mainSplit.Panel2.Controls.Add($rightPanel)

    # Controls in korrekter Dock-Reihenfolge hinzufügen: Fill zuerst, dann Bottom, dann Top
    # (WinForms layoutet das zuletzt hinzugefügte Dock-Control zuerst)
    $form.Controls.Add($mainSplit)
    $form.Controls.Add($statusBar)
    $form.Controls.Add($headerPanel)

    # MinSize und SplitterDistance dynamisch berechnen (kleine Monitore berücksichtigen)
    $availableWidth = $form.ClientSize.Width
    $p1Min = [Math]::Min(380, [int]($availableWidth * 0.30))
    $p1Min = [Math]::Max($p1Min, 300)
    $p2Min = [Math]::Min(400, [int]($availableWidth * 0.40))
    $p2Min = [Math]::Max($p2Min, 250)
    # Sicherstellen, dass Panel1Min + Panel2Min + SplitterWidth <= verfügbare Breite
    if (($p1Min + $p2Min + $mainSplit.SplitterWidth) -gt $availableWidth) {
        $halfAvail = [int](($availableWidth - $mainSplit.SplitterWidth) / 2)
        $p1Min = [int]($halfAvail * 0.45)
        $p2Min = [int]($halfAvail * 0.55)
    }
    $mainSplit.Panel1MinSize = $p1Min
    $mainSplit.Panel2MinSize = $p2Min
    $splitterDist = [Math]::Min(420, [int]($availableWidth * 0.38))
    $splitterDist = [Math]::Max($splitterDist, $p1Min)
    $splitterDist = [Math]::Min($splitterDist, ($availableWidth - $p2Min - $mainSplit.SplitterWidth))
    $mainSplit.SplitterDistance = $splitterDist

    # ========================================
    # EVENT HANDLERS
    # ========================================

    # Datei auswählen
    $selectFileBtn.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = "Synergy Update-Datei auswählen"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $filePathLabel.Text = $openFileDialog.FileName
            $fileValidationLabel.Text = ""
            $fileValidationLabel.ForeColor = $colorSuccess
        }
    })

    # Beenden
    $cancelButton.Add_Click({
        $form.Close()
    })

    # ========================================
    # TIMER für Console-Embedding und Status
    # ========================================
    $global:guiWorkerProcesses = @($null, $null, $null, $null)
    $global:guiWorkerEmbedded = @($false, $false, $false, $false)
    $global:guiWorkerHandles = @([IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero)
    $global:guiLastFontHeight = @(0, 0, 0, 0)
    $global:guiLastPanelSize = @($null, $null, $null, $null)
    $global:guiWorkerNames = @($null, $null, $null, $null)
    $global:guiUpdateRunning = $false

    $logTimer = New-Object System.Windows.Forms.Timer
    $logTimer.Interval = 500

    $logTimer.Add_Tick({
        if (-not $global:guiUpdateRunning) { return }
        
        for ($i = 0; $i -lt 4; $i++) {
            $proc = $global:guiWorkerProcesses[$i]
            if ($null -eq $proc) { continue }
            
            # Konsolenfenster einbetten, wenn noch nicht geschehen
            if (-not $global:guiWorkerEmbedded[$i]) {
                try {
                    # Window-Handle aus Datei lesen (Worker schreibt Handle beim Start)
                    $hWnd = [IntPtr]::Zero
                    $wn = $global:guiWorkerNames[$i]
                    if ($wn -and -not $proc.HasExited) {
                        $handlePath = Join-Path (Join-Path $global:scriptFolder "OneView_Update") "${wn}_Handle.txt"
                        if (Test-Path $handlePath) {
                            try {
                                $handleValue = [long](Get-Content $handlePath -Raw).Trim()
                                if ($handleValue -ne 0) {
                                    $hWnd = [IntPtr]::new($handleValue)
                                }
                            } catch { }
                        }
                    }
                    if ($hWnd -ne [IntPtr]::Zero) {
                        $hostPanel = $workerHostPanels[$i]
                        $hostHandle = $hostPanel.Handle
                        
                        # Window-Style ändern: Titelleiste und Rahmen entfernen, als Child setzen
                        $style = [NativeMethods]::GetWindowLong($hWnd, [NativeMethods]::GWL_STYLE)
                        $style = $style -band (-bnot [NativeMethods]::WS_CAPTION)
                        $style = $style -band (-bnot [NativeMethods]::WS_THICKFRAME)
                        $style = $style -band (-bnot [NativeMethods]::WS_BORDER)
                        $style = $style -band (-bnot [NativeMethods]::WS_SYSMENU)
                        $style = $style -bor [NativeMethods]::WS_CHILD
                        [NativeMethods]::SetWindowLong($hWnd, [NativeMethods]::GWL_STYLE, $style) | Out-Null
                        
                        # Extended Style: 3D-Rand entfernen damit Konsole das Panel komplett ausfüllt
                        $exStyle = [NativeMethods]::GetWindowLong($hWnd, [NativeMethods]::GWL_EXSTYLE)
                        $exStyle = $exStyle -band (-bnot [NativeMethods]::WS_EX_CLIENTEDGE)
                        $exStyle = $exStyle -band (-bnot [NativeMethods]::WS_EX_WINDOWEDGE)
                        $exStyle = $exStyle -band (-bnot [NativeMethods]::WS_EX_DLGMODALFRAME)
                        $exStyle = $exStyle -band (-bnot [NativeMethods]::WS_EX_STATICEDGE)
                        [NativeMethods]::SetWindowLong($hWnd, [NativeMethods]::GWL_EXSTYLE, $exStyle) | Out-Null
                        
                        # Fenster in das Host-Panel einbetten
                        [NativeMethods]::SetParent($hWnd, $hostHandle) | Out-Null
                        
                        # Fenster auf Panel-Größe setzen
                        [NativeMethods]::MoveWindow($hWnd, 0, 0, $hostPanel.ClientSize.Width, $hostPanel.ClientSize.Height, $true) | Out-Null
                        [NativeMethods]::ShowWindow($hWnd, [NativeMethods]::SW_SHOW) | Out-Null
                        
                        $global:guiWorkerEmbedded[$i] = $true
                        $global:guiWorkerHandles[$i] = $hWnd
                        
                        # Initiale Schriftgrösse an Panel-Höhe anpassen
                        $desiredH = [math]::Max(8, [math]::Min(22, [int][math]::Floor($hostPanel.ClientSize.Height / 25)))
                        $wn2 = $global:guiWorkerNames[$i]
                        if ($wn2) {
                            $fontPath = Join-Path (Join-Path $global:scriptFolder "OneView_Update") "${wn2}_FontSize.txt"
                            $desiredH | Out-File -FilePath $fontPath -Force
                            $global:guiLastFontHeight[$i] = $desiredH
                        }
                        
                        # Resize-Handler für das Host-Panel registrieren
                        $idx = $i
                        $hostPanel.Add_Resize({
                            try {
                                $myIdx = $idx
                                $myHwnd = $global:guiWorkerHandles[$myIdx]
                                if ($myHwnd -ne [IntPtr]::Zero -and $global:guiWorkerEmbedded[$myIdx]) {
                                    $hp = $workerHostPanels[$myIdx]
                                    if ($hp.ClientSize.Width -gt 0 -and $hp.ClientSize.Height -gt 0) {
                                        [NativeMethods]::MoveWindow($myHwnd, 0, 0, $hp.ClientSize.Width, $hp.ClientSize.Height, $true) | Out-Null
                                    }
                                }
                            } catch { }
                        }.GetNewClosure())
                    }
                } catch {
                    # Prozess hat möglicherweise noch kein Fenster
                }
            } else {
                # Eingebettetes Fenster nur bei Grössenänderung anpassen (Scroll-Position bleibt erhalten)
                try {
                    $hWnd = $global:guiWorkerHandles[$i]
                    if ($hWnd -ne [IntPtr]::Zero) {
                        $hp = $workerHostPanels[$i]
                        $curW = $hp.ClientSize.Width
                        $curH = $hp.ClientSize.Height
                        if ($curW -gt 0 -and $curH -gt 0) {
                            $lastSize = $global:guiLastPanelSize[$i]
                            $sizeChanged = ($null -eq $lastSize) -or ($lastSize.Width -ne $curW) -or ($lastSize.Height -ne $curH)
                            
                            if ($sizeChanged) {
                                [NativeMethods]::MoveWindow($hWnd, 0, 0, $curW, $curH, $true) | Out-Null
                                $global:guiLastPanelSize[$i] = @{ Width = $curW; Height = $curH }
                                
                                # Schriftgrösse dynamisch an Panel-Höhe anpassen
                                $desiredH = [math]::Max(8, [math]::Min(22, [int][math]::Floor($curH / 25)))
                                if ($desiredH -ne $global:guiLastFontHeight[$i]) {
                                    $wn = $global:guiWorkerNames[$i]
                                    if ($wn) {
                                        $fontPath = Join-Path (Join-Path $global:scriptFolder "OneView_Update") "${wn}_FontSize.txt"
                                        $desiredH | Out-File -FilePath $fontPath -Force
                                        $global:guiLastFontHeight[$i] = $desiredH
                                    }
                                }
                            }
                        }
                    }
                } catch { }
            }
            
            # Worker-Status aktualisieren
            if ($null -ne $proc) {
                $wn = $global:guiWorkerNames[$i]
                $markerPath = Join-Path (Join-Path $global:scriptFolder "OneView_Update") "${wn}_Complete.txt"
                if ((Test-Path $markerPath) -or $proc.HasExited) {
                    $workerStatusLabels[$i].Text = "Fertig"
                    $workerStatusLabels[$i].ForeColor = $colorSuccess
                } else {
                    $workerStatusLabels[$i].Text = "Läuft..."
                    $workerStatusLabels[$i].ForeColor = $workerColors[$i]
                }
            }
        }
        
        # Prüfe ob alle Worker fertig
        $allDone = $true
        for ($i = 0; $i -lt 4; $i++) {
            $proc = $global:guiWorkerProcesses[$i]
            if ($null -ne $proc) {
                $wn = $global:guiWorkerNames[$i]
                $markerPath = Join-Path (Join-Path $global:scriptFolder "OneView_Update") "${wn}_Complete.txt"
                if (-not (Test-Path $markerPath) -and -not $proc.HasExited) {
                    $allDone = $false
                    break
                }
            }
        }
        if ($global:guiUpdateRunning -and $allDone) {
            $anyWorkerWasStarted = $false
            for ($i = 0; $i -lt 4; $i++) {
                if ($null -ne $global:guiWorkerProcesses[$i]) { $anyWorkerWasStarted = $true; break }
            }
            if ($anyWorkerWasStarted) {
                $global:guiUpdateRunning = $false
                $statusItem.Text = "Alle Updates abgeschlossen!"
                $statusItem.ForeColor = $colorSuccess
                $startButton.Enabled = $true
                $startButton.Text = "▶  Update starten"
                $startButton.BackColor = $colorSuccess
            }
        }
    })
    $logTimer.Start()

    # ========================================
    # START BUTTON - Update starten
    # ========================================
    $startButton.Add_Click({
        # Validierung
        $selectedFilePath = $filePathLabel.Text
        if ($selectedFilePath -eq "(Keine Datei ausgewählt)" -or [string]::IsNullOrWhiteSpace($selectedFilePath)) {
            $fileValidationLabel.Text = "Bitte eine Update-Datei auswählen!"
            $fileValidationLabel.ForeColor = $colorError
            return
        }

        $enteredVersion = $versionTextBox.Text
        if ([string]::IsNullOrWhiteSpace($enteredVersion)) {
            $fileValidationLabel.Text = "Bitte eine Zielversion eingeben!"
            $fileValidationLabel.ForeColor = $colorError
            return
        }

        if ($enteredVersion -notmatch '^\d{1,2}\.\d{2}\.\d{2}$') {
            $fileValidationLabel.Text = "Ungültiges Format! Bsp: 9.40.01, 10.20.00"
            $fileValidationLabel.ForeColor = $colorError
            return
        }

        try {
            $parsedVersion = ConvertTo-NormalizedVersion $enteredVersion
            if ($null -eq $parsedVersion -or $null -eq $parsedVersion.Major) {
                throw "Ungültige Version"
            }
        } catch {
            $fileValidationLabel.Text = "Ungültige Versionsnummer!"
            $fileValidationLabel.ForeColor = $colorError
            return
        }

        # Ausgewählte Appliances
        $selectedAppliances = @()
        for ($i = 0; $i -lt $applianceListBox.Items.Count; $i++) {
            if ($applianceListBox.GetItemChecked($i)) {
                $selectedAppliances += $applianceListBox.Items[$i]
            }
        }
        if ($selectedAppliances.Count -eq 0) {
            $fileValidationLabel.Text = "Bitte mindestens eine Appliance auswählen!"
            $fileValidationLabel.ForeColor = $colorError
            return
        }

        # Anmeldedaten prüfen
        $guiUserName = $userNameTextBox.Text.Trim()
        $guiPassword = $passwordTextBox.Text
        if ([string]::IsNullOrWhiteSpace($guiUserName) -or [string]::IsNullOrWhiteSpace($guiPassword)) {
            $fileValidationLabel.Text = "Bitte Benutzername und Passwort eingeben!"
            $fileValidationLabel.ForeColor = $colorError
            return
        }
        # Backup-Passphrase ist Pflichtfeld (muss länger sein als Login-Passwort)
        $guiPassphrase = $backupPassTextBox.Text
        if ([string]::IsNullOrWhiteSpace($guiPassphrase)) {
            $fileValidationLabel.Text = "Bitte Backup-Passwort eingeben!"
            $fileValidationLabel.ForeColor = $colorError
            return
        }

        $fileValidationLabel.Text = ""

        # Auswahl in Datei speichern
        $updateFile = Join-Path $global:scriptFolder "oneview_update.txt"
        $selectedAppliances | Set-Content -Path $updateFile

        # Bestätigung
        $modeText = if ($sequentialRadio.Checked) { "Sequenziell" } else { "Parallel ($($workerCombo.SelectedItem) Worker)" }
        $confirmMsg = "Update starten?`n`nDatei: $selectedFilePath`nVersion: $enteredVersion`nAppliances: $($selectedAppliances.Count)`nModus: $modeText"
        $result = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "Bestätigung", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        # UI deaktivieren
        $startButton.Enabled = $false
        $startButton.Text = "Update läuft..."
        $startButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
        $statusItem.Text = "Update wird gestartet..."
        $statusItem.ForeColor = $colorWarning

        # Worker-Panels zurücksetzen
        for ($i = 0; $i -lt 4; $i++) {
            # Alten Worker-Prozess beenden falls noch aktiv
            $oldProc = $global:guiWorkerProcesses[$i]
            if ($null -ne $oldProc -and -not $oldProc.HasExited) {
                try { $oldProc.Kill() } catch { }
            }
            # Alte Marker-Datei und Handle-Datei entfernen
            if ($global:guiWorkerNames[$i]) {
                $oldMarker = Join-Path (Join-Path $global:scriptFolder "OneView_Update") "$($global:guiWorkerNames[$i])_Complete.txt"
                $oldHandle = Join-Path (Join-Path $global:scriptFolder "OneView_Update") "$($global:guiWorkerNames[$i])_Handle.txt"
                $oldFont = Join-Path (Join-Path $global:scriptFolder "OneView_Update") "$($global:guiWorkerNames[$i])_FontSize.txt"
                Remove-Item $oldMarker -Force -ErrorAction SilentlyContinue
                Remove-Item $oldHandle -Force -ErrorAction SilentlyContinue
                Remove-Item $oldFont -Force -ErrorAction SilentlyContinue
            }
            $workerStatusLabels[$i].Text = "Wartend"
            $workerStatusLabels[$i].ForeColor = [System.Drawing.Color]::Gray
            $global:guiWorkerEmbedded[$i] = $false
            $global:guiWorkerHandles[$i] = [IntPtr]::Zero
            $global:guiWorkerProcesses[$i] = $null
            $global:guiWorkerNames[$i] = $null
        }

        if ($sequentialRadio.Checked) {
            # ========================================
            # SEQUENZIELLER MODUS
            # ========================================
            $workerTitleLabels[0].Text = "Sequenziell"
            $workerStatusLabels[0].Text = "Läuft..."
            $workerStatusLabels[0].ForeColor = $colorWorker1
            
            # Worker 2-4 ausblenden
            for ($i = 1; $i -lt 4; $i++) {
                $workerTitleLabels[$i].Text = "Worker $($i+1) (nicht aktiv)"
                $workerStatusLabels[$i].Text = "-"
                $workerStatusLabels[$i].ForeColor = [System.Drawing.Color]::DarkGray
            }

            $logDir = Join-Path $global:scriptFolder "OneView_Update"
            if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
            
            # Erstelle ein Worker-Script für den sequenziellen Modus (gleiche Logik, aber als externer Prozess)
            $seqWorkerPath = Join-Path $global:scriptFolder "Sequential_Update.ps1"
            New-WorkerScript -WorkerScriptPath $seqWorkerPath -ApplianceGroup $selectedAppliances -DesiredVersion $parsedVersion -UpdateFilePath $selectedFilePath -LogDir $logDir -WorkerName "Sequential" -UserName $guiUserName -Password $guiPassword -Passphrase $guiPassphrase

            # PowerShell 7 suchen
            $pwsh7Path = $null
            $possiblePaths = @(
                "pwsh.exe",
                "C:\Program Files\PowerShell\7\pwsh.exe",
                "C:\Program Files (x86)\PowerShell\7\pwsh.exe",
                "$env:ProgramFiles\PowerShell\7\pwsh.exe",
                "$env:LOCALAPPDATA\Microsoft\powershell\7\pwsh.exe"
            )
            foreach ($path in $possiblePaths) {
                try {
                    $expandedPath = [System.Environment]::ExpandEnvironmentVariables($path)
                    if (Test-Path $expandedPath -PathType Leaf) {
                        $pwsh7Path = $expandedPath
                        break
                    } elseif (Get-Command $path -ErrorAction SilentlyContinue) {
                        $pwsh7Path = $path
                        break
                    }
                } catch { continue }
            }
            if (-not $pwsh7Path) { $pwsh7Path = "powershell.exe" }

            $global:guiWorkerNames[0] = "Sequential"
            $workerArgs = @("-ExecutionPolicy", "Bypass", "-NoExit", "-File", $seqWorkerPath)
            $proc = Start-Process -FilePath $pwsh7Path -ArgumentList $workerArgs -WindowStyle Hidden -PassThru -ErrorAction Stop
            $global:guiWorkerProcesses[0] = $proc
            $global:guiUpdateRunning = $true

            $statusItem.Text = "Sequenzielles Update läuft - $($selectedAppliances.Count) Appliances..."
            $statusItem.ForeColor = $colorWorker1

        } else {
            # ========================================
            # PARALLELER MODUS
            # ========================================
            $workerCount = [int]$workerCombo.SelectedItem
            if ($workerCount -gt $selectedAppliances.Count) {
                $workerCount = $selectedAppliances.Count
            }
            
            $logDir = Join-Path $global:scriptFolder "OneView_Update"
            if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

            # Appliances Round-Robin verteilen
            $groups = @()
            for ($i = 0; $i -lt $workerCount; $i++) { $groups += ,@() }
            for ($i = 0; $i -lt $selectedAppliances.Count; $i++) {
                $groupIndex = $i % $workerCount
                $groups[$groupIndex] += $selectedAppliances[$i]
            }

            # Worker-Panels konfigurieren
            for ($w = 0; $w -lt 4; $w++) {
                if ($w -lt $workerCount) {
                    $appList = ($groups[$w] -join ", ")
                    $workerTitleLabels[$w].Text = "Worker $($w+1) ($($groups[$w].Count) Appl.)"
                    $workerStatusLabels[$w].Text = "Starte..."
                    $workerStatusLabels[$w].ForeColor = $workerColors[$w]
                } else {
                    $workerTitleLabels[$w].Text = "Worker $($w+1) (nicht aktiv)"
                    $workerStatusLabels[$w].Text = "-"
                    $workerStatusLabels[$w].ForeColor = [System.Drawing.Color]::DarkGray
                }
            }

            # PowerShell 7 suchen
            $pwsh7Path = $null
            $possiblePaths = @(
                "pwsh.exe",
                "C:\Program Files\PowerShell\7\pwsh.exe",
                "C:\Program Files (x86)\PowerShell\7\pwsh.exe",
                "$env:ProgramFiles\PowerShell\7\pwsh.exe",
                "$env:LOCALAPPDATA\Microsoft\powershell\7\pwsh.exe"
            )
            foreach ($path in $possiblePaths) {
                try {
                    $expandedPath = [System.Environment]::ExpandEnvironmentVariables($path)
                    if (Test-Path $expandedPath -PathType Leaf) {
                        $pwsh7Path = $expandedPath
                        break
                    } elseif (Get-Command $path -ErrorAction SilentlyContinue) {
                        $pwsh7Path = $path
                        break
                    }
                } catch { continue }
            }
            if (-not $pwsh7Path) { $pwsh7Path = "powershell.exe" }

            # Worker starten
            for ($w = 0; $w -lt $workerCount; $w++) {
                $workerName = "Worker$($w+1)"
                $workerScriptPath = Join-Path $global:scriptFolder "${workerName}_Update.ps1"
                
                New-WorkerScript -WorkerScriptPath $workerScriptPath -ApplianceGroup $groups[$w] -DesiredVersion $parsedVersion -UpdateFilePath $selectedFilePath -LogDir $logDir -WorkerName $workerName -UserName $guiUserName -Password $guiPassword -Passphrase $guiPassphrase

                $global:guiWorkerNames[$w] = $workerName
                $workerArgs = @("-ExecutionPolicy", "Bypass", "-NoExit", "-File", $workerScriptPath)
                $proc = Start-Process -FilePath $pwsh7Path -ArgumentList $workerArgs -WindowStyle Hidden -PassThru -ErrorAction Stop
                $global:guiWorkerProcesses[$w] = $proc

                Start-Sleep -Seconds 3
            }

            $global:guiUpdateRunning = $true
            $statusItem.Text = "Paralleles Update läuft - $workerCount Worker, $($selectedAppliances.Count) Appliances..."
            $statusItem.ForeColor = $colorWorker3
        }
    })

    # ========================================
    # Form schließen - Cleanup
    # ========================================
    $form.Add_FormClosing({
        $logTimer.Stop()
        $logTimer.Dispose()
        
        # Eingebettete Konsolenfenster aus GUI lösen bevor Form geschlossen wird
        for ($i = 0; $i -lt 4; $i++) {
            $hWnd = $global:guiWorkerHandles[$i]
            if ($hWnd -ne [IntPtr]::Zero) {
                # Fenster aus Panel lösen (zurück auf Desktop)
                [NativeMethods]::SetParent($hWnd, [IntPtr]::Zero) | Out-Null
                # Window-Style wiederherstellen (eigenes Fenster)
                $style = [NativeMethods]::GetWindowLong($hWnd, [NativeMethods]::GWL_STYLE)
                $style = $style -band (-bnot [NativeMethods]::WS_CHILD)
                $style = $style -bor [NativeMethods]::WS_CAPTION
                $style = $style -bor [NativeMethods]::WS_THICKFRAME
                $style = $style -bor [NativeMethods]::WS_SYSMENU
                [NativeMethods]::SetWindowLong($hWnd, [NativeMethods]::GWL_STYLE, $style) | Out-Null
                # Extended Style wiederherstellen
                $exStyle = [NativeMethods]::GetWindowLong($hWnd, [NativeMethods]::GWL_EXSTYLE)
                $exStyle = $exStyle -bor [NativeMethods]::WS_EX_CLIENTEDGE
                $exStyle = $exStyle -bor [NativeMethods]::WS_EX_WINDOWEDGE
                [NativeMethods]::SetWindowLong($hWnd, [NativeMethods]::GWL_EXSTYLE, $exStyle) | Out-Null
                [NativeMethods]::ShowWindow($hWnd, [NativeMethods]::SW_SHOW) | Out-Null
            }
            
            # Worker-Prozess beenden und aufräumen
            $proc = $global:guiWorkerProcesses[$i]
            if ($null -ne $proc) {
                if (-not $proc.HasExited) {
                    try { $proc.Kill() } catch { }
                }
                $wn = $global:guiWorkerNames[$i]
                if ($wn) {
                    $scriptPath = Join-Path $global:scriptFolder "${wn}_Update.ps1"
                    $markerPath = Join-Path (Join-Path $global:scriptFolder "OneView_Update") "${wn}_Complete.txt"
                    $handlePath = Join-Path (Join-Path $global:scriptFolder "OneView_Update") "${wn}_Handle.txt"
                    $fontPath = Join-Path (Join-Path $global:scriptFolder "OneView_Update") "${wn}_FontSize.txt"
                    Remove-Item $scriptPath -Force -ErrorAction SilentlyContinue
                    Remove-Item $markerPath -Force -ErrorAction SilentlyContinue
                    Remove-Item $handlePath -Force -ErrorAction SilentlyContinue
                    Remove-Item $fontPath -Force -ErrorAction SilentlyContinue
                }
            }
        }
    })

    # Form anzeigen
    [void]$form.ShowDialog()
    $form.Dispose()
}

# ============================================
# Hauptteil des Skripts
# ============================================
Write-Host "Programm startet..." -ForegroundColor Cyan

# PowerShell-Version prüfen
Test-PowerShellVersion

try {
    # Konsolenfenster zentrieren
    Set-ConsoleWindowPosition -WindowPixelWidth 800 -WindowPixelHeight 600
} catch {
    Write-Host "Warnung: Konsolenfenster konnte nicht zentriert werden: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Unified GUI starten
try {
    Show-UnifiedGUI
} catch {
    Write-Host "`nFEHLER beim Starten der GUI:" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Zeile: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    Read-Host "Drücke Enter zum Beenden..."
    exit 1
}
