<#
.SYNOPSIS
    Erstellt ein Icon (.ico) und eine Desktop-Verknüpfung für den OneView Manager.

.DESCRIPTION
    - Generiert eine OneView-Manager-Icon-Datei (OneViewManager.ico) im Script-Verzeichnis,
      sofern noch nicht vorhanden (mehrere Auflösungen 16/32/48/64/128/256).
    - Legt auf dem Windows-Desktop des aktuellen Benutzers eine Verknüpfung
      "OneView Manager.lnk" an, die OneView-Manager-GUI.ps1 mit pwsh.exe startet.

.NOTES
    Auf Windows ausführen:
        powershell -ExecutionPolicy Bypass -File .\Create-DesktopShortcut.ps1
#>

[CmdletBinding()]
param(
    [string]$ShortcutName = "OneView Manager",
    [switch]$Force
)

Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$targetPs1   = Join-Path $scriptDir 'OneView-Manager-GUI.ps1'
$iconPath    = Join-Path $scriptDir 'OneViewManager.ico'

if (-not (Test-Path $targetPs1)) {
    throw "OneView-Manager-GUI.ps1 nicht gefunden: $targetPs1"
}

# ---------------------------------------------------------------------------
# 1) Icon erzeugen (falls nicht vorhanden oder -Force)
# ---------------------------------------------------------------------------
function New-OneViewIcon {
    param(
        [string]$Path,
        [int[]]$Sizes = @(16, 32, 48, 64, 128, 256)
    )

    # Pro Größe ein Bitmap mit blauem Verlauf, weißem "OV"-Schriftzug und Akzent
    $bitmaps = New-Object 'System.Collections.Generic.List[System.Drawing.Bitmap]'
    try {
        foreach ($s in $Sizes) {
            $bmp = New-Object System.Drawing.Bitmap($s, $s, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $g   = [System.Drawing.Graphics]::FromImage($bmp)
            try {
                $g.SmoothingMode    = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $g.InterpolationMode= [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $g.TextRenderingHint= [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
                $g.Clear([System.Drawing.Color]::Transparent)

                # Hintergrund: abgerundetes Rechteck mit Verlauf
                $rect = New-Object System.Drawing.RectangleF -ArgumentList ([single]0, [single]0, [single]$s, [single]$s)
                $radius = [int][Math]::Max(2, [int]($s * 0.18))
                $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
                $d = [int]($radius * 2)
                $right  = [int]($s - $d)
                $bottom = [int]($s - $d)
                $gp.AddArc(0, 0, $d, $d, 180, 90)
                $gp.AddArc($right, 0, $d, $d, 270, 90)
                $gp.AddArc($right, $bottom, $d, $d, 0, 90)
                $gp.AddArc(0, $bottom, $d, $d, 90, 90)
                $gp.CloseFigure()

                $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush -ArgumentList @(
                    $rect,
                    [System.Drawing.Color]::FromArgb(255, 0, 100, 180),
                    [System.Drawing.Color]::FromArgb(255, 0, 50, 110),
                    [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal)
                $g.FillPath($brush, $gp)
                $brush.Dispose()

                # Subtiler innerer Rand
                $penEdge = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(80, 255, 255, 255), [Math]::Max(1.0, $s/64.0))
                $g.DrawPath($penEdge, $gp)
                $penEdge.Dispose()
                $gp.Dispose()

                # Akzent-Streifen oben (HPE-orange-artig)
                $stripeH = [int][Math]::Max(2, [int]($s * 0.10))
                $stripeRect = New-Object System.Drawing.RectangleF -ArgumentList ([single]0, [single]0, [single]$s, [single]$stripeH)
                $stripeBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 255, 140, 0))
                $g.FillRectangle($stripeBrush, $stripeRect)
                $stripeBrush.Dispose()

                if ($s -ge 24) {
                    # "OV" Schriftzug
                    $fontSize = [single]($s * 0.46)
                    $font = New-Object System.Drawing.Font("Segoe UI", $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
                    $sf = New-Object System.Drawing.StringFormat
                    $sf.Alignment     = [System.Drawing.StringAlignment]::Center
                    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
                    $textRect = New-Object System.Drawing.RectangleF -ArgumentList ([single]0, [single]($stripeH * 0.4), [single]$s, [single]($s - $stripeH * 0.4))
                    $textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
                    $g.DrawString("OV", $font, $textBrush, $textRect, $sf)
                    $textBrush.Dispose()
                    $font.Dispose()
                    $sf.Dispose()
                } else {
                    # Bei kleinen Größen nur ein weißer Mittelpunkt
                    $dot = [int]($s * 0.45)
                    $dx  = [int](($s - $dot) / 2)
                    $dy  = [int]((($s + $stripeH) - $dot) / 2)
                    $dotBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
                    $g.FillEllipse($dotBrush, $dx, $dy, $dot, $dot)
                    $dotBrush.Dispose()
                }
            }
            finally { $g.Dispose() }
            $bitmaps.Add($bmp)
        }

        # ICO-Datei zusammensetzen (Header + Verzeichnis + PNG-Daten je Größe)
        $entries = New-Object 'System.Collections.Generic.List[object]'
        foreach ($bmp in $bitmaps) {
            $ms = New-Object System.IO.MemoryStream
            $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
            $entries.Add([pscustomobject]@{
                Width  = $bmp.Width
                Height = $bmp.Height
                Bytes  = $ms.ToArray()
            })
            $ms.Dispose()
        }

        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
        $bw = New-Object System.IO.BinaryWriter($fs)
        try {
            # ICONDIR
            $bw.Write([uint16]0)                # reserved
            $bw.Write([uint16]1)                # type = icon
            $bw.Write([uint16]$entries.Count)   # count

            $headerSize = 6 + (16 * $entries.Count)
            $offset     = $headerSize

            # ICONDIRENTRYs
            foreach ($e in $entries) {
                $w = if ($e.Width  -ge 256) { 0 } else { [byte]$e.Width }
                $h = if ($e.Height -ge 256) { 0 } else { [byte]$e.Height }
                $bw.Write([byte]$w)
                $bw.Write([byte]$h)
                $bw.Write([byte]0)              # palette
                $bw.Write([byte]0)              # reserved
                $bw.Write([uint16]1)            # color planes
                $bw.Write([uint16]32)           # bits per pixel
                $bw.Write([uint32]$e.Bytes.Length)
                $bw.Write([uint32]$offset)
                $offset += $e.Bytes.Length
            }
            # Bilddaten
            foreach ($e in $entries) {
                $bw.Write($e.Bytes)
            }
        }
        finally {
            $bw.Dispose()
            $fs.Dispose()
        }
    }
    finally {
        foreach ($bmp in $bitmaps) { $bmp.Dispose() }
    }
}

if ((-not (Test-Path $iconPath)) -or $Force) {
    Write-Host "Erzeuge Icon: $iconPath"
    New-OneViewIcon -Path $iconPath
} else {
    Write-Host "Icon vorhanden: $iconPath (mit -Force neu erzeugen)"
}

# ---------------------------------------------------------------------------
# 2) Verknüpfung auf dem Desktop anlegen
# ---------------------------------------------------------------------------
if (-not $IsWindows -and ($PSVersionTable.PSEdition -eq 'Core') -and ($env:OS -ne 'Windows_NT')) {
    Write-Warning "Verknüpfungserstellung wird nur unter Windows unterstützt. Icon wurde aber erzeugt: $iconPath"
    return
}

# pwsh.exe bevorzugt, sonst powershell.exe
$pwshExe = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Source
if (-not $pwshExe) {
    $pwshExe = (Get-Command powershell.exe -ErrorAction SilentlyContinue).Source
}
if (-not $pwshExe) {
    throw "Weder pwsh.exe noch powershell.exe gefunden."
}

$desktop      = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop ("{0}.lnk" -f $ShortcutName)

$wshShell = New-Object -ComObject WScript.Shell
$sc = $wshShell.CreateShortcut($shortcutPath)
$sc.TargetPath       = $pwshExe
$sc.Arguments        = "-NoProfile -ExecutionPolicy Bypass -File `"$targetPs1`""
$sc.WorkingDirectory = $scriptDir
$sc.IconLocation     = "$iconPath,0"
$sc.Description      = "OneView Manager GUI"
$sc.WindowStyle      = 1
$sc.Save()

Write-Host "Desktop-Verknüpfung erstellt: $shortcutPath"
Write-Host "Ziel : $pwshExe $($sc.Arguments)"
Write-Host "Icon : $iconPath"
