#Requires -Version 5.1
<#
.SYNOPSIS
    Erzeugt eine Word-Vorlage (.dotx) für ein OneView Recovery-Handbuch.

.DESCRIPTION
    Dieses Script baut über die Word-COM-Automation eine strukturierte
    Word-Vorlage auf, die als Grundlage für ein Recovery-Handbuch
    (Wiederherstellungs- / Notfallhandbuch) für HPE OneView Appliances dient.

    Die Vorlage enthält:
      - Deckblatt mit Platzhaltern (Kunde, Standort, Version, Datum)
      - Inhaltsverzeichnis (automatisch generiert beim Öffnen)
      - Kapitelstruktur für ein OneView Recovery-Szenario
      - Tabellen für Appliance-Inventar, Ansprechpartner und
        Wiederherstellungsschritte
      - Kopf- und Fußzeile mit Klassifikationshinweis

    Voraussetzungen:
      - Windows mit installiertem Microsoft Word
      - PowerShell 7.x

.PARAMETER OutputPath
    Zielpfad der zu erzeugenden Vorlage. Standard: Skriptverzeichnis\
    Recovery_Handbook_Template.dotx

.PARAMETER Customer
    Optionaler Kundenname, der auf dem Deckblatt vorausgefüllt wird.

.EXAMPLE
    .\New-RecoveryHandbookTemplate.ps1

.EXAMPLE
    .\New-RecoveryHandbookTemplate.ps1 -OutputPath C:\Temp\Recovery.dotx -Customer "Musterfirma AG"
#>

[CmdletBinding()]
param(
    [string]$OutputPath,
    [string]$Customer = "<Kunde>"
)

# Plattform-Check (auch unter Windows PowerShell 5.1, wo $IsWindows nicht definiert ist)
$isWin = if ($null -ne $IsWindows) { [bool]$IsWindows } else { $true }
if (-not $isWin) {
    Write-Error "Dieses Script benötigt Windows mit installiertem Microsoft Word."
    return
}

# Pfad bestimmen
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $MyInvocation.MyCommand.Path -Parent }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path -Path $scriptDir -ChildPath "Recovery_Handbook_Template.dotx"
}

# Falls Vorlage existiert, abfragen
if (Test-Path $OutputPath) {
    Write-Host "Vorlage existiert bereits: $OutputPath" -ForegroundColor Yellow
    $answer = Read-Host "Überschreiben? (j/N)"
    if ($answer -notmatch '^[jJyY]') {
        Write-Host "Abgebrochen." -ForegroundColor Yellow
        return
    }
    Remove-Item -Path $OutputPath -Force -ErrorAction SilentlyContinue
}

# Word starten
try {
    $word = New-Object -ComObject Word.Application -ErrorAction Stop
}
catch {
    Write-Error "Microsoft Word konnte nicht gestartet werden. Ist Word installiert? Fehler: $_"
    return
}
$word.Visible = $false
$word.DisplayAlerts = 0  # wdAlertsNone

# Word-Konstanten
$wdAlignParagraphCenter   = 1
$wdAlignParagraphLeft     = 0
$wdStyleHeading1          = -2
$wdStyleHeading2          = -3
$wdStyleHeading3          = -4
$wdStyleTitle             = -63
$wdStyleSubtitle          = -75
$wdStyleNormal            = -1
$wdSeekCurrentPageHeader  = 9
$wdSeekCurrentPageFooter  = 10
$wdSeekMainDocument       = 0
$wdPageBreak              = 7
$wdFormatXMLTemplate      = 14   # .dotx
$wdCollapseEnd            = 0
$wdLineStyleSingle        = 1
$wdLineWidth050pt         = 4

function Add-Heading {
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][int]$Level
    )
    $sel = $word.Selection
    switch ($Level) {
        1 { $sel.Style = $wdStyleHeading1 }
        2 { $sel.Style = $wdStyleHeading2 }
        3 { $sel.Style = $wdStyleHeading3 }
    }
    $sel.TypeText($Text)
    $sel.TypeParagraph()
    $sel.Style = $wdStyleNormal
}

function Add-Paragraph {
    param([Parameter(Mandatory)][string]$Text)
    $sel = $word.Selection
    $sel.Style = $wdStyleNormal
    $sel.TypeText($Text)
    $sel.TypeParagraph()
}

function Add-BulletList {
    param([Parameter(Mandatory)][string[]]$Items)
    $sel = $word.Selection
    $sel.Style = $wdStyleNormal
    $sel.Range.ListFormat.ApplyBulletDefault()
    foreach ($item in $Items) {
        $sel.TypeText($item)
        $sel.TypeParagraph()
    }
    $sel.Range.ListFormat.RemoveNumbers()
}

function Add-Table {
    param(
        [Parameter(Mandatory)][string[]]$Headers,
        [Parameter(Mandatory)][int]$DataRows
    )
    $sel = $word.Selection
    $rows = $DataRows + 1
    $cols = $Headers.Count
    $range = $sel.Range
    $table = $doc.Tables.Add($range, $rows, $cols)
    $table.Borders.InsideLineStyle  = $wdLineStyleSingle
    $table.Borders.OutsideLineStyle = $wdLineStyleSingle
    $table.AllowAutoFit = $true
    for ($c = 0; $c -lt $cols; $c++) {
        $cell = $table.Cell(1, $c + 1)
        $cell.Range.Bold = $true
        $cell.Range.Shading.BackgroundPatternColor = 14737632  # helles Grau
        $cell.Range.Text = $Headers[$c]
    }
    # Cursor hinter Tabelle setzen
    $end = $table.Range.End
    $word.Selection.SetRange($end, $end)
    $word.Selection.TypeParagraph()
}

# Neues Dokument erzeugen
$doc = $word.Documents.Add()

try {
    # Seitenränder etwas reduzieren
    $doc.PageSetup.TopMargin    = $word.CentimetersToPoints(2.0)
    $doc.PageSetup.BottomMargin = $word.CentimetersToPoints(2.0)
    $doc.PageSetup.LeftMargin   = $word.CentimetersToPoints(2.2)
    $doc.PageSetup.RightMargin  = $word.CentimetersToPoints(2.2)

    # ---------------- Kopf- / Fußzeile ----------------
    $section = $doc.Sections.Item(1)
    $header = $section.Headers.Item(1)
    $header.Range.Text = "Recovery-Handbuch HPE OneView  |  $Customer"
    $header.Range.Font.Size = 9
    $header.Range.Font.Italic = $true

    $footer = $section.Footers.Item(1)
    $footer.Range.Text = "Vertraulich – nur für internen Gebrauch`tSeite "
    $footer.Range.ParagraphFormat.Alignment = $wdAlignParagraphLeft
    # Seitenzahlfeld einfügen
    $footerRange = $footer.Range
    $footerRange.Collapse($wdCollapseEnd)
    $doc.Fields.Add($footerRange, 33) | Out-Null  # wdFieldPage = 33
    $footer.Range.Font.Size = 9

    # ---------------- Deckblatt ----------------
    $sel = $word.Selection
    $sel.ParagraphFormat.Alignment = $wdAlignParagraphCenter
    $sel.TypeParagraph(); $sel.TypeParagraph(); $sel.TypeParagraph()

    $sel.Style = $wdStyleTitle
    $sel.TypeText("Recovery-Handbuch")
    $sel.TypeParagraph()

    $sel.Style = $wdStyleSubtitle
    $sel.TypeText("HPE OneView Appliances")
    $sel.TypeParagraph()
    $sel.TypeParagraph()

    $sel.Style = $wdStyleNormal
    $sel.Font.Size = 14
    $sel.TypeText("Kunde: $Customer")
    $sel.TypeParagraph()
    $sel.TypeText("Standort: <Standort>")
    $sel.TypeParagraph()
    $sel.TypeText("Version: 1.0")
    $sel.TypeParagraph()
    $sel.TypeText("Stand: $(Get-Date -Format 'dd.MM.yyyy')")
    $sel.TypeParagraph()
    $sel.TypeText("Autor: <Autor>")
    $sel.TypeParagraph()
    $sel.Font.Size = 11

    # Seitenumbruch
    $sel.InsertBreak($wdPageBreak)
    $sel.ParagraphFormat.Alignment = $wdAlignParagraphLeft

    # ---------------- Änderungshistorie ----------------
    Add-Heading -Text "Änderungshistorie" -Level 1
    Add-Paragraph -Text "Übersicht über alle Änderungen an diesem Dokument."
    Add-Table -Headers @("Version", "Datum", "Autor", "Beschreibung") -DataRows 4

    $sel.InsertBreak($wdPageBreak)

    # ---------------- Inhaltsverzeichnis ----------------
    Add-Heading -Text "Inhaltsverzeichnis" -Level 1
    $tocRange = $sel.Range
    $doc.TablesOfContents.Add($tocRange, $true, 1, 3) | Out-Null
    $sel.EndKey(6) | Out-Null   # wdStory = 6
    $sel.InsertBreak($wdPageBreak)

    # ---------------- 1. Einleitung ----------------
    Add-Heading -Text "1. Einleitung" -Level 1
    Add-Heading -Text "1.1 Zweck des Dokuments" -Level 2
    Add-Paragraph -Text "Dieses Handbuch beschreibt die Vorgehensweise zur Wiederherstellung der HPE OneView Appliances im Fehler- oder Katastrophenfall. Es richtet sich an Administratoren und Betriebspersonal."

    Add-Heading -Text "1.2 Geltungsbereich" -Level 2
    Add-Paragraph -Text "Das Handbuch gilt für alle in Abschnitt 2 aufgeführten OneView Appliances und die zugehörigen Synergy-/BladeSystem-Umgebungen."

    Add-Heading -Text "1.3 Mitgeltende Dokumente" -Level 2
    Add-BulletList -Items @(
        "Betriebshandbuch HPE OneView",
        "Netzwerkkonzept (VLAN-Konzept, Uplink-Sets)",
        "Backup-Konzept der Infrastruktur",
        "Notfall- und Wiederanlaufplan (BCM/DR-Plan)"
    )

    Add-Heading -Text "1.4 Begriffe und Abkürzungen" -Level 2
    Add-Table -Headers @("Abkürzung", "Bedeutung") -DataRows 6

    $sel.InsertBreak($wdPageBreak)

    # ---------------- 2. Umgebungsübersicht ----------------
    Add-Heading -Text "2. Umgebungsübersicht" -Level 1
    Add-Heading -Text "2.1 Appliance-Inventar" -Level 2
    Add-Paragraph -Text "Übersicht aller OneView Appliances inklusive Version, Standort und Funktion."
    Add-Table -Headers @("Appliance (FQDN/IP)", "OneView-Version", "Standort", "Funktion", "Verwaltete Frames") -DataRows 6

    Add-Heading -Text "2.2 Ansprechpartner" -Level 2
    Add-Table -Headers @("Rolle", "Name", "Telefon", "E-Mail", "Erreichbarkeit") -DataRows 6

    Add-Heading -Text "2.3 Abhängigkeiten" -Level 2
    Add-BulletList -Items @(
        "Active Directory / LDAP",
        "DNS / NTP",
        "SMTP-Relay (Alerts)",
        "Backup-Ziel (SMB/CIFS-Share)",
        "Monitoring (SNMP / Syslog)"
    )

    $sel.InsertBreak($wdPageBreak)

    # ---------------- 3. Backup-Strategie ----------------
    Add-Heading -Text "3. Backup-Strategie" -Level 1
    Add-Heading -Text "3.1 Übersicht" -Level 2
    Add-Paragraph -Text "Beschreibung der eingesetzten Backup-Verfahren für OneView (Appliance-Backup, Konfigurations-Export, Server-Profile, Network-Sets)."

    Add-Heading -Text "3.2 Backup-Zeitplan" -Level 2
    Add-Table -Headers @("Backup-Typ", "Zeitplan", "Aufbewahrung", "Speicherort", "Verantwortlich") -DataRows 5

    Add-Heading -Text "3.3 Backup-Verifikation" -Level 2
    Add-BulletList -Items @(
        "Prüfung des Backup-Logs (täglich)",
        "Stichprobenartiger Restore-Test (vierteljährlich)",
        "Integritätsprüfung der Backup-Dateien (Hash/Checksum)"
    )

    $sel.InsertBreak($wdPageBreak)

    # ---------------- 4. Recovery-Szenarien ----------------
    Add-Heading -Text "4. Recovery-Szenarien" -Level 1
    Add-Paragraph -Text "Die folgenden Abschnitte beschreiben typische Wiederherstellungsszenarien. Pro Szenario sind Voraussetzungen, Schritte, erwartete Dauer und Verifikation dokumentiert."

    $scenarios = @(
        @{ Title = "4.1 Wiederherstellung einer OneView Appliance aus Backup"; Desc = "Vollständige Wiederherstellung einer Appliance aus einem OneView-Backup auf neuer/zurückgesetzter Hardware." },
        @{ Title = "4.2 Wiederherstellung nach Konfigurationsfehler"; Desc = "Rollback einer fehlerhaften Konfigurationsänderung (z.B. Network-Set, Uplink-Set, Server-Profile-Template)." },
        @{ Title = "4.3 Ausfall einer Synergy Composer / Frame Link Module"; Desc = "Wiederherstellung der Frame-Management-Verbindung und Re-Sync mit OneView." },
        @{ Title = "4.4 Wiederherstellung eines einzelnen Server-Profils"; Desc = "Re-Import oder Neuzuweisung eines verlorenen / beschädigten Server-Profils." },
        @{ Title = "4.5 Recovery nach Zertifikats-Problemen"; Desc = "Erneuerung / Wiederherstellung des Appliance-Zertifikats und der Vertrauensstellungen." },
        @{ Title = "4.6 Komplettausfall des Standorts"; Desc = "Disaster-Recovery-Vorgehen bei vollständigem Ausfall der Management-Infrastruktur." }
    )

    foreach ($s in $scenarios) {
        Add-Heading -Text $s.Title -Level 2

        Add-Heading -Text "Beschreibung" -Level 3
        Add-Paragraph -Text $s.Desc

        Add-Heading -Text "Voraussetzungen" -Level 3
        Add-BulletList -Items @(
            "<Voraussetzung 1>",
            "<Voraussetzung 2>",
            "<Voraussetzung 3>"
        )

        Add-Heading -Text "Vorgehen" -Level 3
        Add-Table -Headers @("Schritt", "Aktion", "Befehl / Tool", "Erwartetes Ergebnis", "Verantwortlich") -DataRows 6

        Add-Heading -Text "Verifikation" -Level 3
        Add-BulletList -Items @(
            "<Prüfschritt 1>",
            "<Prüfschritt 2>"
        )

        Add-Heading -Text "Geschätzte Dauer / RTO" -Level 3
        Add-Paragraph -Text "<z.B. 2 Stunden>"

        $sel.InsertBreak($wdPageBreak)
    }

    # ---------------- 5. Tests & Übungen ----------------
    Add-Heading -Text "5. Tests und Übungen" -Level 1
    Add-Paragraph -Text "Regelmäßige Tests stellen sicher, dass die Recovery-Prozesse funktionieren und das Personal mit dem Vorgehen vertraut ist."
    Add-Table -Headers @("Testfall", "Frequenz", "Letzte Durchführung", "Ergebnis", "Verantwortlich") -DataRows 5

    $sel.InsertBreak($wdPageBreak)

    # ---------------- 6. Eskalation ----------------
    Add-Heading -Text "6. Eskalation und Support" -Level 1
    Add-Heading -Text "6.1 Eskalationsstufen" -Level 2
    Add-Table -Headers @("Stufe", "Auslöser", "Zuständig", "Reaktionszeit") -DataRows 4

    Add-Heading -Text "6.2 Hersteller-Support (HPE)" -Level 2
    Add-BulletList -Items @(
        "HPE Support-Vertrag / Service-Agreement-ID: <SAID>",
        "HPE Support Hotline: <Telefonnummer>",
        "HPE Support Portal: https://support.hpe.com"
    )

    $sel.InsertBreak($wdPageBreak)

    # ---------------- 7. Anhang ----------------
    Add-Heading -Text "7. Anhang" -Level 1
    Add-Heading -Text "7.1 Checkliste Recovery" -Level 2
    Add-BulletList -Items @(
        "Aktuelles Backup verfügbar?",
        "Hardware betriebsbereit?",
        "Netzwerk-Konnektivität verfügbar?",
        "Zugangsdaten verfügbar (Admin / Hardware-iLO)?",
        "DNS-Einträge korrekt?",
        "Lizenz-Schlüssel verfügbar?"
    )

    Add-Heading -Text "7.2 Referenzen" -Level 2
    Add-BulletList -Items @(
        "HPE OneView Administrator Guide",
        "HPE OneView Online Help",
        "Interne Runbooks"
    )

    # ---------------- TOC aktualisieren ----------------
    foreach ($t in $doc.TablesOfContents) { $t.Update() }

    # ---------------- Als .dotx speichern ----------------
    $doc.SaveAs([ref]$OutputPath, [ref]$wdFormatXMLTemplate) | Out-Null
    Write-Host "Vorlage erfolgreich erstellt: $OutputPath" -ForegroundColor Green
}
catch {
    Write-Error "Fehler beim Erstellen der Vorlage: $_"
}
finally {
    try { $doc.Close([ref]$false) } catch {}
    try { $word.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
