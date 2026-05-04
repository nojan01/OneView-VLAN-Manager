# HPE OneView Manager

Grafische Oberfläche (WinForms GUI) zur Verwaltung von **Ethernet Networks**, **Network Sets**, **Server Profiles** und **Server Profile Templates** in HPE OneView. Unterstützt mehrere Appliances gleichzeitig und verwendet ausschliesslich die **HPE OneView RESTful API** (ohne das HPE OneView PowerShell-Modul).

Darüber hinaus enthält das Projekt mehrere **eigenständige Tools** rund um den Betrieb von HPE OneView/Synergy: Appliance-Update, Config-Backup, User-Verwaltung, Alerts und ein Synergy-eFuse-Tool.

## Features

### Eigenständige Tools (in der Haupt-GUI verlinkt)

Die Haupt-GUI `OneView-Manager-GUI.ps1` enthält oben einen Bereich **„Eigenständige Tools"**, der folgende separate Skripte startet (jedes Tool hat eine eigene GUI und eine eigene Anmeldung):

| Button | Skript | Funktion |
|---|---|---|
| **Appliance Update** | `Oneview_Update/OneView_Update_GUI.ps1` | OneView/Synergy Update-Workflow |
| **Config Backup** | `Oneview_Backup/Backup_OneView_GUI.ps1` | OneView Konfigurations-Backup (manuell + geplant) |
| **User Manager** | `Oneview_UserManager/OneView-UserManager-GUI.ps1` | Benutzerverwaltung über mehrere Appliances (OV 6.60 und 11.10) |
| **Alerts** | `Oneview_Alerts/Oneview_Alerts_GUI.ps1` | Alert-Auswertung (manuell + geplant) |
| **Synergy eFuse** | `Synergy_eFuse/eFuse_GUI.ps1` | eFuse für Synergy-Komponenten (Device/ICM/Appliance/FLM) |

### GUI (OneView-Manager-GUI.ps1)

Die zentrale Anwendung bietet folgende Hauptfunktionen über Buttons:

| Button | Funktion | Beschreibung |
|---|---|---|
| **Netzwerk erstellen (Import)** | Einzelne Appliance | Erstellt/aktualisiert Ethernet Networks aus einer Excel-Datei auf einer ausgewählten Appliance |
| **VLAN Backup (Multi)** | Mehrere Appliances | Exportiert alle Ethernet Networks von mehreren Appliances gleichzeitig als Excel-Dateien |
| **Netzwerk erstellen (Multi)** | Mehrere Appliances | Erstellt ein einzelnes Netzwerk auf mehreren Appliances mit optionaler Network Set Zuweisung |
| **Network Sets importieren** | Einzelne Appliance | Erstellt/aktualisiert Network Sets aus einer Excel-Datei auf einer ausgewählten Appliance |
| **Network Set Backup (Multi)** | Mehrere Appliances | Exportiert alle Network Sets von mehreren Appliances gleichzeitig als Excel-Dateien |
| **SP exportieren (JSON)** | Mehrere Appliances | Exportiert alle Server Profiles als JSON-Dateien |
| **SP importieren (JSON)** | Einzelne Appliance | Importiert Server Profiles aus JSON-Dateien (Auto/Create/Update) |
| **SP verwalten** | Einzelne Appliance | CRUD-Dialog: Profile anzeigen, erstellen, bearbeiten, löschen |
| **SP JSON Editor** | Einzelne Appliance | Vollständiger JSON-Editor für alle Profil-Felder mit Create/Update in OneView |
| **SPT exportieren (JSON)** | Mehrere Appliances | Exportiert alle Server Profile Templates als JSON-Dateien |
| **SPT importieren (JSON)** | Einzelne Appliance | Importiert Server Profile Templates aus JSON-Dateien (Auto/Create/Update) |
| **SPT verwalten** | Einzelne Appliance | CRUD-Dialog: Templates anzeigen, aus Profil erstellen, löschen, Profil erzeugen |
| **SPT JSON Editor** | Einzelne Appliance | Vollständiger JSON-Editor für alle Template-Felder mit Create/Update in OneView |

### Appliance-Typen und Filterung

Appliances können in der `Appliances.txt` mit einem Typ versehen werden (z.B. ESXi, VDI), um verschiedene Synergy Frames zu unterscheiden. Der Appliance-Auswahldialog zeigt den Typ in Klammern an und bietet dynamische Filter-Buttons:

- **Alle auswählen** / **Keine auswählen** – Standard-Auswahl
- **Alle ESXi** / **Alle VDI** – Filtert nach Typ (Buttons werden automatisch aus den vorhandenen Typen generiert)

### Ethernet Network Management

- **Import aus Excel**: Netzwerke werden aus einer Excel-Datei gelesen und auf der Appliance erstellt
- **Update bestehender Netzwerke**: Bereits existierende Netzwerke werden erkannt und bei Abweichungen aktualisiert (Purpose, SmartLink, PrivateNetwork, Bandwidth, Scope, Network Set)
- **Duplikatserkennung**: Netzwerke mit gleichem Namen und VLAN-ID werden übersprungen, wenn keine Änderungen nötig sind
- **Bandwidth-Konfiguration**: Preferred und Maximum Bandwidth werden über Connection Templates gesetzt
- **Scope-Unterstützung**: Netzwerke können Scopes zugewiesen werden
- **Network Set Zuweisung**: Netzwerke können beim Import direkt Network Sets zugewiesen werden

### Network Set Management

- **Import aus Excel**: Network Sets werden mit zugehörigen Netzwerken aus einer Excel-Datei erstellt
- **Update bestehender Network Sets**: Änderungen an Netzwerkzuweisungen und Bandwidth werden erkannt und aktualisiert
- **Backup**: Export aller Network Sets inkl. zugeordneter Netzwerke als Excel-Datei

### Server Profile Management

- **Export**: Alle Server Profiles als einzelne JSON-Dateien exportieren (inkl. Index-Datei)
- **Import**: JSON-Dateien importieren mit drei Modi: Auto (erkennt automatisch), Create (nur neue), Update (nur bestehende)
- **Verwalten**: Interaktiver Dialog mit Profil-Liste, Detail-Ansicht und Aktionen (Neu, Bearbeiten, Löschen, Exportieren)
- **JSON Editor**: Vollzugriff auf alle Server Profile Felder (Firmware, BIOS, Connections, Storage, Boot etc.) – JSON von Datei laden, von OneView laden, editieren und als neues Profil anlegen oder bestehendes updaten

### Server Profile Template Management

- **Export**: Alle Server Profile Templates als einzelne JSON-Dateien exportieren (inkl. Index-Datei mit HW-Type und Enclosure Group)
- **Import**: JSON-Dateien importieren mit drei Modi: Auto (erkennt automatisch), Create (nur neue), Update (nur bestehende)
- **Verwalten**: Interaktiver Dialog mit Template-Liste (Name, Status, HW Type, Enclosure Group), Detail-Ansicht und Aktionen:
  - **Aus Profil erstellen**: Neues Template basierend auf einem bestehenden Server Profile anlegen
  - **Profil erzeugen**: Neues Server Profile aus einem ausgewählten Template erstellen (mit Hardware-Auswahl)
  - **Löschen**, **Exportieren**, **JSON Editor** starten
- **JSON Editor**: Vollzugriff auf alle Template-Felder – JSON von Datei laden, von OneView laden, editieren und als neues Template anlegen oder bestehendes updaten

### Multi-Deploy

Erstellt ein einzelnes Netzwerk auf mehreren Appliances gleichzeitig:

1. Netzwerk-Parameter eingeben (Name, VLAN-ID, Typ, Purpose, Bandwidth usw.)
2. Ziel-Appliances auswählen (mit Typ-Filter)
3. Optional: Network Sets pro Appliance zuweisen (TreeView-Dialog)
4. Bestätigung und automatisches Deployment auf alle ausgewählten Appliances

### API-Version Auto-Detection

Die API-Version wird automatisch pro Appliance über `GET /rest/version` erkannt. Ein Fallback auf den in der `config.json` konfigurierten Wert ist vorhanden.

### Konsolidierte Logs

Bei Multi-Appliance-Operationen (Backup, Multi-Deploy) wird ein einzelnes Log pro Vorgang geschrieben, statt ein separates Log pro Appliance. Die Logdateien werden im Unterverzeichnis `Logs/` abgelegt.

### Live-Output

Die GUI bleibt während langer Operationen (z.B. Import von 250+ Netzwerken) responsiv. Die Ausgabe der Subprozesse wird zeilenweise in das Protokollfenster geschrieben.

## Synergy eFuse Tool

Das Tool [Synergy_eFuse/eFuse_GUI.ps1](Synergy_eFuse/eFuse_GUI.ps1) löst gezielt einen **eFuse** auf einer Synergy-Komponente aus (entspricht `Reset-OVEnclosureDevice -eFuse`). Es verwendet im Gegensatz zur Haupt-GUI das offizielle **HPEOneView**-PowerShell-Modul.

Ablauf:

1. Login (Benutzername / Kennwort)
2. Auswahl der Appliance aus `Synergy_eFuse/Oneview.txt` (eine Appliance pro Zeile)
3. Verbinden – Enclosures werden automatisch geladen
4. Auswahl von **Enclosure**, **Komponente** und **Slot-Nummer**
   - Komponenten: `Device (BladeServer)`, `ICM (VirtualConnectModule)`, `Appliance (Composer)`, `FLM (FrameLinkModule)`
5. Bestätigungsdialog mit allen Parametern
6. Ausführung und Protokollierung im integrierten Log-Fenster

Voraussetzung: Installiertes HPE OneView PowerShell-Modul (`Install-Module HPEOneView.*`).

## Weitere eigenständige Tools

- **Appliance Update** (`Oneview_Update/OneView_Update_GUI.ps1`) – Geführter Update-Workflow für OneView/Synergy. Benötigt das `HPEOneView.1000`-Modul.
- **Config Backup** (`Oneview_Backup/`) – Manuelles Backup über `Backup_OneView_GUI.ps1` sowie geplante Backups via `Backup_OneView_Scheduled.ps1` (konfigurierbar über `Backup_OneView_TaskConfig_GUI.ps1`).
- **User Manager** (`Oneview_UserManager/OneView-UserManager-GUI.ps1`) – Direkte REST-API-Verwaltung von OneView-Benutzern, parallel auf mehreren Appliances; unterstützt unterschiedliche API-Versionen (OV 6.60 / 11.10) durch automatische Erkennung.
- **Alerts** (`Oneview_Alerts/`) – `Oneview_Alerts_GUI.ps1` für interaktive Auswertung, `Oneview_Alerts_Scheduled.ps1` für regelmässige Reports, Konfiguration über `Oneview_Alerts_TaskConfig_GUI.ps1`.

## Projektstruktur

```
OneView-VLAN-Manager/
├── OneView-Manager-GUI.ps1            # Haupt-GUI (WinForms) – VLANs, Network Sets, SP/SPT
├── OneView-NetworkManager-GUI.ps1     # Alternative Netzwerk-GUI
├── Create-EthernetNetworks.ps1        # Ethernet Networks erstellen/aktualisieren
├── Create-NetworkSets.ps1             # Network Sets erstellen/aktualisieren
├── Export-EthernetNetworks.ps1        # Ethernet Networks nach Excel exportieren
├── Export-NetworkSets.ps1             # Network Sets nach Excel exportieren
├── Export-ServerProfiles.ps1          # Server Profiles als JSON exportieren
├── Import-ServerProfiles.ps1          # Server Profiles aus JSON importieren
├── Export-ServerProfileTemplates.ps1  # Server Profile Templates als JSON exportieren
├── Import-ServerProfileTemplates.ps1  # Server Profile Templates aus JSON importieren
├── New-ExcelTemplate.ps1              # Excel-Vorlage mit Beispieldaten generieren
├── Change_Oneview_Password.ps1        # Hilfsskript zum Ändern des OneView-Passworts
├── config.json                        # Konfiguration (API-Version, Defaults)
├── Appliances.txt                     # Liste der OneView Appliances mit Typ
├── README.md                          # Diese Datei
│
├── Oneview_Update/                    # Eigenständiges Tool: Appliance-Update
│   └── OneView_Update_GUI.ps1
├── Oneview_Backup/                    # Eigenständiges Tool: Config-Backup
│   ├── Backup_OneView_GUI.ps1
│   ├── Backup_OneView_Scheduled.ps1
│   └── Backup_OneView_TaskConfig_GUI.ps1
├── Oneview_UserManager/               # Eigenständiges Tool: User Manager
│   └── OneView-UserManager-GUI.ps1
├── Oneview_Alerts/                    # Eigenständiges Tool: Alerts
│   ├── Oneview_Alerts_GUI.ps1
│   ├── Oneview_Alerts_Scheduled.ps1
│   └── Oneview_Alerts_TaskConfig_GUI.ps1
└── Synergy_eFuse/                     # Eigenständiges Tool: Synergy eFuse
    └── eFuse_GUI.ps1
```

## Voraussetzungen

| Komponente | Mindestversion |
|---|---|
| PowerShell | 7.x (Windows) – alle GUIs und Tools |
| Modul `ImportExcel` | aktuell (wird bei Bedarf automatisch installiert) |
| Modul `HPEOneView.*` | nur für `Synergy_eFuse` und `Oneview_Update` erforderlich |
| HPE OneView Appliance | API Version 3800 (OneView 6.60) API Version 8400 (OneView 11.10) |

## Einrichtung

### 1. Appliances konfigurieren

Bearbeiten Sie die Datei `Appliances.txt` – eine Appliance pro Zeile im Format `Hostname ; Typ`:

```
# Zeilen mit # werden ignoriert
# Format: Hostname ; Typ (z.B. ESXi, VDI)
oneview01.domain.local ; ESXi
oneview02.domain.local ; ESXi
oneview03.domain.local ; VDI
oneview04.domain.local ; VDI
```

Der Typ ist optional. Zeilen ohne Semikolon werden als Appliance ohne Typ behandelt:

```
oneview-legacy.domain.local
```

### 2. Konfiguration anpassen (optional)

Die Datei `config.json` enthält Standardwerte für neue Netzwerke und Network Sets:

```json
{
    "ApiVersion": 8000,
    "ExcelFilePath": ".\\VLANs.xlsx",
    "ExcelSheetName": "VLANs",
    "DefaultSettings": {
        "Purpose": "General",
        "SmartLink": true,
        "PrivateNetwork": false,
        "EthernetNetworkType": "Tagged",
        "PreferredBandwidthGb": 2.5,
        "MaximumBandwidthGb": 50
    },
    "NetworkSetExcelFilePath": ".\\NetworkSets.xlsx",
    "NetworkSetExcelSheetName": "NetworkSets",
    "NetworkSetDefaultSettings": {
        "PreferredBandwidthGb": 2.5,
        "MaximumBandwidthGb": 20
    }
}
```

Die `ApiVersion` dient als Fallback, falls die automatische Erkennung fehlschlägt.

### 3. GUI starten

```powershell
.\OneView-Manager-GUI.ps1
```

Benutzername und Kennwort werden in der GUI eingegeben und sicher an die Subprozesse übergeben.

### 4. Excel-Vorlage erzeugen (optional)

```powershell
.\New-ExcelTemplate.ps1
```

Erstellt eine `VLANs.xlsx` mit Beispieldaten und allen erforderlichen Spalten.

## Excel-Format

### Ethernet Networks (VLANs.xlsx)

| Spalte | Pflicht | Beschreibung | Gültige Werte |
|---|---|---|---|
| `NetworkName` | ✅ | Name des Ethernet Networks | Freitext |
| `VlanId` | ✅ | VLAN-ID | 1–4094 (Tagged), 0 (Untagged) |
| `Purpose` | ❌ | Zweck des Netzwerks | `General`, `Management`, `VMMigration`, `FaultTolerance`, `ISCSI` |
| `EthernetNetworkType` | ❌ | Netzwerktyp | `Tagged`, `Untagged`, `Tunnel` |
| `SmartLink` | ❌ | SmartLink aktivieren | `True` / `False` |
| `PrivateNetwork` | ❌ | Privates Netzwerk | `True` / `False` |
| `PreferredBandwidthGb` | ❌ | Typische Bandbreite (Gb) | Dezimalzahl (z.B. 2.5) |
| `MaximumBandwidthGb` | ❌ | Maximale Bandbreite (Gb) | Dezimalzahl (z.B. 50) |
| `Scope` | ❌ | Scope-Zuweisung | Freitext |
| `NetworkSet` | ❌ | Network Set Zuweisung | Name(n), mehrere mit "; " getrennt |

### Network Sets (NetworkSets.xlsx)

| Spalte | Pflicht | Beschreibung | Gültige Werte |
|---|---|---|---|
| `NetworkSetName` | ✅ | Name des Network Sets | Freitext |
| `Networks` | ✅ | Zugeordnete Netzwerke | Name(n), mehrere mit "; " getrennt |
| `PreferredBandwidthGb` | ❌ | Typische Bandbreite (Gb) | Dezimalzahl |
| `MaximumBandwidthGb` | ❌ | Maximale Bandbreite (Gb) | Dezimalzahl |

## Kommandozeilen-Nutzung (ohne GUI)

Die Scripts können auch direkt aufgerufen werden:

```powershell
# Ethernet Networks erstellen
.\Create-EthernetNetworks.ps1 -ConfigPath ".\config.json"

# Ethernet Networks exportieren
.\Export-EthernetNetworks.ps1 -ConfigPath ".\config.json" -OutputPath ".\Backup.xlsx"

# Network Sets erstellen
.\Create-NetworkSets.ps1 -ConfigPath ".\config.json"

# Network Sets exportieren
.\Export-NetworkSets.ps1 -ConfigPath ".\config.json" -OutputPath ".\NS_Backup.xlsx"
```

## Verwendete API-Endpunkte

| Methode | URI | Beschreibung |
|---|---|---|
| `GET` | `/rest/version` | API-Version der Appliance ermitteln |
| `POST` | `/rest/login-sessions` | Authentifizierung |
| `DELETE` | `/rest/login-sessions` | Session abmelden |
| `GET` | `/rest/ethernet-networks` | Alle Ethernet Networks abrufen |
| `POST` | `/rest/ethernet-networks` | Neues Ethernet Network erstellen |
| `PUT` | `/rest/ethernet-networks/{id}` | Ethernet Network aktualisieren |
| `GET` | `/rest/connection-templates/{id}` | Connection Template abrufen |
| `PUT` | `/rest/connection-templates/{id}` | Bandwidth aktualisieren |
| `GET` | `/rest/network-sets` | Alle Network Sets abrufen |
| `POST` | `/rest/network-sets` | Neues Network Set erstellen |
| `PUT` | `/rest/network-sets/{id}` | Network Set aktualisieren |
| `GET` | `/rest/scopes` | Scopes abrufen |
| `GET` | `/rest/server-profiles` | Alle Server Profiles abrufen |
| `POST` | `/rest/server-profiles` | Neues Server Profile erstellen |
| `PUT` | `/rest/server-profiles/{id}` | Server Profile aktualisieren |
| `DELETE` | `/rest/server-profiles/{id}` | Server Profile löschen |
| `GET` | `/rest/server-profile-templates` | Alle Server Profile Templates abrufen |
| `POST` | `/rest/server-profile-templates` | Neues Template erstellen |
| `PUT` | `/rest/server-profile-templates/{id}` | Template aktualisieren |
| `DELETE` | `/rest/server-profile-templates/{id}` | Template löschen |
| `GET` | `/rest/server-profile-templates/{id}/new-profile` | Profil aus Template generieren |
| `GET` | `/rest/server-hardware-types` | Hardware-Typen abrufen |
| `GET` | `/rest/enclosure-groups` | Enclosure Groups abrufen |

## API-Versionen (Referenz)

| OneView Version | API Version |
|---|---|
| 6.60 | 3800 |
| 10.00 | 7600 |
| 10.20 | 8000 |
| 11.10 | 8400 |

## Sicherheitshinweise

- **Zertifikate**: Die Scripts verwenden `-SkipCertificateCheck` für selbst-signierte Zertifikate. In Produktionsumgebungen sollte dies durch ein gültiges Zertifikat ersetzt werden.
- **Anmeldedaten**: Benutzername und Passwort werden in der GUI eingegeben und über Umgebungsvariablen an die Subprozesse übergeben – keine Passwörter im Klartext in Dateien.
- **Session-Management**: Sessions werden im `finally`-Block immer sauber abgemeldet.

## Fehlerbehandlung

- **Duplikatserkennung**: Bereits existierende Netzwerke/Network Sets werden erkannt und bei Bedarf aktualisiert statt doppelt erstellt.
- **Validierung**: VLAN-IDs, Purpose und EthernetNetworkType werden vor der Erstellung validiert.
- **Protokollierung**: Jede Operation erzeugt eine Logdatei im Verzeichnis `Logs/`. Multi-Appliance-Operationen schreiben ein konsolidiertes Log.
- **Live-Feedback**: Die GUI zeigt den Fortschritt in Echtzeit im Protokollfenster an.

## Lizenz

Intern – nicht zur Weitergabe bestimmt.
