# HPE OneView – Ethernet Networks aus Excel erstellen

Dieses Projekt erstellt **Ethernet Networks** in HPE OneView basierend auf einer Excel-Datei. Es verwendet ausschliesslich die **HPE OneView RESTful API** (ohne das HPE OneView PowerShell-Modul).

## Projektstruktur

```
OneView_VLAN_Projekt/
├── Create-EthernetNetworks.ps1   # Hauptskript – erstellt die Ethernet Networks
├── New-ExcelTemplate.ps1         # Hilfsskript – erzeugt eine Excel-Vorlage
├── config.json                   # Konfigurationsdatei (Appliances, API-Version, Pfade)
├── VLANs.xlsx                    # Excel-Datei mit VLAN-Definitionen (wird generiert)
└── README.md                     # Diese Datei
```

## Voraussetzungen

| Komponente | Mindestversion |
|---|---|
| PowerShell | 7.x |
| Modul `ImportExcel` | aktuell |
| HPE OneView Appliance | API Version 5600+ (OneView 8.50+) |

Das Modul `ImportExcel` wird bei Bedarf automatisch installiert.

## Schnellstart

### 1. Excel-Vorlage erzeugen (optional)

```powershell
.\New-ExcelTemplate.ps1
```

Erstellt eine `VLANs.xlsx` mit Beispieldaten und allen erforderlichen Spalten.

### 2. Konfiguration anpassen

Bearbeiten Sie die Datei `config.json`:

```json
{
    "OneViewAppliances": [
        {
            "Name": "OneView-Prod-01",
            "Hostname": "oneview01.domain.local",
            "Description": "Produktiv OneView Appliance 1"
        }
    ],
    "ApiVersion": 8000,
    "ExcelFilePath": ".\\VLANs.xlsx",
    "ExcelSheetName": "VLANs",
    "DefaultSettings": {
        "Purpose": "General",
        "SmartLink": true,
        "PrivateNetwork": false,
        "EthernetNetworkType": "Tagged",
        "BandwidthTypicalMbps": 2500,
        "BandwidthMaximumMbps": 10000
    }
}
```

**Wichtige Einstellungen:**

- **OneViewAppliances** – Eine oder mehrere Appliances (alle werden nacheinander abgearbeitet)
- **ApiVersion** – Die API-Version Ihrer OneView-Installation (siehe Tabelle unten)
- **ExcelFilePath** – Pfad zur Excel-Datei (relativ oder absolut)
- **DefaultSettings** – Standardwerte, falls eine Spalte in der Excel-Datei leer ist

### 3. Excel-Datei befüllen

Die Excel-Datei benötigt folgende Spalten:

| Spalte | Pflicht | Beschreibung | Gültige Werte |
|---|---|---|---|
| `NetworkName` | ✅ | Name des Ethernet Networks | Freitext |
| `VlanId` | ✅ | VLAN-ID | 1–4094 (Tagged), 0 (Untagged) |
| `Purpose` | ❌ | Zweck des Netzwerks | `General`, `Management`, `VMMigration`, `FaultTolerance`, `ISCSI` |
| `EthernetNetworkType` | ❌ | Netzwerktyp | `Tagged`, `Untagged`, `Tunnel` |
| `SmartLink` | ❌ | SmartLink aktivieren | `True` / `False` |
| `PrivateNetwork` | ❌ | Privates Netzwerk | `True` / `False` |
| `BandwidthTypicalMbps` | ❌ | Typische Bandbreite (Mbps) | Ganzzahl |
| `BandwidthMaximumMbps` | ❌ | Maximale Bandbreite (Mbps) | Ganzzahl |
| `Subnet` | ❌ | Subnetz (nur informativ) | z.B. `10.0.100.0/24` |
| `Description` | ❌ | Beschreibung | Freitext |

### 4. Skript ausführen

```powershell
# Produktiv-Lauf
.\Create-EthernetNetworks.ps1

# Simulation (keine Änderungen)
.\Create-EthernetNetworks.ps1 -WhatIf

# Mit benutzerdefinierter Konfiguration
.\Create-EthernetNetworks.ps1 -ConfigPath "C:\Config\prod-config.json"
```

## Funktionsweise

```
┌──────────────┐     ┌───────────────────┐     ┌──────────────────────┐
│  config.json │     │    VLANs.xlsx     │     │   OneView Appliance  │
│  (Appliance  │     │  (VLAN-Daten)     │     │                      │
│   Settings)  │     │                   │     │  REST API:           │
└──────┬───────┘     └────────┬──────────┘     │  POST /rest/         │
       │                      │                │   login-sessions     │
       ▼                      ▼                │  GET  /rest/         │
┌──────────────────────────────────────┐       │   ethernet-networks  │
│   Create-EthernetNetworks.ps1        │──────▶│  POST /rest/         │
│                                      │       │   ethernet-networks  │
│  1. Config laden                     │       │  PUT  /rest/         │
│  2. Excel importieren & validieren   │       │   connection-        │
│  3. Login via REST API               │       │   templates/{id}     │
│  4. Duplikate prüfen                 │       │  DELETE /rest/       │
│  5. Netzwerke erstellen              │       │   login-sessions     │
│  6. Bandwidth setzen                 │       └──────────────────────┘
│  7. Session abmelden                 │
│  8. Protokoll speichern              │
└──────────────────────────────────────┘
```

## Verwendete API-Endpunkte

| Methode | URI | Beschreibung |
|---|---|---|
| `POST` | `/rest/login-sessions` | Authentifizierung, gibt `sessionID` zurück |
| `DELETE` | `/rest/login-sessions` | Session abmelden |
| `GET` | `/rest/ethernet-networks` | Alle Ethernet Networks abrufen |
| `POST` | `/rest/ethernet-networks` | Neues Ethernet Network erstellen |
| `GET` | `/rest/connection-templates/{id}` | Connection Template abrufen |
| `PUT` | `/rest/connection-templates/{id}` | Bandwidth-Einstellungen aktualisieren |

## API-Versionen (Referenz)

| OneView Version | API Version |
|---|---|
| 8.50 | 5600 |
| 9.00 | 6600 |
| 10.00 | 7600 |
| 10.20 | 8000 |

## Sicherheitshinweise

- **Zertifikate**: Das Skript verwendet `-SkipCertificateCheck` für selbst-signierte Zertifikate. In Produktionsumgebungen sollte dies durch ein gültiges Zertifikat ersetzt werden.
- **Anmeldedaten**: Das Skript fragt Benutzername/Passwort interaktiv via `Get-Credential` ab – keine Passwörter im Klartext.
- **Session-Management**: Die Session wird im `finally`-Block immer sauber abgemeldet.

## Fehlerbehandlung

- **Duplikatserkennung**: Bereits existierende Netzwerke werden automatisch übersprungen.
- **Validierung**: VLAN-IDs, Purpose und EthernetNetworkType werden vor der Erstellung validiert.
- **Protokoll**: Jede Ausführung erzeugt eine Logdatei (`VLAN_Import_YYYYMMDD_HHmmss.log`).
- **WhatIf**: Mit `-WhatIf` kann ein Trockenlauf durchgeführt werden.

## Lizenz

Intern – nicht zur Weitergabe bestimmt.
