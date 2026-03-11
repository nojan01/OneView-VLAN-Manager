<#
.SYNOPSIS
    Erzeugt eine Excel-Vorlage für die VLAN-Konfiguration.

.DESCRIPTION
    Dieses Hilfsskript erstellt eine Beispiel-Excel-Datei (VLANs.xlsx),
    die als Vorlage für den Import von Ethernet Networks in HPE OneView dient.
    Voraussetzung: Das PowerShell-Modul "ImportExcel" muss installiert sein.

.NOTES
    Autor:  OneView VLAN Projekt
    Datum:  2026-02-06
#>

# Modul prüfen / installieren
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Das Modul 'ImportExcel' wird installiert..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$excelPath = Join-Path $scriptDir "VLANs.xlsx"

# Beispieldaten – alle Felder gemäss OneView "Create Network" Dialog
$sampleData = @(
    [PSCustomObject]@{
        NetworkName          = "VLAN_100_Management"
        VlanId               = 100
        EthernetNetworkType  = "Tagged"
        Purpose              = "Management"
        SmartLink            = $true
        PrivateNetwork       = $false
        PreferredBandwidthGb = 2.5
        MaximumBandwidthGb   = 50
        Scope                = "Scope_Prod"
        NetworkSet           = "Management_Set"
        IPv4SubnetId         = ""
        IPv6SubnetId         = ""
        Description          = "Management Netzwerk"
    },
    [PSCustomObject]@{
        NetworkName          = "VLAN_200_Production"
        VlanId               = 200
        EthernetNetworkType  = "Tagged"
        Purpose              = "General"
        SmartLink            = $true
        PrivateNetwork       = $false
        PreferredBandwidthGb = 2.5
        MaximumBandwidthGb   = 50
        Scope                = "Scope_Prod"
        NetworkSet           = "Production_Set"
        IPv4SubnetId         = ""
        IPv6SubnetId         = ""
        Description          = "Produktionsnetzwerk"
    },
    [PSCustomObject]@{
        NetworkName          = "VLAN_201_AppServer"
        VlanId               = 201
        EthernetNetworkType  = "Tagged"
        Purpose              = "General"
        SmartLink            = $true
        PrivateNetwork       = $false
        PreferredBandwidthGb = 2.5
        MaximumBandwidthGb   = 50
        Scope                = "Scope_Prod"
        NetworkSet           = "Production_Set"
        IPv4SubnetId         = ""
        IPv6SubnetId         = ""
        Description          = "Application Server Netzwerk"
    },
    [PSCustomObject]@{
        NetworkName          = "VLAN_300_Backup"
        VlanId               = 300
        EthernetNetworkType  = "Tagged"
        Purpose              = "General"
        SmartLink            = $true
        PrivateNetwork       = $false
        PreferredBandwidthGb = 5
        MaximumBandwidthGb   = 50
        Scope                = ""
        NetworkSet           = ""
        IPv4SubnetId         = ""
        IPv6SubnetId         = ""
        Description          = "Backup Netzwerk"
    },
    [PSCustomObject]@{
        NetworkName          = "VLAN_400_DMZ"
        VlanId               = 400
        EthernetNetworkType  = "Tagged"
        Purpose              = "General"
        SmartLink            = $true
        PrivateNetwork       = $false
        PreferredBandwidthGb = 2.5
        MaximumBandwidthGb   = 10
        Scope                = "Scope_DMZ"
        NetworkSet           = "DMZ_Set"
        IPv4SubnetId         = ""
        IPv6SubnetId         = ""
        Description          = "DMZ Netzwerk"
    },
    [PSCustomObject]@{
        NetworkName          = "Untagged_iSCSI"
        VlanId               = 0
        EthernetNetworkType  = "Untagged"
        Purpose              = "ISCSI"
        SmartLink            = $false
        PrivateNetwork       = $false
        PreferredBandwidthGb = 10
        MaximumBandwidthGb   = 20
        Scope                = ""
        NetworkSet           = ""
        IPv4SubnetId         = ""
        IPv6SubnetId         = ""
        Description          = "iSCSI Storage Netzwerk (Untagged)"
    }
)

# Excel exportieren (ohne Titelzeile, damit Zeile 1 = Spaltenüberschriften)
$sampleData | Export-Excel -Path $excelPath `
    -WorksheetName "VLANs" `
    -AutoSize `
    -FreezeTopRow `
    -BoldTopRow

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Excel-Vorlage erfolgreich erstellt!" -ForegroundColor Green
Write-Host " Pfad: $excelPath" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Spalten-Beschreibung:" -ForegroundColor Yellow
Write-Host "  NetworkName          - Name des Ethernet Networks in OneView"
Write-Host "  VlanId               - VLAN-ID (1-4094 für Tagged, 0 für Untagged)"
Write-Host "  EthernetNetworkType  - Typ: Tagged, Untagged oder Tunnel"
Write-Host "  Purpose              - Zweck: General, Management, VMMigration,"
Write-Host "                         FaultTolerance, ISCSI"
Write-Host "  SmartLink            - SmartLink aktivieren (True/False)"
Write-Host "  PrivateNetwork       - Privates Netzwerk (True/False)"
Write-Host "  PreferredBandwidthGb - Bevorzugte Bandbreite in Gb/s (z.B. 2.5)"
Write-Host "  MaximumBandwidthGb   - Maximale Bandbreite in Gb/s (z.B. 50)"
Write-Host "  Scope                - Scope-Name (optional, muss in OneView existieren)"
Write-Host "  NetworkSet           - Network Set Name (optional, muss in OneView existieren)"
Write-Host "  IPv4SubnetId         - IPv4 Subnet ID (optional, URI aus OneView)"
Write-Host "  IPv6SubnetId         - IPv6 Subnet ID (optional, URI aus OneView)"
Write-Host "  Description          - Beschreibung (optional)"
Write-Host ""
