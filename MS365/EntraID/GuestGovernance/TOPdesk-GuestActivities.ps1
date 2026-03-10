<#
.SYNOPSIS
    TOPdesk Integration fuer M365 Guest Governance
.DESCRIPTION
    Erstellt operative Aktivitaeten in TOPdesk SaaS basierend auf Audit-CSV-Daten
    des M365-GuestGovernance.ps1 Skripts.

    Features:
    - Liest Audit-CSV des Governance-Tools als Eingabe
    - Erstellt operative Aktivitaeten via TOPdesk REST API
    - Flexible Konfiguration (Tenant-URL, Credentials, Kategorien)
    - WhatIf-Modus fuer Testlaeufe
    - Bulk-Erstellung fuer mehrere Gastkonten
    - Filtert nach Schweregrad, Freemailer-Status, Inaktivitaet etc.
    - Erstellt eine Sammel-Aktivitaet oder einzelne pro Konto
.NOTES
    Voraussetzung: TOPdesk API-Konto mit Berechtigung fuer Operations Management
    API-Dokumentation: https://developers.topdesk.com/explorer/?page=operations-management
.LINK
    https://developers.topdesk.com/
#>

#Requires -Version 5.1

[CmdletBinding()]
param()

# ============================================================================
# Konfiguration
# ============================================================================
$Script:TopDeskConfig = @{
    # TOPdesk SaaS Instanz-URL (ohne /tas/api)
    # Wird beim ersten Start abgefragt, wenn nicht gesetzt
    BaseUrl            = ""

    # API-Endpunkt (wird automatisch zusammengesetzt)
    ApiBase            = "/tas/api"

    # Operative Aktivitaeten Endpunkt
    ActivityEndpoint   = "/operationalActivities"

    # Standard-Werte fuer neue Aktivitaeten (IDs oder Namen je nach Umgebung)
    # Diese koennen ueber die Konfigurationsdatei oder Parameter ueberschrieben werden
    Defaults           = @{
        # Kategorie fuer Gastkonto-Aktivitaeten (Name oder ID)
        Category       = ""
        Subcategory    = ""
        # Operator / Operatorgruppe (Name oder ID)
        OperatorGroup  = ""
        Operator       = ""
        # Aktivitaetstyp
        ActivityType   = ""
    }

    # Pfad zur Konfigurationsdatei (wird automatisch gesetzt)
    ConfigFilePath     = ""

    # Credential-Cache (Sitzungsbezogen)
    Credential         = $null
}

# ============================================================================
# Konfigurationsdatei-Verwaltung
# ============================================================================

function Get-TopDeskConfigPath {
    <#
    .SYNOPSIS
        Gibt den Pfad zur TOPdesk-Konfigurationsdatei zurueck.
    #>
    return (Join-Path $PSScriptRoot "topdesk_config.json")
}

function Import-TopDeskConfig {
    <#
    .SYNOPSIS
        Laedt die TOPdesk-Konfiguration aus der JSON-Datei.
    #>
    [CmdletBinding()]
    param()

    $configPath = Get-TopDeskConfigPath

    if (Test-Path $configPath) {
        try {
            $json = Get-Content -Path $configPath -Raw | ConvertFrom-Json

            if ($json.BaseUrl)  { $Script:TopDeskConfig.BaseUrl = $json.BaseUrl }
            if ($json.Defaults) {
                if ($json.Defaults.Category)      { $Script:TopDeskConfig.Defaults.Category      = $json.Defaults.Category }
                if ($json.Defaults.Subcategory)    { $Script:TopDeskConfig.Defaults.Subcategory   = $json.Defaults.Subcategory }
                if ($json.Defaults.OperatorGroup)  { $Script:TopDeskConfig.Defaults.OperatorGroup = $json.Defaults.OperatorGroup }
                if ($json.Defaults.Operator)       { $Script:TopDeskConfig.Defaults.Operator      = $json.Defaults.Operator }
                if ($json.Defaults.ActivityType)   { $Script:TopDeskConfig.Defaults.ActivityType  = $json.Defaults.ActivityType }
            }

            Write-Host "[OK] Konfiguration geladen: $configPath" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Warning "Konfigurationsdatei fehlerhaft: $_"
            return $false
        }
    }

    return $false
}

function Export-TopDeskConfig {
    <#
    .SYNOPSIS
        Speichert die aktuelle TOPdesk-Konfiguration in die JSON-Datei.
    #>
    [CmdletBinding()]
    param()

    $configPath = Get-TopDeskConfigPath

    $configObj = [ordered]@{
        BaseUrl  = $Script:TopDeskConfig.BaseUrl
        Defaults = [ordered]@{
            Category      = $Script:TopDeskConfig.Defaults.Category
            Subcategory   = $Script:TopDeskConfig.Defaults.Subcategory
            OperatorGroup = $Script:TopDeskConfig.Defaults.OperatorGroup
            Operator      = $Script:TopDeskConfig.Defaults.Operator
            ActivityType  = $Script:TopDeskConfig.Defaults.ActivityType
        }
    }

    $configObj | ConvertTo-Json -Depth 3 | Out-File -FilePath $configPath -Encoding UTF8
    Write-Host "[OK] Konfiguration gespeichert: $configPath" -ForegroundColor Green
}

function Initialize-TopDeskConfig {
    <#
    .SYNOPSIS
        Initialisiert die TOPdesk-Konfiguration interaktiv, falls nicht vorhanden.
    #>
    [CmdletBinding()]
    param(
        [switch]$Force
    )

    # Bestehende Konfig laden
    $loaded = Import-TopDeskConfig

    if ($loaded -and -not $Force) {
        Write-Host "  URL: $($Script:TopDeskConfig.BaseUrl)" -ForegroundColor Gray
        return $true
    }

    Write-Host "`n--- TOPdesk Konfiguration ---" -ForegroundColor Cyan
    Write-Host "Bitte geben Sie die Verbindungsdaten fuer Ihre TOPdesk-Instanz ein.`n" -ForegroundColor Yellow

    # Tenant-URL
    $currentUrl = $Script:TopDeskConfig.BaseUrl
    $urlPrompt = "TOPdesk URL (z.B. https://firma.topdesk.net)"
    if ($currentUrl) { $urlPrompt += " [$currentUrl]" }
    $inputUrl = Read-Host $urlPrompt
    if ($inputUrl) {
        # Trailing slash und /tas/api entfernen
        $inputUrl = $inputUrl.TrimEnd('/')
        $inputUrl = $inputUrl -replace '/tas/api/?$', ''
        $inputUrl = $inputUrl -replace '/tas/?$', ''
        $Script:TopDeskConfig.BaseUrl = $inputUrl
    }
    elseif (-not $currentUrl) {
        Write-Host "[FEHLER] URL ist erforderlich." -ForegroundColor Red
        return $false
    }

    # Optionale Standard-Werte
    Write-Host "`nOptionale Standard-Werte (Enter = ueberspringen):" -ForegroundColor Yellow

    $catInput = Read-Host "  Kategorie (Name oder ID)"
    if ($catInput) { $Script:TopDeskConfig.Defaults.Category = $catInput }

    $subInput = Read-Host "  Unterkategorie (Name oder ID)"
    if ($subInput) { $Script:TopDeskConfig.Defaults.Subcategory = $subInput }

    $grpInput = Read-Host "  Operatorgruppe (Name oder ID)"
    if ($grpInput) { $Script:TopDeskConfig.Defaults.OperatorGroup = $grpInput }

    $opInput = Read-Host "  Operator (Name oder ID)"
    if ($opInput) { $Script:TopDeskConfig.Defaults.Operator = $opInput }

    $typeInput = Read-Host "  Aktivitaetstyp (Name oder ID)"
    if ($typeInput) { $Script:TopDeskConfig.Defaults.ActivityType = $typeInput }

    # Speichern
    Export-TopDeskConfig
    return $true
}

# ============================================================================
# Authentifizierung
# ============================================================================

function Connect-TopDesk {
    <#
    .SYNOPSIS
        Stellt eine Verbindung zu TOPdesk her (Basic Auth mit Application Password).
    .DESCRIPTION
        TOPdesk API verwendet Basic Authentication.
        Benutzername = Operator Login-Name
        Passwort = Application Password (erstellt im Operator-Profil)

        Credential-Reihenfolge:
        1. Uebergebener -Credential Parameter
        2. Gespeicherte Credentials (DPAPI-verschluesselt)
        3. Interaktive Abfrage via Get-Credential
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [PSCredential]$Credential,

        [switch]$Force
    )

    # 1. Bereits im Speicher?
    if ($Script:TopDeskConfig.Credential -and -not $Force) {
        Write-Host "[OK] Bereits authentifiziert." -ForegroundColor Green
        return $true
    }

    $fromSaved = $false

    if (-not $Credential) {
        # 2. Gespeicherte Credentials laden (DPAPI)
        $savedCred = Get-SavedTopDeskCredential
        if ($savedCred) {
            $Credential = $savedCred
            $fromSaved = $true
        }
    }

    if (-not $Credential) {
        # 3. Interaktive Abfrage
        Write-Host "`n--- TOPdesk Anmeldung ---" -ForegroundColor Cyan
        Write-Host "Benutzername: Operator Login-Name" -ForegroundColor Gray
        Write-Host "Passwort: Application Password (aus Operator-Profil > Autorisierung)" -ForegroundColor Gray
        Write-Host ""
        $Credential = Get-Credential -Message "TOPdesk API Anmeldedaten (Login + Application Password)"
    }

    if (-not $Credential) {
        Write-Host "[FEHLER] Keine Anmeldedaten angegeben." -ForegroundColor Red
        return $false
    }

    # Verbindung testen
    $baseUrl = $Script:TopDeskConfig.BaseUrl
    $testUri = "$baseUrl$($Script:TopDeskConfig.ApiBase)/version"

    $pair = "$($Credential.UserName):$($Credential.GetNetworkCredential().Password)"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)

    try {
        $headers = @{
            Authorization = "Basic $base64"
        }
        $response = Invoke-RestMethod -Uri $testUri -Method GET -Headers $headers -ErrorAction Stop
        Write-Host "[OK] Verbunden mit TOPdesk ($baseUrl)" -ForegroundColor Green
        if ($response.version) {
            Write-Host "  API-Version: $($response.version)" -ForegroundColor Gray
        }

        # Im Speicher cachen
        $Script:TopDeskConfig.Credential = $Credential

        # Anbieten zu speichern (nur wenn nicht bereits aus Datei geladen)
        if (-not $fromSaved) {
            $credPath = Get-TopDeskCredentialPath
            $alreadySaved = Test-Path $credPath

            if (-not $alreadySaved) {
                Write-Host ""
                $saveChoice = Read-Host "Credentials sicher speichern fuer zukuenftige Ausfuehrungen? (Ja/Nein)"
                if ($saveChoice -in @("Ja", "J", "ja", "j")) {
                    Save-TopDeskCredential -Credential $Credential
                }
            }
        }

        return $true
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
        }

        if ($statusCode -eq 401) {
            Write-Host "[FEHLER] Authentifizierung fehlgeschlagen (401). Bitte pruefen Sie Login und Application Password." -ForegroundColor Red
            # Wenn gespeicherte Credentials fehlschlagen: loeschen und neu fragen
            if ($fromSaved) {
                Write-Host "  Gespeicherte Credentials sind ungueltig. Datei wird geloescht." -ForegroundColor DarkYellow
                Remove-TopDeskCredential | Out-Null
                Write-Host "  Bitte starten Sie das Skript erneut." -ForegroundColor Yellow
            }
        }
        elseif ($statusCode -eq 403) {
            Write-Host "[FEHLER] Zugriff verweigert (403). API-Berechtigungen pruefen." -ForegroundColor Red
        }
        else {
            Write-Host "[FEHLER] Verbindung fehlgeschlagen: $_" -ForegroundColor Red
        }
        return $false
    }
}

function Get-TopDeskAuthHeaders {
    <#
    .SYNOPSIS
        Gibt die Authentifizierungs-Header fuer TOPdesk-API-Aufrufe zurueck.
    #>
    $cred = $Script:TopDeskConfig.Credential
    if (-not $cred) {
        throw "Nicht authentifiziert. Bitte zuerst Connect-TopDesk ausfuehren."
    }

    $pair  = "$($cred.UserName):$($cred.GetNetworkCredential().Password)"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)

    return @{
        Authorization  = "Basic $base64"
        'Content-Type' = 'application/json; charset=utf-8'
    }
}

# ============================================================================
# Credential-Verwaltung (DPAPI-verschluesselt)
# ============================================================================

function Get-TopDeskCredentialPath {
    <#
    .SYNOPSIS
        Gibt den Pfad zur verschluesselten Credential-Datei zurueck.
    #>
    return (Join-Path $PSScriptRoot "topdesk.credential")
}

function Save-TopDeskCredential {
    <#
    .SYNOPSIS
        Speichert TOPdesk-Credentials sicher via DPAPI (Windows Data Protection API).
    .DESCRIPTION
        Die Credentials werden als verschluesselte XML-Datei gespeichert.
        Die Verschluesselung ist an den aktuellen Windows-Benutzer UND die
        Maschine gebunden. Die Datei kann nur vom selben Benutzer auf
        derselben Maschine wieder entschluesselt werden.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCredential]$Credential
    )

    $credPath = Get-TopDeskCredentialPath

    try {
        $Credential | Export-Clixml -Path $credPath -Force
        Write-Host "[OK] Credentials sicher gespeichert: $credPath" -ForegroundColor Green
        Write-Host "  Verschluesselung: DPAPI (Benutzer + Maschine)" -ForegroundColor Gray
        Write-Host "  Benutzername: $($Credential.UserName)" -ForegroundColor Gray
        return $true
    }
    catch {
        Write-Warning "Credentials konnten nicht gespeichert werden: $_"
        return $false
    }
}

function Get-SavedTopDeskCredential {
    <#
    .SYNOPSIS
        Laedt gespeicherte TOPdesk-Credentials aus der verschluesselten Datei.
    .DESCRIPTION
        Gibt $null zurueck wenn keine Datei vorhanden oder die Entschluesselung
        fehlschlaegt (anderer Benutzer oder andere Maschine).
    #>
    [CmdletBinding()]
    param()

    $credPath = Get-TopDeskCredentialPath

    if (-not (Test-Path $credPath)) {
        return $null
    }

    try {
        $credential = Import-Clixml -Path $credPath
        if ($credential -is [PSCredential]) {
            Write-Host "[OK] Gespeicherte Credentials geladen (Benutzer: $($credential.UserName))" -ForegroundColor Green
            return $credential
        }
        Write-Warning "Credential-Datei hat ungueltiges Format."
        return $null
    }
    catch {
        Write-Warning "Gespeicherte Credentials konnten nicht entschluesselt werden."
        Write-Host "  Moegliche Ursache: Anderer Windows-Benutzer oder andere Maschine." -ForegroundColor DarkYellow
        Write-Host "  Loeschen Sie die Datei und speichern Sie die Credentials erneut." -ForegroundColor DarkYellow
        return $null
    }
}

function Remove-TopDeskCredential {
    <#
    .SYNOPSIS
        Loescht die gespeicherte Credential-Datei.
    #>
    [CmdletBinding()]
    param()

    $credPath = Get-TopDeskCredentialPath

    if (Test-Path $credPath) {
        Remove-Item -Path $credPath -Force
        Write-Host "[OK] Gespeicherte Credentials geloescht." -ForegroundColor Green
        $Script:TopDeskConfig.Credential = $null
        return $true
    }
    else {
        Write-Host "Keine gespeicherten Credentials vorhanden." -ForegroundColor Yellow
        return $false
    }
}

# ============================================================================
# TOPdesk API-Hilfsfunktionen
# ============================================================================

function Invoke-TopDeskApi {
    <#
    .SYNOPSIS
        Fuehrt einen API-Aufruf gegen TOPdesk aus.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Endpoint,

        [Parameter()]
        [ValidateSet("GET", "POST", "PATCH", "PUT", "DELETE")]
        [string]$Method = "GET",

        [Parameter()]
        [object]$Body,

        [Parameter()]
        [hashtable]$QueryParams = @{}
    )

    $baseUrl = $Script:TopDeskConfig.BaseUrl
    $apiBase = $Script:TopDeskConfig.ApiBase
    $uri = "$baseUrl$apiBase$Endpoint"

    # Query-Parameter anfuegen
    if ($QueryParams.Count -gt 0) {
        $queryString = ($QueryParams.GetEnumerator() | ForEach-Object {
            "$([System.Uri]::EscapeDataString($_.Key))=$([System.Uri]::EscapeDataString($_.Value))"
        }) -join '&'
        $uri += "?$queryString"
    }

    $headers = Get-TopDeskAuthHeaders

    $params = @{
        Uri     = $uri
        Method  = $Method
        Headers = $headers
    }

    if ($Body) {
        $jsonBody = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
        $params.Body = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)
    }

    try {
        $response = Invoke-RestMethod @params -ErrorAction Stop
        return $response
    }
    catch {
        $statusCode = $null
        $errorBody = $null

        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $errorBody = $reader.ReadToEnd()
                $reader.Close()
            }
            catch {}
        }

        $errorMsg = "TOPdesk API Fehler"
        if ($statusCode) { $errorMsg += " ($statusCode)" }
        if ($errorBody) { $errorMsg += ": $errorBody" }
        else { $errorMsg += ": $($_.Exception.Message)" }

        throw $errorMsg
    }
}

function Get-TopDeskOperationalSettings {
    <#
    .SYNOPSIS
        Ruft die verfuegbaren Einstellungen fuer operative Aktivitaeten ab
        (Kategorien, Unterkategorien, Status, etc.).
    #>
    [CmdletBinding()]
    param()

    try {
        $settings = Invoke-TopDeskApi -Endpoint "$($Script:TopDeskConfig.ActivityEndpoint)/settings"
        return $settings
    }
    catch {
        Write-Warning "Einstellungen konnten nicht abgerufen werden: $_"
        return $null
    }
}

# ============================================================================
# Audit-CSV Import
# ============================================================================

function Import-GuestAuditCsv {
    <#
    .SYNOPSIS
        Importiert die Audit-CSV des M365-GuestGovernance Skripts.
    .PARAMETER CsvPath
        Pfad zur CSV-Datei. Wenn nicht angegeben, wird die neueste CSV im Reports-Ordner gesucht.
    .PARAMETER Filter
        Filter fuer die Ergebnisse: Alle, Inaktiv, Abgelaufen, Kritisch, Freemailer, Deaktiviert
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$CsvPath,

        [Parameter()]
        [ValidateSet("Alle", "Inaktiv", "Abgelaufen", "Kritisch", "Freemailer", "Deaktiviert")]
        [string]$Filter = "Alle"
    )

    # CSV-Pfad ermitteln
    if (-not $CsvPath) {
        $reportDir = Join-Path $PSScriptRoot "Reports"
        if (Test-Path $reportDir) {
            $csvFiles = Get-ChildItem -Path $reportDir -Filter "GuestAudit_*.csv" -File |
                Sort-Object LastWriteTime -Descending
            if ($csvFiles.Count -gt 0) {
                $CsvPath = $csvFiles[0].FullName
                Write-Host "[Auto] Neueste Audit-CSV: $CsvPath" -ForegroundColor Gray
            }
        }
    }

    if (-not $CsvPath -or -not (Test-Path $CsvPath)) {
        Write-Host "[FEHLER] Keine Audit-CSV gefunden." -ForegroundColor Red
        Write-Host "Bitte zuerst M365-GuestGovernance.ps1 ausfuehren oder -CsvPath angeben." -ForegroundColor Yellow
        return $null
    }

    # CSV importieren
    $data = Import-Csv -Path $CsvPath -Delimiter ";"

    $totalCount = ($data | Measure-Object).Count
    Write-Host "[OK] $totalCount Eintraege aus CSV geladen." -ForegroundColor Green

    # Filter anwenden
    $filtered = switch ($Filter) {
        "Inaktiv"     { $data | Where-Object { $_.IsInactive -eq "True" } }
        "Abgelaufen"  { $data | Where-Object { $_.IsExpired -eq "True" } }
        "Kritisch"    { $data | Where-Object { $_.Severity -eq "Hoch" } }
        "Freemailer"  { $data | Where-Object { $_.IsFreemailer -eq "True" } }
        "Deaktiviert" { $data | Where-Object { $_.AccountEnabled -eq "False" } }
        default       { $data }
    }

    $filteredCount = ($filtered | Measure-Object).Count
    if ($Filter -ne "Alle") {
        Write-Host "[Filter] '$Filter': $filteredCount von $totalCount Eintraegen" -ForegroundColor Cyan
    }

    return @{
        Data      = $filtered
        CsvPath   = $CsvPath
        Filter    = $Filter
        Total     = $totalCount
        Filtered  = $filteredCount
    }
}

# ============================================================================
# Operative Aktivitaeten erstellen
# ============================================================================

function New-GuestActivityDescription {
    <#
    .SYNOPSIS
        Erzeugt den Beschreibungstext fuer eine operative Aktivitaet basierend auf Audit-Daten.
    .PARAMETER Accounts
        Array von Gastkonto-Objekten aus der Audit-CSV.
    .PARAMETER Mode
        'Sammel' fuer eine Sammelaktivitaet oder 'Einzel' fuer individuelle Beschreibung.
    .PARAMETER AuditDate
        Datum des zugrundeliegenden Audits.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Accounts,

        [Parameter()]
        [ValidateSet("Sammel", "Einzel")]
        [string]$Mode = "Sammel",

        [Parameter()]
        [string]$AuditDate = (Get-Date -Format "dd.MM.yyyy")
    )

    if ($Mode -eq "Einzel") {
        $acc = $Accounts[0]
        $lines = @(
            "M365 Gastkonto-Governance: Pruefung erforderlich",
            "",
            "Gastkonto: $($acc.DisplayName)",
            "E-Mail: $($acc.Mail)",
            "Domain: $($acc.GuestDomain)"
        )

        if ($acc.IsFreemailer -eq "True") {
            $lines += "DLP-Risiko: Freemailer-Domain erkannt!"
        }

        $lines += @(
            "Erstellt am: $($acc.CreatedDateTime)",
            "Letzte Aktivitaet: $(if ($acc.LastActivity) { $acc.LastActivity } else { 'Nie' })",
            "Tage inaktiv: $($acc.DaysSinceActivity)",
            "Kontoalter (Tage): $($acc.DaysSinceCreation)",
            "Kontostatus: $(if ($acc.AccountEnabled -eq 'True') { 'Aktiv' } else { 'Deaktiviert' })",
            "Einladungsstatus: $($acc.InvitationStatus)",
            "Eingeladen von: $($acc.InviterName)",
            "Schweregrad: $($acc.Severity)",
            "Flags: $($acc.Flags)",
            "",
            "Audit-Datum: $AuditDate",
            "---",
            "Automatisch erstellt durch M365-GuestGovernance"
        )

        return ($lines -join "`n")
    }
    else {
        # Sammel-Beschreibung
        $totalCount = $Accounts.Count
        $inactiveCount  = ($Accounts | Where-Object { $_.IsInactive -eq "True" }).Count
        $expiredCount   = ($Accounts | Where-Object { $_.IsExpired -eq "True" }).Count
        $criticalCount  = ($Accounts | Where-Object { $_.Severity -eq "Hoch" }).Count
        $freemailerCount = ($Accounts | Where-Object { $_.IsFreemailer -eq "True" }).Count
        $disabledCount  = ($Accounts | Where-Object { $_.AccountEnabled -eq "False" }).Count

        $lines = @(
            "M365 Gastkonto-Governance: $totalCount Konten erfordern Pruefung",
            "",
            "=== Zusammenfassung ===",
            "Gesamt markiert: $totalCount",
            "Inaktiv: $inactiveCount",
            "Abgelaufen: $expiredCount",
            "Kritisch (beides): $criticalCount",
            "Freemailer (DLP): $freemailerCount",
            "Deaktiviert: $disabledCount",
            "",
            "=== Betroffene Konten ==="
        )

        foreach ($acc in $Accounts) {
            $status = if ($acc.AccountEnabled -eq "True") { "Aktiv" } else { "Deaktiviert" }
            $fm = if ($acc.IsFreemailer -eq "True") { " [FREEMAILER]" } else { "" }
            $lines += "- $($acc.DisplayName) ($($acc.Mail)) [$($acc.GuestDomain)]$fm - Schwere: $($acc.Severity) - Status: $status - Inaktiv: $($acc.DaysSinceActivity)d"
        }

        $lines += @(
            "",
            "=== Domain-Verteilung ==="
        )

        $domainGroups = $Accounts | Group-Object -Property GuestDomain | Sort-Object Count -Descending
        foreach ($dg in $domainGroups) {
            $domainName = if ($dg.Name) { $dg.Name } else { "(unbekannt)" }
            $lines += "- $($domainName): $($dg.Count) Konten"
        }

        $lines += @(
            "",
            "Audit-Datum: $AuditDate",
            "---",
            "Automatisch erstellt durch M365-GuestGovernance"
        )

        return ($lines -join "`n")
    }
}

function New-TopDeskGuestActivity {
    <#
    .SYNOPSIS
        Erstellt eine oder mehrere operative Aktivitaeten in TOPdesk
        basierend auf Audit-CSV-Daten.
    .PARAMETER CsvPath
        Pfad zur Audit-CSV-Datei. Wenn nicht angegeben, wird die neueste CSV verwendet.
    .PARAMETER Filter
        Filter fuer die Konten: Alle, Inaktiv, Abgelaufen, Kritisch, Freemailer, Deaktiviert
    .PARAMETER Mode
        'Sammel' = eine Aktivitaet fuer alle Konten (Standard)
        'Einzel' = je eine Aktivitaet pro Konto
    .PARAMETER BriefDescription
        Kurzbeschreibung fuer die Aktivitaet. Wird automatisch generiert wenn nicht angegeben.
    .PARAMETER Category
        Kategorie (Name oder ID). Ueberschreibt den Standardwert aus der Konfiguration.
    .PARAMETER Subcategory
        Unterkategorie (Name oder ID). Ueberschreibt den Standardwert.
    .PARAMETER OperatorGroup
        Operatorgruppe (Name oder ID). Ueberschreibt den Standardwert.
    .PARAMETER Operator
        Operator (Name oder ID). Ueberschreibt den Standardwert.
    .PARAMETER PlannedStartDate
        Geplanter Start. Standard: Jetzt.
    .PARAMETER PlannedEndDate
        Geplantes Ende. Standard: Jetzt + 7 Tage.
    .PARAMETER WhatIf
        Zeigt an, was passieren wuerde, ohne Aenderungen vorzunehmen.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string]$CsvPath,

        [Parameter()]
        [ValidateSet("Alle", "Inaktiv", "Abgelaufen", "Kritisch", "Freemailer", "Deaktiviert")]
        [string]$Filter = "Alle",

        [Parameter()]
        [ValidateSet("Sammel", "Einzel")]
        [string]$Mode = "Sammel",

        [Parameter()]
        [string]$BriefDescription,

        [Parameter()]
        [string]$Category,

        [Parameter()]
        [string]$Subcategory,

        [Parameter()]
        [string]$OperatorGroup,

        [Parameter()]
        [string]$Operator,

        [Parameter()]
        [datetime]$PlannedStartDate = (Get-Date),

        [Parameter()]
        [datetime]$PlannedEndDate = ((Get-Date).AddDays(7))
    )

    # CSV importieren
    $import = Import-GuestAuditCsv -CsvPath $CsvPath -Filter $Filter
    if (-not $import -or $import.Filtered -eq 0) {
        Write-Host "Keine Konten zum Verarbeiten. Abbruch." -ForegroundColor Yellow
        return
    }

    $accounts = $import.Data
    $auditDate = if ($import.CsvPath) {
        (Get-Item $import.CsvPath).LastWriteTime.ToString("dd.MM.yyyy HH:mm")
    }
    else { Get-Date -Format "dd.MM.yyyy HH:mm" }

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host " TOPdesk Aktivitaet erstellen" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Modus      : $Mode"
    Write-Host "  Filter     : $Filter"
    Write-Host "  Konten     : $($import.Filtered)"
    Write-Host "  Audit-CSV  : $($import.CsvPath)"
    Write-Host "  Zeitraum   : $($PlannedStartDate.ToString('dd.MM.yyyy')) - $($PlannedEndDate.ToString('dd.MM.yyyy'))"
    Write-Host "  WhatIf     : $($WhatIfPreference -or $PSBoundParameters.ContainsKey('WhatIf'))"
    Write-Host "========================================`n" -ForegroundColor Cyan

    # Standard-Werte zusammenbauen
    $effectiveCategory    = if ($Category)      { $Category }      else { $Script:TopDeskConfig.Defaults.Category }
    $effectiveSubcategory = if ($Subcategory)    { $Subcategory }   else { $Script:TopDeskConfig.Defaults.Subcategory }
    $effectiveGroup       = if ($OperatorGroup)  { $OperatorGroup } else { $Script:TopDeskConfig.Defaults.OperatorGroup }
    $effectiveOperator    = if ($Operator)       { $Operator }      else { $Script:TopDeskConfig.Defaults.Operator }
    $effectiveType        = $Script:TopDeskConfig.Defaults.ActivityType

    # Aktivitaeten erstellen
    $createdActivities = @()
    $errorCount = 0

    if ($Mode -eq "Sammel") {
        # Eine Sammel-Aktivitaet fuer alle Konten
        $brief = if ($BriefDescription) {
            $BriefDescription
        }
        else {
            $filterLabel = if ($Filter -ne "Alle") { " ($Filter)" } else { "" }
            "M365 Guest Governance: $($import.Filtered) Konten pruefen$filterLabel"
        }

        # Max. 80 Zeichen fuer briefDescription
        if ($brief.Length -gt 80) {
            $brief = $brief.Substring(0, 77) + "..."
        }

        $description = New-GuestActivityDescription -Accounts $accounts -Mode "Sammel" -AuditDate $auditDate

        $body = Build-ActivityRequestBody `
            -BriefDescription $brief `
            -Description $description `
            -Category $effectiveCategory `
            -Subcategory $effectiveSubcategory `
            -OperatorGroup $effectiveGroup `
            -Operator $effectiveOperator `
            -ActivityType $effectiveType `
            -PlannedStartDate $PlannedStartDate `
            -PlannedEndDate $PlannedEndDate

        if ($WhatIfPreference) {
            Write-Host "[WhatIf] WUERDE Sammel-Aktivitaet erstellen:" -ForegroundColor Magenta
            Write-Host "  Titel     : $brief" -ForegroundColor Magenta
            Write-Host "  Konten    : $($import.Filtered)" -ForegroundColor Magenta
            Write-Host "  Kategorie : $effectiveCategory" -ForegroundColor Magenta
            Write-Host "  Gruppe    : $effectiveGroup" -ForegroundColor Magenta
            Write-Host "  Zeitraum  : $($PlannedStartDate.ToString('dd.MM.yyyy')) - $($PlannedEndDate.ToString('dd.MM.yyyy'))" -ForegroundColor Magenta
            Write-Host "`n[WhatIf] Keine Aenderungen vorgenommen.`n" -ForegroundColor Magenta
        }
        else {
            try {
                Write-Host "Erstelle Sammel-Aktivitaet..." -ForegroundColor Yellow
                $result = Invoke-TopDeskApi -Endpoint $Script:TopDeskConfig.ActivityEndpoint -Method POST -Body $body
                $activityNumber = if ($result.number) { $result.number } else { $result.id }
                Write-Host "[OK] Aktivitaet erstellt: $activityNumber" -ForegroundColor Green
                Write-Host "  Titel: $brief" -ForegroundColor Gray
                $createdActivities += $result
            }
            catch {
                Write-Warning "Fehler beim Erstellen: $_"
                $errorCount++
            }
        }
    }
    else {
        # Einzel-Modus: je eine Aktivitaet pro Konto
        $counter = 0
        foreach ($acc in $accounts) {
            $counter++
            Write-Progress -Activity "Erstelle Aktivitaeten" -Status "$counter von $($import.Filtered)" `
                -PercentComplete (($counter / $import.Filtered) * 100)

            $brief = if ($BriefDescription) {
                "$BriefDescription - $($acc.DisplayName)"
            }
            else {
                $fm = if ($acc.IsFreemailer -eq "True") { " [DLP]" } else { "" }
                "Guest Governance: $($acc.DisplayName)$fm ($($acc.Severity))"
            }

            if ($brief.Length -gt 80) {
                $brief = $brief.Substring(0, 77) + "..."
            }

            $description = New-GuestActivityDescription -Accounts @($acc) -Mode "Einzel" -AuditDate $auditDate

            $body = Build-ActivityRequestBody `
                -BriefDescription $brief `
                -Description $description `
                -Category $effectiveCategory `
                -Subcategory $effectiveSubcategory `
                -OperatorGroup $effectiveGroup `
                -Operator $effectiveOperator `
                -ActivityType $effectiveType `
                -PlannedStartDate $PlannedStartDate `
                -PlannedEndDate $PlannedEndDate

            if ($WhatIfPreference) {
                Write-Host "[WhatIf] WUERDE erstellen: $brief" -ForegroundColor Magenta
            }
            else {
                try {
                    $result = Invoke-TopDeskApi -Endpoint $Script:TopDeskConfig.ActivityEndpoint -Method POST -Body $body
                    $activityNumber = if ($result.number) { $result.number } else { $result.id }
                    Write-Host "  [OK] $activityNumber - $($acc.DisplayName)" -ForegroundColor Green
                    $createdActivities += $result
                }
                catch {
                    Write-Warning "  [FEHLER] $($acc.DisplayName): $_"
                    $errorCount++
                }
            }
        }
        Write-Progress -Activity "Erstelle Aktivitaeten" -Completed
    }

    # Ergebnis
    if (-not $WhatIfPreference) {
        Write-Host "`n--- Ergebnis ---" -ForegroundColor Cyan
        Write-Host "  Erstellt : $($createdActivities.Count)" -ForegroundColor Green
        Write-Host "  Fehler   : $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })

        # Ergebnis-Log speichern
        if ($createdActivities.Count -gt 0) {
            $logDir = Join-Path $PSScriptRoot "Reports"
            if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
            $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
            $logPath = Join-Path $logDir "TopDesk_Activities_$timestamp.json"

            $logData = [PSCustomObject]@{
                CreatedAt    = (Get-Date -Format "o")
                TopDeskUrl   = $Script:TopDeskConfig.BaseUrl
                Mode         = $Mode
                Filter       = $Filter
                SourceCsv    = $import.CsvPath
                AccountCount = $import.Filtered
                Activities   = $createdActivities
            }
            $logData | ConvertTo-Json -Depth 5 | Out-File -FilePath $logPath -Encoding UTF8
            Write-Host "  Log      : $logPath" -ForegroundColor Gray
        }
        Write-Host ""
    }

    return $createdActivities
}

function Build-ActivityRequestBody {
    <#
    .SYNOPSIS
        Baut den JSON-Body fuer die TOPdesk API zusammen.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BriefDescription,

        [Parameter()]
        [string]$Description,

        [Parameter()]
        [string]$Category,

        [Parameter()]
        [string]$Subcategory,

        [Parameter()]
        [string]$OperatorGroup,

        [Parameter()]
        [string]$Operator,

        [Parameter()]
        [string]$ActivityType,

        [Parameter()]
        [datetime]$PlannedStartDate,

        [Parameter()]
        [datetime]$PlannedEndDate
    )

    $body = @{
        briefDescription = $BriefDescription
    }

    if ($Description) {
        $body.request = $Description
    }

    # Referenz-Felder: ID (GUID) oder Name verwenden
    if ($Category) {
        if ($Category -match '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$') {
            $body.category = @{ id = $Category }
        }
        else {
            $body.category = @{ name = $Category }
        }
    }

    if ($Subcategory) {
        if ($Subcategory -match '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$') {
            $body.subcategory = @{ id = $Subcategory }
        }
        else {
            $body.subcategory = @{ name = $Subcategory }
        }
    }

    if ($OperatorGroup) {
        if ($OperatorGroup -match '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$') {
            $body.operatorGroup = @{ id = $OperatorGroup }
        }
        else {
            $body.operatorGroup = @{ groupName = $OperatorGroup }
        }
    }

    if ($Operator) {
        if ($Operator -match '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$') {
            $body.operator = @{ id = $Operator }
        }
        else {
            $body.operator = @{ name = $Operator }
        }
    }

    if ($ActivityType) {
        if ($ActivityType -match '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$') {
            $body.activityType = @{ id = $ActivityType }
        }
        else {
            $body.activityType = @{ name = $ActivityType }
        }
    }

    if ($PlannedStartDate) {
        $body.plannedStartDate = $PlannedStartDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    }

    if ($PlannedEndDate) {
        $body.plannedEndDate = $PlannedEndDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    }

    return $body
}

# ============================================================================
# Interaktives Menue
# ============================================================================

function Show-TopDeskMenu {
    <#
    .SYNOPSIS
        Zeigt ein interaktives Menue fuer die TOPdesk-Integration.
    #>
    [CmdletBinding()]
    param()

    # Konfiguration laden
    if (-not (Initialize-TopDeskConfig)) {
        Write-Host "[FEHLER] Konfiguration konnte nicht geladen werden." -ForegroundColor Red
        return
    }

    # Verbindung herstellen
    if (-not (Connect-TopDesk)) {
        return
    }

    while ($true) {
        # Audit-CSV Status
        $reportDir = Join-Path $PSScriptRoot "Reports"
        $latestCsv = $null
        $csvInfo = $null
        if (Test-Path $reportDir) {
            $csvFiles = Get-ChildItem -Path $reportDir -Filter "GuestAudit_*.csv" -File |
                Sort-Object LastWriteTime -Descending
            if ($csvFiles.Count -gt 0) {
                $latestCsv = $csvFiles[0].FullName
                $csvInfo = $csvFiles[0]
            }
        }
        $csvStatus = if ($csvInfo) {
            "$($csvInfo.Name) ($($csvInfo.LastWriteTime.ToString('dd.MM.yyyy HH:mm')))"
        }
        else {
            "Keine CSV vorhanden - Bitte zuerst M365-GuestGovernance.ps1 ausfuehren!"
        }

        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host " TOPdesk Guest Governance Integration" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "  TOPdesk: $($Script:TopDeskConfig.BaseUrl)" -ForegroundColor Gray
        Write-Host "  CSV:     $csvStatus" -ForegroundColor $(if ($latestCsv) { 'Gray' } else { 'DarkYellow' })
        Write-Host ""
        Write-Host "  [1] Sammel-Aktivitaet erstellen (alle markierten Konten)" -ForegroundColor White
        Write-Host "      Eine Aktivitaet mit allen betroffenen Konten" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  [2] Einzel-Aktivitaeten erstellen (je Konto)" -ForegroundColor White
        Write-Host "      Separate Aktivitaet pro Gastkonto" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  [3] Nur Freemailer melden (DLP-Risiko)" -ForegroundColor Yellow
        Write-Host "      Sammel-Aktivitaet nur fuer Freemailer-Konten" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  [4] Nur kritische Konten melden" -ForegroundColor Red
        Write-Host "      Sammel-Aktivitaet fuer Konten mit Schweregrad 'Hoch'" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  [5] API-Einstellungen anzeigen" -ForegroundColor DarkGray
        Write-Host "      Verfuegbare Kategorien, Status, etc. abrufen" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  [6] Konfiguration aendern" -ForegroundColor DarkGray
        Write-Host "      TOPdesk URL, Kategorien, Operator, etc." -ForegroundColor DarkGray
        Write-Host ""

        # Credential-Status anzeigen
        $credPath = Get-TopDeskCredentialPath
        $credStatus = if (Test-Path $credPath) { "gespeichert" } else { "nicht gespeichert" }
        Write-Host "  [7] Anmeldedaten verwalten" -ForegroundColor DarkGray
        Write-Host "      Credentials sicher speichern/loeschen (Status: $credStatus)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  [Q] Beenden" -ForegroundColor DarkGray
        Write-Host "========================================" -ForegroundColor Cyan

        $choice = Read-Host "`nAuswahl"

        switch ($choice) {
            # ===== 1. Sammel-Aktivitaet =====
            "1" {
                if (-not $latestCsv) {
                    Write-Host "`n[FEHLER] Keine Audit-CSV vorhanden!" -ForegroundColor Red
                    continue
                }

                $filterChoice = Read-Host "Filter (Alle/Inaktiv/Abgelaufen/Kritisch/Freemailer/Deaktiviert) [Alle]"
                if (-not $filterChoice) { $filterChoice = "Alle" }

                # WhatIf zuerst
                New-TopDeskGuestActivity -CsvPath $latestCsv -Filter $filterChoice -Mode "Sammel" -WhatIf

                $execute = Read-Host "`nAktivitaet erstellen? (Ja/Nein)"
                if ($execute -in @("Ja", "J", "ja", "j")) {
                    New-TopDeskGuestActivity -CsvPath $latestCsv -Filter $filterChoice -Mode "Sammel"
                }
                else {
                    Write-Host "Abgebrochen.`n" -ForegroundColor Yellow
                }
            }

            # ===== 2. Einzel-Aktivitaeten =====
            "2" {
                if (-not $latestCsv) {
                    Write-Host "`n[FEHLER] Keine Audit-CSV vorhanden!" -ForegroundColor Red
                    continue
                }

                $filterChoice = Read-Host "Filter (Alle/Inaktiv/Abgelaufen/Kritisch/Freemailer/Deaktiviert) [Alle]"
                if (-not $filterChoice) { $filterChoice = "Alle" }

                # WhatIf zuerst
                New-TopDeskGuestActivity -CsvPath $latestCsv -Filter $filterChoice -Mode "Einzel" -WhatIf

                $execute = Read-Host "`nAktivitaeten erstellen? (Ja/Nein)"
                if ($execute -in @("Ja", "J", "ja", "j")) {
                    New-TopDeskGuestActivity -CsvPath $latestCsv -Filter $filterChoice -Mode "Einzel"
                }
                else {
                    Write-Host "Abgebrochen.`n" -ForegroundColor Yellow
                }
            }

            # ===== 3. Nur Freemailer =====
            "3" {
                if (-not $latestCsv) {
                    Write-Host "`n[FEHLER] Keine Audit-CSV vorhanden!" -ForegroundColor Red
                    continue
                }

                New-TopDeskGuestActivity -CsvPath $latestCsv -Filter "Freemailer" -Mode "Sammel" `
                    -BriefDescription "M365 DLP: Freemailer-Gastkonten erkannt" -WhatIf

                $execute = Read-Host "`nAktivitaet erstellen? (Ja/Nein)"
                if ($execute -in @("Ja", "J", "ja", "j")) {
                    New-TopDeskGuestActivity -CsvPath $latestCsv -Filter "Freemailer" -Mode "Sammel" `
                        -BriefDescription "M365 DLP: Freemailer-Gastkonten erkannt"
                }
                else {
                    Write-Host "Abgebrochen.`n" -ForegroundColor Yellow
                }
            }

            # ===== 4. Nur Kritische =====
            "4" {
                if (-not $latestCsv) {
                    Write-Host "`n[FEHLER] Keine Audit-CSV vorhanden!" -ForegroundColor Red
                    continue
                }

                New-TopDeskGuestActivity -CsvPath $latestCsv -Filter "Kritisch" -Mode "Sammel" `
                    -BriefDescription "M365 Guest Governance: Kritische Gastkonten" -WhatIf

                $execute = Read-Host "`nAktivitaet erstellen? (Ja/Nein)"
                if ($execute -in @("Ja", "J", "ja", "j")) {
                    New-TopDeskGuestActivity -CsvPath $latestCsv -Filter "Kritisch" -Mode "Sammel" `
                        -BriefDescription "M365 Guest Governance: Kritische Gastkonten"
                }
                else {
                    Write-Host "Abgebrochen.`n" -ForegroundColor Yellow
                }
            }

            # ===== 5. API-Einstellungen =====
            "5" {
                Write-Host "`n--- TOPdesk API-Einstellungen ---" -ForegroundColor Cyan
                $settings = Get-TopDeskOperationalSettings
                if ($settings) {
                    Write-Host "`nVerfuegbare Einstellungen:" -ForegroundColor Yellow
                    $settings | ConvertTo-Json -Depth 5 | Write-Host -ForegroundColor Gray
                }
                else {
                    Write-Host "Einstellungen konnten nicht abgerufen werden." -ForegroundColor Yellow
                    Write-Host "Hinweis: GET /operationalActivities/settings erfordert Operations-Management-Lizenz." -ForegroundColor DarkGray
                }
            }

            # ===== 6. Konfiguration aendern =====
            "6" {
                Initialize-TopDeskConfig -Force
            }

            # ===== 7. Anmeldedaten verwalten =====
            "7" {
                $credPath = Get-TopDeskCredentialPath
                $credExists = Test-Path $credPath

                Write-Host "`n--- Anmeldedaten verwalten ---" -ForegroundColor Cyan

                if ($credExists) {
                    $credInfo = Get-Item $credPath
                    Write-Host "  Status   : Gespeichert" -ForegroundColor Green
                    Write-Host "  Datei    : $credPath" -ForegroundColor Gray
                    Write-Host "  Erstellt : $($credInfo.LastWriteTime.ToString('dd.MM.yyyy HH:mm'))" -ForegroundColor Gray
                    Write-Host "  Schutz   : DPAPI (Windows-Benutzer + Maschine)" -ForegroundColor Gray
                    Write-Host ""
                    Write-Host "  [1] Neu speichern (ueberschreiben)" -ForegroundColor White
                    Write-Host "  [2] Loeschen" -ForegroundColor Red
                    Write-Host "  [3] Zurueck" -ForegroundColor DarkGray

                    $credChoice = Read-Host "`nAuswahl"
                    switch ($credChoice) {
                        "1" {
                            $newCred = Get-Credential -Message "Neue TOPdesk API Anmeldedaten"
                            if ($newCred) {
                                Save-TopDeskCredential -Credential $newCred
                                # Verbindung mit neuen Credentials testen
                                $Script:TopDeskConfig.Credential = $null
                                Connect-TopDesk -Credential $newCred
                            }
                        }
                        "2" {
                            $confirmDel = Read-Host "Gespeicherte Credentials wirklich loeschen? (Ja/Nein)"
                            if ($confirmDel -in @("Ja", "J", "ja", "j")) {
                                Remove-TopDeskCredential
                            }
                        }
                    }
                }
                else {
                    Write-Host "  Status: Nicht gespeichert" -ForegroundColor Yellow
                    Write-Host "  Aktuell werden Credentials bei jedem Start abgefragt." -ForegroundColor Gray
                    Write-Host ""
                    $saveNow = Read-Host "Aktuelle Credentials jetzt sicher speichern? (Ja/Nein)"
                    if ($saveNow -in @("Ja", "J", "ja", "j")) {
                        if ($Script:TopDeskConfig.Credential) {
                            Save-TopDeskCredential -Credential $Script:TopDeskConfig.Credential
                        }
                        else {
                            $newCred = Get-Credential -Message "TOPdesk API Anmeldedaten speichern"
                            if ($newCred) {
                                Save-TopDeskCredential -Credential $newCred
                            }
                        }
                    }
                }
            }

            # ===== Beenden =====
            { $_ -in @("Q", "q") } {
                Write-Host "`nAuf Wiedersehen!`n" -ForegroundColor Cyan
                return
            }
            default {
                Write-Warning "Ungueltige Auswahl."
            }
        }
    }
}

# ============================================================================
# Auto-Start
# ============================================================================

# Wenn direkt ausgefuehrt: interaktives Menue starten
if ($MyInvocation.InvocationName -ne '.') {
    Show-TopDeskMenu
}
