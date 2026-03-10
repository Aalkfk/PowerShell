<#
.SYNOPSIS
    M365 Guest Account Governance Tool
.DESCRIPTION
    PowerShell-basiertes Governance-Tool fuer Microsoft 365 Gastkonten.
    - Audit: Gastkonten auslesen, Inaktivitaet und Alter pruefen, Berichte generieren (HTML + CSV/JSON)
    - Cleanup: Gastkonten deaktivieren/loeschen mit WhatIf-Unterstuetzung und CSV-Massenimport
    - Domain-Whitelist: Bestimmte Domains vom Cleanup ausschliessen
.NOTES
    Erfordert: Microsoft.Graph PowerShell SDK
    Berechtigungen: User.Read.All, AuditLog.Read.All, Directory.ReadWrite.All (fuer Cleanup)
    Verwendet Beta-API fuer zuverlaessige SignInActivity-Daten und Sponsor-Abfragen.
#>

#Requires -Version 5.1

[CmdletBinding()]
param()

# ============================================================================
# Konfiguration
# ============================================================================
$Script:Config = @{
    InactiveDaysThreshold = 60
    MaxAgeDaysThreshold   = 365
    ReportOutputDir       = (Join-Path $PSScriptRoot "Reports")
    DefaultScopes         = @(
        "User.Read.All",
        "AuditLog.Read.All",
        "Directory.ReadWrite.All"
    )
    # Domains die beim Cleanup ignoriert werden (z.B. Partner, Dienstleister)
    # Kann auch ueber Parameter oder Datei geladen werden
    ExcludedDomains       = @()
    # Bekannte Freemailer-Domains (DLP-Risiko: private E-Mail-Adressen)
    FreemailerDomains     = @(
        # Google
        "gmail.com", "googlemail.com",
        # Microsoft
        "outlook.com", "outlook.de", "hotmail.com", "hotmail.de", "live.com", "live.de", "msn.com",
        # Deutsche Freemailer
        "gmx.de", "gmx.net", "gmx.at", "gmx.ch",
        "web.de",
        "t-online.de",
        "freenet.de",
        "arcor.de",
        "online.de",
        "email.de",
        "mail.de",
        "posteo.de",
        "mailbox.org",
        # Yahoo
        "yahoo.com", "yahoo.de", "ymail.com",
        # Apple
        "icloud.com", "me.com", "mac.com",
        # Weitere internationale
        "aol.com", "aol.de",
        "protonmail.com", "proton.me", "pm.me",
        "zoho.com",
        "tutanota.com", "tuta.io",
        "gmx.com",
        "mail.com"
    )
    # Beta-API Properties fuer zuverlaessige SignInActivity
    GraphUserProperties   = @(
        "id",
        "displayName",
        "mail",
        "userPrincipalName",
        "createdDateTime",
        "accountEnabled",
        "signInActivity",
        "externalUserState",
        "externalUserStateChangeDateTime"
    )
}

# ============================================================================
# Hilfsfunktionen
# ============================================================================

function Test-Prerequisites {
    <#
    .SYNOPSIS
        Prueft ob die benoetigten PowerShell-Module installiert sind.
    #>
    [CmdletBinding()]
    param()

    $requiredModules = @("Microsoft.Graph.Authentication")
    $missing = @()

    foreach ($mod in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $mod)) {
            $missing += $mod
        }
    }

    if ($missing.Count -gt 0) {
        Write-Warning "Fehlende Module: $($missing -join ', ')"
        Write-Host "`nInstallation mit:" -ForegroundColor Yellow
        Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Cyan
        return $false
    }

    Write-Host "[OK] Alle benoetigten Module vorhanden." -ForegroundColor Green
    return $true
}

function Connect-M365Governance {
    <#
    .SYNOPSIS
        Stellt eine Verbindung zu Microsoft Graph her.
    .PARAMETER TenantId
        Optional: Tenant-ID fuer die Verbindung.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$TenantId
    )

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if ($context) {
        Write-Host "[OK] Bereits verbunden als: $($context.Account)" -ForegroundColor Green
        return $true
    }

    $connectParams = @{
        Scopes = $Script:Config.DefaultScopes
    }
    if ($TenantId) {
        $connectParams.TenantId = $TenantId
    }

    try {
        Connect-MgGraph @connectParams
        $context = Get-MgContext
        Write-Host "[OK] Verbunden mit Tenant: $($context.TenantId)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Verbindung fehlgeschlagen: $_"
        return $false
    }
}

function Test-FreemailerDomain {
    <#
    .SYNOPSIS
        Prueft ob eine Domain ein bekannter Freemailer ist (DLP-Risiko).
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Domain
    )

    if (-not $Domain) { return $false }
    return ($Script:Config.FreemailerDomains -contains $Domain.ToLower())
}

function Get-AllGraphPages {
    <#
    .SYNOPSIS
        Ruft alle Seiten einer paginierten Graph-API-Antwort ab.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter()]
        [hashtable]$Headers = @{}
    )

    $allResults = @()
    $currentUri = $Uri

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $currentUri -Headers $Headers
        if ($response.value) {
            $allResults += $response.value
        }
        $currentUri = $response.'@odata.nextLink'
    } while ($currentUri)

    return $allResults
}

function Get-DomainFromAddress {
    <#
    .SYNOPSIS
        Extrahiert die Domain aus einer E-Mail-Adresse oder einem UPN.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Address
    )

    if (-not $Address) { return $null }
    if ($Address -match '@(.+)$') {
        return $Matches[1].ToLower()
    }
    # UPN Format mit #EXT# (z.B. user_extern.de#EXT#@tenant.onmicrosoft.com)
    if ($Address -match '_([^#]+)#EXT#') {
        return $Matches[1].ToLower()
    }
    return $null
}

function Get-GuestInviter {
    <#
    .SYNOPSIS
        Ermittelt den Einlader/Sponsor eines Gastkontos ueber die Beta-API.
        Prueft zuerst die Sponsor-Beziehung, dann Audit-Logs als Fallback.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter()]
        [hashtable]$AuditCache = @{}
    )

    # 1. Versuch: Sponsor-Beziehung abfragen (Beta-API)
    try {
        $sponsors = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/beta/users/$UserId/sponsors" `
            -ErrorAction Stop

        if ($sponsors.value -and $sponsors.value.Count -gt 0) {
            $sponsor = $sponsors.value[0]
            return [PSCustomObject]@{
                InviterId   = $sponsor.id
                InviterName = $sponsor.displayName
                InviterMail = $sponsor.mail
                Source       = "Sponsor"
            }
        }
    }
    catch {
        # Sponsor-Endpunkt nicht verfuegbar oder keine Berechtigung - weiter mit Fallback
    }

    # 2. Versuch: Aus vorgeladenem Audit-Cache
    if ($AuditCache.ContainsKey($UserId)) {
        return $AuditCache[$UserId]
    }

    return [PSCustomObject]@{
        InviterId   = $null
        InviterName = "Unbekannt"
        InviterMail = $null
        Source       = "Nicht ermittelbar"
    }
}

function Get-InvitationAuditCache {
    <#
    .SYNOPSIS
        Laedt Audit-Log-Eintraege fuer Gasteinladungen und baut einen Lookup-Cache.
        Durchsucht mehrere Aktivitaetstypen fuer maximale Abdeckung.
        Hinweis: Audit-Logs sind je nach Lizenz auf 30 Tage (Free/P1) begrenzt.
        Mit Entra ID P2 sind bis zu 30 Tage verfuegbar, mit Log Analytics-Export mehr.
    #>
    [CmdletBinding()]
    param()

    $cache = @{}

    # Mehrere Aktivitaetstypen abfragen fuer bessere Abdeckung
    $activityTypes = @(
        "Invite external user",
        "Add user",
        "Redeem external user invite"
    )

    foreach ($activityType in $activityTypes) {
        try {
            Write-Host "  Suche Audit-Logs: '$activityType'..." -ForegroundColor Gray -NoNewline

            $auditUri = "https://graph.microsoft.com/beta/auditLogs/directoryAudits?" +
                "`$filter=activityDisplayName eq '$activityType'&" +
                "`$select=targetResources,initiatedBy,activityDateTime&" +
                "`$top=999"

            $auditLogs = Get-AllGraphPages -Uri $auditUri
            $matchCount = 0

            foreach ($log in $auditLogs) {
                # Ziel-User-ID aus targetResources extrahieren
                $targetId = $null
                if ($log.targetResources) {
                    foreach ($target in $log.targetResources) {
                        if ($target.type -eq "User" -and $target.id) {
                            $targetId = $target.id
                            break
                        }
                    }
                }

                # Nur hinzufuegen wenn noch nicht im Cache (erste Quelle gewinnt)
                if ($targetId -and -not $cache.ContainsKey($targetId)) {
                    $initiator = $log.initiatedBy
                    $inviterName = $null
                    $inviterMail = $null
                    $inviterId = $null

                    if ($initiator.user) {
                        $inviterName = $initiator.user.displayName
                        $inviterMail = $initiator.user.userPrincipalName
                        $inviterId = $initiator.user.id
                    }
                    elseif ($initiator.app) {
                        $inviterName = "App: $($initiator.app.displayName)"
                    }

                    # Nur cachen wenn wir einen Namen haben
                    if ($inviterName) {
                        $cache[$targetId] = [PSCustomObject]@{
                            InviterId   = $inviterId
                            InviterName = $inviterName
                            InviterMail = $inviterMail
                            Source       = "Audit-Log ($activityType)"
                        }
                        $matchCount++
                    }
                }
            }

            Write-Host " $matchCount neue Zuordnungen" -ForegroundColor Gray
        }
        catch {
            Write-Host " Fehler" -ForegroundColor DarkYellow
            Write-Warning "    Audit-Logs '$activityType' konnten nicht geladen werden: $($_.Exception.Message)"
        }
    }

    $totalCached = $cache.Count
    Write-Host "  [OK] Gesamt: $totalCached Einlader-Zuordnungen aus Audit-Logs." -ForegroundColor Gray
    if ($totalCached -eq 0) {
        Write-Host "  [Hinweis] Audit-Logs sind lizenzabhaengig auf 30 Tage begrenzt." -ForegroundColor DarkYellow
        Write-Host "            Fuer aeltere Gaeste wird der Einlader ueber die Sponsor-Beziehung ermittelt." -ForegroundColor DarkYellow
    }

    return $cache
}

function Get-GuestActivityStatus {
    <#
    .SYNOPSIS
        Ermittelt den Aktivitaetsstatus eines Gastkontos.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$User,

        [Parameter()]
        [object]$InviterInfo,

        [Parameter()]
        [datetime]$ReferenceDate = (Get-Date)
    )

    $createdDate = if ($User.createdDateTime) { [datetime]$User.createdDateTime } else { $null }
    $lastSignIn  = $null
    $lastNonInteractive = $null

    # SignInActivity aus Beta-API (Hashtable, nicht Objekt)
    $signInActivity = $User.signInActivity
    if ($signInActivity) {
        if ($signInActivity.lastSignInDateTime) {
            $lastSignIn = [datetime]$signInActivity.lastSignInDateTime
        }
        if ($signInActivity.lastNonInteractiveSignInDateTime) {
            $lastNonInteractive = [datetime]$signInActivity.lastNonInteractiveSignInDateTime
        }
    }

    # Letzte Aktivitaet = neuester Wert aus interaktivem und nicht-interaktivem Sign-In
    $lastActivity = $null
    if ($lastSignIn -and $lastNonInteractive) {
        $lastActivity = @($lastSignIn, $lastNonInteractive) | Sort-Object -Descending | Select-Object -First 1
    }
    elseif ($lastSignIn) {
        $lastActivity = $lastSignIn
    }
    elseif ($lastNonInteractive) {
        $lastActivity = $lastNonInteractive
    }

    $daysSinceActivity = if ($lastActivity) { ($ReferenceDate - $lastActivity).Days } else { -1 }
    $daysSinceCreation = if ($createdDate) { ($ReferenceDate - $createdDate).Days } else { -1 }

    $flags = @()
    $isInactive = $false
    $isExpired  = $false

    # Inaktivitaetspruefung
    if ($daysSinceActivity -eq -1) {
        $flags += "KEINE_ANMELDUNG"
        $isInactive = $true
    }
    elseif ($daysSinceActivity -gt $Script:Config.InactiveDaysThreshold) {
        $flags += "INAKTIV_${daysSinceActivity}_TAGE"
        $isInactive = $true
    }

    # Alterspruefung
    if ($daysSinceCreation -gt $Script:Config.MaxAgeDaysThreshold) {
        $flags += "ABGELAUFEN_${daysSinceCreation}_TAGE"
        $isExpired = $true
    }

    # Einladungsstatus mappen
    $externalState = $User.externalUserState
    $invitationStatus = switch ($externalState) {
        "Accepted"          { "Akzeptiert" }
        "PendingAcceptance" { "Ausstehend" }
        default             { if ($externalState) { $externalState } else { "Unbekannt" } }
    }

    # Einladungsstatus-Flags
    if ($externalState -eq "PendingAcceptance") {
        $flags += "EINLADUNG_AUSSTEHEND"
    }

    # Einlader-Informationen
    $inviterName = "Unbekannt"
    $inviterMail = $null
    $inviterSource = "Nicht ermittelbar"
    if ($InviterInfo) {
        $inviterName = $InviterInfo.InviterName
        $inviterMail = $InviterInfo.InviterMail
        $inviterSource = $InviterInfo.Source
    }

    # Domain extrahieren fuer Whitelist-Pruefung
    $guestDomain = Get-DomainFromAddress -Address $User.mail
    if (-not $guestDomain) {
        $guestDomain = Get-DomainFromAddress -Address $User.userPrincipalName
    }

    # Freemailer-Pruefung (DLP-Risiko)
    $isFreemailer = Test-FreemailerDomain -Domain $guestDomain
    if ($isFreemailer) {
        $flags += "FREEMAILER"
    }

    return [PSCustomObject]@{
        UserId                = $User.id
        DisplayName           = $User.displayName
        Mail                  = $User.mail
        UserPrincipalName     = $User.userPrincipalName
        GuestDomain           = $guestDomain
        CreatedDateTime       = $createdDate
        LastSignIn            = $lastSignIn
        LastNonInteractive    = $lastNonInteractive
        LastActivity          = $lastActivity
        DaysSinceActivity     = $daysSinceActivity
        DaysSinceCreation     = $daysSinceCreation
        AccountEnabled        = $User.accountEnabled
        InvitationStatus      = $invitationStatus
        ExternalUserState     = $externalState
        InviterName           = $inviterName
        InviterMail           = $inviterMail
        InviterSource         = $inviterSource
        IsFreemailer          = $isFreemailer
        IsInactive            = $isInactive
        IsExpired             = $isExpired
        Flags                 = ($flags -join "; ")
        Severity              = if ($isInactive -and $isExpired) { "Hoch" } elseif ($isExpired) { "Mittel" } elseif ($isInactive) { "Niedrig" } else { "OK" }
    }
}

# ============================================================================
# Funktion 1: Audit & Report
# ============================================================================

function Get-M365GuestReport {
    <#
    .SYNOPSIS
        Liest alle Gastkonten aus, prueft Aktivitaet und Alter, erstellt Berichte.
    .DESCRIPTION
        - Liest alle Gastkonten ueber die Microsoft Graph Beta-API
        - Prueft letzte Anmeldung (inaktiv > 60 Tage) inkl. nicht-interaktiver Sign-Ins
        - Prueft Erstellungsdatum (aelter als 365 Tage)
        - Ermittelt Einladungsstatus und Einlader/Sponsor
        - Generiert HTML-Report (grafisch) und CSV + JSON (technisch)
    .PARAMETER OutputDirectory
        Zielverzeichnis fuer die Reports. Standard: ./Reports
    .PARAMETER InactiveDays
        Schwellenwert fuer Inaktivitaet in Tagen. Standard: 60
    .PARAMETER MaxAgeDays
        Maximales Kontoalter in Tagen. Standard: 365
    .PARAMETER SkipInviterLookup
        Ueberspringt die Einlader-Ermittlung (beschleunigt den Audit).
    .PARAMETER PassThru
        Gibt die Ergebnisse zusaetzlich als Objekte zurueck.
    .EXAMPLE
        Get-M365GuestReport
    .EXAMPLE
        Get-M365GuestReport -InactiveDays 90 -MaxAgeDays 180 -PassThru
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$OutputDirectory = $Script:Config.ReportOutputDir,

        [Parameter()]
        [int]$InactiveDays = $Script:Config.InactiveDaysThreshold,

        [Parameter()]
        [int]$MaxAgeDays = $Script:Config.MaxAgeDaysThreshold,

        [Parameter()]
        [switch]$SkipInviterLookup,

        [Parameter()]
        [switch]$PassThru
    )

    # Schwellenwerte aktualisieren
    $Script:Config.InactiveDaysThreshold = $InactiveDays
    $Script:Config.MaxAgeDaysThreshold = $MaxAgeDays

    # Output-Verzeichnis erstellen
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
    }

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host " M365 Gastkonto-Audit" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " Inaktivitaet-Schwelle : $InactiveDays Tage"
    Write-Host " Max. Kontoalter       : $MaxAgeDays Tage"
    Write-Host " Einlader-Abfrage      : $(if ($SkipInviterLookup) {'Uebersprungen'} else {'Aktiv'})"
    Write-Host " Ausgabeverzeichnis    : $OutputDirectory"
    Write-Host " API-Endpunkt          : Beta (zuverlaessige SignInActivity)"
    Write-Host "========================================`n" -ForegroundColor Cyan

    # -----------------------------------------------------------------------
    # Gastkonten ueber Beta-API abrufen (zuverlaessige SignInActivity-Daten)
    # -----------------------------------------------------------------------
    Write-Host "Lade Gastkonten ueber Graph Beta-API..." -ForegroundColor Yellow

    $selectProperties = $Script:Config.GraphUserProperties -join ","
    $graphUri = "https://graph.microsoft.com/beta/users?" +
        "`$filter=userType eq 'Guest'&" +
        "`$select=$selectProperties&" +
        "`$count=true&" +
        "`$top=999"

    try {
        $guests = Get-AllGraphPages -Uri $graphUri -Headers @{ "ConsistencyLevel" = "eventual" }
    }
    catch {
        Write-Error "Fehler beim Abrufen der Gastkonten: $_"
        Write-Host "`nMoegliche Ursachen:" -ForegroundColor Yellow
        Write-Host "  - Fehlende Berechtigung: AuditLog.Read.All (fuer SignInActivity)" -ForegroundColor Yellow
        Write-Host "  - Fehlende Berechtigung: User.Read.All" -ForegroundColor Yellow
        Write-Host "  - Verbindung nicht hergestellt (Connect-MgGraph)" -ForegroundColor Yellow
        return
    }

    $totalGuests = ($guests | Measure-Object).Count
    Write-Host "[OK] $totalGuests Gastkonten gefunden." -ForegroundColor Green

    if ($totalGuests -eq 0) {
        Write-Host "Keine Gastkonten vorhanden. Audit beendet." -ForegroundColor Yellow
        return
    }

    # Debug: Pruefen ob SignInActivity geladen wurde
    $withSignIn = ($guests | Where-Object { $_.signInActivity -ne $null } | Measure-Object).Count
    Write-Host "  -> $withSignIn von $totalGuests mit SignInActivity-Daten" -ForegroundColor Gray

    # -----------------------------------------------------------------------
    # Einlader/Sponsor ermitteln
    # -----------------------------------------------------------------------
    $auditCache = @{}
    if (-not $SkipInviterLookup) {
        Write-Host "`nErmittle Einlader/Sponsoren..." -ForegroundColor Yellow
        $auditCache = Get-InvitationAuditCache
    }

    # -----------------------------------------------------------------------
    # Aktivitaetsstatus ermitteln
    # -----------------------------------------------------------------------
    Write-Host "`nAnalysiere Aktivitaetsstatus..." -ForegroundColor Yellow
    $results = @()
    $counter = 0
    foreach ($guest in $guests) {
        $counter++
        Write-Progress -Activity "Analysiere Gastkonten" -Status "$counter von $totalGuests" -PercentComplete (($counter / $totalGuests) * 100)

        # Einlader ermitteln
        $inviterInfo = $null
        if (-not $SkipInviterLookup) {
            $inviterInfo = Get-GuestInviter -UserId $guest.id -AuditCache $auditCache
        }

        $results += Get-GuestActivityStatus -User $guest -InviterInfo $inviterInfo
    }
    Write-Progress -Activity "Analysiere Gastkonten" -Completed

    # Statistiken
    $pendingCount   = ($results | Where-Object { $_.InvitationStatus -eq "Ausstehend" }).Count
    $freemailerCount = ($results | Where-Object { $_.IsFreemailer }).Count
    $stats = [PSCustomObject]@{
        Gesamt           = $totalGuests
        Aktiv            = ($results | Where-Object { -not $_.IsInactive -and -not $_.IsExpired }).Count
        Inaktiv          = ($results | Where-Object { $_.IsInactive }).Count
        Abgelaufen       = ($results | Where-Object { $_.IsExpired }).Count
        Deaktiviert      = ($results | Where-Object { -not $_.AccountEnabled }).Count
        KriteriumHoch    = ($results | Where-Object { $_.Severity -eq "Hoch" }).Count
        Ausstehend       = $pendingCount
        Freemailer       = $freemailerCount
    }

    # Einlader-Statistik
    $inviterKnown   = ($results | Where-Object { $_.InviterName -ne "Unbekannt" }).Count
    $inviterUnknown = ($results | Where-Object { $_.InviterName -eq "Unbekannt" }).Count

    # Freemailer-Domain-Aufschluesselung
    $freemailerBreakdown = $results | Where-Object { $_.IsFreemailer } |
        Group-Object -Property GuestDomain |
        Sort-Object Count -Descending

    Write-Host "`n--- Zusammenfassung ---" -ForegroundColor Cyan
    Write-Host "  Gesamt              : $($stats.Gesamt)"
    Write-Host "  Aktiv               : $($stats.Aktiv)" -ForegroundColor Green
    Write-Host "  Inaktiv (>$InactiveDays d)    : $($stats.Inaktiv)" -ForegroundColor Yellow
    Write-Host "  Abgelaufen (>$MaxAgeDays d)  : $($stats.Abgelaufen)" -ForegroundColor Red
    Write-Host "  Deaktiviert         : $($stats.Deaktiviert)" -ForegroundColor DarkGray
    Write-Host "  Kritisch (beides)   : $($stats.KriteriumHoch)" -ForegroundColor Magenta
    Write-Host "  Einladung ausstehend: $($stats.Ausstehend)" -ForegroundColor DarkYellow
    Write-Host "  Einlader bekannt    : $inviterKnown / $($stats.Gesamt)" -ForegroundColor $(if ($inviterUnknown -eq 0) { "Green" } else { "Yellow" })
    if ($freemailerCount -gt 0) {
        Write-Host "  Freemailer (DLP)    : $freemailerCount" -ForegroundColor Red
        foreach ($fm in $freemailerBreakdown) {
            Write-Host "    -> $($fm.Name): $($fm.Count)" -ForegroundColor DarkRed
        }
    }
    else {
        Write-Host "  Freemailer (DLP)    : 0" -ForegroundColor Green
    }
    Write-Host ""

    # Timestamp fuer Dateinamen
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"

    # --- CSV Export ---
    $csvPath = Join-Path $OutputDirectory "GuestAudit_$timestamp.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Host "[Export] CSV : $csvPath" -ForegroundColor Green

    # --- JSON Export ---
    $jsonPath = Join-Path $OutputDirectory "GuestAudit_$timestamp.json"
    $jsonOutput = [PSCustomObject]@{
        Meta = [PSCustomObject]@{
            GeneratedAt       = (Get-Date -Format "o")
            TenantId          = (Get-MgContext).TenantId
            InactiveDays      = $InactiveDays
            MaxAgeDays        = $MaxAgeDays
            TotalGuests       = $totalGuests
            ApiEndpoint       = "Beta"
        }
        Statistics = $stats
        Guests     = $results
    }
    $jsonOutput | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonPath -Encoding UTF8
    Write-Host "[Export] JSON: $jsonPath" -ForegroundColor Green

    # --- HTML Report ---
    $htmlPath = Join-Path $OutputDirectory "GuestAudit_$timestamp.html"
    $htmlContent = New-GuestHtmlReport -Results $results -Stats $stats -InactiveDays $InactiveDays -MaxAgeDays $MaxAgeDays
    $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Host "[Export] HTML: $htmlPath" -ForegroundColor Green

    Write-Host "`n[OK] Audit abgeschlossen. Alle Reports gespeichert.`n" -ForegroundColor Green

    if ($PassThru) {
        return $results
    }
}

function New-GuestHtmlReport {
    <#
    .SYNOPSIS
        Generiert einen grafischen HTML-Report.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Results,

        [Parameter(Mandatory)]
        [object]$Stats,

        [Parameter()]
        [int]$InactiveDays = 60,

        [Parameter()]
        [int]$MaxAgeDays = 365
    )

    $tenantId = (Get-MgContext).TenantId
    $reportDate = Get-Date -Format "dd.MM.yyyy HH:mm"

    $flaggedResults   = $Results | Where-Object { $_.IsInactive -or $_.IsExpired } | Sort-Object Severity -Descending
    $okResults        = $Results | Where-Object { -not $_.IsInactive -and -not $_.IsExpired }
    $flaggedCount     = ($flaggedResults | Measure-Object).Count
    $okCount          = ($okResults | Measure-Object).Count
    $complianceRate   = if ($Stats.Gesamt -gt 0) { [math]::Round(($okCount / $Stats.Gesamt) * 100, 1) } else { 100 }

    # Einladungsstatus-Badge erzeugen
    function Get-InvitationBadge($status) {
        switch ($status) {
            "Akzeptiert"  { return '<span class="badge badge-accepted">Akzeptiert</span>' }
            "Ausstehend"  { return '<span class="badge badge-pending">Ausstehend</span>' }
            default       { return '<span class="badge badge-unknown">Unbekannt</span>' }
        }
    }

    # Hilfsfunktion: ISO-Sortiertwert fuer Datum
    function Get-SortDate($dt) {
        if ($dt) { return $dt.ToString("yyyy-MM-dd") } else { return "0000-00-00" }
    }

    # Flagged Tabellenzeilen
    $flaggedRows = ""
    foreach ($r in $flaggedResults) {
        $severityClass = switch ($r.Severity) {
            "Hoch"    { "severity-high" }
            "Mittel"  { "severity-medium" }
            "Niedrig" { "severity-low" }
            default   { "" }
        }
        $statusBadge     = if ($r.AccountEnabled) { '<span class="badge badge-active">Aktiv</span>' } else { '<span class="badge badge-disabled">Deaktiviert</span>' }
        $invitationBadge = Get-InvitationBadge $r.InvitationStatus
        $lastAct         = if ($r.LastActivity) { $r.LastActivity.ToString("dd.MM.yyyy") } else { '<span class="no-data">Nie</span>' }
        $lastActSort     = Get-SortDate $r.LastActivity
        $lastSignInFmt   = if ($r.LastSignIn) { $r.LastSignIn.ToString("dd.MM.yyyy") } else { "-" }
        $lastNonIntFmt   = if ($r.LastNonInteractive) { $r.LastNonInteractive.ToString("dd.MM.yyyy") } else { "-" }
        $created         = if ($r.CreatedDateTime) { $r.CreatedDateTime.ToString("dd.MM.yyyy") } else { "Unbekannt" }
        $createdSort     = Get-SortDate $r.CreatedDateTime
        $inviterDisplay  = if ($r.InviterMail) { "<span title=`"$($r.InviterMail) (Quelle: $($r.InviterSource))`">$($r.InviterName)</span>" } else { "<span title=`"Quelle: $($r.InviterSource)`">$($r.InviterName)</span>" }
        $severitySort    = switch ($r.Severity) { "Hoch" { "1" } "Mittel" { "2" } "Niedrig" { "3" } default { "4" } }
        $domainBadge     = if ($r.IsFreemailer) { "<span class=`"badge badge-freemailer`" title=`"DLP-Risiko: Private E-Mail-Adresse`">$($r.GuestDomain)</span>" } else { $r.GuestDomain }

        # Data-Attribute fuer korrekte Sortierung und Filterung
        $flaggedRows += @"
        <tr class="$severityClass$(if ($r.IsFreemailer) {' freemailer-row'})" data-severity="$($r.Severity)" data-inactive="$($r.IsInactive)" data-expired="$($r.IsExpired)" data-disabled="$(-not $r.AccountEnabled)" data-invitation="$($r.InvitationStatus)" data-freemailer="$($r.IsFreemailer)">
            <td title="$($r.UserId)">$($r.DisplayName)</td>
            <td>$($r.Mail)</td>
            <td>$domainBadge</td>
            <td data-sort="$createdSort">$created</td>
            <td data-sort="$lastActSort" title="Interaktiv: $lastSignInFmt | Nicht-interaktiv: $lastNonIntFmt">$lastAct</td>
            <td data-sort="$($r.DaysSinceActivity)">$($r.DaysSinceActivity)</td>
            <td data-sort="$($r.DaysSinceCreation)">$($r.DaysSinceCreation)</td>
            <td>$statusBadge</td>
            <td>$invitationBadge</td>
            <td>$inviterDisplay</td>
            <td data-sort="$severitySort"><span class="severity-badge $severityClass">$($r.Severity)</span></td>
            <td class="flags">$($r.Flags)</td>
        </tr>
"@
    }

    # OK Tabellenzeilen
    $okRows = ""
    foreach ($r in $okResults) {
        $statusBadge     = if ($r.AccountEnabled) { '<span class="badge badge-active">Aktiv</span>' } else { '<span class="badge badge-disabled">Deaktiviert</span>' }
        $invitationBadge = Get-InvitationBadge $r.InvitationStatus
        $lastAct         = if ($r.LastActivity) { $r.LastActivity.ToString("dd.MM.yyyy") } else { '<span class="no-data">Nie</span>' }
        $lastActSort     = Get-SortDate $r.LastActivity
        $lastSignInFmt   = if ($r.LastSignIn) { $r.LastSignIn.ToString("dd.MM.yyyy") } else { "-" }
        $lastNonIntFmt   = if ($r.LastNonInteractive) { $r.LastNonInteractive.ToString("dd.MM.yyyy") } else { "-" }
        $created         = if ($r.CreatedDateTime) { $r.CreatedDateTime.ToString("dd.MM.yyyy") } else { "Unbekannt" }
        $createdSort     = Get-SortDate $r.CreatedDateTime
        $inviterDisplay  = if ($r.InviterMail) { "<span title=`"$($r.InviterMail) (Quelle: $($r.InviterSource))`">$($r.InviterName)</span>" } else { "<span title=`"Quelle: $($r.InviterSource)`">$($r.InviterName)</span>" }
        $domainBadge     = if ($r.IsFreemailer) { "<span class=`"badge badge-freemailer`" title=`"DLP-Risiko: Private E-Mail-Adresse`">$($r.GuestDomain)</span>" } else { $r.GuestDomain }

        $okRows += @"
        <tr$(if ($r.IsFreemailer) {' class="freemailer-row"'}) data-severity="OK" data-inactive="False" data-expired="False" data-disabled="$(-not $r.AccountEnabled)" data-invitation="$($r.InvitationStatus)" data-freemailer="$($r.IsFreemailer)">
            <td title="$($r.UserId)">$($r.DisplayName)</td>
            <td>$($r.Mail)</td>
            <td>$domainBadge</td>
            <td data-sort="$createdSort">$created</td>
            <td data-sort="$lastActSort" title="Interaktiv: $lastSignInFmt | Nicht-interaktiv: $lastNonIntFmt">$lastAct</td>
            <td data-sort="$($r.DaysSinceActivity)">$($r.DaysSinceActivity)</td>
            <td data-sort="$($r.DaysSinceCreation)">$($r.DaysSinceCreation)</td>
            <td>$statusBadge</td>
            <td>$invitationBadge</td>
            <td>$inviterDisplay</td>
            <td><span class="severity-badge severity-ok">OK</span></td>
        </tr>
"@
    }

    return @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>M365 Gastkonto-Audit Report</title>
    <style>
        :root {
            --primary: #0078d4;
            --danger: #d13438;
            --warning: #ffaa44;
            --success: #107c10;
            --muted: #605e5c;
            --bg: #faf9f8;
            --card-bg: #ffffff;
            --border: #edebe9;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', -apple-system, sans-serif;
            background: var(--bg);
            color: #323130;
            line-height: 1.5;
        }
        .container { max-width: 1600px; margin: 0 auto; padding: 24px; }

        /* Header */
        .header {
            background: linear-gradient(135deg, #0078d4 0%, #005a9e 100%);
            color: white;
            padding: 32px;
            border-radius: 8px;
            margin-bottom: 24px;
        }
        .header h1 { font-size: 28px; font-weight: 600; margin-bottom: 8px; }
        .header-meta { display: flex; gap: 24px; font-size: 14px; opacity: 0.9; flex-wrap: wrap; }

        /* Stat Cards */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }
        .stat-card {
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 20px;
            text-align: center;
            transition: box-shadow 0.2s, transform 0.15s;
            cursor: pointer;
            user-select: none;
        }
        .stat-card:hover { box-shadow: 0 2px 8px rgba(0,0,0,0.1); transform: translateY(-2px); }
        .stat-card.active-filter { border: 2px solid var(--primary); box-shadow: 0 2px 12px rgba(0,120,212,0.3); }
        .stat-card .stat-value { font-size: 36px; font-weight: 700; }
        .stat-card .stat-label { font-size: 13px; color: var(--muted); margin-top: 4px; }
        .stat-card.total .stat-value { color: var(--primary); }
        .stat-card.ok .stat-value { color: var(--success); }
        .stat-card.warn .stat-value { color: var(--warning); }
        .stat-card.danger .stat-value { color: var(--danger); }
        .stat-card.critical .stat-value { color: #881798; }
        .stat-card.pending .stat-value { color: #ca5010; }

        /* Active filter indicator */
        .filter-active-hint {
            display: none;
            text-align: center;
            padding: 8px 16px;
            margin-bottom: 16px;
            background: #deecf9;
            border: 1px solid var(--primary);
            border-radius: 6px;
            font-size: 13px;
            color: #0078d4;
        }
        .filter-active-hint.visible { display: block; }
        .filter-active-hint button {
            background: none;
            border: 1px solid var(--primary);
            color: var(--primary);
            padding: 2px 10px;
            border-radius: 4px;
            cursor: pointer;
            margin-left: 8px;
            font-size: 12px;
        }
        .filter-active-hint button:hover { background: var(--primary); color: white; }

        /* Compliance Bar */
        .compliance-section {
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 24px;
        }
        .compliance-section h3 { margin-bottom: 12px; font-size: 16px; }
        .progress-bar {
            background: #e1dfdd;
            border-radius: 4px;
            height: 24px;
            overflow: hidden;
            position: relative;
        }
        .progress-fill {
            height: 100%;
            border-radius: 4px;
            transition: width 0.5s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: 600;
            font-size: 13px;
        }
        .compliance-good { background: var(--success); }
        .compliance-warn { background: var(--warning); }
        .compliance-bad { background: var(--danger); }

        /* Policy Info */
        .policy-info {
            background: #f3f2f1;
            border-left: 4px solid var(--primary);
            padding: 16px 20px;
            border-radius: 0 8px 8px 0;
            margin-bottom: 24px;
            font-size: 14px;
        }
        .policy-info strong { color: var(--primary); }

        /* Tables */
        .table-section {
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: 8px;
            overflow: hidden;
            margin-bottom: 24px;
        }
        .table-header {
            padding: 16px 20px;
            border-bottom: 1px solid var(--border);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .table-header h2 { font-size: 18px; }
        .table-header .count {
            background: var(--primary);
            color: white;
            padding: 2px 10px;
            border-radius: 12px;
            font-size: 13px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        thead th {
            background: #f3f2f1;
            padding: 10px 12px;
            text-align: left;
            font-weight: 600;
            font-size: 12px;
            color: var(--muted);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            border-bottom: 2px solid var(--border);
            position: sticky;
            top: 0;
            cursor: pointer;
            white-space: nowrap;
        }
        thead th:hover { background: #e1dfdd; }
        thead th .sort-arrow { font-size: 10px; margin-left: 4px; opacity: 0.4; }
        thead th.sorted-asc .sort-arrow,
        thead th.sorted-desc .sort-arrow { opacity: 1; color: var(--primary); }
        tbody td {
            padding: 8px 12px;
            border-bottom: 1px solid var(--border);
        }
        tbody tr:hover { background: #f3f2f1; }

        /* Severity Rows */
        .severity-high { border-left: 4px solid var(--danger); }
        .severity-medium { border-left: 4px solid var(--warning); }
        .severity-low { border-left: 4px solid #0078d4; }

        /* Badges */
        .badge {
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 600;
            white-space: nowrap;
        }
        .badge-active { background: #dff6dd; color: #107c10; }
        .badge-disabled { background: #f4f4f4; color: #605e5c; }
        .badge-accepted { background: #dff6dd; color: #107c10; }
        .badge-pending { background: #fff4ce; color: #8a6d3b; }
        .badge-unknown { background: #f4f4f4; color: #605e5c; }
        .badge-freemailer { background: #fde7e9; color: #a80000; border: 1px solid #f1bbbc; }
        .freemailer-row { background: #fef6f6; }
        .freemailer-row:hover { background: #fde7e9 !important; }
        .stat-card.freemailer .stat-value { color: #a80000; }
        .severity-badge {
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 600;
            color: white;
            white-space: nowrap;
        }
        .severity-badge.severity-high { background: var(--danger); border: none; }
        .severity-badge.severity-medium { background: var(--warning); color: #323130; border: none; }
        .severity-badge.severity-low { background: var(--primary); border: none; }
        .severity-badge.severity-ok { background: var(--success); border: none; }

        .flags { font-size: 11px; color: var(--muted); max-width: 220px; }
        .no-data { color: var(--danger); font-style: italic; }

        /* Collapsible */
        .collapsible { cursor: pointer; user-select: none; }
        .collapsible-content { display: none; }
        .collapsible-content.open { display: block; }
        .toggle-icon { transition: transform 0.2s; display: inline-block; margin-right: 8px; }
        .toggle-icon.open { transform: rotate(90deg); }

        /* Filter */
        .filter-bar {
            padding: 12px 20px;
            border-bottom: 1px solid var(--border);
            display: flex;
            gap: 12px;
            align-items: center;
            flex-wrap: wrap;
        }
        .filter-bar input {
            padding: 6px 12px;
            border: 1px solid var(--border);
            border-radius: 4px;
            font-size: 14px;
            min-width: 250px;
        }
        .filter-bar input:focus { outline: none; border-color: var(--primary); }

        /* Tooltip on hover for activity breakdown */
        td[title] { cursor: help; }

        /* Footer */
        .footer {
            text-align: center;
            color: var(--muted);
            font-size: 13px;
            padding: 24px;
        }

        /* Print */
        @media print {
            body { background: white; }
            .container { max-width: 100%; padding: 0; }
            .header { background: #0078d4 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
            .filter-bar { display: none; }
            .filter-active-hint { display: none !important; }
            .table-section { break-inside: avoid; }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Header -->
        <div class="header">
            <h1>M365 Gastkonto-Audit Report</h1>
            <div class="header-meta">
                <span>Erstellt: $reportDate</span>
                <span>Tenant: $tenantId</span>
                <span>Inaktivitaet: &gt;$InactiveDays Tage</span>
                <span>Max. Alter: &gt;$MaxAgeDays Tage</span>
                <span>API: Beta-Endpunkt</span>
            </div>
        </div>

        <!-- Policy Info -->
        <div class="policy-info">
            <strong>Richtlinie:</strong> Gastkonten werden als <strong>inaktiv</strong> markiert, wenn keine Anmeldung innerhalb von <strong>$InactiveDays Tagen</strong> erfolgt ist.
            Konten, die aelter als <strong>$MaxAgeDays Tage</strong> sind, gelten als <strong>abgelaufen</strong> und muessen erneuert oder entfernt werden.
            <br><em>Letzte Aktivitaet: Hover ueber das Datum zeigt die Aufschluesselung (interaktiv / nicht-interaktiv).</em>
            <br><em>Stat-Karten anklicken um die Tabelle zu filtern. Erneut klicken um den Filter aufzuheben.</em>
        </div>

        <!-- Stats -->
        <div class="stats-grid">
            <div class="stat-card total" onclick="filterByCard(this, 'all')">
                <div class="stat-value">$($Stats.Gesamt)</div>
                <div class="stat-label">Gastkonten Gesamt</div>
            </div>
            <div class="stat-card ok" onclick="filterByCard(this, 'ok')">
                <div class="stat-value">$okCount</div>
                <div class="stat-label">Konform</div>
            </div>
            <div class="stat-card warn" onclick="filterByCard(this, 'inactive')">
                <div class="stat-value">$($Stats.Inaktiv)</div>
                <div class="stat-label">Inaktiv (&gt;$InactiveDays d)</div>
            </div>
            <div class="stat-card danger" onclick="filterByCard(this, 'expired')">
                <div class="stat-value">$($Stats.Abgelaufen)</div>
                <div class="stat-label">Abgelaufen (&gt;$MaxAgeDays d)</div>
            </div>
            <div class="stat-card critical" onclick="filterByCard(this, 'critical')">
                <div class="stat-value">$($Stats.KriteriumHoch)</div>
                <div class="stat-label">Kritisch (beides)</div>
            </div>
            <div class="stat-card pending" onclick="filterByCard(this, 'pending')">
                <div class="stat-value">$($Stats.Ausstehend)</div>
                <div class="stat-label">Einladung ausstehend</div>
            </div>
            <div class="stat-card freemailer" onclick="filterByCard(this, 'freemailer')">
                <div class="stat-value">$($Stats.Freemailer)</div>
                <div class="stat-label">Freemailer (DLP)</div>
            </div>
            <div class="stat-card" onclick="filterByCard(this, 'disabled')" style="cursor:pointer;">
                <div class="stat-value" style="color: var(--muted);">$($Stats.Deaktiviert)</div>
                <div class="stat-label">Bereits deaktiviert</div>
            </div>
        </div>

        <!-- Filter active hint -->
        <div class="filter-active-hint" id="filterHint">
            <span id="filterHintText"></span>
            <button onclick="clearCardFilter()">Filter aufheben</button>
        </div>

        <!-- Compliance -->
        <div class="compliance-section">
            <h3>Compliance-Rate</h3>
            <div class="progress-bar">
                <div class="progress-fill $(if ($complianceRate -ge 80) {'compliance-good'} elseif ($complianceRate -ge 50) {'compliance-warn'} else {'compliance-bad'})"
                     style="width: ${complianceRate}%;">
                    ${complianceRate}%
                </div>
            </div>
        </div>

        <!-- Flagged Accounts -->
        <div class="table-section">
            <div class="table-header">
                <h2>Markierte Konten</h2>
                <span class="count" id="flaggedCount">$flaggedCount</span>
            </div>
            <div class="filter-bar">
                <input type="text" id="filterFlagged" placeholder="Filtern nach Name, E-Mail, Einlader, Flags..." onkeyup="filterTable('flaggedTable', this.value)">
            </div>
            <div style="overflow-x: auto;">
                <table id="flaggedTable">
                    <thead>
                        <tr>
                            <th onclick="sortTable('flaggedTable', 0)">Name <span class="sort-arrow">&#9650;&#9660;</span></th>
                            <th onclick="sortTable('flaggedTable', 1)">E-Mail <span class="sort-arrow">&#9650;&#9660;</span></th>
                            <th onclick="sortTable('flaggedTable', 2)">Domain <span class="sort-arrow">&#9650;&#9660;</span></th>
                            <th onclick="sortTable('flaggedTable', 3)">Erstellt <span class="sort-arrow">&#9650;&#9660;</span></th>
                            <th onclick="sortTable('flaggedTable', 4)">Letzte Aktivitaet <span class="sort-arrow">&#9650;&#9660;</span></th>
                            <th onclick="sortTable('flaggedTable', 5)">Tage inaktiv <span class="sort-arrow">&#9650;&#9660;</span></th>
                            <th onclick="sortTable('flaggedTable', 6)">Alter (Tage) <span class="sort-arrow">&#9650;&#9660;</span></th>
                            <th onclick="sortTable('flaggedTable', 7)">Status <span class="sort-arrow">&#9650;&#9660;</span></th>
                            <th onclick="sortTable('flaggedTable', 8)">Einladung <span class="sort-arrow">&#9650;&#9660;</span></th>
                            <th onclick="sortTable('flaggedTable', 9)">Eingeladen von <span class="sort-arrow">&#9650;&#9660;</span></th>
                            <th onclick="sortTable('flaggedTable', 10)">Schwere <span class="sort-arrow">&#9650;&#9660;</span></th>
                            <th>Flags</th>
                        </tr>
                    </thead>
                    <tbody>
                        $flaggedRows
                    </tbody>
                </table>
            </div>
        </div>

        <!-- OK Accounts (collapsible) -->
        <div class="table-section">
            <div class="table-header collapsible" onclick="toggleSection('okSection', this)">
                <h2><span class="toggle-icon">&#9654;</span>Konforme Konten</h2>
                <span class="count" id="okCount" style="background: var(--success);">$okCount</span>
            </div>
            <div id="okSection" class="collapsible-content">
                <div class="filter-bar">
                    <input type="text" id="filterOk" placeholder="Filtern nach Name, E-Mail, Einlader..." onkeyup="filterTable('okTable', this.value)">
                </div>
                <div style="overflow-x: auto;">
                    <table id="okTable">
                        <thead>
                            <tr>
                                <th onclick="sortTable('okTable', 0)">Name <span class="sort-arrow">&#9650;&#9660;</span></th>
                                <th onclick="sortTable('okTable', 1)">E-Mail <span class="sort-arrow">&#9650;&#9660;</span></th>
                                <th onclick="sortTable('okTable', 2)">Domain <span class="sort-arrow">&#9650;&#9660;</span></th>
                                <th onclick="sortTable('okTable', 3)">Erstellt <span class="sort-arrow">&#9650;&#9660;</span></th>
                                <th onclick="sortTable('okTable', 4)">Letzte Aktivitaet <span class="sort-arrow">&#9650;&#9660;</span></th>
                                <th onclick="sortTable('okTable', 5)">Tage inaktiv <span class="sort-arrow">&#9650;&#9660;</span></th>
                                <th onclick="sortTable('okTable', 6)">Alter (Tage) <span class="sort-arrow">&#9650;&#9660;</span></th>
                                <th onclick="sortTable('okTable', 7)">Status <span class="sort-arrow">&#9650;&#9660;</span></th>
                                <th onclick="sortTable('okTable', 8)">Einladung <span class="sort-arrow">&#9650;&#9660;</span></th>
                                <th onclick="sortTable('okTable', 9)">Eingeladen von <span class="sort-arrow">&#9650;&#9660;</span></th>
                                <th>Bewertung</th>
                            </tr>
                        </thead>
                        <tbody>
                            $okRows
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Footer -->
        <div class="footer">
            M365 Guest Governance Tool &bull; Report generiert am $reportDate &bull; API: Beta-Endpunkt
        </div>
    </div>

    <script>
        // =============================================
        // Collapsible Sections
        // =============================================
        function toggleSection(id, el) {
            var section = document.getElementById(id);
            var icon = el.querySelector('.toggle-icon');
            section.classList.toggle('open');
            icon.classList.toggle('open');
        }

        // =============================================
        // Text-Filter (Suchfeld)
        // =============================================
        function filterTable(tableId, filter) {
            var table = document.getElementById(tableId);
            var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
            var lowerFilter = filter.toLowerCase();
            for (var i = 0; i < rows.length; i++) {
                var text = rows[i].textContent.toLowerCase();
                // Nur anzeigen wenn Text-Filter UND Card-Filter passen
                var cardHidden = rows[i].getAttribute('data-card-hidden') === 'true';
                rows[i].style.display = (text.indexOf(lowerFilter) > -1 && !cardHidden) ? '' : 'none';
            }
        }

        // =============================================
        // Sortierung mit data-sort Attributen
        // =============================================
        function sortTable(tableId, colIndex) {
            var table = document.getElementById(tableId);
            var tbody = table.getElementsByTagName('tbody')[0];
            var rows = Array.from(tbody.getElementsByTagName('tr'));
            var th = table.getElementsByTagName('thead')[0].getElementsByTagName('th')[colIndex];

            // Sortierrichtung bestimmen
            var asc = !th.classList.contains('sorted-asc');

            // Alle th-Klassen zuruecksetzen
            var allTh = table.getElementsByTagName('thead')[0].getElementsByTagName('th');
            for (var i = 0; i < allTh.length; i++) {
                allTh[i].classList.remove('sorted-asc', 'sorted-desc');
            }
            th.classList.add(asc ? 'sorted-asc' : 'sorted-desc');

            rows.sort(function(a, b) {
                var aCell = a.cells[colIndex];
                var bCell = b.cells[colIndex];

                // data-sort Attribut hat Prioritaet (fuer Datums- und numerische Spalten)
                var aVal = aCell.getAttribute('data-sort');
                var bVal = bCell.getAttribute('data-sort');

                // Fallback auf Textinhalt
                if (aVal === null) aVal = aCell.textContent.trim();
                if (bVal === null) bVal = bCell.textContent.trim();

                // Numerisch sortieren wenn moeglich
                var aNum = parseFloat(aVal);
                var bNum = parseFloat(bVal);
                if (!isNaN(aNum) && !isNaN(bNum)) {
                    return asc ? aNum - bNum : bNum - aNum;
                }

                // String-Sortierung
                return asc ? aVal.localeCompare(bVal, 'de') : bVal.localeCompare(aVal, 'de');
            });

            rows.forEach(function(row) { tbody.appendChild(row); });
        }

        // =============================================
        // Stat-Card Filter
        // =============================================
        var currentCardFilter = null;

        function filterByCard(cardEl, filterType) {
            var cards = document.querySelectorAll('.stat-card');
            var hint = document.getElementById('filterHint');
            var hintText = document.getElementById('filterHintText');

            // Toggle: gleichen Filter nochmal klicken = aufheben
            if (currentCardFilter === filterType) {
                clearCardFilter();
                return;
            }

            currentCardFilter = filterType;

            // Alle Cards deaktivieren, aktive markieren
            cards.forEach(function(c) { c.classList.remove('active-filter'); });
            cardEl.classList.add('active-filter');

            // Filter-Labels
            var filterLabels = {
                'all': 'Alle Gastkonten',
                'ok': 'Nur konforme Konten',
                'inactive': 'Nur inaktive Konten',
                'expired': 'Nur abgelaufene Konten',
                'critical': 'Nur kritische Konten (inaktiv + abgelaufen)',
                'pending': 'Nur Konten mit ausstehender Einladung',
                'freemailer': 'Nur Freemailer-Konten (DLP-Risiko)',
                'disabled': 'Nur deaktivierte Konten'
            };
            hintText.textContent = 'Filter aktiv: ' + (filterLabels[filterType] || filterType);
            hint.classList.add('visible');

            // Beide Tabellen filtern
            applyCardFilter('flaggedTable', filterType);
            applyCardFilter('okTable', filterType);

            // OK-Section oeffnen wenn noetig
            if (filterType === 'ok' || filterType === 'all') {
                var okSection = document.getElementById('okSection');
                if (!okSection.classList.contains('open')) {
                    okSection.classList.add('open');
                    var icon = okSection.previousElementSibling.querySelector('.toggle-icon');
                    if (icon) icon.classList.add('open');
                }
            }
        }

        function applyCardFilter(tableId, filterType) {
            var table = document.getElementById(tableId);
            if (!table) return;
            var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');

            for (var i = 0; i < rows.length; i++) {
                var row = rows[i];
                var show = false;

                var severity   = row.getAttribute('data-severity');
                var inactive   = row.getAttribute('data-inactive') === 'True';
                var expired    = row.getAttribute('data-expired') === 'True';
                var disabled   = row.getAttribute('data-disabled') === 'True';
                var invitation = row.getAttribute('data-invitation');
                var freemailer = row.getAttribute('data-freemailer') === 'True';

                switch (filterType) {
                    case 'all':        show = true; break;
                    case 'ok':         show = !inactive && !expired; break;
                    case 'inactive':   show = inactive; break;
                    case 'expired':    show = expired; break;
                    case 'critical':   show = inactive && expired; break;
                    case 'pending':    show = invitation === 'Ausstehend'; break;
                    case 'freemailer': show = freemailer; break;
                    case 'disabled':   show = disabled; break;
                    default:           show = true;
                }

                row.setAttribute('data-card-hidden', show ? 'false' : 'true');
                row.style.display = show ? '' : 'none';
            }
        }

        function clearCardFilter() {
            currentCardFilter = null;
            var cards = document.querySelectorAll('.stat-card');
            cards.forEach(function(c) { c.classList.remove('active-filter'); });

            var hint = document.getElementById('filterHint');
            hint.classList.remove('visible');

            // Alle Zeilen wieder sichtbar machen
            var tables = ['flaggedTable', 'okTable'];
            tables.forEach(function(tableId) {
                var table = document.getElementById(tableId);
                if (!table) return;
                var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
                for (var i = 0; i < rows.length; i++) {
                    rows[i].setAttribute('data-card-hidden', 'false');
                    rows[i].style.display = '';
                }
            });

            // Text-Filter ggf. neu anwenden
            var filterInput1 = document.getElementById('filterFlagged');
            var filterInput2 = document.getElementById('filterOk');
            if (filterInput1 && filterInput1.value) filterTable('flaggedTable', filterInput1.value);
            if (filterInput2 && filterInput2.value) filterTable('okTable', filterInput2.value);
        }
    </script>
</body>
</html>
"@
}

# ============================================================================
# Funktion 2: Cleanup (Deaktivieren / Loeschen)
# ============================================================================

function Remove-M365GuestAccounts {
    <#
    .SYNOPSIS
        Deaktiviert oder loescht Gastkonten mit WhatIf-Unterstuetzung.
    .DESCRIPTION
        Ermoeglicht das gezielte oder massenhafte Deaktivieren/Loeschen von Gastkonten.
        Unterstuetzt CSV-Import fuer Massenoperationen und WhatIf-Modus.
        Bestimmte Domains koennen ueber -ExcludedDomains oder -ExcludedDomainsFile
        vom Cleanup ausgeschlossen werden.
    .PARAMETER UserIds
        Array von User-IDs zum Verarbeiten.
    .PARAMETER CsvPath
        Pfad zu einer CSV-Datei mit Spalte 'UserId' (oder 'UserPrincipalName').
    .PARAMETER Action
        Auszufuehrende Aktion: 'Disable' (deaktivieren) oder 'Delete' (loeschen).
    .PARAMETER FromAudit
        Verwendet die Ergebnisse eines vorherigen Audits. Erwartet Objekte von Get-M365GuestReport -PassThru.
    .PARAMETER SeverityFilter
        Filtert Audit-Ergebnisse nach Schweregrad: 'Hoch', 'Mittel', 'Niedrig', 'Alle'.
    .PARAMETER ExcludedDomains
        Liste von Domains die vom Cleanup ausgeschlossen werden (z.B. "partner.de","dienstleister.com").
    .PARAMETER ExcludedDomainsFile
        Pfad zu einer Textdatei mit einer Domain pro Zeile die ausgeschlossen werden soll.
    .PARAMETER WhatIf
        Zeigt an, was passieren wuerde, ohne Aenderungen vorzunehmen.
    .PARAMETER Force
        Ueberspringt die Sicherheitsabfrage (nicht empfohlen).
    .EXAMPLE
        Remove-M365GuestAccounts -CsvPath ".\cleanup.csv" -Action Disable -WhatIf
    .EXAMPLE
        Remove-M365GuestAccounts -FromAudit $audit -SeverityFilter "Hoch" -Action Delete -ExcludedDomains "partner.de","trusted.com" -WhatIf
    .EXAMPLE
        Remove-M365GuestAccounts -FromAudit $audit -Action Disable -ExcludedDomainsFile ".\whitelist.txt" -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = "ById")]
    param(
        [Parameter(ParameterSetName = "ById", Mandatory)]
        [string[]]$UserIds,

        [Parameter(ParameterSetName = "ByCsv", Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$CsvPath,

        [Parameter(ParameterSetName = "ByAudit", Mandatory)]
        [object[]]$FromAudit,

        [Parameter(ParameterSetName = "ByAudit")]
        [ValidateSet("Hoch", "Mittel", "Niedrig", "Alle")]
        [string]$SeverityFilter = "Alle",

        [Parameter(Mandatory)]
        [ValidateSet("Disable", "Delete")]
        [string]$Action,

        [Parameter()]
        [string[]]$ExcludedDomains = @(),

        [Parameter()]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$ExcludedDomainsFile,

        [Parameter()]
        [switch]$Force
    )

    # Domain-Whitelist zusammenbauen
    $domainWhitelist = @()
    $domainWhitelist += $ExcludedDomains | ForEach-Object { $_.ToLower().Trim() } | Where-Object { $_ }
    $domainWhitelist += $Script:Config.ExcludedDomains | ForEach-Object { $_.ToLower().Trim() } | Where-Object { $_ }

    if ($ExcludedDomainsFile) {
        $fileEntries = Get-Content -Path $ExcludedDomainsFile |
            ForEach-Object { $_.Trim().ToLower() } |
            Where-Object { $_ -and -not $_.StartsWith('#') }
        $domainWhitelist += $fileEntries
    }
    $domainWhitelist = $domainWhitelist | Select-Object -Unique

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host " M365 Gastkonto-Cleanup" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " Aktion      : $Action"
    Write-Host " WhatIf      : $($WhatIfPreference -or $PSBoundParameters.ContainsKey('WhatIf'))"
    if ($domainWhitelist.Count -gt 0) {
        Write-Host " Geschuetzte Domains ($($domainWhitelist.Count)):" -ForegroundColor Green
        foreach ($d in $domainWhitelist) {
            Write-Host "   - $d" -ForegroundColor Green
        }
    }
    Write-Host "========================================`n" -ForegroundColor Cyan

    # Zielkonten ermitteln
    $targets = @()

    switch ($PSCmdlet.ParameterSetName) {
        "ById" {
            foreach ($id in $UserIds) {
                $targets += [PSCustomObject]@{ UserId = $id; Source = "Parameter"; Mail = $null; GuestDomain = $null }
            }
        }
        "ByCsv" {
            Write-Host "Lade CSV: $CsvPath" -ForegroundColor Yellow
            $csvData = Import-Csv -Path $CsvPath -Delimiter ";"

            foreach ($row in $csvData) {
                $id = if ($row.UserId) { $row.UserId } elseif ($row.UserPrincipalName) { $row.UserPrincipalName } else { $null }
                if ($id) {
                    $mail = $row.Mail
                    $domain = Get-DomainFromAddress -Address $mail
                    if (-not $domain) { $domain = Get-DomainFromAddress -Address $row.UserPrincipalName }

                    $targets += [PSCustomObject]@{
                        UserId      = $id
                        Source      = "CSV"
                        DisplayName = $row.DisplayName
                        Mail        = $mail
                        GuestDomain = $domain
                    }
                }
            }
            Write-Host "[OK] $($targets.Count) Eintraege aus CSV geladen." -ForegroundColor Green
        }
        "ByAudit" {
            $filtered = if ($SeverityFilter -eq "Alle") {
                $FromAudit | Where-Object { $_.IsInactive -or $_.IsExpired }
            }
            else {
                $FromAudit | Where-Object { $_.Severity -eq $SeverityFilter }
            }
            foreach ($item in $filtered) {
                $targets += [PSCustomObject]@{
                    UserId      = $item.UserId
                    Source      = "Audit ($($item.Severity))"
                    DisplayName = $item.DisplayName
                    Mail        = $item.Mail
                    GuestDomain = $item.GuestDomain
                    Flags       = $item.Flags
                }
            }
            Write-Host "[OK] $($targets.Count) Konten aus Audit gefiltert (Schwere: $SeverityFilter)." -ForegroundColor Green
        }
    }

    if ($targets.Count -eq 0) {
        Write-Host "Keine Zielkonten gefunden. Abbruch." -ForegroundColor Yellow
        return
    }

    # Domain-Whitelist anwenden
    if ($domainWhitelist.Count -gt 0) {
        $beforeCount = $targets.Count
        $excluded = @()
        $remaining = @()

        foreach ($t in $targets) {
            $domain = $t.GuestDomain
            if (-not $domain -and $t.Mail) {
                $domain = Get-DomainFromAddress -Address $t.Mail
            }

            if ($domain -and $domainWhitelist -contains $domain) {
                $excluded += $t
            }
            else {
                $remaining += $t
            }
        }

        if ($excluded.Count -gt 0) {
            Write-Host "`n[Whitelist] $($excluded.Count) Konten uebersprungen (geschuetzte Domains):" -ForegroundColor Green
            foreach ($e in $excluded) {
                $name = if ($e.DisplayName) { $e.DisplayName } else { $e.UserId }
                Write-Host "  [GESCHUETZT] $name ($($e.Mail)) -> $($e.GuestDomain)" -ForegroundColor Green
            }
            Write-Host ""
        }

        $targets = $remaining
        Write-Host "[Whitelist] $beforeCount -> $($targets.Count) Konten nach Domain-Filter`n" -ForegroundColor Cyan
    }

    if ($targets.Count -eq 0) {
        Write-Host "Alle Zielkonten wurden durch die Domain-Whitelist ausgeschlossen. Abbruch." -ForegroundColor Yellow
        return
    }

    # Vorschau anzeigen
    Write-Host "Betroffene Konten ($($targets.Count)):" -ForegroundColor Cyan
    Write-Host ("-" * 90)
    $targets | Format-Table -AutoSize -Property @(
        @{N = "UserId"; E = { $_.UserId.Substring(0, [Math]::Min(8, $_.UserId.Length)) + "..." } },
        "DisplayName", "Mail", "GuestDomain", "Source"
    ) | Out-String | Write-Host

    # WhatIf Modus
    if ($WhatIfPreference) {
        Write-Host "[WhatIf] Folgende Aktionen WUERDEN ausgefuehrt werden:" -ForegroundColor Magenta
        Write-Host ("-" * 60) -ForegroundColor Magenta
        foreach ($target in $targets) {
            $name = if ($target.DisplayName) { $target.DisplayName } else { $target.UserId }
            $actionText = if ($Action -eq "Disable") { "DEAKTIVIEREN" } else { "LOESCHEN" }
            Write-Host "  [WhatIf] $actionText : $name ($($target.Mail))" -ForegroundColor Magenta
        }
        Write-Host "`n[WhatIf] Keine Aenderungen vorgenommen. Entfernen Sie -WhatIf um die Aktionen auszufuehren.`n" -ForegroundColor Magenta
        return
    }

    # Sicherheitsabfrage
    if (-not $Force) {
        $actionText = if ($Action -eq "Disable") { "deaktiviert" } else { "GELOESCHT" }
        Write-Host "`n[!] WARNUNG: $($targets.Count) Gastkonto(en) werden $actionText!" -ForegroundColor Red
        $confirm = Read-Host "Fortfahren? (Ja/Nein)"
        if ($confirm -notin @("Ja", "J", "ja", "j", "Yes", "Y", "yes", "y")) {
            Write-Host "Abgebrochen." -ForegroundColor Yellow
            return
        }
    }

    # Aktionen ausfuehren
    $results = @()
    $successCount = 0
    $errorCount = 0

    foreach ($target in $targets) {
        $name = if ($target.DisplayName) { $target.DisplayName } else { $target.UserId }

        try {
            if ($Action -eq "Disable") {
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($target.UserId)" `
                    -Body @{ accountEnabled = $false } | Out-Null
                Write-Host "  [OK] Deaktiviert: $name" -ForegroundColor Green
                $results += [PSCustomObject]@{ UserId = $target.UserId; Name = $name; Domain = $target.GuestDomain; Action = "Disabled"; Status = "Success"; Error = "" }
                $successCount++
            }
            else {
                Invoke-MgGraphRequest -Method DELETE `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($target.UserId)" | Out-Null
                Write-Host "  [OK] Geloescht: $name" -ForegroundColor Green
                $results += [PSCustomObject]@{ UserId = $target.UserId; Name = $name; Domain = $target.GuestDomain; Action = "Deleted"; Status = "Success"; Error = "" }
                $successCount++
            }
        }
        catch {
            Write-Warning "  [FEHLER] $name : $_"
            $results += [PSCustomObject]@{ UserId = $target.UserId; Name = $name; Domain = $target.GuestDomain; Action = $Action; Status = "Error"; Error = $_.Exception.Message }
            $errorCount++
        }
    }

    # Ergebnis-Log speichern
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $logDir = $Script:Config.ReportOutputDir
    if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
    $logPath = Join-Path $logDir "Cleanup_${Action}_$timestamp.csv"
    $results | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    Write-Host "`n--- Ergebnis ---" -ForegroundColor Cyan
    Write-Host "  Erfolgreich : $successCount" -ForegroundColor Green
    Write-Host "  Fehler      : $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
    Write-Host "  Log         : $logPath" -ForegroundColor Gray
    Write-Host ""
}

# ============================================================================
# Interaktive Kontenauswahl
# ============================================================================

function Select-GuestAccounts {
    <#
    .SYNOPSIS
        Zeigt eine interaktive Liste von Gastkonten und erlaubt einzelne Konten abzuwaehlen.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Accounts,

        [Parameter(Mandatory)]
        [string]$ActionLabel
    )

    if ($Accounts.Count -eq 0) {
        Write-Host "Keine Konten zur Auswahl." -ForegroundColor Yellow
        return @()
    }

    Write-Host "`n--- Konten zur $ActionLabel ($($Accounts.Count)) ---" -ForegroundColor Cyan
    Write-Host "Geben Sie die Nummern der Konten ein, die Sie AUSSCHLIESSEN moechten." -ForegroundColor Yellow
    Write-Host "Mehrere Nummern mit Komma trennen (z.B. 1,3,5). Enter = alle behalten.`n" -ForegroundColor Yellow

    # Nummerierte Liste anzeigen
    for ($i = 0; $i -lt $Accounts.Count; $i++) {
        $a = $Accounts[$i]
        $name = if ($a.DisplayName) { $a.DisplayName } else { "N/A" }
        $mail = if ($a.Mail) { $a.Mail } else { $a.UserPrincipalName }
        $domain = if ($a.GuestDomain) { "[$($a.GuestDomain)]" } else { "" }
        $days = if ($a.DaysSinceActivity -ge 0) { "$($a.DaysSinceActivity)d inaktiv" } else { "Nie angemeldet" }
        $severity = if ($a.Severity) { "($($a.Severity))" } else { "" }
        $status = if ($a.AccountEnabled -eq $false) { "[DEAKTIVIERT]" } else { "" }

        Write-Host "  [$($i + 1)] " -ForegroundColor White -NoNewline
        Write-Host "$name" -ForegroundColor Cyan -NoNewline
        Write-Host " - $mail $domain - $days $severity $status" -ForegroundColor Gray
    }

    Write-Host ""
    $excludeInput = Read-Host "Konten ausschliessen (Nummern oder Enter fuer alle)"

    if (-not $excludeInput -or $excludeInput.Trim() -eq "") {
        Write-Host "[OK] Alle $($Accounts.Count) Konten behalten." -ForegroundColor Green
        return $Accounts
    }

    # Nummern parsen
    $excludeIndices = @()
    foreach ($part in ($excludeInput -split ',')) {
        $part = $part.Trim()
        $num = 0
        if ([int]::TryParse($part, [ref]$num)) {
            if ($num -ge 1 -and $num -le $Accounts.Count) {
                $excludeIndices += ($num - 1)
            }
            else {
                Write-Warning "  Nummer $num ist ungueltig (1-$($Accounts.Count))."
            }
        }
    }

    $selected = @()
    for ($i = 0; $i -lt $Accounts.Count; $i++) {
        if ($excludeIndices -contains $i) {
            $name = if ($Accounts[$i].DisplayName) { $Accounts[$i].DisplayName } else { $Accounts[$i].UserId }
            Write-Host "  [AUSGESCHLOSSEN] $name" -ForegroundColor DarkGray
        }
        else {
            $selected += $Accounts[$i]
        }
    }

    Write-Host "`n[OK] $($selected.Count) von $($Accounts.Count) Konten ausgewaehlt." -ForegroundColor Green
    return $selected
}

function Get-DomainWhitelist {
    <#
    .SYNOPSIS
        Baut die Domain-Whitelist aus Config und optionaler Datei zusammen.
    #>
    [CmdletBinding()]
    param()

    $whitelist = @()
    $whitelist += $Script:Config.ExcludedDomains | ForEach-Object { $_.ToLower().Trim() } | Where-Object { $_ }

    # Whitelist-Datei im Skript-Verzeichnis suchen
    $whitelistFile = Join-Path $PSScriptRoot "excluded_domains.txt"
    if (Test-Path $whitelistFile) {
        $fileEntries = Get-Content -Path $whitelistFile |
            ForEach-Object { $_.Trim().ToLower() } |
            Where-Object { $_ -and -not $_.StartsWith('#') }
        $whitelist += $fileEntries
    }

    return ($whitelist | Select-Object -Unique)
}

function Find-LatestAuditCsv {
    <#
    .SYNOPSIS
        Sucht die neueste Audit-CSV-Datei im Report-Verzeichnis.
    #>
    [CmdletBinding()]
    param()

    $reportDir = $Script:Config.ReportOutputDir
    if (-not (Test-Path $reportDir)) {
        return $null
    }

    $csvFiles = Get-ChildItem -Path $reportDir -Filter "GuestAudit_*.csv" -File |
        Sort-Object LastWriteTime -Descending

    if ($csvFiles.Count -gt 0) {
        return $csvFiles[0].FullName
    }
    return $null
}

# ============================================================================
# Hauptmenue (Interaktiver Modus)
# ============================================================================

function Show-GovernanceMenu {
    <#
    .SYNOPSIS
        Zeigt ein interaktives Menue fuer das Governance-Tool.
        Workflow: 1. Audit -> 2. Deaktivieren (Inaktive) -> 3. Loeschen (Deaktivierte)
    #>
    [CmdletBinding()]
    param()

    # Voraussetzungen pruefen
    if (-not (Test-Prerequisites)) {
        return
    }

    # Verbindung herstellen
    if (-not (Connect-M365Governance)) {
        return
    }

    # Domain-Whitelist laden
    $domainWhitelist = Get-DomainWhitelist

    while ($true) {
        # Letzte Audit-CSV pruefen
        $latestCsv = Find-LatestAuditCsv
        $auditStatus = if ($latestCsv) {
            $csvInfo = Get-Item $latestCsv
            "Letzter Audit: $($csvInfo.LastWriteTime.ToString('dd.MM.yyyy HH:mm')) ($($csvInfo.Name))"
        } else {
            "Kein Audit vorhanden - bitte zuerst Schritt 1 ausfuehren!"
        }

        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host " M365 Guest Governance Tool" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  [1] Gastkonto-Audit durchfuehren" -ForegroundColor White
        Write-Host "      Alle Gastkonten auslesen und Reports erzeugen (HTML/CSV/JSON)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  [2] Inaktive Gastkonten deaktivieren" -ForegroundColor Yellow
        Write-Host "      Konten mit >$($Script:Config.InactiveDaysThreshold) Tagen Inaktivitaet (basierend auf Audit-CSV)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  [3] Deaktivierte Gastkonten loeschen" -ForegroundColor Red
        Write-Host "      Bereits deaktivierte Konten endgueltig entfernen (basierend auf Audit-CSV)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  [Q] Beenden" -ForegroundColor DarkGray
        Write-Host ""

        if ($domainWhitelist.Count -gt 0) {
            Write-Host "  Geschuetzte Domains: $($domainWhitelist -join ', ')" -ForegroundColor Green
        }
        Write-Host "  $auditStatus" -ForegroundColor $(if ($latestCsv) { 'Gray' } else { 'DarkYellow' })
        Write-Host "========================================" -ForegroundColor Cyan

        $choice = Read-Host "`nAuswahl"

        switch ($choice) {
            # ==============================================================
            # 1. AUDIT
            # ==============================================================
            "1" {
                Get-M365GuestReport
            }

            # ==============================================================
            # 2. DEAKTIVIEREN (Inaktive Konten >60 Tage)
            # ==============================================================
            "2" {
                # Audit-CSV pruefen
                $csvPath = Find-LatestAuditCsv
                if (-not $csvPath) {
                    Write-Host "`n[FEHLER] Kein Audit gefunden!" -ForegroundColor Red
                    Write-Host "Bitte fuehren Sie zuerst Schritt [1] aus, um einen Audit-Report zu erzeugen.`n" -ForegroundColor Yellow
                    continue
                }

                Write-Host "`n--- Schritt 2: Inaktive Gastkonten deaktivieren ---" -ForegroundColor Yellow
                Write-Host "Lade Audit-Daten: $csvPath" -ForegroundColor Gray

                $auditData = Import-Csv -Path $csvPath -Delimiter ";"

                # Nur inaktive, noch aktive Konten filtern
                $inactiveAccounts = $auditData | Where-Object {
                    $_.IsInactive -eq "True" -and $_.AccountEnabled -eq "True"
                }

                $totalInactive = ($inactiveAccounts | Measure-Object).Count
                Write-Host "[OK] $totalInactive inaktive, noch aktive Gastkonten gefunden.`n" -ForegroundColor Cyan

                if ($totalInactive -eq 0) {
                    Write-Host "Keine inaktiven Konten zum Deaktivieren vorhanden.`n" -ForegroundColor Green
                    continue
                }

                # Domain-Whitelist anzeigen und anwenden
                if ($domainWhitelist.Count -gt 0) {
                    Write-Host "[Whitelist] Folgende Domains werden uebersprungen:" -ForegroundColor Green
                    foreach ($d in $domainWhitelist) {
                        Write-Host "  - $d" -ForegroundColor Green
                    }
                    Write-Host ""

                    $filtered = @()
                    $excluded = @()
                    foreach ($acc in $inactiveAccounts) {
                        $domain = Get-DomainFromAddress -Address $acc.Mail
                        if (-not $domain) { $domain = Get-DomainFromAddress -Address $acc.UserPrincipalName }
                        # Domain auf Objekt setzen fuer Anzeige
                        $acc | Add-Member -NotePropertyName "GuestDomain" -NotePropertyValue $domain -Force

                        if ($domain -and $domainWhitelist -contains $domain) {
                            $excluded += $acc
                        }
                        else {
                            $filtered += $acc
                        }
                    }

                    if ($excluded.Count -gt 0) {
                        Write-Host "[Whitelist] $($excluded.Count) Konten ausgenommen:" -ForegroundColor Green
                        foreach ($e in $excluded) {
                            Write-Host "  [GESCHUETZT] $($e.DisplayName) ($($e.Mail))" -ForegroundColor Green
                        }
                        Write-Host ""
                    }

                    $inactiveAccounts = $filtered
                    Write-Host "[Whitelist] $totalInactive -> $($inactiveAccounts.Count) Konten nach Domain-Filter`n" -ForegroundColor Cyan
                }
                else {
                    foreach ($acc in $inactiveAccounts) {
                        $domain = Get-DomainFromAddress -Address $acc.Mail
                        $acc | Add-Member -NotePropertyName "GuestDomain" -NotePropertyValue $domain -Force
                    }
                }

                if (($inactiveAccounts | Measure-Object).Count -eq 0) {
                    Write-Host "Keine Konten zum Deaktivieren nach Whitelist-Filter.`n" -ForegroundColor Yellow
                    continue
                }

                # Interaktive Auswahl
                $selected = Select-GuestAccounts -Accounts $inactiveAccounts -ActionLabel "Deaktivierung"

                if (($selected | Measure-Object).Count -eq 0) {
                    Write-Host "Keine Konten ausgewaehlt. Abbruch.`n" -ForegroundColor Yellow
                    continue
                }

                # WhatIf ausfuehren
                $userIds = $selected | ForEach-Object { $_.UserId }
                Remove-M365GuestAccounts -UserIds $userIds -Action Disable -WhatIf

                # Ausfuehren?
                $execute = Read-Host "`nDeaktivierung ausfuehren? (Ja/Nein)"
                if ($execute -in @("Ja", "J", "ja", "j")) {
                    Remove-M365GuestAccounts -UserIds $userIds -Action Disable -Force
                }
                else {
                    Write-Host "Abgebrochen.`n" -ForegroundColor Yellow
                }
            }

            # ==============================================================
            # 3. LOESCHEN (Bereits deaktivierte Konten)
            # ==============================================================
            "3" {
                # Audit-CSV pruefen
                $csvPath = Find-LatestAuditCsv
                if (-not $csvPath) {
                    Write-Host "`n[FEHLER] Kein Audit gefunden!" -ForegroundColor Red
                    Write-Host "Bitte fuehren Sie zuerst Schritt [1] aus, um einen Audit-Report zu erzeugen.`n" -ForegroundColor Yellow
                    continue
                }

                Write-Host "`n--- Schritt 3: Deaktivierte Gastkonten loeschen ---" -ForegroundColor Red
                Write-Host "Lade Audit-Daten: $csvPath" -ForegroundColor Gray

                $auditData = Import-Csv -Path $csvPath -Delimiter ";"

                # Nur deaktivierte Gastkonten
                $disabledAccounts = $auditData | Where-Object {
                    $_.AccountEnabled -eq "False"
                }

                $totalDisabled = ($disabledAccounts | Measure-Object).Count
                Write-Host "[OK] $totalDisabled deaktivierte Gastkonten gefunden.`n" -ForegroundColor Cyan

                if ($totalDisabled -eq 0) {
                    Write-Host "Keine deaktivierten Konten zum Loeschen vorhanden." -ForegroundColor Green
                    Write-Host "Tipp: Fuehren Sie zuerst Schritt [2] aus um inaktive Konten zu deaktivieren.`n" -ForegroundColor Gray
                    continue
                }

                # Domain-Whitelist anzeigen und anwenden
                if ($domainWhitelist.Count -gt 0) {
                    Write-Host "[Whitelist] Folgende Domains werden uebersprungen:" -ForegroundColor Green
                    foreach ($d in $domainWhitelist) {
                        Write-Host "  - $d" -ForegroundColor Green
                    }
                    Write-Host ""

                    $filtered = @()
                    $excluded = @()
                    foreach ($acc in $disabledAccounts) {
                        $domain = Get-DomainFromAddress -Address $acc.Mail
                        if (-not $domain) { $domain = Get-DomainFromAddress -Address $acc.UserPrincipalName }
                        $acc | Add-Member -NotePropertyName "GuestDomain" -NotePropertyValue $domain -Force

                        if ($domain -and $domainWhitelist -contains $domain) {
                            $excluded += $acc
                        }
                        else {
                            $filtered += $acc
                        }
                    }

                    if ($excluded.Count -gt 0) {
                        Write-Host "[Whitelist] $($excluded.Count) Konten ausgenommen:" -ForegroundColor Green
                        foreach ($e in $excluded) {
                            Write-Host "  [GESCHUETZT] $($e.DisplayName) ($($e.Mail))" -ForegroundColor Green
                        }
                        Write-Host ""
                    }

                    $disabledAccounts = $filtered
                    Write-Host "[Whitelist] $totalDisabled -> $($disabledAccounts.Count) Konten nach Domain-Filter`n" -ForegroundColor Cyan
                }
                else {
                    foreach ($acc in $disabledAccounts) {
                        $domain = Get-DomainFromAddress -Address $acc.Mail
                        $acc | Add-Member -NotePropertyName "GuestDomain" -NotePropertyValue $domain -Force
                    }
                }

                if (($disabledAccounts | Measure-Object).Count -eq 0) {
                    Write-Host "Keine Konten zum Loeschen nach Whitelist-Filter.`n" -ForegroundColor Yellow
                    continue
                }

                # Interaktive Auswahl
                $selected = Select-GuestAccounts -Accounts $disabledAccounts -ActionLabel "Loeschung"

                if (($selected | Measure-Object).Count -eq 0) {
                    Write-Host "Keine Konten ausgewaehlt. Abbruch.`n" -ForegroundColor Yellow
                    continue
                }

                # WhatIf ausfuehren
                $userIds = $selected | ForEach-Object { $_.UserId }
                Remove-M365GuestAccounts -UserIds $userIds -Action Delete -WhatIf

                # Ausfuehren?
                Write-Host "`n[!] ACHTUNG: Geloeschte Konten koennen nur innerhalb von 30 Tagen wiederhergestellt werden!" -ForegroundColor Red
                $execute = Read-Host "Loeschung ausfuehren? (Ja/Nein)"
                if ($execute -in @("Ja", "J", "ja", "j")) {
                    Remove-M365GuestAccounts -UserIds $userIds -Action Delete -Force
                }
                else {
                    Write-Host "Abgebrochen.`n" -ForegroundColor Yellow
                }
            }

            # ==============================================================
            # BEENDEN
            # ==============================================================
            { $_ -in @("Q", "q") } {
                Write-Host "`nAuf Wiedersehen!`n" -ForegroundColor Cyan
                return
            }
            default {
                Write-Warning "Ungueltige Auswahl. Bitte 1, 2, 3 oder Q eingeben."
            }
        }
    }
}

# ============================================================================
# Modul-Export & Auto-Start
# ============================================================================

# Wenn direkt ausgefuehrt: interaktives Menue starten
if ($MyInvocation.InvocationName -ne '.') {
    Show-GovernanceMenu
}
