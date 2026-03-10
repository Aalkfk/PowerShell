<#
.SYNOPSIS
    M365 Guest Account Governance Tool
.DESCRIPTION
    PowerShell-basiertes Governance-Tool fuer Microsoft 365 Gastkonten.
    - Audit: Gastkonten auslesen, Inaktivitaet und Alter pruefen, Berichte generieren (HTML + CSV/JSON)
    - Cleanup: Gastkonten deaktivieren/loeschen mit WhatIf-Unterstuetzung und CSV-Massenimport
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
        Audit-Logs sind auf max. 30 Tage begrenzt.
    #>
    [CmdletBinding()]
    param()

    $cache = @{}

    try {
        Write-Host "  Lade Einladungs-Audit-Logs (letzte 30 Tage)..." -ForegroundColor Gray
        $auditUri = "https://graph.microsoft.com/beta/auditLogs/directoryAudits?" +
            "`$filter=activityDisplayName eq 'Invite external user'&" +
            "`$top=500&" +
            "`$select=targetResources,initiatedBy,activityDateTime"

        $auditLogs = Get-AllGraphPages -Uri $auditUri

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

            if ($targetId -and -not $cache.ContainsKey($targetId)) {
                $initiator = $log.initiatedBy
                $inviterName = "Unbekannt"
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

                $cache[$targetId] = [PSCustomObject]@{
                    InviterId   = $inviterId
                    InviterName = $inviterName
                    InviterMail = $inviterMail
                    Source       = "Audit-Log"
                }
            }
        }

        Write-Host "  [OK] $($cache.Count) Einladungen aus Audit-Logs geladen." -ForegroundColor Gray
    }
    catch {
        Write-Warning "  Audit-Logs konnten nicht geladen werden: $($_.Exception.Message)"
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
    if ($InviterInfo) {
        $inviterName = $InviterInfo.InviterName
        $inviterMail = $InviterInfo.InviterMail
    }

    return [PSCustomObject]@{
        UserId                = $User.id
        DisplayName           = $User.displayName
        Mail                  = $User.mail
        UserPrincipalName     = $User.userPrincipalName
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
    $pendingCount = ($results | Where-Object { $_.InvitationStatus -eq "Ausstehend" }).Count
    $stats = [PSCustomObject]@{
        Gesamt           = $totalGuests
        Aktiv            = ($results | Where-Object { -not $_.IsInactive -and -not $_.IsExpired }).Count
        Inaktiv          = ($results | Where-Object { $_.IsInactive }).Count
        Abgelaufen       = ($results | Where-Object { $_.IsExpired }).Count
        Deaktiviert      = ($results | Where-Object { -not $_.AccountEnabled }).Count
        KriteriumHoch    = ($results | Where-Object { $_.Severity -eq "Hoch" }).Count
        Ausstehend       = $pendingCount
    }

    Write-Host "`n--- Zusammenfassung ---" -ForegroundColor Cyan
    Write-Host "  Gesamt              : $($stats.Gesamt)"
    Write-Host "  Aktiv               : $($stats.Aktiv)" -ForegroundColor Green
    Write-Host "  Inaktiv (>$InactiveDays d)    : $($stats.Inaktiv)" -ForegroundColor Yellow
    Write-Host "  Abgelaufen (>$MaxAgeDays d)  : $($stats.Abgelaufen)" -ForegroundColor Red
    Write-Host "  Deaktiviert         : $($stats.Deaktiviert)" -ForegroundColor DarkGray
    Write-Host "  Kritisch (beides)   : $($stats.KriteriumHoch)" -ForegroundColor Magenta
    Write-Host "  Einladung ausstehend: $($stats.Ausstehend)" -ForegroundColor DarkYellow
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
        $lastSignInFmt   = if ($r.LastSignIn) { $r.LastSignIn.ToString("dd.MM.yyyy") } else { "-" }
        $lastNonIntFmt   = if ($r.LastNonInteractive) { $r.LastNonInteractive.ToString("dd.MM.yyyy") } else { "-" }
        $created         = if ($r.CreatedDateTime) { $r.CreatedDateTime.ToString("dd.MM.yyyy") } else { "Unbekannt" }
        $inviterDisplay  = if ($r.InviterMail) { "<span title=`"$($r.InviterMail)`">$($r.InviterName)</span>" } else { $r.InviterName }

        $flaggedRows += @"
        <tr class="$severityClass">
            <td title="$($r.UserId)">$($r.DisplayName)</td>
            <td>$($r.Mail)</td>
            <td>$created</td>
            <td title="Interaktiv: $lastSignInFmt | Nicht-interaktiv: $lastNonIntFmt">$lastAct</td>
            <td>$($r.DaysSinceActivity)</td>
            <td>$($r.DaysSinceCreation)</td>
            <td>$statusBadge</td>
            <td>$invitationBadge</td>
            <td>$inviterDisplay</td>
            <td><span class="severity-badge $severityClass">$($r.Severity)</span></td>
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
        $lastSignInFmt   = if ($r.LastSignIn) { $r.LastSignIn.ToString("dd.MM.yyyy") } else { "-" }
        $lastNonIntFmt   = if ($r.LastNonInteractive) { $r.LastNonInteractive.ToString("dd.MM.yyyy") } else { "-" }
        $created         = if ($r.CreatedDateTime) { $r.CreatedDateTime.ToString("dd.MM.yyyy") } else { "Unbekannt" }
        $inviterDisplay  = if ($r.InviterMail) { "<span title=`"$($r.InviterMail)`">$($r.InviterName)</span>" } else { $r.InviterName }

        $okRows += @"
        <tr>
            <td title="$($r.UserId)">$($r.DisplayName)</td>
            <td>$($r.Mail)</td>
            <td>$created</td>
            <td title="Interaktiv: $lastSignInFmt | Nicht-interaktiv: $lastNonIntFmt">$lastAct</td>
            <td>$($r.DaysSinceActivity)</td>
            <td>$($r.DaysSinceCreation)</td>
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
            transition: box-shadow 0.2s;
        }
        .stat-card:hover { box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
        .stat-card .stat-value { font-size: 36px; font-weight: 700; }
        .stat-card .stat-label { font-size: 13px; color: var(--muted); margin-top: 4px; }
        .stat-card.total .stat-value { color: var(--primary); }
        .stat-card.ok .stat-value { color: var(--success); }
        .stat-card.warn .stat-value { color: var(--warning); }
        .stat-card.danger .stat-value { color: var(--danger); }
        .stat-card.critical .stat-value { color: #881798; }
        .stat-card.pending .stat-value { color: #ca5010; }

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
        </div>

        <!-- Stats -->
        <div class="stats-grid">
            <div class="stat-card total">
                <div class="stat-value">$($Stats.Gesamt)</div>
                <div class="stat-label">Gastkonten Gesamt</div>
            </div>
            <div class="stat-card ok">
                <div class="stat-value">$okCount</div>
                <div class="stat-label">Konform</div>
            </div>
            <div class="stat-card warn">
                <div class="stat-value">$($Stats.Inaktiv)</div>
                <div class="stat-label">Inaktiv (&gt;$InactiveDays d)</div>
            </div>
            <div class="stat-card danger">
                <div class="stat-value">$($Stats.Abgelaufen)</div>
                <div class="stat-label">Abgelaufen (&gt;$MaxAgeDays d)</div>
            </div>
            <div class="stat-card critical">
                <div class="stat-value">$($Stats.KriteriumHoch)</div>
                <div class="stat-label">Kritisch (beides)</div>
            </div>
            <div class="stat-card pending">
                <div class="stat-value">$($Stats.Ausstehend)</div>
                <div class="stat-label">Einladung ausstehend</div>
            </div>
            <div class="stat-card">
                <div class="stat-value" style="color: var(--muted);">$($Stats.Deaktiviert)</div>
                <div class="stat-label">Bereits deaktiviert</div>
            </div>
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
                <span class="count">$flaggedCount</span>
            </div>
            <div class="filter-bar">
                <input type="text" id="filterFlagged" placeholder="Filtern nach Name, E-Mail, Einlader, Flags..." onkeyup="filterTable('flaggedTable', this.value)">
            </div>
            <div style="overflow-x: auto;">
                <table id="flaggedTable">
                    <thead>
                        <tr>
                            <th onclick="sortTable('flaggedTable', 0)">Name</th>
                            <th onclick="sortTable('flaggedTable', 1)">E-Mail</th>
                            <th onclick="sortTable('flaggedTable', 2)">Erstellt</th>
                            <th onclick="sortTable('flaggedTable', 3)">Letzte Aktivitaet</th>
                            <th onclick="sortTable('flaggedTable', 4)">Tage inaktiv</th>
                            <th onclick="sortTable('flaggedTable', 5)">Alter (Tage)</th>
                            <th onclick="sortTable('flaggedTable', 6)">Status</th>
                            <th onclick="sortTable('flaggedTable', 7)">Einladung</th>
                            <th onclick="sortTable('flaggedTable', 8)">Eingeladen von</th>
                            <th onclick="sortTable('flaggedTable', 9)">Schwere</th>
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
                <span class="count" style="background: var(--success);">$okCount</span>
            </div>
            <div id="okSection" class="collapsible-content">
                <div class="filter-bar">
                    <input type="text" id="filterOk" placeholder="Filtern nach Name, E-Mail, Einlader..." onkeyup="filterTable('okTable', this.value)">
                </div>
                <div style="overflow-x: auto;">
                    <table id="okTable">
                        <thead>
                            <tr>
                                <th onclick="sortTable('okTable', 0)">Name</th>
                                <th onclick="sortTable('okTable', 1)">E-Mail</th>
                                <th onclick="sortTable('okTable', 2)">Erstellt</th>
                                <th onclick="sortTable('okTable', 3)">Letzte Aktivitaet</th>
                                <th onclick="sortTable('okTable', 4)">Tage inaktiv</th>
                                <th onclick="sortTable('okTable', 5)">Alter (Tage)</th>
                                <th onclick="sortTable('okTable', 6)">Status</th>
                                <th onclick="sortTable('okTable', 7)">Einladung</th>
                                <th onclick="sortTable('okTable', 8)">Eingeladen von</th>
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
        function toggleSection(id, el) {
            var section = document.getElementById(id);
            var icon = el.querySelector('.toggle-icon');
            section.classList.toggle('open');
            icon.classList.toggle('open');
        }

        function filterTable(tableId, filter) {
            var table = document.getElementById(tableId);
            var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
            var lowerFilter = filter.toLowerCase();
            for (var i = 0; i < rows.length; i++) {
                var text = rows[i].textContent.toLowerCase();
                rows[i].style.display = text.indexOf(lowerFilter) > -1 ? '' : 'none';
            }
        }

        function sortTable(tableId, colIndex) {
            var table = document.getElementById(tableId);
            var tbody = table.getElementsByTagName('tbody')[0];
            var rows = Array.from(tbody.getElementsByTagName('tr'));
            var asc = table.getAttribute('data-sort-asc-' + colIndex) !== 'true';
            table.setAttribute('data-sort-asc-' + colIndex, asc);

            rows.sort(function(a, b) {
                var aText = a.cells[colIndex].textContent.trim();
                var bText = b.cells[colIndex].textContent.trim();
                var aNum = parseFloat(aText);
                var bNum = parseFloat(bText);
                if (!isNaN(aNum) && !isNaN(bNum)) {
                    return asc ? aNum - bNum : bNum - aNum;
                }
                return asc ? aText.localeCompare(bText, 'de') : bText.localeCompare(aText, 'de');
            });

            rows.forEach(function(row) { tbody.appendChild(row); });
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
    .PARAMETER WhatIf
        Zeigt an, was passieren wuerde, ohne Aenderungen vorzunehmen.
    .PARAMETER Force
        Ueberspringt die Sicherheitsabfrage (nicht empfohlen).
    .EXAMPLE
        Remove-M365GuestAccounts -CsvPath ".\cleanup.csv" -Action Disable -WhatIf
    .EXAMPLE
        $audit = Get-M365GuestReport -PassThru
        Remove-M365GuestAccounts -FromAudit $audit -SeverityFilter "Hoch" -Action Delete -WhatIf
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
        [switch]$Force
    )

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host " M365 Gastkonto-Cleanup" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " Aktion  : $Action"
    Write-Host " WhatIf  : $($WhatIfPreference -or $PSBoundParameters.ContainsKey('WhatIf'))"
    Write-Host "========================================`n" -ForegroundColor Cyan

    # Zielkonten ermitteln
    $targets = @()

    switch ($PSCmdlet.ParameterSetName) {
        "ById" {
            foreach ($id in $UserIds) {
                $targets += [PSCustomObject]@{ UserId = $id; Source = "Parameter" }
            }
        }
        "ByCsv" {
            Write-Host "Lade CSV: $CsvPath" -ForegroundColor Yellow
            $csvData = Import-Csv -Path $CsvPath -Delimiter ";"

            # Unterstuetzt sowohl 'UserId' als auch 'UserPrincipalName' Spalte
            foreach ($row in $csvData) {
                $id = if ($row.UserId) { $row.UserId } elseif ($row.UserPrincipalName) { $row.UserPrincipalName } else { $null }
                if ($id) {
                    $targets += [PSCustomObject]@{
                        UserId      = $id
                        Source      = "CSV"
                        DisplayName = $row.DisplayName
                        Mail        = $row.Mail
                    }
                }
            }
            Write-Host "[OK] $($targets.Count) Eintraege aus CSV geladen.`n" -ForegroundColor Green
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
                    Flags       = $item.Flags
                }
            }
            Write-Host "[OK] $($targets.Count) Konten aus Audit gefiltert (Schwere: $SeverityFilter).`n" -ForegroundColor Green
        }
    }

    if ($targets.Count -eq 0) {
        Write-Host "Keine Zielkonten gefunden. Abbruch." -ForegroundColor Yellow
        return
    }

    # Vorschau anzeigen
    Write-Host "Betroffene Konten ($($targets.Count)):" -ForegroundColor Cyan
    Write-Host ("-" * 80)
    $targets | Format-Table -AutoSize -Property @(
        @{N = "UserId"; E = { $_.UserId.Substring(0, [Math]::Min(8, $_.UserId.Length)) + "..." } },
        "DisplayName", "Mail", "Source"
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
                # Beta-API fuer Konsistenz
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($target.UserId)" `
                    -Body @{ accountEnabled = $false } | Out-Null
                Write-Host "  [OK] Deaktiviert: $name" -ForegroundColor Green
                $results += [PSCustomObject]@{ UserId = $target.UserId; Name = $name; Action = "Disabled"; Status = "Success"; Error = "" }
                $successCount++
            }
            else {
                Invoke-MgGraphRequest -Method DELETE `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($target.UserId)" | Out-Null
                Write-Host "  [OK] Geloescht: $name" -ForegroundColor Green
                $results += [PSCustomObject]@{ UserId = $target.UserId; Name = $name; Action = "Deleted"; Status = "Success"; Error = "" }
                $successCount++
            }
        }
        catch {
            Write-Warning "  [FEHLER] $name : $_"
            $results += [PSCustomObject]@{ UserId = $target.UserId; Name = $name; Action = $Action; Status = "Error"; Error = $_.Exception.Message }
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
# Hauptmenue (Interaktiver Modus)
# ============================================================================

function Show-GovernanceMenu {
    <#
    .SYNOPSIS
        Zeigt ein interaktives Menue fuer das Governance-Tool.
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

    while ($true) {
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host " M365 Guest Governance Tool" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "  [1] Gastkonto-Audit durchfuehren"
        Write-Host "  [2] Gastkonten deaktivieren (mit WhatIf)"
        Write-Host "  [3] Gastkonten loeschen (mit WhatIf)"
        Write-Host "  [4] Audit + Kritische deaktivieren"
        Write-Host "  [5] CSV-Import: Massenbereinigung"
        Write-Host "  [Q] Beenden"
        Write-Host "========================================" -ForegroundColor Cyan

        $choice = Read-Host "`nAuswahl"

        switch ($choice) {
            "1" {
                Get-M365GuestReport
            }
            "2" {
                $audit = Get-M365GuestReport -PassThru
                if ($audit) {
                    Remove-M365GuestAccounts -FromAudit $audit -SeverityFilter "Alle" -Action Disable -WhatIf
                    $execute = Read-Host "`nMoechten Sie die Deaktivierung ausfuehren? (Ja/Nein)"
                    if ($execute -in @("Ja", "J", "ja", "j")) {
                        Remove-M365GuestAccounts -FromAudit $audit -SeverityFilter "Alle" -Action Disable
                    }
                }
            }
            "3" {
                $audit = Get-M365GuestReport -PassThru
                if ($audit) {
                    Remove-M365GuestAccounts -FromAudit $audit -SeverityFilter "Alle" -Action Delete -WhatIf
                    $execute = Read-Host "`nMoechten Sie die Loeschung ausfuehren? (Ja/Nein)"
                    if ($execute -in @("Ja", "J", "ja", "j")) {
                        Remove-M365GuestAccounts -FromAudit $audit -SeverityFilter "Alle" -Action Delete
                    }
                }
            }
            "4" {
                $audit = Get-M365GuestReport -PassThru
                if ($audit) {
                    Write-Host "`n[WhatIf-Vorschau fuer kritische Konten (Schwere: Hoch)]" -ForegroundColor Yellow
                    Remove-M365GuestAccounts -FromAudit $audit -SeverityFilter "Hoch" -Action Disable -WhatIf
                    $execute = Read-Host "`nKritische Konten deaktivieren? (Ja/Nein)"
                    if ($execute -in @("Ja", "J", "ja", "j")) {
                        Remove-M365GuestAccounts -FromAudit $audit -SeverityFilter "Hoch" -Action Disable
                    }
                }
            }
            "5" {
                $csvFile = Read-Host "Pfad zur CSV-Datei"
                if (Test-Path $csvFile) {
                    $action = Read-Host "Aktion (Disable/Delete)"
                    if ($action -in @("Disable", "Delete")) {
                        Remove-M365GuestAccounts -CsvPath $csvFile -Action $action -WhatIf
                        $execute = Read-Host "`nAktion ausfuehren? (Ja/Nein)"
                        if ($execute -in @("Ja", "J", "ja", "j")) {
                            Remove-M365GuestAccounts -CsvPath $csvFile -Action $action
                        }
                    }
                    else {
                        Write-Warning "Ungueltige Aktion. Bitte 'Disable' oder 'Delete' angeben."
                    }
                }
                else {
                    Write-Warning "Datei nicht gefunden: $csvFile"
                }
            }
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
# Modul-Export & Auto-Start
# ============================================================================

# Wenn direkt ausgefuehrt: interaktives Menue starten
if ($MyInvocation.InvocationName -ne '.') {
    Show-GovernanceMenu
}
