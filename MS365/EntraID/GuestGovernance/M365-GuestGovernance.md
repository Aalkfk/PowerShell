# M365-GuestGovernance.ps1 - Technische Dokumentation

| Eigenschaft        | Wert                                                                 |
|--------------------|----------------------------------------------------------------------|
| **Dateiname**      | `M365-GuestGovernance.ps1`                                           |
| **Version**        | 1.0                                                                  |
| **Erstellt**       | 2026-03                                                              |
| **PowerShell**     | >= 5.1                                                               |
| **Abhaengigkeiten**| Microsoft.Graph PowerShell SDK (`Microsoft.Graph.Authentication`)     |
| **API-Endpunkt**   | Microsoft Graph Beta (`https://graph.microsoft.com/beta/`)           |
| **Speicherort**    | `MS365/EntraID/GuestGovernance/`                                     |

---

## 1. Zweck

Das Skript implementiert einen automatisierten Governance-Prozess fuer externe Gastkonten (B2B) in Microsoft Entra ID (ehemals Azure AD). Es prueft regelmaessig alle Gastkonten auf Inaktivitaet, Kontoalter und DLP-Risiken, erstellt Berichte in mehreren Formaten und ermoeglicht die kontrollierte Deaktivierung bzw. Loeschung nicht mehr benoetigter Konten.

---

## 2. Funktionsumfang

### 2.1 Audit (Schritt 1)

Liest saemtliche Gastkonten (`userType eq 'Guest'`) aus dem Tenant ueber die **Microsoft Graph Beta-API** und fuehrt folgende Pruefungen durch:

| Pruefung                 | Schwellenwert   | Flag                       | Schweregrad |
|--------------------------|-----------------|----------------------------|-------------|
| Inaktivitaet             | > 60 Tage       | `INAKTIV_{n}_TAGE`         | Niedrig     |
| Kontoalter               | > 365 Tage      | `ABGELAUFEN_{n}_TAGE`      | Mittel      |
| Inaktiv + Abgelaufen     | beide erfuellt   | beide Flags                | Hoch        |
| Keine Anmeldung          | kein SignIn      | `KEINE_ANMELDUNG`          | Niedrig     |
| Einladung ausstehend     | PendingAcceptance| `EINLADUNG_AUSSTEHEND`     | -           |
| Freemailer-Domain (DLP)  | bekannte Domain  | `FREEMAILER`               | -           |

**Letzte Aktivitaet** wird als Maximum aus interaktivem Sign-In (`lastSignInDateTime`) und nicht-interaktivem Sign-In (`lastNonInteractiveSignInDateTime`) berechnet.

#### Einlader-Ermittlung

Die Ermittlung des Einladers/Sponsors erfolgt zweistufig:

1. **Sponsor-Beziehung** (Beta-API): `GET /beta/users/{id}/sponsors`
2. **Audit-Log Fallback**: Durchsucht drei Aktivitaetstypen:
   - `Invite external user`
   - `Add user`
   - `Redeem external user invite`

> **Hinweis:** Audit-Logs sind lizenzabhaengig auf ca. 30 Tage begrenzt. Fuer aeltere Eintraege wird ausschliesslich die Sponsor-Beziehung genutzt.

#### Freemailer-Erkennung (DLP)

Das Skript erkennt Gastkonten mit privaten E-Mail-Adressen (DLP-Risiko). Aktuell enthaltene Freemailer-Domains (35+):

- **Google**: gmail.com, googlemail.com
- **Microsoft**: outlook.com, outlook.de, hotmail.com, hotmail.de, live.com, live.de, msn.com
- **Deutsche**: gmx.de, gmx.net, gmx.at, gmx.ch, web.de, t-online.de, freenet.de, arcor.de, online.de, email.de, mail.de, posteo.de, mailbox.org
- **Yahoo**: yahoo.com, yahoo.de, ymail.com
- **Apple**: icloud.com, me.com, mac.com
- **Weitere**: aol.com, aol.de, protonmail.com, proton.me, pm.me, zoho.com, tutanota.com, tuta.io, gmx.com, mail.com

Die Domain-Liste kann im Skript unter `$Script:Config.FreemailerDomains` erweitert werden.

#### Export-Formate

| Format | Dateiname                        | Inhalt                                                       |
|--------|----------------------------------|--------------------------------------------------------------|
| CSV    | `GuestAudit_{timestamp}.csv`     | Alle Felder, Trennzeichen `;`, UTF-8                         |
| JSON   | `GuestAudit_{timestamp}.json`    | Metadaten + Statistiken + alle Gastkonten                    |
| HTML   | `GuestAudit_{timestamp}.html`    | Grafischer Report mit interaktiven Funktionen                |

Alle Reports werden im Verzeichnis `./Reports/` abgelegt.

#### HTML-Report Features

- **Stat-Karten**: Gesamt, Konform, Inaktiv, Abgelaufen, Kritisch, Ausstehend, Freemailer, Deaktiviert
- **Klickbare Filter**: Stat-Karten filtern die Tabellen bei Klick, erneuter Klick hebt den Filter auf
- **Sortierbare Spalten**: Alle Spalten per Klick sortierbar, Datums-/Zahlenspalten verwenden `data-sort`-Attribute fuer korrekte Sortierung
- **Suchfeld**: Freitextsuche ueber beide Tabellen
- **Compliance-Rate**: Visueller Fortschrittsbalken
- **Freemailer-Hervorhebung**: Rote Badges und Hintergrundfarbe fuer DLP-Risiko-Konten
- **Domain-Spalte**: Zeigt die Gast-Domain mit optionalem Freemailer-Badge
- **Tooltip**: Hover ueber Aktivitaetsdatum zeigt Aufschluesselung (interaktiv / nicht-interaktiv)
- **Druckoptimierung**: CSS-Anpassungen fuer Druckausgabe

### 2.2 Deaktivierung (Schritt 2)

Deaktiviert Gastkonten mit einer Inaktivitaet von ueber 60 Tagen:

- Basiert auf der CSV-Datei aus Schritt 1 (Fehler, wenn kein Audit vorhanden)
- Filtert nur Konten mit `IsInactive = True` und `AccountEnabled = True`
- Wendet Domain-Whitelist an (geschuetzte Domains werden uebersprungen)
- Interaktive Kontenauswahl: nummerierte Liste mit Abwahl-Moeglichkeit
- **WhatIf zuerst**: Zeigt geplante Aktionen an, dann Bestaetigung
- API-Aufruf: `PATCH /v1.0/users/{id}` mit `{ accountEnabled: false }`

### 2.3 Loeschung (Schritt 3)

Loescht bereits deaktivierte Gastkonten:

- Basiert auf der CSV-Datei aus Schritt 1
- Filtert nur Konten mit `AccountEnabled = False`
- Gleiche Whitelist- und Auswahllogik wie Schritt 2
- **WhatIf zuerst** mit zusaetzlicher Warnung (30-Tage-Wiederherstellungsfenster)
- API-Aufruf: `DELETE /v1.0/users/{id}`

### 2.4 Domain-Whitelist

Domains koennen auf drei Wegen vom Cleanup ausgeschlossen werden:

| Quelle                        | Konfiguration                                    |
|-------------------------------|--------------------------------------------------|
| Skript-Konfiguration          | `$Script:Config.ExcludedDomains` Array           |
| Datei im Skript-Verzeichnis   | `excluded_domains.txt` (eine Domain pro Zeile)   |
| Parameter                     | `-ExcludedDomains` / `-ExcludedDomainsFile`       |

Whitelist-Domains werden beim **Audit weiterhin erfasst**, aber bei Deaktivierung/Loeschung uebersprungen und als `[GESCHUETZT]` markiert.

---

## 3. Voraussetzungen

### 3.1 PowerShell-Module

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

### 3.2 Microsoft Graph Berechtigungen

| Berechtigung            | Typ          | Zweck                                    |
|-------------------------|--------------|------------------------------------------|
| `User.Read.All`         | Delegated    | Gastkonten lesen                         |
| `AuditLog.Read.All`     | Delegated    | Audit-Logs fuer Einlader-Ermittlung      |
| `Directory.ReadWrite.All`| Delegated   | Konten deaktivieren / loeschen           |

### 3.3 Lizenzen

- **Minimum**: Microsoft Entra ID Free (eingeschraenkte Audit-Log-Aufbewahrung)
- **Empfohlen**: Entra ID P1/P2 (erweiterte Audit-Log-Aufbewahrung, SignInActivity-Daten)

---

## 4. Ausfuehrung

### 4.1 Interaktiver Modus (Standard)

```powershell
.\M365-GuestGovernance.ps1
```

Startet das interaktive Menue mit den Schritten 1-3.

### 4.2 Dot-Sourcing (Funktionen laden)

```powershell
. .\M365-GuestGovernance.ps1
```

Laedt alle Funktionen in die aktuelle Sitzung fuer direkte Aufrufe:

```powershell
# Audit mit angepassten Schwellenwerten
Get-M365GuestReport -InactiveDays 90 -MaxAgeDays 180 -PassThru

# Cleanup per CSV mit WhatIf
Remove-M365GuestAccounts -CsvPath ".\cleanup.csv" -Action Disable -WhatIf

# Cleanup per Audit-Ergebnis
$audit = Get-M365GuestReport -PassThru
Remove-M365GuestAccounts -FromAudit $audit -SeverityFilter "Hoch" -Action Delete -WhatIf
```

---

## 5. Funktionsreferenz

### Oeffentliche Funktionen

| Funktion                     | Beschreibung                                                       |
|------------------------------|--------------------------------------------------------------------|
| `Get-M365GuestReport`        | Hauptfunktion Audit: liest, analysiert, exportiert                 |
| `Remove-M365GuestAccounts`   | Deaktiviert oder loescht Gastkonten (WhatIf-faehig)                |
| `Show-GovernanceMenu`        | Interaktives 3-Schritt-Menue                                       |

### Interne Hilfsfunktionen

| Funktion                     | Beschreibung                                                       |
|------------------------------|--------------------------------------------------------------------|
| `Test-Prerequisites`         | Prueft installierte PowerShell-Module                              |
| `Connect-M365Governance`     | Graph-Verbindung herstellen                                        |
| `Test-FreemailerDomain`      | Prueft ob eine Domain ein bekannter Freemailer ist                 |
| `Get-AllGraphPages`          | Paginierungs-Helper fuer Graph-API                                |
| `Get-DomainFromAddress`      | Extrahiert Domain aus Mail/UPN (inkl. `#EXT#`-Format)             |
| `Get-GuestInviter`           | Zweistufige Einlader-Ermittlung (Sponsor + Audit-Log)             |
| `Get-InvitationAuditCache`   | Baut Audit-Log-Cache fuer Einladungen                              |
| `Get-GuestActivityStatus`    | Berechnet alle Felder und Flags fuer ein Gastkonto                |
| `New-GuestHtmlReport`        | Generiert den interaktiven HTML-Report                             |
| `Select-GuestAccounts`       | Interaktive Shell-Liste fuer Kontenauswahl                         |
| `Get-DomainWhitelist`        | Baut Whitelist aus Config + Datei zusammen                         |
| `Find-LatestAuditCsv`        | Sucht die neueste Audit-CSV im Reports-Verzeichnis                |

---

## 6. Datenmodell

### Ausgabe-Objekt (pro Gastkonto)

| Feld                    | Typ       | Beschreibung                                         |
|-------------------------|-----------|------------------------------------------------------|
| `UserId`                | String    | Entra ID Object-ID                                   |
| `DisplayName`           | String    | Anzeigename                                          |
| `Mail`                  | String    | E-Mail-Adresse                                       |
| `UserPrincipalName`     | String    | UPN (inkl. `#EXT#`-Format)                           |
| `GuestDomain`           | String    | Extrahierte Domain des Gastes                        |
| `CreatedDateTime`       | DateTime  | Erstellungsdatum des Kontos                          |
| `LastSignIn`            | DateTime  | Letzter interaktiver Sign-In                         |
| `LastNonInteractive`    | DateTime  | Letzter nicht-interaktiver Sign-In                   |
| `LastActivity`          | DateTime  | Maximum aus beiden Sign-In-Werten                    |
| `DaysSinceActivity`     | Int       | Tage seit letzter Aktivitaet (-1 = nie)              |
| `DaysSinceCreation`     | Int       | Kontoalter in Tagen                                  |
| `AccountEnabled`        | Boolean   | Kontostatus aktiv/deaktiviert                        |
| `InvitationStatus`      | String    | Akzeptiert / Ausstehend / Unbekannt                  |
| `ExternalUserState`     | String    | Roher Wert aus Entra ID                              |
| `InviterName`           | String    | Name des Einladers                                   |
| `InviterMail`           | String    | E-Mail des Einladers                                 |
| `InviterSource`         | String    | Quelle: Sponsor / Audit-Log / Nicht ermittelbar      |
| `IsFreemailer`          | Boolean   | True wenn Domain ein bekannter Freemailer             |
| `IsInactive`            | Boolean   | True wenn Inaktivitaetsschwelle ueberschritten        |
| `IsExpired`             | Boolean   | True wenn Altersschwelle ueberschritten               |
| `Flags`                 | String    | Semikolon-getrennte Flag-Liste                       |
| `Severity`              | String    | OK / Niedrig / Mittel / Hoch                         |

---

## 7. Konfigurationsparameter

Alle Parameter befinden sich im Block `$Script:Config` am Anfang des Skripts:

| Parameter              | Standard | Beschreibung                                          |
|------------------------|----------|-------------------------------------------------------|
| `InactiveDaysThreshold`| 60       | Schwellenwert fuer Inaktivitaet (Tage)                |
| `MaxAgeDaysThreshold`  | 365      | Maximales Kontoalter (Tage)                           |
| `ReportOutputDir`      | ./Reports| Ausgabeverzeichnis fuer Reports                       |
| `DefaultScopes`        | siehe 3.2| Graph-Berechtigungen fuer `Connect-MgGraph`           |
| `ExcludedDomains`      | leer     | Inline-Whitelist (Array)                              |
| `FreemailerDomains`    | 35+ Domains | Bekannte Freemailer-Domains                       |
| `GraphUserProperties`  | 9 Felder | Abgefragte Beta-API Properties                       |

---

## 8. Dateien und Verzeichnisstruktur

```
GuestGovernance/
├── M365-GuestGovernance.ps1        # Hauptskript (Audit + Cleanup)
├── M365-GuestGovernance.md         # Diese Dokumentation
├── TOPdesk-GuestActivities.ps1     # TOPdesk-Integration (separates Skript)
├── TOPdesk-GuestActivities.md      # TOPdesk-Dokumentation
├── cleanup_template.csv            # CSV-Vorlage fuer Massenoperationen
├── excluded_domains.txt            # Domain-Whitelist (eine Domain pro Zeile)
├── topdesk_config.example.json     # TOPdesk-Konfigurationsvorlage
├── topdesk_config.json             # TOPdesk-Konfiguration (beim Start erstellt)
├── topdesk.credential              # DPAPI-verschluesselte Credentials (optional)
└── Reports/                        # Ausgabeverzeichnis (automatisch erstellt)
    ├── GuestAudit_{timestamp}.csv
    ├── GuestAudit_{timestamp}.json
    ├── GuestAudit_{timestamp}.html
    ├── Cleanup_{Action}_{timestamp}.csv
    └── TopDesk_Activities_{timestamp}.json
```

---

## 9. Technische Hinweise

### 9.1 Beta-API vs. v1.0

Das Skript verwendet die **Beta-API** (`graph.microsoft.com/beta/`) fuer:
- Zuverlaessige `signInActivity`-Daten (v1.0 liefert fuer Gastkonten oft `null`)
- Sponsor-Beziehung (`/users/{id}/sponsors`)
- `externalUserState` Feld

Der Header `ConsistencyLevel: eventual` ist fuer `$count` und `signInActivity`-Abfragen erforderlich.

### 9.2 Datentypen bei Beta-API

Die Beta-API gibt JSON-Hashtables zurueck (camelCase), keine .NET-Objekte (PascalCase):
- Richtig: `$User.signInActivity.lastSignInDateTime`
- Falsch: `$User.SignInActivity.LastSignInDateTime`

### 9.3 Paginierung

Alle Graph-API-Abfragen werden automatisch paginiert (`@odata.nextLink`) ueber die Funktion `Get-AllGraphPages`. Es werden bis zu 999 Eintraege pro Seite angefordert (`$top=999`).

### 9.4 UPN-Format fuer Gastkonten

Gastkonten verwenden das Format: `user_extern.de#EXT#@tenant.onmicrosoft.com`
Die Funktion `Get-DomainFromAddress` extrahiert die Ursprungsdomain korrekt.

---

## 10. Sicherheitsaspekte

| Aspekt                        | Implementierung                                          |
|-------------------------------|----------------------------------------------------------|
| Zugriffskontrolle             | Delegated Permissions via `Connect-MgGraph`              |
| Least Privilege               | Nur `Directory.ReadWrite.All` fuer Cleanup-Aktionen       |
| WhatIf-Modus                  | `SupportsShouldProcess` auf `Remove-M365GuestAccounts`   |
| Interaktive Bestaetigung      | Sicherheitsabfrage vor jeder destruktiven Aktion         |
| Domain-Whitelist              | Schutz fuer definierte Partner-/Dienstleister-Domains    |
| Audit-Trail                   | Cleanup-Ergebnisse werden als CSV protokolliert          |
| Keine Klartext-Credentials    | Verwendet `Connect-MgGraph` mit interaktivem Login       |

---

## 11. Aenderungshistorie

| Datum      | Aenderung                                                          |
|------------|--------------------------------------------------------------------|
| 2026-03    | Initiale Erstellung: Audit + Cleanup + HTML-Report                 |
| 2026-03    | SignInActivity auf Beta-API umgestellt (zuverlaessigere Daten)     |
| 2026-03    | Einlader-/Sponsor-Ermittlung hinzugefuegt (zweistufig)            |
| 2026-03    | Einladungsstatus (externalUserState) hinzugefuegt                  |
| 2026-03    | Domain-Whitelist fuer Cleanup implementiert                        |
| 2026-03    | HTML-Sortierung auf data-sort Attribute umgestellt                 |
| 2026-03    | Klickbare Stat-Karten mit Filterfunktion                           |
| 2026-03    | 3-Schritt-Workflow im interaktiven Menue                           |
| 2026-03    | Freemailer-Erkennung (DLP) mit 35+ Domains                        |
