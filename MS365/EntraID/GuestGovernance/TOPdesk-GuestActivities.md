# TOPdesk-GuestActivities.ps1 - Technische Dokumentation

| Eigenschaft        | Wert                                                                    |
|--------------------|-------------------------------------------------------------------------|
| **Dateiname**      | `TOPdesk-GuestActivities.ps1`                                           |
| **Version**        | 1.0                                                                     |
| **Erstellt**       | 2026-03                                                                 |
| **PowerShell**     | >= 5.1                                                                  |
| **Abhaengigkeiten**| Keine externen Module (nutzt `Invoke-RestMethod`)                       |
| **API-Endpunkt**   | TOPdesk REST API (`/tas/api/operationalActivities`)                     |
| **Speicherort**    | `MS365/EntraID/GuestGovernance/`                                        |
| **Eingabedaten**   | Audit-CSV von `M365-GuestGovernance.ps1`                                |

---

## 1. Zweck

Das Skript stellt die Schnittstelle zwischen dem M365 Guest Governance Audit und dem ITSM-System TOPdesk her. Es erstellt automatisiert **operative Aktivitaeten** in TOPdesk basierend auf den Ergebnissen des Gastkonto-Audits und unterstuetzt damit die lueckenlose Dokumentation von Governance-Massnahmen.

---

## 2. Funktionsumfang

### 2.1 CSV-Import

Das Skript liest die Audit-CSV-Dateien (`GuestAudit_*.csv`), die vom Hauptskript `M365-GuestGovernance.ps1` erzeugt werden. Wenn kein Pfad angegeben wird, sucht es automatisch die neueste CSV im `Reports/`-Verzeichnis.

### 2.2 Filtermodi

Vor der Erstellung koennen die Audit-Daten gefiltert werden:

| Filter         | Beschreibung                                          | Feld                    |
|----------------|-------------------------------------------------------|-------------------------|
| `Alle`         | Alle Eintraege (Standard)                             | -                       |
| `Inaktiv`      | Nur inaktive Konten (>60 Tage)                        | `IsInactive = True`     |
| `Abgelaufen`   | Nur abgelaufene Konten (>365 Tage)                    | `IsExpired = True`      |
| `Kritisch`     | Nur Schweregrad "Hoch" (inaktiv + abgelaufen)         | `Severity = Hoch`       |
| `Freemailer`   | Nur Konten mit Freemailer-Domain (DLP-Risiko)         | `IsFreemailer = True`   |
| `Deaktiviert`  | Nur bereits deaktivierte Konten                       | `AccountEnabled = False`|

### 2.3 Erstellungsmodi

| Modus     | Beschreibung                                                              |
|-----------|---------------------------------------------------------------------------|
| `Sammel`  | **Eine** operative Aktivitaet mit allen betroffenen Konten (Standard)     |
| `Einzel`  | **Je eine** operative Aktivitaet pro betroffenes Gastkonto                |

#### Sammel-Aktivitaet (Beschreibung)

Die automatisch generierte Beschreibung enthaelt:
- Zusammenfassung: Anzahl inaktiv, abgelaufen, kritisch, Freemailer, deaktiviert
- Liste aller betroffenen Konten mit Domain, Schweregrad und Status
- Domain-Verteilung (gruppiert)
- Audit-Datum und Herkunftsverweis

#### Einzel-Aktivitaet (Beschreibung)

Pro Konto wird eine Aktivitaet mit detaillierten Informationen erstellt:
- Alle Kontoattribute (Name, E-Mail, Domain, Erstelldatum, letzte Aktivitaet)
- DLP-Risiko-Warnung bei Freemailer-Domains
- Einladungsstatus und Einlader
- Schweregrad und Flags

### 2.4 WhatIf-Modus

Jede Aktion wird standardmaessig zuerst im **WhatIf-Modus** ausgefuehrt. Das Skript zeigt:
- Titel der geplanten Aktivitaet
- Anzahl betroffener Konten
- Kategorie, Operatorgruppe, Zeitraum

Erst nach expliziter Bestaetigung (`Ja`) erfolgt der API-Aufruf.

---

## 3. Voraussetzungen

### 3.1 TOPdesk API-Konto

Ein TOPdesk Operator-Konto mit folgenden Einstellungen:

| Einstellung                | Wert                                                    |
|----------------------------|---------------------------------------------------------|
| API-Konto Checkbox         | Aktiviert (Operator-Karte > Login-Daten)                |
| Application Password       | Erstellt (Operator-Karte > Autorisierung)               |
| Berechtigungsgruppe        | REST API Zugriff + Operations Management Schreibzugriff |
| Lizenz                     | Frei (API-Konten sind kostenlos)                        |

### 3.2 TOPdesk Modul

Das Operations Management Modul muss in der TOPdesk-Instanz lizenziert und aktiviert sein.

### 3.3 Netzwerk

HTTPS-Zugriff auf `https://<instanz>.topdesk.net` vom ausfuehrenden System.

---

## 4. Konfiguration

### 4.1 Konfigurationsdatei

Beim ersten Start wird interaktiv eine Konfigurationsdatei `topdesk_config.json` erstellt:

```json
{
    "BaseUrl": "https://firma.topdesk.net",
    "Defaults": {
        "Category": "IT-Sicherheit",
        "Subcategory": "Zugriffskontrolle",
        "OperatorGroup": "Security-Team",
        "Operator": "",
        "ActivityType": ""
    }
}
```

| Feld             | Beschreibung                                                         |
|------------------|----------------------------------------------------------------------|
| `BaseUrl`        | TOPdesk SaaS URL (ohne `/tas/api`)                                   |
| `Category`       | Standard-Kategorie fuer neue Aktivitaeten (Name oder GUID)          |
| `Subcategory`    | Standard-Unterkategorie (Name oder GUID)                             |
| `OperatorGroup`  | Standard-Operatorgruppe (Name oder GUID)                             |
| `Operator`       | Standard-Operator (Name oder GUID)                                   |
| `ActivityType`   | Standard-Aktivitaetstyp (Name oder GUID)                             |

> **Hinweis:** Alle Referenz-Felder akzeptieren sowohl Namen als auch GUIDs. Das Skript erkennt GUIDs automatisch am Format und waehlt das korrekte API-Feld (`id` vs. `name`/`groupName`).

### 4.2 Beispielkonfiguration

Eine Vorlage wird mitgeliefert: `topdesk_config.example.json`

### 4.3 Credential-Speicherung

Das Skript bietet drei Ebenen der Credential-Verwaltung:

| Ebene           | Methode                                   | Persistenz                  |
|-----------------|-------------------------------------------|-----------------------------|
| **Parameter**   | `-Credential` an `Connect-TopDesk`        | Nur Aufruf                  |
| **DPAPI-Datei** | `topdesk.credential` (verschluesselt)     | Dauerhaft (benutzergebunden)|
| **Sitzung**     | `$Script:TopDeskConfig.Credential`        | Nur PowerShell-Sitzung      |

#### DPAPI-Verschluesselung (empfohlen fuer Automatisierung)

Beim ersten erfolgreichen Login fragt das Skript, ob die Credentials gespeichert werden sollen. Bei Zustimmung wird eine Datei `topdesk.credential` im Skript-Verzeichnis erstellt.

**Sicherheitsmerkmale:**
- Verschluesselt via **Windows Data Protection API (DPAPI)**
- Gebunden an den **Windows-Benutzer** und die **Maschine** (nicht uebertragbar)
- Entschluesselung durch anderen Benutzer oder auf anderer Maschine schlaegt fehl
- Keine Klartext-Passwörter auf der Festplatte
- Verwaltung ueber Menue-Option [7] (speichern / ueberschreiben / loeschen)

#### Automatisierter Betrieb (Task Scheduler)

```powershell
# Einmalig: Credentials interaktiv speichern
.\TOPdesk-GuestActivities.ps1   # Menue [7] > Speichern

# Danach: Unbeaufsichtigte Ausfuehrung (Credentials werden automatisch geladen)
. .\TOPdesk-GuestActivities.ps1
Connect-TopDesk
New-TopDeskGuestActivity -Filter "Kritisch" -Mode "Sammel"
```

> **Wichtig:** Der geplante Task muss unter dem **gleichen Windows-Benutzer** laufen, der die Credentials gespeichert hat.

---

## 5. Ausfuehrung

### 5.1 Interaktiver Modus (Standard)

```powershell
.\TOPdesk-GuestActivities.ps1
```

Startet das interaktive Menue mit folgenden Optionen:

| Option | Aktion                                           | Filter      | Modus  |
|--------|--------------------------------------------------|-------------|--------|
| [1]    | Sammel-Aktivitaet erstellen                      | Waehlbar    | Sammel |
| [2]    | Einzel-Aktivitaeten erstellen                    | Waehlbar    | Einzel |
| [3]    | Nur Freemailer melden (DLP)                      | Freemailer  | Sammel |
| [4]    | Nur kritische Konten melden                      | Kritisch    | Sammel |
| [5]    | API-Einstellungen anzeigen                       | -           | -      |
| [6]    | Konfiguration aendern                            | -           | -      |
| [7]    | Anmeldedaten verwalten                           | -           | -      |
| [Q]    | Beenden                                          | -           | -      |

### 5.2 Dot-Sourcing (Funktionen laden)

```powershell
. .\TOPdesk-GuestActivities.ps1
```

Ermoeglicht direkte Funktionsaufrufe:

```powershell
# Konfiguration initialisieren
Initialize-TopDeskConfig

# Verbindung herstellen
Connect-TopDesk

# Sammel-Aktivitaet fuer alle inaktiven Konten (WhatIf)
New-TopDeskGuestActivity -Filter "Inaktiv" -Mode "Sammel" -WhatIf

# Einzel-Aktivitaeten fuer Freemailer mit eigener Beschreibung
New-TopDeskGuestActivity -Filter "Freemailer" -Mode "Einzel" `
    -BriefDescription "DLP-Pruefung Gastkonto" `
    -Category "Datenschutz" `
    -OperatorGroup "DLP-Team"

# Bestimmte CSV verwenden
New-TopDeskGuestActivity -CsvPath ".\Reports\GuestAudit_2026-03-10_143000.csv" `
    -Filter "Kritisch" -Mode "Sammel"
```

---

## 6. Funktionsreferenz

### Oeffentliche Funktionen

| Funktion                     | Beschreibung                                                            |
|------------------------------|-------------------------------------------------------------------------|
| `New-TopDeskGuestActivity`   | Erstellt operative Aktivitaeten aus Audit-CSV (Haupt-Funktion)          |
| `Show-TopDeskMenu`           | Interaktives Menue                                                       |
| `Connect-TopDesk`            | Verbindung testen und Credentials speichern                             |
| `Initialize-TopDeskConfig`   | Konfiguration interaktiv erstellen oder laden                           |

### Interne Hilfsfunktionen

| Funktion                          | Beschreibung                                                      |
|-----------------------------------|-------------------------------------------------------------------|
| `Get-TopDeskConfigPath`           | Gibt Pfad zur `topdesk_config.json` zurueck                      |
| `Import-TopDeskConfig`            | Laedt Konfiguration aus JSON                                      |
| `Export-TopDeskConfig`            | Speichert Konfiguration als JSON                                  |
| `Get-TopDeskAuthHeaders`          | Baut Basic Auth Header (Base64)                                   |
| `Get-TopDeskCredentialPath`       | Gibt Pfad zur `topdesk.credential` zurueck                       |
| `Save-TopDeskCredential`          | Speichert Credentials via DPAPI (verschluesselt)                  |
| `Get-SavedTopDeskCredential`      | Laedt gespeicherte Credentials aus DPAPI-Datei                    |
| `Remove-TopDeskCredential`        | Loescht gespeicherte Credential-Datei                             |
| `Invoke-TopDeskApi`               | Generischer API-Aufruf mit Fehlerbehandlung                      |
| `Get-TopDeskOperationalSettings`  | Ruft verfuegbare Kategorien/Status ab                             |
| `Import-GuestAuditCsv`            | Importiert und filtert Audit-CSV                                  |
| `New-GuestActivityDescription`    | Generiert Beschreibungstext (Sammel/Einzel)                       |
| `Build-ActivityRequestBody`       | Baut den JSON-Request-Body fuer die API                           |

---

## 7. API-Schnittstelle

### 7.1 Authentifizierung

| Methode          | Details                                                               |
|------------------|-----------------------------------------------------------------------|
| Typ              | HTTP Basic Authentication                                             |
| Benutzername     | TOPdesk Operator Login-Name                                           |
| Passwort         | Application Password (nicht das Login-Passwort)                       |
| Header           | `Authorization: Basic {base64(user:password)}`                        |
| Content-Type     | `application/json; charset=utf-8`                                     |

### 7.2 Verwendete Endpunkte

| Methode | Endpunkt                               | Beschreibung                              |
|---------|----------------------------------------|-------------------------------------------|
| GET     | `/tas/api/version`                     | Verbindungstest                           |
| POST    | `/tas/api/operationalActivities`       | Operative Aktivitaet erstellen            |
| GET     | `/tas/api/operationalActivities/settings` | Verfuegbare Einstellungen abrufen      |

### 7.3 Request-Body (POST)

```json
{
    "briefDescription": "M365 Guest Governance: 15 Konten pruefen",
    "request": "Detaillierte Beschreibung mit Kontenliste...",
    "category": { "name": "IT-Sicherheit" },
    "subcategory": { "name": "Zugriffskontrolle" },
    "operatorGroup": { "groupName": "Security-Team" },
    "operator": { "name": "Max Mustermann" },
    "activityType": { "name": "Pruefung" },
    "plannedStartDate": "2026-03-10T08:00:00Z",
    "plannedEndDate": "2026-03-17T08:00:00Z"
}
```

> Referenz-Felder koennen alternativ per GUID uebergeben werden: `{ "id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" }`

### 7.4 Fehlerbehandlung

| HTTP-Code | Bedeutung                                                        |
|-----------|------------------------------------------------------------------|
| 401       | Authentifizierung fehlgeschlagen (Login/Password pruefen)        |
| 403       | Fehlende API-Berechtigung                                        |
| 400       | Ungueltige Felder im Request-Body (Kategorie existiert nicht)    |
| 404       | Endpunkt nicht gefunden (Operations Management nicht lizenziert) |

---

## 8. Workflow-Integration

### 8.1 Empfohlener Gesamtprozess

```
  M365-GuestGovernance.ps1                TOPdesk-GuestActivities.ps1
  ========================                ============================

  [1] Gastkonto-Audit
       |
       v
  Reports/GuestAudit_*.csv  ──────────>  CSV einlesen
  Reports/GuestAudit_*.html              |
  Reports/GuestAudit_*.json              v
       |                               [1-4] Aktivitaeten erstellen
       v                                     |
  [2] Deaktivieren                           v
       |                               TOPdesk: Operative Aktivitaet
       v                               Reports/TopDesk_Activities_*.json
  [3] Loeschen
       |
       v
  Reports/Cleanup_*.csv
```

### 8.2 Automatisierung (geplanter Task)

Das Skript kann via Dot-Sourcing in einem geplanten Task (Task Scheduler / Azure Automation) eingebunden werden:

```powershell
# Beispiel: Woechentlicher Audit + TOPdesk-Meldung
. .\M365-GuestGovernance.ps1
. .\TOPdesk-GuestActivities.ps1

# 1. Audit durchfuehren
Connect-M365Governance
Get-M365GuestReport

# 2. Kritische Konten an TOPdesk melden
$cred = Get-StoredCredential -Target "TOPdesk-API"  # z.B. aus Windows Credential Manager
Connect-TopDesk -Credential $cred
New-TopDeskGuestActivity -Filter "Kritisch" -Mode "Sammel"
```

---

## 9. Ergebnis-Protokollierung

Jede erfolgreiche Erstellung wird als JSON-Log gespeichert:

| Datei                                  | Inhalt                                               |
|----------------------------------------|------------------------------------------------------|
| `Reports/TopDesk_Activities_{ts}.json` | Zeitstempel, TOPdesk-URL, Modus, Filter, Quell-CSV, erstellte Aktivitaeten |

---

## 10. Dateien und Verzeichnisstruktur

```
GuestGovernance/
├── M365-GuestGovernance.ps1            # Hauptskript (Audit + Cleanup)
├── M365-GuestGovernance.md             # Dokumentation Hauptskript
├── TOPdesk-GuestActivities.ps1         # TOPdesk-Integration (dieses Skript)
├── TOPdesk-GuestActivities.md          # Diese Dokumentation
├── topdesk_config.json                 # Konfiguration (wird beim Start erstellt)
├── topdesk_config.example.json         # Konfigurationsvorlage
├── topdesk.credential                  # DPAPI-verschluesselte Credentials (optional)
├── cleanup_template.csv                # CSV-Vorlage fuer Massenoperationen
├── excluded_domains.txt                # Domain-Whitelist
└── Reports/                            # Ausgabeverzeichnis
    ├── GuestAudit_*.csv                # Input fuer dieses Skript
    └── TopDesk_Activities_*.json       # Output dieses Skripts
```

> **Achtung:** Die Datei `topdesk.credential` darf **nicht** in ein Git-Repository eingecheckt werden. Sie sollte in `.gitignore` aufgenommen werden.

---

## 11. Sicherheitsaspekte

| Aspekt                          | Implementierung                                                |
|---------------------------------|----------------------------------------------------------------|
| Credential-Speicherung          | DPAPI-verschluesselt (`topdesk.credential`), benutzer- und maschinengebunden |
| Credential-Fallback             | Sitzungscache > DPAPI-Datei > Interaktiver Dialog              |
| Ungueltige gespeicherte Creds   | Automatische Loeschung bei HTTP 401, erneute Abfrage           |
| Zugriffskontrolle               | TOPdesk API-Konto mit minimalen Berechtigungen                 |
| WhatIf-Modus                    | Standardmaessig aktiv, Ausfuehrung nur nach Bestaetigung       |
| Verbindungstest                 | `GET /version` vor produktiven API-Aufrufen                    |
| Eingabevalidierung              | GUID-Erkennung, URL-Bereinigung, Laengenbegrenzung (80 Zeichen)|
| Audit-Trail                     | JSON-Log aller erstellten Aktivitaeten                          |
| Keine Secrets in Config         | `topdesk_config.json` enthaelt keine Passwörter                |
| UTF-8 Encoding                  | Request-Body wird als UTF-8 Bytes gesendet                     |

---

## 12. Aenderungshistorie

| Datum      | Aenderung                                                          |
|------------|--------------------------------------------------------------------|
| 2026-03    | Initiale Erstellung: TOPdesk-Integration fuer Guest Governance     |
| 2026-03    | Sichere Credential-Speicherung via DPAPI hinzugefuegt              |
| 2026-03    | Menue-Option [7] fuer Credential-Verwaltung                        |
| 2026-03    | Automatische Invalidierung bei fehlgeschlagener Authentifizierung   |
