# Heimbas ICS Kalender

Automatischer ICS-Kalender-Generator für Heimbas Einsatz-Vorschau.

## Features
- Multi-User-Unterstützung via USERS_JSON
- Automatische stündliche Updates
- GitHub Pages Hosting
- iPhone-kompatible Kalenderabos

## URLs
- Kombiniert: https://bonii97.github.io/Heimbas-Calender/index.ics
- Pro User: https://bonii97.github.io/Heimbas-Calender/dienstplan_<name>.ics

## Setup
1. Repository Secret USERS_JSON anlegen (als JSON-Array)
2. Workflow manuell starten
3. ICS-URLs am iPhone abonnieren

### Beispiel für USERS_JSON

Folgende Key-Varianten werden unterstützt (`name`/`label`, `user`/`username`, `pass`/`password`).
Verwende keine Klarnamen – nur neutrale Platzhalter:

```json
[
  { "label": "user1", "user": "USER1", "pass": "PASS1" },
  { "name":  "user2", "username": "USER2", "password": "PASS2" }
]
```

---

## Wie der Heimbas‑Kalender funktioniert – und warum

### Überblick
Dieser Workflow meldet sich automatisiert im Heimbas‑Portal an, ruft die Seite „Einsatz‑Vorschau“ auf, liest die Tabelle mit Terminen aus und erzeugt daraus eine abonnierbare Kalenderdatei (ICS). Anschließend werden die Dateien über GitHub Pages bereitgestellt. So lassen sich Einsätze als „Kalenderabo“ auf dem iPhone (und anderen Clients) einbinden und halten sich automatisch aktuell.

### Warum dieser Ansatz?
- Heimbas bietet (nach aktuellem Kenntnisstand) keine offizielle ICS‑Schnittstelle für die Einsatz‑Vorschau.
- ICS ist das Standardformat für Kalenderabos auf iOS, macOS, Outlook, Google Calendar etc.
- GitHub Actions + Pages liefern einen stabilen, kostenlosen und wartungsarmen Hosting‑ und Automatisierungs‑Pfad für reine Lese‑Workflows.

### Technischer Ablauf (high‑level)
1) Login & Navigation (Playwright/Chromium)
   - Startet headless Chromium
   - Füllt Benutzer/Passwort (aus `USERS_JSON`) robust über verschiedene Selektoren (auch alte BBj/GWT‑Oberflächen)
   - Erkennt Login‑Erfolg über „intelligentes Polling“ (URL/Content/Tabellenindikatoren)
   - Klickt auf den Menüpunkt „Einsatz‑Vorschau“ (table‑basierte Menüs werden unterstützt)

2) Tabellen‑Parsing (BeautifulSoup)
   - Sucht eine Tabelle mit typischen Headern wie `Datum`, `Uhrzeit`, `Einsatz/Training`, `Dauer`
   - Extrahiert pro Zeile:
     - Datum (dd.mm.yyyy oder dd.mm.yy)
     - Start/Ende (Formate: `08:00 - 10:00` bzw. `von … bis …`)
     - Beschreibung/Adresse (aus dem längsten/geeignetsten Zellen‑Text)
     - Dauer (falls Spalte vorhanden), Formate wie `2,0`/`2.0` Stunden oder `90 Minuten`

3) ICS‑Erzeugung (icalendar)
   - Zeitzone: `Europe/Berlin`
   - `summary` = erste Zeile/erster Satz der Beschreibung
   - `location` = erkannte Adresse (falls vorhanden)
   - `description` = kompletter Zellen‑Text
   - UID stabil: Hash aus Start/Ende/Ort/Text → verhindert Duplikate
   - Falls „Dauer“ erkannt wird, überschreibt sie die Endzeit (Start + Dauer hat Priorität)

4) Veröffentlichung (GitHub Pages)
   - Workflow lädt die ICS‑Dateien als Pages‑Artifact hoch
   - `site/index.ics` = kombinierte Datei für alle Benutzer
   - `site/dienstplan_<name>.ics` = je Benutzer
   - `site/index.html` listet die Dateien für den Browser
   - `robots.txt` + `<meta name="robots" …>` verhindern Suchmaschinen‑Indexierung

### Multi‑User‑Logik
- `USERS_JSON` kann 1..n Accounts enthalten.
- Für jeden Account werden die Einsätze separat gescraped und in `dienstplan_<name>.ics` gespeichert.
- Zusätzlich wird eine kombinierte `index.ics` erzeugt (Merge aller Termine).
- `<name>` ist der Slug aus `name`/`label` (Kleinbuchstaben, Sonderzeichen entfernt).

### Kalender auf dem iPhone abonnieren
1. Einstellungen → Kalender → Accounts → Account hinzufügen → Andere → „Kalenderabo hinzufügen“
2. URL einfügen, z. B. `https://bonii97.github.io/Heimbas-Calender/index.ics`
3. „Sichern“ – fertig. Das Abo aktualisiert sich automatisch (Workflow läuft stündlich).

### Datenschutz & Sicherheit
- Das Repository ist öffentlich, die ICS‑Dateien auf Pages damit prinzipiell abrufbar (keine Authentifizierung).
- Suchmaschinen‑Indexierung ist deaktiviert (robots.txt + Meta‑Tag). Öffentliche Verlinkungen können die Datei dennoch auffindbar machen.
- Zugangsdaten gehören ausschließlich in Secrets (`USERS_JSON`) – niemals in den Code/Commits.
- Logs sind so ausgelegt, keine sensitiven Inhalte auszugeben. Bei Bedarf kann das Debug‑Level weiter reduziert werden.

### Troubleshooting
- Login/Tabelle nicht gefunden:
  - Prüfe Actions‑Logs (Schritt „Run scraper“). Oft sind Selektoren leicht anzupassen (Portal‑Layout).
  - Ggf. einen Screenshot/HTML‑Schnipsel posten – Selektoren kann man schnell nachschärfen.
- Falsche Dauer:
  - Die Spalte „Dauer“ muss erkennbar sein (Header enthält „Dauer“). Unterstützte Formate: `2,0`/`2.0`‑Stunden, `90 Minuten`.
  - Endzeit wird aus Start + Dauer berechnet, sofern Dauer erkannt wurde.
- Pro‑User‑Datei fehlt:
  - Sicherstellen, dass `name`/`label` im `USERS_JSON` gesetzt sind. Der daraus gebildete Slug bestimmt den Dateinamen.
- Deploy schlägt fehl:
  - Settings → Pages → Source = GitHub Actions
  - Actions → letzter Run → Job „deploy“ → Fehlermeldung prüfen

### Anpassungen / Erweiterungen
- Selektoren (Login, Navigation, Tabelle) befinden sich in `scraper.py` und sind bewusst robust/erweiterbar.
- Cron‑Intervall in `.github/workflows/build.yml` (`0 * * * *`) → stündlicher Lauf.
- Weitere Kalender‑Client‑Kompatibilitäten (Outlook, Google Calendar) sind durch ICS automatisch gegeben.


