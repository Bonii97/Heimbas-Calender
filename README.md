# Heimbas ICS Kalender

Automatischer ICS-Kalender-Generator für Heimbas Einsatz-Vorschau.

## Features
- Multi-User-Unterstützung via USERS_JSON
- Automatische stündliche Updates
- GitHub Pages Hosting
- iPhone-kompatible Kalenderabos

## URLs
- Kombiniert: https://bonii97.github.io/Heimbas-Trustful-DayCare-GmbH-Calender/index.ics
- Pro User: https://bonii97.github.io/Heimbas-Trustful-DayCare-GmbH-Calender/dienstplan_<name>.ics

## Setup
1. Repository Secret USERS_JSON anlegen (als JSON-Array)
2. Workflow manuell starten
3. ICS-URLs am iPhone abonnieren

### Beispiel für USERS_JSON

Folgende Key-Varianten werden unterstützt (`name`/`label`, `user`/`username`, `pass`/`password`):

```json
[
  { "label": "vici",  "user": "V.Hendel", "pass": "<PASS>" },
  { "name":  "felix", "username": "USER2", "password": "PASS2" }
]
```

