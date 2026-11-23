# Angebotsstunden-Auswertung

Kleine Web-App, die eine Steuerungsdatei (XLSX) aus einer Nextcloud liest, die
Jahresstunden je Mitarbeiter*in rechnet und Tabelle, Balken, PDF und Excel ausgibt.

Die App enthält keine Daten: Personen, Wochenstunden, Honorartage und Sollwerte stehen
in der XLSX, die per WebDAV geladen wird.

## Umgebung

| Variable | Bedeutung |
|---|---|
| `NC_URL` | Nextcloud-Basis, z. B. `https://cloud.example.org` |
| `NC_USER` / `NC_PASS` | WebDAV-Zugang (App-Passwort) |
| `NC_PATH` | Pfad der Steuerungsdatei im Benutzerordner |
| `APP_TITLE` | Überschrift (Standard `Angebotsstunden`) |
| `ORG` | Organisation in Kopfzeile und Fußtext |
| `CACHE_SECONDS` | Cache-Dauer der Datei (Standard 300) |
| `LOCAL_FILE` | optional: lokale XLSX statt Nextcloud (Entwicklung) |

## Steuerungsdatei

Blätter `Annahmen` (Arbeitstage, Sollwerte), `Eingabe <Jahr>` (Personen, Wochenstunden,
Zeitraum, AF-5-Stunden, Reserve-Anteil) und `Quellen` (maschinell gelesene Werte).
Fehlende oder leere Zellen werden als 0 gewertet.

## Endpunkte

- `/` — Auswertung mit Jahresumschalter
- `/export.pdf` — PDF mit Tabelle, Kennzahlen und Balken
- `/export.xlsx` — die Steuerungsdatei selbst
- `/healthz` — Status mit Datenstand

## Rechenweg

`Stunden im Jahr = Wochenstunden / 5 × Arbeitstage`. AF 5 wird zuerst aus den
Honorarstunden gedeckt, der Rest aus dem Team; die Reserve folgt dem Rollenanteil,
AF 1 ist der Rest. Die Summen treffen damit die Sollwerte aus dem Blatt `Annahmen`.
