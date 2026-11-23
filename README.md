# Angebotsstunden CIA Spandau

Minimale Web-App: liest die Steuerungsdatei `Angebotsstunden-CIA-Modell.xlsx` aus der
Nextcloud (WebDAV), rechnet das Stundengerüst und exportiert PDF und Excel.

## Umgebung

| Variable | Bedeutung |
|---|---|
| `NC_URL` | Nextcloud-Basis, z. B. `https://nc.cia-spandau.org` |
| `NC_USER` / `NC_PASS` | WebDAV-Zugang (App-Passwort) |
| `NC_PATH` | Pfad der Steuerungsdatei, z. B. `1 Team CIA/Finanzierung CIA/Angebotsstunden-CIA-Modell.xlsx` |
| `CACHE_SECONDS` | Cache-Dauer der Datei, Standard 300 |
| `LOCAL_FILE` | optional: lokale XLSX statt Nextcloud (nur Entwicklung) |

## Endpunkte

- `/` — Auswertung mit Jahresumschalter
- `/export.pdf` — zweistufiges PDF (2026/2027)
- `/export.xlsx` — die Steuerungsdatei selbst
- `/healthz` — Status inkl. Datenstand

## Rechenweg

Stunden im Jahr = Wochenstunden / 5 × Arbeitstage (aus Blatt `Annahmen`).
AF 5 zuerst aus den Honorarstunden, Rest aus dem Team; Reserve nach Rollenanteil,
AF 1 als Rest. Die Summen treffen das Vertragssoll (5.000 + 1.828 h).
