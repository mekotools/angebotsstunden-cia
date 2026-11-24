# Angebotsstunden-Auswertung (CIA Spandau)

Die druckbare **A4-Seite ist das Produkt**. Davor liegt ein Justierformular: jeder
Faktor und jede Zahl lässt sich einzeln ändern, die Vorschau ist das echte PDF.
Datenquelle und Speicher ist ausschließlich die Steuerungsdatei (XLSX) in der
Nextcloud — die App hält keine eigenen Daten.

## Ablauf

```
Nextcloud-XLSX (Annahmen, Eingabe 2026/2027, Texte)
        │  WebDAV
        ▼
params.py  ──►  core.py  ──►  report.py (fpdf2, A4)
        ▲                         │
        │  /sichern               ▼
xlsx_write.py  ◄────────  PDF / Vorschau im iframe
```

* `params.py` — Datenmodell, Lesen aus der XLSX, Lesen des Formulars, `to_form()`
* `core.py` — Rechnung mit den Faktoren k5 (AF-5-Verteilung) und kr (Reserve)
* `report.py` — eine A4-Seite je Jahr, Tabelle/Deckungsbalken/Personenbalken,
  passt sich der Zeilenzahl an (13–15 Personen bleiben auf einer Seite)
* `xlsx_write.py` — Rückweg: Eingabeblätter schreiben, Auswertungsformeln für die
  aktuelle Zeilenzahl neu setzen, Blatt `Quellen` unangetastet lassen

## Was justierbar ist

Global: Werktage, Feiertage, Urlaub, Krankheit (oder Arbeitstage direkt, per Häkchen),
Leistungsstunden gesamt/AF 1/AF 5, Kostensatz, Titel, Organisation, Träger.
Je Jahr: Honorarstunden, Workshoptage, Kennzeichen (IST/PLANUNG), Untertitel, Hinweis.
Je Person: Name, Funktion, Wochenstunden, von/bis-Monat, AF-5-Anteil, Reserve-Anteil —
Zeilen lassen sich hinzufügen und löschen.

## Endpunkte

- `/` — Justierformular mit PDF-Vorschau im iframe
- `/export.pdf` — A4-Ausgabe (`?y=2026`, `?y=2027` oder beide; `&dl=1` als Download)
- `/kennzahlen` — Kontrollzahlen und Warnungen als HTML-Fragment
- `/sichern` — POST: Werte in die Steuerungsdatei zurückschreiben (legt vorher
  `…sicherung-<stempel>.xlsx` in der Nextcloud an)
- `/export.xlsx` — die Steuerungsdatei selbst
- `/healthz` — Liveness plus Datenstand (Bytes, Jahre, Personenzahl)

## Umgebung

| Variable | Bedeutung |
|---|---|
| `NC_URL` | Nextcloud-Basis, z. B. `https://cloud.example.org` |
| `NC_USER` / `NC_PASS` | WebDAV-Zugang (App-Passwort) |
| `NC_PATH` | Pfad der Steuerungsdatei im Benutzerordner |
| `LOCAL_FILE` | optional: lokale XLSX statt Nextcloud (Entwicklung/Tests) |
| `CACHE_SECONDS` | Cache-Dauer der Datei (Standard 300) |

## Steuerungsdatei

Blätter `Annahmen` (Krankheitstage, Arbeitstage, Sollwerte, Kostensatz, Werktage),
`Eingabe <Jahr>` (Personen, Wochenstunden, Zeitraum, AF-5-Stunden, Reserve-Anteil),
`Auswertung <Jahr>` (Formeln inklusive der Faktoren k5 und kr) sowie `Quellen` und
`Texte`. Leere Zellen werden als 0 gewertet.

## Rechenweg

1. `Stunden im Jahr = Wochenstunden / 5 × Arbeitstage × Anteil am Jahr`
2. `AF 5 = AF-5-Anteil × k5`, `k5 = (Soll AF 5 − Honorarstunden) / Summe der AF-5-Anteile`
3. `Reserve = Stunden × Reserve-Anteil × kr` (Anteil 1 = ganze Stelle nur Leitung/Verwaltung),
   `kr = (Kapazität − Vertragssoll − feste Reserve) / Summe der übrigen Ansprüche`
4. `AF 1 = Stunden − AF 5 − Reserve`

Daraus folgen die Invarianten: Summe AF 5 = Soll AF 5, Summe AF 1 = Soll AF 1,
Summe Reserve = Kapazität − Vertragssoll. Die XLSX rechnet mit **denselben** Formeln;
`tests/test_kette.py` prüft das je Person über die Bibliothek `formulas`.

## Tests

```bash
python -m venv .venv && .venv/bin/pip install -r requirements-dev.txt
STEUERDATEI=/pfad/zur/Steuerungsdatei.xlsx .venv/bin/python -m pytest -q tests
```

Die Steuerungsdatei ist absichtlich **nicht** im Repository (echte Personendaten);
ohne `STEUERDATEI` überspringen sich die dateibasierten Prüfungen.

## Betrieb

Coolify-App auf flip, Build aus diesem Repo (`master`), BasicAuth im Caddy-Label.
Zugangsdaten und Nextcloud-Pfad kommen aus dem Container-Env, nie aus dem Repo.
