# Angebotsstunden-Auswertung (CIA Spandau)

Die druckbare **A4-Seite ist das Produkt**. Davor liegt ein Justierformular: jeder
Faktor und jede Zahl lässt sich einzeln ändern, die Vorschau ist das echte PDF.
Datenquelle und Speicher ist ausschließlich die Steuerungsdatei (XLSX) in der
Nextcloud — die App hält keine eigenen Daten.

## Ablauf

```
Nextcloud-XLSX (Basisdaten, Annahmen, Eingabe 2026/2027, Texte)
        │  WebDAV
        ▼
params.py ──► basisdaten.py ──► core.py ──► report.py (fpdf2, A4)
        ▲                                      │
        │  /sichern                            ▼
xlsx_write.py  ◄──────────────────  PDF / Vorschau im iframe
```

* `basisdaten.py` — Werktage, Feiertage und Arbeitstage je **Bundesland und Jahr**,
  aus dem Kalender gerechnet (Osterformel, Landesrecht, Einmal-Feiertage als Liste)
* `params.py` — Datenmodell (Personen mit Zeiträumen), Lesen aus der XLSX, Lesen des
  Formulars, `to_form()`
* `core.py` — Rechnung mit den Faktoren k5 (AF-5-Verteilung) und kr (Reserve)
* `report.py` — eine A4-Seite je Jahr, Tabelle/Deckungsbalken/Personenbalken,
  passt sich der Zeilenzahl an (13–15 Personen bleiben auf einer Seite)
* `xlsx_write.py` — Rückweg: Eingabeblätter schreiben, Auswertungsformeln für die
  aktuelle Zeilenzahl neu setzen, Blatt `Quellen` unangetastet lassen

## Basisdaten je Jahr

Werktage, Feiertage und Arbeitstage werden **nicht gepflegt, sondern gerechnet**:
`basisdaten.py` kennt die gesetzlichen Feiertage aller 16 Länder (bewegliche Feste
über die Osterformel, Einmal-Feiertage als Liste, z. B. 8. Mai 2025 in Berlin),
zählt die Werktage Montag–Freitag und zieht die Feiertage ab. Das Bundesland steht
im Blatt `Quellen` und im Formular.

Beispiele Berlin: 2026 = 261 Werktage − 7 Feiertage auf Werktagen = **254
Arbeitstage**, 2027 = **255**, 2028 = 260 − 9 = **251**. Die Werte sind gegen
veröffentlichte Tabellen geprüft (`tests/test_basisdaten.py`, alle 16 Länder).

## Personen mit Zeiträumen

Eine Person hat **einen oder mehrere Zeiträume** (Monat von … bis), jeder mit
eigenen Wochenstunden, Urlaubs- und Krankheitstagen — damit lässt sich abbilden,
dass jemand ab Juli seine Stunden erhöht. In der Personentabelle ist das „gleicher
Name in Folgezeilen"; der Knopf **＋** teilt einen Zeitraum in der Mitte und legt
den zweiten an. Leere Urlaubs-/Krankheitsfelder erben anteilig den Jahresanspruch
aus den Annahmen. AF-5- und Reserve-Anteil gelten je Zeitraum.

## Was justierbar ist

Basisdaten: Bundesland, Urlaubsanspruch und Krankheitstage je Jahr.
Vertrag: Leistungsstunden gesamt/AF 1/AF 5, Kostensatz, Titel, Organisation, Träger.
Je Jahr: Honorarstunden, Workshoptage, Kennzeichen (IST/PLANUNG), Untertitel, Hinweis.
Je Zeitraum: Name, Funktion, von/bis-Monat, Wochenstunden, Urlaub, Krankheit,
AF-5-Anteil, Reserve-Anteil — Zeilen lassen sich hinzufügen, teilen und löschen.

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

`Basisdaten <Jahr>` (Monatstabelle Werktage/Feiertage/Arbeitstage plus Feiertagsliste),
`Annahmen` (Krankheitstage, Sollwerte, Kostensatz, Bundesland, Urlaubsanspruch),
`Eingabe <Jahr>` (**eine Zeile je Zeitraum**; Arbeitstage und Stunden sind Formeln),
`Auswertung <Jahr>` (eine Zeile je Person, Formeln inklusive k5 und kr) sowie
`Quellen` und `Texte`. Leere Zellen werden als 0 gewertet; beim Sichern werden
anteilige Urlaubs-/Krankheitswerte als Zahl geschrieben.

## Rechenweg

1. `Arbeitstage eines Zeitraums = Werktage − Feiertage` (Blatt `Basisdaten`),
   `Netto = Arbeitstage − Urlaub − Krankheit`, `Stunden = Wochenstunden / 5 × Netto`;
   die Stunden einer Person sind die Summe ihrer Zeiträume
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

Coolify-App auf flip, Build aus diesem Repo (`master`).
Zugangsdaten und Nextcloud-Pfad kommen aus dem Container-Env, nie aus dem Repo.

**Zugang:** über CIA-Authentik per Forward-Auth (Proxy-Anbieter „Angebotsstunden",
eigener Outpost-Container `angebotsstunden-outpost` auf flip). Der Router steht in
`/data/coolify/proxy/dynamic/angebotsstunden-auth.yml` auf flip und hat bewusst
höhere Priorität als der Coolify-Router; die BasicAuth-Kennzeichnung in Coolify
bleibt als Rückfallebene gesetzt, wird aber nicht mehr wirksam.
