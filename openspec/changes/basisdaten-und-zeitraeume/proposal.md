# Basisdaten je Jahr und Zeiträume je Person

## Warum

Die Steuerungsdatei führte Werktage (261) und Feiertage (8) als **einen** Wert für
beide Jahre, und die Personen hatten nur einen Zeitraum (Monat von … bis). Zwei
Lücken:

1. **Basisdaten je Jahr.** Die Zahl der Arbeitstage hängt vom Jahr und vom
   Bundesland ab — 2026 hat Berlin 254 Arbeitstage, 2027 255, 2028 251 bei
   260 Werktagen. Ein Handwert gilt höchstens für ein Jahr.
2. **Änderung mitten im Jahr.** Wer im Juli seine Wochenstunden erhöht, ließ sich
   nicht abbilden; die Jahresstunden waren eine einzige Multiplikation.

## Was sich ändert

* Neues Modul `basisdaten.py`: Feiertage je Bundesland aus dem Landesrecht
  (bewegliche Feste über die Osterformel, Einmal-Feiertage als Liste), Werktage
  Mo–Fr, Arbeitstage = Werktage − Feiertage auf Werktagen; Monatstabelle und
  Zeitraumsumme. Bundesland steht im Blatt `Quellen`.
* `params.py`: `Abschnitt` (Monat von … bis, Wochenstunden, Urlaub, Krankheit,
  AF-5-Anteil, Reserve-Anteil); `Person` hält eine Liste von Abschnitten.
  Leere Urlaubs-/Krankheitsfelder erben anteilig den Jahresanspruch.
* `core.py`: Arbeitstage je Abschnitt monatsgenau, Stunden je Abschnitt,
  Personenwerte sind die Summe ihrer Abschnitte. Die Verteilung (k5, kr) bleibt
  unverändert und trifft weiterhin die Vertragssummen exakt.
* `xlsx_write.py`: neues Blatt `Basisdaten <Jahr>` (Monatstabelle + Feiertagsliste),
  `Eingabe <Jahr>` mit **einer Zeile je Zeitraum** und den Spalten Arbeitstage und
  Stunden als Formeln, `Auswertung <Jahr>` mit einer Zeile je Person und
  Formeln, die die Zeitraumzeilen summieren.
* `report.py`: A4-Seite zeigt je Person die Zeiträume („20 → 35“, „01.–06. ·
  07.–12.“) und in der Fußzeile die Basisdaten des Jahres.
* Formular: Bundesland-Auswahl, Jahresanspruch Urlaub/Krankheit, Tabelle
  Basisdaten je Jahr, Personentabelle mit Urlaub/Krankheit je Zeitraum und
  „+ Zeitraum“.

## Folgen für die Zahlen

Die Ankerwerte der abgenommenen Seite verschieben sich bewusst, weil jetzt
gerechnet statt geschätzt wird:

* Kapazität 2026 7.915 → 7.950 h, 2027 7.967 → 8.029 h (je ein Feiertag weniger
  auf Werktagen als der Handwert 8).
* Unterjährige Personen (z. B. Januar–März) bekommen die Arbeitstage ihrer Monate
  statt eines linearen Zwölftels.

Die Deckung bleibt 100 %, weil AF 1, AF 5 und Reserve weiterhin auf das
Vertragssoll skaliert werden.

## Abnahme

* `tests/test_basisdaten.py` prüft die Feiertagszahlen 2027 für alle 16 Länder und
  die Arbeitstage 2025–2027 gegen veröffentlichte Tabellen.
* `tests/test_kette.py` prüft Zeitraumteilung, Vererbung von Urlaub/Krankheit,
  den Excel-Rückweg (zwei Zeiträume überleben das Sichern) und weiterhin die
  Excel-Gleichheit über die `formulas`-Auswertung.

## Nicht enthalten

* Urlaub/Krankheit als eigene Abwesenheitsliste je Datum (nur Tage je Zeitraum).
* Automatische Prüfung, ob die Zeiträume einer Person lückenlos sind.
