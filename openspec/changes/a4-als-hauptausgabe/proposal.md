# Change: A4-Seite als Hauptausgabe, alle Faktoren justierbar

## Warum

Die Web-Auswertung zeigte eine Tabelle, das PDF war ein Nebenprodukt. Gebraucht wird
das Gegenteil: Die A4-Seite für den Jugendamtsleiter ist das Produkt, die Web-Ansicht
dient nur dem Justieren. Außerdem wich die App in der Reserve-Verteilung von der
Steuerungsdatei ab (Excel rechnete unskaliert, die App skalierte auf das Vertragssoll).

## Was

1. `params.py` — jedes Feld, das das Ergebnis beeinflusst, ist ein benanntes Feld und
   kommt aus der XLSX oder aus dem Formular; `to_form()`/`from_form()` sind verlustfrei.
2. `core.py` — Rechnung folgt den Excel-Formeln einschließlich k5 und kr; Invarianten
   (Summe AF 1 = Soll AF 1, Summe AF 5 = Soll AF 5, Reserve = Kapazität − Soll).
3. `report.py` — A4-Seite je Jahr, dynamisch kompaktierend, Herleitung im Fuß.
4. `xlsx_write.py` — Rückweg in die Steuerungsdatei; Auswertungsformeln werden für die
   aktuelle Zeilenzahl neu gesetzt, damit Excel und App gleich rechnen.
5. `templates/index.html` — Justierformular mit PDF-Vorschau im `iframe`; „In Excel
   sichern" mit sichtbarem Ergebnis.

## Abnahme

- `pytest` mit `STEUERDATEI`: 22 Prüfungen, darunter der unabhängige Abgleich der
  geschriebenen Mappe über `formulas` je Person und Jahr (< 0,01 h Abweichung).
- PDF gerendert und als Bild geprüft (keine Überlappung, 2 Seiten, 15 Personen passen).
- Live: `/healthz` `data_ok: true`, `/export.pdf` mit geänderter Zahl im Text.
- „In Excel sichern" gegen die Nextcloud, mit Sicherungsdatei und Rückprobe.
