"""Schreibt die Parameter in die Steuerungsdatei zurück.

Die XLSX bleibt die Datenquelle: Blätter ``Annahmen`` und ``Eingabe <Jahr>``
tragen die Eingaben, ``Auswertung <Jahr>`` rechnet sie mit **denselben** Formeln
wie die App (Faktoren k5 und kr als eigene Zellen), ``Texte`` hält Überschriften.
Blatt ``Quellen`` bleibt inhaltlich erhalten.

So zeigen App und Excel nach dem Sichern dieselben Zahlen — auch wenn Zeilen
hinzukommen oder wegfallen, weil die Formeln für die neue Zeilenzahl neu gesetzt
werden.
"""
from __future__ import annotations

from io import BytesIO

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from params import Params, STANDARD_TEXTE

PETROL = "FF1E4B5A"
GELB = "FFFFF2CC"
GRAU = "FFF2F5F6"

JAHRES_NAMEN = {2026: "", 2027: ""}


def _kopf(ws, spalten: list[str], zeile: int = 3):
    for i, titel in enumerate(spalten, start=1):
        c = ws.cell(row=zeile, column=i, value=titel)
        c.font = Font(bold=True, color="FFFFFFFF", size=9)
        c.fill = PatternFill("solid", fgColor=PETROL)
        c.alignment = Alignment(horizontal="left" if i == 1 else "center", vertical="center")


def _wert(ws, zeile: int, spalte: int, wert, gelb: bool = True):
    c = ws.cell(row=zeile, column=spalte, value=wert)
    if gelb:
        c.fill = PatternFill("solid", fgColor=GELB)
    return c


def _leeren(ws, ab_zeile: int, bis_spalte: int = 8):
    for r in range(ab_zeile, ws.max_row + 1):
        for c in range(1, bis_spalte + 1):
            ws.cell(row=r, column=c).value = None


def _herkunft_alt(ws) -> dict[str, str]:
    d = {}
    r = 4
    while ws.cell(row=r, column=1).value:
        d[str(ws.cell(row=r, column=1).value).strip()] = str(ws.cell(row=r, column=8).value or "")
        r += 1
    return d


def _annahmen(wb, p: Params):
    ws = wb["Annahmen"]
    werte = [
        ("Krankheitstage je Jahr", p.krank, "Annahme; aus den Nachweisen nicht vollständig ablesbar"),
        ("Arbeitstage", p.arbeitstage, "Werktage − Feiertage − Urlaub − Krankheit" if not p.arbeitstage_fix
         else "von Hand gesetzt"),
        ("Produktive Wochen", round(p.arbeitstage / 5, 2), "Arbeitstage / 5"),
        ("Leistungsstunden gesamt", p.soll, "Vertrag § 5 Nr. 1"),
        ("davon AF 1", p.soll_af1, ""),
        ("davon AF 5", p.soll_af5, ""),
        ("Kostensatz (€/h)", p.satz, "Leistungsvertrag § 5 Nr. 1"),
        ("Werktage je Jahr", p.werktage, "Kalender Mo–Fr (Blatt Quellen)"),
        ("Feiertage auf Werktagen", p.feiertage, "Blatt Quellen"),
        ("Urlaubsanspruch (Tage)", p.urlaub, "Blatt Quellen"),
    ]
    for i, (label, wert, grund) in enumerate(werte, start=4):
        a = ws.cell(row=i, column=1, value=label)
        a.font = Font(size=10)
        _wert(ws, i, 2, wert)
        c = ws.cell(row=i, column=3, value=grund)
        c.font = Font(size=9, color="FF6B7C85")
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 52


def _eingabe(wb, p: Params, jahr: int):
    ws = wb[f"Eingabe {jahr}"]
    alt = _herkunft_alt(ws)
    _leeren(ws, 1)
    titel = ws.cell(row=1, column=1, value=f"Eingabe {jahr} (gelb = änderbar)")
    titel.font = Font(bold=True, size=11)
    _kopf(ws, ["Mitarbeiter*in", "Funktion", "Wochenstunden", "von Monat", "bis Monat",
               "AF 5 (h)", "Reserve-Anteil", "Herkunft"])
    for i, q in enumerate(p.jahre[jahr].personen):
        r = 4 + i
        a = ws.cell(row=r, column=1, value=q.name)
        a.font = Font(bold=True, size=10)
        ws.cell(row=r, column=2, value=q.rolle).font = Font(size=10)
        _wert(ws, r, 3, q.wochenstunden)
        _wert(ws, r, 4, q.von)
        _wert(ws, r, 5, q.bis)
        _wert(ws, r, 6, q.af5)
        _wert(ws, r, 7, q.res)
        c = ws.cell(row=r, column=8, value=alt.get(q.name, "im Formular gepflegt"))
        c.font = Font(size=9, color="FF6B7C85")
        for sp in (3, 6, 7):
            ws.cell(row=r, column=sp).number_format = "#,##0.0" if sp == 3 else "#,##0.0##"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 30
    for sp, breite in (("C", 12), ("D", 9), ("E", 9), ("F", 9), ("G", 12), ("H", 46)):
        ws.column_dimensions[sp].width = breite
    ws.freeze_panes = "A4"


def _auswertung(wb, p: Params, jahr: int):
    """Tabelle plus Summenblock. Die Faktoren k5/kr stehen als Zellen darunter,
    damit Excel und App nachweislich gleich rechnen."""
    ws = wb[f"Auswertung {jahr}"]
    _leeren(ws, 1, 8)
    n = len(p.jahre[jahr].personen)
    erste, letzte = 5, 4 + n

    ws.cell(row=1, column=1, value=f"Auswertung {jahr}").font = Font(bold=True, size=11)
    ws.cell(row=2, column=1,
            value="Rechnet ausschließlich aus den Blättern Eingabe und Annahmen.").font = Font(
        size=9, color="FF6B7C85")
    _kopf(ws, ["Mitarbeiter*in", "Stunden im Jahr", "AF 1", "AF 5", "Reserve"], zeile=4)

    r_sum = letzte + 1
    r_hon = r_sum + 1
    r_af5 = r_hon + 1
    r_kap = r_af5 + 1
    r_soll = r_kap + 1
    r_deck = r_soll + 1
    r_res = r_deck + 1
    r_k5 = r_res + 2
    r_kapz = r_k5 + 1
    r_fix = r_kapz + 1
    r_var = r_fix + 1
    r_kr = r_var + 1

    q = f"'Eingabe {jahr}'"
    q_erste, q_letzte = 4, 3 + n   # Eingabe hat den Kopf in Zeile 3, Daten ab Zeile 4
    for i in range(n):
        r = erste + i      # Zeile im Blatt Auswertung
        qr = 4 + i         # zugehörige Zeile im Blatt Eingabe
        ws.cell(row=r, column=1, value=f"={q}!A{qr}")
        ws.cell(row=r, column=2, value=f"={q}!C{qr}/5*Annahmen!$B$5*({q}!E{qr}-{q}!D{qr}+1)/12")
        ws.cell(row=r, column=3, value=f"=B{r}-D{r}-E{r}")
        ws.cell(row=r, column=4, value=f"={q}!F{qr}*$B${r_k5}")
        ws.cell(row=r, column=5,
                value=f"=IF({q}!G{qr}>=0.999,B{r},B{r}*{q}!G{qr}*$B${r_kr})")
    ws.cell(row=r_sum, column=1, value="Summe Team").font = Font(bold=True)
    for sp in "BCDE":
        ws[f"{sp}{r_sum}"] = f"=SUM({sp}{erste}:{sp}{letzte})"
        ws[f"{sp}{r_sum}"].font = Font(bold=True)

    ws.cell(row=r_hon, column=1, value="Honorarkräfte")
    ws.cell(row=r_hon, column=4, value=p.jahre[jahr].honorarstunden)
    ws.cell(row=r_af5, column=1, value="Summe AF 5").font = Font(bold=True)
    ws.cell(row=r_af5, column=4, value=f"=D{r_sum}+D{r_hon}")
    ws.cell(row=r_kap, column=1, value="Kapazität").font = Font(bold=True)
    ws.cell(row=r_kap, column=2, value=f"=B{r_sum}+D{r_hon}")
    ws.cell(row=r_soll, column=1, value="Vertragssoll").font = Font(bold=True)
    ws.cell(row=r_soll, column=2, value="=Annahmen!$B$7")
    ws.cell(row=r_deck, column=1, value="Deckung").font = Font(bold=True)
    ws.cell(row=r_deck, column=2, value=f"=(C{r_sum}+D{r_af5})/B{r_soll}")
    ws.cell(row=r_deck, column=2).number_format = "0.0%"
    ws.cell(row=r_res, column=1, value="Reserve = Kapazität − Soll").font = Font(size=9)
    ws.cell(row=r_res, column=2, value=f"=B{r_kap}-B{r_soll}")

    # Faktoren (offen dokumentiert, damit nachvollziehbar bleibt, wie verteilt wird)
    ws.cell(row=r_k5, column=1, value="Faktor AF 5 (k5)").font = Font(size=9, italic=True)
    ws.cell(row=r_k5, column=2,
            value=f"=(Annahmen!$B$9-D{r_hon})/SUM({q}!$F${q_erste}:$F${q_letzte})")
    ws.cell(row=r_kapz, column=1, value="Reserve-Soll (Kapazität − Soll)").font = Font(size=9, italic=True)
    ws.cell(row=r_kapz, column=2, value=f"=B{r_kap}-Annahmen!$B$7")
    ws.cell(row=r_fix, column=1, value="davon feste Reserve (Anteil 1)").font = Font(size=9, italic=True)
    ws.cell(row=r_fix, column=2,
            value=f'=SUMIF({q}!$G${q_erste}:$G${q_letzte},">=0.999",$B${erste}:$B${letzte})')
    ws.cell(row=r_var, column=1, value="übrige Reserve-Ansprüche (Stunden × Anteil)").font = Font(
        size=9, italic=True)
    ws.cell(row=r_var, column=2,
            value=f"=SUMPRODUCT(({q}!$G${q_erste}:$G${q_letzte}<0.999)*$B${erste}:$B${letzte}"
                  f"*{q}!$G${q_erste}:$G${q_letzte})")
    ws.cell(row=r_kr, column=1, value="Faktor Reserve (kr)").font = Font(size=9, italic=True)
    ws.cell(row=r_kr, column=2, value=f"=(B{r_kapz}-B{r_fix})/B{r_var}")

    for sp, breite in (("A", 30), ("B", 18), ("C", 14), ("D", 14), ("E", 14)):
        ws.column_dimensions[sp].width = breite


def _texte(wb, p: Params):
    if "Texte" in wb.sheetnames:
        ws = wb["Texte"]
        _leeren(ws, 1, 4)
    else:
        ws = wb.create_sheet("Texte")
    ws.cell(row=1, column=1, value="Texte und Kopfzeilen (gelb = änderbar)").font = Font(bold=True, size=11)
    _kopf(ws, ["Schlüssel", "Text"], zeile=3)
    zeilen = [
        ("titel", p.titel),
        ("org", p.org),
        ("traeger", p.traeger),
    ]
    for jahr, j in sorted(p.jahre.items()):
        zeilen += [(f"marker {jahr}", j.marker), (f"untertitel {jahr}", j.untertitel),
                   (f"hinweis {jahr}", j.hinweis)]
    for i, (k, v) in enumerate(zeilen, start=4):
        ws.cell(row=i, column=1, value=k).font = Font(size=10)
        _wert(ws, i, 2, v)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 92


def _quellen(wb, p: Params):
    if "Quellen" not in wb.sheetnames:
        return
    ws = wb["Quellen"]
    for r in range(1, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        if not isinstance(label, str):
            continue
        for jahr in p.jahre:
            if label.strip() == f"Workshoptage {jahr}":
                ws.cell(row=r, column=2, value=p.jahre[jahr].workshoptage)
        if label.strip() == "Werktage Mo–Fr 2026" or label.strip() == "Werktage Mo-Fr 2026":
            ws.cell(row=r, column=2, value=p.werktage)
        if label.strip() == "Feiertage auf Werktagen 2026":
            ws.cell(row=r, column=2, value=p.feiertage)
        if label.strip() == "Urlaubsanspruch":
            ws.cell(row=r, column=2, value=p.urlaub)


def schreiben(data: bytes, p: Params) -> bytes:
    wb = openpyxl.load_workbook(BytesIO(data))
    if "Texte" not in wb.sheetnames:
        wb.create_sheet("Texte")
    _annahmen(wb, p)
    for jahr in p.jahre:
        if f"Eingabe {jahr}" not in wb.sheetnames:
            ws = wb.create_sheet(f"Eingabe {jahr}")
            ws.cell(row=1, column=1, value=f"Eingabe {jahr} (gelb = änderbar)")
        if f"Auswertung {jahr}" not in wb.sheetnames:
            wb.create_sheet(f"Auswertung {jahr}")
        _eingabe(wb, p, jahr)
        _auswertung(wb, p, jahr)
    _texte(wb, p)
    _quellen(wb, p)
    wb.properties.title = f"{p.titel} — Steuerungsdatei"
    out = BytesIO()
    wb.save(out)
    return out.getvalue()
