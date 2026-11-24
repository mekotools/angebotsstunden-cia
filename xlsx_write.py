"""Schreibt die Parameter in die Steuerungsdatei zurück.

Die XLSX bleibt die Datenquelle:

* ``Basisdaten <Jahr>`` — Werktage, Feiertage und Arbeitstage je Monat eines
  Jahres, berechnet aus dem Kalender (Bundesland steht im Blatt ``Quellen``).
* ``Annahmen`` — Vertrag, Jahresannahmen.
* ``Eingabe <Jahr>`` — **eine Zeile je Zeitraum** (Monat von … bis) mit eigenen
  Wochenstunden, Urlaubs- und Krankheitstagen. Arbeitstage und Stunden rechnet
  das Blatt selbst.
* ``Auswertung <Jahr>`` — eine Zeile je Person, summiert ihre Zeiträume und
  rechnet mit denselben Formeln wie die App (k5 und kr als eigene Zellen).
* ``Texte`` — Überschriften der A4-Seite.

So zeigen App und Excel nach dem Sichern dieselben Zahlen — auch wenn Zeilen
hinzukommen oder wegfallen, weil die Formeln für die neue Zeilenzahl neu gesetzt
werden.
"""
from __future__ import annotations

from io import BytesIO

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import basisdaten
from params import Params

PETROL = "FF1E4B5A"
GELB = "FFFFF2CC"
GRAU = "FFF2F5F6"

MONATSNAMEN = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
               "August", "September", "Oktober", "November", "Dezember"]


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


def _leeren(ws, ab_zeile: int, bis_spalte: int = 12):
    for r in range(ab_zeile, ws.max_row + 1):
        for c in range(1, bis_spalte + 1):
            ws.cell(row=r, column=c).value = None


def _annahmen(wb, p: Params):
    """Annahmen-Blatt neu schreiben (alte Zeilen auch löschen, sonst bleiben
    Reste wie ein zweiter 'Urlaubsanspruch' stehen)."""
    ws = wb["Annahmen"]
    _leeren(ws, 4, 4)
    werte = [
        ("Krankheitstage je Jahr (Vorgabe)", p.krank,
         "Annahme; gilt anteilig für Zeiträume ohne eigene Angabe"),
        ("Urlaub bei Vollzeit (Tage)", p.urlaub_vollzeit,
         "Jahresanspruch einer Vollzeitstelle"),
        ("Vollzeit-Wochenstunden", p.vollzeit_stunden,
         "Bezug für den Urlaub: Tage = Vollzeiturlaub × Wochenstunden / Vollzeit"),
        ("Arbeitstage erstes Jahr", p.arbeitstage(p.jahr_liste()[0]),
         "Werktage − Feiertage, Blatt Basisdaten (Kalender, nicht geschätzt)"),
        ("Arbeitstage zweites Jahr", p.arbeitstage(p.jahr_liste()[-1]),
         "Werktage − Feiertage, Blatt Basisdaten"),
        ("Leistungsstunden gesamt", p.soll, "Vertrag § 5 Nr. 1"),
        ("davon AF 1", p.soll_af1, ""),
        ("davon AF 5", p.soll_af5, ""),
        ("Kostensatz (€/h)", p.satz, "Leistungsvertrag § 5 Nr. 1"),
        ("Bundesland", p.land, "bestimmt die Feiertage (Blatt Basisdaten)"),
    ]
    for i, (label, wert, grund) in enumerate(werte, start=4):
        a = ws.cell(row=i, column=1, value=label)
        a.font = Font(size=10)
        _wert(ws, i, 2, wert)
        c = ws.cell(row=i, column=3, value=grund)
        c.font = Font(size=9, color="FF6B7C85")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 56
    # Adressen zurückgeben, damit die Auswertung nicht auf Zeilennummern baut
    return {
        "krank": "B4",
        "soll": f"B{4 + 5}",
        "af5": f"B{4 + 7}",
        "land": f"B{4 + 9}",
    }


def _basisdaten(wb, p: Params, jahr: int) -> str:
    """Blatt mit der Monatstabelle; liefert den Blattnamen für Formeln."""
    name = f"Basisdaten {jahr}"
    if name in wb.sheetnames:
        ws = wb[name]
        _leeren(ws, 1)
    else:
        ws = wb.create_sheet(name)
    b = p.basis(jahr)
    ws.cell(row=1, column=1, value=f"Basisdaten {jahr} — {b['land']}").font = Font(bold=True, size=11)
    ws.cell(row=2, column=1,
            value="Aus dem Kalender gerechnet: Werktage = Mo–Fr, Feiertage nach Landesrecht, "
                  "Arbeitstage = Werktage − Feiertage auf Werktagen.").font = Font(
        size=9, color="FF6B7C85")
    _kopf(ws, ["Monat", "Werktage (Mo–Fr)", "Feiertage werktags", "Arbeitstage"], zeile=3)
    for i, z in enumerate(b["monate"]):
        r = 4 + i
        ws.cell(row=r, column=1, value=MONATSNAMEN[i]).font = Font(size=10)
        for sp, wert in ((2, z["werktage"]), (3, z["feiertage"]), (4, z["arbeitstage"])):
            c = ws.cell(row=r, column=sp, value=wert)
            c.alignment = Alignment(horizontal="center")
    r_sum = 4 + len(b["monate"])
    ws.cell(row=r_sum, column=1, value="Jahr gesamt").font = Font(bold=True, size=10)
    for sp in (2, 3, 4):
        c = ws.cell(row=r_sum, column=sp, value=f"=SUM({chr(64 + sp)}4:{chr(64 + sp)}{r_sum - 1})")
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")

    r_fe = r_sum + 2
    ws.cell(row=r_fe, column=1, value="Gesetzliche Feiertage").font = Font(bold=True, size=10)
    ws.cell(row=r_fe + 1, column=1, value="Datum").font = Font(size=9, color="FF6B7C85")
    ws.cell(row=r_fe + 1, column=2, value="Name").font = Font(size=9, color="FF6B7C85")
    for i, (datum, name_feiertag) in enumerate(b["daten"]):
        ws.cell(row=r_fe + 2 + i, column=1, value=datum).alignment = Alignment(horizontal="left")
        ws.cell(row=r_fe + 2 + i, column=2, value=name_feiertag).font = Font(size=9)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 14
    return name


def _eingabe(wb, p: Params, jahr: int, basis_blatt: str):
    """Eine Zeile je Zeitraum.

    Gelb ist, was man setzt; weiß, was daraus folgt. Urlaub und Krankheit haben
    je eine Eingabespalte und eine Spalte „wirksam": bleibt die Eingabe leer,
    rechnet Excel denselben anteiligen Wert wie die App (Urlaub aus
    Vollzeitanspruch × Wochenstunden ÷ Vollzeit-Wochenstunden × Monatsanteil).
    """
    ws = wb[f"Eingabe {jahr}"]
    _leeren(ws, 1, 14)
    titel = ws.cell(row=1, column=1, value=f"Eingabe {jahr} (gelb = änderbar)")
    titel.font = Font(bold=True, size=11)
    ws.cell(row=2, column=1,
            value="Eine Zeile je Zeitraum. Gleicher Name = gleiche Person "
                  "(z. B. Stundenänderung ab Juli). Urlaub/Krankheit leer = anteilig.").font = Font(
        size=9, color="FF6B7C85")
    _kopf(ws, ["Mitarbeiter*in", "Funktion", "von Monat", "bis Monat", "Wochenstunden",
               "Urlaub (Eingabe)", "Urlaub wirksam", "Krank (Eingabe)", "Krank wirksam",
               "AF 5 (h)", "Reserve-Anteil", "Arbeitstage", "Stunden"])
    q = f"'{basis_blatt}'"
    r = 4
    for person in p.jahre[jahr].personen:
        for a in person.abschnitte:
            z = ws.cell(row=r, column=1, value=person.name)
            z.font = Font(bold=True, size=10)
            ws.cell(row=r, column=2, value=person.rolle).font = Font(size=10)
            _wert(ws, r, 3, a.von)
            _wert(ws, r, 4, a.bis)
            _wert(ws, r, 5, a.wochenstunden)
            _wert(ws, r, 6, "" if a.urlaub is None else a.urlaub)
            _wert(ws, r, 8, "" if a.krank is None else a.krank)
            _wert(ws, r, 10, a.af5)
            _wert(ws, r, 11, a.res)
            monate = "(D{r}-C{r}+1)/12".format(r=r)
            ws.cell(row=r, column=7,
                    value=f'=IF(F{r}="",Annahmen!$B$5*E{r}/Annahmen!$B$6*{monate},F{r})')
            ws.cell(row=r, column=9,
                    value=f'=IF(H{r}="",Annahmen!$B$4*{monate},H{r})')
            at = "+".join(f"{q}!D{4 + m - 1}" for m in range(a.von, a.bis + 1))
            ws.cell(row=r, column=12, value=f"={at}")
            ws.cell(row=r, column=13, value=f"=E{r}/5*(L{r}-G{r}-I{r})")
            for sp in (5, 10, 11):
                ws.cell(row=r, column=sp).number_format = "#,##0.0##"
            for sp in (6, 7, 8, 9):
                ws.cell(row=r, column=sp).number_format = "#,##0.0#"
            for sp in (12, 13):
                ws.cell(row=r, column=sp).number_format = "#,##0.0"
            r += 1
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 26
    for sp, breite in (("C", 9), ("D", 9), ("E", 11), ("F", 11), ("G", 11), ("H", 10),
                       ("I", 11), ("J", 9), ("K", 11), ("L", 10), ("M", 10)):
        ws.column_dimensions[sp].width = breite
    ws.freeze_panes = "A4"
    return r - 1  # letzte Datenzeile


def _auswertung(wb, p: Params, jahr: int, letzte_eingabe: int, an: dict):
    """Eine Zeile je Person. Die Faktoren k5/kr stehen als Zellen darunter,
    damit Excel und App nachweisbar gleich rechnen."""
    ws = wb[f"Auswertung {jahr}"]
    _leeren(ws, 1, 8)
    j = p.jahre[jahr]
    n = len(j.personen)
    erste, letzte = 5, 4 + n

    ws.cell(row=1, column=1, value=f"Auswertung {jahr}").font = Font(bold=True, size=11)
    ws.cell(row=2, column=1,
            value="Rechnet ausschließlich aus den Blättern Eingabe, Basisdaten und Annahmen.").font = Font(
        size=9, color="FF6B7C85")
    _kopf(ws, ["Mitarbeiter*in", "Stunden im Jahr", "AF 1", "AF 5", "Reserve",
               "Reserve-Anteil"], zeile=4)

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
    for i, person in enumerate(j.personen):
        z = erste + i
        # Zeilen dieser Person im Blatt Eingabe
        zeilen = []
        r_eingabe = 4
        for vorherige in j.personen:
            if vorherige is person:
                zeilen = list(range(r_eingabe, r_eingabe + len(person.abschnitte)))
                break
            r_eingabe += len(vorherige.abschnitte)
        ws.cell(row=z, column=1, value=person.name)
        ws.cell(row=z, column=2, value="=" + "+".join(f"{q}!M{x}" for x in zeilen))
        ws.cell(row=z, column=4, value="=(" + "+".join(f"{q}!J{x}" for x in zeilen) + f")*$B${r_k5}")
        ws.cell(row=z, column=6, value="=" + "+".join(f"{q}!K{x}" for x in zeilen))
        ws.cell(row=z, column=5, value=f"=IF(F{z}>=0.999,B{z},B{z}*F{z}*$B${r_kr})")
        ws.cell(row=z, column=3, value=f"=B{z}-D{z}-E{z}")

    ws.cell(row=r_sum, column=1, value="Summe Team").font = Font(bold=True)
    for sp in "BCDE":
        ws[f"{sp}{r_sum}"] = f"=SUM({sp}{erste}:{sp}{letzte})"
        ws[f"{sp}{r_sum}"].font = Font(bold=True)

    ws.cell(row=r_hon, column=1, value="Honorarkräfte")
    ws.cell(row=r_hon, column=4, value=j.honorarstunden)
    ws.cell(row=r_af5, column=1, value="Summe AF 5").font = Font(bold=True)
    ws.cell(row=r_af5, column=4, value=f"=D{r_sum}+D{r_hon}")
    ws.cell(row=r_kap, column=1, value="Kapazität").font = Font(bold=True)
    ws.cell(row=r_kap, column=2, value=f"=B{r_sum}+D{r_hon}")
    ws.cell(row=r_soll, column=1, value="Vertragssoll").font = Font(bold=True)
    ws.cell(row=r_soll, column=2, value=f"=Annahmen!${an['soll'][0]}${an['soll'][1:]}")
    ws.cell(row=r_deck, column=1, value="Deckung").font = Font(bold=True)
    ws.cell(row=r_deck, column=2, value=f"=(C{r_sum}+D{r_af5})/B{r_soll}")
    ws.cell(row=r_deck, column=2).number_format = "0.0%"
    ws.cell(row=r_res, column=1, value="Reserve = Kapazität − Soll").font = Font(size=9)
    ws.cell(row=r_res, column=2, value=f"=B{r_kap}-B{r_soll}")

    # Faktoren (offen dokumentiert, damit nachvollziehbar bleibt, wie verteilt wird)
    ws.cell(row=r_k5, column=1, value="Faktor AF 5 (k5)").font = Font(size=9, italic=True)
    ws.cell(row=r_k5, column=2,
            value=f"=(Annahmen!${an['af5'][0]}${an['af5'][1:]}-D{r_hon})"
                  f"/SUM({q}!$J$4:$J${letzte_eingabe})")
    ws.cell(row=r_kapz, column=1, value="Reserve-Soll (Kapazität − Soll)").font = Font(size=9, italic=True)
    ws.cell(row=r_kapz, column=2, value=f"=B{r_kap}-Annahmen!${an['soll'][0]}${an['soll'][1:]}")
    ws.cell(row=r_fix, column=1, value="davon feste Reserve (Anteil 1)").font = Font(size=9, italic=True)
    ws.cell(row=r_fix, column=2, value=f'=SUMIF(F{erste}:F{letzte},">=0.999",B{erste}:B{letzte})')
    ws.cell(row=r_var, column=1, value="übrige Reserve-Ansprüche (Stunden × Anteil)").font = Font(
        size=9, italic=True)
    ws.cell(row=r_var, column=2, value=f"=SUMPRODUCT((F{erste}:F{letzte}<0.999)*B{erste}:B{letzte}"
                                       f"*F{erste}:F{letzte})")
    ws.cell(row=r_kr, column=1, value="Faktor Reserve (kr)").font = Font(size=9, italic=True)
    ws.cell(row=r_kr, column=2, value=f"=(B{r_kapz}-B{r_fix})/B{r_var}")

    for sp, breite in (("A", 30), ("B", 18), ("C", 14), ("D", 14), ("E", 14), ("F", 14)):
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
    """Blatt Quellen: Bundesland und Jahreswerte aktualisieren, Fehlendes anhängen."""
    if "Quellen" not in wb.sheetnames:
        ws = wb.create_sheet("Quellen")
        ws.cell(row=1, column=1, value="Quellen und Jahreswerte").font = Font(bold=True, size=11)
    else:
        ws = wb["Quellen"]
    vorhanden = {}
    for r in range(1, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        if isinstance(label, str) and label.strip():
            vorhanden.setdefault(label.strip(), r)

    zu_setzen: dict[str, object] = {
        "Bundesland": p.land,
        "Urlaub bei Vollzeit": p.urlaub_vollzeit,
        "Vollzeit-Wochenstunden": p.vollzeit_stunden,
    }
    for jahr in p.jahr_liste():
        b = p.basis(jahr)
        zu_setzen[f"Werktage Mo–Fr {jahr}"] = b["werktage"]
        zu_setzen[f"Feiertage auf Werktagen {jahr}"] = b["feiertage_werktags"]
        zu_setzen[f"Feiertage gesamt {jahr}"] = b["feiertage"]
        zu_setzen[f"Arbeitstage {jahr}"] = b["arbeitstage"]
        zu_setzen[f"Workshoptage {jahr}"] = p.jahre[jahr].workshoptage

    # alter Einzelwert ist überholt (er war die 24 aus Till Steinmetz' Blatt)
    for alt_label in ("Urlaubsanspruch",):
        r = vorhanden.get(alt_label)
        if r is not None:
            ws.cell(row=r, column=2, value=None)
            ws.cell(row=r, column=1, value=f"{alt_label} (früher)")

    naechste = ws.max_row + 1
    for label, wert in zu_setzen.items():
        r = vorhanden.get(label)
        if r is None:
            r = naechste
            naechste += 1
            ws.cell(row=r, column=1, value=label).font = Font(size=10)
        ws.cell(row=r, column=2, value=wert)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 16


def schreiben(data: bytes, p: Params) -> bytes:
    wb = openpyxl.load_workbook(BytesIO(data))
    if "Texte" not in wb.sheetnames:
        wb.create_sheet("Texte")
    an = _annahmen(wb, p)
    for jahr in p.jahre:
        if f"Eingabe {jahr}" not in wb.sheetnames:
            ws = wb.create_sheet(f"Eingabe {jahr}")
            ws.cell(row=1, column=1, value=f"Eingabe {jahr} (gelb = änderbar)")
        if f"Auswertung {jahr}" not in wb.sheetnames:
            wb.create_sheet(f"Auswertung {jahr}")
        basis_blatt = _basisdaten(wb, p, jahr)
        letzte = _eingabe(wb, p, jahr, basis_blatt)
        _auswertung(wb, p, jahr, letzte, an)
    _texte(wb, p)
    _quellen(wb, p)
    wb.properties.title = f"{p.titel} — Steuerungsdatei"
    out = BytesIO()
    wb.save(out)
    return out.getvalue()
