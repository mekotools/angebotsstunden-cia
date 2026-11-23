"""Liest die Steuerungs-XLSX und rechnet das Stundengerüst. Keine Seiteneffekte."""
from __future__ import annotations
from io import BytesIO
import calendar
import openpyxl

SOLL_AF1 = 5000
SOLL_AF5 = 1828
SOLL_GES = SOLL_AF1 + SOLL_AF5


def _zeitraum(von, bis):
    if von == 1 and bis == 12:
        return "01.01.–31.12."
    letzter = calendar.monthrange(2026, bis)[1]
    return f"01.{von:02d}.–{letzter:02d}.{bis:02d}."


def load(data: bytes) -> dict:
    wb = openpyxl.load_workbook(BytesIO(data), data_only=False)
    an = wb["Annahmen"]
    arbeitstage = float(an["B5"].value)
    soll_af5 = float(an["B9"].value)
    soll_af1 = float(an["B8"].value)

    jahre = {}
    for jahr in (2026, 2027):
        ws = wb[f"Eingabe {jahr}"]
        personen = []
        r = 4
        while ws.cell(row=r, column=1).value:
            von = int(ws.cell(row=r, column=4).value)
            bis = int(ws.cell(row=r, column=5).value)
            personen.append(dict(
                name=ws.cell(row=r, column=1).value,
                rolle=ws.cell(row=r, column=2).value,
                wochenstunden=float(ws.cell(row=r, column=3).value),
                von=von, bis=bis, zeitraum=_zeitraum(von, bis),
                af5=float(ws.cell(row=r, column=6).value or 0),
                res=float(ws.cell(row=r, column=7).value or 0),
                quelle=ws.cell(row=r, column=8).value,
            ))
            r += 1
        jahre[jahr] = personen

    # Honorarstunden aus dem Quellenblatt lesen (Zeilen mit „Honorarstunden“)
    ws = wb["Quellen"]
    hon = {}
    for r in range(1, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        if isinstance(label, str) and label.startswith("Honorarstunden"):
            jahr = 2026 if "2026" in label else 2027
            hon[jahr] = float(ws.cell(row=r, column=2).value)
    hon.setdefault(2026, 1630.0)
    hon.setdefault(2027, 1200.0)

    ergebnis = {}
    for jahr, personen in jahre.items():
        hon_h = hon[jahr]
        af5_team = soll_af5 - hon_h
        daten = []
        for p in personen:
            anteil = (p["bis"] - p["von"] + 1) / 12
            netto = round(p["wochenstunden"] / 5 * arbeitstage * anteil, 1)
            daten.append(dict(p, anteil=round(anteil, 4), netto=netto,
                              fix=p["res"] >= 0.999))
        roh = sum(d["af5"] for d in daten)
        k5 = af5_team / roh if roh else 0
        for d in daten:
            d["af5"] = d["af5"] * k5
        kap = sum(d["netto"] for d in daten) + hon_h
        res_soll = kap - SOLL_GES
        res_fix = sum(d["netto"] for d in daten if d["fix"])
        kr = (res_soll - res_fix) / sum(d["netto"] * d["res"] for d in daten if not d["fix"])
        for d in daten:
            if d["fix"]:
                d["reserve"], d["af1"] = d["netto"], 0.0
            else:
                d["reserve"] = d["netto"] * d["res"] * kr
                d["af1"] = d["netto"] - d["af5"] - d["reserve"]
        ergebnis[jahr] = dict(
            jahr=jahr, personen=daten, honorarstunden=hon_h, kapazitaet=kap,
            reserve=res_soll,
            summen=dict(netto=sum(d["netto"] for d in daten),
                        af1=sum(d["af1"] for d in daten),
                        af5_team=sum(d["af5"] for d in daten),
                        af5_honorar=hon_h,
                        reserve=sum(d["reserve"] for d in daten)),
            soll=dict(af1=soll_af1, af5=soll_af5, ges=soll_af1 + soll_af5),
            reserve_pct=res_soll / kap * 100 if kap else 0,
        )
    return dict(jahre=ergebnis, arbeitstage=arbeitstage,
                wochen=round(arbeitstage / 5, 1))
