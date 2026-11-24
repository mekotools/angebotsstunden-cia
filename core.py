"""Rechnet das Stundengerüst aus den Parametern. Keine Seiteneffekte.

Der Rechenweg ist derselbe wie in der Steuerungsdatei (Blätter Annahmen,
Eingabe, Auswertung) und trifft die Vertragssummen exakt:

1. Stunden im Jahr  = Wochenstunden / 5 × Arbeitstage × Anteil am Jahr
2. AF 5 je Person   = AF-5-Anteil × Faktor k5, mit
                      k5 = (Soll AF 5 − Honorarstunden) / Summe der AF-5-Anteile
3. Reserve je Person= Stunden × Reserve-Anteil × Faktor kr (volle Reserve bei Anteil 1)
                      kr = (Kapazität − Vertragssoll − feste Reserve) / Summe der
                           übrigen Reserve-Ansprüche
4. AF 1             = Stunden − AF 5 − Reserve

Daraus folgt: Summe AF 5 = Soll AF 5, Summe AF 1 = Soll AF 1, Summe Reserve =
Kapazität − Vertragssoll. Diese drei Gleichungen prüft tests/test_calc.py.
"""
from __future__ import annotations

from params import Params, Person


def _jahr(p: Params, jahr: int, at: float) -> dict:
    j = p.jahre[jahr]
    personen = []
    for q in j.personen:
        anteil = q.anteil()
        stunden = q.wochenstunden / 5 * at * anteil
        personen.append(dict(
            name=q.name, rolle=q.rolle, wochenstunden=q.wochenstunden,
            von=q.von, bis=q.bis, zeitraum=q.zeitraum(jahr), anteil=anteil,
            stunden=stunden, af5_anteil=q.af5, res_anteil=q.res, fix=q.fix,
        ))

    # AF 5: Honorarkräfte decken zuerst, der Rest wird nach AF-5-Anteil verteilt
    af5_team = p.soll_af5 - j.honorarstunden
    roh = sum(d["af5_anteil"] for d in personen)
    k5 = af5_team / roh if roh else 0.0
    for d in personen:
        d["af5"] = d["af5_anteil"] * k5
        d["k5"] = k5

    kapazitaet = sum(d["stunden"] for d in personen) + j.honorarstunden
    reserve_soll = kapazitaet - p.soll
    fix_reserve = sum(d["stunden"] for d in personen if d["fix"])
    variabel = sum(d["stunden"] * d["res_anteil"] for d in personen if not d["fix"])
    kr = (reserve_soll - fix_reserve) / variabel if variabel else 0.0
    for d in personen:
        if d["fix"]:
            d["reserve"] = d["stunden"]
        else:
            d["reserve"] = d["stunden"] * d["res_anteil"] * kr
        d["af1"] = d["stunden"] - d["af5"] - d["reserve"]
        d["kr"] = kr

    summen = dict(
        stunden=sum(d["stunden"] for d in personen),
        af1=sum(d["af1"] for d in personen),
        af5_team=sum(d["af5"] for d in personen),
        af5_honorar=j.honorarstunden,
        af5=sum(d["af5"] for d in personen) + j.honorarstunden,
        reserve=sum(d["reserve"] for d in personen),
    )
    gedeckt = summen["af1"] + summen["af5"]
    return dict(
        jahr=jahr, personen=personen, honorarstunden=j.honorarstunden,
        workshoptage=j.workshoptage, kapazitaet=kapazitaet,
        reserve=reserve_soll, reserve_pct=(reserve_soll / kapazitaet * 100) if kapazitaet else 0.0,
        deckung=(gedeckt / p.soll * 100) if p.soll else 0.0, gedeckt=gedeckt,
        k5=k5, kr=kr, summen=summen,
        soll=dict(gesamt=p.soll, af1=p.soll_af1, af5=p.soll_af5),
        marker=j.marker, untertitel=j.untertitel, hinweis=j.hinweis,
    )


def compute(p: Params) -> dict:
    at = p.arbeitstage
    jahre = {jahr: _jahr(p, jahr, at) for jahr in p.jahr_liste()}
    return dict(
        jahre=jahre, arbeitstage=at, wochen=at / 5,
        werktage=p.werktage, feiertage=p.feiertage, urlaub=p.urlaub, krank=p.krank,
        arbeitstage_fix=p.arbeitstage_fix,
        soll=dict(gesamt=p.soll, af1=p.soll_af1, af5=p.soll_af5), satz=p.satz,
        titel=p.titel, org=p.org, traeger=p.traeger,
    )


def warnungen(p: Params, model: dict) -> list[str]:
    """Klartext-Hinweise für das Formular. Leer = alles plausibel."""
    w: list[str] = []
    if p.arbeitstage <= 0:
        w.append("Arbeitstage sind 0 oder negativ — Jahresstunden werden alle 0.")
    if p.soll and abs(p.soll_af1 + p.soll_af5 - p.soll) > 1:
        w.append(
            f"Soll gesamt ({p.soll:.0f} h) ist nicht die Summe aus AF 1 ({p.soll_af1:.0f} h) "
            f"und AF 5 ({p.soll_af5:.0f} h)."
        )
    for jahr, m in model["jahre"].items():
        for d in m["personen"]:
            if d["wochenstunden"] <= 0:
                w.append(f"{jahr}: {d['name']} hat keine Wochenstunden.")
        neg = [d["name"] for d in m["personen"] if d["af1"] < -0.5]
        if neg:
            w.append(f"{jahr}: AF 1 wird negativ bei {', '.join(neg)} — Reserve- oder AF-5-Anteil zu hoch.")
        if p.soll_af5 - m["honorarstunden"] < -0.5:
            w.append(f"{jahr}: Honorarstunden übersteigen das AF-5-Soll.")
        if len(m["personen"]) > 12:
            w.append(f"{jahr}: {len(m['personen'])} Zeilen — mehr als 12 passen nicht mehr sauber auf eine A4-Seite.")
        if m["reserve"] < -0.5:
            w.append(f"{jahr}: Kapazität liegt unter dem Vertragssoll — die Reserve wird negativ.")
    return w
