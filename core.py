"""Rechnet das Stundengerüst aus den Parametern. Keine Seiteneffekte.

Der Rechenweg ist derselbe wie in der Steuerungsdatei (Blätter Basisdaten,
Annahmen, Eingabe, Auswertung) und trifft die Vertragssummen exakt:

1. Arbeitstage eines Abschnitts = Werktage − Feiertage des Zeitraums (aus
   ``basisdaten``, je Bundesland und Jahr) — abzüglich Urlaub und Krankheit
2. Stunden im Abschnitt = Wochenstunden / 5 × Arbeitstage des Abschnitts
   Die Stunden einer Person sind die Summe ihrer Abschnitte, damit eine
   Stundenänderung mitten im Jahr tagesgenau (monatsgenau) wirkt.
3. AF 5 je Person   = AF-5-Anteil × Faktor k5, mit
                      k5 = (Soll AF 5 − Honorarstunden) / Summe der AF-5-Anteile
4. Reserve je Person= Stunden × Reserve-Anteil × Faktor kr (volle Reserve bei Anteil 1)
                      kr = (Kapazität − Vertragssoll − feste Reserve) / Summe der
                           übrigen Reserve-Ansprüche
5. AF 1             = Stunden − AF 5 − Reserve

Daraus folgt: Summe AF 5 = Soll AF 5, Summe AF 1 = Soll AF 1, Summe Reserve =
Kapazität − Vertragssoll. Diese drei Gleichungen prüft tests/test_kette.py.
"""
from __future__ import annotations

import basisdaten
from params import Params


def _abschnitt(p: Params, jahr: int, a) -> dict:
    at = basisdaten.arbeitstage(p.land, jahr, a.von, a.bis)
    u = a.urlaub if a.urlaub is not None else p.urlaub * a.anteil
    k = a.krank if a.krank is not None else p.krank * a.anteil
    netto = max(0.0, at - u - k)
    return dict(
        von=a.von, bis=a.bis, zeitraum=a.zeitraum(), monate=a.monate,
        wochenstunden=a.wochenstunden, arbeitstage=at, urlaub=u, krank=k,
        netto=netto, stunden=a.wochenstunden / 5 * netto,
        af5_anteil=a.af5, res_anteil=a.res, fix=a.fix,
    )


def _jahr(p: Params, jahr: int) -> dict:
    j = p.jahre[jahr]
    b = p.basis(jahr)
    personen = []
    for q in j.personen:
        abschnitte = [_abschnitt(p, jahr, a) for a in q.abschnitte]
        stunden = sum(x["stunden"] for x in abschnitte)
        personen.append(dict(
            name=q.name, rolle=q.rolle, abschnitte=abschnitte,
            wochenstunden=abschnitte[0]["wochenstunden"] if abschnitte else 0.0,
            wochenstunden_text=q.ws_text,
            zeitraum=abschnitte[0]["zeitraum"] if len(abschnitte) == 1 else q.zeitraum_text,
            zeitraum_text=q.zeitraum_text,
            stunden=stunden,
            af5_anteil=sum(x["af5_anteil"] for x in abschnitte),
            fix=any(x["fix"] for x in abschnitte),
            fix_reserve=sum(x["stunden"] for x in abschnitte if x["fix"]),
            variabel=sum(x["stunden"] * x["res_anteil"] for x in abschnitte if not x["fix"]),
            urlaub=sum(x["urlaub"] for x in abschnitte),
            krank=sum(x["krank"] for x in abschnitte),
            arbeitstage=sum(x["arbeitstage"] for x in abschnitte),
        ))

    # AF 5: Honorarkräfte decken zuerst, der Rest wird nach AF-5-Anteil verteilt
    af5_team = p.soll_af5 - j.honorarstunden
    roh = sum(d["af5_anteil"] for d in personen)
    k5 = af5_team / roh if roh else 0.0

    kapazitaet = sum(d["stunden"] for d in personen) + j.honorarstunden
    reserve_soll = kapazitaet - p.soll
    fix_reserve = sum(d["fix_reserve"] for d in personen)
    variabel = sum(d["variabel"] for d in personen)
    kr = (reserve_soll - fix_reserve) / variabel if variabel else 0.0

    for d in personen:
        d["k5"], d["kr"] = k5, kr
        for x in d["abschnitte"]:
            x["af5"] = x["af5_anteil"] * k5
            x["reserve"] = x["stunden"] if x["fix"] else x["stunden"] * x["res_anteil"] * kr
            x["af1"] = x["stunden"] - x["af5"] - x["reserve"]
        d["af5"] = sum(x["af5"] for x in d["abschnitte"])
        d["reserve"] = sum(x["reserve"] for x in d["abschnitte"])
        d["af1"] = sum(x["af1"] for x in d["abschnitte"])

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
        basis=b, land=b["land"],
        arbeitstage=b["arbeitstage"], werktage=b["werktage"],
        feiertage=b["feiertage"], feiertage_werktags=b["feiertage_werktags"],
        reserve=reserve_soll, reserve_pct=(reserve_soll / kapazitaet * 100) if kapazitaet else 0.0,
        deckung=(gedeckt / p.soll * 100) if p.soll else 0.0, gedeckt=gedeckt,
        k5=k5, kr=kr, summen=summen,
        soll=dict(gesamt=p.soll, af1=p.soll_af1, af5=p.soll_af5),
        marker=j.marker, untertitel=j.untertitel, hinweis=j.hinweis,
    )


def compute(p: Params) -> dict:
    jahre = {jahr: _jahr(p, jahr) for jahr in p.jahr_liste()}
    return dict(
        jahre=jahre, land=p.land,
        urlaub=p.urlaub, krank=p.krank,
        soll=dict(gesamt=p.soll, af1=p.soll_af1, af5=p.soll_af5), satz=p.satz,
        titel=p.titel, org=p.org, traeger=p.traeger,
    )


def warnungen(p: Params, model: dict) -> list[str]:
    """Klartext-Hinweise für das Formular. Leer = alles plausibel."""
    w: list[str] = []
    if p.soll and abs(p.soll_af1 + p.soll_af5 - p.soll) > 1:
        w.append(
            f"Soll gesamt ({p.soll:.0f} h) ist nicht die Summe aus AF 1 ({p.soll_af1:.0f} h) "
            f"und AF 5 ({p.soll_af5:.0f} h)."
        )
    for jahr, m in model["jahre"].items():
        zeilen = sum(len(d["abschnitte"]) for d in m["personen"])
        for d in m["personen"]:
            for x in d["abschnitte"]:
                if x["wochenstunden"] <= 0:
                    w.append(f"{jahr}: {d['name']} hat im Zeitraum {x['zeitraum']} "
                             f"keine Wochenstunden.")
                if x["netto"] <= 0:
                    w.append(f"{jahr}: {d['name']} hat im Zeitraum {x['zeitraum']} "
                             f"keine Arbeitstage (Urlaub/Krankheit zu hoch?).")
        neg = [d["name"] for d in m["personen"] if d["af1"] < -0.5]
        if neg:
            w.append(f"{jahr}: AF 1 wird negativ bei {', '.join(neg)} — Reserve- oder AF-5-Anteil zu hoch.")
        if p.soll_af5 - m["honorarstunden"] < -0.5:
            w.append(f"{jahr}: Honorarstunden übersteigen das AF-5-Soll.")
        if zeilen > 12:
            w.append(f"{jahr}: {zeilen} Zeiträume — mehr als 12 Zeilen passen nicht mehr "
                     f"sauber auf eine A4-Seite.")
        if m["reserve"] < -0.5:
            w.append(f"{jahr}: Kapazität liegt unter dem Vertragssoll — die Reserve wird negativ.")
    return w
