"""Prüfungen der Kette Excel → Rechnung → A4-PDF → Excel.

Aufruf:  STEUERDATEI=/pfad/zur/Steuerungsdatei.xlsx python -m pytest -q tests
Die Steuerungsdatei ist absichtlich **nicht** im Repository (echte Personendaten).
Ohne sie werden die dateibasierten Prüfungen übersprungen.

Die Ankerzahlen haben sich mit den Basisdaten je Jahr bewusst verschoben: statt
der von Hand gesetzten 261 Werktage / 8 Feiertage rechnet das Modell jetzt
261 − 7 = 254 Arbeitstage (Berlin 2026, Blatt Basisdaten) und Arbeitstage je
Zeitraum monatsgenau statt linear über das Jahr.
"""
from __future__ import annotations

import os
import re
import sys
import tempfile
from io import BytesIO

import pytest

HIER = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HIER))

import basisdaten  # noqa: E402
import core  # noqa: E402
import params as P  # noqa: E402
import report  # noqa: E402
import xlsx_write  # noqa: E402

STEUERDATEI = os.environ.get("STEUERDATEI", os.path.join(HIER, "steuerdatei.xlsx"))


@pytest.fixture(scope="module")
def rohdaten() -> bytes:
    if not os.path.exists(STEUERDATEI):
        pytest.skip(f"Steuerungsdatei fehlt: {STEUERDATEI} (STEUERDATEI setzen)")
    with open(STEUERDATEI, "rb") as fh:
        return fh.read()


@pytest.fixture(scope="module")
def params(rohdaten):
    p = P.from_xlsx(rohdaten)
    p.quelle = "Testdatei"
    p.stand = "01.01.2026 00:00"
    return p


@pytest.fixture(scope="module")
def model(params):
    return core.compute(params)


# --------------------------------------------------------------- Rechnung

def test_summen_treffen_den_vertrag(params, model):
    for jahr, m in model["jahre"].items():
        s = m["summen"]
        assert s["af5"] == pytest.approx(params.soll_af5, abs=0.5), jahr
        assert s["af1"] == pytest.approx(params.soll_af1, abs=0.5), jahr
        assert s["reserve"] == pytest.approx(m["kapazitaet"] - params.soll, abs=0.5), jahr


def test_jede_person_ist_vollstaendig_verteilt(model):
    for m in model["jahre"].values():
        for d in m["personen"]:
            assert d["stunden"] == pytest.approx(d["af1"] + d["af5"] + d["reserve"], abs=0.01)
            for x in d["abschnitte"]:
                assert x["stunden"] == pytest.approx(x["af1"] + x["af5"] + x["reserve"], abs=0.01)


def test_basisdaten_je_jahr_kommen_aus_dem_kalender(params, model):
    """Kein Handwert mehr: Werktage/Feiertage/Arbeitstage je Jahr aus basisdaten."""
    assert params.land == "Berlin"
    for jahr, m in model["jahre"].items():
        b = basisdaten.jahr(params.land, jahr)
        assert m["werktage"] == b["werktage"] and m["arbeitstage"] == b["arbeitstage"]
        assert m["feiertage_werktags"] == b["feiertage_werktags"]
        assert m["arbeitstage"] == m["werktage"] - m["feiertage_werktags"]
    assert model["jahre"][2026]["arbeitstage"] == 254
    assert model["jahre"][2027]["arbeitstage"] == 255


def test_zahlen_des_freigegebenen_entwurfs(model):
    """Referenz der abgenommenen A4-Seite, nach der Basisdaten-Korrektur."""
    m = model["jahre"][2026]
    soll = {
        "Julia Hoffmann": (348.8, 130.9, 0, 217.9),
        "Johannes Wiener": (1329.0, 1178.8, 29.5, 120.8),
        "Till Steinmetz": (1408.0, 1195.8, 84.3, 127.9),
        "Felix Ramm": (742.4, 531.6, 0, 210.8),
        "Sandra Rahn": (169.0, 0, 0, 169.0),
    }
    ist = {d["name"]: (d["stunden"], d["af1"], d["af5"], d["reserve"]) for d in m["personen"]}
    for name, (std, af1, af5, res) in soll.items():
        assert ist[name][0] == pytest.approx(std, abs=0.6), name
        assert ist[name][1] == pytest.approx(af1, abs=0.6), name
        assert ist[name][2] == pytest.approx(af5, abs=0.6), name
        assert ist[name][3] == pytest.approx(res, abs=0.6), name
    assert round(m["kapazitaet"]) == 8078
    assert model["jahre"][2027]["kapazitaet"] == pytest.approx(8189, abs=0.6)


def test_zeitraeume_teilen_die_stunden(params):
    """Dieselbe Person mit zwei Zeiträumen: Stunden werden monatsgenau summiert."""
    p = P.from_xlsx(open(STEUERDATEI, "rb").read())
    j = p.jahre[2026]
    person = next(q for q in j.personen if len(q.abschnitte) == 1 and q.abschnitte[0].wochenstunden)
    af5_gesamt = person.af5
    person.abschnitte = [
        P.Abschnitt(von=1, bis=6, wochenstunden=20, af5=af5_gesamt / 2, res=0.1),
        P.Abschnitt(von=7, bis=12, wochenstunden=40, af5=af5_gesamt / 2, res=0.1),
    ]
    m = core.compute(p)["jahre"][2026]
    d = next(x for x in m["personen"] if x["name"] == person.name)
    at1 = basisdaten.arbeitstage(p.land, 2026, 1, 6)
    at2 = basisdaten.arbeitstage(p.land, 2026, 7, 12)
    u1 = p.urlaub_anteilig(person.abschnitte[0])   # 20 h, halbes Jahr = 7,5 Tage
    u2 = p.urlaub_anteilig(person.abschnitte[1])   # 40 h, halbes Jahr = 15 Tage
    assert len(d["abschnitte"]) == 2
    assert d["abschnitte"][0]["stunden"] == pytest.approx(20 / 5 * (at1 - u1 - 5))
    assert d["abschnitte"][1]["stunden"] == pytest.approx(40 / 5 * (at2 - u2 - 5))
    assert d["stunden"] == pytest.approx(d["abschnitte"][0]["stunden"] + d["abschnitte"][1]["stunden"])
    assert d["wochenstunden_text"] == "20 → 40"
    # Vertragssummen stimmen weiterhin
    assert m["summen"]["af5"] == pytest.approx(p.soll_af5, abs=0.5)
    assert m["summen"]["reserve"] == pytest.approx(m["kapazitaet"] - p.soll, abs=0.5)


def test_urlaub_haengt_an_den_wochenstunden(params):
    """30 Tage bei Vollzeit (40 h): 20 h / halbes Jahr = 7,5 Tage Urlaub."""
    a = P.Abschnitt(von=7, bis=12, wochenstunden=20)
    p = P.modify(params, jahre={2026: P.Jahr(jahr=2026, personen=[
        P.Person(name="Halbjahr", abschnitte=[a])])})
    m = core.compute(p)["jahre"][2026]
    assert m["personen"][0]["urlaub"] == pytest.approx(7.5)
    assert m["personen"][0]["krank"] == pytest.approx(5)
    # Vollzeit, ganzes Jahr = 30 Tage; 32 h = 24 Tage
    assert params.urlaub_anteilig(P.Abschnitt(wochenstunden=40)) == pytest.approx(30)
    assert params.urlaub_anteilig(P.Abschnitt(wochenstunden=32)) == pytest.approx(24)


def test_formular_versteht_komma():
    p = P.from_form({"soll": "6.828", "af1": "5.000", "af5": "1.828", "satz": "38,08",
                     "n26": "1", "p2026_0_name": "A", "p2026_0_ws": "3,5", "p2026_0_res": "0,08"})
    assert p.soll == 6828 and p.satz == 38.08
    assert p.jahre[2026].personen[0].wochenstunden == 3.5
    assert p.jahre[2026].personen[0].res == 0.08


# --------------------------------------------------------- Excel-Rückweg

def test_rueckschreiben_bleibt_lesbar(rohdaten, params):
    neu = xlsx_write.schreiben(rohdaten, params)
    p2 = P.from_xlsx(neu)
    assert p2.soll == params.soll and p2.soll_af1 == params.soll_af1
    assert p2.land == params.land and p2.urlaub_vollzeit == params.urlaub_vollzeit
    for jahr in params.jahre:
        assert p2.arbeitstage(jahr) == pytest.approx(params.arbeitstage(jahr))
        assert len(p2.jahre[jahr].personen) == len(params.jahre[jahr].personen)
        assert p2.jahre[jahr].honorarstunden == params.jahre[jahr].honorarstunden
    blaetter = __import__("openpyxl").load_workbook(BytesIO(neu)).sheetnames
    for muss in ("Texte", "Basisdaten 2026", "Basisdaten 2027", "Eingabe 2026", "Auswertung 2026"):
        assert muss in blaetter, muss


def test_neue_zeile_wandert_ins_excel(rohdaten):
    p = P.from_xlsx(rohdaten)
    p.jahre[2026].personen.append(P.Person(
        name="Neue Kraft", rolle="Test",
        abschnitte=[P.Abschnitt(von=3, bis=12, wochenstunden=10, af5=40, res=0.1)]))
    neu = xlsx_write.schreiben(rohdaten, p)
    p2 = P.from_xlsx(neu)
    assert [q.name for q in p2.jahre[2026].personen][-1] == "Neue Kraft"
    m = core.compute(p2)["jahre"][2026]
    assert any(d["name"] == "Neue Kraft" for d in m["personen"])


def test_zwei_zeitraeume_ueberleben_das_excel(rohdaten):
    """Die Zeitraum-Aufteilung steht nach dem Sichern wieder in der Datei."""
    p = P.from_xlsx(rohdaten)
    q = p.jahre[2026].personen[0]
    q.abschnitte = [P.Abschnitt(von=1, bis=6, wochenstunden=20, af5=20, res=0.1),
                    P.Abschnitt(von=7, bis=12, wochenstunden=35, af5=20, res=0.1)]
    neu = xlsx_write.schreiben(rohdaten, p)
    p2 = P.from_xlsx(neu)
    q2 = p2.jahre[2026].personen[0]
    assert len(q2.abschnitte) == 2
    assert (q2.abschnitte[0].von, q2.abschnitte[0].bis) == (1, 6)
    assert (q2.abschnitte[1].von, q2.abschnitte[1].bis) == (7, 12)
    assert q2.abschnitte[0].wochenstunden == 20 and q2.abschnitte[1].wochenstunden == 35


def test_sichern_bricht_bei_verstellten_zeitraeumen_ab(rohdaten, params):
    """Selbstprüfung: passt die geschriebene Datei nicht zu den Eingaben, wird
    nicht gespeichert — sonst verbiegen Zeiträume still."""
    neu = xlsx_write.schreiben(rohdaten, params)
    verstellt = P.from_xlsx(neu)
    verstellt.jahre[2026].personen[0].abschnitte[0].bis = 12
    with pytest.raises(ValueError):
        xlsx_write._selbstpruefung(neu, verstellt)


def test_excel_formeln_rechnen_wie_die_app(rohdaten, params):
    """Unabhängiger Prüfer: die 'formulas'-Bibliothek wertet die geschriebene
    Arbeitsmappe aus. Die Zahlen müssen mit core.compute übereinstimmen — sonst
    zeigt Excel etwas anderes als das PDF."""
    formulas = pytest.importorskip("formulas")
    neu = xlsx_write.schreiben(rohdaten, params)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as fh:
        fh.write(neu)
        pfad = fh.name
    try:
        wb = formulas.ExcelModel().loads(pfad).finish()
        loesung = wb.calculate()
        m = core.compute(params)

        index = {}
        for schluessel, v in loesung.items():
            treffer = re.match(r"^'\[[^\]]+\](.+)'!\$?([A-Z]+)\$?(\d+)$", schluessel)
            if not treffer:
                continue
            try:
                zahl = float(v.value[0, 0])
            except Exception:
                try:
                    zahl = float(v)
                except Exception:
                    continue
            index[(treffer.group(1).upper(), treffer.group(2), int(treffer.group(3)))] = zahl

        def wert(blatt, zelle):
            treffer = re.match(r"^([A-Z]+)(\d+)$", zelle)
            schluessel = (blatt.upper(), treffer.group(1), int(treffer.group(2)))
            if schluessel not in index:
                raise AssertionError(f"{blatt}!{zelle} fehlt in der Excel-Auswertung")
            return index[schluessel]

        for jahr, mm in m["jahre"].items():
            blatt = f"Auswertung {jahr}"
            for i, d in enumerate(mm["personen"]):
                r = 5 + i
                assert wert(blatt, f"B{r}") == pytest.approx(d["stunden"], abs=0.01), (jahr, d["name"])
                assert wert(blatt, f"C{r}") == pytest.approx(d["af1"], abs=0.01), (jahr, d["name"])
                assert wert(blatt, f"D{r}") == pytest.approx(d["af5"], abs=0.01), (jahr, d["name"])
                assert wert(blatt, f"E{r}") == pytest.approx(d["reserve"], abs=0.01), (jahr, d["name"])
    finally:
        os.unlink(pfad)


# --------------------------------------------------------------------- PDF

def pymupdf():
    return pytest.importorskip("pymupdf")


def test_pdf_hat_eine_seite_je_jahr(params, model):
    pdf = pymupdf().open(stream=report.build(model, params, [2026, 2027]))
    assert pdf.page_count == 2
    assert "Angebotsstunden 2026" in pdf[0].get_text()
    assert "Angebotsstunden 2027" in pdf[1].get_text()


def test_pdf_einzeljahr(params, model):
    pdf = pymupdf().open(stream=report.build(model, params, [2027]))
    assert pdf.page_count == 1
    assert "2027" in pdf[0].get_text()


def test_pdf_zeigt_basisdaten_und_zeitraeume(params):
    p = P.from_xlsx(open(STEUERDATEI, "rb").read())
    q = p.jahre[2026].personen[0]
    q.abschnitte = [P.Abschnitt(von=1, bis=6, wochenstunden=20, af5=20, res=0.1),
                    P.Abschnitt(von=7, bis=12, wochenstunden=35, af5=20, res=0.1)]
    m = core.compute(p)
    text = pymupdf().open(stream=report.build(m, p, [2026]))[0].get_text().replace("\n", " ")
    assert "254 Arbeitstage" in text and "261 Werktage" in text
    assert "20 → 35" in text and "01.–06." in text


def test_pdf_zeigt_justierte_zahl(params, model):
    """Eine geänderte Annahme muss im PDF ankommen — Beweis, dass nichts fest verdrahtet ist."""
    p2 = P.from_form({**params.to_form(), "k": "15"})
    m2 = core.compute(p2)
    kapazitaet = f"{round(m2['jahre'][2026]['kapazitaet']):,}".replace(",", ".")
    assert kapazitaet != f"{round(model['jahre'][2026]['kapazitaet']):,}".replace(",", ".")
    text = pymupdf().open(stream=report.build(m2, p2, [2026]))[0].get_text()
    assert kapazitaet in text


def test_pdf_bleibt_auf_einer_seite(params):
    p = P.from_xlsx(open(STEUERDATEI, "rb").read())
    for _ in range(6):
        p.jahre[2026].personen.append(P.Person(
            name="Zusatzkraft", rolle="Test",
            abschnitte=[P.Abschnitt(von=1, bis=12, wochenstunden=8, af5=20, res=0.1)]))
    m = core.compute(p)
    pdf = pymupdf().open(stream=report.build(m, p, [2026]))
    assert pdf.page_count == 1
    bloecke = [b for b in pdf[0].get_text("blocks")]
    assert max(b[3] for b in bloecke) < 830  # Punkte: A4 = 842 hoch, nichts läuft heraus
    unten = [b for b in bloecke if b[1] > 250]
    assert unten, "Fußzeile mit der Herleitung fehlt"


# ------------------------------------------------------------------ Warnungen

def test_warnung_bei_schiefem_soll(params, model):
    p = P.from_form({**params.to_form(), "soll": "7000"})
    assert any("nicht die Summe" in w for w in core.warnungen(p, core.compute(p)))


def test_warnung_bei_zu_vielen_zeilen(params):
    p = P.from_xlsx(open(STEUERDATEI, "rb").read())
    for _ in range(5):
        p.jahre[2026].personen.append(P.Person(
            name=f"X{len(p.jahre[2026].personen)}",
            abschnitte=[P.Abschnitt(von=1, bis=12, wochenstunden=4, res=0.1)]))
    assert any("A4-Seite" in w for w in core.warnungen(p, core.compute(p)))


def test_overlay_ueberschreibt_nur_das_uebergebene(rohdaten):
    p = P.from_xlsx(rohdaten)
    q = P.overlay(p, {"k": "15"})
    assert q.krank == 15
    # Arbeitstage kommen aus dem Kalender, nicht aus Urlaub/Krankheit
    assert q.arbeitstage(2026) == 254
    assert q.soll == p.soll and q.satz == p.satz
    assert [x.name for x in q.jahre[2026].personen] == [x.name for x in p.jahre[2026].personen]
    # Zeilenzahl im Formular ersetzt die Personenliste
    q2 = P.overlay(p, {"n26": "1", "p2026_0_name": "Nur eine", "p2026_0_ws": "20",
                       "p2026_0_res": "0,1"})
    assert [x.name for x in q2.jahre[2026].personen] == ["Nur eine"]
    assert len(q2.jahre[2027].personen) == len(p.jahre[2027].personen)
