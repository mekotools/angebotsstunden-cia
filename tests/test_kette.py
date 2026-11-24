"""Prüfungen der Kette Excel → Rechnung → A4-PDF → Excel.

Aufruf:  STEUERDATEI=/pfad/zur/Steuerungsdatei.xlsx python -m pytest -q tests
Die Steuerungsdatei ist absichtlich **nicht** im Repository (echte Personendaten).
Ohne sie werden die dateibasierten Prüfungen übersprungen.
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


def test_zahlen_des_freigegebenen_entwurfs(model):
    """Die von Hand erzeugte A4-Seite (Entwurf 23.11.2026) als Referenz."""
    m = model["jahre"][2026]
    soll = {
        "Julia Hoffmann": (350, 159, 0, 191),
        "Johannes Wiener": (1314, 1180, 29, 104),
        "Till Steinmetz": (1402, 1206, 84, 111),
        "Felix Ramm": (701, 527, 0, 174),
        "Sandra Rahn": (153, 0, 0, 153),
    }
    ist = {d["name"]: (d["stunden"], d["af1"], d["af5"], d["reserve"]) for d in m["personen"]}
    for name, (std, af1, af5, res) in soll.items():
        assert ist[name][0] == pytest.approx(std, abs=0.6), name
        assert ist[name][1] == pytest.approx(af1, abs=0.6), name
        assert ist[name][2] == pytest.approx(af5, abs=0.6), name
        assert ist[name][3] == pytest.approx(res, abs=0.6), name
    assert round(m["kapazitaet"]) == 7915
    assert model["jahre"][2027]["kapazitaet"] == pytest.approx(7967, abs=0.6)


def test_arbeitstage_von_hand(params):
    p = P.from_form({
        "w": "261", "f": "8", "u": "24", "k": "10", "soll": "6828", "af1": "5000", "af5": "1828",
        "at_fix": "1", "at": "200",
        "n26": "1", "p2026_0_name": "Test", "p2026_0_ws": "40", "p2026_0_von": "1",
        "p2026_0_bis": "12", "p2026_0_af5": "0", "p2026_0_res": "0",
        "h26": "1630", "n27": "0", "h27": "0",
    })
    assert p.arbeitstage == 200
    m = core.compute(p)["jahre"][2026]
    assert m["personen"][0]["stunden"] == pytest.approx(40 / 5 * 200)


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
    assert p2.arbeitstage == pytest.approx(params.arbeitstage)
    for jahr in params.jahre:
        assert len(p2.jahre[jahr].personen) == len(params.jahre[jahr].personen)
        assert p2.jahre[jahr].honorarstunden == params.jahre[jahr].honorarstunden
    assert "Texte" in __import__("openpyxl").load_workbook(BytesIO(neu)).sheetnames


def test_neue_zeile_wandert_ins_excel(rohdaten, params):
    p = P.from_xlsx(rohdaten)
    p.jahre[2026].personen.append(P.Person(name="Neue Kraft", rolle="Test", wochenstunden=10,
                                           von=3, bis=12, af5=40, res=0.1))
    neu = xlsx_write.schreiben(rohdaten, p)
    p2 = P.from_xlsx(neu)
    assert [q.name for q in p2.jahre[2026].personen][-1] == "Neue Kraft"
    # und die Rechnung nimmt die Zeile mit
    m = core.compute(p2)["jahre"][2026]
    assert any(d["name"] == "Neue Kraft" for d in m["personen"])


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


def test_pdf_zeigt_justierte_zahl(params, model):
    """Eine geänderte Zahl muss im PDF stehen — Beweis, dass nichts fest verdrahtet ist."""
    p2 = P.from_form({**params.to_form(), "k": "15"})
    m2 = core.compute(p2)
    text = pymupdf().open(stream=report.build(m2, p2, [2026]))[0].get_text()
    assert "Arbeitstage 214" in text.replace("\n", " ")  # 261 − 8 − 24 − 15
    assert f"{round(m2['jahre'][2026]['kapazitaet']):,}".replace(",", ".") in text


def test_pdf_bleibt_auf_einer_seite(params):
    p = P.from_xlsx(open(STEUERDATEI, "rb").read())
    for _ in range(6):
        p.jahre[2026].personen.append(P.Person(name="Zusatzkraft", rolle="Test",
                                               wochenstunden=8, von=1, bis=12, af5=20, res=0.1))
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
        p.jahre[2026].personen.append(P.Person(name="X", wochenstunden=4, res=0.1))
    assert any("A4-Seite" in w for w in core.warnungen(p, core.compute(p)))
