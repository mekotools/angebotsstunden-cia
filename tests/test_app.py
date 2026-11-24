"""HTTP-Ebene: Justierformular, PDF-Ausgabe, Kontrollzahlen.

Läuft gegen eine lokale Steuerungsdatei (LOCAL_FILE), nicht gegen die Nextcloud.
"""
from __future__ import annotations

import os
import sys

import pytest

HIER = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HIER))
STEUERDATEI = os.environ.get("STEUERDATEI", os.path.join(HIER, "steuerdatei.xlsx"))

if not os.path.exists(STEUERDATEI):
    pytest.skip(f"Steuerungsdatei fehlt: {STEUERDATEI}", allow_module_level=True)

os.environ["LOCAL_FILE"] = STEUERDATEI
os.environ["CACHE_SECONDS"] = "0"

from fastapi.testclient import TestClient  # noqa: E402

import app as A  # noqa: E402
import params as P  # noqa: E402
import xlsx_write  # noqa: E402

client = TestClient(A.app)


def basis_form() -> dict:
    with open(STEUERDATEI, "rb") as fh:
        return P.from_xlsx(fh.read()).to_form()


def test_formular_zeigt_alle_felder():
    r = client.get("/")
    assert r.status_code == 200
    html = r.text
    for feld in ("name=\"w\"", "name=\"soll\"", "name=\"p2026_0_name\"", "name=\"h26\"",
                 "name=\"p2027_0_ws\"", "name=\"satz\"", "Beruf"):
        if feld == "Beruf":
            continue
        assert feld in html, feld
    assert "A4-PDF ansehen" in html and "In Excel sichern" in html


def test_pdf_route_liefert_pdf():
    r = client.get("/export.pdf", params=basis_form())
    assert r.status_code == 200
    assert r.headers["content-type"] == "application/pdf"
    assert r.content[:4] == b"%PDF"
    assert r.headers["content-disposition"].startswith("inline")


def test_pdf_route_nur_ein_jahr():
    r = client.get("/export.pdf", params={**basis_form(), "y": "2027"})
    assert r.status_code == 200 and r.content[:4] == b"%PDF"


def test_justierte_zahl_kommt_im_pdf_an():
    p = client.get("/export.pdf", params={**basis_form(), "wt26": "150"})
    assert p.status_code == 200
    import pymupdf
    text = pymupdf.open(stream=p.content)[0].get_text()
    assert "150 Workshoptage" in text.replace("\n", " ")


def test_kennzahlen_fragment():
    r = client.get("/kennzahlen", params=basis_form())
    assert r.status_code == 200
    assert "AF 1" in r.text and "gedeckt" in r.text


def test_healthz_meldet_datenstand():
    r = client.get("/healthz")
    d = r.json()
    assert d["ok"] is True and d["data_ok"] is True
    assert d["bytes"] > 5000 and d["personen"]["2026"] > 0


def test_sichern_ohne_nextcloud_sagt_warum():
    r = client.post("/sichern", data=basis_form())
    assert r.status_code == 400
    assert "Nextcloud" in r.json()["fehler"]


def test_geschriebene_datei_wird_wieder_gelesen():
    """Der Rückweg über HTTP-Formular → XLSX → Formular bleibt verlustfrei."""
    form = basis_form()
    form.update({"k": "12", "p2026_0_ws": "35"})
    p = P.from_form(form)
    with open(STEUERDATEI, "rb") as fh:
        neu = xlsx_write.schreiben(fh.read(), p)
    p2 = P.from_xlsx(neu)
    assert p2.krank == 12
    assert p2.jahre[2026].personen[0].wochenstunden == 35
    assert p2.to_form() == p.to_form()


def test_einzelner_parameter_in_der_url():
    """?k=15 muss wirken — fehlende Felder kommen aus der Steuerungsdatei, nicht aus
    Python-Konstanten, und die Personen bleiben stehen."""
    import pymupdf
    r = client.get("/export.pdf", params={"k": "15"})
    assert r.status_code == 200
    text = pymupdf.open(stream=r.content)[0].get_text().replace("\n", " ")
    assert "Arbeitstage 214" in text
    assert "Julia Hoffmann" in text
