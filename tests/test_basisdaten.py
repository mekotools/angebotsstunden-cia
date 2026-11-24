"""Basisdaten: Arbeitstage, Werktage und Feiertage je Bundesland und Jahr.

Die Sollzahlen stammen aus veröffentlichten Tabellen (schnelle-online.info,
arbeitstage.org, 5-Tage-Woche Mo–Fr) und werden hier gegen die eigene Rechnung
aus dem Kalender geprüft — nicht umgekehrt.
"""
from __future__ import annotations

import os
import sys

import pytest

HIER = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HIER))

import basisdaten as B  # noqa: E402


def test_alle_sechzehn_laender_bekannt():
    assert len(B.BUNDESLAENDER) == 16
    assert len(set(B.BUNDESLAENDER)) == 16
    for land in B.BUNDESLAENDER:
        assert len(B.feiertage(land, 2027)) >= 9, land


def test_ostern_bekannte_jahre():
    assert B.ostern(2024).isoformat() == "2024-03-31"
    assert B.ostern(2025).isoformat() == "2025-04-20"
    assert B.ostern(2026).isoformat() == "2026-04-05"
    assert B.ostern(2027).isoformat() == "2027-03-28"


@pytest.mark.parametrize("land,anzahl", [
    ("Baden-Württemberg", 12), ("Bayern", 13), ("Berlin", 10), ("Brandenburg", 12),
    ("Bremen", 10), ("Hamburg", 10), ("Hessen", 10), ("Mecklenburg-Vorpommern", 11),
    ("Niedersachsen", 10), ("Nordrhein-Westfalen", 11), ("Rheinland-Pfalz", 11),
    ("Saarland", 12), ("Sachsen", 11), ("Sachsen-Anhalt", 11),
    ("Schleswig-Holstein", 10), ("Thüringen", 11),
])
def test_feiertage_2027_wie_veroeffentlicht(land, anzahl):
    assert len(B.feiertage(land, 2027)) == anzahl


@pytest.mark.parametrize("jahr,arbeitstage", [(2025, 251), (2026, 254), (2027, 255)])
def test_berlin_arbeitstage_wie_veroeffentlicht(jahr, arbeitstage):
    d = B.jahr("Berlin", jahr)
    assert d["werktage"] == 261
    assert d["arbeitstage"] == arbeitstage


def test_einmal_feiertag_2025_berlin():
    """Der 8. Mai 2025 war in Berlin einmalig frei — ohne ihn 252 statt 251."""
    f = B.feiertage("Berlin", 2025)
    assert any(d.isoformat() == "2025-05-08" for d in f), "Einmal-Feiertag fehlt"
    assert B.jahr("Berlin", 2025)["arbeitstage"] == 251


@pytest.mark.parametrize("land,arbeitstage", [
    ("Bayern", 252), ("Baden-Württemberg", 252), ("Nordrhein-Westfalen", 253),
    ("Sachsen", 253), ("Berlin", 254),
])
def test_arbeitstage_2026_wie_veroeffentlicht(land, arbeitstage):
    assert B.jahr(land, 2026)["arbeitstage"] == arbeitstage


def test_monate_ergeben_das_jahr():
    for land in ("Berlin", "Bayern"):
        d = B.jahr(land, 2026)
        assert sum(m["werktage"] for m in d["monate"]) == d["werktage"]
        assert sum(m["arbeitstage"] for m in d["monate"]) == d["arbeitstage"]
        assert d["feiertage_werktags"] + d["arbeitstage"] == d["werktage"]


def test_zeitraum_summiert_die_monate():
    ganz = B.arbeitstage("Berlin", 2026, 1, 12)
    halb = B.arbeitstage("Berlin", 2026, 1, 6) + B.arbeitstage("Berlin", 2026, 7, 12)
    assert ganz == halb == 254
    assert B.arbeitstage("Berlin", 2026, 6, 3) == 0  # verdrehter Zeitraum
