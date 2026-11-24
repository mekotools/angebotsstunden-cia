"""Basisdaten: Arbeitstage, Werktage und Feiertage je Bundesland und Jahr.

Keine Tabelle, die abgeschrieben wird, sondern die Rechnung aus dem Kalender:

* Werktage = Tage Montag–Freitag im Zeitraum
* Feiertage je Land aus dem Landesrecht (bewegliche Feste über die Osterformel,
  Einmal-Feiertage als eigene Liste)
* Arbeitstage = Werktage − Feiertage, die auf einen Werktag fallen

Damit ist jedes Jahr und jedes Land reproduzierbar, ohne dass jemand Zahlen
pflegt. Die Tabelle im Netz (z. B. schnelle-online.info) dient nur als
Gegenprobe — tests/test_basisdaten.py prüft dagegen.
"""
from __future__ import annotations

import calendar
import datetime as dt

BUNDESLAENDER = [
    "Baden-Württemberg", "Bayern", "Berlin", "Brandenburg", "Bremen", "Hamburg",
    "Hessen", "Mecklenburg-Vorpommern", "Niedersachsen", "Nordrhein-Westfalen",
    "Rheinland-Pfalz", "Saarland", "Sachsen", "Sachsen-Anhalt",
    "Schleswig-Holstein", "Thüringen",
]

# Kürzel für die Anzeige (Kopfzeilen, PDF)
KUERZEL = {
    "Baden-Württemberg": "BW", "Bayern": "BY", "Berlin": "BE", "Brandenburg": "BB",
    "Bremen": "HB", "Hamburg": "HH", "Hessen": "HE", "Mecklenburg-Vorpommern": "MV",
    "Niedersachsen": "NI", "Nordrhein-Westfalen": "NW", "Rheinland-Pfalz": "RP",
    "Saarland": "SL", "Sachsen": "SN", "Sachsen-Anhalt": "ST",
    "Schleswig-Holstein": "SH", "Thüringen": "TH",
}

# Überall gesetzlich (§ 2 Abs. 3 Feiertagsgesetze), zusätzlich die landesspezifischen.
_BUNDESWEIT = {
    "Neujahr", "Karfreitag", "Ostermontag", "Tag der Arbeit", "Christi Himmelfahrt",
    "Pfingstmontag", "Tag der Deutschen Einheit", "1. Weihnachtstag", "2. Weihnachtstag",
}

# Land → zusätzliche Feiertage. "e" = bewegliches Fest (Tage nach Ostern),
# "m-t" = festes Datum (Monat-Tag).
_LAND = {
    "Baden-Württemberg": {"Heilige Drei Könige": ("m", 1, 6), "Fronleichnam": ("e", 60),
                          "Allerheiligen": ("m", 11, 1)},
    "Bayern": {"Heilige Drei Könige": ("m", 1, 6), "Fronleichnam": ("e", 60),
               "Mariä Himmelfahrt": ("m", 8, 15), "Allerheiligen": ("m", 11, 1)},
    "Berlin": {"Internationaler Frauentag": ("m", 3, 8)},
    "Brandenburg": {"Ostersonntag": ("e", 0), "Pfingstsonntag": ("e", 49),
                    "Reformationstag": ("m", 10, 31)},
    "Bremen": {"Reformationstag": ("m", 10, 31)},
    "Hamburg": {"Reformationstag": ("m", 10, 31)},
    "Hessen": {"Fronleichnam": ("e", 60)},
    "Mecklenburg-Vorpommern": {"Internationaler Frauentag": ("m", 3, 8),
                               "Reformationstag": ("m", 10, 31)},
    "Niedersachsen": {"Reformationstag": ("m", 10, 31)},
    "Nordrhein-Westfalen": {"Fronleichnam": ("e", 60), "Allerheiligen": ("m", 11, 1)},
    "Rheinland-Pfalz": {"Fronleichnam": ("e", 60), "Allerheiligen": ("m", 11, 1)},
    "Saarland": {"Fronleichnam": ("e", 60), "Mariä Himmelfahrt": ("m", 8, 15),
                 "Allerheiligen": ("m", 11, 1)},
    "Sachsen": {"Reformationstag": ("m", 10, 31), "Buß- und Bettag": ("buss",)},
    "Sachsen-Anhalt": {"Heilige Drei Könige": ("m", 1, 6), "Reformationstag": ("m", 10, 31)},
    "Schleswig-Holstein": {"Reformationstag": ("m", 10, 31)},
    "Thüringen": {"Weltkindertag": ("m", 9, 20), "Reformationstag": ("m", 10, 31)},
}

# Einmalige Feiertage aus besonderem Anlass (Bekanntmachung des Landes).
_EINMALIG = {
    ("Berlin", 2025): [("5-8", "Tag der Befreiung (80. Jahrestag)")],
}


def ostern(jahr: int) -> dt.date:
    """Ostersonntag nach der Gaußschen Osterformel (gregorianischer Kalender)."""
    a = jahr % 19
    b, c = divmod(jahr, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    monat, tag = divmod(h + l - 7 * m + 114, 31)
    return dt.date(jahr, monat, tag + 1)


def _buss_und_bettag(jahr: int) -> dt.date:
    """Mittwoch vor dem 23. November."""
    d = dt.date(jahr, 11, 23)
    while d.weekday() != 2:
        d -= dt.timedelta(days=1)
    return d


def feiertage(land: str, jahr: int) -> dict[dt.date, str]:
    """Alle gesetzlichen Feiertage eines Landes im Jahr (Datum → Name)."""
    if land not in _LAND:
        raise ValueError(f"Unbekanntes Bundesland: {land}")
    o = ostern(jahr)
    f: dict[dt.date, str] = {
        dt.date(jahr, 1, 1): "Neujahr",
        o - dt.timedelta(days=2): "Karfreitag",
        o + dt.timedelta(days=1): "Ostermontag",
        dt.date(jahr, 5, 1): "Tag der Arbeit",
        o + dt.timedelta(days=39): "Christi Himmelfahrt",
        o + dt.timedelta(days=50): "Pfingstmontag",
        dt.date(jahr, 10, 3): "Tag der Deutschen Einheit",
        dt.date(jahr, 12, 25): "1. Weihnachtstag",
        dt.date(jahr, 12, 26): "2. Weihnachtstag",
    }
    for name, regel in _LAND[land].items():
        art = regel[0]
        if art == "e":
            f[o + dt.timedelta(days=regel[1])] = name
        elif art == "buss":
            f[_buss_und_bettag(jahr)] = name
        else:  # ("m", Monat, Tag)
            f[dt.date(jahr, regel[1], regel[2])] = name
    for datum, name in _EINMALIG.get((land, jahr), []):
        monat, tag = (int(x) for x in datum.split("-"))
        f[dt.date(jahr, monat, tag)] = name
    return dict(sorted(f.items()))


def werktage(jahr: int, monat: int | None = None) -> int:
    """Tage Montag–Freitag, im Monat oder im ganzen Jahr."""
    if monat is None:
        return sum(werktage(jahr, m) for m in range(1, 13))
    tage = calendar.monthrange(jahr, monat)[1]
    return sum(1 for t in range(1, tage + 1)
               if dt.date(jahr, monat, t).weekday() < 5)


def monatstabelle(land: str, jahr: int) -> list[dict]:
    """Je Monat: Werktage, Feiertage auf Werktagen, Arbeitstage."""
    f = feiertage(land, jahr)
    zeilen = []
    for m in range(1, 13):
        w = werktage(jahr, m)
        fe = sum(1 for d in f if d.month == m and d.weekday() < 5)
        zeilen.append(dict(monat=m, werktage=w, feiertage=fe, arbeitstage=w - fe))
    return zeilen


def jahr(land: str, jahr: int) -> dict:
    """Basisdaten eines Jahres: Summen plus Monatstabelle."""
    tab = monatstabelle(land, jahr)
    f = feiertage(land, jahr)
    return dict(
        land=land, jahr=jahr,
        werktage=sum(z["werktage"] for z in tab),
        feiertage=len(f),
        feiertage_werktags=sum(z["feiertage"] for z in tab),
        arbeitstage=sum(z["arbeitstage"] for z in tab),
        monate=tab,
        daten=[(d.isoformat(), n) for d, n in f.items()],
    )


def arbeitstage(land: str, jahr: int, von: int = 1, bis: int = 12) -> float:
    """Arbeitstage vom Monat ``von`` bis ``bis`` (einschließlich)."""
    von, bis = max(1, min(12, int(von))), max(1, min(12, int(bis)))
    if bis < von:
        return 0.0
    tab = monatstabelle(land, jahr)
    return float(sum(z["arbeitstage"] for z in tab[von - 1:bis]))
