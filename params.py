"""Parameter-Modell: alles, was die Auswertung beeinflusst, an einer Stelle.

Quelle der Standardwerte ist die Steuerungsdatei (XLSX) in der Nextcloud.
Jeder Wert lässt sich im Formular einzeln überschreiben; ``from_form`` und
``to_form`` sind die Brücke zwischen HTTP und Modell.

Grundsätze:
* Keine Rechnung hier — nur lesen, prüfen, halten.
* Zahlen aus Formularen dürfen Komma oder Punkt als Dezimaltrenner haben.
"""
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass, field, replace
from io import BytesIO
from typing import Iterable

import openpyxl

# Voreinstellungen, die (noch) in keiner Spalte der Steuerungsdatei stehen.
# Sie landen beim ersten Sichern im Blatt "Texte" der XLSX.
STANDARD_TEXTE = {
    "titel": "Angebotsstunden",
    "org": "Medienkompetenzzentrum CIA Spandau",
    "traeger": "Staakkato Kinder und Jugend e.V.",
}

STANDARD_JAHR = {
    2026: {
        "marker": "IST",
        "untertitel": "Wer arbeitet wie viel woran — Aufstellung für das Jugendamt",
    },
    2027: {
        "marker": "PLANUNG",
        "untertitel": "Was sich ändert — und wie das Vertragssoll auch 2027 gedeckt wird",
    },
}

MAX_ZEILEN = 20
JAHR_MIN, JAHR_MAX = 2020, 2100


def zahl(v, default: float = 0.0) -> float:
    """Formular- oder Zellwert zu float. '1.234,5' und '1234.5' ergeben beide 1234.5."""
    if v is None or v == "":
        return float(default)
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("\u00a0", "").replace(" ", "")
    s = s.replace("€", "").replace("%", "").replace("h", "").replace("+", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    elif re.match(r"^-?\d{1,3}(\.\d{3})+$", s):
        # "6.828" ist im deutschen Formular eine Tausenderangabe, keine 6,828
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return float(default)


def ganz(v, default: int = 0) -> int:
    return int(round(zahl(v, default)))


def monat(v, default: int = 1) -> int:
    m = ganz(v, default)
    return min(12, max(1, m))


@dataclass
class Person:
    name: str = ""
    rolle: str = ""
    wochenstunden: float = 0.0
    von: int = 1
    bis: int = 12
    af5: float = 0.0      # gewünschter AF-5-Anteil in Stunden (wird auf das Soll skaliert)
    res: float = 0.0      # Reserve-Anteil 0..1; 1 = volle Stelle nur Leitung/Verwaltung

    @property
    def fix(self) -> bool:
        return self.res >= 0.999

    def zeitraum(self, jahr: int) -> str:
        if self.von == 1 and self.bis == 12:
            return "01.01.–31.12."
        letzter = calendar.monthrange(jahr, self.bis)[1]
        return f"01.{self.von:02d}.–{letzter:02d}.{self.bis:02d}."

    def anteil(self) -> float:
        return (self.bis - self.von + 1) / 12


@dataclass
class Jahr:
    jahr: int
    honorarstunden: float = 0.0
    workshoptage: float = 0.0
    marker: str = ""
    untertitel: str = ""
    hinweis: str = ""
    personen: list[Person] = field(default_factory=list)

    def __post_init__(self):
        if not self.marker:
            self.marker = STANDARD_JAHR.get(self.jahr, {}).get("marker", "PLANUNG")
        if not self.untertitel:
            self.untertitel = STANDARD_JAHR.get(self.jahr, {}).get("untertitel", "")


@dataclass
class Params:
    # Annahmen
    werktage: float = 261
    feiertage: float = 8
    urlaub: float = 24
    krank: float = 10
    arbeitstage_fix: bool = False
    arbeitstage_wert: float = 219

    # Vertrag
    soll: float = 6828
    soll_af1: float = 5000
    soll_af5: float = 1828
    satz: float = 38.08

    # Kopfzeilen
    titel: str = STANDARD_TEXTE["titel"]
    org: str = STANDARD_TEXTE["org"]
    traeger: str = STANDARD_TEXTE["traeger"]

    jahre: dict[int, Jahr] = field(default_factory=dict)

    # Herkunft der Standardwerte (nur Anzeige)
    quelle: str = ""
    stand: str = ""

    # ---------- abgeleitet ----------

    @property
    def arbeitstage(self) -> float:
        """Arbeitstage: automatisch aus Werktagen − Feiertagen − Urlaub − Krankheit,
        oder der von Hand gesetzte Wert, wenn ``arbeitstage_fix`` gesetzt ist."""
        if self.arbeitstage_fix:
            return self.arbeitstage_wert
        return max(0.0, self.werktage - self.feiertage - self.urlaub - self.krank)

    @property
    def wochen(self) -> float:
        return self.arbeitstage / 5

    def jahr_liste(self) -> list[int]:
        return sorted(self.jahre)

    def ist_jahr(self, jahr: int) -> bool:
        return jahr in self.jahre

    # ---------- Serialisierung ----------

    def to_form(self) -> dict[str, str]:
        """Alles, was das Formular braucht, als flache Zeichenketten."""
        d = {
            "w": fmt(self.werktage, 1), "f": fmt(self.feiertage, 1), "u": fmt(self.urlaub, 1),
            "k": fmt(self.krank, 1),
            "at_fix": "1" if self.arbeitstage_fix else "",
            "at": fmt(self.arbeitstage_wert if self.arbeitstage_fix else self.arbeitstage, 1),
            "soll": fmt(self.soll), "af1": fmt(self.soll_af1), "af5": fmt(self.soll_af5),
            "satz": fmt(self.satz, 2),
            "titel": self.titel, "org": self.org, "traeger": self.traeger,
        }
        for jahr, j in self.jahre.items():
            k = f"p{jahr}"
            d[f"h{k[-2:]}"] = fmt(j.honorarstunden, 1)
            d[f"wt{k[-2:]}"] = fmt(j.workshoptage, 1)
            d[f"m{k[-2:]}"] = j.marker
            d[f"un{k[-2:]}"] = j.untertitel
            d[f"hn{k[-2:]}"] = j.hinweis
            d[f"n{k[-2:]}"] = str(len(j.personen))
            for i, p in enumerate(j.personen):
                b = f"{k}_{i}"
                d[f"{b}_name"] = p.name
                d[f"{b}_rolle"] = p.rolle
                d[f"{b}_ws"] = fmt(p.wochenstunden, 1)
                d[f"{b}_von"] = str(p.von)
                d[f"{b}_bis"] = str(p.bis)
                d[f"{b}_af5"] = fmt(p.af5, 1)
                d[f"{b}_res"] = fmt(p.res, 3)
        return d

    def to_query(self) -> str:
        from urllib.parse import urlencode

        return urlencode(self.to_form())


def fmt(v: float, stellen: int = 0) -> str:
    """Zahl für ein Eingabefeld: ohne unnötige Nullen, mit Komma."""
    if v is None:
        return ""
    r = round(float(v), stellen)
    if r == int(r):
        return str(int(r))
    s = f"{r:.{stellen}f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


# ---------------------------------------------------------------- aus XLSX

def _zelle(ws, adresse, default=None):
    v = ws[adresse].value
    return default if v is None else v


def _honorar(ws_auswertung) -> float:
    """Honorarstunden stehen im Block der Auswertung (Zeile „Honorarkräfte", Spalte D).
    Die Zeile wird über ihre Beschriftung gesucht, damit sie nach dem Sichern
    an beliebiger Stelle stehen darf."""
    for r in range(1, ws_auswertung.max_row + 1):
        label = ws_auswertung.cell(row=r, column=1).value
        if isinstance(label, str) and label.strip().lower().startswith("honorar"):
            return zahl(ws_auswertung.cell(row=r, column=4).value)
    return 0.0


def _quellen(ws) -> dict[str, float]:
    """Blatt Quellen: Beschriftung in A, Zahl in B."""
    d = {}
    for r in range(1, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        if isinstance(label, str):
            d[label.strip()] = zahl(ws.cell(row=r, column=2).value)
    return d


def from_xlsx(data: bytes) -> Params:
    wb = openpyxl.load_workbook(BytesIO(data), data_only=False)
    if "Annahmen" not in wb.sheetnames:
        raise ValueError("Blatt 'Annahmen' fehlt in der Steuerungsdatei")
    an = wb["Annahmen"]
    p = Params(
        krank=zahl(_zelle(an, "B4", 10)),
        arbeitstage_wert=zahl(_zelle(an, "B5", 219)),
        soll=zahl(_zelle(an, "B7", 6828)),
        soll_af1=zahl(_zelle(an, "B8", 5000)),
        soll_af5=zahl(_zelle(an, "B9", 1828)),
        satz=zahl(_zelle(an, "B10", 38.08), 38.08),
    )
    q = _quellen(wb["Quellen"]) if "Quellen" in wb.sheetnames else {}
    p.werktage = q.get("Werktage Mo–Fr 2026") or q.get("Werktage Mo-Fr 2026") or 261
    p.feiertage = q.get("Feiertage auf Werktagen 2026") or 8
    p.urlaub = q.get("Urlaubsanspruch") or 24
    # Arbeitstage automatisch, solange die gespeicherte Zahl zur Rechnung passt
    auto = max(0.0, p.werktage - p.feiertage - p.urlaub - p.krank)
    p.arbeitstage_fix = abs(auto - p.arbeitstage_wert) > 0.05
    if p.arbeitstage_fix:
        p.arbeitstage_wert = zahl(_zelle(an, "B5", auto))
    else:
        p.arbeitstage_wert = auto

    texte = {}
    if "Texte" in wb.sheetnames:
        wt = wb["Texte"]
        for r in range(1, wt.max_row + 1):
            k = wt.cell(row=r, column=1).value
            v = wt.cell(row=r, column=2).value
            if isinstance(k, str) and v not in (None, ""):
                texte[k.strip()] = str(v)
    p.titel = texte.get("titel", p.titel)
    p.org = texte.get("org", p.org)
    p.traeger = texte.get("traeger", p.traeger)

    for jahr in (2026, 2027):
        blatt = f"Eingabe {jahr}"
        if blatt not in wb.sheetnames:
            continue
        ws = wb[blatt]
        personen: list[Person] = []
        r = 4
        while ws.cell(row=r, column=1).value:
            name = str(ws.cell(row=r, column=1).value).strip()
            if not name:
                break
            personen.append(Person(
                name=name,
                rolle=str(ws.cell(row=r, column=2).value or "").strip(),
                wochenstunden=zahl(ws.cell(row=r, column=3).value),
                von=monat(ws.cell(row=r, column=4).value, 1),
                bis=monat(ws.cell(row=r, column=5).value, 12),
                af5=zahl(ws.cell(row=r, column=6).value),
                res=zahl(ws.cell(row=r, column=7).value),
            ))
            r += 1
        j = Jahr(
            jahr=jahr,
            honorarstunden=_honorar(wb[f"Auswertung {jahr}"]) if f"Auswertung {jahr}" in wb.sheetnames else 0.0,
            workshoptage=q.get(f"Workshoptage {jahr}", 0.0),
            marker=texte.get(f"marker {jahr}", ""),
            untertitel=texte.get(f"untertitel {jahr}", ""),
            hinweis=texte.get(f"hinweis {jahr}", ""),
            personen=personen,
        )
        p.jahre[jahr] = j
    if not p.jahre:
        raise ValueError("Keine Blätter 'Eingabe <Jahr>' gefunden")
    p.quelle = texte.get("quelle", "")
    return p


# ---------------------------------------------------------------- aus HTTP

def from_form(form) -> Params:
    """Baut Parameter aus Formular-/Query-Daten. Fehlende Felder = Standard."""
    werte = dict(form)
    base = Params()
    p = Params(
        werktage=zahl(werte.get("w"), base.werktage),
        feiertage=zahl(werte.get("f"), base.feiertage),
        urlaub=zahl(werte.get("u"), base.urlaub),
        krank=zahl(werte.get("k"), base.krank),
        arbeitstage_fix=bool(str(werte.get("at_fix") or "").strip()),
        arbeitstage_wert=zahl(werte.get("at"), 0),
        soll=zahl(werte.get("soll"), base.soll),
        soll_af1=zahl(werte.get("af1"), base.soll_af1),
        soll_af5=zahl(werte.get("af5"), base.soll_af5),
        satz=zahl(werte.get("satz"), base.satz),
        titel=str(werte.get("titel") or base.titel),
        org=str(werte.get("org") or base.org),
        traeger=str(werte.get("traeger") or base.traeger),
    )
    if p.arbeitstage_fix and p.arbeitstage_wert <= 0:
        p.arbeitstage_wert = round(p.werktage - p.feiertage - p.urlaub - p.krank, 1)
    if p.arbeitstage_fix is False:
        p.arbeitstage_wert = round(p.werktage - p.feiertage - p.urlaub - p.krank, 1)

    for jahr in (2026, 2027):
        k = str(jahr)[-2:]
        n = ganz(werte.get(f"n{k}"), 0)
        if n <= 0 and not any(f"p{jahr}_" in s for s in werte):
            n = 0
        n = min(n, MAX_ZEILEN)
        personen = []
        for i in range(n):
            b = f"p{jahr}_{i}"
            name = str(werte.get(f"{b}_name") or "").strip()
            if not name:
                continue
            personen.append(Person(
                name=name,
                rolle=str(werte.get(f"{b}_rolle") or "").strip(),
                wochenstunden=zahl(werte.get(f"{b}_ws")),
                von=monat(werte.get(f"{b}_von"), 1),
                bis=monat(werte.get(f"{b}_bis"), 12),
                af5=zahl(werte.get(f"{b}_af5")),
                res=zahl(werte.get(f"{b}_res")),
            ))
        p.jahre[jahr] = Jahr(
            jahr=jahr,
            honorarstunden=zahl(werte.get(f"h{k}")),
            workshoptage=zahl(werte.get(f"wt{k}")),
            marker=str(werte.get(f"m{k}") or STANDARD_JAHR.get(jahr, {}).get("marker", "")),
            untertitel=str(werte.get(f"un{k}") or STANDARD_JAHR.get(jahr, {}).get("untertitel", "")),
            hinweis=str(werte.get(f"hn{k}") or ""),
            personen=personen,
        )
    p.quelle = str(werte.get("quelle") or "")
    p.stand = str(werte.get("stand") or "")
    return p


SCHLUESSEL = ("w", "f", "u", "k", "at_fix", "at", "soll", "af1", "af5", "satz",
              "titel", "org", "traeger", "h26", "h27", "wt26", "wt27", "n26", "n27")


def betroffen(form) -> bool:
    """Wurde mindestens ein bekannter Parameter übergeben?"""
    return any(k in form for k in SCHLUESSEL)


def overlay(base: Params, form) -> Params:
    """Legt nur die tatsächlich übergebenen Felder über die Werte aus der XLSX.

    Damit funktioniert auch ein einzelner Parameter in der URL (`?k=15`) — und
    fehlende Felder fallen auf die Steuerungsdatei zurück statt auf Python-Konstanten.
    """
    w = dict(form)

    def pick(key: str, current, cast=zahl):
        if key not in w or str(w[key]).strip() == "":
            return current
        return cast(w[key])

    p = replace(
        base,
        werktage=pick("w", base.werktage),
        feiertage=pick("f", base.feiertage),
        urlaub=pick("u", base.urlaub),
        krank=pick("k", base.krank),
        soll=pick("soll", base.soll),
        soll_af1=pick("af1", base.soll_af1),
        soll_af5=pick("af5", base.soll_af5),
        satz=pick("satz", base.satz),
        titel=str(w.get("titel") or base.titel),
        org=str(w.get("org") or base.org),
        traeger=str(w.get("traeger") or base.traeger),
    )
    if "at_fix" in w:
        p.arbeitstage_fix = bool(str(w.get("at_fix") or "").strip())
    if "at" in w and str(w["at"]).strip():
        p.arbeitstage_wert = zahl(w["at"])
    elif not p.arbeitstage_fix:
        p.arbeitstage_wert = round(p.werktage - p.feiertage - p.urlaub - p.krank, 1)

    jahre = {}
    for jahr, j in base.jahre.items():
        s = str(jahr)[-2:]
        if f"n{s}" in w:
            n = min(ganz(w[f"n{s}"], 0), MAX_ZEILEN)
            personen = []
            for i in range(n):
                b = f"p{jahr}_{i}"
                name = str(w.get(f"{b}_name") or "").strip()
                if not name:
                    continue
                personen.append(Person(
                    name=name,
                    rolle=str(w.get(f"{b}_rolle") or "").strip(),
                    wochenstunden=zahl(w.get(f"{b}_ws")),
                    von=monat(w.get(f"{b}_von"), 1),
                    bis=monat(w.get(f"{b}_bis"), 12),
                    af5=zahl(w.get(f"{b}_af5")),
                    res=zahl(w.get(f"{b}_res")),
                ))
        else:
            personen = j.personen
        jahre[jahr] = Jahr(
            jahr=jahr,
            honorarstunden=pick(f"h{s}", j.honorarstunden),
            workshoptage=pick(f"wt{s}", j.workshoptage),
            marker=str(w.get(f"m{s}") or j.marker),
            untertitel=str(w.get(f"un{s}") or j.untertitel),
            hinweis=str(w[f"hn{s}"]) if f"hn{s}" in w else j.hinweis,
            personen=personen,
        )
    p.jahre = jahre
    return p


def modify(p: Params, **werte) -> Params:
    return replace(p, **werte)
