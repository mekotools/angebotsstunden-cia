"""Parameter-Modell: alles, was die Auswertung beeinflusst, an einer Stelle.

Quelle der Standardwerte ist die Steuerungsdatei (XLSX) in der Nextcloud.
Jeder Wert lässt sich im Formular einzeln überschreiben; ``from_form`` und
``to_form`` sind die Brücke zwischen HTTP und Modell.

Zwei Dinge sind hier neu gegenüber einer reinen Jahrestabelle:

* **Basisdaten je Bundesland und Jahr** kommen aus ``basisdaten`` (Kalender,
  keine abgeschriebene Tabelle). Werktage, Feiertage und Arbeitstage stehen
  damit für jedes Jahr fest und werden nicht von Hand gepflegt.
* **Zeiträume je Person.** Eine Person hat einen oder mehrere Abschnitte
  (Monat von … bis) mit eigenen Wochenstunden, Urlaubs- und Krankheitstagen —
  damit lässt sich abbilden, dass jemand im Juli seine Stunden erhöht.
  Leere Urlaubs-/Krankheitsfelder eines Abschnitts erben anteilig den
  Jahresanspruch.

Grundsätze:
* Keine Rechnung hier — nur lesen, prüfen, halten.
* Zahlen aus Formularen dürfen Komma oder Punkt als Dezimaltrenner haben.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field, replace
from io import BytesIO
from typing import Iterable

import openpyxl

import basisdaten

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

MAX_ZEILEN = 40
JAHR_MIN, JAHR_MAX = 2020, 2100
STANDARD_LAND = "Berlin"


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


def optional_zahl(v) -> float | None:
    """Leeres Feld = ``None`` (anteilig), sonst die Zahl."""
    if v is None or str(v).strip() == "":
        return None
    return zahl(v)


@dataclass
class Abschnitt:
    """Ein Beschäftigungszeitraum: Monate ``von``..``bis`` mit eigenen Werten."""

    von: int = 1
    bis: int = 12
    wochenstunden: float = 0.0
    urlaub: float | None = None   # None = anteilig aus dem Jahresanspruch
    krank: float | None = None    # None = anteilig aus dem Jahresanspruch
    af5: float = 0.0              # gewünschter AF-5-Anteil in Stunden (wird auf das Soll skaliert)
    res: float = 0.0              # Reserve-Anteil 0..1; 1 = volle Stelle nur Leitung/Verwaltung

    def __post_init__(self):
        self.von = monat(self.von, 1)
        self.bis = monat(self.bis, 12)
        if self.bis < self.von:
            self.von, self.bis = self.bis, self.von

    @property
    def monate(self) -> int:
        return self.bis - self.von + 1

    @property
    def anteil(self) -> float:
        """Anteil am Jahr, für die anteilige Vererbung von Urlaub und Krankheit."""
        return self.monate / 12

    @property
    def fix(self) -> bool:
        return self.res >= 0.999

    def zeitraum(self) -> str:
        if self.von == 1 and self.bis == 12:
            return "01.–12."
        return f"{self.von:02d}.–{self.bis:02d}."


@dataclass
class Person:
    """Eine Person mit einem oder mehreren Abschnitten."""

    name: str = ""
    rolle: str = ""
    abschnitte: list[Abschnitt] = field(default_factory=list)

    def __post_init__(self):
        if not self.abschnitte:
            self.abschnitte = [Abschnitt()]

    # ---- Bequemlichkeit für Anzeige und alte Aufrufer ----

    @property
    def wochenstunden(self) -> float:
        return self.abschnitte[0].wochenstunden if self.abschnitte else 0.0

    @property
    def ws_text(self) -> str:
        werte = [a.wochenstunden for a in self.abschnitte]
        if len(werte) == 1:
            return fmt(werte[0], 1)
        return " → ".join(fmt(w, 1) for w in werte)

    @property
    def af5(self) -> float:
        return sum(a.af5 for a in self.abschnitte)

    @property
    def res(self) -> float:
        return self.abschnitte[0].res if self.abschnitte else 0.0

    @property
    def fix(self) -> bool:
        return any(a.fix for a in self.abschnitte)

    @property
    def zeitraum_text(self) -> str:
        if len(self.abschnitte) == 1:
            return self.abschnitte[0].zeitraum()
        return " · ".join(a.zeitraum() for a in self.abschnitte)


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
    # Basisdaten
    land: str = STANDARD_LAND
    urlaub: float = 24.0      # Jahresanspruch, Vorgabe für Abschnitte ohne eigene Angabe
    krank: float = 10.0       # Jahresannahme, ebenso

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

    _basis: dict = field(default_factory=dict, repr=False)

    # ---------- Basisdaten ----------

    def basis(self, jahr: int) -> dict:
        """Werktage, Feiertage, Arbeitstage und Monatstabelle eines Jahres."""
        schluessel = (self.land, jahr)
        if schluessel not in self._basis:
            self._basis[schluessel] = basisdaten.jahr(self.land, jahr)
        return self._basis[schluessel]

    def arbeitstage(self, jahr: int) -> float:
        return float(self.basis(jahr)["arbeitstage"])

    def netto_tage(self, jahr: int, a: Abschnitt) -> float:
        """Arbeitstage eines Abschnitts nach Abzug von Urlaub und Krankheit."""
        at = basisdaten.arbeitstage(self.land, jahr, a.von, a.bis)
        u = a.urlaub if a.urlaub is not None else self.urlaub * a.anteil
        k = a.krank if a.krank is not None else self.krank * a.anteil
        return max(0.0, at - u - k)

    def jahr_liste(self) -> list[int]:
        return sorted(self.jahre)

    def ist_jahr(self, jahr: int) -> bool:
        return jahr in self.jahre

    # ---------- Serialisierung ----------

    def to_form(self) -> dict[str, str]:
        """Alles, was das Formular braucht, als flache Zeichenketten.

        Die Tabelle im Formular hat **eine Zeile je Abschnitt**; ``n<jj>`` ist
        entsprechend die Zahl der Abschnittszeilen, nicht die der Personen.
        """
        d = {
            "land": self.land,
            "u": fmt(self.urlaub, 1), "k": fmt(self.krank, 1),
            "soll": fmt(self.soll), "af1": fmt(self.soll_af1), "af5": fmt(self.soll_af5),
            "satz": fmt(self.satz, 2),
            "titel": self.titel, "org": self.org, "traeger": self.traeger,
        }
        for jahr, j in self.jahre.items():
            s = str(jahr)[-2:]
            b = self.basis(jahr)
            d[f"bd{s}_w"] = fmt(b["werktage"])
            d[f"bd{s}_f"] = fmt(b["feiertage_werktags"])
            d[f"bd{s}_at"] = fmt(b["arbeitstage"])
            d[f"h{s}"] = fmt(j.honorarstunden, 1)
            d[f"wt{s}"] = fmt(j.workshoptage, 1)
            d[f"m{s}"] = j.marker
            d[f"un{s}"] = j.untertitel
            d[f"hn{s}"] = j.hinweis
            zeilen = [(q, a) for q in j.personen for a in q.abschnitte]
            d[f"n{s}"] = str(len(zeilen))
            for i, (q, a) in enumerate(zeilen):
                r = f"p{jahr}_{i}"
                d[f"{r}_name"] = q.name
                d[f"{r}_rolle"] = q.rolle
                d[f"{r}_von"] = str(a.von)
                d[f"{r}_bis"] = str(a.bis)
                d[f"{r}_ws"] = fmt(a.wochenstunden, 1)
                d[f"{r}_u"] = "" if a.urlaub is None else fmt(a.urlaub, 1)
                d[f"{r}_k"] = "" if a.krank is None else fmt(a.krank, 1)
                d[f"{r}_af5"] = fmt(a.af5, 1)
                d[f"{r}_res"] = fmt(a.res, 3)
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
    for r in range(1, ws_auswertung.max_row + 1):
        label = ws_auswertung.cell(row=r, column=1).value
        if isinstance(label, str) and label.strip().lower().startswith("honorar"):
            return zahl(ws_auswertung.cell(row=r, column=4).value)
    return 0.0


def _quellen(ws) -> dict[str, str]:
    """Blatt Quellen: Beschriftung in A, Wert in B (Zahl oder Text)."""
    d = {}
    for r in range(1, ws.max_row + 1):
        label = ws.cell(row=r, column=1).value
        wert = ws.cell(row=r, column=2).value
        if isinstance(label, str):
            d[label.strip()] = wert
    return d


def _personen_aus_blatt(ws, alt: bool, urlaub_jahr: float, krank_jahr: float) -> list[Person]:
    """Liest das Blatt 'Eingabe <Jahr>'.

    Neues Format: A Name, B Funktion, C von, D bis, E Wochenstunden,
    F Urlaub, G Krankheit, H AF 5, I Reserve-Anteil.
    Altes Format: A Name, B Funktion, C Wochenstunden, D von, E bis,
    F AF 5, G Reserve-Anteil — Urlaub/Krankheit kamen aus den Annahmen.
    """
    personen: list[Person] = []
    index: dict[str, Person] = {}
    r = 4
    while ws.cell(row=r, column=1).value:
        name = str(ws.cell(row=r, column=1).value).strip()
        if not name:
            break
        if alt:
            a = Abschnitt(
                von=monat(ws.cell(row=r, column=4).value, 1),
                bis=monat(ws.cell(row=r, column=5).value, 12),
                wochenstunden=zahl(ws.cell(row=r, column=3).value),
                af5=zahl(ws.cell(row=r, column=6).value),
                res=zahl(ws.cell(row=r, column=7).value),
            )
        else:
            a = Abschnitt(
                von=monat(ws.cell(row=r, column=3).value, 1),
                bis=monat(ws.cell(row=r, column=4).value, 12),
                wochenstunden=zahl(ws.cell(row=r, column=5).value),
                urlaub=optional_zahl(ws.cell(row=r, column=6).value),
                krank=optional_zahl(ws.cell(row=r, column=7).value),
                af5=zahl(ws.cell(row=r, column=8).value),
                res=zahl(ws.cell(row=r, column=9).value),
            )
        if name in index:
            index[name].abschnitte.append(a)
        else:
            index[name] = Person(name=name,
                                 rolle=str(ws.cell(row=r, column=2).value or "").strip(),
                                 abschnitte=[a])
            personen.append(index[name])
        r += 1
    return personen


def _ist_neues_format(ws) -> bool:
    kopf = str(ws.cell(row=3, column=3).value or "").strip().lower()
    return kopf.startswith("von")


def from_xlsx(data: bytes) -> Params:
    wb = openpyxl.load_workbook(BytesIO(data), data_only=False)
    if "Annahmen" not in wb.sheetnames:
        raise ValueError("Blatt 'Annahmen' fehlt in der Steuerungsdatei")
    an = wb["Annahmen"]
    q = _quellen(wb["Quellen"]) if "Quellen" in wb.sheetnames else {}
    land = str(q.get("Bundesland") or STANDARD_LAND).strip()
    if land not in basisdaten.BUNDESLAENDER:
        land = STANDARD_LAND
    p = Params(
        land=land,
        urlaub=zahl(q.get("Urlaubsanspruch"), 24),
        krank=zahl(_zelle(an, "B4", 10)),
        soll=zahl(_zelle(an, "B7", 6828)),
        soll_af1=zahl(_zelle(an, "B8", 5000)),
        soll_af5=zahl(_zelle(an, "B9", 1828)),
        satz=zahl(_zelle(an, "B10", 38.08), 38.08),
    )

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

    for jahr in sorted({int(s.split()[-1]) for s in wb.sheetnames
                        if s.startswith("Eingabe ") and s.split()[-1].isdigit()}):
        ws = wb[f"Eingabe {jahr}"]
        personen = _personen_aus_blatt(ws, alt=not _ist_neues_format(ws),
                                       urlaub_jahr=p.urlaub, krank_jahr=p.krank)
        j = Jahr(
            jahr=jahr,
            honorarstunden=_honorar(wb[f"Auswertung {jahr}"])
            if f"Auswertung {jahr}" in wb.sheetnames else 0.0,
            workshoptage=zahl(q.get(f"Workshoptage {jahr}"), 0),
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

def _personen_aus_form(werte, jahr: int, urlaub_jahr: float, krank_jahr: float) -> list[Person]:
    """Formularzeilen → Personen. Gleicher Name in Folgezeilen = gleiche Person."""
    s = str(jahr)[-2:]
    n = min(ganz(werte.get(f"n{s}"), 0), MAX_ZEILEN)
    personen: list[Person] = []
    index: dict[str, Person] = {}
    for i in range(n):
        r = f"p{jahr}_{i}"
        name = str(werte.get(f"{r}_name") or "").strip()
        if not name:
            continue
        a = Abschnitt(
            von=monat(werte.get(f"{r}_von"), 1),
            bis=monat(werte.get(f"{r}_bis"), 12),
            wochenstunden=zahl(werte.get(f"{r}_ws")),
            urlaub=optional_zahl(werte.get(f"{r}_u")),
            krank=optional_zahl(werte.get(f"{r}_k")),
            af5=zahl(werte.get(f"{r}_af5")),
            res=zahl(werte.get(f"{r}_res")),
        )
        if name in index:
            index[name].abschnitte.append(a)
        else:
            index[name] = Person(name=name, rolle=str(werte.get(f"{r}_rolle") or "").strip(),
                                 abschnitte=[a])
            personen.append(index[name])
    return personen


def from_form(form) -> Params:
    """Baut Parameter aus Formular-/Query-Daten. Fehlende Felder = Standard."""
    werte = dict(form)
    base = Params()
    land = str(werte.get("land") or base.land).strip()
    if land not in basisdaten.BUNDESLAENDER:
        land = base.land
    p = Params(
        land=land,
        urlaub=zahl(werte.get("u"), base.urlaub),
        krank=zahl(werte.get("k"), base.krank),
        soll=zahl(werte.get("soll"), base.soll),
        soll_af1=zahl(werte.get("af1"), base.soll_af1),
        soll_af5=zahl(werte.get("af5"), base.soll_af5),
        satz=zahl(werte.get("satz"), base.satz),
        titel=str(werte.get("titel") or base.titel),
        org=str(werte.get("org") or base.org),
        traeger=str(werte.get("traeger") or base.traeger),
    )
    for jahr in (2026, 2027):
        s = str(jahr)[-2:]
        p.jahre[jahr] = Jahr(
            jahr=jahr,
            honorarstunden=zahl(werte.get(f"h{s}")),
            workshoptage=zahl(werte.get(f"wt{s}")),
            marker=str(werte.get(f"m{s}") or STANDARD_JAHR.get(jahr, {}).get("marker", "")),
            untertitel=str(werte.get(f"un{s}") or STANDARD_JAHR.get(jahr, {}).get("untertitel", "")),
            hinweis=str(werte.get(f"hn{s}") or ""),
            personen=_personen_aus_form(werte, jahr, p.urlaub, p.krank),
        )
    p.quelle = str(werte.get("quelle") or "")
    p.stand = str(werte.get("stand") or "")
    return p


SCHLUESSEL = ("land", "u", "k", "soll", "af1", "af5", "satz",
              "titel", "org", "traeger", "h26", "h27", "wt26", "wt27", "n26", "n27")


def betroffen(form) -> bool:
    """Wurde mindestens ein bekannter Parameter übergeben?"""
    return any(k in form for k in SCHLUESSEL)


def overlay(base: Params, form) -> Params:
    """Legt nur die tatsächlich übergebenen Felder über die Werte aus der XLSX.

    Ein einzelner Parameter in der URL (`?k=15`) wirkt damit, ohne die Personen
    zu verlieren: ``u`` und ``k`` gelten dann anteilig für alle Abschnitte.
    Kommen Personen-Zeilen mit, gewinnen deren eigene Werte.
    """
    w = dict(form)

    def pick(key: str, current, cast=zahl):
        if key not in w or str(w[key]).strip() == "":
            return current
        return cast(w[key])

    land = str(w.get("land") or base.land).strip()
    if land not in basisdaten.BUNDESLAENDER:
        land = base.land
    p = replace(
        base, land=land, _basis={},
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

    jahre = {}
    for jahr, j in base.jahre.items():
        s = str(jahr)[-2:]
        if f"n{s}" in w:
            personen = _personen_aus_form(w, jahr, p.urlaub, p.krank)
        else:
            personen = j.personen
        # URL-Einzelwerte für Urlaub/Krankheit wirken auf alle Abschnitte anteilig
        if "u" in w or "k" in w:
            neu = []
            for q in personen:
                abschnitte = []
                for a in q.abschnitte:
                    abschnitte.append(replace(
                        a,
                        urlaub=None if "u" in w else a.urlaub,
                        krank=None if "k" in w else a.krank))
                neu.append(replace(q, abschnitte=abschnitte))
            personen = neu
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
