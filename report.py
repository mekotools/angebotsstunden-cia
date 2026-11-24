"""A4-Ausgabe (PDF) des Stundengerüsts.

Eine Seite je Jahr, Aufbau: Kopf → vier Kennzahlen → Tabelle → Deckungsbalken →
Balken je Person → Herleitung im Fuß. Alle Zahlen kommen aus ``core.compute``,
alle Texte und Faktoren aus den Parametern; die Seite passt sich der Zeilenzahl an.
"""
from __future__ import annotations

from pathlib import Path

from fpdf import FPDF

PETROL = (30, 75, 90)
GREEN = (46, 125, 107)
AMBER = (201, 127, 46)
SAND = (226, 168, 100)
GREY = (198, 205, 209)
TEXT = (34, 49, 58)
MUTED = (107, 124, 133)
LIGHT = (242, 245, 246)
LINE = (226, 232, 235)
WHITE = (255, 255, 255)

HERE = Path(__file__).parent
FONT = str(HERE / "assets" / "DejaVuSans.ttf")
FONTB = str(HERE / "assets" / "DejaVuSans-Bold.ttf")

W = 178.0
LEFT = 16.0
COLS = [52, 16, 26, 18, 23, 22, 21]  # Summe 178
FUSS_Y = 286.0


def fmt(n) -> str:
    return f"{round(float(n)):,}".replace(",", ".")


def fmt1(n) -> str:
    r = round(float(n), 1)
    if r == int(r):
        return fmt(r)
    return f"{r:.1f}".replace(".", ",")


def fmt2(n) -> str:
    r = round(float(n), 2)
    return f"{r:.2f}".replace(".", ",")


def _one_line(pdf: FPDF, text: str, breite: float, groesse: float, stil: str = "") -> str:
    """Kürzt Text mit … bis er in die Breite passt."""
    pdf.set_font("D", stil, groesse)
    if pdf.get_string_width(text) <= breite:
        return text
    t = text
    while len(t) > 4 and pdf.get_string_width(t + "…") > breite:
        t = t[:-1]
    return t.rstrip() + "…"


class Seite(FPDF):
    def __init__(self, org: str, traeger: str):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.org = org
        self.traeger = traeger
        self.seiten = 0
        self.add_font("D", "", FONT)
        self.add_font("D", "B", FONTB)
        self.set_auto_page_break(False)
        self.set_margins(LEFT, 12, 16)

    def footer(self):
        self.set_font("D", "", 6.9)
        self.set_text_color(*MUTED)
        self.set_xy(LEFT, -11)
        self.cell(120, 4, f"{self.org} · {self.traeger}", align="L")
        self.set_xy(150, -11)
        self.cell(44, 4, f"Seite {self.page_no()} von {self.seiten}", align="R")


def _kpi(pdf, x, y, w, label, value, sub, accent=PETROL):
    pdf.set_fill_color(*WHITE)
    pdf.set_draw_color(*LINE)
    pdf.set_line_width(0.3)
    pdf.rect(x, y, w, 20, style="DF")
    pdf.set_fill_color(*accent)
    pdf.rect(x, y, 1.6, 20, style="F")
    pdf.set_font("D", "B", 6.2)
    pdf.set_text_color(*MUTED)
    pdf.set_xy(x + 5, y + 3.6)
    pdf.cell(w - 6, 3, _one_line(pdf, label, w - 6, 6.2, "B"))
    pdf.set_font("D", "B", 15)
    pdf.set_text_color(*TEXT)
    pdf.set_xy(x + 5, y + 8.0)
    pdf.cell(w - 6, 7, value)
    pdf.set_font("D", "", 6.2)
    pdf.set_text_color(*MUTED)
    pdf.set_xy(x + 5, y + 15.4)
    pdf.cell(w - 6, 3, _one_line(pdf, sub, w - 6, 6.2))


def _kopfzeile_tabelle(pdf, y):
    titles = [("Mitarbeiter*in", None), ("Wochen-", "stunden"), ("im Projekt", None),
              ("Stunden", "im Jahr"), ("Freie Jugend-", "arbeit (AF 1)"),
              ("Schulwork-", "shops (AF 5)"), ("Leitung, Verw.", "IT (Reserve)")]
    pdf.set_fill_color(*PETROL)
    pdf.set_text_color(*WHITE)
    pdf.rect(LEFT, y, W, 8.4, style="F")
    x = LEFT
    for (a, b), cw in zip(titles, COLS):
        al = "L" if b is None and cw > 20 else "R"
        pdf.set_font("D", "B", 6.9)
        if b:
            pdf.set_xy(x + 1.5, y + 1.0)
            pdf.cell(cw - 3, 3.1, a, align="R")
            pdf.set_xy(x + 1.5, y + 4.5)
            pdf.cell(cw - 3, 3.1, b, align="R")
        else:
            pdf.set_xy(x + 1.5, y + 1.0)
            pdf.cell(cw - 3, 6.4, a, align=al)
        x += cw
    return y + 8.4


def _tabelle(pdf, m, y):
    n = len(m["personen"])
    h = 7.6 if n <= 9 else max(5.4, 7.6 * 9 / n)
    klein = h < 6.6
    y = _kopfzeile_tabelle(pdf, y)
    for i, d in enumerate(m["personen"]):
        if i % 2 == 0:
            pdf.set_fill_color(*LIGHT)
            pdf.rect(LEFT, y, W, h, style="F")
        pdf.set_text_color(*TEXT)
        pdf.set_font("D", "B", 8.2 if not klein else 7.6)
        pdf.set_xy(17.5, y + (0.7 if not klein else 0.4))
        pdf.cell(49, 3.6, _one_line(pdf, d["name"], 49, 8.2 if not klein else 7.6, "B"))
        pdf.set_font("D", "", 6.2)
        pdf.set_text_color(*MUTED)
        pdf.set_xy(17.5, y + (4.4 if not klein else 3.6))
        pdf.cell(49, 3.0, _one_line(pdf, d["rolle"] or "", 49, 6.2))
        pdf.set_text_color(*TEXT)
        cells = [
            (d["wochenstunden_text"], "R"),
            (d["zeitraum_text"], "L"),
            (fmt(d["stunden"]), "R"),
            (fmt(d["af1"]) if abs(d["af1"]) > 0.5 else "–", "R"),
            (fmt(d["af5"]) if d["af5"] > 0.5 else "–", "R"),
            (fmt(d["reserve"]), "R"),
        ]
        x = LEFT + COLS[0]
        for (v, al), cw in zip(cells, COLS[1:]):
            pdf.set_font("D", "", 8.2 if not klein else 7.6)
            pdf.set_xy(x + 1.5, y + (h - 6.0) / 2)
            pdf.cell(cw - 3, 5.4, v, align=al)
            x += cw
        y += h

    pdf.set_draw_color(*LINE)
    pdf.line(LEFT, y, LEFT + W, y)
    pdf.set_text_color(*TEXT)
    pdf.set_font("D", "B", 8.2)
    pdf.set_xy(17.5, y + 2.0)
    pdf.cell(49, 4, "Honorarkräfte (Pool)")
    pdf.set_font("D", "", 6.2)
    pdf.set_text_color(*MUTED)
    pdf.set_xy(17.5, y + 4.8)
    tage = m.get("workshoptage") or 0
    pdf.cell(49, 3, f"externe Workshopleitungen · {fmt(tage)} Workshoptage")
    pdf.set_font("D", "B", 8.2)
    pdf.set_text_color(*TEXT)
    x = LEFT + sum(COLS[:5])
    pdf.set_xy(x + 1.5, y + 2.2)
    pdf.cell(COLS[5] - 3, 4, fmt(m["honorarstunden"]), align="R")
    y += 7.4

    s = m["summen"]
    pdf.set_fill_color(*PETROL)
    pdf.rect(LEFT, y, W, 8.2, style="F")
    pdf.set_text_color(*WHITE)
    pdf.set_font("D", "B", 8.2)
    pdf.set_xy(17.5, y + 2.2)
    pdf.cell(49, 4, "Summe")
    vals = ["", "", fmt(s["stunden"] + m["honorarstunden"]), fmt(s["af1"]),
            fmt(s["af5"]), fmt(s["reserve"])]
    x = LEFT + COLS[0]
    for v, cw in zip(vals, COLS[1:]):
        pdf.set_xy(x + 1.5, y + 2.2)
        pdf.cell(cw - 3, 4, v, align="R")
        x += cw
    return y + 8.2


def _deckung(pdf, m, y):
    pdf.set_font("D", "B", 8.2)
    pdf.set_text_color(*PETROL)
    pdf.set_xy(LEFT, y)
    pdf.cell(66, 4, "DECKUNG DES VERTRAGSSOLLS")
    pdf.set_font("D", "", 7.2)
    pdf.set_text_color(*MUTED)
    pdf.set_xy(LEFT + 68, y)
    pdf.cell(W - 68, 4,
             f"{fmt(m['gedeckt'])} von {fmt(m['soll']['gesamt'])} h gebunden · "
             f"{fmt(m['reserve'])} h Reserve ({m['reserve_pct']:.0f} %) für Leitung, Verwaltung, IT")
    y += 5
    ges = m["kapazitaet"] or 1
    s = m["summen"]
    x = LEFT
    for v, col in ((s["af1"], GREEN), (m["honorarstunden"], AMBER),
                   (s["af5_team"], SAND), (s["reserve"], GREY)):
        w = W * (max(0.0, v) / ges)
        pdf.set_fill_color(*col)
        pdf.rect(x, y, w, 9, style="F")
        x += w
    pdf.set_font("D", "B", 7.2)
    pdf.set_text_color(*WHITE)
    pdf.set_xy(LEFT + 1.5, y + 2)
    pdf.cell(80, 5, f"AF 1  {fmt(s['af1'])} h")
    pdf.set_text_color(*TEXT)
    pdf.set_font("D", "", 7.2)
    pdf.set_xy(LEFT + W + 2.5, y + 2)
    pdf.cell(20, 5, f"{fmt(ges)} h")
    return y + 9


def _balken(pdf, m, y, unten):
    p = m["personen"]
    if not p:
        return y
    label_w, bar_w = 42.0, 104.0
    pitch = 7.0
    while pitch > 3.6 and (y + pitch * len(p) + 12) > unten:
        pitch -= 0.2
    maxv = (max(d["stunden"] for d in p) or 1) * 1.06
    x0 = LEFT + label_w
    hoehe = max(2.0, pitch - 2.6)
    pdf.set_font("D", "", 6.4)
    pdf.set_text_color(*MUTED)
    schritt = 250 if maxv < 2000 else 500
    for tick in range(0, int(maxv) + 1, schritt):
        xx = x0 + bar_w * tick / maxv
        pdf.set_draw_color(232, 237, 239)
        pdf.set_line_width(0.2)
        pdf.line(xx, y, xx, y + pitch * len(p))
        pdf.set_xy(xx - 5, y + pitch * len(p) + 0.6)
        pdf.cell(10, 3, str(tick), align="C")
    pdf.set_xy(x0, y + pitch * len(p) + 3.8)
    pdf.cell(bar_w, 3, "Stunden im Jahr", align="C")
    for i, d in enumerate(p):
        yy = y + i * pitch
        pdf.set_font("D", "", 7.2)
        pdf.set_text_color(*TEXT)
        name = _one_line(pdf, d["name"], label_w - 4, 7.2)
        pdf.set_xy(LEFT, yy + (pitch - 3.6) / 2)
        pdf.cell(label_w - 3, 3.6, name)
        x = x0
        for v, col in ((d["af1"], GREEN), (d["af5"], AMBER), (d["reserve"], GREY)):
            w = bar_w * max(0.0, v) / maxv
            if w > 0:
                pdf.set_fill_color(*col)
                pdf.rect(x, yy + (pitch - hoehe) / 2, w, hoehe, style="F")
            x += w
        pdf.set_font("D", "", 6.8)
        pdf.set_text_color(*MUTED)
        pdf.set_xy(x0 + bar_w * d["stunden"] / maxv + 2, yy + (pitch - 3.4) / 2)
        pdf.cell(16, 3.4, f"{fmt(d['stunden'])} h")
    y += pitch * len(p) + 8
    lx = LEFT
    pdf.set_font("D", "", 6.9)
    for lab, col in (("Freie Jugendarbeit (AF 1)", GREEN),
                     ("Schulworkshops / curricular (AF 5)", AMBER),
                     ("Leitung, Verwaltung, IT", GREY)):
        pdf.set_fill_color(*col)
        pdf.rect(lx, y, 2.6, 2.6, style="F")
        pdf.set_text_color(*MUTED)
        pdf.set_xy(lx + 3.6, y - 0.7)
        pdf.cell(pdf.get_string_width(lab) + 1, 3.6, lab)
        lx += 3.6 + pdf.get_string_width(lab) + 7
    return y + 5


def _fuss(pdf, m, model, p, y_unten):
    s = m["summen"]
    quelle = p.quelle or "Steuerungsdatei in der Nextcloud"
    herkunft = (
        f"{m['jahr']}: {m['land']} {m['arbeitstage']:g} Arbeitstage = {m['werktage']:g} Werktage "
        f"− {m['feiertage_werktags']:g} Feiertage auf Werktagen (Kalender, Blatt Basisdaten); "
        f"Urlaub {p.urlaub:g} / Krankheit {p.krank:g} Tage je Jahr, in Zeiträumen anteilig."
    )
    zeilen = [
        f"Basis: {quelle}" + (f", Stand {p.stand}." if p.stand else "."),
        f"{herkunft} Stunden = Wochenstunden / 5 × Arbeitstage des Zeitraums.",
        f"AF 5 aus dem Team {fmt(s['af5_team'])} h = Soll AF 5 {fmt(p.soll_af5)} h − Honorar "
        f"{fmt(m['honorarstunden'])} h, verteilt nach AF-5-Anteil. Reserve {fmt(s['reserve'])} h "
        f"= Kapazität {fmt(m['kapazitaet'])} h − Vertragssoll {fmt(m['soll']['gesamt'])} h.",
    ]
    if m.get("hinweis"):
        zeilen.append(m["hinweis"])
    hoehe = 3.2
    y = y_unten - hoehe * len(zeilen)
    pdf.set_font("D", "", 6.4)
    pdf.set_text_color(*MUTED)
    pdf.set_xy(LEFT, y)
    for z in zeilen:
        pdf.multi_cell(W, hoehe, z)
        pdf.set_x(LEFT)
    return y


def build(model: dict, p, jahre: list[int] | None = None) -> bytes:
    """PDF mit einer Seite je Jahr. ``jahre`` wählt die Seiten aus."""
    seiten = [j for j in (jahre or model["jahre"].keys()) if j in model["jahre"]]
    seiten.sort()
    pdf = Seite(p.org, p.traeger)
    pdf.seiten = len(seiten)
    for jahr in seiten:
        m = model["jahre"][jahr]
        accent = AMBER if jahr == min(model["jahre"]) else GREEN
        pdf.add_page()
        pdf.set_font("D", "B", 7.2)
        pdf.set_text_color(*PETROL)
        pdf.set_xy(LEFT, 10)
        pdf.cell(100, 3.6, _one_line(pdf, p.org.upper(), 100, 7.2, "B"))
        pdf.set_font("D", "", 7.2)
        pdf.set_text_color(*MUTED)
        pdf.set_xy(110, 10)
        pdf.cell(84, 3.6, _one_line(pdf, p.traeger, 84, 7.2), align="R")
        pdf.set_draw_color(214, 221, 224)
        pdf.set_line_width(0.3)
        pdf.line(LEFT, 15, LEFT + W, 15)

        pdf.set_font("D", "B", 20)
        pdf.set_text_color(*TEXT)
        pdf.set_xy(LEFT, 20.5)
        pdf.cell(120, 9, _one_line(pdf, f"{p.titel} {jahr}", 120, 20, "B"))
        pdf.set_font("D", "B", 11)
        pdf.set_text_color(*accent)
        pdf.set_xy(LEFT + 128, 22.5)
        pdf.cell(50, 5, _one_line(pdf, m["marker"], 50, 11, "B"), align="R")
        pdf.set_font("D", "", 9.4)
        pdf.set_text_color(*MUTED)
        pdf.set_xy(LEFT, 30)
        pdf.cell(W, 4, _one_line(pdf, m["untertitel"], W, 9.4))
        pdf.line(LEFT, 35, LEFT + W, 35)

        s = m["summen"]
        tw = (W - 3 * 2.0) / 4
        _kpi(pdf, LEFT, 38.5, tw, "VERTRAGLICH GESCHULDET", f"{fmt(m['soll']['gesamt'])} h",
             f"{fmt(m['soll']['af1'])} + {fmt(m['soll']['af5'])} h · {fmt2(model['satz'])} €/h")
        _kpi(pdf, LEFT + (tw + 2), 38.5, tw, "FREIE JUGENDARBEIT (AF 1)",
             f"{fmt(m['soll']['af1'])} h", "standortgebundene Angebote", GREEN)
        _kpi(pdf, LEFT + 2 * (tw + 2), 38.5, tw, "CURRICULAR (AF 5)",
             f"{fmt(m['soll']['af5'])} h", f"davon {fmt(m['honorarstunden'])} h Honorarkräfte", AMBER)
        _kpi(pdf, LEFT + 3 * (tw + 2), 38.5, tw, "DECKUNG DES SOLLS",
             f"{m['deckung']:.0f} %", f"{fmt(m['gedeckt'])} von {fmt(m['soll']['gesamt'])} h")

        y = _tabelle(pdf, m, 62)
        y += 4
        y = _deckung(pdf, m, y)
        y += 5
        fuss_start = FUSS_Y - 3.2 * (4 if m.get("hinweis") else 3) - 4
        _balken(pdf, m, y, fuss_start)
        _fuss(pdf, m, model, p, fuss_start)
    return bytes(pdf.output())
