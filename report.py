"""PDF-Export des Stundengerüsts (2 Seiten, ohne matplotlib)."""
from __future__ import annotations
from pathlib import Path
from fpdf import FPDF

PETROL = (30, 75, 90)
GREEN = (46, 125, 107)
AMBER = (201, 127, 46)
GREY = (198, 205, 209)
TEXT = (34, 49, 58)
MUTED = (107, 124, 133)
LIGHT = (242, 245, 246)
WHITE = (255, 255, 255)

HERE = Path(__file__).parent
FONT = str(HERE / "assets" / "DejaVuSans.ttf")
FONTB = str(HERE / "assets" / "DejaVuSans-Bold.ttf")

W = 178
COLS = [52, 16, 26, 18, 23, 22, 21]  # Summe 178


def fmt(n) -> str:
    return f"{round(n):,}".replace(",", ".")


class PDF(FPDF):
    def footer(self):
        self.set_font("D", "", 6.9)
        self.set_text_color(*MUTED)
        self.set_xy(16, -11)
        self.cell(120, 4, "Medienkompetenzzentrum CIA Spandau · Staakkato Kinder und Jugend e.V.", align="L")
        self.set_xy(150, -11)
        self.cell(44, 4, f"Seite {self.page_no()} von 2", align="R")


def _kpi(pdf, x, y, w, label, value, sub, accent=PETROL):
    pdf.set_fill_color(*WHITE); pdf.set_draw_color(226, 232, 235); pdf.set_line_width(0.3)
    pdf.rect(x, y, w, 20, style="DF")
    pdf.set_fill_color(*accent); pdf.rect(x, y, 1.6, 20, style="F")
    pdf.set_font("D", "B", 6.8); pdf.set_text_color(*MUTED); pdf.set_xy(x + 5, y + 3.8)
    pdf.cell(w - 6, 3, label)
    pdf.set_font("D", "B", 16); pdf.set_text_color(*TEXT); pdf.set_xy(x + 5, y + 8.2)
    pdf.cell(w - 6, 7, value)
    pdf.set_font("D", "", 6.9); pdf.set_text_color(*MUTED); pdf.set_xy(x + 5, y + 15.4)
    pdf.cell(w - 6, 3, sub)


def _kopfzeile(pdf, cols, y):
    titles = [("Mitarbeiter*in", None), ("Wochen-", "stunden"), ("im Projekt", None),
              ("Stunden", "im Jahr"), ("Freie Jugend-", "arbeit (AF 1)"),
              ("Schulwork-", "shops (AF 5)"), ("Leitung, Verw.", "IT (Reserve)")]
    pdf.set_fill_color(*PETROL); pdf.set_text_color(*WHITE); pdf.rect(16, y, W, 8.4, style="F")
    x = 16
    for (a, b), cw in zip(titles, cols):
        al = "L" if b is None and cw > 20 else "R"
        pdf.set_font("D", "B", 6.9)
        if b:
            pdf.set_xy(x + 1.5, y + 1.0); pdf.cell(cw - 3, 3.1, a, align="R")
            pdf.set_xy(x + 1.5, y + 4.5); pdf.cell(cw - 3, 3.1, b, align="R")
        else:
            pdf.set_xy(x + 1.5, y + 1.0); pdf.cell(cw - 3, 6.4, a, align=al)
        x += cw
    return y + 8.4


def _tabelle(pdf, m, y):
    y = _kopfzeile(pdf, COLS, y)
    for i, d in enumerate(m["personen"]):
        h = 7.6
        if i % 2 == 0:
            pdf.set_fill_color(*LIGHT); pdf.rect(16, y, W, h, style="F")
        pdf.set_text_color(*TEXT)
        pdf.set_font("D", "B", 8.2); pdf.set_xy(17.5, y + 0.7); pdf.cell(49, 3.8, d["name"])
        pdf.set_font("D", "", 6.2); pdf.set_text_color(*MUTED)
        pdf.set_xy(17.5, y + 4.4); pdf.cell(49, 3.0, d["rolle"] or "")
        pdf.set_text_color(*TEXT)
        cells = [
            (f"{d['wochenstunden']:g}".replace(".", ","), "R"),
            (d["zeitraum"], "L"),
            (fmt(d["netto"]), "R"),
            (fmt(d["af1"]) if d["af1"] > 0.5 else "–", "R"),
            (fmt(d["af5"]) if d["af5"] > 0.5 else "–", "R"),
            (fmt(d["reserve"]), "R"),
        ]
        x = 16 + COLS[0]
        for (v, al), cw in zip(cells, COLS[1:]):
            pdf.set_font("D", "", 8.2)
            pdf.set_xy(x + 1.5, y + 0.6); pdf.cell(cw - 3, 6.0, v, align=al)
            x += cw
        y += h
    pdf.set_draw_color(226, 232, 235); pdf.line(16, y, 194, y)
    pdf.set_text_color(*TEXT)
    pdf.set_font("D", "B", 8.2); pdf.set_xy(17.5, y + 2.2); pdf.cell(49, 4, "Honorarkräfte (Pool)")
    pdf.set_font("D", "", 6.2); pdf.set_text_color(*MUTED)
    pdf.set_xy(17.5, y + 5.0); pdf.cell(49, 3, "externe Workshopleitungen")
    pdf.set_font("D", "B", 8.2); pdf.set_text_color(*TEXT)
    x = 16 + sum(COLS[:5])
    pdf.set_xy(x + 1.5, y + 2.2); pdf.cell(COLS[5] - 3, 4, fmt(m["honorarstunden"]), align="R")
    y += 7.4

    s = m["summen"]
    pdf.set_fill_color(*PETROL); pdf.rect(16, y, W, 8.2, style="F")
    pdf.set_text_color(*WHITE); pdf.set_font("D", "B", 8.2)
    pdf.set_xy(17.5, y + 2.2); pdf.cell(49, 4, "Summe")
    vals = ["", "", fmt(s["netto"] + m["honorarstunden"]), fmt(s["af1"]),
            fmt(s["af5_team"] + s["af5_honorar"]), fmt(s["reserve"])]
    x = 16 + COLS[0]
    for v, cw in zip(vals, COLS[1:]):
        pdf.set_xy(x + 1.5, y + 2.2); pdf.cell(cw - 3, 4, v, align="R")
        x += cw
    return y + 8.2


def _balken(pdf, m, y):
    p = m["personen"]
    label_w, bar_w, row = 42.0, 104.0, 7.0
    maxv = max(d["netto"] for d in p) * 1.06
    x0 = 16 + label_w
    pdf.set_font("D", "", 6.4); pdf.set_text_color(*MUTED)
    for tick in range(0, int(maxv) + 1, 250):
        xx = x0 + bar_w * tick / maxv
        pdf.set_draw_color(232, 237, 239); pdf.set_line_width(0.2)
        pdf.line(xx, y, xx, y + row * len(p))
        pdf.set_xy(xx - 5, y + row * len(p) + 0.8); pdf.cell(10, 3, str(tick), align="C")
    pdf.set_xy(x0, y + row * len(p) + 4.0); pdf.cell(bar_w, 3, "Stunden im Jahr", align="C")
    for i, d in enumerate(p):
        yy = y + i * row
        pdf.set_font("D", "", 7.4); pdf.set_text_color(*TEXT)
        name = d["name"]
        while pdf.get_string_width(name) > label_w - 4 and len(name) > 4:
            name = name[:-2] + "…"
        pdf.set_xy(16, yy + 1.8); pdf.cell(label_w - 3, 3.6, name)
        x = x0
        for v, col in ((d["af1"], GREEN), (d["af5"], AMBER), (d["reserve"], GREY)):
            w = bar_w * v / maxv
            if w > 0:
                pdf.set_fill_color(*col); pdf.rect(x, yy, w, row - 2.4, style="F")
            x += w
        pdf.set_font("D", "", 6.8); pdf.set_text_color(*MUTED)
        pdf.set_xy(x0 + bar_w * d["netto"] / maxv + 2, yy + 1.6)
        pdf.cell(16, 3.4, f"{fmt(d['netto'])} h")
    y += row * len(p) + 9
    lx = 16.0
    pdf.set_font("D", "", 6.9)
    for lab, col in (("Freie Jugendarbeit (AF 1)", GREEN),
                     ("Schulworkshops / curricular (AF 5)", AMBER),
                     ("Leitung, Verwaltung, IT", GREY)):
        pdf.set_fill_color(*col); pdf.rect(lx, y, 2.6, 2.6, style="F")
        pdf.set_text_color(*MUTED); pdf.set_xy(lx + 3.6, y - 0.7)
        pdf.cell(pdf.get_string_width(lab) + 1, 3.6, lab)
        lx += 3.6 + pdf.get_string_width(lab) + 7
    return y + 5


def build(model: dict, stand: str) -> bytes:
    pdf = PDF(orientation="P", unit="mm", format="A4")
    pdf.add_font("D", "", FONT)
    pdf.add_font("D", "B", FONTB)
    pdf.set_auto_page_break(False)
    pdf.set_margins(16, 12, 16)

    for jahr, marker, accent in ((2026, "IST", AMBER), (2027, "PLANUNG", GREEN)):
        m = model["jahre"][jahr]
        pdf.add_page()
        pdf.set_font("D", "B", 7.2); pdf.set_text_color(*PETROL)
        pdf.set_xy(16, 10); pdf.cell(100, 3.6, "MEDIENKOMPETENZZENTRUM CIA SPANDAU")
        pdf.set_font("D", "", 7.2); pdf.set_text_color(*MUTED)
        pdf.set_xy(120, 10); pdf.cell(74, 3.6, "Staakkato Kinder und Jugend e.V.", align="R")
        pdf.set_draw_color(214, 221, 224); pdf.set_line_width(0.3); pdf.line(16, 15, 194, 15)
        pdf.set_font("D", "B", 20); pdf.set_text_color(*TEXT)
        pdf.set_xy(16, 20.5); pdf.cell(120, 9, f"Angebotsstunden {jahr}")
        pdf.set_font("D", "B", 11); pdf.set_text_color(*accent)
        pdf.set_xy(148, 22.5); pdf.cell(30, 5, marker, align="R")
        pdf.set_font("D", "", 9.4); pdf.set_text_color(*MUTED)
        pdf.set_xy(16, 30); pdf.cell(178, 4, "Wer arbeitet wie viel woran — Aufstellung für das Jugendamt Spandau")
        pdf.line(16, 35, 194, 35)

        soll = m["soll"]
        _kpi(pdf, 16, 38.5, 57, "VERTRAGLICH GESCHULDET", f"{fmt(soll['ges'])} h",
             "§ 5 Nr. 1 · 5.000 + 1.828 · 38,08 €/h")
        _kpi(pdf, 75.5, 38.5, 57, "FREIE JUGENDARBEIT (AF 1)", f"{fmt(soll['af1'])} h",
             "standortgebundene Angebote")
        _kpi(pdf, 135, 38.5, 57, "CURRICULAR (AF 5)", f"{fmt(soll['af5'])} h",
             f"davon {fmt(m['honorarstunden'])} h Honorarkräfte")

        y = _tabelle(pdf, m, 62)
        y += 4
        pdf.set_font("D", "B", 8.2); pdf.set_text_color(*PETROL)
        pdf.set_xy(16, y); pdf.cell(66, 4, "DECKUNG DES VERTRAGSSOLLS")
        pdf.set_font("D", "", 7.2); pdf.set_text_color(*MUTED)
        pdf.set_xy(84, y)
        pdf.cell(110, 4, f"{fmt(soll['ges'])} h gebunden · {fmt(m['reserve'])} h Reserve "
                         f"({m['reserve_pct']:.0f} %) für Leitung, Verwaltung, IT")
        y += 5
        ges = m["kapazitaet"]
        x = 16
        for v, col in ((m["summen"]["af1"], GREEN), (m["honorarstunden"], AMBER),
                       (m["summen"]["af5_team"], (226, 168, 100)), (m["reserve"], GREY)):
            w = 148 * v / ges
            pdf.set_fill_color(*col); pdf.rect(x, y, w, 9, style="F")
            x += w
        pdf.set_font("D", "B", 7.2); pdf.set_text_color(*WHITE)
        pdf.set_xy(17.5, y + 2); pdf.cell(80, 5, f"AF 1  {fmt(m['summen']['af1'])} h")
        pdf.set_text_color(*TEXT); pdf.set_font("D", "", 7.2)
        pdf.set_xy(16 + 148 + 2.5, y + 2); pdf.cell(24, 5, f"{fmt(ges)} h")
        y += 14
        y = _balken(pdf, m, y)

        pdf.set_font("D", "", 6.6); pdf.set_text_color(*MUTED)
        pdf.set_xy(16, max(y + 2, 250))
        pdf.multi_cell(W, 3.0,
                       f"Basis: Steuerungsdatei in Nextcloud (Blätter Eingabe/Annahmen), Leistungsvertrag "
                       f"§ 5 Nr. 1, Planung 2026.ods. Stunden im Jahr = Wochenstunden / 5 × "
                       f"{model['arbeitstage']:.0f} Arbeitstage. Stand der Daten: {stand}.")

    return bytes(pdf.output())
