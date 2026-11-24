"""Web-Auswertung Angebotsstunden CIA Spandau.

Hauptausgabe ist die druckbare A4-Seite (PDF). Die Seite davor ist ein
Justierformular: jeder Faktor und jede Zahl lässt sich ändern, die Vorschau ist
das echte PDF. Datenquelle und Speicher ist die Steuerungsdatei (XLSX) in der
Nextcloud — die App hält keine eigenen Daten.
"""
from __future__ import annotations

import base64
import os
import threading
import time
import urllib.parse
import urllib.request
from io import BytesIO

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.templating import Jinja2Templates

import basisdaten
import core
import params as P
import report
import xlsx_write

app = FastAPI(title="Angebotsstunden CIA Spandau")
templates = Jinja2Templates(directory=os.path.join(os.path.dirname(__file__), "templates"))
templates.env.filters["de"] = lambda n: f"{round(float(n)):,}".replace(",", ".")
templates.env.filters["g"] = lambda n: P.fmt(n)
templates.env.filters["g1"] = lambda n: P.fmt(n, 1)
templates.env.filters["g3"] = lambda n: P.fmt(n, 3)


def _norm_url(u: str) -> str:
    u = (u or "").strip().rstrip("/")
    if u and not u.startswith(("http://", "https://")):
        u = "https://" + u
    return u or "http://localhost"


NC_URL = _norm_url(os.environ.get("NC_URL", ""))
NC_USER = os.environ.get("NC_USER", "")
NC_PASS = os.environ.get("NC_PASS", "")
NC_PATH = os.environ.get(
    "NC_PATH", "1 Team CIA/Finanzierung CIA/Angebotsstunden-CIA-Modell.xlsx")
LOCAL_FILE = os.environ.get("LOCAL_FILE", "")
CACHE_SECONDS = int(os.environ.get("CACHE_SECONDS", "300"))

_lock = threading.Lock()
_cache: dict = {"bytes": None, "ts": 0.0, "stand": "", "quelle": "", "backup": ""}


def _dav_url(pfad: str) -> str:
    return (f"{NC_URL}/remote.php/dav/files/{urllib.parse.quote(NC_USER)}/"
            + urllib.parse.quote(pfad))


def _dav(methode: str, pfad: str, daten: bytes | None = None) -> bytes:
    req = urllib.request.Request(_dav_url(pfad), data=daten, method=methode)
    token = base64.b64encode(f"{NC_USER}:{NC_PASS}".encode()).decode()
    req.add_header("Authorization", f"Basic {token}")
    if daten is not None:
        req.add_header("Content-Type", "application/vnd.openxmlformats-officedocument."
                                        "spreadsheetml.sheet")
    with urllib.request.urlopen(req, timeout=45) as r:
        return r.read()


def _load(force: bool = False) -> dict:
    with _lock:
        frisch = _cache["bytes"] is not None and (time.time() - _cache["ts"] < CACHE_SECONDS)
        if frisch and not force:
            return _cache
        if LOCAL_FILE:
            data = open(LOCAL_FILE, "rb").read()
            quelle = f"Datei {LOCAL_FILE}"
        else:
            if not (NC_USER and NC_PASS):
                raise HTTPException(500, "NC_USER/NC_PASS sind nicht gesetzt")
            data = _dav("GET", NC_PATH)
            quelle = f"Nextcloud-Datei „{NC_PATH}“"
        _cache.update(bytes=data, ts=time.time(), quelle=quelle,
                      stand=time.strftime("%d.%m.%Y %H:%M", time.localtime()))
        return _cache


def _basis_params() -> P.Params:
    c = _load()
    try:
        p = P.from_xlsx(c["bytes"])
    except Exception as e:
        raise HTTPException(500, f"Steuerungsdatei nicht lesbar: {e}")
    p.quelle = c["quelle"]
    p.stand = c["stand"]
    return p


def _aus_form(form) -> tuple[P.Params, dict]:
    """Parameter aus der Steuerungsdatei, überschrieben von dem, was im Formular steht."""
    p = _basis_params()
    if P.betroffen(form):
        p = P.overlay(p, form)
    return p, core.compute(p)


def _seiten(jahre_param: str, p: P.Params) -> list[int]:
    gewaehlt = [j for j in p.jahr_liste() if str(j) in (jahre_param or "")]
    return gewaehlt or p.jahr_liste()


@app.get("/healthz")
def healthz():
    """Liveness für den Container-Healthcheck: immer 200, Datenstand als Feld."""
    try:
        c = _load()
        p = P.from_xlsx(c["bytes"])
        return {"ok": True, "data_ok": True, "bytes": len(c["bytes"]), "stand": c["stand"],
                "quelle": c["quelle"], "jahre": p.jahr_liste(),
                "personen": {j: len(x.personen) for j, x in p.jahre.items()}}
    except Exception as e:
        return {"ok": True, "data_ok": False, "error": str(e)[:200],
                "quelle": f"{NC_URL} · {NC_PATH}"}


@app.get("/", response_class=HTMLResponse)
def index(request: Request, refresh: int = 0):
    if refresh:
        _load(force=True)
    try:
        p, model = _aus_form(request.query_params)
    except HTTPException as e:
        return HTMLResponse(
            f"<h1>Steuerungsdatei nicht verfügbar</h1><p>{e.detail}</p>"
            f"<p>Quelle: {NC_URL} · {NC_PATH}</p>", status_code=503)
    form = p.to_form()
    alle = ",".join(str(j) for j in p.jahr_liste())
    seiten = request.query_params.get("y") or alle
    if not set(seiten.split(",")) & {str(j) for j in p.jahr_liste()}:
        seiten = alle
    return templates.TemplateResponse(request, "index.html", {
        "p": p, "m": model, "f": form, "jahre": p.jahr_liste(),
        "warnungen": core.warnungen(p, model),
        "query": p.to_query(), "stand": p.stand, "quelle": p.quelle,
        "seiten": seiten,
        "max_zeilen": P.MAX_ZEILEN,
        "laender": basisdaten.BUNDESLAENDER,
    })


@app.get("/kennzahlen", response_class=HTMLResponse)
def kennzahlen(request: Request):
    """Kontrollzahlen und Hinweise als Fragment — nach jeder Änderung neu geholt."""
    try:
        p, model = _aus_form(request.query_params)
    except HTTPException as e:
        return HTMLResponse(f"<p class='warn'>Steuerungsdatei nicht lesbar: {e.detail}</p>")
    return templates.TemplateResponse(request, "_fragment.html", {
        "p": p, "m": model, "warnungen": core.warnungen(p, model),
    })


@app.get("/export.pdf")
def export_pdf(request: Request, dl: int = 0):
    p, model = _aus_form(request.query_params)
    seiten = _seiten(request.query_params.get("y", ""), p)
    daten = report.build(model, p, seiten)
    name = f"{p.titel.strip().replace(' ', '-')}-{'-'.join(str(j) for j in seiten)}.pdf"
    return Response(daten, media_type="application/pdf", headers={
        "Content-Disposition": f'{"attachment" if dl else "inline"}; filename="{name}"'})


@app.get("/export.xlsx")
def export_xlsx():
    c = _load()
    return Response(c["bytes"],
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": 'attachment; filename="Angebotsstunden-CIA.xlsx"'})


@app.post("/sichern")
async def sichern(request: Request):
    """Schreibt die justierten Werte in die Steuerungsdatei zurück (mit Sicherung)."""
    # Bewusst ohne python-multipart: der Posten kommt als urlencoded, das parst die
    # Standardbibliothek. (Starlette verlangt für request.form() ein Zusatzpaket.)
    roh = await request.body()
    felder = {k: v[0] for k, v in urllib.parse.parse_qs(
        roh.decode("utf-8", "replace"), keep_blank_values=True).items()}
    p, _ = _aus_form(felder)
    if not (NC_USER and NC_PASS) or LOCAL_FILE:
        return JSONResponse({"ok": False,
                             "fehler": "Ohne Nextcloud-Zugang wird nicht gesichert."}, status_code=400)
    try:
        alt = _load(force=True)["bytes"]
        neu = xlsx_write.schreiben(alt, p)
        stempel = time.strftime("%Y%m%d-%H%M%S")
        basis = NC_PATH
        sicherung = basis.rsplit(".xlsx", 1)[0] + f".sicherung-{stempel}.xlsx"
        _dav("PUT", sicherung, alt)
        _dav("PUT", basis, neu)
        with _lock:
            _cache.update(bytes=neu, ts=time.time(), backup=sicherung,
                          stand=time.strftime("%d.%m.%Y %H:%M", time.localtime()))
        return {"ok": True, "bytes": len(neu), "sicherung": sicherung,
                "stand": _cache["stand"], "personen": {j: len(x.personen) for j, x in p.jahre.items()}}
    except Exception as e:
        return JSONResponse({"ok": False, "fehler": str(e)[:300]}, status_code=502)


@app.get("/neu")
def neu():
    return Response(status_code=303, headers={"Location": "/?refresh=1"})
