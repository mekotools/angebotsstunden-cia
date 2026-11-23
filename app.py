"""Minimale Web-App: liest die Steuerungs-XLSX aus Nextcloud, zeigt die Auswertung, exportiert PDF/Excel."""
from __future__ import annotations

import base64
import os
import threading
import time
import urllib.parse
import urllib.request
from io import BytesIO

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates

import core
import report

app = FastAPI(title="Angebotsstunden CIA Spandau")
templates = Jinja2Templates(directory=os.path.join(os.path.dirname(__file__), "templates"))
templates.env.filters["de"] = lambda n: f"{round(float(n)):,}".replace(",", ".")
templates.env.filters["g"] = lambda n: f"{float(n):g}".replace(".", ",")

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
APP_TITLE = os.environ.get("APP_TITLE", "Angebotsstunden")
ORG = os.environ.get("ORG", "Medienkompetenzzentrum CIA Spandau")

_lock = threading.Lock()
_cache: dict = {"bytes": None, "ts": 0.0, "stand": "", "quelle": ""}


def _dav_get() -> bytes:
    url = f"{NC_URL}/remote.php/dav/files/{urllib.parse.quote(NC_USER)}/" + \
          urllib.parse.quote(NC_PATH)
    req = urllib.request.Request(url)
    token = base64.b64encode(f"{NC_USER}:{NC_PASS}".encode()).decode()
    req.add_header("Authorization", f"Basic {token}")
    with urllib.request.urlopen(req, timeout=30) as r:
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
            data = _dav_get()
            quelle = f"Nextcloud: {NC_PATH}"
        _cache.update(bytes=data, ts=time.time(), quelle=quelle,
                      stand=time.strftime("%d.%m.%Y %H:%M", time.localtime()))
        return _cache


def _model(force: bool = False) -> tuple[dict, dict]:
    c = _load(force)
    try:
        return core.load(c["bytes"]), c
    except Exception as e:  # pragma: no cover
        raise HTTPException(500, f"Steuerungsdatei nicht lesbar: {e}")


@app.get("/healthz")
def healthz():
    """Liveness für den Container-Healthcheck: immer 200, Datenstand als Feld."""
    try:
        c = _load()
        return {"ok": True, "data_ok": True, "bytes": len(c["bytes"]),
                "stand": c["stand"], "quelle": c["quelle"]}
    except Exception as e:
        return {"ok": True, "data_ok": False, "error": str(e)[:200],
                "quelle": f"{NC_URL} · {NC_PATH}"}


@app.get("/", response_class=HTMLResponse)
def index(request: Request, jahr: int = 2026, refresh: int = 0):
    try:
        model, c = _model(bool(refresh))
    except HTTPException as e:
        return HTMLResponse(
            f"<h1>Daten nicht verfügbar</h1><p>{e.detail}</p>"
            f"<p>Quelle: {NC_URL} · {NC_PATH}</p>", status_code=503)
    if jahr not in model["jahre"]:
        jahr = 2026
    return templates.TemplateResponse(request, "index.html", {
        "m": model["jahre"][jahr], "jahr": jahr,
        "jahre": sorted(model["jahre"]), "arbeitstage": model["arbeitstage"],
        "wochen": model["wochen"], "stand": c["stand"], "quelle": c["quelle"],
        "app_title": APP_TITLE, "org": ORG,
    })


@app.get("/export.pdf")
def export_pdf():
    model, c = _model()
    data = report.build(model, c["stand"], ORG)
    return Response(data, media_type="application/pdf", headers={
        "Content-Disposition": 'inline; filename="Angebotsstunden-CIA.pdf"'})


@app.get("/export.xlsx")
def export_xlsx():
    c = _load()
    return Response(c["bytes"],
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": 'attachment; filename="Angebotsstunden-CIA.xlsx"'})


@app.get("/refresh")
def refresh():
    _load(force=True)
    return Response(status_code=303, headers={"Location": "/"})
