"""FastAPI application for accounting.90ten.life.

    uvicorn accounting.main:app --host 0.0.0.0 --port 8000

Routers are registered here; each package owns its own `router.py`. The web process does not poll
IMAP or run reminders — `accounting.worker` does, as a separate container.
"""
from __future__ import annotations

import logging
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.exceptions import HTTPException as StarletteHTTPException

from accounting import APP_NAME, APP_VERSION
from accounting.config import settings
from accounting.db.pool import close_pool, get_pool
from accounting.identity import crypto
from accounting.identity.sessions import resolve
from accounting.web.templating import render

log = logging.getLogger("accounting")


@asynccontextmanager
async def lifespan(app: FastAPI):
    crypto.field_key()  # fail fast outside test mode without a key
    get_pool()
    yield
    close_pool()


app = FastAPI(title=APP_NAME, version=APP_VERSION, docs_url=None, redoc_url=None, openapi_url=None, lifespan=lifespan)
app.mount("/static", StaticFiles(directory=str(Path(__file__).parent / "web" / "static")), name="static")


@app.middleware("http")
async def security_headers(request: Request, call_next):
    response: Response = await call_next(request)
    # A signed-in page is never cached anywhere (the MaxxDash lesson: `private` is not `per user`).
    response.headers["Cache-Control"] = "no-store"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self'; "
        "frame-ancestors 'none'; form-action 'self'; base-uri 'none'"
    )
    if settings.secure_cookies:
        response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    return response


@app.exception_handler(StarletteHTTPException)
async def http_error(request: Request, exc: StarletteHTTPException):
    if exc.status_code == 404:
        return render(request, "404.html", {}, status_code=404)
    if exc.status_code == 403:
        return render(request, "error.html", {"message": exc.detail or "Not allowed."}, status_code=403)
    return render(request, "error.html", {"message": "Something went wrong. Nothing was lost — try again."},
                  status_code=exc.status_code)


@app.get("/healthz")
def healthz():
    return {"ok": True, "version": APP_VERSION}


@app.get("/")
def home(request: Request):
    user = resolve(request)
    if user is None:
        return RedirectResponse("/login", status_code=303)
    request.state.user = user
    from accounting.reports.dashboard import home_tiles
    return render(request, "home.html", {"tiles": home_tiles(user)})


from accounting.identity.router import router as identity_router  # noqa: E402
from accounting.web.routers import all_routers  # noqa: E402

app.include_router(identity_router)
for r in all_routers():
    app.include_router(r)
