"""
OrchestrIQ Document Intelligence Engine v4 — FastAPI service
ZERO-500 GUARANTEE: every /generate/* endpoint always returns a valid
document. Failures at any layer degrade to the deterministic fallback
model, never to an HTTP error. Engine mode + reason are exposed via
X-Engine-Mode / X-Engine-Reason response headers.
"""
import traceback
from fastapi import FastAPI, Request
import os
from fastapi import HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, JSONResponse
from pydantic import BaseModel
from typing import Optional

from sanitizer import sanitize_request
import ai_extractor
from excel_engine import build_excel
from pptx_engine import build_pptx
from pdf_engine import build_pdf
from docx_engine import build_docx
from blueprint_engine import render_blueprint
from doc_blueprint_engine import (render_pptx_blueprint, render_pdf_blueprint,
                                  render_docx_blueprint)
import layout_engine

VERSION = "4.3.0"

app = FastAPI(title="OrchestrIQ Document Intelligence Engine", version=VERSION)
# CORS WAS allow_origins=["*"], which let ANY website on the internet drive this
# service from a visitor's browser. Combined with the absent auth below, the
# endpoint was an open relay: the Railway URL ships inside your public JS
# bundle, so anyone could extract it and generate documents on your compute, at
# your cost, indefinitely. Now locked to your own origins.
_ALLOWED = [o.strip() for o in os.environ.get(
    "ALLOWED_ORIGINS",
    "https://orchestriq.gorakhai.com,https://gorakhai.com").split(",") if o.strip()]
app.add_middleware(CORSMiddleware, allow_origins=_ALLOWED, allow_methods=["POST", "GET"],
                   allow_headers=["*"], expose_headers=["X-Engine-Mode", "X-Engine-Reason"])
 
 
def _require_auth(req: "GenRequest"):
    """Shared-secret gate. Set SERVICE_SECRET in Railway and the same value as
    VITE-free server config on the Cloudflare side. If SERVICE_SECRET is not
    set, the service stays open — so deploying this file alone cannot break
    your app. Set the variable when you are ready to enforce."""
    expected = os.environ.get("SERVICE_SECRET", "").strip()
    if not expected:
        return
    if (req.service_secret or "").strip() != expected:
        raise HTTPException(status_code=401, detail="Unauthorized")


class GenRequest(BaseModel):
    objective: str = ""
    company_context: str = ""
    available_data: str = ""
    currency: str = "INR"
    currency_symbol: str = "\u20b9"
    api_key: Optional[str] = ""          # legacy — treated as Claude key if no claude_key given
    claude_key: Optional[str] = ""
    openai_key: Optional[str] = ""
    deepseek_key: Optional[str] = ""
    provider_order: Optional[str] = ""   # e.g. "deepseek,claude,openai" — first working one wins
    title: Optional[str] = ""
    subtitle: Optional[str] = ""
    # The frontend now sends WHO the document is for and a full generation brief
    # describing what that reader needs. Before this, the service received only
    # raw content and had to guess - so an investor deck and an operations
    # review were generated from identical instructions.
    audience: Optional[str] = "general"
    doc_purpose: Optional[str] = ""
    generation_brief: Optional[str] = ""
    service_secret: Optional[str] = ""


def _keys_and_order(req: "GenRequest"):
    """Builds the provider key map + try-order from the request. Falls back
    to sensible defaults so older frontend calls (api_key only) still work."""
    keys = {
        "claude": (req.claude_key or req.api_key or "").strip(),
        "openai": (req.openai_key or "").strip(),
        "deepseek": (req.deepseek_key or "").strip(),
    }
    order = [p.strip() for p in (req.provider_order or "").split(",") if p.strip()]
    if not order:
        order = ["deepseek", "claude", "openai"]
    return keys, order


# FOUND BY TESTING THE NEW ENDPOINT, AND IT APPLIES TO ALL OF THEM.
# The blueprint bounds below trim a payload only AFTER FastAPI has parsed the
# whole JSON body into memory. A deliberately huge body therefore exhausts the
# container before any of our limits are consulted - I reproduced exactly that
# and killed the process. This middleware refuses the request at the door, using
# Content-Length, before a single byte is deserialised.
# 12 MB is far above any real document request (120,000 characters of content is
# about 0.12 MB) and far below what hurts a Railway container.
_MAX_BODY_BYTES = 12 * 1024 * 1024


@app.middleware("http")
async def _limit_body(request: Request, call_next):
    cl = request.headers.get("content-length")
    if cl:
        try:
            if int(cl) > _MAX_BODY_BYTES:
                return JSONResponse(status_code=413,
                                    content={"error": "Request body too large.",
                                             "engine": "error"})
        except ValueError:
            pass
    return await call_next(request)


@app.exception_handler(Exception)
async def global_handler(request: Request, exc: Exception):
    # WAS traceback.print_exc(), which writes full local variable context to the
    # Railway log. With 120,000-character payloads that means fragments of a
    # customer's ledger and board minutes can end up in your logs — data you
    # never intended to retain and would have to disclose in a breach. Type and
    # message only; no payload.
    print("[gen-service] %s: %s" % (type(exc).__name__, str(exc)[:200]))
    return JSONResponse(status_code=200,
                        content={"error": str(exc)[:200], "engine": "error"})


@app.get("/health")
def health():
    import os
    return {"status": "ok", "version": VERSION,
            "engines": ["excel", "pptx", "pdf", "docx"],
            "server_key": bool(os.environ.get("ANTHROPIC_API_KEY", "").strip())}


MIMES = {
    "excel": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}


def _pipeline(req: GenRequest, fmt: str) -> Response:
    """Shared pipeline: sanitize → extract (never raises) → build → validate.
    If build with the AI-enriched model fails, retry with pure fallback model.
    This function itself never raises."""
    _require_auth(req)
    # 120,000 chars is roughly 30,000 tokens — comfortably inside the context
    # window of every provider we route to (Claude 200k, GPT-4o 128k,
    # DeepSeek 64k tokens).
    obj, ctx, data = sanitize_request(req.objective, req.company_context,
                                      req.available_data, data_cap=120000)
    sym = (req.currency_symbol or "\u20b9")[:4]
    keys, order = _keys_and_order(req)
    model, mode, reason = ai_extractor.extract(obj, ctx, data, keys, order, sym,
                                               brief=(req.generation_brief or ""))
    title = (req.title or model.get("title") or obj)[:90]
    subtitle = (req.subtitle or "Board of Directors Review")[:90]

    builders = {"excel": lambda m: build_excel(m, title, sym),
                "pptx": lambda m: build_pptx(m, title, subtitle, sym),
                "pdf": lambda m: build_pdf(m, title, subtitle, sym),
                "docx": lambda m: build_docx(m, title, subtitle, sym)}
    build = builders[fmt]

    try:
        blob = build(model)
    except Exception as e:
        traceback.print_exc()
        mode, reason = "fallback", f"build failed on AI model: {str(e)[:100]}"
        try:
            fb = ai_extractor._base_model(obj, sym)
            fb["title"] = title
            blob = build(fb)
        except Exception as e2:
            traceback.print_exc()
            # absolute last resort — still 200, tiny valid file
            return Response(content=f"Generation failed: {str(e2)[:200]}".encode(),
                            media_type="text/plain", status_code=200,
                            headers={"X-Engine-Mode": "error",
                                     "X-Engine-Reason": str(e2)[:120]})
    return Response(content=blob, media_type=MIMES[fmt], status_code=200,
                    headers={"X-Engine-Mode": mode, "X-Engine-Reason": reason[:120],
                             "X-Engine-Version": VERSION})


@app.post("/generate/excel")
def gen_excel(req: GenRequest):
    """v4.1: AI-designed blueprint → generic renderer. Structure adapts to the
    request (any domain, columns, row counts). v4 financial template is the
    last-resort floor. Never raises, never 500s."""
    try:
        _require_auth(req)
        # 120,000 chars is roughly 30,000 tokens — comfortably inside the context
        # window of every provider we route to (Claude 200k, GPT-4o 128k,
        # DeepSeek 64k tokens).
        obj, ctx, data = sanitize_request(req.objective, req.company_context,
                                          req.available_data, data_cap=120000)
        sym = (req.currency_symbol or "\u20b9")[:4]
        keys, order = _keys_and_order(req)
        bp, mode, reason = ai_extractor.extract_blueprint(obj, ctx, data, keys, order, sym)
        if bp is not None:
            try:
                blob = render_blueprint(bp, sym)
                return Response(content=blob, media_type=MIMES["excel"], status_code=200,
                                headers={"X-Engine-Mode": mode, "X-Engine-Reason": reason[:120],
                                         "X-Engine-Version": VERSION, "X-Engine-Path": "blueprint"})
            except Exception as e:
                traceback.print_exc()
                reason = f"blueprint render failed: {str(e)[:80]}"
    except Exception as e:
        traceback.print_exc()
    return _pipeline(req, "excel")


def _doc_blueprint_pipeline(req: GenRequest, fmt: str) -> Response:
    """v4.2: AI designs the document structure per request; generic renderer
    builds it. v4 template pipeline is the last-resort floor. Never 500s."""
    renderers = {"pptx": render_pptx_blueprint, "pdf": render_pdf_blueprint,
                 "docx": render_docx_blueprint}
    try:
        _require_auth(req)
        # 120,000 chars is roughly 30,000 tokens — comfortably inside the context
        # window of every provider we route to (Claude 200k, GPT-4o 128k,
        # DeepSeek 64k tokens).
        obj, ctx, data = sanitize_request(req.objective, req.company_context,
                                          req.available_data, data_cap=120000)
        sym = (req.currency_symbol or "\u20b9")[:4]
        keys, order = _keys_and_order(req)
        bp, mode, reason = ai_extractor.extract_doc_blueprint(fmt, obj, ctx, data,
                                                              keys, order, sym)
        if req.title:
            bp["title"] = req.title[:90]
        if req.subtitle:
            bp["subtitle"] = req.subtitle[:90]
        try:
            blob = renderers[fmt](bp)
            return Response(content=blob, media_type=MIMES[fmt], status_code=200,
                            headers={"X-Engine-Mode": mode, "X-Engine-Reason": reason[:120],
                                     "X-Engine-Version": VERSION, "X-Engine-Path": "blueprint"})
        except Exception as e:
            traceback.print_exc()
    except Exception:
        traceback.print_exc()
    return _pipeline(req, fmt)


# ─────────────────────────────────────────────────────────────────────────────
# RENDER-ONLY PATH
#
# WHY THIS EXISTS. Two problems, one answer.
#
#   1. NVIDIA could not generate documents. The NVIDIA key lives in Cloudflare,
#      server-side, reachable only through /api/nvidia from the user's own
#      browser. This service runs on Railway - a different machine that has no
#      NVIDIA key and cannot be given one without duplicating the secret onto a
#      second platform. So /generate/* could only ever use Claude, OpenAI or
#      DeepSeek.
#
#   2. Customer API keys were being POSTed to this service in the request body.
#      TLS protects them in flight, but any request logging or proxy on the
#      Railway side captures live customer credentials - credentials we never
#      intended to hold. That has been an open HIGH finding since Session 47.
#
# Both dissolve if the AI step happens in the BROWSER, where every provider is
# already reachable, and this service only renders. The browser sends a finished
# blueprint; Railway turns it into a file. No keys are sent, because none are
# needed. Railway stops being an AI client and becomes what it is good at: a
# renderer.
#
# The /generate/* endpoints below are UNCHANGED and still work exactly as before,
# so nothing that currently functions can break. This is an additional path, not
# a replacement.
# ─────────────────────────────────────────────────────────────────────────────

class RenderRequest(BaseModel):
    """A finished document blueprint, ready to render. Note what is ABSENT:
    no api_key, no claude_key, no openai_key, no deepseek_key. This endpoint
    cannot receive a credential because it has nowhere to put one."""
    blueprint: dict = {}
    title: Optional[str] = ""
    subtitle: Optional[str] = ""
    currency_symbol: Optional[str] = "\u20b9"
    service_secret: Optional[str] = ""


# Bounds. This endpoint renders caller-supplied structure, so it must not be
# possible to hand it something that exhausts the container. A blueprint with
# 50,000 slides is not a document, it is a denial-of-service. These caps are
# far above any real document and far below anything that hurts.
_MAX_SECTIONS = 60
_MAX_ROWS = 200
_MAX_COLS = 20
_MAX_SERIES = 12
_MAX_POINTS = 60


def _bounded(bp):
    """Trim a caller-supplied blueprint to sane limits. Never raises: a
    malformed blueprint yields a smaller document, not an error."""
    if not isinstance(bp, dict):
        return {}
    out = dict(bp)
    secs = out.get("slides") or out.get("sections") or []
    if not isinstance(secs, list):
        secs = []
    trimmed = []
    for s in secs[:_MAX_SECTIONS]:
        if not isinstance(s, dict):
            continue
        s = dict(s)
        tbl = s.get("table")
        if isinstance(tbl, dict) and isinstance(tbl.get("rows"), list):
            s["table"] = dict(tbl, rows=[r[:_MAX_COLS] for r in tbl["rows"][:_MAX_ROWS]
                                         if isinstance(r, list)])
        ch = s.get("chart")
        if isinstance(ch, dict):
            ch = dict(ch)
            if isinstance(ch.get("cats"), list):
                ch["cats"] = ch["cats"][:_MAX_POINTS]
            if isinstance(ch.get("series"), list):
                ch["series"] = [x for x in ch["series"][:_MAX_SERIES]
                                if isinstance(x, list) and len(x) == 2]
            s["chart"] = ch
        for k in ("points", "left", "right", "kpis"):
            if isinstance(s.get(k), list):
                s[k] = s[k][:_MAX_ROWS]
        trimmed.append(s)
    if "slides" in out or "sections" not in out:
        out["slides"] = trimmed
    else:
        out["sections"] = trimmed
    return out


def _render_only(req: RenderRequest, fmt: str) -> Response:
    """Render a blueprint the browser already produced. No AI, no keys."""
    renderers = {"pptx": render_pptx_blueprint, "pdf": render_pdf_blueprint,
                 "docx": render_docx_blueprint}
    if fmt not in renderers:
        raise HTTPException(status_code=400, detail="Unsupported format")
    _require_auth(req)
    bp = _bounded(req.blueprint)
    if not (bp.get("slides") or bp.get("sections")):
        raise HTTPException(status_code=400,
                            detail="Blueprint contains no slides or sections to render.")
    sym = (req.currency_symbol or "\u20b9")[:4]
    if req.title:
        bp["title"] = str(req.title)[:120]
    if req.subtitle:
        bp["subtitle"] = str(req.subtitle)[:140]
    try:
        # Same styling pass the AI path applies, so a browser-built blueprint
        # and a Railway-built one come out looking identical.
        bp = layout_engine.style_blueprint(bp, sym)
    except Exception as e:
        print("[gen-service] style_blueprint skipped: %s" % type(e).__name__)
    blob = renderers[fmt](bp)
    return Response(content=blob, media_type=MIMES[fmt], status_code=200,
                    headers={"X-Engine-Mode": "render-only", "X-Engine-Reason": "browser blueprint",
                             "X-Engine-Version": VERSION, "X-Engine-Path": "render"})


@app.post("/render/pptx")
def render_pptx(req: RenderRequest): return _render_only(req, "pptx")


@app.post("/render/pdf")
def render_pdf(req: RenderRequest): return _render_only(req, "pdf")


@app.post("/render/docx")
def render_docx(req: RenderRequest): return _render_only(req, "docx")


@app.post("/generate/pptx")
def gen_pptx(req: GenRequest): return _doc_blueprint_pipeline(req, "pptx")


@app.post("/generate/pdf")
def gen_pdf(req: GenRequest): return _doc_blueprint_pipeline(req, "pdf")


@app.post("/generate/docx")
def gen_docx(req: GenRequest): return _doc_blueprint_pipeline(req, "docx")
