"""
OrchestrIQ Document Intelligence Engine v4 — FastAPI service
ZERO-500 GUARANTEE: every /generate/* endpoint always returns a valid
document. Failures at any layer degrade to the deterministic fallback
model, never to an HTTP error. Engine mode + reason are exposed via
X-Engine-Mode / X-Engine-Reason response headers.
"""
import traceback
from fastapi import FastAPI, Request
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

VERSION = "4.3.0"

app = FastAPI(title="OrchestrIQ Document Intelligence Engine", version=VERSION)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"],
                   allow_headers=["*"], expose_headers=["X-Engine-Mode", "X-Engine-Reason"])


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


@app.exception_handler(Exception)
async def global_handler(request: Request, exc: Exception):
    traceback.print_exc()
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
    obj, ctx, data = sanitize_request(req.objective, req.company_context, req.available_data)
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
        obj, ctx, data = sanitize_request(req.objective, req.company_context, req.available_data)
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
        obj, ctx, data = sanitize_request(req.objective, req.company_context, req.available_data)
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


@app.post("/generate/pptx")
def gen_pptx(req: GenRequest): return _doc_blueprint_pipeline(req, "pptx")


@app.post("/generate/pdf")
def gen_pdf(req: GenRequest): return _doc_blueprint_pipeline(req, "pdf")


@app.post("/generate/docx")
def gen_docx(req: GenRequest): return _doc_blueprint_pipeline(req, "docx")
