"""
OrchestrIQ Generation Service v3 — FastAPI
Railway deployment: https://orchestriq-gen-service-production.up.railway.app
"""
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pydantic import BaseModel
from typing import Optional
import traceback

from excel_engine  import build_excel
from pptx_engine   import build_pptx
from pdf_engine    import build_pdf
from docx_engine   import build_docx
from ai_extractor  import (
    extract_excel_schema, extract_pptx_schema,
    extract_pdf_schema, extract_docx_schema
)

app = FastAPI(
    title="OrchestrIQ Generation Service",
    version="3.0.0",
    description="CFO/Board-grade Excel, PowerPoint, PDF, Word generation"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class GenerateRequest(BaseModel):
    objective: str
    company_context: Optional[str] = ""
    available_data: Optional[str] = ""
    currency: Optional[str] = "INR"
    currency_symbol: Optional[str] = "₹"
    api_key: Optional[str] = ""

@app.get("/health")
def health():
    return {"status": "ok", "service": "OrchestrIQ Generation Service v3", "version": "3.0.0"}

@app.get("/")
def root():
    return {"status": "ok", "endpoints": ["/generate/excel", "/generate/pptx", "/generate/pdf", "/generate/docx"]}

@app.post("/generate/excel")
async def generate_excel(req: GenerateRequest):
    try:
        schema = await extract_excel_schema(
            req.objective, req.company_context or "",
            req.available_data or "", req.currency or "INR",
            req.currency_symbol or "₹", req.api_key or ""
        )
        file_bytes = build_excel(schema, req.currency_symbol or "₹")
        filename = (schema.get("title","Report") or "Report").replace(" ","-")[:50] + ".xlsx"
        return Response(
            content=file_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/generate/pptx")
async def generate_pptx(req: GenerateRequest):
    try:
        schema = await extract_pptx_schema(
            req.objective, req.company_context or "",
            req.available_data or "", req.currency or "INR",
            req.currency_symbol or "₹", req.api_key or ""
        )
        file_bytes = build_pptx(schema, req.currency_symbol or "₹")
        filename = (schema.get("title","Presentation") or "Presentation").replace(" ","-")[:50] + ".pptx"
        return Response(
            content=file_bytes,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/generate/pdf")
async def generate_pdf(req: GenerateRequest):
    try:
        schema = await extract_pdf_schema(
            req.objective, req.company_context or "",
            req.available_data or "", req.currency or "INR",
            req.currency_symbol or "₹", req.api_key or ""
        )
        file_bytes = build_pdf(schema, req.currency_symbol or "₹")
        filename = (schema.get("title","Report") or "Report").replace(" ","-")[:50] + ".pdf"
        return Response(
            content=file_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/generate/docx")
async def generate_docx(req: GenerateRequest):
    try:
        schema = await extract_docx_schema(
            req.objective, req.company_context or "",
            req.available_data or "", req.currency or "INR",
            req.currency_symbol or "₹", req.api_key or ""
        )
        file_bytes = build_docx(schema, req.currency_symbol or "₹")
        filename = (schema.get("title","Document") or "Document").replace(" ","-")[:50] + ".docx"
        return Response(
            content=file_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))
