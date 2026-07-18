"""
OrchestrIQ Generation Service
FastAPI backend for CFO/CEO-grade document generation.
Deploy to Railway. Set ANTHROPIC_API_KEY in environment variables.
"""

import io
import os
import traceback
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Optional

from ai_extractor import extract_parameters
from excel_engine import build_excel
from pptx_engine import build_pptx
from pdf_engine import build_pdf
from docx_engine import build_docx

app = FastAPI(title="OrchestrIQ Generation Service", version="1.0.0")

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
    api_key: Optional[str] = None  # user BYOK; falls back to env key


def get_api_key(request: GenerateRequest) -> str:
    key = request.api_key or os.getenv("ANTHROPIC_API_KEY", "")
    if not key:
        raise HTTPException(status_code=400, detail="No API key provided.")
    return key


@app.get("/health")
def health():
    return {"status": "ok", "service": "OrchestrIQ Generation Service"}


@app.post("/generate/excel")
async def generate_excel(req: GenerateRequest):
    try:
        api_key = get_api_key(req)
        params = await extract_parameters(
            objective=req.objective,
            context=req.company_context,
            data=req.available_data,
            currency=req.currency,
            currency_symbol=req.currency_symbol,
            doc_type="excel",
            api_key=api_key,
        )
        file_bytes = build_excel(params)
        filename = f"{params.get('filename', 'Financial-Model')}.xlsx"
        return StreamingResponse(
            io.BytesIO(file_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/generate/pptx")
async def generate_pptx(req: GenerateRequest):
    try:
        api_key = get_api_key(req)
        params = await extract_parameters(
            objective=req.objective,
            context=req.company_context,
            data=req.available_data,
            currency=req.currency,
            currency_symbol=req.currency_symbol,
            doc_type="pptx",
            api_key=api_key,
        )
        file_bytes = build_pptx(params)
        filename = f"{params.get('filename', 'Presentation')}.pptx"
        return StreamingResponse(
            io.BytesIO(file_bytes),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/generate/pdf")
async def generate_pdf(req: GenerateRequest):
    try:
        api_key = get_api_key(req)
        params = await extract_parameters(
            objective=req.objective,
            context=req.company_context,
            data=req.available_data,
            currency=req.currency,
            currency_symbol=req.currency_symbol,
            doc_type="pdf",
            api_key=api_key,
        )
        file_bytes = build_pdf(params)
        filename = f"{params.get('filename', 'Report')}.pdf"
        return StreamingResponse(
            io.BytesIO(file_bytes),
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/generate/docx")
async def generate_docx(req: GenerateRequest):
    try:
        api_key = get_api_key(req)
        params = await extract_parameters(
            objective=req.objective,
            context=req.company_context,
            data=req.available_data,
            currency=req.currency,
            currency_symbol=req.currency_symbol,
            doc_type="docx",
            api_key=api_key,
        )
        file_bytes = build_docx(params)
        filename = f"{params.get('filename', 'Document')}.docx"
        return StreamingResponse(
            io.BytesIO(file_bytes),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", 8001)))
