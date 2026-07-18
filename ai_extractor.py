"""
AI Parameter Extractor
Uses Claude to extract precise, structured financial parameters from natural language.
The Python engines use these parameters deterministically — AI does intelligence,
Python does construction. This is the architecture that produces CFO-grade output.
"""

import json
import re
import anthropic
from typing import Any, Dict

# ─── EXTRACTION PROMPTS BY DOC TYPE ──────────────────────────────────────────

EXCEL_EXTRACTION_PROMPT = """You are a CFO-grade financial analyst extracting parameters from a business objective.
Your output will be used by a deterministic Python engine to build a professional Excel workbook.
You MUST return only a valid JSON object. No preamble. No explanation. No markdown fences.

EXTRACTION RULES:
1. Extract or infer ALL financial numbers from the context. If not provided, use realistic industry benchmarks.
2. Numbers must be raw integers/floats — NEVER strings like "₹5L" or "$500K". Write 500000.
3. Percentages as fractions: 15% = 0.15. NEVER 15 or "15%".
4. All currency values in the stated currency unit (full amount, not lakhs/crores shorthand).
5. Choose template_type based on the objective — do not default to financial_dashboard if something more specific fits.
6. num_periods should be 24 for monthly projections (2 years). Use 12 if the user specifies 1 year.
7. revenue_growth_rate_monthly: if user says "15% annual growth", convert: (1.15)^(1/12)-1 ≈ 0.0117.
8. If multiple revenue streams are mentioned, list them separately. If not mentioned, create 1 realistic stream.
9. Extract company name from context. Use "Your Company" if not found.
10. The filename field must be URL-safe (hyphens, no spaces, no special chars).

TEMPLATE TYPES (choose the most specific fit):
- financial_dashboard: P&L + Cash Flow + Balance Sheet. Use for general "financial model", "dashboard", "management accounts"
- saas_metrics: MRR, Churn, LTV, CAC, ARR. Use for SaaS, subscription, recurring revenue businesses
- runway_planner: Burn rate, cash position, funding runway. Use for "runway", "burn rate", "how long will cash last"
- budget_vs_actual: Two-column comparison. Use for "budget", "forecast vs actual", "variance analysis"
- unit_economics: Per-unit profitability. Use for "unit economics", "contribution margin", "per-customer economics"
- headcount_model: People costs. Use for "headcount", "org chart", "staffing plan", "salary budget"
- general_model: Flexible template for anything that doesn't fit the above

OUTPUT SCHEMA (return this exact structure):
{
  "template_type": "financial_dashboard",
  "company_name": "Acme Corp",
  "title": "Financial Dashboard — Acme Corp FY2025-26",
  "subtitle": "Prepared for Board Review · July 2025 · Management Estimates",
  "filename": "Acme-Corp-Financial-Dashboard-FY2025",
  "currency": "INR",
  "currency_symbol": "₹",
  "num_periods": 24,
  "start_month": "Jan 2025",
  "assumptions": {
    "revenue_base_monthly": 500000,
    "revenue_growth_rate_monthly": 0.05,
    "revenue_streams": [
      {"name": "Product Sales", "monthly_base": 350000, "growth_rate": 0.05},
      {"name": "Services", "monthly_base": 150000, "growth_rate": 0.08}
    ],
    "cogs_pct": 0.35,
    "salaries_annual": 3600000,
    "marketing_annual": 600000,
    "rent_annual": 360000,
    "technology_annual": 240000,
    "other_opex_annual": 300000,
    "opening_cash": 5000000,
    "ar_days": 30,
    "ap_days": 45,
    "capex_annual": 500000,
    "depreciation_years": 5,
    "tax_rate": 0.25,
    "headcount": 12
  },
  "saas_metrics": {
    "mrr_base": 0,
    "monthly_churn_rate": 0,
    "new_customers_monthly": 0,
    "cac": 0,
    "arpu_monthly": 0,
    "ltv_months": 0
  },
  "runway_inputs": {
    "current_cash": 0,
    "monthly_burn_categories": []
  },
  "budget_data": {
    "categories": [],
    "budget_values": [],
    "actual_values": []
  },
  "key_context": "Any important additional context the engine should know",
  "sections": []
}"""

PPTX_EXTRACTION_PROMPT = """You are a McKinsey-grade analyst creating a structured presentation outline.
Extract content parameters for a professional PowerPoint deck.
Return ONLY valid JSON. No preamble. No fences.

OUTPUT SCHEMA:
{
  "title": "Board Presentation — Q2 FY2025",
  "subtitle": "Strategic Review & Financial Performance",
  "company_name": "Acme Corp",
  "filename": "Acme-Corp-Q2-Board-Presentation",
  "currency": "INR",
  "currency_symbol": "₹",
  "presenter": "Management Team",
  "date": "July 2025",
  "audience": "Board of Directors",
  "classification": "Confidential",
  "executive_summary": {
    "headline": "Revenue grew 34% YoY; runway extends to 18 months post Series A",
    "key_points": [
      "₹7.2Cr revenue in Q2, up 34% YoY, driven by enterprise segment",
      "Gross margin expanded 400bps to 68% on operational leverage",
      "Cash position ₹22Cr; 18-month runway with current burn of ₹1.2Cr/month"
    ]
  },
  "financial_data": {
    "revenue": [4200000, 4800000, 5500000, 6100000, 6800000, 7200000],
    "gross_profit": [2700000, 3100000, 3600000, 4000000, 4500000, 4900000],
    "ebitda": [420000, 580000, 770000, 920000, 1100000, 1260000],
    "net_profit": [300000, 430000, 600000, 750000, 900000, 1080000],
    "cash": [25000000, 23800000, 22600000, 21400000, 20200000, 22000000],
    "period_labels": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
    "kpis": [
      {"label": "Revenue Q2", "value": "₹7.2Cr", "change": "+34% YoY", "status": "good"},
      {"label": "Gross Margin", "value": "68%", "change": "+400bps", "status": "good"},
      {"label": "EBITDA Margin", "value": "17.5%", "change": "+6pp", "status": "good"},
      {"label": "Runway", "value": "18 months", "change": "Post Series A", "status": "neutral"},
      {"label": "Burn Rate", "value": "₹1.2Cr/mo", "change": "-8% MoM", "status": "good"},
      {"label": "Headcount", "value": "47 FTE", "change": "+12 QoQ", "status": "neutral"}
    ]
  },
  "slides": [
    {
      "layout": "title",
      "title": "Q2 FY2025 Board Review",
      "subtitle": "Financial Performance & Strategic Update"
    },
    {
      "layout": "exec_summary",
      "title": "Revenue grew 34% YoY; EBITDA positive for second consecutive quarter",
      "bullets": ["Point 1", "Point 2", "Point 3"]
    },
    {
      "layout": "kpi_dashboard",
      "title": "Key Performance Indicators — Q2 FY2025"
    },
    {
      "layout": "revenue_chart",
      "title": "Revenue trajectory on track; enterprise mix improving",
      "insight": "Enterprise segment now 62% of revenue, up from 45% in Q2 FY24",
      "bullets": ["Insight 1", "Insight 2", "Insight 3"]
    },
    {
      "layout": "pl_table",
      "title": "P&L Summary — strong margin expansion across all levels",
      "table_data": {
        "headers": ["Metric", "Q1 FY25", "Q2 FY25", "QoQ", "Q2 FY24", "YoY"],
        "rows": [
          ["Revenue (₹Cr)", "5.9", "7.2", "+22%", "5.4", "+34%"],
          ["Gross Profit (₹Cr)", "3.8", "4.9", "+29%", "3.4", "+44%"],
          ["Gross Margin %", "64%", "68%", "+4pp", "63%", "+5pp"],
          ["EBITDA (₹Cr)", "0.82", "1.26", "+54%", "0.31", "+3.1x"],
          ["Net Profit (₹Cr)", "0.60", "1.08", "+80%", "0.18", "+5.6x"]
        ]
      }
    },
    {
      "layout": "cash_runway",
      "title": "Cash position secured; 18-month runway post-Series A close",
      "insight": "Net cash build of ₹1.8Cr in Q2; positive free cash flow expected from Q3"
    },
    {
      "layout": "full_text",
      "title": "Strategic priorities for H2 FY2025",
      "bullets": ["Priority 1", "Priority 2", "Priority 3", "Priority 4"]
    },
    {
      "layout": "next_steps",
      "title": "Decisions required from the Board",
      "bullets": ["Decision 1", "Decision 2", "Decision 3"]
    },
    {
      "layout": "closing",
      "title": "Thank You"
    }
  ]
}"""

PDF_EXTRACTION_PROMPT = """You are a Big4-grade analyst creating a structured report.
Extract content for a professional PDF report. Return ONLY valid JSON. No preamble. No fences.

OUTPUT SCHEMA:
{
  "title": "Business Performance Review — Q2 FY2025",
  "subtitle": "Prepared for Executive Management",
  "company_name": "Acme Corp",
  "filename": "Acme-Corp-Q2-Performance-Review",
  "classification": "Confidential",
  "date": "July 2025",
  "currency": "INR",
  "currency_symbol": "₹",
  "executive_summary": "3-5 sentence precise summary with numbers.",
  "key_findings": [
    "Finding 1 with specific metric",
    "Finding 2 with specific metric",
    "Finding 3 with specific metric"
  ],
  "recommendations": [
    "Recommendation 1 — specific and actionable",
    "Recommendation 2 — specific and actionable"
  ],
  "sections": [
    {
      "title": "Financial Performance",
      "level": 1,
      "content": "Detailed content in markdown. Use **bold** for key terms. Use tables with | syntax.",
      "tables": [
        {
          "title": "P&L Summary (₹ Lakhs)",
          "headers": ["Metric", "Q1 FY25", "Q2 FY25", "Change"],
          "rows": [
            ["Revenue", "59.0", "72.0", "+22%"],
            ["Gross Profit", "38.0", "49.0", "+29%"]
          ]
        }
      ]
    },
    {
      "title": "Revenue Analysis",
      "level": 1,
      "content": "Detailed content here."
    }
  ],
  "appendices": [
    {
      "title": "Detailed Financial Tables",
      "content": "Supporting data here."
    }
  ]
}"""

DOCX_EXTRACTION_PROMPT = """You are a management consultant creating a structured Word document.
Extract content for a professional Word document. Return ONLY valid JSON. No preamble. No fences.

Use the same schema as the PDF extraction prompt but with document-appropriate content.
OUTPUT SCHEMA: Same as PDF schema above."""

PROMPTS = {
    "excel": EXCEL_EXTRACTION_PROMPT,
    "pptx": PPTX_EXTRACTION_PROMPT,
    "pdf": PDF_EXTRACTION_PROMPT,
    "docx": DOCX_EXTRACTION_PROMPT,
}


# ─── EXTRACTOR ────────────────────────────────────────────────────────────────

async def extract_parameters(
    objective: str,
    context: str,
    data: str,
    currency: str,
    currency_symbol: str,
    doc_type: str,
    api_key: str,
) -> Dict[str, Any]:
    """
    Call Claude to extract structured parameters from natural language.
    Returns a dict that the appropriate engine uses deterministically.
    """
    client = anthropic.Anthropic(api_key=api_key)
    system_prompt = PROMPTS.get(doc_type, PROMPTS["excel"])

    user_message = f"""OBJECTIVE: {objective}

COMPANY CONTEXT:
{context or "Not provided — use realistic defaults for an Indian startup."}

AVAILABLE DATA / NUMBERS:
{data or "No specific data provided. Infer realistic parameters from the objective and context."}

CURRENCY: {currency_symbol} ({currency})

Extract all parameters and return the JSON object now.
Remember: ALL numbers must be raw integers/floats. ALL percentages as fractions (0.15 = 15%).
The output will be used directly by a Python code generator — accuracy is critical."""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
    )

    raw = response.content[0].text.strip()

    # Aggressive JSON extraction — 4 strategies
    params = _try_parse(raw)
    if not params:
        m = re.search(r"\{[\s\S]*\}", raw)
        if m:
            params = _try_parse(m.group(0))
    if not params:
        params = _repair_json(raw)
    if not params:
        # Ultimate fallback — return safe defaults
        params = _safe_defaults(objective, currency, currency_symbol)

    # Ensure critical fields exist
    params.setdefault("currency", currency)
    params.setdefault("currency_symbol", currency_symbol)
    params.setdefault("company_name", "Your Company")
    params.setdefault("title", objective[:80])
    params.setdefault("filename", _safe_filename(objective))

    return params


def _try_parse(s: str):
    try:
        return json.loads(s)
    except Exception:
        return None


def _repair_json(text: str):
    """Best-effort repair of truncated JSON."""
    try:
        s = text.strip()
        start = s.find("{")
        if start == -1:
            return None
        s = s[start:]
        stack, in_str, esc, last_good = [], False, False, 0
        for i, ch in enumerate(s):
            if esc:
                esc = False
                continue
            if ch == "\\":
                esc = True
                continue
            if ch == '"':
                in_str = not in_str
                continue
            if in_str:
                continue
            if ch == "{":
                stack.append("}")
            elif ch == "[":
                stack.append("]")
            elif ch in "}]":
                if stack:
                    stack.pop()
                if not stack:
                    last_good = i + 1
                    break
            if ch in "},]":
                last_good = i + 1
        body = s[: last_good or len(s)].rstrip(",")
        closers = "".join(reversed(stack))
        return _try_parse(body + closers)
    except Exception:
        return None


def _safe_filename(text: str) -> str:
    name = re.sub(r"[^\w\s-]", "", text)
    name = re.sub(r"\s+", "-", name.strip())
    return name[:60] or "Document"


def _safe_defaults(objective: str, currency: str, sym: str) -> dict:
    """Minimal safe defaults when AI extraction fails."""
    return {
        "template_type": "financial_dashboard",
        "company_name": "Your Company",
        "title": f"Financial Model — {objective[:50]}",
        "subtitle": "Management Estimates",
        "filename": _safe_filename(objective),
        "currency": currency,
        "currency_symbol": sym,
        "num_periods": 24,
        "start_month": "Jan 2025",
        "assumptions": {
            "revenue_base_monthly": 1000000,
            "revenue_growth_rate_monthly": 0.05,
            "revenue_streams": [{"name": "Revenue", "monthly_base": 1000000, "growth_rate": 0.05}],
            "cogs_pct": 0.35,
            "salaries_annual": 3600000,
            "marketing_annual": 600000,
            "rent_annual": 360000,
            "technology_annual": 240000,
            "other_opex_annual": 300000,
            "opening_cash": 5000000,
            "ar_days": 30,
            "ap_days": 45,
            "capex_annual": 500000,
            "depreciation_years": 5,
            "tax_rate": 0.25,
            "headcount": 10,
        },
        "saas_metrics": {"mrr_base": 0, "monthly_churn_rate": 0, "new_customers_monthly": 0, "cac": 0, "arpu_monthly": 0},
        "key_context": objective,
    }
