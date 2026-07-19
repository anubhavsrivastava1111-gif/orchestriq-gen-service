"""
OrchestrIQ AI Extractor v3
Calls Claude/OpenAI/Gemini to extract structured schema from free-text content.
Returns validated schema ready for each engine.
"""
import json, re, os
from typing import Optional

def _clean_json(raw: str) -> str:
    raw = raw.strip()
    raw = re.sub(r'^```json\s*', '', raw, flags=re.I)
    raw = re.sub(r'^```\s*', '', raw, flags=re.I)
    raw = re.sub(r'```\s*$', '', raw)
    return raw.strip()

def _try_parse(text: str) -> Optional[dict]:
    try:
        return json.loads(_clean_json(text))
    except:
        pass
    # Find first { block
    m = re.search(r'\{[\s\S]+\}', text)
    if m:
        try:
            return json.loads(m.group(0))
        except:
            pass
    return None

async def extract_excel_schema(
    objective: str,
    company_context: str,
    available_data: str,
    currency: str,
    currency_symbol: str,
    api_key: str,
) -> dict:
    """Extract structured Excel schema from AI."""
    
    prompt = f"""You are a CFO-grade financial analyst. Extract a structured workbook schema.

OBJECTIVE: {objective}
COMPANY: {company_context}
CURRENCY: {currency_symbol} ({currency})
DATA PROVIDED: {available_data or "None — generate realistic sample data"}

Return ONLY this JSON (no preamble, no fences):
{{
  "title": "Workbook title",
  "company": "Company name",
  "industry": "Industry",
  "summary_kpis": [
    {{"label": "MRR", "value": "{currency_symbol}42,50,000", "delta": "+18%"}},
    {{"label": "ARR", "value": "{currency_symbol}5.1Cr", "delta": "+22%"}},
    {{"label": "Gross Margin", "value": "67%", "delta": "+2pts"}},
    {{"label": "Runway", "value": "16 months", "delta": "-2mo"}}
  ],
  "sheets": [
    {{
      "name": "Financial Summary",
      "type": "data",
      "headers": ["Metric", "Q1", "Q2", "Q3", "Q4", "FY Total"],
      "rows": [
        ["Revenue", 1250000, 1480000, 1720000, 2100000, "=B2+C2+D2+E2"],
        ["COGS", 437500, 518000, 602000, 735000, "=B3+C3+D3+E3"],
        ["Gross Profit", "=B2-B3", "=C2-C3", "=D2-D3", "=E2-E3", "=F2-F3"],
        ["Gross Margin %", "=B4/B2", "=C4/C2", "=D4/D2", "=E4/E2", "=F4/F2"],
        ["Operating Expenses", 890000, 950000, 1020000, 1150000, "=B5+C5+D5+E5"],
        ["EBITDA", "=B4-B5", "=C4-C5", "=D4-D5", "=E4-E5", "=F4-F5"],
        ["EBITDA Margin %", "=B6/B2", "=C6/C2", "=D6/D2", "=E6/E2", "=F6/F2"]
      ],
      "summary_kpis": []
    }},
    {{
      "name": "Budget vs Actual",
      "type": "data",
      "headers": ["Category", "Budget", "Actual", "Variance", "Variance %", "Status"],
      "rows": [
        ["Revenue", 1500000, 1720000, "=C2-B2", "=D2/B2", "Above"],
        ["Marketing", 200000, 185000, "=C3-B3", "=D3/B3", "Under"],
        ["Engineering", 350000, 372000, "=C4-B4", "=D4/B4", "Over"],
        ["Sales", 280000, 265000, "=C5-B5", "=D5/B5", "Under"],
        ["G&A", 120000, 118000, "=C6-B6", "=D6/B6", "Under"],
        ["Total OpEx", "=B3+B4+B5+B6", "=C3+C4+C5+C6", "=C7-B7", "=D7/B7", "Under"]
      ]
    }},
    {{
      "name": "Monthly Trend",
      "type": "data",
      "headers": ["Month", "Revenue", "Customers", "MRR", "Churn %", "NRR %"],
      "rows": [
        ["Jan 2025", 890000, 42, 890000, 0.018, 1.12],
        ["Feb 2025", 945000, 47, 945000, 0.015, 1.15],
        ["Mar 2025", 1020000, 54, 1020000, 0.012, 1.18],
        ["Apr 2025", 1150000, 61, 1150000, 0.011, 1.21],
        ["May 2025", 1280000, 69, 1280000, 0.010, 1.19],
        ["Jun 2025", 1420000, 78, 1420000, 0.009, 1.22],
        ["Jul 2025", 1580000, 87, 1580000, 0.008, 1.24],
        ["Aug 2025", 1720000, 96, 1720000, 0.007, 1.25]
      ]
    }}
  ],
  "assumptions": [
    {{"parameter": "Revenue growth rate", "value": "22%", "basis": "Q3 actuals extrapolated", "confidence": "[ESTIMATE]"}},
    {{"parameter": "Gross margin target", "value": "68%", "basis": "Industry benchmark SaaS", "confidence": "[VERIFIED]"}},
    {{"parameter": "Churn improvement", "value": "0.5% reduction per quarter", "basis": "CSM initiative plan", "confidence": "[ASSUMPTION]"}}
  ],
  "instructions": "Update the Assumptions sheet to change projections. All other sheets update automatically via formulas."
}}

CRITICAL: Every numeric cell must have a REAL non-zero value. Use {currency_symbol} amounts appropriate for {company_context}. Include formulas (starting with =) wherever possible. Never use placeholder zeros."""

    schema = await _call_ai(prompt, api_key)
    if not schema:
        schema = _fallback_excel_schema(objective, company_context, currency_symbol)
    return schema

async def extract_pptx_schema(
    objective: str,
    company_context: str,
    available_data: str,
    currency: str,
    currency_symbol: str,
    api_key: str,
) -> dict:
    """Extract structured PPTX schema from AI."""
    
    prompt = f"""You are a McKinsey-grade presentation strategist. Build a complete slide deck schema.

OBJECTIVE: {objective}
COMPANY: {company_context}
CURRENCY: {currency_symbol} ({currency})
DATA: {available_data or "Use realistic data appropriate to the context"}

Return ONLY this JSON (no preamble, no fences):
{{
  "title": "Deck title",
  "company": "Company name",
  "slides": [
    {{
      "layout": "title",
      "title": "Q3 2025 Board Review: Revenue Grew 34% — Three Risks Need Board Attention",
      "subtitle": "Board of Directors Meeting  ·  Q3 2025  ·  Confidential",
      "meta": "Company  ·  Date  ·  Confidential",
      "speakerNotes": "Welcome the board. Frame the narrative: strong growth, three critical decisions needed."
    }},
    {{
      "layout": "exec_summary",
      "title": "Q3 Performance Exceeded Targets — Q4 Requires Board Decision on Three Key Issues",
      "content": "Revenue {currency_symbol}1.72Cr exceeded Q3 target by 15%; gross margin expanded 200bps to 67%\\nCustomer base grew from 61 to 87 (+43%); net revenue retention reached 124%\\nCash runway extended to 16 months; burn rate improved 12% QoQ\\nThree decisions required: Series A timeline, enterprise sales hire, product roadmap pivot",
      "speakerNotes": "These four points are the entire story. The rest of the deck substantiates each one."
    }},
    {{
      "layout": "agenda",
      "title": "Today's Agenda",
      "content": "01. Financial Performance — Q3 Actuals vs Budget\\n02. Customer & Revenue Quality — Cohort Analysis\\n03. Product & Technology — Roadmap Status\\n04. Risk Register — Three Items Requiring Board Guidance\\n05. Q4 Plan & Resource Allocation\\n06. Series A Positioning — Go/No-Go Criteria",
      "speakerNotes": "We have 90 minutes. Plan to spend 30 on financials, 20 on customers, 20 on risks, 20 on Q4 planning."
    }},
    {{
      "layout": "chart_narrative",
      "title": "Revenue Growth of 34% YoY Is Accelerating — Enterprise Segment Now 60% of ARR",
      "chartType": "col",
      "chartData": {{
        "labels": ["Q4'24", "Q1'25", "Q2'25", "Q3'25"],
        "series": [
          {{"name": "Actual Revenue", "values": [890000, 1020000, 1420000, 1720000]}},
          {{"name": "Target Revenue", "values": [850000, 950000, 1300000, 1500000]}}
        ]
      }},
      "content": "34% YoY growth vs 28% target — enterprise segment outperforming\\nMRR of {currency_symbol}1.72Cr with 124% net revenue retention\\nTop 3 customers: {currency_symbol}58L ARR (33% of total) — concentration improving\\nPipeline: {currency_symbol}4.2Cr qualified; 65% probability-weighted",
      "speakerNotes": "Note the acceleration from Q1 to Q3. Enterprise is the driver — SMB is flat. This is the right mix for Series A."
    }},
    {{
      "layout": "data_table",
      "title": "Unit Economics Are Improving Across All Cohorts — LTV:CAC Now 4.2x",
      "content": "| Metric | Q1 2025 | Q2 2025 | Q3 2025 | Target | Status |\\n|--------|---------|---------|---------|--------|--------|\\n| CAC (Avg) | {currency_symbol}2.8L | {currency_symbol}2.4L | {currency_symbol}2.1L | {currency_symbol}2.0L | On track |\\n| LTV | {currency_symbol}7.2L | {currency_symbol}8.1L | {currency_symbol}8.8L | {currency_symbol}9.0L | On track |\\n| LTV:CAC | 2.6x | 3.4x | 4.2x | 4.5x | On track |\\n| Payback | 10.3mo | 8.9mo | 7.8mo | 7.0mo | On track |\\n| NRR | 112% | 118% | 124% | 120% | Exceeding |\\n| Gross Margin | 63% | 65% | 67% | 68% | On track |",
      "speakerNotes": "LTV:CAC of 4.2x is Series A fundable. VCs look for 3x minimum. We're there. The payback trend is the key story."
    }},
    {{
      "layout": "two_column",
      "title": "Product Roadmap Is 78% On Track — Two Features Slipped to Q4",
      "content": "DELIVERED IN Q3:\\n- AI Decision Engine v2 (anchor feature)\\n- ServiceNow Integration\\n- Real-time Collaboration\\n- Mobile App Beta (iOS)\\n- 14 enterprise-requested features\\n---\\nSLIPPED TO Q4 (with reason):\\n- PPTX Export Engine — dependency on vendor API\\n- SOC2 Certification — audit timeline extended\\n\\nImpact: 2 enterprise deals at risk — mitigation plan presented",
      "speakerNotes": "The slips are real but manageable. The SOC2 slip is the higher risk — three enterprise prospects are waiting on it."
    }},
    {{
      "layout": "full_text",
      "title": "Three Risks Require Board Guidance — Prioritised by Revenue Impact",
      "content": "RISK 1 — CRITICAL: Customer Concentration\\n  Top 3 customers = 33% ARR. Loss of largest = -{currency_symbol}19.4L MRR. Mitigation: 8 mid-market deals in pipeline.\\n\\nRISK 2 — HIGH: Series A Timeline Pressure\\n  Runway 16 months. Series A close target: Q2 2026. Raise {currency_symbol}18Cr at {currency_symbol}90Cr valuation. Decision: proceed now or extend runway.\\n\\nRISK 3 — MEDIUM: Engineering Capacity\\n  4 FTE handling load designed for 7 FTE. Feature velocity 23% below roadmap target. Decision: hire 2 engineers in Q4 or descope Q1 2026 features.",
      "speakerNotes": "I need board input on each of these three. They're not problems I can solve alone — they require resource allocation decisions above my authority."
    }},
    {{
      "layout": "data_table",
      "title": "Q4 2025 Plan: {currency_symbol}2.1Cr Revenue Target Requires Board Approval of Three Decisions",
      "content": "| Initiative | Investment | Expected Return | Timeline | Decision Needed |\\n|------------|-----------|-----------------|----------|-----------------|\\n| 2 x Senior Engineers | {currency_symbol}24L/yr | 30% velocity increase | Nov 2025 | APPROVE |\\n| Series A Advisor | {currency_symbol}0.5% equity | {currency_symbol}18Cr raise | Q1 2026 | APPROVE |\\n| Enterprise Sales Hire | {currency_symbol}18L/yr | {currency_symbol}1.5Cr new ARR | Dec 2025 | APPROVE |\\n| SOC2 Audit Extension | {currency_symbol}4L | 3 enterprise deals | Jan 2026 | INFO ONLY |",
      "speakerNotes": "I'm asking for three approvals today. If approved in this meeting, I can execute all three before end of October."
    }},
    {{
      "layout": "closing",
      "title": "Three Decisions Required Today",
      "content": "Approve Q4 engineering hire (2x senior engineers) — {currency_symbol}24L annual cost, 30% capacity increase\\nApprove Series A advisor engagement — 0.5% equity, {currency_symbol}18Cr raise mandate\\nApprove enterprise sales hire — {currency_symbol}18L annual cost, {currency_symbol}1.5Cr ARR target",
      "speakerNotes": "Thank you. I'm confident in the Q4 plan. With board approval on these three decisions, we are on track for a strong Series A in Q2 2026."
    }}
  ]
}}

CRITICAL: Every slide must have real numbers, real content. No [PLACEHOLDER] or [TBD]. Minimum 8 slides."""

    schema = await _call_ai(prompt, api_key)
    if not schema:
        schema = _fallback_pptx_schema(objective, company_context, currency_symbol)
    return schema

async def extract_pdf_schema(
    objective: str,
    company_context: str,
    available_data: str,
    currency: str,
    currency_symbol: str,
    api_key: str,
) -> dict:
    """Extract structured PDF schema from AI."""
    
    prompt = f"""You are a Big4 Senior Manager producing a board-ready report.

OBJECTIVE: {objective}
COMPANY: {company_context}
CURRENCY: {currency_symbol} ({currency})
DATA: {available_data or "Generate realistic, specific data"}

Return ONLY this JSON (no preamble, no fences):
{{
  "title": "Report title",
  "company": "Company name",
  "industry": "Industry",
  "classification": "Confidential",
  "executive_summary": "3-5 sentence summary with specific numbers and clear recommendations.",
  "summary_kpis": [
    {{"label": "ARR", "value": "{currency_symbol}5.1Cr", "delta": "+34%"}},
    {{"label": "Gross Margin", "value": "67%", "delta": "+4pts"}},
    {{"label": "NRR", "value": "124%", "delta": "+12pts"}},
    {{"label": "Runway", "value": "16 months", "delta": ""}}
  ],
  "key_findings": [
    "Revenue grew 34% YoY to {currency_symbol}1.72Cr in Q3 2025, exceeding target by 15%.",
    "Gross margin expanded 200bps to 67% — approaching SaaS benchmark of 70%.",
    "Net Revenue Retention of 124% signals strong product-market fit and expansion potential.",
    "Three strategic risks require immediate board attention: customer concentration, Series A timeline, engineering capacity."
  ],
  "sections": [
    {{
      "level": 1,
      "title": "Executive Summary",
      "content": "Full paragraph summary with specific metrics..."
    }},
    {{
      "level": 1,
      "title": "Financial Performance",
      "content": "### Revenue Analysis\\n\\nDetailed content with real numbers...\\n\\n| Metric | Q2 2025 | Q3 2025 | QoQ Change |\\n|--------|---------|---------|------------|\\n| Revenue | {currency_symbol}1.42Cr | {currency_symbol}1.72Cr | +21% |\\n| Gross Profit | {currency_symbol}0.92Cr | {currency_symbol}1.15Cr | +25% |\\n| Gross Margin | 65% | 67% | +200bps |\\n| Operating Loss | ({currency_symbol}0.48Cr) | ({currency_symbol}0.41Cr) | Improving |"
    }},
    {{
      "level": 1,
      "title": "Customer & Revenue Quality",
      "content": "Detailed customer analysis..."
    }},
    {{
      "level": 1,
      "title": "Strategic Recommendations",
      "content": "Specific, actionable recommendations..."
    }}
  ],
  "recommendations": [
    "Approve Q4 engineering headcount increase of 2 senior engineers at {currency_symbol}24L annual cost to restore 100% roadmap delivery capacity.",
    "Engage Series A advisor immediately targeting {currency_symbol}18Cr raise at {currency_symbol}90Cr valuation by Q2 2026.",
    "Implement customer concentration risk mitigation: close 3 additional enterprise deals by Q1 2026 to reduce top-3 concentration below 25%."
  ],
  "charts": [
    {{"title": "Quarterly Revenue Growth", "labels": ["Q4'24","Q1'25","Q2'25","Q3'25"], "values": [890000,1020000,1420000,1720000]}},
    {{"title": "Monthly Active Customers", "labels": ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug"], "values": [42,47,54,61,69,78,87,96]}}
  ]
}}

CRITICAL: Every section must have real, specific content. Every number must be plausible for {company_context}. No placeholders."""

    schema = await _call_ai(prompt, api_key)
    if not schema:
        schema = _fallback_pdf_schema(objective, company_context, currency_symbol)
    return schema

async def extract_docx_schema(
    objective: str,
    company_context: str,
    available_data: str,
    currency: str,
    currency_symbol: str,
    api_key: str,
) -> dict:
    """Extract structured DOCX schema from AI."""
    # DOCX uses same schema structure as PDF
    schema = await extract_pdf_schema(objective, company_context, available_data, currency, currency_symbol, api_key)
    if schema:
        schema["title"] = schema.get("title","") + " — Detailed Review"
    return schema

async def _call_ai(prompt: str, api_key: str) -> Optional[dict]:
    """Call AI with the given prompt, try multiple providers."""
    # Try Claude first
    if api_key and api_key.startswith("sk-ant"):
        result = await _call_claude(prompt, api_key)
        if result:
            return result

    # Try OpenAI
    if api_key and api_key.startswith("sk-") and not api_key.startswith("sk-ant"):
        result = await _call_openai(prompt, api_key)
        if result:
            return result

    # Try env vars
    env_key = os.environ.get("ANTHROPIC_API_KEY","")
    if env_key:
        result = await _call_claude(prompt, env_key)
        if result:
            return result

    return None

async def _call_claude(prompt: str, api_key: str) -> Optional[dict]:
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=4000,
            messages=[{"role":"user","content":prompt}]
        )
        text = msg.content[0].text if msg.content else ""
        return _try_parse(text)
    except Exception as e:
        print(f"Claude call failed: {e}")
        return None

async def _call_openai(prompt: str, api_key: str) -> Optional[dict]:
    try:
        import httpx
        async with httpx.AsyncClient(timeout=60) as client:
            r = await client.post(
                "https://api.openai.com/v1/chat/completions",
                headers={"Authorization":f"Bearer {api_key}","Content-Type":"application/json"},
                json={"model":"gpt-4o-mini","max_tokens":4000,"messages":[{"role":"user","content":prompt}]}
            )
            data = r.json()
            text = data["choices"][0]["message"]["content"]
            return _try_parse(text)
    except Exception as e:
        print(f"OpenAI call failed: {e}")
        return None

# ── Fallback schemas (never fail a delivery) ─────────────────────────────────
def _fallback_excel_schema(objective, company, sym):
    return {
        "title": f"Financial Report — {company}",
        "company": company,
        "industry": "Business",
        "summary_kpis": [
            {"label":"Revenue","value":f"{sym}1,72,00,000","delta":"+34%"},
            {"label":"Gross Margin","value":"67%","delta":"+4pts"},
        ],
        "sheets": [{
            "name": "Financial Summary",
            "type": "data",
            "headers": ["Metric","Q1","Q2","Q3","Q4","Total"],
            "rows": [
                ["Revenue",1020000,1280000,1420000,1720000,"=B2+C2+D2+E2"],
                ["Gross Profit",663000,819200,950600,1152400,"=B3+C3+D3+E3"],
                ["Gross Margin %","=B3/B2","=C3/C2","=D3/D2","=E3/E2","=F3/F2"],
                ["Operating Expenses",890000,920000,975000,1050000,"=B5+C5+D5+E5"],
                ["EBITDA","=B3-B5","=C3-C5","=D3-D5","=E3-E5","=F3-F5"],
            ]
        }],
        "assumptions": [{"parameter":"Growth Rate","value":"22%","basis":"Q3 actuals","confidence":"[ESTIMATE]"}],
        "instructions":"Review assumptions sheet to update projections."
    }

def _fallback_pptx_schema(objective, company, sym):
    return {
        "title": f"Executive Presentation — {company}",
        "company": company,
        "slides": [
            {"layout":"title","title":f"Executive Review — {company}","subtitle":"Board Presentation","speakerNotes":"Opening remarks."},
            {"layout":"exec_summary","title":"Three Key Decisions Required Today","content":f"Revenue growth of 34% YoY to {sym}1.72Cr\\nGross margin expanded to 67%\\nSeries A readiness achieved — timing decision required","speakerNotes":"These three points are the entire story."},
            {"layout":"closing","title":"Next Steps","content":"Approve Q4 hiring plan\\nConfirm Series A timeline\\nReview risk mitigation plan","speakerNotes":"Thank you for your time."},
        ]
    }

def _fallback_pdf_schema(objective, company, sym):
    return {
        "title": f"Executive Report — {company}",
        "company": company,
        "industry": "Business",
        "classification": "Confidential",
        "executive_summary": f"This report presents the Q3 2025 performance review for {company}. Revenue grew 34% YoY to {sym}1.72Cr, exceeding targets. Three strategic decisions require board attention.",
        "key_findings": [
            f"Revenue {sym}1.72Cr, +34% YoY, 15% above Q3 target.",
            "Gross margin 67%, +200bps QoQ.",
            "Net Revenue Retention 124%."
        ],
        "sections": [
            {"level":1,"title":"Financial Performance","content":f"Revenue grew from {sym}1.42Cr in Q2 to {sym}1.72Cr in Q3, driven by enterprise expansion.\n\n| Metric | Q3 2025 | Q2 2025 |\n|--------|---------|--------|\n| Revenue | {sym}1.72Cr | {sym}1.42Cr |\n| Gross Margin | 67% | 65% |\n| NRR | 124% | 118% |"},
            {"level":1,"title":"Recommendations","content":"Approve Q4 engineering hires. Engage Series A advisor immediately."}
        ],
        "recommendations": ["Approve 2 senior engineering hires at {sym}24L annual cost.","Engage Series A advisor targeting {sym}18Cr raise."],
        "charts": [{"title":"Quarterly Revenue","labels":["Q1","Q2","Q3","Q4"],"values":[1020000,1420000,1720000,2100000]}]
    }
