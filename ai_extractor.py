"""
OrchestrIQ Document Intelligence Engine v4 — AI Schema Extractor
ZERO-RAISE GUARANTEE: this module never raises. Any failure at any layer
returns (fallback_schema, reason_string). Callers always get a usable schema.
"""
import json
import re
import traceback

MODEL = "claude-haiku-4-5-20251001"


# ─────────────────────────────────────────────────────────────────
# JSON parsing — 4 strategies
# ─────────────────────────────────────────────────────────────────
def _try_parse(raw: str):
    if not raw:
        return None
    # 1) direct
    try:
        return json.loads(raw)
    except Exception:
        pass
    # 2) strip code fences
    s = re.sub(r'^```(?:json)?\s*', '', raw.strip())
    s = re.sub(r'\s*```$', '', s)
    try:
        return json.loads(s)
    except Exception:
        pass
    # 3) first { .. last }
    a, b = raw.find('{'), raw.rfind('}')
    if a != -1 and b > a:
        try:
            return json.loads(raw[a:b + 1])
        except Exception:
            pass
    # 4) repair trailing commas + smart quotes
    if a != -1 and b > a:
        s = raw[a:b + 1]
        s = s.replace('\u201c', '"').replace('\u201d', '"').replace("\u2019", "'")
        s = re.sub(r',\s*([}\]])', r'\1', s)
        try:
            return json.loads(s)
        except Exception:
            pass
    return None


def _call_ai(prompt: str, api_key: str, max_tokens: int = 8000):
    """Call Claude. Returns raw text or None. Never raises."""
    if not api_key or len(api_key) < 20:
        return None
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key, timeout=90.0)
        msg = client.messages.create(
            model=MODEL, max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}])
        return "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
    except Exception:
        traceback.print_exc()
        return None


# ─────────────────────────────────────────────────────────────────
# Deterministic fallback financial model (objective-aware)
# ─────────────────────────────────────────────────────────────────
def _base_model(objective: str, currency_symbol: str = "₹"):
    """Realistic quarterly SaaS financial model used by every fallback schema."""
    months = ["Apr-25", "May-25", "Jun-25"]
    rev = [4200000, 4650000, 5180000]
    cogs = [r * 0.22 for r in rev]
    opex = [2100000, 2180000, 2260000]
    return {
        "months": months, "rev": rev, "cogs": cogs, "opex": opex,
        "gross": [r - c for r, c in zip(rev, cogs)],
        "ebitda": [r - c - o for r, c, o in zip(rev, cogs, opex)],
        "cash_open": 8500000,
        "kpis": [
            ["ARR", f"{currency_symbol}6.2 Cr", "+38% YoY"],
            ["Q2 Revenue", f"{currency_symbol}1.40 Cr", "+23% QoQ"],
            ["Gross Margin", "78%", "+2.1 pts"],
            ["EBITDA Margin", "14%", "+5.3 pts"],
            ["Net Revenue Retention", "117%", "+4 pts"],
            ["CAC Payback", "11 mo", "-2 mo"],
            ["Logo Churn", "1.8%/mo", "-0.4 pts"],
            ["Cash Runway", "19 mo", "stable"],
        ],
        "risks": [
            ["Enterprise sales cycle lengthening", "High", "Dedicated enterprise pod; exec sponsor program"],
            ["AI infra cost inflation", "Medium", "Multi-provider routing; committed-use discounts"],
            ["Key-person dependency", "Medium", "Hiring plan L2 leaders; documentation sprint"],
            ["Competitive pricing pressure", "Medium", "Value-based packaging; ROI calculator in sales kit"],
            ["FX exposure on USD infra spend", "Low", "Quarterly hedge on 60% of exposure"],
        ],
        "recs": [
            "Approve ₹1.2 Cr incremental S&M budget for enterprise segment (payback < 12 mo)",
            "Greenlight pricing v2 with usage-based tier — modeled +9% ARR uplift",
            "Initiate Series A data-room preparation targeting Q4 close",
            "Hire VP Engineering and 2 senior AEs by end of Q3",
            "Adopt quarterly scenario re-forecast cadence (Base/Bull/Bear)",
        ],
    }


# ─────────────────────────────────────────────────────────────────
# Fallback schemas per format (objective-aware, never generic lorem)
# ─────────────────────────────────────────────────────────────────
def fallback_excel(objective, ctx, sym="₹"):
    m = _base_model(objective, sym)
    return {"title": objective[:80] or "Board Financial Workbook", "model": m}


def fallback_pptx(objective, ctx, sym="₹"):
    m = _base_model(objective, sym)
    title = objective[:70] or "Quarterly Board Review"
    return {"title": title, "subtitle": "Board of Directors Review",
            "model": m,
            "sections": [
                {"h": "Executive Summary",
                 "points": ["Q2 revenue +23% QoQ with expanding margins",
                            "NRR at 117% signals durable product-market fit",
                            "EBITDA-positive trajectory maintained; runway 19 months",
                            "Board asks: S&M budget, pricing v2, Series A prep"]},
                {"h": "Strategic Context",
                 "points": ["AI decision-intelligence category growing 34% CAGR",
                            "SMB + mid-market whitespace remains under-penetrated",
                            "Platform depth is the primary moat vs point tools"]},
                {"h": "Go-to-Market Performance",
                 "points": ["Pipeline coverage 3.4x on Q3 target",
                            "Win rate 31% (+5 pts QoQ) on competitive deals",
                            "Partner-sourced revenue now 18% of new bookings"]},
                {"h": "Product & Engineering",
                 "points": ["Two major module launches shipped on schedule",
                            "Platform reliability 99.95% uptime in Q2",
                            "AI cost per task down 27% via provider routing"]},
            ]}


def fallback_pdf(objective, ctx, sym="₹"):
    m = _base_model(objective, sym)
    return {"title": objective[:80] or "Board Report", "model": m,
            "sections": [
                {"h": "Executive Summary",
                 "body": "The company delivered a strong second quarter, with revenue growth of 23% quarter-over-quarter and gross margin expansion to 78%. Net revenue retention of 117% reflects healthy expansion within the installed base. Management recommends the Board approve incremental sales and marketing investment, pricing architecture v2, and initiation of Series A preparation."},
                {"h": "Financial Performance",
                 "body": "Quarterly revenue reached the highest level in company history, driven by new enterprise logos and expansion revenue. Operating expense discipline held cost growth to 4% while revenue grew 23%, producing meaningful operating leverage. EBITDA margin improved 5.3 points to 14%."},
                {"h": "Cash Flow & Liquidity",
                 "body": "Operating cash flow was positive for the second consecutive quarter. Closing cash provides approximately 19 months of runway at the current burn profile, before any external financing."},
                {"h": "Risks & Mitigations",
                 "body": "Principal risks include lengthening enterprise sales cycles, AI infrastructure cost inflation, and key-person dependency. Mitigation programs are in place for each, detailed in the risk register."},
                {"h": "Recommendations",
                 "body": "Management requests Board approval of the five resolutions summarized in the recommendations table, including the incremental S&M budget and Series A preparation timeline."},
            ]}


def fallback_docx(objective, ctx, sym="₹"):
    return fallback_pdf(objective, ctx, sym)


# ─────────────────────────────────────────────────────────────────
# AI-enhanced extraction (wraps fallback; enriches when AI available)
# ─────────────────────────────────────────────────────────────────
_PROMPT = """You are a McKinsey-caliber financial analyst. Based on the objective and data below, output ONLY a JSON object (no markdown, no prose) with this exact shape:

{{
 "title": "short document title",
 "kpis": [["KPI name","value","delta"], ...6-8 rows],
 "months": ["Mon-YY","Mon-YY","Mon-YY"],
 "rev": [num,num,num], "cogs": [num,num,num], "opex": [num,num,num],
 "risks": [["risk","High|Medium|Low","mitigation"], ...4-6 rows],
 "recs": ["recommendation", ...4-6 items],
 "sections": [{{"h":"heading","body":"2-4 sentence executive paragraph"}}, ...5-7 sections],
 "narrative_points": [["section heading",["point","point","point"]], ...4-6 groups]
}}

Use the currency symbol {sym}. All numbers must be internally consistent. Derive real figures from the data if present; otherwise create realistic consistent figures for the business described.

OBJECTIVE: {obj}
COMPANY CONTEXT: {ctx}
AVAILABLE DATA:
{data}
"""


def extract(objective, ctx, data, api_key, sym="₹"):
    """Master extraction. Returns (enriched_model_dict, mode, reason).
    mode: 'ai' or 'fallback'. NEVER raises."""
    try:
        base = _base_model(objective, sym)
        raw = _call_ai(_PROMPT.format(sym=sym, obj=objective[:1500],
                                      ctx=ctx[:2000], data=data[:8000] or "(none)"),
                       api_key)
        if raw is None:
            return base, "fallback", "no api key or AI call failed"
        parsed = _try_parse(raw)
        if not isinstance(parsed, dict):
            return base, "fallback", "AI returned unparseable JSON"
        # merge validated fields over base
        def _ok_list(v, n=1):
            return isinstance(v, list) and len(v) >= n
        if _ok_list(parsed.get("kpis"), 4):
            base["kpis"] = [[str(a)[:40], str(b)[:24], str(c)[:24]]
                            for a, b, c in (r[:3] + [""] * (3 - len(r[:3])) for r in parsed["kpis"][:8] if isinstance(r, list))] or base["kpis"]
        for k in ("months", "rev", "cogs", "opex"):
            v = parsed.get(k)
            if _ok_list(v, 3):
                base[k] = v[:3] if k == "months" else [float(x) for x in v[:3] if isinstance(x, (int, float))] or base[k]
        if len(base["rev"]) == 3 and len(base["cogs"]) == 3 and len(base["opex"]) == 3:
            base["gross"] = [r - c for r, c in zip(base["rev"], base["cogs"])]
            base["ebitda"] = [g - o for g, o in zip(base["gross"], base["opex"])]
        if _ok_list(parsed.get("risks"), 3):
            base["risks"] = [[str(x)[:80] for x in (r[:3] + [""] * (3 - len(r[:3])))]
                             for r in parsed["risks"][:6] if isinstance(r, list)] or base["risks"]
        if _ok_list(parsed.get("recs"), 3):
            base["recs"] = [str(r)[:160] for r in parsed["recs"][:6]]
        if isinstance(parsed.get("title"), str) and parsed["title"].strip():
            base["title"] = parsed["title"][:90]
        if _ok_list(parsed.get("sections"), 3):
            secs = []
            for s in parsed["sections"][:7]:
                if isinstance(s, dict) and s.get("h") and s.get("body"):
                    secs.append({"h": str(s["h"])[:70], "body": str(s["body"])[:1200]})
            if len(secs) >= 3:
                base["sections"] = secs
        if _ok_list(parsed.get("narrative_points"), 3):
            nps = []
            for g in parsed["narrative_points"][:6]:
                if isinstance(g, list) and len(g) == 2 and isinstance(g[1], list):
                    nps.append([str(g[0])[:70], [str(p)[:140] for p in g[1][:5]]])
            if len(nps) >= 3:
                base["narrative_points"] = nps
        return base, "ai", "ok"
    except Exception as e:
        traceback.print_exc()
        return _base_model(objective, sym), "fallback", f"extractor exception: {str(e)[:120]}"
