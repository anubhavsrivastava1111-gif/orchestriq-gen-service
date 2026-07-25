"""
OrchestrIQ Document Intelligence Engine v4.3 - AI Schema Extractor
ZERO-RAISE GUARANTEE: this module never raises. Any failure at any layer
returns (fallback_schema, reason_string). Callers always get a usable schema.

v4.3 additions:
- domain_detector: picks appropriate fallback model per business domain
- layout_engine: post-processes blueprints with theme + smart chart selection
"""
import json
import re
import traceback

try:
    import domain_detector  # flat-file import
except Exception:
    domain_detector = None

try:
    import layout_engine  # flat-file import
except Exception:
    layout_engine = None

CLAUDE_MODEL = "claude-haiku-4-5-20251001"
OPENAI_MODEL = "gpt-4o-mini"
DEEPSEEK_MODEL = "deepseek-v4-pro"
DEFAULT_PROVIDER_ORDER = ["deepseek", "claude", "openai"]


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


def _call_claude(prompt: str, key: str, max_tokens: int):
    if not key or len(key) < 20:
        return None
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=key, timeout=90.0)
        msg = client.messages.create(
            model=CLAUDE_MODEL, max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}])
        return "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
    except Exception:
        traceback.print_exc()
        return None


def _call_openai_compatible(prompt: str, key: str, base_url: str, model: str, max_tokens: int):
    """Shared caller for OpenAI and DeepSeek — both speak the same REST
    format. Uses only Python's built-in networking, no extra libraries."""
    if not key or len(key) < 10:
        return None
    try:
        import json as _json
        from urllib import request as _urlreq
        body = _json.dumps({
            "model": model, "max_tokens": max_tokens,
            "messages": [{"role": "user", "content": prompt}],
        }).encode()
        req = _urlreq.Request(base_url, data=body, method="POST",
                              headers={"Content-Type": "application/json",
                                       "Authorization": "Bearer " + key.strip()})
        with _urlreq.urlopen(req, timeout=90) as resp:
            data = _json.loads(resp.read().decode())
        return data.get("choices", [{}])[0].get("message", {}).get("content", "")
    except Exception:
        traceback.print_exc()
        return None


def _call_openai(prompt: str, key: str, max_tokens: int):
    return _call_openai_compatible(prompt, key,
        "https://api.openai.com/v1/chat/completions", OPENAI_MODEL, max_tokens)


def _call_deepseek(prompt: str, key: str, max_tokens: int):
    return _call_openai_compatible(prompt, key,
        "https://api.deepseek.com/v1/chat/completions", DEEPSEEK_MODEL, max_tokens)


def _call_ai(prompt: str, keys: dict, order=None, max_tokens: int = 8000):
    """Tries each provider in `order`, using the matching key from `keys`.
    keys example: {"claude":"...", "openai":"...", "deepseek":"..."}.
    Falls back to server environment variables when a specific key is
    missing, so the service still works if the frontend sends nothing.
    Returns (raw_text, provider_used) — or (None, None) if every
    provider is missing a key or fails. Never raises."""
    import os
    order = order or DEFAULT_PROVIDER_ORDER
    keys = keys or {}
    callers = {
        "claude": lambda k: _call_claude(prompt, k or os.environ.get("ANTHROPIC_API_KEY", ""), max_tokens),
        "openai": lambda k: _call_openai(prompt, k or os.environ.get("OPENAI_API_KEY", ""), max_tokens),
        "deepseek": lambda k: _call_deepseek(prompt, k or os.environ.get("DEEPSEEK_API_KEY", ""), max_tokens),
    }
    for provider in order:
        caller = callers.get(provider)
        if not caller:
            continue
        text = caller(keys.get(provider, ""))
        if text:
            return text, provider
    return None, None


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

Use the currency symbol {sym}. CRITICAL: if DATA below contains figures, use EXACTLY those figures everywhere (no invented numbers); derive additional values only via arithmetic on them. Otherwise create realistic consistent figures. All numbers internally consistent.

OBJECTIVE: {obj}
COMPANY CONTEXT: {ctx}
AVAILABLE DATA:
{data}
"""


def extract(objective, ctx, data, keys, order=None, sym="₹"):
    """Master extraction. Returns (enriched_model_dict, mode, reason).
    mode: 'ai:<provider>' or 'fallback'. NEVER raises."""
    try:
        base = _base_model(objective, sym)
        raw, used_provider = _call_ai(_PROMPT.format(sym=sym, obj=objective[:1500],
                                      ctx=ctx[:2000], data=data[:8000] or "(none)"),
                       keys, order)
        if raw is None:
            if domain_detector is not None:
                try:
                    return domain_detector.get_fallback_model(objective, sym), "fallback", "no api key or AI call failed (domain fallback)"
                except Exception:
                    pass
            return base, "fallback", "no api key or AI call failed"
        parsed = _try_parse(raw)
        if not isinstance(parsed, dict):
            if domain_detector is not None:
                try:
                    return domain_detector.get_fallback_model(objective, sym), "fallback", "AI returned unparseable JSON (domain fallback)"
                except Exception:
                    pass
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
        return base, "ai:" + used_provider, "ok"
    except Exception as e:
        traceback.print_exc()
        if domain_detector is not None:
            try:
                return domain_detector.get_fallback_model(objective, sym), "fallback", f"extractor exception: {str(e)[:120]} (domain fallback)"
            except Exception:
                pass
        return _base_model(objective, sym), "fallback", f"extractor exception: {str(e)[:120]}"


# ═════════════════════════════════════════════════════════════════
# v4.1 BLUEPRINT EXTRACTION — AI designs the workbook structure
# ═════════════════════════════════════════════════════════════════
_BP_PROMPT = """You are a McKinsey-caliber analyst and Excel architect. Design a complete Excel workbook BLUEPRINT for the request below. Output ONLY JSON (no markdown, no prose).

Blueprint schema:
{{
 "title": "workbook title",
 "sheets": [
  {{"name":"Assumptions","type":"kv","rows":[["label",value,"rationale"],...8-15 rows]}},
  {{"name":"<Data sheet name>","type":"table","row_count":<N from request, default 50, max 500>,
   "columns":[
     {{"h":"<header>","gen":{{"kind":"id","prefix":"EMP-","start":1001}}}},
     {{"h":"<header>","gen":{{"kind":"name"}}}},
     {{"h":"<header>","gen":{{"kind":"choice","values":["...5-8 realistic values..."]}}}},
     {{"h":"<header>","gen":{{"kind":"choice_dependent","on":"<other col>","map":{{"<val>":["..."]}},"default":["..."]}}}},
     {{"h":"<header>","gen":{{"kind":"number","min":X,"max":Y,"decimals":D}},"format":"number|hours|currency|percent"}},
     {{"h":"<calculated header>","formula":"{{Col A}}-{{Col B}}","format":"number"}}
   ]}},
  {{"name":"<Group> Summary","type":"summary","source":"<data sheet name>","group_by":"<choice column>",
   "aggregates":[{{"h":"Headcount","kind":"count"}},{{"h":"Total X","kind":"sum","col":"<col>","format":"number"}},{{"h":"Avg Y","kind":"avg","col":"<col>","format":"percent"}}]}},
  {{"name":"Dashboard","type":"dashboard",
   "kpis":[{{"label":"...","ref":{{"sheet":"<data sheet>","agg":"count|sum|avg","col":"<col>"}},"format":"..."}} ...8-12 kpis],
   "charts":[{{"title":"...","type":"bar|line|pie","source":"<summary sheet>","cat_col":"<group>","val_col":"<agg header>"}} ...2-4 charts]}}
 ]
}}

RULES:
- Include EVERY column the user listed, in their order. Requested calculated fields use "formula" with {{Column Name}} tokens referencing THIS sheet's columns (native Excel math only: + - * / and parentheses; IFERROR allowed).
- percent-format columns use decimals 2-3 with min/max between 0 and 1.
- Honor requested row counts and category counts exactly (e.g. "150 employees across 5 departments" → row_count 150, 5 department values).
- Realistic values for the domain, currency amounts sized for {sym}. CRITICAL: if DATA PROVIDED contains actual figures or tables, reproduce EXACTLY those values in the workbook (as const gen kinds or kv rows), not invented ones.
- 2-3 summary sheets if multiple groupings are requested (department, team, etc).
- Dashboard KPIs must cover the user's requested dashboard metrics.

REQUEST:
{obj}

CONTEXT: {ctx}
DATA PROVIDED:
{data}
"""


def _fallback_blueprint_workforce(objective, sym):
    """Deterministic workforce/FTE blueprint — parses N employees and dept count."""
    m = re.search(r'(\d{2,4})\s*(?:employees|staff|agents|workers|people)', objective, re.I)
    n = min(int(m.group(1)), 500) if m else 150
    m2 = re.search(r'(\d{1,2})\s*departments?', objective, re.I)
    nd = min(int(m2.group(1)), 8) if m2 else 5
    depts = ["Customer Support", "Technical Support", "Sales Operations",
             "Back Office", "Quality Assurance", "Finance Ops", "HR Services",
             "IT Helpdesk"][:nd]
    team_map = {d: [f"{d.split()[0]} Team {i}" for i in (1, 2, 3)] for d in depts}
    cols = [
        {"h": "Employee ID", "gen": {"kind": "id", "prefix": "EMP-", "start": 10001}},
        {"h": "Employee Name", "gen": {"kind": "name"}},
        {"h": "Department", "gen": {"kind": "choice", "values": depts, "sequential": True}},
        {"h": "Team", "gen": {"kind": "choice_dependent", "on": "Department", "map": team_map, "default": ["Team 1"]}},
        {"h": "Manager", "gen": {"kind": "name"}},
        {"h": "Location", "gen": {"kind": "choice", "values": ["Lucknow", "Noida", "Bengaluru", "Hyderabad", "Manila"]}},
        {"h": "Shift", "gen": {"kind": "choice", "values": ["Morning", "Evening", "Night", "Split"]}},
        {"h": "Employment Type", "gen": {"kind": "choice", "values": ["Full-time", "Full-time", "Full-time", "Part-time"]}},
        {"h": "Planned Working Days", "gen": {"kind": "number", "min": 22, "max": 22}},
        {"h": "Actual Working Days", "gen": {"kind": "number", "min": 19, "max": 22}},
        {"h": "Scheduled Hours", "formula": "{Planned Working Days}*8", "format": "hours"},
        {"h": "Leave Hours", "gen": {"kind": "number", "min": 0, "max": 24}, "format": "hours"},
        {"h": "Training Hours", "gen": {"kind": "number", "min": 4, "max": 10}, "format": "hours"},
        {"h": "Meeting Hours", "gen": {"kind": "number", "min": 4, "max": 8}, "format": "hours"},
        {"h": "Break Hours", "formula": "{Actual Working Days}*1.5", "format": "hours"},
        {"h": "Overtime Hours", "gen": {"kind": "number", "min": 0, "max": 16}, "format": "hours"},
        {"h": "Available Hours", "formula": "{Scheduled Hours}-{Leave Hours}", "format": "hours"},
        {"h": "Shrinkage Hours", "formula": "{Training Hours}+{Meeting Hours}+{Break Hours}", "format": "hours"},
        {"h": "Capacity Hours", "formula": "{Available Hours}-{Shrinkage Hours}", "format": "hours"},
        {"h": "Occupancy %", "gen": {"kind": "number", "min": 0.74, "max": 0.9, "decimals": 3}, "format": "percent"},
        {"h": "Productive Hours", "formula": "{Capacity Hours}*{Occupancy %}", "format": "hours"},
        {"h": "Lost Hours", "formula": "{Capacity Hours}-{Productive Hours}", "format": "hours"},
        {"h": "Billable Hours", "formula": "{Productive Hours}*0.85", "format": "hours"},
        {"h": "Non-Billable Hours", "formula": "{Productive Hours}*0.15", "format": "hours"},
        {"h": "Attendance %", "formula": "IFERROR({Actual Working Days}/{Planned Working Days},0)", "format": "percent"},
        {"h": "Utilization %", "formula": "IFERROR({Productive Hours}/{Available Hours},0)", "format": "percent"},
        {"h": "Shrinkage %", "formula": "IFERROR({Shrinkage Hours}/{Available Hours},0)", "format": "percent"},
        {"h": "FTE", "formula": "IFERROR({Available Hours}/176,0)", "format": "decimal"},
        {"h": "Hourly Cost", "gen": {"kind": "number", "min": 180, "max": 650}, "format": "currency"},
        {"h": "Monthly Salary", "formula": "{Hourly Cost}*{Scheduled Hours}", "format": "currency"},
        {"h": "Cost per Productive Hour", "formula": "IFERROR({Monthly Salary}/{Productive Hours},0)", "format": "currency"},
        {"h": "Variance vs 85% Target", "formula": "{Utilization %}-0.85", "format": "percent"},
        {"h": "Cost Center", "gen": {"kind": "choice", "values": ["CC-1001", "CC-1002", "CC-1003", "CC-1004", "CC-1005"]}},
    ]
    aggs = [
        {"h": "Headcount", "kind": "count"},
        {"h": "Total FTE", "kind": "sum", "col": "FTE", "format": "decimal"},
        {"h": "Capacity Hours", "kind": "sum", "col": "Capacity Hours", "format": "hours"},
        {"h": "Productive Hours", "kind": "sum", "col": "Productive Hours", "format": "hours"},
        {"h": "Lost Hours", "kind": "sum", "col": "Lost Hours", "format": "hours"},
        {"h": "Avg Utilization", "kind": "avg", "col": "Utilization %", "format": "percent"},
        {"h": "Avg Occupancy", "kind": "avg", "col": "Occupancy %", "format": "percent"},
        {"h": "Avg Shrinkage", "kind": "avg", "col": "Shrinkage %", "format": "percent"},
        {"h": "Payroll Cost", "kind": "sum", "col": "Monthly Salary", "format": "currency"},
    ]
    return {
        "title": objective[:80] or "Workforce Capacity Planning & FTE Calculator",
        "sheets": [
            {"name": "Assumptions", "type": "kv", "rows": [
                ["Standard work hours/day", 8, "Company policy"],
                ["Working days/month", 22, "Excluding weekends & holidays"],
                ["Lunch break/day", "1 hour", "Unpaid"],
                ["Tea breaks/day", "30 min", "Paid"],
                ["Team meetings/month", "6 hours", "All-hands + team syncs"],
                ["Training/month", "8 hours", "Compliance + upskilling"],
                ["FTE basis", "176 hours", "22 days x 8 hours"],
                ["Target occupancy", "82%", "Industry benchmark"],
                ["Target shrinkage", "18%", "Industry benchmark"],
                ["Target utilization", "85%", "Board-approved target"],
                ["Currency", sym, "Reporting currency"]]},
            {"name": "Employee Data", "type": "table", "row_count": n, "columns": cols},
            {"name": "Department Summary", "type": "summary", "source": "Employee Data",
             "group_by": "Department", "aggregates": aggs},
            {"name": "Shift Summary", "type": "summary", "source": "Employee Data",
             "group_by": "Shift", "aggregates": aggs[:6]},
            {"name": "Dashboard", "type": "dashboard",
             "kpis": [
                 {"label": "Total Employees", "ref": {"sheet": "Employee Data", "agg": "count", "col": "Employee ID"}},
                 {"label": "Total FTE", "ref": {"sheet": "Employee Data", "agg": "sum", "col": "FTE"}, "format": "decimal"},
                 {"label": "Capacity Hours", "ref": {"sheet": "Employee Data", "agg": "sum", "col": "Capacity Hours"}, "format": "hours"},
                 {"label": "Productive Hours", "ref": {"sheet": "Employee Data", "agg": "sum", "col": "Productive Hours"}, "format": "hours"},
                 {"label": "Lost Hours", "ref": {"sheet": "Employee Data", "agg": "sum", "col": "Lost Hours"}, "format": "hours"},
                 {"label": "Avg Utilization", "ref": {"sheet": "Employee Data", "agg": "avg", "col": "Utilization %"}, "format": "percent"},
                 {"label": "Avg Occupancy", "ref": {"sheet": "Employee Data", "agg": "avg", "col": "Occupancy %"}, "format": "percent"},
                 {"label": "Avg Shrinkage", "ref": {"sheet": "Employee Data", "agg": "avg", "col": "Shrinkage %"}, "format": "percent"},
                 {"label": "Total Payroll", "ref": {"sheet": "Employee Data", "agg": "sum", "col": "Monthly Salary"}, "format": "currency"},
                 {"label": "Total Overtime Hrs", "ref": {"sheet": "Employee Data", "agg": "sum", "col": "Overtime Hours"}, "format": "hours"},
                 {"label": "Total Leave Hrs", "ref": {"sheet": "Employee Data", "agg": "sum", "col": "Leave Hours"}, "format": "hours"},
                 {"label": "Avg Attendance", "ref": {"sheet": "Employee Data", "agg": "avg", "col": "Attendance %"}, "format": "percent"}],
             "charts": [
                 {"title": "FTE by Department", "type": "bar", "source": "Department Summary", "val_col": "Total FTE"},
                 {"title": "Avg Utilization by Department", "type": "bar", "source": "Department Summary", "val_col": "Avg Utilization"},
                 {"title": "Payroll Cost by Department", "type": "pie", "source": "Department Summary", "val_col": "Payroll Cost"},
                 {"title": "Headcount by Shift", "type": "bar", "source": "Shift Summary", "val_col": "Headcount"}]}
        ]}


def _fallback_blueprint_generic(objective, sym):
    """Generic analysis workbook when domain is unknown."""
    return {
        "title": objective[:80] or "Business Analysis Workbook",
        "sheets": [
            {"name": "Assumptions", "type": "kv", "rows": [
                ["Scope", objective[:70] or "Business analysis", "From request"],
                ["Currency", sym, "Reporting currency"],
                ["Period", "Current month", "Default"]]},
            {"name": "Data", "type": "table", "row_count": 50, "columns": [
                {"h": "Item ID", "gen": {"kind": "id", "prefix": "ITM-", "start": 1001}},
                {"h": "Category", "gen": {"kind": "choice", "values": ["Category A", "Category B", "Category C", "Category D"]}},
                {"h": "Owner", "gen": {"kind": "name"}},
                {"h": "Quantity", "gen": {"kind": "number", "min": 10, "max": 500}, "format": "number"},
                {"h": "Unit Value", "gen": {"kind": "number", "min": 100, "max": 5000}, "format": "currency"},
                {"h": "Total Value", "formula": "{Quantity}*{Unit Value}", "format": "currency"},
                {"h": "Score %", "gen": {"kind": "number", "min": 0.4, "max": 0.98, "decimals": 3}, "format": "percent"}]},
            {"name": "Category Summary", "type": "summary", "source": "Data",
             "group_by": "Category", "aggregates": [
                 {"h": "Count", "kind": "count"},
                 {"h": "Total Value", "kind": "sum", "col": "Total Value", "format": "currency"},
                 {"h": "Avg Score", "kind": "avg", "col": "Score %", "format": "percent"}]},
            {"name": "Dashboard", "type": "dashboard", "kpis": [
                {"label": "Total Items", "ref": {"sheet": "Data", "agg": "count", "col": "Item ID"}},
                {"label": "Total Value", "ref": {"sheet": "Data", "agg": "sum", "col": "Total Value"}, "format": "currency"},
                {"label": "Avg Score", "ref": {"sheet": "Data", "agg": "avg", "col": "Score %"}, "format": "percent"}],
             "charts": [
                 {"title": "Value by Category", "type": "bar", "source": "Category Summary", "val_col": "Total Value"}]}
        ]}


_WF_KEYWORDS = re.compile(r'\b(workforce|fte|employee|staffing|headcount|shrinkage|occupancy|roster|capacity plan|wfm|attendance)\b', re.I)
_FIN_KEYWORDS = re.compile(r'\b(board|quarterly|p&l|profit|cash flow|financial statement|budget|forecast|investor|revenue model)\b', re.I)


def extract_blueprint(objective, ctx, data, keys, order=None, sym="\u20b9"):
    """Returns (blueprint, mode, reason). NEVER raises.
    AI designs the structure; deterministic fallbacks by domain keywords."""
    try:
        from blueprint_engine import validate_blueprint
        raw, used_provider = _call_ai(_BP_PROMPT.format(sym=sym, obj=objective[:5000],
                                         ctx=ctx[:2000], data=(data[:6000] or "(none)")),
                       keys, order, max_tokens=14000)
        if raw is not None:
            bp = _try_parse(raw)
            if validate_blueprint(bp):
                # cap row counts
                for s in bp.get("sheets", []):
                    if isinstance(s, dict) and s.get("type") == "table":
                        s["row_count"] = min(int(s.get("row_count", 50) or 50), 500)
                # v4.3: theme + smart chart selection
                if layout_engine is not None:
                    try:
                        bp = layout_engine.style_blueprint(bp, sym)
                    except Exception:
                        pass
                return bp, "ai:" + used_provider, "ok"
            reason = "AI blueprint invalid; domain fallback used"
        else:
            reason = "no api key or AI call failed; domain fallback used"
    except Exception as e:
        traceback.print_exc()
        reason = f"blueprint exception: {str(e)[:100]}"
    # v4.3: prefer domain_detector (broader domain coverage). Keep legacy
    # keyword routing as safety net when detector unavailable.
    if domain_detector is not None:
        try:
            domain = domain_detector.detect_domain(objective, ctx)
            if domain == "financial":
                return None, "fallback_v4_template", reason
            if domain == "workforce":
                return _fallback_blueprint_workforce(objective, sym), "fallback", reason
            return _fallback_blueprint_generic(objective, sym), "fallback", reason
        except Exception:
            pass
    if _FIN_KEYWORDS.search(objective):
        return None, "fallback_v4_template", reason
    if _WF_KEYWORDS.search(objective):
        return _fallback_blueprint_workforce(objective, sym), "fallback", reason
    return _fallback_blueprint_generic(objective, sym), "fallback", reason


# ═════════════════════════════════════════════════════════════════
# v4.2 DOCUMENT BLUEPRINTS — AI designs PPTX / PDF / DOCX structure
# ═════════════════════════════════════════════════════════════════
_PPTX_PROMPT = """You are a McKinsey-caliber presentation designer. Design a boardroom-quality PowerPoint BLUEPRINT for the request below. Output ONLY JSON.

Schema:
{{"title":"...","subtitle":"...",
 "slides":[
  {{"type":"bullets","h":"heading","kicker":"section label","points":["insight sentence",...3-5],"notes":"speaker note"}},
  {{"type":"kpi","h":"...","kpis":[["label","value","delta"],...5-8],"notes":"..."}},
  {{"type":"chart","h":"...","chart":{{"ctype":"bar|line|pie","title":"...","cats":["..."],"series":[["name",[numbers]]]}},"notes":"..."}},
  {{"type":"table","h":"...","table":{{"rows":[["hdr",...],["...",...]]}},"notes":"..."}},
  {{"type":"two_col","h":"...","left":["..."],"right":["..."],"notes":"..."}}
 ]}}

RULES:
- 10-16 slides tailored EXACTLY to the request topic and audience. No generic filler.
- At least 3 chart slides. CRITICAL: if DATA contains actual figures, use EXACTLY those in charts/KPIs; invent nothing that contradicts them (currency {sym}).
- At least 1 kpi and 1 table slide. Every slide has a specific, useful speaker note.
- Executive storytelling arc: situation → analysis → insight → recommendation → next steps.

REQUEST: {obj}
CONTEXT: {ctx}
DATA: {data}
"""

_DOC_PROMPT = """You are a McKinsey-caliber consultant. Design a publication-quality {kind} BLUEPRINT for the request below. Output ONLY JSON.

Schema:
{{"title":"...","subtitle":"...",
 "sections":[
  {{"h":"heading","body":"3-6 sentence executive paragraph","bullets":["optional point",...0-6],
    "table":{{"rows":[["hdr",...],["...",...]]}},
    "chart":{{"ctype":"bar","title":"...","cats":["..."],"series":[["name",[numbers]]]}} }}
 ]}}

RULES:
- 6-10 sections tailored EXACTLY to the request. Executive summary first, recommendations near the end.
- Include at least 2 tables and 1 chart. CRITICAL: if DATA contains actual figures, use EXACTLY those; derive additional values only via arithmetic ({sym}).
- Substantive analytical writing, no generic AI filler. table/chart keys optional per section.

REQUEST: {obj}
CONTEXT: {ctx}
DATA: {data}
"""


def _model_to_pptx_bp(model, title, subtitle):
    """Convert the v4 base/enriched model to a pptx blueprint (fallback floor)."""
    months, rev, gross, ebitda = model["months"], model["rev"], model["gross"], model["ebitda"]
    naps = model.get("narrative_points") or []
    slides = [
        {"type": "bullets", "h": "Executive Summary", "kicker": "Overview",
         "points": [s["h"] + ": " + s["body"].split(".")[0] + "." for s in (model.get("sections") or [])[:4]] or
                   ["Strong performance with expanding margins", "Growth funded by installed-base expansion"],
         "notes": "Frame the quarter in 60 seconds."},
        {"type": "kpi", "h": "KPI Scorecard", "kpis": model["kpis"][:8], "notes": "Highlight the two KPIs the audience tracks most."},
        {"type": "chart", "h": "Revenue Trajectory",
         "chart": {"ctype": "bar", "title": "Revenue & Gross Profit", "cats": months,
                   "series": [["Revenue", rev], ["Gross Profit", gross]]}, "notes": "Sequential growth story."},
        {"type": "chart", "h": "Profitability Trend",
         "chart": {"ctype": "line", "title": "EBITDA by Month", "cats": months,
                   "series": [["EBITDA", ebitda]]}, "notes": "Operating leverage in one chart."},
    ]
    for h, pts in naps[:3]:
        slides.append({"type": "bullets", "h": h, "kicker": "Business Review",
                       "points": pts[:5], "notes": f"Keep {h} to 90 seconds."})
    slides.append({"type": "table", "h": "Risk Register",
                   "table": {"rows": [["Risk", "Severity", "Mitigation"]] + [r[:3] for r in model["risks"][:5]]},
                   "notes": "Surface the top risk proactively."})
    slides.append({"type": "bullets", "h": "Recommendations", "kicker": "Decisions Requested",
                   "points": model["recs"][:6], "notes": "Pause for discussion."})
    return {"title": title, "subtitle": subtitle, "slides": slides}


def _model_to_doc_bp(model, title, subtitle):
    secs = list(model.get("sections") or [])
    out = [{"h": s["h"], "body": s["body"]} for s in secs]
    if out:
        out[0]["table"] = {"rows": [["KPI", "Value", "Δ"]] + [k[:3] for k in model["kpis"][:8]]}
    out.append({"h": "Risk Register", "body": "Principal risks and mitigations are summarized below.",
                "table": {"rows": [["Risk", "Severity", "Mitigation"]] + [r[:3] for r in model["risks"][:6]]}})
    out.append({"h": "Recommendations",
                "body": "Management requests approval of the following actions.",
                "bullets": model["recs"][:6],
                "chart": {"ctype": "bar", "title": "Revenue by Month",
                          "cats": model["months"], "series": [["Revenue", model["rev"]]]}})
    return {"title": title, "subtitle": subtitle, "sections": out}


def extract_doc_blueprint(fmt, objective, ctx, data, keys, order=None, sym="\u20b9"):
    """fmt in {'pptx','pdf','docx'}. Returns (blueprint, mode, reason). Never raises."""
    try:
        from doc_blueprint_engine import validate_pptx_blueprint, validate_doc_blueprint
        if fmt == "pptx":
            prompt = _PPTX_PROMPT.format(sym=sym, obj=objective[:5000], ctx=ctx[:2000],
                                         data=(data[:6000] or "(none)"))
            valid = validate_pptx_blueprint
        else:
            prompt = _DOC_PROMPT.format(kind=("PDF report" if fmt == "pdf" else "Word document"),
                                        sym=sym, obj=objective[:5000], ctx=ctx[:2000],
                                        data=(data[:6000] or "(none)"))
            valid = validate_doc_blueprint
        raw, used_provider = _call_ai(prompt, keys, order, max_tokens=12000)
        if raw is not None:
            bp = _try_parse(raw)
            if valid(bp):
                # v4.3: theme + smart chart selection
                if layout_engine is not None:
                    try:
                        bp = layout_engine.style_blueprint(bp, sym)
                    except Exception:
                        pass
                return bp, "ai:" + used_provider, "ok"
            reason = "AI doc blueprint invalid; model fallback used"
        else:
            reason = "no working API key for any configured provider; model fallback used"
    except Exception as e:
        traceback.print_exc()
        reason = f"doc blueprint exception: {str(e)[:100]}"
    # v4.3: derive blueprint from domain-appropriate model when possible.
    model = None
    if domain_detector is not None:
        try:
            model = domain_detector.get_fallback_model(objective, sym)
        except Exception:
            model = None
    if model is None:
        model, _m2, _r2 = extract(objective, ctx, data, keys, order, sym)
    title = model.get("title") or objective[:80] or "Business Document"
    subtitle = "Prepared by OrchestrIQ"
    bp = _model_to_pptx_bp(model, title, subtitle) if fmt == "pptx" else \
         _model_to_doc_bp(model, title, subtitle)
    if layout_engine is not None:
        try:
            bp = layout_engine.style_blueprint(bp, sym)
        except Exception:
            pass
    return bp, "fallback", reason
