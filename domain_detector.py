"""
OrchestrIQ Document Intelligence Engine v4.3 — Domain Detector
Detects business domain from free-text and returns appropriate
deterministic fallback models and blueprints. Flat-file, no packages.
Never raises. Financial, workforce, marketing, product, legal, generic.
"""
from __future__ import annotations
import re

# Order matters - first match wins.
_DOMAIN_PATTERNS = [
    (re.compile(r'\b(board|quarterly|p&l|profit|loss|cash\s+flow|financial\s+statement|budget|forecast|investor|revenue\s+model|earnings|eps|ebitda|arr|mrr)\b', re.I),
     "financial"),
    (re.compile(r'\b(workforce|fte|employee|staffing|headcount|shrinkage|occupancy|roster|capacity\s+plan|wfm|attendance|shift|payroll|hr\s+operations|utilization)\b', re.I),
     "workforce"),
    (re.compile(r'\b(marketing|campaign|lead\s+generation|conversion|funnel|brand|awareness|pricing|promotion|go.to.market|gtm|cac|roas|mql)\b', re.I),
     "marketing"),
    (re.compile(r'\b(product\s+spec|feature\s+roadmap|technical\s+documentation|api\s+spec|sprint|scrum|devops|release\s+plan|product\s+roadmap)\b', re.I),
     "product"),
    (re.compile(r'\b(legal|compliance|regulation|policy|terms\s+and\s+conditions|privacy|gdpr|hipaa|audit|risk\s+assessment|litigation)\b', re.I),
     "legal"),
]

_GENERIC = "generic"


def detect_domain(objective: str, company_context: str = "") -> str:
    """Return one of: financial, workforce, marketing, product, legal, generic. Never raises."""
    try:
        text = f"{objective or ''} {company_context or ''}"
        for pattern, label in _DOMAIN_PATTERNS:
            if pattern.search(text):
                return label
    except Exception:
        pass
    return _GENERIC


# ═════════════════════════════════════════════════════════════════
# Flat model fallbacks (v4-model dict shape)
# ═════════════════════════════════════════════════════════════════
def get_fallback_model(objective: str, currency_symbol: str = "\u20b9") -> dict:
    """Return a v4-shaped model dict appropriate to the detected domain.
    Never raises."""
    try:
        domain = detect_domain(objective, "")
        if domain == "financial":
            # Reuse the existing base model - preserves v4 backward compatibility.
            from ai_extractor import _base_model
            m = _base_model(objective, currency_symbol)
            m["_domain"] = "financial"
            return m
        if domain == "workforce":
            return _fallback_workforce_model(objective, currency_symbol)
        if domain == "marketing":
            return _fallback_marketing_model(objective, currency_symbol)
        if domain == "product":
            return _fallback_product_model(objective, currency_symbol)
        if domain == "legal":
            return _fallback_legal_model(objective, currency_symbol)
        return _fallback_generic_model(objective, currency_symbol)
    except Exception:
        return _fallback_generic_model(objective, currency_symbol)


def _fallback_workforce_model(objective: str, sym: str) -> dict:
    months = ["Apr-25", "May-25", "Jun-25"]
    return {
        "_domain": "workforce",
        "title": objective[:80] or "Workforce Performance Review",
        "months": months,
        "rev": [0, 0, 0], "cogs": [0, 0, 0], "opex": [0, 0, 0],
        "gross": [0, 0, 0], "ebitda": [0, 0, 0], "cash_open": 0,
        "kpis": [
            ["Total Headcount", "150", "+8 vs prior"],
            ["Avg Utilization", "76%", "+2 pts"],
            ["Attrition (annualized)", "18%", "-2 pts"],
            ["Avg Occupancy", "82%", "+1 pt"],
            ["Avg Shrinkage", "18%", "in-band"],
            ["Attendance Rate", "96%", "+0.5 pt"],
            ["Overtime Hours", "2,400", "-8%"],
            ["Training Hours/FTE", "8.4", "+0.6"],
        ],
        "risks": [
            ["Key-person dependency in Ops", "Medium", "Cross-training program; documented SOPs"],
            ["Shift-coverage gaps in night window", "Medium", "Flex-pool of 15 agents; incentive pay"],
            ["Attrition spike above 25%", "High", "Career pathing; retention bonuses; exit-interview analytics"],
            ["Skill gaps in emerging tools", "Low", "Quarterly upskilling; certification funding"],
        ],
        "recs": [
            "Approve incremental 12 FTE for peak-season coverage",
            "Roll out flexible shift-swap tool across 3 largest teams",
            "Introduce productivity dashboards at team-lead level",
            "Adopt monthly workforce planning review cadence",
        ],
        "sections": [
            {"h": "Workforce Overview", "body": "Total headcount, department distribution, and utilization trends across the period. Occupancy remained above 80% target; shrinkage held within the 18% band."},
            {"h": "Operational Excellence", "body": "Initiatives to improve scheduling accuracy, reduce overtime spend, and lift first-call resolution. Automation of roster generation cut planning time by 40%."},
            {"h": "Talent & Retention", "body": "Attrition trended down 2 points versus prior quarter. Career-pathing and skill-certification programs are the primary levers."},
            {"h": "Risks & Actions", "body": "Principal workforce risks span coverage, retention, and skills. Mitigations are in place with owners and timelines."},
        ],
        "narrative_points": [
            ["Capacity Performance", ["Utilization at 76% (target 75%)", "Occupancy at 82% (target 80%)", "Shrinkage held at 18% within band"]],
            ["People Metrics", ["Attrition down 2 pts QoQ", "Training hours per FTE up 8%", "Employee NPS at +32"]],
        ],
    }


def _fallback_marketing_model(objective: str, sym: str) -> dict:
    months = ["Apr-25", "May-25", "Jun-25"]
    return {
        "_domain": "marketing",
        "title": objective[:80] or "Marketing Performance Review",
        "months": months,
        "rev": [1200000, 1350000, 1520000],
        "cogs": [240000, 270000, 304000],
        "opex": [600000, 660000, 720000],
        "gross": [960000, 1080000, 1216000],
        "ebitda": [360000, 420000, 496000],
        "cash_open": 2000000,
        "kpis": [
            ["MQLs (month)", "1,240", "+15% MoM"],
            ["CAC", f"{sym}850", "-5% MoM"],
            ["LTV:CAC", "4.2x", "+0.3"],
            ["Conversion Rate", "3.4%", "+0.4 pt"],
            ["ROAS", "5.8x", "+0.6"],
            ["Brand Search Volume", "18.4k", "+22%"],
            ["NPS", "48", "+2"],
            ["Pipeline Coverage", "3.4x", "+0.2"],
        ],
        "risks": [
            ["Channel fatigue on paid social", "Medium", "Diversify creative; test 3 new channels"],
            ["Privacy-driven attribution loss", "Medium", "Server-side tracking; MMM model"],
            ["Rising CPC in top segments", "Low", "Shift mix to earned + partner"],
        ],
        "recs": [
            "Increase branded-search budget by 20% (payback 2 months)",
            "Launch A/B test program for top-5 landing pages",
            "Consolidate reporting into single-source dashboard",
            "Initiate partner marketing motion in Q3",
        ],
        "sections": [
            {"h": "Campaign Performance", "body": "ROAS and funnel metrics across paid, owned, and earned channels. Paid social continues to dominate top-of-funnel; branded search delivers highest efficiency."},
            {"h": "Brand Health", "body": "Awareness, consideration, and NPS all improved. Share of voice up 4 points versus top competitor."},
            {"h": "Pipeline & Attribution", "body": "Pipeline coverage sits above 3x on next-quarter target. Attribution model migration on track for Q3 completion."},
        ],
        "narrative_points": [
            ["Growth Levers", ["Branded search efficiency", "Partner motion ramp-up", "Landing-page CRO program"]],
            ["Risk Watch", ["Channel diversification", "Attribution accuracy", "CPC inflation"]],
        ],
    }


def _fallback_product_model(objective: str, sym: str) -> dict:
    months = ["Apr-25", "May-25", "Jun-25"]
    return {
        "_domain": "product",
        "title": objective[:80] or "Product Performance Review",
        "months": months,
        "rev": [800000, 880000, 970000],
        "cogs": [160000, 176000, 194000],
        "opex": [420000, 460000, 500000],
        "gross": [640000, 704000, 776000],
        "ebitda": [220000, 244000, 276000],
        "cash_open": 1500000,
        "kpis": [
            ["Active Users", "42k", "+8% MoM"],
            ["Feature Adoption", "62%", "+3 pts"],
            ["NPS", "48", "+2"],
            ["Churn (monthly)", "1.2%", "-0.1 pt"],
            ["Time-to-value (days)", "14", "-3"],
            ["Uptime", "99.95%", "in-target"],
            ["Deploy Frequency", "42/mo", "+8"],
            ["Escaped Defects", "3", "-2"],
        ],
        "risks": [
            ["Technical-debt accrual", "Medium", "Quarterly refactor sprints (15% capacity)"],
            ["Competitive feature parity", "Medium", "API ecosystem differentiation"],
            ["Roadmap slippage", "Low", "Weekly progress reviews; scope discipline"],
        ],
        "recs": [
            "Allocate 15% of engineering capacity to tech-debt reduction",
            "Publish public roadmap for Q3",
            "Ship API v2 by end of quarter",
            "Introduce release-notes automation",
        ],
        "sections": [
            {"h": "Product Usage", "body": "Adoption, retention, and satisfaction metrics per feature area. Core feature adoption held above 60%; new modules trending on plan."},
            {"h": "Release Health", "body": "Cycle time, escape defects, and rollback frequency all improved. Deployment cadence up to 42/month with reduced incident count."},
            {"h": "Roadmap Progress", "body": "12 of 14 committed features shipped on time. Two features moved to next quarter with clear justification."},
        ],
        "narrative_points": [
            ["Engineering Excellence", ["Deploy frequency up", "Escape defects down", "Uptime above target"]],
            ["User Experience", ["NPS trending up", "TTV shortened", "Churn declining"]],
        ],
    }


def _fallback_legal_model(objective: str, sym: str) -> dict:
    months = ["Apr-25", "May-25", "Jun-25"]
    return {
        "_domain": "legal",
        "title": objective[:80] or "Legal & Compliance Review",
        "months": months,
        "rev": [0, 0, 0], "cogs": [0, 0, 0], "opex": [0, 0, 0],
        "gross": [0, 0, 0], "ebitda": [0, 0, 0], "cash_open": 0,
        "kpis": [
            ["Open Matters", "24", "-2"],
            ["Avg Resolution Time", "18 days", "+1 day"],
            ["Compliance Score", "94%", "+1 pt"],
            ["Contracts Reviewed", "148", "+12"],
            ["Policy Updates", "6", "+2"],
            ["Training Completion", "97%", "+3 pts"],
            ["Data Requests Handled", "42", "+8"],
            ["Audit Findings (open)", "3", "-4"],
        ],
        "risks": [
            ["Regulatory-change exposure", "Medium", "Monthly monitoring committee; external counsel briefings"],
            ["Data-subject request volume", "Medium", "Automate intake workflow; self-serve portal"],
            ["Vendor-contract non-compliance", "Low", "Quarterly vendor reviews; standard clauses"],
        ],
        "recs": [
            "Update data-retention policy to reflect new regulation",
            "Run quarterly tabletop exercise for incident response",
            "Automate DSR intake workflow",
            "Refresh vendor-onboarding legal review checklist",
        ],
        "sections": [
            {"h": "Litigation Overview", "body": "Summary of active cases, exposures, and reserves. Two matters approaching disposition; reserves adequate."},
            {"h": "Compliance Program", "body": "Status of controls, training, and third-party assessments. Overall compliance score improved to 94%."},
            {"h": "Regulatory Watch", "body": "Key legislation and regulatory actions expected to impact the business over the next two quarters."},
        ],
        "narrative_points": [
            ["Governance Health", ["Compliance score up", "Training completion at 97%", "Audit findings declining"]],
            ["Risk Watch", ["Regulatory pipeline", "DSR volume", "Vendor obligations"]],
        ],
    }


def _fallback_generic_model(objective: str, sym: str) -> dict:
    months = ["Apr-25", "May-25", "Jun-25"]
    return {
        "_domain": "generic",
        "title": objective[:80] or "Business Analysis",
        "months": months,
        "rev": [500000, 550000, 600000],
        "cogs": [200000, 220000, 240000],
        "opex": [180000, 190000, 200000],
        "gross": [300000, 330000, 360000],
        "ebitda": [120000, 140000, 160000],
        "cash_open": 800000,
        "kpis": [
            ["Total Value", f"{sym}1.65M", "+9%"],
            ["Active Items", "245", "+18"],
            ["Completion Rate", "82%", "+3 pts"],
            ["Efficiency", "76%", "+2 pts"],
        ],
        "risks": [
            ["Scope drift", "Medium", "Weekly steering-committee reviews"],
            ["Resource constraint", "Low", "Cross-training and demand planning"],
        ],
        "recs": [
            "Standardize monthly KPI review with variance analysis",
            "Prioritize top-3 initiatives for the next 90 days",
            "Establish quarterly business review with clear owners",
        ],
        "sections": [
            {"h": "Overview", "body": "Summary of key metrics and trends over the period. Refine the request with more specific business context for a tailored analysis."},
            {"h": "Recommendations", "body": "Suggested next actions and prioritization framework for the leadership team."},
        ],
        "narrative_points": [
            ["Performance", ["Metrics trending positive", "Efficiency gains", "Steady execution"]],
        ],
    }


# ═════════════════════════════════════════════════════════════════
# Blueprint fallbacks - referenced by ai_extractor.extract_blueprint
# ═════════════════════════════════════════════════════════════════
def _fallback_blueprint_workforce(objective: str, sym: str = "\u20b9") -> dict:
    """Delegate to the existing v4.1 workforce blueprint in ai_extractor to
    preserve tested behaviour (150-employee sheet, department summary, etc.)."""
    try:
        from ai_extractor import _fallback_blueprint_workforce as _wf
        return _wf(objective, sym)
    except Exception:
        return _fallback_blueprint_generic(objective, sym)


def _fallback_blueprint_generic(objective: str, sym: str = "\u20b9") -> dict:
    """Delegate to existing v4.1 generic blueprint; local fallback if unavailable."""
    try:
        from ai_extractor import _fallback_blueprint_generic as _gen
        return _gen(objective, sym)
    except Exception:
        pass
    return {
        "title": (objective[:80] or "Business Analysis Workbook"),
        "sheets": [
            {"name": "Assumptions", "type": "kv", "rows": [
                ["Scope", objective[:70] or "Business analysis", "From request"],
                ["Currency", sym, "Reporting currency"],
                ["Period", "Current month", "Default"]]},
            {"name": "Data", "type": "table", "row_count": 50, "columns": [
                {"h": "Item ID", "gen": {"kind": "id", "prefix": "ITM-", "start": 1001}},
                {"h": "Category", "gen": {"kind": "choice", "values": ["A", "B", "C", "D"]}},
                {"h": "Owner", "gen": {"kind": "name"}},
                {"h": "Quantity", "gen": {"kind": "number", "min": 10, "max": 500}, "format": "number"},
                {"h": "Unit Value", "gen": {"kind": "number", "min": 100, "max": 5000}, "format": "currency"},
                {"h": "Total Value", "formula": "{Quantity}*{Unit Value}", "format": "currency"},
                {"h": "Score %", "gen": {"kind": "number", "min": 0.4, "max": 0.98, "decimals": 3}, "format": "percent"}]},
            {"name": "Summary", "type": "summary", "source": "Data",
             "group_by": "Category", "aggregates": [
                 {"h": "Count", "kind": "count"},
                 {"h": "Total Value", "kind": "sum", "col": "Total Value", "format": "currency"},
                 {"h": "Avg Score", "kind": "avg", "col": "Score %", "format": "percent"}]},
            {"name": "Dashboard", "type": "dashboard", "kpis": [
                {"label": "Total Items", "ref": {"sheet": "Data", "agg": "count", "col": "Item ID"}},
                {"label": "Total Value", "ref": {"sheet": "Data", "agg": "sum", "col": "Total Value"}, "format": "currency"},
                {"label": "Avg Score", "ref": {"sheet": "Data", "agg": "avg", "col": "Score %"}, "format": "percent"}],
             "charts": [
                 {"title": "Value by Category", "type": "bar", "source": "Summary", "val_col": "Total Value"}]}]}
