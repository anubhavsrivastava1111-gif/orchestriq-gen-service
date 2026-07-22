"""
OrchestrIQ Document Intelligence Engine v4.3 - Layout & Theme Engine
Post-processes any blueprint (Excel workbook, PPTX presentation, or
PDF/DOCX document) to attach theme metadata and optimise chart type
selection based on data shape. Never raises - if styling fails, the
original blueprint is returned unchanged, preserving the zero-500
guarantee of the pipeline.
"""
from __future__ import annotations
import copy
import os
import json
from typing import Any, Dict


DEFAULT_THEME = {
    "font_name": "Calibri",
    "font_family_headings": "Calibri",
    "font_size_body": 11,
    "font_size_h1": 24,
    "font_size_h2": 16,
    "primary_color": "#1E3A5F",
    "accent_color": "#14B8A6",
    "neutral_color": "#64748B",
    "light_color": "#F1F5F9",
    "success_color": "#16A34A",
    "warning_color": "#D97706",
    "danger_color": "#DC2626",
    "chart_palette": ["#1E3A5F", "#14B8A6", "#D97706", "#16A34A", "#DC2626", "#9333EA", "#0EA5E9", "#F97316"],
    "logo_placeholder": "OrchestrIQ",
}


def _get_theme() -> Dict[str, Any]:
    """Read theme from ORCHESTRIQ_THEME env (JSON), falling back to defaults."""
    theme = DEFAULT_THEME.copy()
    try:
        override = os.environ.get("ORCHESTRIQ_THEME", "").strip()
        if override:
            parsed = json.loads(override)
            if isinstance(parsed, dict):
                theme.update({k: v for k, v in parsed.items() if v is not None})
    except Exception:
        pass
    return theme


# ─── Smart chart selection ─────────────────────────────────────────
def _looks_like_time(categories: list) -> bool:
    if not categories:
        return False
    try:
        sample = str(categories[0]).lower()
        for token in ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug",
                      "sep", "oct", "nov", "dec", "q1", "q2", "q3", "q4",
                      "week", "month", "day", "-25", "-24", "-26", "-23",
                      "2023", "2024", "2025", "2026", "2027"):
            if token in sample:
                return True
    except Exception:
        pass
    return False


def _series_values(series: Any) -> list:
    """Extract numeric values from a series entry [[name,[values]], ...]."""
    out = []
    try:
        if isinstance(series, list) and len(series) >= 2 and isinstance(series[1], list):
            out = [v for v in series[1] if isinstance(v, (int, float))]
    except Exception:
        pass
    return out


def _choose_chart_type(chart_spec: Dict[str, Any]) -> str:
    """Pick the most appropriate chart type. Respects a strong AI hint."""
    try:
        ctype = str(chart_spec.get("ctype", "")).lower()
        cats = chart_spec.get("cats") or chart_spec.get("categories") or []
        series = chart_spec.get("series") or []
        title = str(chart_spec.get("title", "")).lower()

        # If AI already gave a valid type, respect it.
        if ctype in {"bar", "line", "pie", "stacked", "area", "scatter", "hbar"}:
            return ctype

        # Share/percentage language + single series -> pie
        if len(series) == 1 and any(w in title for w in ("share", "portion", "percentage", "mix", "split", "breakdown", "composition")):
            return "pie"

        # Time series -> line if many points, bar if few
        if _looks_like_time(cats):
            if len(series) == 1 and len(cats) >= 6:
                return "line"
            if len(series) >= 2:
                return "line"
            return "bar"

        # Single series over many categories -> horizontal bar
        if len(series) == 1 and len(cats) > 10:
            return "hbar"

        # Two-plus series categorical -> grouped bar
        return "bar"
    except Exception:
        return chart_spec.get("ctype", "bar") or "bar"


# ─── Public entry point ────────────────────────────────────────────
def style_blueprint(blueprint: Any, currency_symbol: str = "\u20b9") -> Any:
    """Return a styled copy of the blueprint. Never raises."""
    if not isinstance(blueprint, dict):
        return blueprint
    try:
        bp = copy.deepcopy(blueprint)
        bp["_theme"] = _get_theme()
        bp["_currency_symbol"] = currency_symbol
        bp["_layout_hints"] = {
            "min_slide_margin_in": 0.5,
            "max_data_density": 0.7,
            "prefer_whitespace": True,
        }
        _optimise_charts(bp)
        return bp
    except Exception:
        return blueprint


def _optimise_charts(obj: Any) -> None:
    """Walk any blueprint structure and rewrite chart ctype in place."""
    try:
        if isinstance(obj, dict):
            # PPTX slide-shape chart
            if obj.get("type") == "chart" and isinstance(obj.get("chart"), dict):
                obj["chart"]["ctype"] = _choose_chart_type(obj["chart"])
            # Direct chart spec on document section
            if obj.get("chart") and isinstance(obj["chart"], dict) and "ctype" in obj["chart"]:
                obj["chart"]["ctype"] = _choose_chart_type(obj["chart"])
            # Excel dashboard chart list
            if "charts" in obj and isinstance(obj["charts"], list):
                for ch in obj["charts"]:
                    if isinstance(ch, dict) and "type" in ch:
                        # Only rewrite when the AI didn't give a valid hint
                        cur = str(ch.get("type", "")).lower()
                        if cur not in {"bar", "line", "pie", "hbar", "stacked"}:
                            chosen = _choose_chart_type({"ctype": "",
                                                         "cats": ch.get("cats") or obj.get("cats") or [],
                                                         "series": ch.get("series") or obj.get("series") or [],
                                                         "title": ch.get("title", "")})
                            ch["type"] = chosen
            for v in obj.values():
                _optimise_charts(v)
        elif isinstance(obj, list):
            for v in obj:
                _optimise_charts(v)
    except Exception:
        pass
