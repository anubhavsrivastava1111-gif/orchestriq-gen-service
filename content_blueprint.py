"""
content_blueprint.py — build a document FROM THE USER'S OWN CONTENT.

WHY THIS EXISTS
---------------
When the AI structuring step failed, the pipeline called
domain_detector.get_fallback_model(objective, currency). Look at that signature:
it receives the objective and the currency symbol. It never receives the user's
content at all.

So a failure did not produce an error. It produced a polished, professional,
completely generic document — "Scope drift", "Revenue by Month", "Total Value
1.65M" — none of which the user ever wrote. That is worse than an error,
because it looks finished. A person could send it to a board.

This module removes the need for that guess. The content coming out of
Executive Chat and the Boardroom is already well-structured Markdown: headings,
tables, bullets, bold figures. That structure can be turned into a document
directly, deterministically, with no model involved and nothing invented.

A document built this way is not a downgrade on the AI path — for material that
is already organised it is frequently better, because nothing is summarised
away and every number survives exactly as written.
"""

import re

_MAX_SECTIONS = 24
_MAX_SLIDES = 26
_MAX_ROWS = 30
_MAX_COLS = 8


def _clean(s):
    """Strip Markdown emphasis that the renderers cannot draw."""
    t = str(s or "")
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", t)
    t = re.sub(r"__(.+?)__", r"\1", t)
    t = re.sub(r"`(.+?)`", r"\1", t)
    t = re.sub(r"\[(.+?)\]\((.*?)\)", r"\1", t)
    t = t.replace("*", "").replace("#", "")
    return re.sub(r"[ \t]+", " ", t).strip()


def _is_table_row(line):
    return line.strip().startswith("|") and line.count("|") >= 2


def _is_divider(line):
    s = line.strip()
    return bool(s) and set(s.replace("|", "").replace(" ", "")) <= {"-", ":"}


def _parse_table(lines, i):
    """Read a Markdown table starting at i. Returns (rows, next_index)."""
    rows = []
    while i < len(lines) and _is_table_row(lines[i]):
        if _is_divider(lines[i]):
            i += 1
            continue
        cells = [_clean(c) for c in lines[i].strip().strip("|").split("|")]
        if any(c for c in cells):
            rows.append(cells[:_MAX_COLS])
        i += 1
    # Pad ragged rows so the renderer never receives a jagged table.
    if rows:
        w = max(len(r) for r in rows)
        rows = [r + [""] * (w - len(r)) for r in rows[:_MAX_ROWS]]
    return rows, i


_NUM = re.compile(r"[₹$€£]?\s*[\d][\d,]*\.?\d*\s*%?")


def _numbers_in(text):
    """Every figure in a block, so a chart can be offered where one is useful."""
    return [m.group(0).strip() for m in _NUM.finditer(text or "")]


def parse_markdown(text):
    """Markdown -> [{heading, level, body, bullets, table}] in document order."""
    if not text:
        return []
    lines = str(text).replace("\r\n", "\n").split("\n")
    blocks, cur = [], None
    i = 0

    def flush():
        if cur and (cur["body"].strip() or cur["bullets"] or cur["table"]):
            cur["body"] = cur["body"].strip()
            blocks.append(cur)

    while i < len(lines):
        line = lines[i]
        m = re.match(r"^(#{1,4})\s+(.+)$", line.strip())
        if m:
            flush()
            cur = {"heading": _clean(m.group(2)), "level": len(m.group(1)),
                   "body": "", "bullets": [], "table": None}
            i += 1
            continue

        if cur is None:
            cur = {"heading": "", "level": 2, "body": "", "bullets": [], "table": None}

        if _is_table_row(line):
            rows, i = _parse_table(lines, i)
            if rows:
                # A section can carry one table. A second starts a new section so
                # neither is lost.
                if cur["table"] is None:
                    cur["table"] = rows
                else:
                    flush()
                    cur = {"heading": cur["heading"] + " (continued)", "level": cur["level"],
                           "body": "", "bullets": [], "table": rows}
            continue

        b = re.match(r"^\s*(?:[-*•]|\d+[.)])\s+(.+)$", line)
        if b:
            t = _clean(b.group(1))
            t = re.sub(r"^\[[ xX]\]\s*", "", t)   # checklist boxes
            if t:
                cur["bullets"].append(t)
            i += 1
            continue

        if line.strip():
            cur["body"] += (" " if cur["body"] else "") + _clean(line)
        i += 1

    flush()
    return [b for b in blocks if b["heading"] or b["body"] or b["bullets"] or b["table"]]


def _chart_from_table(rows, title):
    """A chart only when a table genuinely holds comparable numbers.
    Inventing one from labels would be worse than having none."""
    if not rows or len(rows) < 3:
        return None
    header, body = rows[0], rows[1:]
    num_cols = []
    for c in range(1, len(header)):
        vals, ok = [], True
        for r in body:
            if c >= len(r):
                ok = False
                break
            cell = (r[c] or "").strip()
            raw = re.sub(r"[^\d.\-]", "", cell)
            if raw in ("", "-", "."):
                ok = False
                break
            # The cell must be MOSTLY a number. Without this, "Day 7" and
            # "Q3 2026" become chart data and you get a bar chart of deadlines -
            # which is exactly the kind of confident nonsense a reader cannot
            # tell apart from real analysis.
            if len(raw) < len(cell) * 0.55:
                ok = False
                break
            try:
                vals.append(float(raw))
            except ValueError:
                ok = False
                break
        if ok and len(vals) == len(body) and any(v for v in vals):
            num_cols.append((header[c] or ("Series %d" % c), vals))
    # At least two different values, or the "chart" is a flat line saying nothing.
    num_cols = [(n, v) for n, v in num_cols if len(set(v)) > 1]
    if not num_cols:
        return None
    cats = [(r[0] or "")[:24] for r in body][:12]
    series = [[n[:40], v[:12]] for n, v in num_cols[:3]]
    return {"ctype": "bar", "title": (title or "Comparison")[:78],
            "cats": cats, "series": series}


def build_doc_blueprint(text, title, subtitle=""):
    """PDF / Word blueprint built entirely from the user's own content."""
    blocks = parse_markdown(text)
    if not blocks:
        return None
    sections = []
    for b in blocks[:_MAX_SECTIONS]:
        sec = {"h": (b["heading"] or "Overview")[:110]}
        if b["body"]:
            sec["body"] = b["body"][:4000]
        if b["bullets"]:
            sec["bullets"] = b["bullets"][:10]
        if b["table"]:
            sec["table"] = {"rows": b["table"]}
            ch = _chart_from_table(b["table"], b["heading"])
            if ch:
                sec["chart"] = ch
        # A heading with nothing under it is noise in a finished document.
        if sec.get("body") or sec.get("bullets") or sec.get("table"):
            sections.append(sec)
    if not sections:
        return None
    return {"title": (title or "Report")[:110],
            "subtitle": (subtitle or "")[:140],
            "sections": sections}


def build_pptx_blueprint(text, title, subtitle=""):
    """Deck blueprint built entirely from the user's own content.
    Long prose is split across slides rather than crammed onto one."""
    blocks = parse_markdown(text)
    if not blocks:
        return None
    slides = []
    for b in blocks:
        if len(slides) >= _MAX_SLIDES:
            break
        head = (b["heading"] or "Overview")[:110]

        if b["table"]:
            ch = _chart_from_table(b["table"], b["heading"])
            if ch:
                slides.append({"type": "chart", "h": head, "chart": ch,
                               "notes": (b["body"] or "")[:700]})
            slides.append({"type": "table", "h": head if not ch else head + " — detail",
                           "table": {"rows": b["table"]},
                           "notes": (b["body"] or "")[:700]})
            continue

        pts = list(b["bullets"])
        if not pts and b["body"]:
            # Sentences make better slide points than a wall of prose.
            pts = [s.strip() for s in re.split(r"(?<=[.!?])\s+", b["body"]) if len(s.strip()) > 15]
        if not pts:
            continue
        # Six points per slide keeps a slide readable; the rest continue over.
        for n in range(0, min(len(pts), 18), 6):
            chunk = pts[n:n + 6]
            slides.append({
                "type": "bullets",
                "h": head if n == 0 else head + " (cont.)",
                "kicker": "",
                "points": [p[:200] for p in chunk],
                "notes": (b["body"] or "")[:700] if n == 0 else "",
            })
    if not slides:
        return None
    return {"title": (title or "Presentation")[:110],
            "subtitle": (subtitle or "")[:140],
            "slides": slides}


def build(fmt, text, title, subtitle=""):
    """Entry point. Returns a blueprint, or None when there is nothing usable."""
    try:
        if fmt == "pptx":
            return build_pptx_blueprint(text, title, subtitle)
        return build_doc_blueprint(text, title, subtitle)
    except Exception:
        return None
