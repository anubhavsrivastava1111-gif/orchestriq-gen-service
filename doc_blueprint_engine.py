"""
OrchestrIQ Document Intelligence Engine v4.2 — Document Blueprint Engine
Generic renderers for AI-designed PPTX / PDF / DOCX blueprints.
The AI decides the structure per request; these functions render ANY spec.

PPTX blueprint:
{"title","subtitle","slides":[
  {"type":"bullets","h":"...","kicker":"...","points":["..."],"notes":"..."},
  {"type":"chart","h":"...","chart":{"ctype":"bar|line|pie","cats":[...],"series":[["name",[nums]],...]},"notes":"..."},
  {"type":"table","h":"...","table":{"rows":[[...],...]},"notes":"..."},
  {"type":"two_col","h":"...","left":["..."],"right":["..."],"notes":"..."},
  {"type":"kpi","h":"...","kpis":[["label","value","delta"],...],"notes":"..."}]}

DOC blueprint (PDF + DOCX share it):
{"title","subtitle","sections":[
  {"h":"...","body":"paragraph","bullets":["..."],
   "table":{"rows":[[...],...]},
   "chart":{"ctype":"bar","cats":[...],"series":[["name",[nums]]]}}]}
"""
import io

# ═════════════ PPTX ═════════════
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.chart import XL_CHART_TYPE
from pptx_engine import (_blank, _bar, _txt, _bullets, _notes, _header,
                         _chart, _table, NAVY, TEAL, WHITE, GREY, SW, SH)

_CT = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "line": XL_CHART_TYPE.LINE_MARKERS,
       "pie": XL_CHART_TYPE.PIE, "stacked": XL_CHART_TYPE.COLUMN_STACKED,
       "hbar": XL_CHART_TYPE.BAR_CLUSTERED}


def _safe_chart_data(ch):
    cats = [str(c)[:24] for c in (ch.get("cats") or ["A", "B", "C"])[:12]]
    series = []
    for s in (ch.get("series") or [])[:4]:
        if isinstance(s, list) and len(s) == 2 and isinstance(s[1], list):
            vals = [float(v) if isinstance(v, (int, float)) else 0 for v in s[1][:12]]
            vals += [0] * (len(cats) - len(vals))
            series.append((str(s[0])[:30], vals[:len(cats)]))
    if not series:
        series = [("Series 1", [1.0] * len(cats))]
    return cats, series


def render_pptx_blueprint(bp: dict) -> bytes:
    prs = Presentation()
    prs.slide_width = SW; prs.slide_height = SH
    title = str(bp.get("title", "Presentation"))[:80]
    subtitle = str(bp.get("subtitle", ""))[:80]

    # Title slide
    s = _blank(prs)
    _bar(s, 0, 0, SW, SH, NAVY)
    _bar(s, 0, Inches(4.6), SW, Inches(0.08), TEAL)
    _txt(s, Inches(0.9), Inches(2.6), Inches(11.5), Inches(1.4), title, 40, True, WHITE)
    if subtitle:
        _txt(s, Inches(0.9), Inches(4.9), Inches(11), Inches(0.6), subtitle, 20, False, TEAL)
    _notes(s, bp.get("opening_notes", "Open with the one-line framing of this deck."))

    # Agenda from slide headings
    heads = [str(sl.get("h", ""))[:60] for sl in (bp.get("slides") or []) if sl.get("h")]
    if len(heads) >= 3:
        s = _blank(prs); _header(s, "Agenda", "Overview")
        half = (len(heads[:12]) + 1) // 2
        _bullets(s, Inches(0.8), Inches(1.9), Inches(5.8), Inches(5), heads[:half], 16)
        _bullets(s, Inches(7.0), Inches(1.9), Inches(5.8), Inches(5), heads[half:12], 16)
        _notes(s, "Walk the agenda in 20 seconds.")

    for sl in (bp.get("slides") or [])[:24]:
        t = sl.get("type", "bullets")
        h = str(sl.get("h", "Section"))[:70]
        kick = str(sl.get("kicker", ""))[:40]
        s = _blank(prs); _header(s, h, kick)
        if t == "chart" and sl.get("chart"):
            cats, series = _safe_chart_data(sl["chart"])
            ctype = _CT.get(sl["chart"].get("ctype", "bar"), _CT["bar"])
            _chart(s, ctype, cats, series, Inches(0.9), Inches(1.8),
                   Inches(11.5), Inches(5.1), str(sl["chart"].get("title", h))[:60])
        elif t == "table" and (sl.get("table") or {}).get("rows"):
            rows = [[str(c)[:60] for c in r[:6]] for r in sl["table"]["rows"][:10]]
            w = min(11.9, 2.2 * len(rows[0]) + 2)
            _table(s, rows, Inches(0.7), Inches(1.9), Inches(w), Inches(4.6))
        elif t == "kpi" and sl.get("kpis"):
            rows = [["Metric", "Value", "Δ"]] + [[str(x)[:36] for x in (k[:3] + [""] * (3 - len(k[:3])))]
                                                for k in sl["kpis"][:8]]
            _table(s, rows, Inches(1.2), Inches(1.9), Inches(10.9), Inches(4.6))
        elif t == "two_col":
            _bullets(s, Inches(0.8), Inches(1.9), Inches(5.8), Inches(4.8),
                     [str(p)[:120] for p in (sl.get("left") or [])[:6]], 15)
            _bullets(s, Inches(7.0), Inches(1.9), Inches(5.8), Inches(4.8),
                     [str(p)[:120] for p in (sl.get("right") or [])[:6]], 15)
        else:
            _bullets(s, Inches(0.8), Inches(1.9), Inches(11.8), Inches(4.6),
                     [str(p)[:160] for p in (sl.get("points") or ["Content"])[:7]], 17)
        _notes(s, str(sl.get("notes", f"Speak to: {h}"))[:400])

    # Closing
    s = _blank(prs)
    _bar(s, 0, 0, SW, SH, NAVY)
    _txt(s, Inches(0.9), Inches(2.8), Inches(11.5), Inches(1.2), "Thank You", 44, True, WHITE)
    _txt(s, Inches(0.9), Inches(4.2), Inches(11.5), Inches(0.8),
         "Questions & Discussion", 20, False, TEAL)
    _notes(s, "Open the floor for questions.")

    # Floor: >=10 slides
    while len(prs.slides) < 10:
        s = _blank(prs); _header(s, "Supplementary", "Appendix")
        _bullets(s, Inches(0.8), Inches(1.9), Inches(11.8), Inches(4), 
                 ["Additional detail available on request"], 16)
        _notes(s, "Auto-appendix.")
    buf = io.BytesIO(); prs.save(buf)
    return buf.getvalue()


def validate_pptx_blueprint(bp) -> bool:
    return isinstance(bp, dict) and isinstance(bp.get("slides"), list) and len(bp["slides"]) >= 4


# ═════════════ PDF ═════════════
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph,
                                Spacer, Table, PageBreak, NextPageTemplate)
from pdf_engine import (_cover, _hf, _tstyle, _bar_drawing, _wrapped_table,
                        S_H1, S_BODY, S_TOC, S_SMALL, W, H)


def render_pdf_blueprint(bp: dict) -> bytes:
    title = str(bp.get("title", "Report"))[:90]
    subtitle = str(bp.get("subtitle", "Business Report"))[:90]
    buf = io.BytesIO()
    doc = BaseDocTemplate(buf, pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm,
                          topMargin=2.4 * cm, bottomMargin=2 * cm)
    frame = Frame(2 * cm, 2 * cm, W - 4 * cm, H - 4.6 * cm, id="f")
    doc.addPageTemplates([
        PageTemplate(id="cover", frames=[frame], onPage=lambda c, d: _cover(c, d, title, subtitle)),
        PageTemplate(id="body", frames=[frame], onPage=lambda c, d: _hf(c, d, title))])
    el = [NextPageTemplate("body"), PageBreak()]
    secs = (bp.get("sections") or [])[:14]
    el.append(Paragraph("Table of Contents", S_H1))
    for i, s in enumerate(secs, 1):
        el.append(Paragraph(f"{i}. {str(s.get('h','Section'))[:70]}", S_TOC))
    el.append(PageBreak())
    for s in secs:
        el.append(Paragraph(str(s.get("h", "Section"))[:80], S_H1))
        if s.get("body"):
            el.append(Paragraph(str(s["body"])[:6000], S_BODY))
        for b in (s.get("bullets") or [])[:8]:
            el.append(Paragraph("•  " + str(b)[:600], S_BODY))
        tab = (s.get("table") or {}).get("rows")
        if tab:
            # Wrap every cell (Paragraph) so long text wraps inside its column
            # instead of overflowing the right margin. Even column widths that
            # sum to exactly the frame width keep the table inside the page.
            rows = [[str(c) for c in r[:5]] for r in tab[:12]]
            ncol = len(rows[0]) if rows else 1
            cw = (W - 4 * cm) / ncol
            t = _wrapped_table(rows, [cw] * ncol)
            el.append(Spacer(1, 8)); el.append(t); el.append(Spacer(1, 8))
        ch = s.get("chart")
        if ch and ch.get("series"):
            cats, series = _safe_chart_data(ch)
            el.append(Spacer(1, 8))
            el.append(_bar_drawing(cats[:8], [(series[0][0], series[0][1][:8])],
                                   str(ch.get("title", s.get("h", "Chart")))[:60]))
            el.append(Spacer(1, 8))
    el.append(Spacer(1, 12))
    el.append(Paragraph("Generated by the OrchestrIQ Document Intelligence Engine.", S_SMALL))
    doc.build(el)
    return buf.getvalue()


# ═════════════ DOCX ═════════════
from docx import Document
from docx.shared import Pt as DPt, RGBColor as DRGB
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_engine import _toc_field, _page_numbers, _styled_table, NAVY as DNAVY, TEAL as DTEAL, GREY as DGREY


def render_docx_blueprint(bp: dict) -> bytes:
    title = str(bp.get("title", "Document"))[:90]
    subtitle = str(bp.get("subtitle", "Business Document"))[:90]
    doc = Document()
    normal = doc.styles["Normal"]; normal.font.name = "Calibri"; normal.font.size = DPt(11)
    for lvl, sz in [("Heading 1", 16), ("Heading 2", 13)]:
        st = doc.styles[lvl]; st.font.name = "Calibri"; st.font.size = DPt(sz)
        st.font.color.rgb = DNAVY; st.font.bold = True
    for _ in range(6): doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title); r.font.size = DPt(28); r.font.bold = True; r.font.color.rgb = DNAVY
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(subtitle); r.font.size = DPt(14); r.font.color.rgb = DTEAL
    for _ in range(10): doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Confidential"); r.font.size = DPt(10); r.font.color.rgb = DGREY
    doc.add_page_break()
    doc.add_heading("Table of Contents", level=1); _toc_field(doc); doc.add_page_break()
    for s in (bp.get("sections") or [])[:16]:
        doc.add_heading(str(s.get("h", "Section"))[:80], level=1)
        if s.get("body"):
            doc.add_paragraph(str(s["body"])[:3000])
        for b in (s.get("bullets") or [])[:10]:
            doc.add_paragraph(str(b)[:300], style="List Bullet")
        tab = (s.get("table") or {}).get("rows")
        if tab:
            _styled_table(doc, [[str(c)[:50] for c in r[:6]] for r in tab[:15]])
            doc.add_paragraph()
    _page_numbers(doc)
    buf = io.BytesIO(); doc.save(buf)
    return buf.getvalue()


def validate_doc_blueprint(bp) -> bool:
    return isinstance(bp, dict) and isinstance(bp.get("sections"), list) and len(bp["sections"]) >= 3
