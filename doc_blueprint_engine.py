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
                         _chart, _table, _metric_cards, _two_col, _hero, _round,
                         NAVY, TEAL, WHITE, GREY, SW, SH)

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

    # ── Title ──
    s = _blank(prs)
    _bar(s, 0, 0, SW, SH, NAVY)
    _bar(s, 0, Inches(4.7), SW, Inches(0.09), TEAL)
    _bar(s, Inches(0.9), Inches(2.35), Inches(1.4), Inches(0.14), TEAL)
    _txt(s, Inches(0.9), Inches(2.6), Inches(11.5), Inches(1.5), title, 40, True, WHITE)
    if subtitle:
        _txt(s, Inches(0.9), Inches(4.95), Inches(11), Inches(0.6), subtitle, 20, False, TEAL)
    _notes(s, bp.get("opening_notes", "Open with the one-line framing of this deck."))

    slides_in = (bp.get("slides") or [])

    for sl in slides_in[:24]:
        t = sl.get("type", "bullets")
        h = str(sl.get("h", "Section"))[:70]
        kick = str(sl.get("kicker", ""))[:40]
        s = _blank(prs); _header(s, h, kick)

        if t == "chart" and sl.get("chart"):
            cats, series = _safe_chart_data(sl["chart"])
            ctype = _CT.get(sl["chart"].get("ctype", "bar"), _CT["bar"])
            _chart(s, ctype, cats, series, Inches(0.8), Inches(1.85),
                   Inches(11.7), Inches(5.0), str(sl["chart"].get("title", h))[:60])

        elif t == "table" and (sl.get("table") or {}).get("rows"):
            rows = [[str(c)[:70] for c in r[:6]] for r in sl["table"]["rows"][:9]]
            nc = len(rows[0]); tw = Inches(12.1)
            cw = [int(tw / nc)] * nc
            _table(s, rows, Inches(0.6), Inches(1.95), tw, Inches(4.6), cw)

        elif t == "kpi" and sl.get("kpis"):
            kp = sl["kpis"]
            cards = [(str(k[0])[:22] if len(k) > 0 else "",
                      str(k[1])[:14] if len(k) > 1 else "",
                      str(k[2])[:20] if len(k) > 2 else "") for k in kp[:4]]
            _metric_cards(s, cards)
            if len(kp) > 4:
                rows = [["Metric", "Value", "Δ"]] + [[str(x)[:40] for x in (k[:3] + [""] * (3 - len(k[:3])))] for k in kp[4:10]]
                _table(s, rows, Inches(0.6), Inches(3.95), Inches(12.1), Inches(2.9),
                       [Inches(5.5), Inches(3.3), Inches(3.3)])

        elif t == "two_col":
            lt = str(sl.get("left_title", "Perspective A"))[:34]
            rt = str(sl.get("right_title", "Perspective B"))[:34]
            _two_col(s, lt, [str(p)[:130] for p in (sl.get("left") or [])[:6]],
                     rt, [str(p)[:130] for p in (sl.get("right") or [])[:6]])

        elif t == "hero" and sl.get("hero"):
            big = str(sl["hero"].get("value", "\u2014"))[:12]
            cap = str(sl["hero"].get("caption", ""))[:90]
            sup = [str(p)[:120] for p in (sl.get("points") or sl["hero"].get("support") or [])[:6]]
            _hero(s, big, cap, sup)

        else:
            pts = [str(p)[:160] for p in (sl.get("points") or ["Content"])[:7]]
            if len(pts) >= 4:
                # split long bullet lists into a left-rail styled layout for polish
                _bar(s, 0, Inches(1.7), Inches(0.9), SH - Inches(1.7), NAVY)
                _bullets(s, Inches(1.3), Inches(1.95), Inches(11.4), Inches(4.7), pts, 16)
            else:
                _bullets(s, Inches(0.8), Inches(1.95), Inches(11.8), Inches(4.6), pts, 17)

        _notes(s, str(sl.get("notes", f"Speak to: {h}"))[:400])

    # ── Closing ──
    s = _blank(prs)
    _bar(s, 0, 0, SW, SH, NAVY)
    _bar(s, 0, Inches(3.7), SW, Inches(0.06), TEAL)
    _txt(s, Inches(0.9), Inches(2.7), Inches(11.5), Inches(1.1), "Thank You", 44, True, WHITE)
    _txt(s, Inches(0.9), Inches(4.0), Inches(11.5), Inches(0.7), "Questions & Discussion", 20, False, TEAL)
    _notes(s, "Open the floor for questions.")

    while len(prs.slides) < 10:
        s = _blank(prs); _header(s, "Supplementary Analysis", "Appendix")
        _bullets(s, Inches(0.8), Inches(1.95), Inches(11.8), Inches(4),
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
from pdf_engine import (_cover, _hf, _tstyle, _bar_drawing,
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
            el.append(Paragraph(str(s["body"])[:2500], S_BODY))
        for b in (s.get("bullets") or [])[:8]:
            el.append(Paragraph("•  " + str(b)[:250], S_BODY))
        tab = (s.get("table") or {}).get("rows")
        if tab:
            rows = [[str(c)[:45] for c in r[:5]] for r in tab[:12]]
            cw = (W - 4 * cm) / len(rows[0])
            t = Table(rows, colWidths=[cw] * len(rows[0])); t.setStyle(_tstyle())
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
