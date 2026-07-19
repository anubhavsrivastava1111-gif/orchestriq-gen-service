"""
OrchestrIQ PowerPoint Engine v3 — McKinsey/BCG Grade
Real embedded charts, master slide, brand typography, speaker notes.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import ChartData, CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
import io
from datetime import datetime
from lxml import etree

# ── Brand ──────────────────────────────────────────────────────────────────────
def rgb(h): return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

NAVY  = rgb("1E3A5F")
TEAL  = rgb("14B8A6")
WHITE = rgb("FFFFFF")
LIGHT = rgb("F1F5F9")
MUTED = rgb("94A3B8")
DARK  = rgb("0F172A")
GREEN = rgb("10B981")
RED   = rgb("EF4444")
AMBER = rgb("F59E0B")

W = Inches(13.333)
H = Inches(7.5)

def _add_text(tf, text, size=12, bold=False, color=None, align=PP_ALIGN.LEFT):
    p = tf.paragraphs[0] if tf.paragraphs else tf.add_paragraph()
    p.alignment = align
    run = p.add_run()
    run.text = str(text or "")
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = "Calibri"
    if color:
        run.font.color.rgb = color

def _solid_fill(shape, color_rgb):
    """Apply solid fill to a shape."""
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = color_rgb

def _add_rect(slide, left, top, width, height, color_rgb):
    shape = slide.shapes.add_shape(1, left, top, width, height)  # MSO_SHAPE_TYPE.RECTANGLE
    _solid_fill(shape, color_rgb)
    shape.line.fill.background()
    return shape

def _add_header(slide, title_text, slide_num=None, total=None):
    """Standard header bar on every content slide."""
    bar = _add_rect(slide, 0, 0, W, Inches(0.85), NAVY)
    accent = _add_rect(slide, 0, 0, Inches(0.22), Inches(0.85), TEAL)

    tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.1), Inches(11.5), Inches(0.65))
    _add_text(tb.text_frame, title_text, size=20, bold=True, color=WHITE)

    if slide_num and total:
        nb = slide.shapes.add_textbox(Inches(12.3), Inches(0.1), Inches(0.9), Inches(0.65))
        nb.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
        run = nb.text_frame.paragraphs[0].add_run()
        run.text = f"{slide_num}/{total}"
        run.font.size = Pt(9)
        run.font.color.rgb = MUTED
        run.font.name = "Calibri"

def _add_footer(slide, company="", date=None):
    fb = slide.shapes.add_textbox(Inches(0.4), Inches(7.1), Inches(12.5), Inches(0.3))
    run = fb.text_frame.paragraphs[0].add_run()
    run.text = f"{company}  ·  Confidential  ·  {date or datetime.now().strftime('%d %b %Y')}"
    run.font.size = Pt(8)
    run.font.color.rgb = MUTED
    run.font.name = "Calibri"

def build_pptx(schema: dict, currency_symbol: str = "₹") -> bytes:
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    slides_data = schema.get("slides", [])
    total = len(slides_data)
    company = schema.get("company", "")
    today = datetime.now().strftime("%d %b %Y")

    # Remove default blank layout issues
    blank_layout = prs.slide_layouts[6]  # blank

    for idx, sd in enumerate(slides_data, 1):
        layout_type = str(sd.get("layout","full_text")).lower()
        slide = prs.slides.add_slide(blank_layout)

        if layout_type == "title":
            _build_title_slide(slide, sd, company, today)
        elif layout_type == "exec_summary":
            _build_exec_summary(slide, sd, idx, total, company, today)
        elif layout_type == "agenda":
            _build_agenda(slide, sd, idx, total, company, today)
        elif layout_type == "chart_narrative":
            _build_chart_narrative(slide, sd, idx, total, company, today, currency_symbol)
        elif layout_type == "two_column":
            _build_two_column(slide, sd, idx, total, company, today)
        elif layout_type == "data_table":
            _build_data_table(slide, sd, idx, total, company, today)
        elif layout_type == "closing":
            _build_closing(slide, sd, company, today)
        elif layout_type == "section_divider":
            _build_section_divider(slide, sd, idx)
        else:
            _build_full_text(slide, sd, idx, total, company, today)

        # Speaker notes
        notes = sd.get("speakerNotes") or sd.get("notes") or sd.get("speaker_notes","")
        if notes:
            tf = slide.notes_slide.notes_text_frame
            tf.text = str(notes)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()

def _build_title_slide(slide, sd, company, today):
    bg = _add_rect(slide, 0, 0, W, H, NAVY)
    accent = _add_rect(slide, 0, Inches(2.9), Inches(0.32), Inches(1.6), TEAL)

    tb = slide.shapes.add_textbox(Inches(0.6), Inches(2.7), Inches(11.5), Inches(1.0))
    _add_text(tb.text_frame, sd.get("title",""), size=38, bold=True, color=WHITE)

    sub = sd.get("subtitle") or sd.get("content","")
    if sub:
        sb = slide.shapes.add_textbox(Inches(0.6), Inches(3.8), Inches(11.5), Inches(0.7))
        _add_text(sb.text_frame, sub, size=20, color=TEAL)

    meta = sd.get("meta") or f"{company}  ·  {today}  ·  Confidential"
    mb = slide.shapes.add_textbox(Inches(0.6), Inches(4.7), Inches(11.5), Inches(0.4))
    _add_text(mb.text_frame, meta, size=13, color=MUTED)

    dt = slide.shapes.add_textbox(Inches(0.6), Inches(6.9), Inches(11.5), Inches(0.35))
    _add_text(dt.text_frame, f"Confidential  ·  Generated {today}", size=9, color=MUTED)

def _build_exec_summary(slide, sd, idx, total, company, today):
    _add_rect(slide, 0, 0, W, H, rgb("0A0E1A"))
    _add_header(slide, sd.get("title","Executive Summary"), idx, total)
    
    bullets = [b for b in str(sd.get("content","")).split("\n") if b.strip()]
    for i, b in enumerate(bullets[:5]):
        y = Inches(1.3) + i * Inches(0.95)
        # Number circle
        circ = slide.shapes.add_shape(9, Inches(0.35), y + Inches(0.08), Inches(0.46), Inches(0.46))
        _solid_fill(circ, TEAL)
        circ.line.fill.background()
        ct = circ.text_frame
        ct.word_wrap = False
        p = ct.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = str(i+1)
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = WHITE
        run.font.name = "Calibri"

        tb = slide.shapes.add_textbox(Inches(1.0), y, Inches(11.8), Inches(0.85))
        _add_text(tb.text_frame, b.lstrip("-•* "), size=15, color=WHITE)

    _add_footer(slide, company, today)

def _build_agenda(slide, sd, idx, total, company, today):
    _add_rect(slide, 0, 0, W, H, rgb("0A0E1A"))
    _add_header(slide, "AGENDA", idx, total)

    items = [b for b in str(sd.get("content","")).split("\n") if b.strip()]
    for i, item in enumerate(items[:7]):
        y = Inches(1.3) + i * Inches(0.74)
        num = slide.shapes.add_textbox(Inches(0.4), y, Inches(0.6), Inches(0.6))
        _add_text(num.text_frame, str(i+1).zfill(2), size=20, bold=True, color=TEAL)

        tb = slide.shapes.add_textbox(Inches(1.1), y + Inches(0.05), Inches(11.5), Inches(0.55))
        _add_text(tb.text_frame, item.lstrip("0123456789.-• "), size=16, color=WHITE)

        # Separator
        line = _add_rect(slide, Inches(0.4), y + Inches(0.62), Inches(12.5), Pt(0.5), rgb("263050"))

    _add_footer(slide, company, today)

def _build_chart_narrative(slide, sd, idx, total, company, today, currency_symbol):
    _add_rect(slide, 0, 0, W, H, rgb("0A0E1A"))
    _add_header(slide, sd.get("title",""), idx, total)

    chart_data_spec = sd.get("chartData") or {}
    labels = chart_data_spec.get("labels", [])
    series = chart_data_spec.get("series", [])
    chart_type_str = str(sd.get("chartType","bar")).lower()

    if labels and series:
        try:
            cd = CategoryChartData()
            cd.categories = [str(l) for l in labels[:12]]
            for s in series[:3]:
                cd.add_series(str(s.get("name","")), [float(v) if v is not None else 0 for v in s.get("values",[])[:12]])

            ct_map = {
                "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                "col": XL_CHART_TYPE.COLUMN_CLUSTERED,
                "line": XL_CHART_TYPE.LINE,
                "pie": XL_CHART_TYPE.PIE,
            }
            ct = ct_map.get(chart_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED)
            chart = slide.shapes.add_chart(ct, Inches(0.3), Inches(1.0), Inches(8.0), Inches(5.9), cd).chart
            chart.has_title = False
            chart.plots[0].has_data_labels = True
        except Exception:
            pass

    # Narrative bullets (right side)
    bullets = [b for b in str(sd.get("content","")).split("\n") if b.strip()]
    for i, b in enumerate(bullets[:6]):
        y = Inches(1.2) + i * Inches(0.85)
        tb = slide.shapes.add_textbox(Inches(8.5), y, Inches(4.6), Inches(0.8))
        p = tb.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = "▸  " + b.lstrip("-•* ")
        run.font.size = Pt(12)
        run.font.color.rgb = WHITE
        run.font.name = "Calibri"

    _add_footer(slide, company, today)

def _build_two_column(slide, sd, idx, total, company, today):
    _add_rect(slide, 0, 0, W, H, rgb("0A0E1A"))
    _add_header(slide, sd.get("title",""), idx, total)
    _add_rect(slide, Inches(6.5), Inches(1.0), Pt(1), Inches(6.0), rgb("263050"))

    content = str(sd.get("content",""))
    parts = content.split("---")
    left = parts[0].strip()
    right = parts[1].strip() if len(parts) > 1 else ""

    lt = slide.shapes.add_textbox(Inches(0.4), Inches(1.1), Inches(5.8), Inches(6.0))
    lt.text_frame.word_wrap = True
    for line in left.split("\n")[:8]:
        if line.strip():
            p = lt.text_frame.add_paragraph()
            run = p.add_run()
            run.text = line.lstrip("-•* ")
            run.font.size = Pt(13)
            run.font.color.rgb = WHITE
            run.font.name = "Calibri"
            p.space_after = Pt(6)

    rt = slide.shapes.add_textbox(Inches(6.8), Inches(1.1), Inches(6.2), Inches(6.0))
    rt.text_frame.word_wrap = True
    for line in right.split("\n")[:8]:
        if line.strip():
            p = rt.text_frame.add_paragraph()
            run = p.add_run()
            run.text = line.lstrip("-•* ")
            run.font.size = Pt(13)
            run.font.color.rgb = WHITE
            run.font.name = "Calibri"
            p.space_after = Pt(6)

    _add_footer(slide, company, today)

def _build_data_table(slide, sd, idx, total, company, today):
    _add_rect(slide, 0, 0, W, H, rgb("0A0E1A"))
    _add_header(slide, sd.get("title",""), idx, total)

    content = str(sd.get("content",""))
    table_lines = [l for l in content.split("\n") if "|" in l and not l.strip().replace("|","").replace("-","").replace(":","").replace(" ","") == ""]
    
    if not table_lines:
        tb = slide.shapes.add_textbox(Inches(0.4), Inches(1.1), Inches(12.5), Inches(6.0))
        _add_text(tb.text_frame, content, size=12, color=WHITE)
        _add_footer(slide, company, today)
        return

    rows = []
    for l in table_lines:
        cells = [c.strip() for c in l.split("|") if c.strip()]
        if cells:
            rows.append(cells)

    if len(rows) < 2:
        _add_footer(slide, company, today)
        return

    ncols = max(len(r) for r in rows)
    nrows = min(len(rows), 9)
    
    try:
        tbl = slide.shapes.add_table(nrows, ncols, Inches(0.4), Inches(1.1), Inches(12.5), Inches(5.8)).table
        for ri, row in enumerate(rows[:nrows]):
            for ci, val in enumerate(row[:ncols]):
                cell = tbl.cell(ri, ci)
                cell.text = val
                tf = cell.text_frame
                p = tf.paragraphs[0]
                run = p.add_run()
                run.font.name = "Calibri"
                run.font.size = Pt(11 if ri > 0 else 12)
                run.font.bold = ri == 0
                run.font.color.rgb = WHITE if ri == 0 else LIGHT
                p.alignment = PP_ALIGN.CENTER
                # Fill header row
                if ri == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = NAVY
                elif ri % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = rgb("131825")
    except:
        tb = slide.shapes.add_textbox(Inches(0.4), Inches(1.1), Inches(12.5), Inches(6.0))
        _add_text(tb.text_frame, content, size=10, color=WHITE)

    _add_footer(slide, company, today)

def _build_full_text(slide, sd, idx, total, company, today):
    _add_rect(slide, 0, 0, W, H, rgb("0A0E1A"))
    _add_header(slide, sd.get("title",""), idx, total)

    bullets = [b for b in str(sd.get("content","")).split("\n") if b.strip()]
    for i, b in enumerate(bullets[:8]):
        y = Inches(1.1) + i * Inches(0.72)
        tb = slide.shapes.add_textbox(Inches(0.5), y, Inches(12.5), Inches(0.65))
        p = tb.text_frame.paragraphs[0]
        run = p.add_run()
        is_sub = b.startswith("  ") or b.startswith("\t")
        run.text = "▸  " + b.lstrip("-•* \t")
        run.font.size = Pt(12 if is_sub else 14)
        run.font.color.rgb = MUTED if is_sub else WHITE
        run.font.name = "Calibri"

    _add_footer(slide, company, today)

def _build_closing(slide, sd, company, today):
    _add_rect(slide, 0, 0, W, H, NAVY)
    _add_rect(slide, 0, Inches(3.3), Inches(0.35), Inches(2.0), TEAL)

    tb = slide.shapes.add_textbox(Inches(0.7), Inches(3.1), Inches(11.5), Inches(1.0))
    _add_text(tb.text_frame, sd.get("title","Recommendations & Next Steps"), size=34, bold=True, color=WHITE)

    content = str(sd.get("content",""))
    actions = [a for a in content.split("\n") if a.strip()]
    for i, a in enumerate(actions[:5]):
        y = Inches(4.2) + i * Inches(0.65)
        tb = slide.shapes.add_textbox(Inches(0.7), y, Inches(11.5), Inches(0.6))
        _add_text(tb.text_frame, f"{i+1}.  {a.lstrip('0123456789.-• ')}", size=15, color=WHITE)

    ft = slide.shapes.add_textbox(Inches(0.7), Inches(7.0), Inches(11.5), Inches(0.35))
    _add_text(ft.text_frame, f"{company}  ·  Confidential  ·  {today}", size=9, color=MUTED)

def _build_section_divider(slide, sd, idx):
    _add_rect(slide, 0, 0, W, H, NAVY)
    num = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(3.0), Inches(3.0))
    p = num.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = str(idx).zfill(2)
    run.font.size = Pt(96)
    run.font.bold = True
    run.font.color.rgb = TEAL
    run.font.name = "Calibri"
    p.alignment = PP_ALIGN.LEFT

    tb = slide.shapes.add_textbox(Inches(3.2), Inches(2.8), Inches(9.8), Inches(1.2))
    _add_text(tb.text_frame, sd.get("title",""), size=36, bold=True, color=WHITE)

    sub = sd.get("content","")
    if sub:
        sb = slide.shapes.add_textbox(Inches(3.2), Inches(4.1), Inches(9.8), Inches(0.8))
        _add_text(sb.text_frame, sub, size=16, color=MUTED)
