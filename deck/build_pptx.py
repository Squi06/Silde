# -*- coding: utf-8 -*-
"""Sinh bộ slide 7UP (.pptx) – 16:9, thiết kế native để import vào Canva."""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from PIL import Image, ImageDraw

import content as C
from design import (
    INK, DEEP, GREEN, GREEN_DK, LIME, CREAM, PAPER, RED, MUTED, SOFT, LINE,
    WHITE, MINT, FONT_DISPLAY, FONT_BODY,
    text, P, para, rect, circle, hline, vline, set_bg, rgb, report, measure,
)

HERE = Path(__file__).resolve().parent
SW, SH = 13.333, 7.5
ML, CW = 0.85, 11.633
TOTAL = 12


# ------------------------------------------------------------------ chrome
def header(slide, kicker, title, size=29, dark=False):
    ink = CREAM if dark else INK
    text(slide, ML, 0.55, 8.0, 0.24,
         [para(P(kicker, 9, "bold", LIME if dark else GREEN, "display", spc=2.4))])
    text(slide, ML, 0.84, 11.4, 1.0,
         [para(P(title, size, "bold", ink, "display"), line=1.02)])
    brand(slide, 11.55, 0.62, dark)


def brand(slide, x, y, dark=False):
    text(slide, x, y, 0.95, 0.26,
         [para([P("7", 13, "black", CREAM if dark else INK, "display"),
                P("UP", 13, "black", CREAM if dark else INK, "display")], align="right")])
    circle(slide, x + 0.97, y + 0.02, 0.11, RED)


def footer(slide, n, dark=False):
    col = "2E5646" if dark else LINE
    muted = "8FB2A3" if dark else MUTED
    hline(slide, ML, 6.98, CW, col, 0.75)
    text(slide, ML, 7.08, 6.0, 0.2,
         [para(P("7UP · BRAND DECK · 2026", 7.5, "semi", muted, "display", spc=2.0))])
    text(slide, ML + CW - 1.2, 7.08, 1.2, 0.2,
         [para(P(f"{n:02d} / {TOTAL}", 7.5, "semi", muted, "display", spc=1.4), align="right")])


def new_slide(prs, bg=CREAM):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, bg)
    return s


def card(slide, x, y, w, h, fill=PAPER, radius=0.07):
    return rect(slide, x, y, w, h, fill=fill, radius=radius, shape=MSO_SHAPE.ROUNDED_RECTANGLE)


def chip(slide, x, y, label, fg=GREEN_DK, bg=None, size=8, pad=0.3, h=0.3):
    w = measure(label, "display", "bold", size) / 72 + pad
    if bg:
        rect(slide, x, y, w, h, fill=bg, radius=0.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(slide, x + pad / 2, y + (h - size / 72 * 1.35) / 2 + 0.01, w - pad, h,
         [para(P(label, size, "semi", fg, "display", spc=0.6))])
    return w




def circle_photo(slide, src, x, y, d):
    """Chèn ảnh và cắt theo hình tròn bằng prstGeom=ellipse (vẫn là ảnh vector-crop trong Canva)."""
    pic = slide.shapes.add_picture(str(src), Inches(x), Inches(y), Inches(d), Inches(d))
    spPr = pic._element.spPr
    geom = spPr.find(qn("a:prstGeom"))
    if geom is not None:
        geom.set("prst", "ellipse")
    return pic


# ------------------------------------------------------------------ slides
def s01_cover(prs):
    s = new_slide(prs, DEEP)
    circle_photo(s, HERE / "assets/hero-lime.jpg", 8.05, 0.88, 4.95)
    circle(s, 7.55, 0.52, 0.95, RED)

    text(s, 0.82, 0.6, 7.0, 0.24,
         [para(P(C.COVER["kicker"], 9.5, "bold", LIME, "display", spc=3.0))])
    text(s, 0.75, 1.18, 7.6, 2.75,
         [para(P(C.COVER["wordmark"], 186, "black", CREAM, "display"), line=0.94)])
    text(s, 0.85, 4.28, 6.6, 1.5,
         [para(P(line, 28, "bold", CREAM, "display"), line=1.1) for line in
          C.COVER["tagline"].split("\n")])
    text(s, 0.85, 5.62, 6.4, 0.85,
         [para(P(line, 11.5, "regular", "A5C2B4", "body"), line=1.4) for line in
          C.COVER["sub"].split("\n")])
    hline(s, 0.85, 6.62, 6.55, "2E5646", 0.75)
    text(s, 0.85, 6.74, 3.0, 0.22,
         [para(P(C.COVER["meta_left"].upper(), 8, "semi", "8FB2A3", "display", spc=2.2))])
    text(s, 4.4, 6.74, 3.0, 0.22,
         [para(P(C.COVER["meta_right"].upper(), 8, "semi", "8FB2A3", "display", spc=2.2), align="right")])
    footer(s, 1, dark=True)
    return s


def s02_agenda(prs):
    s = new_slide(prs)
    header(s, C.AGENDA["kicker"], "Sáu phần,\nmột câu chuyện thương hiệu", 30)
    circle(s, 11.35, -1.5, 3.2, LIME)
    circle(s, 10.95, 0.05, 0.42, RED)
    x0, w, h, gap = ML, 5.66, 1.34, 0.19
    ys = [2.32, 3.85 - 0.19, 5.38 - 0.38]
    for i, (num, title, desc) in enumerate(C.AGENDA["items"]):
        col = i // 3
        row = i % 3
        x = x0 + col * (w + gap + 0.31)
        y = ys[row]
        hline(s, x, y, w, LINE, 0.75)
        text(s, x, y + 0.16, 0.85, 0.55,
             [para(P(num, 26, "black", "D6DECB", "display"), line=1.0)])
        text(s, x + 0.85, y + 0.22, w - 0.85, 0.3,
             [para(P(title, 13.5, "bold", INK, "display"))])
        text(s, x + 0.85, y + 0.62, w - 0.9, 0.5,
             [para(P(desc, 9.5, "regular", MUTED, "body"), line=1.35)])
        report(f"agenda item {num}", [(desc, "body", "regular", 9.5)], w - 0.9, 0.5)
    footer(s, 2)
    return s


def s03_snapshot(prs):
    s = new_slide(prs)
    header(s, C.SNAPSHOT["kicker"], "Một thương hiệu 95 năm,\nhiện diện hơn 140 quốc gia", 28)

    tile_w, gap = 2.72, 0.25
    for i, (big, label, sub) in enumerate(C.SNAPSHOT["stats"]):
        x = ML + i * (tile_w + gap)
        card(s, x, 2.3, tile_w, 1.42)
        text(s, x + 0.28, 2.44, tile_w - 0.5, 0.6,
             [para(P(big, 30, "xbold", GREEN, "display"), line=1.0)])
        text(s, x + 0.28, 3.06, tile_w - 0.5, 0.24,
             [para(P(label, 9, "bold", INK, "display"))])
        text(s, x + 0.28, 3.34, tile_w - 0.5, 0.24,
             [para(P(sub, 8.5, "regular", MUTED, "body"))])

    text(s, ML, 4.02, 6.55, 1.05,
         [para(P(C.SNAPSHOT["lead"], 13.5, "medium", INK, "display"), line=1.4)])
    text(s, ML, 5.18, 6.5, 1.4,
         [para(P(C.SNAPSHOT["body"], 10, "regular", MUTED, "body"), line=1.5)])
    report("snapshot body", [(C.SNAPSHOT["body"], "body", "regular", 10)], 6.5, 1.4, 1.5)

    fx, fw = 7.85, 4.63
    for i, (label, val) in enumerate(C.SNAPSHOT["facts"]):
        y = 4.02 + i * 0.94
        hline(s, fx, y, fw, LINE, 0.75)
        text(s, fx, y + 0.12, fw, 0.2,
             [para(P(label.upper(), 8, "bold", GREEN_DK, "display", spc=1.8))])
        text(s, fx, y + 0.36, fw, 0.5,
             [para(P(val, 9.5, "regular", INK, "body"), line=1.35)])
        report(f"fact {label}", [(val, "body", "regular", 9.5)], fw, 0.5, 1.35)
    footer(s, 3)
    return s


def s04_timeline(prs):
    s = new_slide(prs)
    header(s, C.TIMELINE["kicker"], "Tám dấu mốc trong 95 năm", 29)
    text(s, ML, 1.62, 10.6, 0.3,
         [para(P(C.TIMELINE["lead"], 10, "regular", MUTED, "body"), line=1.4)])
    report("timeline lead", [(C.TIMELINE["lead"], "body", "regular", 10)], 10.6, 0.3)

    w, gap = 2.66, 0.33
    for i, (year, head, desc) in enumerate(C.TIMELINE["nodes"]):
        row, col = i // 4, i % 4
        x = ML + col * (w + gap)
        y = 2.3 + row * 2.42
        hline(s, x, y, w, LINE, 0.75)
        circle(s, x - 0.01, y - 0.055, 0.11, GREEN if row == 0 else RED)
        text(s, x, y + 0.14, w, 0.32,
             [para(P(year, 17, "xbold", INK, "display"))])
        text(s, x, y + 0.5, w, 0.26,
             [para(P(head, 10.5, "bold", GREEN_DK, "display", spc=0.4))])
        text(s, x, y + 0.78, w - 0.05, 0.85,
             [para(P(desc, 9, "regular", MUTED, "body"), line=1.4)])
        report(f"timeline {year}", [(desc, "body", "regular", 9)], w - 0.05, 0.85, 1.4)
    footer(s, 4)
    return s


def s05_positioning(prs):
    s = new_slide(prs)
    header(s, C.POSITIONING["kicker"], "\u201cUncola\u201d – đứng ngoài cuộc chơi cola", 29)

    panel = rect(s, ML, 2.02, CW, 1.5, fill=DEEP, radius=0.06, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    circle(s, 1.18, 2.69, 0.16, LIME)
    text(s, 1.56, 2.02, 10.45, 1.5,
         [para(P(C.POSITIONING["statement"], 14, "medium", CREAM, "display"), line=1.45)],
         anchor="middle")
    report("positioning statement", [(C.POSITIONING["statement"], "display", "medium", 14)], 10.45, 1.0, 1.45)

    w, gap, y = 2.72, 0.25, 3.78
    for i, (head, body_) in enumerate(C.POSITIONING["pillars"]):
        x = ML + i * (w + gap)
        card(s, x, y, w, 2.35)
        text(s, x + 0.28, y + 0.26, w - 0.56, 0.24,
             [para(P(f"0{i+1}", 8.5, "bold", GREEN, "display", spc=1.8))])
        text(s, x + 0.28, y + 0.54, w - 0.56, 0.55,
             [para(P(head, 12.5, "bold", INK, "display"), line=1.15)])
        text(s, x + 0.28, y + 1.12, w - 0.54, 1.1,
             [para(P(body_, 9.5, "regular", MUTED, "body"), line=1.45)])
        report(f"pillar {head}", [(body_, "body", "regular", 9.5)], w - 0.54, 1.1, 1.45)

    text(s, ML, 6.36, 11.4, 0.28,
         [para(P(C.POSITIONING["promise"].upper(), 10, "bold", INK, "display", spc=4.0), align="center")])
    footer(s, 5)
    return s


def s06_portfolio(prs):
    s = new_slide(prs)
    header(s, C.PORTFOLIO["kicker"], "Một hương vị gốc, nhiều biến thể theo nhu cầu", 28)
    w, gap = 3.72, 0.235
    for i, (name, tag, desc) in enumerate(C.PORTFOLIO["items"]):
        row, col = i // 3, i % 3
        x = ML + col * (w + gap)
        y = 2.05 + row * 2.08
        card(s, x, y, w, 1.92)
        text(s, x + 0.3, y + 0.22, w - 0.6, 0.32,
             [para(P(name, 13, "bold", INK, "display"))])
        text(s, x + 0.3, y + 0.6, w - 0.6, 0.22,
             [para(P(tag.upper(), 8, "bold", GREEN_DK, "display", spc=1.6))])
        text(s, x + 0.3, y + 0.88, w - 0.58, 0.9,
             [para(P(desc, 9, "regular", MUTED, "body"), line=1.45)])
        report(f"portfolio {name}", [(desc, "body", "regular", 9)], w - 0.58, 0.9, 1.45)
    text(s, ML, 6.35, 11.4, 0.22,
         [para(P("Danh mục khác nhau theo thị trường; Việt Nam tập trung vào 7UP Original và 7UP Revive.",
                 8.5, "regular", MUTED, "body", italic=True))])
    footer(s, 6)
    return s


def s07_audience(prs):
    s = new_slide(prs)
    header(s, C.AUDIENCE["kicker"], "Ba nhóm người uống, ba dịp uống", 29)
    w, gap = 3.72, 0.235
    for i, (letter, rng, name, desc, tags) in enumerate(C.AUDIENCE["personas"]):
        x = ML + i * (w + gap)
        y = 2.02
        card(s, x, y, w, 3.62)
        circle(s, x + 0.3, y + 0.3, 0.46, DEEP)
        text(s, x + 0.3, y + 0.39, 0.46, 0.28,
             [para(P(letter, 11, "bold", CREAM, "display"), align="center")])
        text(s, x + 0.92, y + 0.29, w - 1.2, 0.22,
             [para(P(rng, 9, "bold", GREEN_DK, "display", spc=1.4))])
        text(s, x + 0.92, y + 0.52, w - 1.2, 0.3,
             [para(P(name, 12.5, "bold", INK, "display"))])
        text(s, x + 0.3, y + 1.02, w - 0.58, 1.15,
             [para(P(desc, 9, "regular", MUTED, "body"), line=1.5)])
        report(f"persona {letter}", [(desc, "body", "regular", 9)], w - 0.58, 1.1, 1.5)
        text(s, x + 0.3, y + 1.82, w - 0.6, 0.2,
             [para(P("DỊP UỐNG & ĐIỂM CHẠM", 7.5, "bold", MUTED, "display", spc=1.6))])
        for j, tag in enumerate(tags):
            chip(s, x + 0.3, y + 2.12 + j * 0.38, tag, GREEN_DK, MINT, 8.5, 0.34, 0.3)

    bar = rect(s, ML, 5.86, CW, 0.86, fill=DEEP, radius=0.08, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, ML, 5.86, 0.09, 0.86, fill=LIME)
    text(s, ML + 0.4, 5.98, 7.5, 0.3,
         [para(P(C.AUDIENCE["insight"], 12, "medium", CREAM, "display"), line=1.3)])
    text(s, ML + 8.1, 6.04, 3.15, 0.55,
         [para(P(C.AUDIENCE["insight_note"], 8.5, "regular", "A5C2B4", "body"), line=1.4)])
    report("insight note", [(C.AUDIENCE["insight_note"], "body", "regular", 8.5)], 3.15, 0.55, 1.4)
    footer(s, 7)
    return s


def s08_marketing(prs):
    s = new_slide(prs)
    header(s, C.MARKETING["kicker"], "Bốn chiến dịch định hình cách nói \u201ckhác biệt\u201d", 28)
    w, gap = 2.72, 0.25
    for i, (year, name, desc) in enumerate(C.MARKETING["campaigns"]):
        x = ML + i * (w + gap)
        card(s, x, 2.02, w, 3.18)
        rect(s, x, 2.02, w, 0.07, fill=GREEN if i % 2 == 0 else RED)
        text(s, x + 0.26, 2.28, w - 0.5, 0.24,
             [para(P(year, 9.5, "bold", GREEN_DK, "display", spc=2.0))])
        text(s, x + 0.26, 2.56, w - 0.5, 0.62,
             [para(P(name, 14, "bold", INK, "display"), line=1.15)])
        text(s, x + 0.26, 3.26, w - 0.52, 1.7,
             [para(P(desc, 8.8, "regular", MUTED, "body"), line=1.5)])
        report(f"campaign {name}", [(desc, "body", "regular", 8.8)], w - 0.52, 1.7, 1.5)

    y = 5.52
    hline(s, ML, y, CW, LINE, 0.75)
    wlabel = text(s, ML, y + 0.18, 1.7, 0.24,
                  [para(P("VŨ KHÍ LẶP LẠI", 8, "bold", MUTED, "display", spc=1.8))])
    x = ML + 1.62
    for tag in C.MARKETING["tactics"]:
        wd = chip(s, x, y + 0.12, tag, INK, SOFT, 8.5, 0.3, 0.34)
        x += wd + 0.14
    footer(s, 8)
    return s


def _table_style(tbl_shape):
    tbl = tbl_shape.table
    tbl.first_row = True
    tbl.horz_banding = False
    tbl.vert_banding = False
    tblPr = tbl._tbl.tblPr
    for attr in ("firstRow", "bandRow"):
        tblPr.set(attr, "0")
    style_id = tblPr.find(qn("a:tableStyleId"))
    if style_id is None:
        style_id = tblPr.makeelement(qn("a:tableStyleId"), {})
        tblPr.append(style_id)
    style_id.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    return tbl


def _cell_border(cell, edges=("B",), color=LINE, width=0.75):
    tcPr = cell._tc.get_or_add_tcPr()
    order = {"L": "a:lnL", "R": "a:lnR", "T": "a:lnT", "B": "a:lnB"}
    for e in edges:
        tag = order[e]
        el = tcPr.find(qn(tag))
        if el is not None:
            tcPr.remove(el)
        ln = tcPr.makeelement(qn(tag), {"w": str(int(width * 12700)), "cap": "flat",
                                        "cmpd": "sng", "algn": "ctr"})
        fill = ln.makeelement(qn("a:solidFill"), {})
        clr = fill.makeelement(qn("a:srgbClr"), {"val": color})
        fill.append(clr)
        ln.append(fill)
        # chèn đúng thứ tự schema (lnL, lnR, lnT, lnB trước fill)
        tcPr.insert(0, ln)


def s09_competition(prs):
    s = new_slide(prs)
    header(s, "06A · THỊ TRƯỜNG", "Sân chơi chanh–lime: 7UP giữ vị thế người khai phá", 28)

    cols = [2.35, 2.45, 2.25, 2.35, 2.233]
    nrows = len(C.COMPETITION["rows"]) + 1
    row_h = [0.56] + [0.6] * (nrows - 1)
    gf = s.shapes.add_table(nrows, 5, Inches(ML), Inches(2.1), Inches(CW), Inches(sum(row_h)))
    tbl = _table_style(gf)
    for i, cw in enumerate(cols):
        tbl.columns[i].width = Inches(cw)
    for i, rh in enumerate(row_h):
        tbl.rows[i].height = Inches(rh)

    for r in range(nrows):
        for c in range(5):
            cell = tbl.cell(r, c)
            cell.margin_left = Inches(0.14)
            cell.margin_right = Inches(0.1)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = rgb(DEEP)
            elif c == 1:
                cell.fill.fore_color.rgb = rgb("E4F3D6")
            else:
                cell.fill.fore_color.rgb = rgb(PAPER if r % 2 else "FBFAF4")
            if r > 0:
                _cell_border(cell, ("B",), LINE, 0.75)

    for c, head in enumerate(C.COMPETITION["columns"]):
        cell = tbl.cell(0, c)
        cell.text = ""
        p = cell.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = head if head else "Tiêu chí"
        run.font.name = FONT_DISPLAY
        run.font.size = Pt(9.5)
        run.font.bold = True
        run.font.color.rgb = rgb(LIME if c == 1 else CREAM)

    for r, row in enumerate(C.COMPETITION["rows"], start=1):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = ""
            p = cell.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = val
            run.font.name = FONT_DISPLAY if c == 0 else FONT_BODY
            run.font.size = Pt(9 if c == 0 else 10)
            run.font.bold = c == 0 or c == 1
            run.font.color.rgb = rgb(MUTED if c == 0 else INK)

    ny = 2.1 + sum(row_h) + 0.28
    for i, note in enumerate(C.COMPETITION["notes"]):
        y = ny + i * 0.45
        circle(s, ML + 0.02, y + 0.09, 0.07, GREEN)
        text(s, ML + 0.25, y, 11.1, 0.42,
             [para(P(note, 9.5, "regular", MUTED, "body"), line=1.35)])
        report(f"competition note {i}", [(note, "body", "regular", 9.5)], 11.1, 0.42, 1.35)
    footer(s, 9)
    return s


def s10_swot(prs):
    s = new_slide(prs)
    header(s, "06B · SWOT", "Điểm mạnh nhất cũng là điểm yếu dễ thấy nhất", 28)
    w, h = 5.72, 2.08
    pos = [(ML, 2.02), (ML + w + 0.19, 2.02), (ML, 4.25), (ML + w + 0.19, 4.25)]
    accents = [GREEN, "B9A24B", GREEN_DK, RED]
    for i, (letter, name, items) in enumerate(C.SWOT["quadrants"]):
        x, y = pos[i]
        card(s, x, y, w, h)
        rect(s, x, y, 0.075, h, fill=accents[i])
        text(s, x + 0.28, y + 0.18, 0.5, 0.42,
             [para(P(letter, 24, "black", accents[i], "display"), line=1.0)])
        text(s, x + 0.8, y + 0.26, 4.6, 0.24,
             [para(P(name.upper(), 8.5, "bold", MUTED, "display", spc=2.2))])
        paras = []
        for it in items:
            paras.append(para([P("–  ", 9.5, "bold", accents[i], "body"),
                               P(it, 9.5, "regular", INK, "body")], line=1.9))
        text(s, x + 0.28, y + 0.72, w - 0.6, 1.25, paras)
        for it in items:
            report(f"swot {letter}", [("–  ", "body", "bold", 9.5), (it, "body", "regular", 9.5)],
                   w - 0.6, 1.25 / len(items) - 0.03, 1.9)
    footer(s, 10)
    return s


def s11_recommend(prs):
    s = new_slide(prs)
    header(s, "06C · KHUYẾN NGHỊ", "Bốn việc 7UP nên làm trong 18 tháng tới", 28)

    w, gap = 3.72, 0.235
    for i, (head, desc) in enumerate(C.RECOMMEND["trends"]):
        x = ML + i * (w + gap)
        rect(s, x, 1.98, w, 0.98, fill=SOFT, radius=0.08, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text(s, x + 0.24, 2.08, w - 0.46, 0.24,
             [para(P(head, 10, "bold", INK, "display"))])
        text(s, x + 0.24, 2.34, w - 0.46, 0.56,
             [para(P(desc, 8.3, "regular", MUTED, "body"), line=1.35)])
        report(f"trend {head}", [(desc, "body", "regular", 8.3)], w - 0.46, 0.56, 1.35)

    for i, (num, head, desc) in enumerate(C.RECOMMEND["actions"]):
        y = 3.2 + i * 0.85
        hline(s, ML, y, CW, LINE, 0.75)
        text(s, ML, y + 0.2, 0.6, 0.3,
             [para(P(num, 11, "bold", GREEN, "display", spc=1.2))])
        text(s, ML + 0.62, y + 0.16, 3.4, 0.6,
             [para(P(head, 11.5, "bold", INK, "display"), line=1.2)])
        text(s, ML + 4.2, y + 0.18, 7.35, 0.55,
             [para(P(desc, 9.2, "regular", MUTED, "body"), line=1.4)])
        report(f"action {num}", [(desc, "body", "regular", 9.2)], 7.35, 0.55, 1.4)
    hline(s, ML, 3.2 + 4 * 0.85, CW, LINE, 0.75)
    footer(s, 11)
    return s


def s12_closing(prs):
    s = new_slide(prs, DEEP)
    circle(s, 10.35, 4.35, 3.75, LIME)
    circle(s, 10.62, 4.72, 0.52, RED)

    text(s, ML, 0.62, 6.0, 0.24,
         [para(P(C.CLOSING["kicker"], 9.5, "bold", LIME, "display", spc=3.2))])
    text(s, ML, 0.95, 8.4, 1.6,
         [para(P(C.CLOSING["title"], 31, "bold", CREAM, "display"), line=1.1)])
    report("closing title", [(C.CLOSING["title"], "display", "bold", 31)], 8.4, 1.6, 1.1)

    for i, pt_ in enumerate(C.CLOSING["points"]):
        y = 2.95 + i * 0.62
        circle(s, ML + 0.02, y + 0.1, 0.1, GREEN)
        text(s, ML + 0.35, y, 6.9, 0.55,
             [para(P(pt_, 11.5, "regular", "D6E7DC", "body"), line=1.35)])
        report(f"closing point {i}", [(pt_, "body", "regular", 11.5)], 6.9, 0.55, 1.35)

    hline(s, ML, 5.28, 6.9, "2E5646", 0.75)
    text(s, ML, 5.42, 6.9, 0.5,
         [para(P(C.CLOSING["thanks"], 22, "bold", CREAM, "display"))])
    text(s, ML, 6.02, 6.9, 0.3,
         [para(P(C.CLOSING["qa"].upper(), 9, "semi", LIME, "display", spc=2.6))])
    text(s, ML, 6.55, 8.2, 0.35,
         [para(P(C.CLOSING["sources"], 7.3, "regular", "7FA294", "body"), line=1.3)])
    brand(s, 11.55, 0.62, dark=True)
    return s


def build(out=None):
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    print("Đo & kiểm tra chữ:")
    for fn in (s01_cover, s02_agenda, s03_snapshot, s04_timeline, s05_positioning,
               s06_portfolio, s07_audience, s08_marketing, s09_competition,
               s10_swot, s11_recommend, s12_closing):
        fn(prs)
    for i, slide in enumerate(prs.slides, start=1):
        note = C.NOTES.get(i)
        if note:
            slide.notes_slide.notes_text_frame.text = note

    out = Path(out or HERE / "7UP-Brand-Deck.pptx")
    prs.core_properties.title = "7UP – Brand Deck"
    prs.core_properties.author = "Silde"
    prs.core_properties.subject = "Phân tích thương hiệu 7UP"
    prs.save(out)
    print(f"\nĐã lưu: {out} ({out.stat().st_size/1024:.0f} KB)")
    return out


if __name__ == "__main__":
    build()
