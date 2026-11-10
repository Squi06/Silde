# -*- coding: utf-8 -*-
"""Design system dùng chung: màu, font, đo chữ và helper cho python-pptx."""

from pathlib import Path
from PIL import ImageFont
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

HERE = Path(__file__).resolve().parent

# ---------------------------------------------------------------- palette
INK = "0E2A22"
DEEP = "06382A"
GREEN = "00A94F"
GREEN_DK = "00763A"
LIME = "D3EE72"
CREAM = "F7F6EE"
PAPER = "FFFFFF"
RED = "E4002B"
MUTED = "5D7169"
SOFT = "E9EDE0"
LINE = "D5DBCB"
WHITE = "FFFFFF"
MINT = "EBF6E7"

FONT_DISPLAY = "Montserrat"
FONT_BODY = "Open Sans"

FONTS = {
    ("display", "regular"): (HERE / "assets/fonts/Montserrat.ttf", "Regular"),
    ("display", "medium"): (HERE / "assets/fonts/Montserrat.ttf", "Medium"),
    ("display", "semi"): (HERE / "assets/fonts/Montserrat.ttf", "SemiBold"),
    ("display", "bold"): (HERE / "assets/fonts/Montserrat.ttf", "Bold"),
    ("display", "xbold"): (HERE / "assets/fonts/Montserrat.ttf", "ExtraBold"),
    ("display", "black"): (HERE / "assets/fonts/Montserrat.ttf", "Black"),
    ("body", "regular"): (HERE / "assets/fonts/OpenSans.ttf", "Regular"),
    ("body", "semi"): (HERE / "assets/fonts/OpenSans.ttf", "SemiBold"),
    ("body", "bold"): (HERE / "assets/fonts/OpenSans.ttf", "Bold"),
}

_cache = {}


def measure(text, family="body", weight="regular", size=12):
    # PPTX chỉ có regular/bold: "semi" được vẽ thành Bold nên đo bằng Bold cho khớp
    if weight == "semi":
        weight = "bold"
    key = (family, weight, size)
    if key not in _cache:
        path, inst = FONTS[(family, weight)]
        f = ImageFont.truetype(str(path), size=size)
        try:
            f.set_variation_by_name(inst)
        except Exception:
            pass
        _cache[key] = f
    return _cache[key].getlength(text)


def wrap_lines(runs, width_pt):
    """runs: list[(text, family, weight, size)] -> số dòng sau khi wrap (tôn trọng \\n)."""
    words = []
    for text, fam, wgt, size in runs:
        segs = text.split("\n")
        for si, seg in enumerate(segs):
            if si > 0:
                words.append(("\n", fam, wgt, size, False))
            for i, chunk in enumerate(seg.split(" ")):
                if chunk == "":
                    continue
                words.append((chunk, fam, wgt, size, i > 0))
    lines, cur = 1, 0.0
    space_w = measure(" ", "body", "regular", 12) / 12.0
    for word, fam, wgt, size, joined in words:
        if word == "\n":
            lines += 1
            cur = 0.0
            continue
        w = measure(word, fam, wgt, size)
        add = w + (space_w * size if joined and cur > 0 else 0)
        if cur > 0 and cur + add > width_pt:
            lines += 1
            cur = w
        else:
            cur += add
    return lines


def block_height(runs, width_pt, spacing=1.22):
    """Ước lượng chiều cao (pt) của một đoạn văn đã wrap."""
    lines = wrap_lines(runs, width_pt)
    size = max(r[3] for r in runs)
    return lines * size * spacing


def report(name, runs, w_in, h_in, spacing=1.22, quiet=False):
    need = block_height(runs, w_in * 72, spacing)
    have = h_in * 72
    status = "OK " if need <= have + 1 else "OVER"
    if not quiet:
        print(f"  [{status}] {name:<34} need {need/72:5.2f}in / box {have/72:5.2f}in")
    return need <= have + 1


# ---------------------------------------------------------------- pptx helpers
def rgb(hexstr):
    return RGBColor.from_string(hexstr)


def set_bg(slide, hexcolor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(hexcolor)


def _apply_run_font(run, family, weight, size, color, spc=0, italic=False):
    f = run.font
    f.name = FONT_DISPLAY if family == "display" else FONT_BODY
    f.size = Pt(size)
    f.bold = weight in ("semi", "bold", "xbold", "black")
    f.italic = italic
    f.color.rgb = rgb(color)
    rPr = f._rPr
    rPr.set("lang", "vi-VN")
    if spc:
        rPr.set("spc", str(int(spc * 100)))
    # đảm bảo font áp cho cả east-asian / complex script
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", FONT_DISPLAY if family == "display" else FONT_BODY)


def text(slide, x, y, w, h, paras, anchor="top", margins=(0, 0, 0, 0), wrap=True):
    """paras: list of dict(runs=[dict(t,size,weight,color,fam,spc)], align, line, before, after)"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    l, t, r, b = margins
    tf.margin_left, tf.margin_top = Inches(l), Inches(t)
    tf.margin_right, tf.margin_bottom = Inches(r), Inches(b)
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                          "bottom": MSO_ANCHOR.BOTTOM}[anchor]
    first = True
    for p in paras:
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        para.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                          "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}[p.get("align", "left")]
        if p.get("line"):
            para.line_spacing = p["line"]
        if p.get("before"):
            para.space_before = Pt(p["before"])
        if p.get("after"):
            para.space_after = Pt(p["after"])
        for rd in p["runs"]:
            parts = rd["t"].split("\n")
            for i, part in enumerate(parts):
                if i > 0:
                    para.add_line_break()
                if not part:
                    continue
                run = para.add_run()
                run.text = part
                _apply_run_font(run, rd.get("fam", "body"), rd.get("weight", "regular"),
                                rd["size"], rd.get("color", INK), rd.get("spc", 0),
                                rd.get("italic", False))
    return box


def P(t, size, weight="regular", color=INK, fam="body", **kw):
    d = {"t": t, "size": size, "weight": weight, "color": color, "fam": fam}
    d.update(kw)
    return d


def para(runs, align="left", line=None, before=0, after=0):
    return {"runs": runs if isinstance(runs, list) else [runs],
            "align": align, "line": line, "before": before, "after": after}


def rect(slide, x, y, w, h, fill=None, line=None, line_w=1.0, radius=None, shape=MSO_SHAPE.RECTANGLE,
         transparency=0):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
        if transparency:
            try:
                spPr = shp._element.spPr
                srgb = spPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
                alpha = srgb.makeelement(qn("a:alpha"),
                                         {"val": str(int((1 - transparency) * 100000))})
                srgb.append(alpha)
            except Exception:
                pass
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(line_w)
    shp.text_frame.word_wrap = True
    return shp


def circle(slide, x, y, d, fill, line=None, line_w=1.0):
    return rect(slide, x, y, d, d, fill=fill, line=line, line_w=line_w, shape=MSO_SHAPE.OVAL)


def hline(slide, x, y, w, color=LINE, weight=0.75):
    from pptx.enum.shapes import MSO_CONNECTOR
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y),
                                    Inches(x + w), Inches(y))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(weight)
    return ln


def vline(slide, x, y, h, color=LINE, weight=0.75):
    from pptx.enum.shapes import MSO_CONNECTOR
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y),
                                    Inches(x), Inches(y + h))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(weight)
    return ln
