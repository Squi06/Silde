# -*- coding: utf-8 -*-
"""Renderer nhẹ: đọc .pptx và vẽ PNG preview (để tự kiểm tra bố cục khi không có PowerPoint)."""

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR

HERE = Path(__file__).resolve().parent
PPI = 96  # 72dpi x 1.333 -> slide 1280x720

FONT_FILES = {
    "Montserrat": HERE / "assets/fonts/Montserrat.ttf",
    "Open Sans": HERE / "assets/fonts/OpenSans.ttf",
}
_fonts = {}


def font_for(name, bold, px):
    fam = "Montserrat" if name and "Montserrat" in name else "Open Sans"
    px = max(1, int(round(px)))
    key = (fam, bool(bold), px)
    if key not in _fonts:
        f = ImageFont.truetype(str(FONT_FILES[fam]), size=px)
        f.set_variation_by_name("Bold" if bold else "Regular")
        _fonts[key] = f
    return _fonts[key]


def px(emu):
    return float(emu) / 914400 * PPI


def color_of(obj, default=(14, 42, 34)):
    try:
        rgb = obj.rgb
        return tuple(int(str(rgb)[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return default


def solid_fill(shape):
    try:
        if shape.fill.type is not None and str(shape.fill.type).startswith("SOLID"):
            return color_of(shape.fill.fore_color)
    except Exception:
        pass
    return None


def line_style(shape):
    try:
        ln = shape.line
        if ln.fill.type is not None and str(ln.fill.type).startswith("SOLID"):
            w = ln.width.pt if ln.width else 1.0
            return color_of(ln.color), max(1, int(round(w * PPI / 72)))
    except Exception:
        pass
    return None, 0


def runs_of(p):
    """Trả về token theo đúng thứ tự XML, gồm cả line-break."""
    from pptx.oxml.ns import qn
    out = []
    for child in p._p:
        if child.tag == qn("a:br"):
            out.append({"br": True})
        elif child.tag == qn("a:r"):
            run = p.runs[len([o for o in out if "text" in o])] if False else None
            rPr = child.find(qn("a:rPr"))
            t = child.find(qn("a:t"))
            size = 18
            bold = False
            name = "Open Sans"
            color = (14, 42, 34)
            if rPr is not None:
                if rPr.get("sz"):
                    size = int(rPr.get("sz")) / 100
                bold = rPr.get("b") == "1"
                if rPr.get("i") is None:
                    pass
                latin = rPr.find(qn("a:latin"))
                if latin is not None:
                    name = latin.get("typeface") or name
                fill = rPr.find(qn("a:solidFill"))
                if fill is not None:
                    clr = fill.find(qn("a:srgbClr"))
                    if clr is not None:
                        v = clr.get("val")
                        color = tuple(int(v[i:i + 2], 16) for i in (0, 2, 4))
            out.append({"text": t.text or "", "size": size, "bold": bold,
                        "name": name, "color": color})
    return out


def wrap_paragraph(p, maxw):
    """Trả về list[list[(word, style, width, leading_space)]] theo dòng."""
    tokens = []
    for rs in runs_of(p):
        if rs.get("br"):
            tokens.append(("br", "", rs, False))
            continue
        words = rs["text"].split(" ")
        for i, w in enumerate(words):
            if w == "":
                continue
            tokens.append(("word", w, rs, i > 0))
    lines, cur, curw = [], [], 0.0
    for kind, word, rs, joined in tokens:
        if kind == "br":
            lines.append(cur)
            cur, curw = [], 0.0
            continue
        f = font_for(rs["name"], rs["bold"], rs["size"] * PPI / 72)
        wpx = f.getlength(word)
        sp = f.getlength(" ")
        need = wpx + (sp if (joined and cur) else 0)
        if cur and curw + need > maxw:
            lines.append(cur)
            cur, curw = [(word, rs, wpx, 0)], wpx
        else:
            cur.append((word, rs, wpx, sp if (joined and cur) else 0))
            curw += need
    if cur:
        lines.append(cur)
    return [ln for ln in lines if ln]


def draw_textbox(d, shape, alpha_img=None):
    tf = shape.text_frame
    x = px(shape.left) + px(tf.margin_left)
    y = px(shape.top) + px(tf.margin_top)
    maxw = px(shape.width) - px(tf.margin_left) - px(tf.margin_right)
    boxh = px(shape.height) - px(tf.margin_top) - px(tf.margin_bottom)

    blocks = []
    for p in tf.paragraphs:
        if not p.runs:
            continue
        spacing = p.line_spacing if isinstance(p.line_spacing, float) else 1.2
        lines = wrap_paragraph(p, maxw)
        if not lines:
            continue
        before = (p.space_before.pt if p.space_before else 0) * PPI / 72
        after = (p.space_after.pt if p.space_after else 0) * PPI / 72
        align = str(p.alignment) if p.alignment is not None else "LEFT"
        blocks.append((p, lines, spacing, before, after, align, maxw))

    total = sum(b[3] + sum(max(t[1]["size"] for t in ln) * b[2] * PPI / 72 for ln in b[1]) + b[4]
                for b in blocks)
    anchor = tf.vertical_anchor
    if anchor == MSO_ANCHOR.MIDDLE:
        y += (boxh - total) / 2
    elif anchor == MSO_ANCHOR.BOTTOM:
        y += boxh - total

    for p, lines, spacing, before, after, align, _ in blocks:
        y += before
        for ln in lines:
            lh = max(t[1]["size"] for t in ln) * spacing * PPI / 72
            lw = sum(t[2] + t[3] for t in ln)
            xoff = x
            if "CENTER" in align:
                xoff = x + (maxw - lw) / 2
            elif "RIGHT" in align:
                xoff = x + maxw - lw
            cx = xoff
            for word, rs, wpx, sp in ln:
                cx += sp
                f = font_for(rs["name"], rs["bold"], rs["size"] * PPI / 72)
                asc, desc = f.getmetrics()
                d.text((cx, y + (lh - (asc + desc)) / 2 + asc), word, font=f, fill=rs["color"],
                       anchor="ls")
                cx += wpx
            y += lh
        y += after


def draw_table(d, shape):
    tbl = shape.table
    x0, y0 = px(shape.left), px(shape.top)
    ys = y0
    for r, row in enumerate(tbl.rows):
        rh = px(row.height)
        xs = x0
        for c, col in enumerate(tbl.columns):
            cw = px(col.width)
            cell = tbl.cell(r, c)
            fill = color_of(cell.fill.fore_color) if cell.fill.type is not None else None
            if fill:
                d.rectangle([xs, ys, xs + cw, ys + rh], fill=fill)
            d.line([xs, ys + rh, xs + cw, ys + rh], fill=(213, 219, 203), width=1)
            ml = px(cell.margin_left)
            mt = px(cell.margin_top)
            text_y = ys + mt
            for p in cell.text_frame.paragraphs:
                if not p.runs:
                    continue
                lh_run = p.runs[0]
                size = lh_run.font.size.pt if lh_run.font.size else 12
                lh = size * 1.25 * PPI / 72
                cx = xs + ml
                for run in p.runs:
                    f = font_for(run.font.name, bool(run.font.bold),
                                 (run.font.size.pt if run.font.size else 12) * PPI / 72)
                    asc, desc = f.getmetrics()
                    d.text((cx, text_y + (lh - (asc + desc)) / 2 + asc), run.text, font=f,
                           fill=color_of(run.font.color), anchor="ls")
                    cx += f.getlength(run.text)
                text_y += lh
            xs += cw
        ys += rh


def render_slide(slide, W, H):
    img = Image.new("RGB", (W, H), (247, 246, 238))
    d = ImageDraw.Draw(img)
    try:
        bg = slide.background.fill
        if bg.type is not None and str(bg.type).startswith("SOLID"):
            d.rectangle([0, 0, W, H], fill=color_of(bg.fore_color))
    except Exception:
        pass

    for shape in slide.shapes:
        try:
            if shape.has_table:
                draw_table(d, shape)
                continue
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                import io
                blob = shape.image.blob
                im = Image.open(io.BytesIO(blob)).convert("RGBA")
                im = im.resize((max(1, int(px(shape.width))), max(1, int(px(shape.height)))),
                               Image.LANCZOS)
                try:
                    from pptx.oxml.ns import qn as _qn
                    if shape._element.spPr.find(_qn("a:prstGeom")).get("prst") == "ellipse":
                        from PIL import ImageDraw as _ID
                        mask = Image.new("L", im.size, 0)
                        _ID.Draw(mask).ellipse((0, 0, im.width - 1, im.height - 1), fill=255)
                        im.putalpha(mask)
                except Exception:
                    pass
                img.paste(im, (int(px(shape.left)), int(px(shape.top))), im)
                continue
            if shape.shape_type == MSO_SHAPE_TYPE.LINE:
                lc, lw = line_style(shape)
                if lc:
                    d.line([px(shape.left), px(shape.top),
                            px(shape.left) + px(shape.width), px(shape.top) + px(shape.height)],
                           fill=lc, width=lw)
                continue
            x, y = px(shape.left), px(shape.top)
            w, h = px(shape.width), px(shape.height)
            fill = solid_fill(shape)
            lc, lw = line_style(shape)
            try:
                geom = str(shape.auto_shape_type)
            except Exception:
                geom = ""

            if fill or lc:
                if "OVAL" in geom:
                    d.ellipse([x, y, x + w, y + h], fill=fill, outline=lc, width=lw or 1)
                elif "ROUNDED" in geom:
                    rad = 0.1
                    try:
                        rad = float(shape.adjustments[0])
                    except Exception:
                        pass
                    d.rounded_rectangle([x, y, x + w, y + h], radius=max(2, rad * min(w, h)),
                                        fill=fill, outline=lc, width=lw or 1)
                else:
                    d.rectangle([x, y, x + w, y + h], fill=fill, outline=lc, width=lw or 1)
            if shape.has_text_frame and shape.text_frame.text.strip():
                draw_textbox(d, shape)
        except Exception as e:
            print("  ! shape lỗi:", e)
    return img


def main(pptx_path=None, outdir=None):
    pptx_path = Path(pptx_path or HERE / "7UP-Brand-Deck.pptx")
    outdir = Path(outdir or HERE / "preview")
    outdir.mkdir(parents=True, exist_ok=True)
    prs = Presentation(str(pptx_path))
    W = int(px(prs.slide_width))
    H = int(px(prs.slide_height))
    imgs = []
    for i, slide in enumerate(prs.slides, start=1):
        img = render_slide(slide, W, H)
        p = outdir / f"slide-{i:02d}.png"
        img.save(p)
        imgs.append(img)
        print("  ->", p.name)

    cols, rows = 3, 4
    tw, th = W // 2, H // 2
    gap = 16
    sheet = Image.new("RGB", (cols * tw + gap * (cols + 1), rows * th + gap * (rows + 1)),
                      (28, 32, 28))
    for i, im in enumerate(imgs):
        r, c = divmod(i, cols)
        sheet.paste(im.resize((tw, th), Image.LANCZOS),
                    (gap + c * (tw + gap), gap + r * (th + gap)))
    sheet.save(outdir / "contact-sheet.png")
    print("  -> contact-sheet.png")
    return imgs


if __name__ == "__main__":
    main()
