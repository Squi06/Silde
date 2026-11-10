# -*- coding: utf-8 -*-
"""Chuyển .pptx -> một file HTML tự chứa (font nhúng base64, điều hướng bằng bàn phím)."""

import base64
import io
import subprocess
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn

HERE = Path(__file__).resolve().parent
FONT_DIR = HERE / "assets" / "fonts"
PX_IN = 96.0
EMU_IN = 914400.0

UNICODES = ("U+0000-024F,U+0259,U+1E00-1EFF,U+02B0-02FF,U+0300-036F,"
            "U+2000-206F,U+20A0-20BF,U+2100-214F,U+2190-21BB,U+2212,U+25A0-25FF")


def subset_woff2(src, out):
    if out.exists() and out.stat().st_mtime > src.stat().st_mtime:
        return out
    subprocess.run([sys.executable, "-m", "fontTools.subset", str(src),
                    f"--output-file={out}", f"--unicodes={UNICODES}",
                    "--layout-features=*", "--flavor=woff2", "--with-zopfli"],
                   check=True, capture_output=True)
    return out


def font_css():
    faces = []
    for src, fam, weights in (("Montserrat.ttf", "Montserrat", "100 900"),
                              ("OpenSans.ttf", "Open Sans", "300 800")):
        woff = subset_woff2(FONT_DIR / src, FONT_DIR / (Path(src).stem.lower() + "-subset.woff2"))
        b64 = base64.b64encode(woff.read_bytes()).decode()
        faces.append(
            "@font-face{font-family:'%s';font-style:normal;font-weight:%s;font-display:swap;"
            "src:url(data:font/woff2;base64,%s) format('woff2');}" % (fam, weights, b64))
    return "\n".join(faces)


# ------------------------------------------------------------------ helpers
def px(emu):
    return round(float(emu) / EMU_IN * PX_IN, 2)


def hexcolor(el):
    clr = el.find(qn("a:srgbClr")) if el is not None else None
    return "#" + clr.get("val") if clr is not None else None


def shape_fill(shape):
    try:
        if shape.fill.type is not None and str(shape.fill.type).startswith("SOLID"):
            return "#" + str(shape.fill.fore_color.rgb)
    except Exception:
        pass
    return None


def shape_line(shape):
    try:
        ln = shape.line
        if ln.fill.type is not None and str(ln.fill.type).startswith("SOLID"):
            w = ln.width.pt if ln.width else 1
            return "#" + str(ln.color.rgb), round(w * PX_IN / 72, 2)
    except Exception:
        pass
    return None, 0


def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def style_box(shape):
    x, y = px(shape.left), px(shape.top)
    w, h = px(shape.width), px(shape.height)
    fill = shape_fill(shape)
    line, lw = shape_line(shape)
    geom = ""
    try:
        geom = str(shape.auto_shape_type)
    except Exception:
        geom = ""
    css = [f"left:{x}px", f"top:{y}px", f"width:{w}px", f"height:{h}px"]
    if fill:
        css.append(f"background:{fill}")
    if line:
        css.append(f"border:{lw}px solid {line}")
    if "OVAL" in geom:
        css.append("border-radius:50%")
    elif "ROUNDED" in geom:
        try:
            rad = float(shape.adjustments[0]) * min(w, h) / 2
        except Exception:
            rad = 8
        css.append(f"border-radius:{max(0, rad):.1f}px")
    return ";".join(css)


def text_html(shape):
    tf = shape.text_frame
    body = tf._txBody
    bodyPr = body.find(qn("a:bodyPr"))
    anchor = bodyPr.get("anchor", "t") if bodyPr is not None else "t"
    justify = {"ctr": "center", "b": "flex-end"}.get(anchor, "flex-start")
    pad = (px(tf.margin_left), px(tf.margin_top), px(tf.margin_right), px(tf.margin_bottom))

    paras = []
    for p in body.findall(qn("a:p")):
        pPr = p.find(qn("a:pPr"))
        align = "left"
        lh = 1.2
        before = after = 0
        if pPr is not None:
            algn = pPr.get("algn")
            align = {"ctr": "center", "r": "right", "just": "justify"}.get(algn, "left")
            ln = pPr.find(qn("a:lnSpc"))
            if ln is not None:
                pct = ln.find(qn("a:spcPct"))
                if pct is not None:
                    lh = int(pct.get("val")) / 100000
            for tag, name in ((qn("a:spcBef"), "before"), (qn("a:spcAft"), "after")):
                el = pPr.find(tag)
                if el is not None:
                    pts = el.find(qn("a:spcPts"))
                    v = int(pts.get("val")) / 100 if pts is not None else 0
                    if name == "before":
                        before = v
                    else:
                        after = v
        parts, sizes = [], []
        for child in p:
            if child.tag == qn("a:br"):
                parts.append("<br>")
                continue
            if child.tag != qn("a:r"):
                continue
            rPr = child.find(qn("a:rPr"))
            t = child.find(qn("a:t"))
            text = esc(t.text or "") if t is not None else ""
            css = []
            size = 18
            if rPr is not None:
                if rPr.get("sz"):
                    size = int(rPr.get("sz")) / 100
                sizes.append(size)
                latin = rPr.find(qn("a:latin"))
                fam = latin.get("typeface") if latin is not None else "Open Sans"
                css.append("font-family:'%s'" % fam)
                css.append(f"font-size:{size * PX_IN / 72:.2f}px")
                css.append("font-weight:700" if rPr.get("b") == "1" else "font-weight:400")
                if rPr.get("i") == "1":
                    css.append("font-style:italic")
                fill = rPr.find(qn("a:solidFill"))
                col = hexcolor(fill)
                if col:
                    css.append(f"color:{col}")
                if rPr.get("spc"):
                    css.append(f"letter-spacing:{int(rPr.get('spc'))/100 * PX_IN/72:.2f}px")
            parts.append(f"<span style=\"{';'.join(css)}\">{text}</span>")
        if not sizes:
            sizes = [18]
        # line-height theo cỡ chữ lớn nhất của đoạn
        base = max(sizes)
        paras.append(
            f"<p style=\"margin:{before * PX_IN/72:.2f}px 0 {after * PX_IN/72:.2f}px 0;"
            f"text-align:{align};line-height:{lh};min-height:{base * lh * PX_IN/72:.2f}px\">"
            + "".join(parts) + "</p>")

    return (f"<div class=\"tx\" style=\"left:{px(shape.left)}px;top:{px(shape.top)}px;"
            f"width:{px(shape.width)}px;height:{px(shape.height)}px;"
            f"padding:{pad[1]}px {pad[2]}px {pad[3]}px {pad[0]}px;"
            f"justify-content:{justify}\">{''.join(paras)}</div>")


def table_html(shape):
    tbl = shape.table
    x, y = px(shape.left), px(shape.top)
    col_w = [px(c.width) for c in tbl.columns]
    rows = []
    for r, row in enumerate(tbl.rows):
        cells = []
        for c, col in enumerate(tbl.columns):
            cell = tbl.cell(r, c)
            fill = "#" + str(cell.fill.fore_color.rgb) if cell.fill.type is not None else "transparent"
            inner = []
            for p in cell.text_frame._txBody.findall(qn("a:p")):
                runs = []
                for child in p:
                    if child.tag != qn("a:r"):
                        continue
                    rPr = child.find(qn("a:rPr"))
                    t = child.find(qn("a:t"))
                    css = []
                    if rPr is not None:
                        sz = int(rPr.get("sz", 1800)) / 100
                        latin = rPr.find(qn("a:latin"))
                        fam = latin.get("typeface") if latin is not None else "Open Sans"
                        col_ = hexcolor(rPr.find(qn("a:solidFill"))) or "#0E2A22"
                        css = [f"font-family:'{fam}'", f"font-size:{sz*PX_IN/72:.2f}px",
                               "font-weight:%s" % (700 if rPr.get("b") == "1" else 400),
                               f"color:{col_}"]
                    runs.append(f"<span style=\"{';'.join(css)}\">{esc(t.text or '')}</span>")
                inner.append("".join(runs))
            cells.append(
                f"<td style=\"background:{fill};border-bottom:1px solid #D5DBCB;"
                f"padding:{px(cell.margin_top)}px {px(cell.margin_right)}px "
                f"{px(cell.margin_bottom)}px {px(cell.margin_left)}px;"
                f"height:{px(row.height)}px;vertical-align:middle\">"
                + "<br>".join(inner) + "</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    cols = "".join(f"<col style=\"width:{w}px\">" for w in col_w)
    return (f"<table class=\"tbl\" style=\"left:{x}px;top:{y}px;border-collapse:collapse;"
            f"table-layout:fixed\"><colgroup>{cols}</colgroup>" + "".join(rows) + "</table>")


def slide_html(slide):
    try:
        bg = "#" + str(slide.background.fill.fore_color.rgb)
    except Exception:
        bg = "#F7F6EE"
    out = [f"<section class=\"slide\" style=\"background:{bg}\">"]
    for shape in slide.shapes:
        try:
            if shape.has_table:
                out.append(table_html(shape))
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                b64 = base64.b64encode(shape.image.blob).decode()
                mime = shape.image.content_type
                try:
                    prst = shape._element.spPr.find(qn("a:prstGeom")).get("prst")
                except Exception:
                    prst = "rect"
                radius = "border-radius:50%;" if prst == "ellipse" else ""
                out.append(
                    f"<img src=\"data:{mime};base64,{b64}\" style=\"position:absolute;{radius}"
                    f"left:{px(shape.left)}px;top:{px(shape.top)}px;width:{px(shape.width)}px;"
                    f"height:{px(shape.height)}px\">")
            elif shape.shape_type == MSO_SHAPE_TYPE.LINE:
                col, lw = shape_line(shape)
                x1, y1 = px(shape.left), px(shape.top)
                x2, y2 = x1 + px(shape.width), y1 + px(shape.height)
                out.append(f"<div style=\"position:absolute;left:{min(x1,x2)}px;top:{min(y1,y2)}px;"
                           f"width:{max(abs(x2-x1), 0.6)}px;height:{max(abs(y2-y1), 0.6)}px;"
                           f"background:{col or '#D5DBCB'}\"></div>")
            else:
                css = style_box(shape)
                has_text = shape.has_text_frame and shape.text_frame.text.strip()
                if has_text:
                    out.append(f"<div style=\"position:absolute;{css}\"></div>")
                    out.append(text_html(shape))
                else:
                    out.append(f"<div style=\"position:absolute;{css}\"></div>")
        except Exception as e:
            print("  ! bỏ qua shape:", e)
    out.append("</section>")
    return "\n".join(out)


CSS = """
*{box-sizing:border-box;margin:0;padding:0}
:root{--ink:#0E2A22;--cream:#F7F6EE}
html,body{height:100%;background:#0B1410;font-family:'Open Sans',system-ui,sans-serif;
 overflow:hidden;color:var(--ink)}
#stage{position:fixed;inset:0;display:flex;align-items:center;justify-content:center}
#deck{width:1280px;height:720px;position:relative;transform-origin:center center;
 box-shadow:0 40px 120px rgba(0,0,0,.55)}
.slide{position:absolute;inset:0;width:1280px;height:720px;opacity:0;pointer-events:none;
 transition:opacity .28s ease;overflow:hidden}
.slide.on{opacity:1;pointer-events:auto}
.tx{position:absolute;display:flex;flex-direction:column;white-space:normal;
 overflow-wrap:break-word}
.tbl{position:absolute;font-size:13.33px}
.ui{position:fixed;left:0;right:0;bottom:0;display:flex;align-items:center;justify-content:center;
 gap:18px;padding:14px;font-size:12px;color:#9DB3A8;letter-spacing:.08em;z-index:20;
 background:linear-gradient(to top,rgba(6,15,11,.9),transparent);font-family:'Montserrat',sans-serif}
.ui button{all:unset;cursor:pointer;padding:6px 12px;border:1px solid #2E463C;border-radius:999px;
 color:#C9D8D0;font-size:11px;letter-spacing:.1em}
.ui button:hover{border-color:#00A94F;color:#fff}
#bar{position:fixed;left:0;top:0;height:3px;background:#00A94F;width:0;z-index:30;
 transition:width .25s ease}
#notes{position:fixed;left:50%;bottom:58px;transform:translateX(-50%);width:min(880px,92vw);
 max-height:36vh;overflow:auto;background:#061F17f2;color:#D6E7DC;border:1px solid #28483A;
 border-radius:14px;padding:16px 20px;font-size:13.5px;line-height:1.6;display:none;z-index:25;
 font-family:'Open Sans',sans-serif}
#notes b{color:#D3EE72;font-family:'Montserrat';letter-spacing:.12em;font-size:10px;
 display:block;margin-bottom:6px}
body.print{overflow:visible}
body.print #deck{transform:none!important;box-shadow:none}
body.print .ui,body.print #bar,body.print #notes{display:none}
body.print #stage{position:static;display:block}
body.print{background:#fff}
body.print .slide{position:relative;opacity:1;page-break-after:always;margin:0 auto}
@media print{
 @page{size:1280px 720px;margin:0}
 body{background:#fff}
 #stage{position:static;display:block}
 #deck{transform:none!important;box-shadow:none}
 .slide{position:relative;opacity:1;page-break-after:always}
 .ui,#bar,#notes{display:none}
}
"""

JS = """
const slides=[...document.querySelectorAll('.slide')];
const notes=N => `%NOTES%`;
let i=0;
const bar=document.getElementById('bar');
const counter=document.getElementById('counter');
function show(n){
  i=Math.max(0,Math.min(slides.length-1,n));
  slides.forEach((s,k)=>s.classList.toggle('on',k===i));
  bar.style.width=((i+1)/slides.length*100)+'%';
  counter.textContent=String(i+1).padStart(2,'0')+' / '+String(slides.length).padStart(2,'0');
  const nt=document.getElementById('notes');
  if(!nt.hidden){
    const t=document.getElementById('notesBody');
    t.textContent=%NOTES_JSON%[i]||'—';
  }
  history.replaceState(null,'','#'+(i+1));
}
function fit(){
  const s=Math.min(innerWidth/1280,(innerHeight-76)/720)*0.98;
  document.getElementById('deck').style.transform=`scale(${s})`;
}
addEventListener('resize',fit);
addEventListener('keydown',e=>{
  if(['ArrowRight','ArrowDown','PageDown',' '].includes(e.key)){e.preventDefault();show(i+1);}
  if(['ArrowLeft','ArrowUp','PageUp'].includes(e.key)){e.preventDefault();show(i-1);}
  if(e.key==='Home')show(0);
  if(e.key==='End')show(slides.length-1);
  if(e.key==='n'||e.key==='N'){document.getElementById('notes').hidden=!document.getElementById('notes').hidden;show(i);}
  if(e.key==='p'||e.key==='P'){document.body.classList.toggle('print');}
});
document.getElementById('prev').onclick=()=>show(i-1);
document.getElementById('next').onclick=()=>show(i+1);
document.getElementById('grid').onclick=()=>{document.getElementById('notes').hidden=true;show(0)};
const start=parseInt(location.hash.replace('#',''))||1;
fit();show(start-1);
"""


def build(pptx_path=None, out=None):
    pptx_path = Path(pptx_path or HERE / "7UP-Brand-Deck.pptx")
    prs = Presentation(str(pptx_path))
    slides = [slide_html(s) for s in prs.slides]
    notes = []
    for s in prs.slides:
        try:
            notes.append(s.notes_slide.notes_text_frame.text.replace("\n", " ").strip())
        except Exception:
            notes.append("")
    import json
    js = JS.replace("%NOTES_JSON%", json.dumps(notes, ensure_ascii=False))
    html = f"""<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>7UP · Brand Deck</title>
<style>
{font_css()}
{CSS}
</style>
</head>
<body>
<div id="bar"></div>
<div id="stage"><div id="deck">
{''.join(slides)}
</div></div>
<div id="notes" hidden><b>GHI CHÚ THUYẾT TRÌNH · NHẤN N ĐỂ ẨN</b><div id="notesBody"></div></div>
<div class="ui">
  <button id="prev">← TRƯỚC</button>
  <span id="counter">01 / 12</span>
  <button id="next">SAU →</button>
  <span style="opacity:.55">← → điều hướng · N ghi chú · P chế độ in</span>
</div>
<script>
{js}
</script>
</body>
</html>
"""
    out = Path(out or HERE / "7up-slides.html")
    out.write_text(html, encoding="utf-8")
    print(f"Đã lưu: {out} ({out.stat().st_size/1024:.0f} KB)")
    return out


if __name__ == "__main__":
    build()
