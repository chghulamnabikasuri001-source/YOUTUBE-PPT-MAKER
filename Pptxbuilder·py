"""
Builds rich, designed .pptx files from structured slide data.

Every slide can carry a "visual": nested circles, flow chart, cards, numbered steps,
table, tree, neural-network diagram, code block, comparison panels or a screenshot.
All shapes are native PowerPoint objects, so the teacher can edit everything.
"""
import io
import math
import re

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

W, H = 13.333, 7.5

# acc = [teal-ish, orange-ish, blue-ish, navy-ish]  (used in different orders by each visual)
THEMES = {
    "Classroom": dict(dark="111A2E", light="F6F3EC", surface_dark="1F2D4E", surface_light="FFFFFF",
                      ink="1B2540", on_dark="F6F3EC", muted="5B6478",
                      acc=["1FB5A3", "F28C3E", "2A7090", "1F2D4E"], head="Trebuchet MS", body="Calibri"),
    "Ocean Blue": dict(dark="0B2A4A", light="F0F7FF", surface_dark="16406B", surface_light="FFFFFF",
                       ink="0F2A47", on_dark="F0F7FF", muted="53708F",
                       acc=["2E9BE8", "FFB142", "1B6FB5", "123B66"], head="Georgia", body="Calibri"),
    "Forest": dict(dark="12372A", light="F4F8F1", surface_dark="1E5240", surface_light="FFFFFF",
                   ink="14301F", on_dark="F4F8F1", muted="5A7266",
                   acc=["2E9E6B", "F2A541", "4C8C7A", "1F5C45"], head="Georgia", body="Trebuchet MS"),
    "Sunset": dict(dark="2B1B3D", light="FFF6F0", surface_dark="3E2A59", surface_light="FFFFFF",
                   ink="2B1B3D", on_dark="FFF6F0", muted="7A6A8A",
                   acc=["FF6B6B", "FFA94D", "845EF7", "3B2A5E"], head="Trebuchet MS", body="Verdana"),
    "Graphite": dict(dark="1F2328", light="F5F5F4", surface_dark="2D333B", surface_light="FFFFFF",
                     ink="1F2328", on_dark="F5F5F4", muted="64707D",
                     acc=["10B981", "F59E0B", "3B82F6", "374151"], head="Calibri", body="Calibri"),
}
DEFAULT_THEME = "Classroom"

CHAR_W = {"Calibri": 0.49, "Trebuchet MS": 0.54, "Georgia": 0.55, "Verdana": 0.60, "Arial": 0.52}


# ----------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------
def _rgb(h):
    return RGBColor.from_string(h)


def _lum(h):
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return 0.299 * r + 0.587 * g + 0.114 * b


def _tint(h, t):
    """Mix a colour with white. t=0 -> original, t=1 -> white."""
    rgb = [int(h[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{int(c + (255 - c) * t):02X}" for c in rgb)


class Ctx:
    def __init__(self, T, dark, rtl):
        self.T, self.dark, self.rtl = T, dark, rtl

    @property
    def text(self):
        return self.T["on_dark"] if self.dark else self.T["ink"]

    @property
    def surface(self):
        return self.T["surface_dark"] if self.dark else self.T["surface_light"]

    @property
    def line(self):
        return "3E4F78" if self.dark else "D5DDE8"

    @property
    def bg(self):
        return self.T["dark"] if self.dark else self.T["light"]

    @property
    def head_font(self):
        return "Arial" if self.rtl else self.T["head"]

    @property
    def body_font(self):
        return "Arial" if self.rtl else self.T["body"]

    def on(self, fill):
        return self.T["ink"] if _lum(fill) > 0.55 else self.T["on_dark"]

    def pal(self):  # order used by diagrams (navy, blue, teal, orange)
        a = self.T["acc"]
        return [a[3], a[2], a[0], a[1]]

    def cards(self):  # order used by cards (teal, orange, blue, navy)
        a = self.T["acc"]
        return [a[0], a[1], a[2], a[3]]


def _est_lines(text, size, width, font="Calibri", bold=False, mono=False):
    cw = size * (0.6 if mono else CHAR_W.get(font, 0.52) * (1.07 if bold else 1.0)) / 72
    cpl = max(1, int(width / cw))
    lines = 0
    for para in str(text).split("\n"):
        cur, n = 0, 1
        for wd in para.split():
            L = len(wd) + (1 if cur else 0)
            if cur + L > cpl and cur > 0:
                n, cur = n + 1, len(wd)
            else:
                cur += L
        lines += n
    return lines


def _fit(items, width, height, max_size, min_size, font="Calibri", bold=False, gap=0.12, indent=0.0):
    """Biggest font size (pt) at which all items fit inside width x height (inches)."""
    items = [re.sub(r"\*\*", "", str(i)) for i in items] or [""]
    cw_em = CHAR_W.get(font, 0.52) * (1.07 if bold else 1.0) * 1.12
    longest = max((len(w) for t in items for w in t.split()), default=1)
    max_size = min(max_size, max(min_size, int((width - indent) * 72 / (longest * cw_em))))
    s = max_size
    while s > min_size:
        h = sum(_est_lines(t, s, width - indent, font, bold) * s * 1.2 / 72 + gap for t in items)
        if h <= height:
            return s
        s -= 1
    return min_size


def _add_runs(p, text, size, color, font, bold=False):
    for part in re.split(r"(\*\*.+?\*\*)", str(text)):
        if not part:
            continue
        b = bold
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            part, b = part[2:-2], True
        r = p.add_run()
        r.text = part
        r.font.size = Pt(size)
        r.font.bold = b
        r.font.name = font
        r.font.color.rgb = _rgb(color)


def _bulletize(p):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(0.32))))
    pPr.set("indent", str(-int(Inches(0.32))))
    bu = etree.SubElement(pPr, qn("a:buChar"))
    bu.set("char", "•")


def _fill_tf(tf, ctx, paras, anchor=MSO_ANCHOR.TOP, margins=(0.05, 0.03)):
    """paras: list of dicts {t,size,color,bold,head,font,align,space,bullet}"""
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margins[0])
    tf.margin_top = tf.margin_bottom = Inches(margins[1])
    for i, d in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        font = d.get("font") or (ctx.head_font if d.get("head") else ctx.body_font)
        text = d["t"]
        if d.get("bullet") and ctx.rtl:
            text = "•  " + text
        _add_runs(p, text, d["size"], d["color"], font, d.get("bold", False))
        al = d.get("align")
        if ctx.rtl:
            p._p.get_or_add_pPr().set("rtl", "1")
            p.alignment = al if al is not None else PP_ALIGN.RIGHT
        elif al is not None:
            p.alignment = al
        if d.get("space") is not None:
            p.space_after = Pt(d["space"])
        if d.get("bullet") and not ctx.rtl:
            _bulletize(p)


def add_text(slide, ctx, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    _fill_tf(tb.text_frame, ctx, paras, anchor)
    return tb


def add_shape(slide, shape, x, y, w, h, fill=None, line=None, line_w=1.0, radius=None):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = _rgb(line)
        shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = radius
    shp.shadow.inherit = False
    st = shp._element.find(qn('p:style'))
    if st is not None:
        shp._element.remove(st)
    return shp


def shape_text(shp, ctx, paras, anchor=MSO_ANCHOR.MIDDLE, margins=(0.1, 0.05)):
    _fill_tf(shp.text_frame, ctx, paras, anchor, margins)


def add_line(slide, x1, y1, x2, y2, color, width=1.5):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = _rgb(color)
    c.line.width = Pt(width)
    return c


def _label(x):
    if isinstance(x, dict):
        return str(x.get("label") or x.get("title") or x.get("text") or x.get("name") or "").strip()
    return str(x).strip()


# ----------------------------------------------------------------------
# Visual normalisation (the AI output is never trusted blindly)
# ----------------------------------------------------------------------
def _strs(lst, n):
    out = [_label(x) for x in (lst or [])] if isinstance(lst, list) else []
    return [s for s in out if s][:n]


def _norm_tree(node, depth=0):
    if isinstance(node, str):
        node = {"label": node}
    if not isinstance(node, dict):
        return None
    lab = _label(node)
    if not lab:
        return None
    kids = []
    if depth < 2:
        for c in (node.get("children") or [])[:5]:
            k = _norm_tree(c, depth + 1)
            if k:
                kids.append(k)
    return {"label": lab, "children": kids}


def _norm_items(items, n):
    out = []
    for it in (items or [])[:n]:
        if isinstance(it, dict):
            title = str(it.get("title") or it.get("label") or it.get("name") or "").strip()
            text = str(it.get("text") or it.get("description") or it.get("subtitle") or "").strip()
            icon = str(it.get("icon") or it.get("emoji") or "").strip()
        else:
            title, text, icon = str(it).strip(), "", ""
        if title or text:
            out.append({"icon": icon[:4], "title": title, "text": text})
    return out


def norm_visual(v):
    if not isinstance(v, dict):
        return None
    t = str(v.get("type", "")).lower().strip()
    if t in ("nested", "venn", "circles", "concentric"):
        items = _strs(v.get("items"), 4)
        return {"type": "nested", "items": items} if len(items) >= 2 else None
    if t in ("flow", "process", "flowchart", "pipeline"):
        items = _strs(v.get("items"), 6)
        d = "horizontal" if str(v.get("direction", "")).lower().startswith("h") else "vertical"
        return {"type": "flow", "items": items, "direction": d} if len(items) >= 2 else None
    if t in ("cards", "steps"):
        items = _norm_items(v.get("items"), 4 if t == "cards" else 5)
        return {"type": t, "items": items} if len(items) >= 2 else None
    if t == "table":
        headers = [_label(x) for x in (v.get("headers") or [])][:4] if isinstance(v.get("headers"), list) else []
        if sum(1 for h in headers if h) < 1:
            headers = []
        rows = []
        for r in (v.get("rows") or [])[:7]:
            if isinstance(r, list):
                cells = [str(c).strip() for c in r][: len(headers)]
                cells += [""] * (len(headers) - len(cells))
                rows.append(cells)
        return {"type": "table", "headers": headers, "rows": rows} if headers and rows else None
    if t in ("tree", "hierarchy", "mindmap"):
        root = _norm_tree(v.get("root"))
        return {"type": "tree", "root": root} if root and root["children"] else None
    if t in ("network", "neural_network", "nn"):
        layers = []
        for x in (v.get("layers") or [])[:6]:
            try:
                layers.append(max(1, min(8, int(x))))
            except Exception:
                pass
        labels = _strs(v.get("labels"), 3)
        return {"type": "network", "layers": layers, "labels": labels} if len(layers) >= 3 else None
    if t == "code":
        code = str(v.get("code", "")).rstrip()
        return {"type": "code", "code": code, "language": str(v.get("language", ""))[:14]} if code else None
    if t in ("compare", "comparison", "versus"):
        sides = []
        for k in ("left", "right"):
            d = v.get(k)
            if isinstance(d, dict) and (d.get("title") or d.get("points")):
                sides.append({"title": str(d.get("title", "")).strip(), "points": _strs(d.get("points"), 5)})
        return {"type": "compare", "left": sides[0], "right": sides[1]} if len(sides) == 2 else None
    return None


# ----------------------------------------------------------------------
# Visual renderers.  Each draws inside the box (x, y, w, h) in inches.
# ----------------------------------------------------------------------
def draw_nested(slide, ctx, items, x, y, w, h):
    n = len(items)
    D = min(w, h)
    cx = x + w / 2
    bottom = y + (h + D) / 2
    ratios = {2: [1, .62], 3: [1, .68, .36], 4: [1, .76, .52, .28]}[n]
    cols = ctx.pal()
    diam = [D * r for r in ratios]
    for i, d in enumerate(diam):
        shp = add_shape(slide, MSO_SHAPE.OVAL, cx - d / 2, bottom - d, d, d, fill=cols[i],
                        line=ctx.line if (i == 0 and ctx.dark) else None, line_w=2)
    for i, d in enumerate(diam):
        fill = cols[i]
        if i < n - 1:
            band = d - diam[i + 1]
            ty, th = bottom - d + 0.04, max(band * 0.92, 0.3)
            anchor = MSO_ANCHOR.MIDDLE
        else:
            ty, th, anchor = bottom - d, d, MSO_ANCHOR.MIDDLE
        tw = d * (0.82 if i < n - 1 else 0.78)
        size = _fit([items[i]], tw, th * 0.9, 24, 10, ctx.head_font, True)
        add_text(slide, ctx, cx - tw / 2, ty, tw, th,
                 [dict(t=items[i], size=size, color=ctx.on(fill), bold=True, head=True, align=PP_ALIGN.CENTER)],
                 anchor)


def draw_flow(slide, ctx, items, x, y, w, h, direction="vertical"):
    n = len(items)
    cols = ctx.pal()
    if direction == "horizontal":
        gap = 0.55
        bw = min((w - (n - 1) * gap) / n, 2.7)
        bh = min(h, 1.6)
        total = n * bw + (n - 1) * gap
        x0, y0 = x + (w - total) / 2, y + (h - bh) / 2
        for i, t in enumerate(items):
            bx = x0 + i * (bw + gap)
            shp = add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, bx, y0, bw, bh, fill=cols[i % 4], radius=0.14)
            size = _fit([t], bw - 0.2, bh - 0.15, 20, 10, ctx.head_font, True)
            shape_text(shp, ctx, [dict(t=t, size=size, color=ctx.on(cols[i % 4]), bold=True, head=True,
                                       align=PP_ALIGN.CENTER)])
            if i < n - 1:
                add_shape(slide, MSO_SHAPE.RIGHT_ARROW, bx + bw + 0.08, y0 + bh / 2 - 0.17, gap - 0.16, 0.34,
                          fill=ctx.text)
    else:
        gap = 0.42
        bw = min(w, 5.2)
        bh = min((h - (n - 1) * gap) / n, 1.05)
        total = n * bh + (n - 1) * gap
        x0, y0 = x + (w - bw) / 2, y + (h - total) / 2
        for i, t in enumerate(items):
            by = y0 + i * (bh + gap)
            shp = add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x0, by, bw, bh, fill=cols[i % 4], radius=0.18)
            size = _fit([t], bw - 0.3, bh - 0.12, 22, 11, ctx.head_font, True)
            shape_text(shp, ctx, [dict(t=t, size=size, color=ctx.on(cols[i % 4]), bold=True, head=True,
                                       align=PP_ALIGN.CENTER)])
            if i < n - 1:
                add_shape(slide, MSO_SHAPE.DOWN_ARROW, x0 + bw / 2 - 0.16, by + bh + 0.05, 0.32, gap - 0.1,
                          fill=ctx.text)


def draw_cards(slide, ctx, items, x, y, w, h, numbered=False):
    n = len(items)
    if numbered:
        fills = [_tint(c, 0.72) for c in ctx.cards()]
        txt = ctx.T["ink"]
        badge_cols = ctx.cards()
    else:
        fills = ctx.cards()
        txt = None
    stacked = w < 8
    if not stacked:
        gap = 0.3
        cw = (w - (n - 1) * gap) / n
        ch = min(h, 2.9 if numbered else 3.7)
        cy = y + (h - ch) / 2
        for i, it in enumerate(items):
            cx = x + i * (cw + gap)
            fill = fills[i % 4]
            tc = txt or ctx.on(fill)
            add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, cw, ch, fill=fill, radius=0.08)
            top = cy + 0.25
            if numbered:
                b = add_shape(slide, MSO_SHAPE.OVAL, cx + 0.25, top, 0.55, 0.55, fill=badge_cols[i % 4])
                shape_text(b, ctx, [dict(t=str(i + 1), size=18, color=ctx.on(badge_cols[i % 4]), bold=True,
                                         head=True, align=PP_ALIGN.CENTER)], margins=(0, 0))
                top += 0.75
            elif it["icon"]:
                add_text(slide, ctx, cx + 0.2, top - 0.05, cw - 0.4, 0.85,
                         [dict(t=it["icon"], size=36, color=tc, align=PP_ALIGN.LEFT)])
                top += 0.95
            tw = cw - 0.5
            avail = cy + ch - top - 0.2
            tsize = _fit([it["title"]], tw, avail * 0.45, 26, 13, ctx.head_font, True)
            paras = [dict(t=it["title"], size=tsize, color=tc, bold=True, head=True, space=6)]
            if it["text"]:
                th = avail - _est_lines(it["title"], tsize, tw, ctx.head_font, True) * tsize * 1.2 / 72
                ssize = _fit([it["text"]], tw, max(th, 0.6), 20, 11, ctx.body_font)
                paras.append(dict(t=it["text"], size=ssize, color=tc))
            add_text(slide, ctx, cx + 0.25, top, tw + 0.05, avail, paras)
    else:
        gap = 0.22
        rh = min((h - (n - 1) * gap) / n, 1.5)
        total = n * rh + (n - 1) * gap
        y0 = y + (h - total) / 2
        for i, it in enumerate(items):
            ry = y0 + i * (rh + gap)
            fill = fills[i % 4]
            tc = txt or ctx.on(fill)
            add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, ry, w, rh, fill=fill, radius=0.14)
            lead_w = 0.9
            if numbered:
                b = add_shape(slide, MSO_SHAPE.OVAL, x + 0.2, ry + rh / 2 - 0.3, 0.6, 0.6, fill=badge_cols[i % 4])
                shape_text(b, ctx, [dict(t=str(i + 1), size=18, color=ctx.on(badge_cols[i % 4]), bold=True,
                                         head=True, align=PP_ALIGN.CENTER)], margins=(0, 0))
            elif it["icon"]:
                add_text(slide, ctx, x + 0.1, ry, lead_w, rh,
                         [dict(t=it["icon"], size=30, color=tc, align=PP_ALIGN.CENTER)], MSO_ANCHOR.MIDDLE)
            else:
                lead_w = 0.25
            tw = w - lead_w - 0.3
            tsize = _fit([it["title"]], tw, rh * 0.5, 20, 12, ctx.head_font, True)
            paras = [dict(t=it["title"], size=tsize, color=tc, bold=True, head=True, space=2)]
            if it["text"]:
                ssize = _fit([it["text"]], tw, rh * 0.5, 16, 10, ctx.body_font)
                paras.append(dict(t=it["text"], size=ssize, color=tc))
            add_text(slide, ctx, x + lead_w + 0.1, ry, tw, rh, paras, MSO_ANCHOR.MIDDLE)


def draw_table(slide, ctx, data, x, y, w, h):
    headers, rows = data["headers"], data["rows"]
    nc, nr = len(headers), len(rows) + 1
    rh = min(0.8, h / nr)
    wts = []
    for c in range(nc):
        L = max([len(headers[c])] + [len(r[c]) for r in rows])
        wts.append(min(max(L, 8), 30))
    tot = sum(wts)
    widths = [w * wt / tot for wt in wts]
    gs = slide.shapes.add_table(nr, nc, Inches(x), Inches(y + (h - rh * nr) / 2 if rh * nr < h else y),
                                Inches(w), Inches(rh * nr))
    tbl = gs.table
    for c in range(nc):
        tbl.columns[c].width = Inches(widths[c])
    for r in range(nr):
        tbl.rows[r].height = Inches(rh)
    longest = max(len(c) for r in rows for c in r) if rows else 10
    size = 20 if longest < 22 and nr <= 5 else 17 if longest < 36 else 14
    header_fill = ctx.T["acc"][3]
    if ctx.dark:
        header_fill = ctx.T["acc"][2]
    for r in range(nr):
        for c in range(nc):
            cell = tbl.cell(r, c)
            text = headers[c] if r == 0 else rows[r - 1][c]
            if r == 0:
                fill, color, bold = header_fill, ctx.on(header_fill), True
            else:
                base = ctx.surface
                fill = base if r % 2 else (_tint(base, 0.0) if ctx.dark else "F1F5F9")
                color, bold = ctx.text, (c == 0)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(fill)
            cell.margin_left = cell.margin_right = Inches(0.18)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            _add_runs(p, text, size, color, ctx.head_font if r == 0 else ctx.body_font, bold)
            if ctx.rtl:
                p._p.get_or_add_pPr().set("rtl", "1")
                p.alignment = PP_ALIGN.RIGHT


def draw_tree(slide, ctx, root, x, y, w, h):
    # count leaves & depth
    def leaves(n):
        return 1 if not n["children"] else sum(leaves(c) for c in n["children"])

    def depth(n):
        return 1 if not n["children"] else 1 + max(depth(c) for c in n["children"])

    L = depth(root)
    nleaf = leaves(root)
    slot = w / nleaf
    nw = max(min(2.4, slot * 0.92), 0.95)
    nh = 0.8 if nleaf <= 6 else 0.7
    gap = 0.9 if L == 2 else 0.6
    gap = min(gap, (h - L * nh) / max(L - 1, 1))
    y0 = y + (h - (L * nh + (L - 1) * gap)) / 2
    cols = ctx.pal()
    counter = [0]
    placed = []  # (node, cx, level)

    def place(n, lvl):
        if not n["children"]:
            cx = x + slot * (counter[0] + 0.5)
            counter[0] += 1
        else:
            xs = [place(c, lvl + 1) for c in n["children"]]
            cx = (xs[0] + xs[-1]) / 2
        placed.append((n, cx, lvl))
        return cx

    place(root, 0)
    pos = {id(n): (cx, lvl) for n, cx, lvl in placed}
    line_col = ctx.T["muted"] if not ctx.dark else "8FA0C8"

    def links(n):
        if not n["children"]:
            return
        px, pl = pos[id(n)]
        py = y0 + pl * (nh + gap) + nh
        mid = py + gap / 2
        add_line(slide, px, py, px, mid, line_col, 1.75)
        cxs = [pos[id(c)][0] for c in n["children"]]
        if len(cxs) > 1:
            add_line(slide, min(cxs), mid, max(cxs), mid, line_col, 1.75)
        for c in n["children"]:
            cx, cl = pos[id(c)]
            add_line(slide, cx, mid, cx, y0 + cl * (nh + gap), line_col, 1.75)
            links(c)

    links(root)
    for n, cx, lvl in placed:
        fill = [cols[0], cols[2], cols[3], cols[1]][min(lvl, 3)] if lvl else cols[0]
        fill = {0: cols[0], 1: ctx.cards()[0], 2: cols[1]}[min(lvl, 2)]
        shp = add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, cx - nw / 2, y0 + lvl * (nh + gap), nw, nh,
                        fill=fill, radius=0.2)
        size = _fit([n["label"]], nw - 0.15, nh - 0.1, 17, 9, ctx.head_font, True)
        shape_text(shp, ctx, [dict(t=n["label"], size=size, color=ctx.on(fill), bold=True, head=True,
                                   align=PP_ALIGN.CENTER)], margins=(0.05, 0.03))


def draw_network(slide, ctx, layers, labels, x, y, w, h):
    n = len(layers)
    lab_h = 0.5
    area_h = h - lab_h - 0.1
    mx = max(layers)
    d = min(0.42, area_h / (mx * 1.55))
    xs = [x + d / 2 + i * (w - d) / (n - 1) for i in range(n)]
    sp = min(d * 1.75, (area_h - d) / max(mx - 1, 1))

    def ys(m):
        tot = d + (m - 1) * sp
        top = y + (area_h - tot) / 2
        return [top + d / 2 + k * sp for k in range(m)]

    all_y = [ys(m) for m in layers]
    line_col = "4A5A8A" if ctx.dark else "9FB0C8"
    for i in range(n - 1):
        for ya in all_y[i]:
            for yb in all_y[i + 1]:
                add_line(slide, xs[i], ya, xs[i + 1], yb, line_col, 0.75)
    a = ctx.T["acc"]
    out_fill = ctx.T["on_dark"] if ctx.dark else ctx.T["ink"]
    for i in range(n):
        fill = a[1] if i == 0 else (out_fill if i == n - 1 else a[0])
        for yy in all_y[i]:
            add_shape(slide, MSO_SHAPE.OVAL, xs[i] - d / 2, yy - d / 2, d, d, fill=fill)
    labs = labels if len(labels) == 3 else ["Input", "Hidden layers", "Output"]
    ly = y + h - lab_h
    mid_x = (xs[1] + xs[-2]) / 2 if n > 3 else xs[1]
    for txt, cx in zip(labs, (xs[0], mid_x, xs[-1])):
        add_text(slide, ctx, cx - 1.1, ly, 2.2, lab_h,
                 [dict(t=txt, size=15, color=ctx.text, align=PP_ALIGN.CENTER)], MSO_ANCHOR.MIDDLE)


def draw_code(slide, ctx, data, x, y, w, h):
    code = data["code"].replace("\t", "    ")
    lines = code.split("\n")[:16]
    bar = 0.5
    aw = w - 0.5
    longest = max(len(l) for l in lines)
    size = min(20, aw * 72 / (max(longest, 1) * 0.6), (h - bar - 0.4) * 72 / (len(lines) * 1.3))
    size = max(int(size), 10)
    maxc = int(aw * 72 / (size * 0.6))
    lines = [l if len(l) <= maxc else l[: maxc - 1] + "…" for l in lines]
    box_h = min(h, bar + len(lines) * size * 1.35 / 72 + 0.45)
    y = y + (h - box_h) / 2
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, box_h, fill="0F172A", radius=0.04)
    for i, c in enumerate(("FF5F56", "FFBD2E", "27C93F")):
        add_shape(slide, MSO_SHAPE.OVAL, x + 0.22 + i * 0.26, y + 0.18, 0.14, 0.14, fill=c)
    if data.get("language"):
        add_text(slide, ctx, x + w - 1.6, y + 0.1, 1.45, 0.3,
                 [dict(t=data["language"], size=11, color="94A3B8", font="Consolas", align=PP_ALIGN.RIGHT)])
    add_text(slide, ctx, x + 0.25, y + bar + 0.05, aw, box_h - bar - 0.1,
             [dict(t="\n".join(lines), size=size, color="E2E8F0", font="Consolas")])


def draw_compare(slide, ctx, data, x, y, w, h):
    gap = 0.4
    pw = (w - gap) / 2
    fills = [ctx.cards()[0], ctx.cards()[1]]
    npts = max(len(data["left"]["points"]), len(data["right"]["points"]), 1)
    allpts = data["left"]["points"] + data["right"]["points"]
    bsize = _fit(allpts, pw - 0.6, h - 1.5, 22, 12, ctx.body_font, gap=0.18, indent=0.32)
    ph = min(h, 1.35 + npts * (bsize * 1.45 / 72 + 0.16) + 0.25)
    py = y + (h - ph) / 2
    for i, side in enumerate((data["left"], data["right"])):
        px = x + i * (pw + gap)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, px, py, pw, ph, fill=ctx.surface, line=ctx.line, radius=0.05)
        hd = add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, px, py, pw, 0.95, fill=fills[i], radius=0.2)
        size = _fit([side["title"]], pw - 0.4, 0.8, 28, 14, ctx.head_font, True)
        shape_text(hd, ctx, [dict(t=side["title"], size=size, color=ctx.on(fills[i]), bold=True, head=True,
                                  align=PP_ALIGN.CENTER)])
        add_text(slide, ctx, px + 0.3, py + 1.15, pw - 0.6, ph - 1.25,
                 [dict(t=pt, size=bsize, color=ctx.text, bullet=True, space=8) for pt in side["points"]])


def draw_image(slide, ctx, img_bytes, x, y, w, h):
    iw, ih = Image.open(io.BytesIO(img_bytes)).size
    scale = min(w / iw, h / ih)
    pw, ph = iw * scale, ih * scale
    px, py = x + (w - pw) / 2, y + (h - ph) / 2
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, px - 0.08, py - 0.08, pw + 0.16, ph + 0.16,
              fill=ctx.surface, line=ctx.line, radius=0.02)
    slide.shapes.add_picture(io.BytesIO(img_bytes), Inches(px), Inches(py), Inches(pw), Inches(ph))


def draw_rows(slide, ctx, items, x, y, w, h, numbered=True):
    """items: list of (main, sub). Used for takeaways, quiz, answers, and plain bullet slides."""
    n = len(items)
    gap = 0.2
    rh = min((h - (n - 1) * gap) / n, 1.15)
    y0 = y + (h - (n * rh + (n - 1) * gap)) / 2
    cols = ctx.cards()
    for i, (main, sub) in enumerate(items):
        ry = y0 + i * (rh + gap)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, ry, w, rh, fill=ctx.surface, line=ctx.line, radius=0.2)
        c = cols[i % 4]
        bx = x + w - 0.76 if ctx.rtl else x + 0.22
        b = add_shape(slide, MSO_SHAPE.OVAL, bx, ry + rh / 2 - 0.27, 0.54, 0.54, fill=c)
        if numbered:
            shape_text(b, ctx, [dict(t=str(i + 1), size=16, color=ctx.on(c), bold=True, head=True,
                                     align=PP_ALIGN.CENTER)], margins=(0, 0))
        tw = w - 1.2
        both = [main] + ([sub] if sub else [])
        size = _fit(both, tw, rh - 0.1, 26, 12, ctx.body_font, gap=0.05)
        paras = [dict(t=main, size=size, color=ctx.text, bold=bool(sub), space=2)]
        if sub:
            paras.append(dict(t=sub, size=max(size - 1, 10), color=ctx.T["acc"][2] if not ctx.dark else ctx.T["acc"][0]))
        add_text(slide, ctx, x + (0.25 if ctx.rtl else 0.95), ry, tw, rh, paras, MSO_ANCHOR.MIDDLE)


# ----------------------------------------------------------------------
# Slide assembly
# ----------------------------------------------------------------------
FULL_TYPES = {"table", "compare", "tree"}


def _new_slide(prs, ctx):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = _rgb(ctx.bg)
    return s


def _mx(ctx, x, w):
    """Mirror x for right-to-left languages."""
    return W - x - w if ctx.rtl else x


def _draw_visual(slide, ctx, vis, img, x, y, w, h):
    t = vis["type"] if vis else "screenshot"
    if t == "nested":
        draw_nested(slide, ctx, vis["items"], x, y, w, h)
    elif t == "flow":
        draw_flow(slide, ctx, vis["items"], x, y, w, h, vis["direction"] if w > 8 else "vertical")
    elif t == "cards":
        draw_cards(slide, ctx, vis["items"], x, y, w, h)
    elif t == "steps":
        draw_cards(slide, ctx, vis["items"], x, y, w, h, numbered=True)
    elif t == "table":
        draw_table(slide, ctx, vis, x, y, w, h)
    elif t == "tree":
        draw_tree(slide, ctx, vis["root"], x, y, w, h)
    elif t == "network":
        draw_network(slide, ctx, vis["layers"], vis["labels"], x, y, w, h)
    elif t == "code":
        draw_code(slide, ctx, vis, x, y, w, h)
    elif t == "compare":
        draw_compare(slide, ctx, vis, x, y, w, h)
    elif t == "screenshot":
        draw_image(slide, ctx, img, x, y, w, h)


def _footer(slide, ctx, text):
    if not text:
        return
    col = ctx.T["acc"][0] if ctx.dark else ctx.T["acc"][2]
    size = _fit([text], 11.9, 0.7, 22, 13, ctx.head_font, True)
    add_text(slide, ctx, 0.7, 6.45, 11.93, 0.75,
             [dict(t=text, size=size, color=col, bold=True, head=True)], MSO_ANCHOR.MIDDLE)


def _page_no(slide, ctx, i, n):
    add_text(slide, ctx, W - 1.6 if not ctx.rtl else 0.2, 7.05, 1.4, 0.3,
             [dict(t=f"{i} / {n}", size=11, color=ctx.T["muted"], align=PP_ALIGN.RIGHT if not ctx.rtl else PP_ALIGN.LEFT)])


def _heading_size(text, width, max_size=38):
    for s in range(max_size, 21, -2):
        if _est_lines(text, s, width, "Trebuchet MS", True) <= 2:
            return s
    return 22


def build_content_slide(prs, ctx, s, idx, total):
    heading = s["heading"]
    bullets = s["bullets"]
    vis = s["visual"]
    img = s.get("image")
    footer = s["footer"]
    slide = _new_slide(prs, ctx)

    vtype = vis["type"] if vis else ("screenshot" if img else None)
    if vtype is None:
        # plain bullets -> designed rows
        hs = _heading_size(heading, 11.9)
        add_text(slide, ctx, 0.7, 0.45, 11.93, 1.2, [dict(t=heading, size=hs, color=ctx.text, bold=True, head=True)],
                 MSO_ANCHOR.MIDDLE)
        items = [(b, None) for b in bullets] or [(heading, None)]
        draw_rows(slide, ctx, items[:5], 0.7, 1.85, 11.93, 4.5, numbered=False)
    else:
        full = vtype in FULL_TYPES or not bullets
        if vtype in FULL_TYPES and bullets and not footer:
            footer = " • ".join(bullets[:2])
        if full:
            hs = _heading_size(heading, 11.9)
            add_text(slide, ctx, 0.7, 0.45, 11.93, 1.2,
                     [dict(t=heading, size=hs, color=ctx.text, bold=True, head=True)], MSO_ANCHOR.MIDDLE)
            vh = 4.4 if footer else 5.1
            _draw_visual(slide, ctx, vis, img, 0.7, 1.8, 11.93, vh)
        else:
            col_w, vis_w = 5.7, 5.9
            tx, vx = _mx(ctx, 0.7, col_w), _mx(ctx, 6.75, vis_w)
            hs = max(_heading_size(heading, col_w, 34) - 4, 24)
            head_h = _est_lines(heading, hs, col_w, ctx.head_font, True) * hs * 1.2 / 72 + 0.2
            avail = 5.2 - head_h
            bs = _fit(bullets, col_w - 0.1, avail, 24, 14, ctx.body_font, gap=0.14, indent=0.32)
            paras = [dict(t=heading, size=hs, color=ctx.text, bold=True, head=True, space=14)]
            paras += [dict(t=b, size=bs, color=ctx.text, bullet=True, space=9) for b in bullets]
            add_text(slide, ctx, tx, 0.8, col_w, 5.5, paras, MSO_ANCHOR.MIDDLE)
            vh = 4.9 if footer else 5.6
            _draw_visual(slide, ctx, vis, img, vx, 0.95 if not footer else 1.0, vis_w, vh)
        _footer(slide, ctx, footer)
    _page_no(slide, ctx, idx, total)
    if s.get("notes"):
        slide.notes_slide.notes_text_frame.text = s["notes"]
    return slide


def _delete_slide(prs, slide):
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        if prs.part.related_part(sid.rId) is slide.part:
            prs.part.drop_rel(sid.rId)
            lst.remove(sid)


def build_title_slide(prs, ctx, title, tagline, credit, cover):
    slide = _new_slide(prs, ctx)
    lw = 6.9
    lx = _mx(ctx, 0.8, lw)
    ts = 54 if len(title) <= 28 else 46 if len(title) <= 48 else 38
    paras = []
    if tagline:
        paras.append(dict(t=tagline, size=22, color=ctx.T["acc"][0], bold=True, head=True, space=10))
    paras.append(dict(t=title, size=ts, color=ctx.text, bold=True, head=True))
    add_text(slide, ctx, lx, 1.3, lw, 4.6, paras, MSO_ANCHOR.MIDDLE)
    if credit:
        add_text(slide, ctx, lx, 6.4, lw, 0.5, [dict(t=credit, size=14, color=ctx.T["muted"] if not ctx.dark else "9AA7C0")])
    vx = _mx(ctx, 7.7, 5.0)
    try:
        if cover:
            _draw_visual(slide, ctx, cover, None, vx, 1.0, 5.0, 5.5)
        else:
            raise ValueError
    except Exception:
        draw_nested(slide, ctx, ["", "", ""], vx, 1.0, 5.0, 5.5)
    return slide


def build_pptx(title, tagline, channel, video_url, slides, takeaways, quiz, rtl,
               theme_name=DEFAULT_THEME, cover_visual=None) -> bytes:
    """
    slides: [{"heading","bullets","visual","footer","notes","image"(bytes, optional)}]
    takeaways: [str]   quiz: [{"q","a"}]
    """
    T = THEMES.get(theme_name, THEMES[DEFAULT_THEME])
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)

    credit = (f"ماخذ: {channel}" if rtl else f"Based on a lecture by {channel}") if channel else ""
    build_title_slide(prs, Ctx(T, True, rtl), title, tagline, credit, norm_visual(cover_visual))

    # normalise slides and expand screenshot+visual slides into two
    flat = []
    for s in slides:
        ns = dict(
            heading=str(s.get("heading", "")).strip() or "Untitled",
            bullets=_strs(s.get("bullets"), 5),
            visual=norm_visual(s.get("visual")),
            footer=str(s.get("footer") or "").strip(),
            notes=str(s.get("notes") or "").strip(),
            image=s.get("image"),
        )
        if ns["image"] and ns["visual"]:
            extra = dict(ns, visual=None, bullets=[], footer="", notes="",
                         heading=("From the lecture: " if not rtl else "") + ns["heading"])
            ns["image"] = None
            flat.append(ns)
            flat.append(extra)
        else:
            flat.append(ns)

    total = len(flat)
    for i, ns in enumerate(flat, 1):
        dark = (i % 3 == 0)
        ctx = Ctx(T, dark, rtl)
        try:
            build_content_slide(prs, ctx, ns, i, total)
        except Exception:
            # drop the broken slide and rebuild it as a simple bullets slide
            _delete_slide(prs, prs.slides[len(prs.slides) - 1])
            safe = dict(ns, visual=None, image=None)
            build_content_slide(prs, ctx, safe, i, total)

    if takeaways:
        ctx = Ctx(T, True, rtl)
        s = _new_slide(prs, ctx)
        add_text(s, ctx, 0.7, 0.45, 11.93, 1.2,
                 [dict(t="اہم نکات" if rtl else "Key takeaways", size=38, color=ctx.text, bold=True, head=True)],
                 MSO_ANCHOR.MIDDLE)
        draw_rows(s, ctx, [(t, None) for t in takeaways[:6]], 0.7, 1.8, 11.93, 5.2)
    if quiz:
        ctx = Ctx(T, False, rtl)
        s = _new_slide(prs, ctx)
        add_text(s, ctx, 0.7, 0.45, 11.93, 1.2,
                 [dict(t="کوئز" if rtl else "Quick quiz", size=38, color=ctx.text, bold=True, head=True)], MSO_ANCHOR.MIDDLE)
        draw_rows(s, ctx, [(str(q.get("q", "")), None) for q in quiz[:5]], 0.7, 1.8, 11.93, 5.2)
        s2 = _new_slide(prs, ctx)
        add_text(s2, ctx, 0.7, 0.45, 11.93, 1.2,
                 [dict(t="جوابات" if rtl else "Answers", size=38, color=ctx.text, bold=True, head=True)], MSO_ANCHOR.MIDDLE)
        draw_rows(s2, ctx, [(str(q.get("q", "")), str(q.get("a", ""))) for q in quiz[:5]], 0.7, 1.8, 11.93, 5.2)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
