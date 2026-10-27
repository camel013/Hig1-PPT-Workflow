# -*- coding: utf-8 -*-
"""Shared layout primitives for the Zhiwen teaching deck.

All measurements are in 1280x720 design pixels; 1 px = 9525 EMU.
Chinese: FangSong for headings and labels, STZhongsong for body.
Latin letters and figures: Times New Roman.
"""
import math
import re
from functools import lru_cache
from pathlib import Path

from lxml import etree
from PIL import ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

PX = 9525
W, H = 1280, 720

# Palette sampled from the supplied cover.
PAPER = "F3EDE4"
PAPER_WARM = "EDE3D6"
WHITE = "FFFFFF"
INK = "141A20"
MUTED = "6B6862"
RULE = "D8CCBE"
CORAL = "D56A58"
CORAL_DARK = "A64A3D"
CORAL_LIGHT = "E9A79B"
INDIGO = "1B4F8B"
INDIGO_DEEP = "0E305F"
NAVY = "0C182B"
TEAL = "649788"
TEAL_DARK = "3E6B60"
GOLD = "B09263"
GOLD_LIGHT = "D9C4A3"

FANG = "仿宋"
SONG = "华文中宋"
TNR = "Times New Roman"

FONT_FILES = {
    FANG: r"C:\Windows\Fonts\simfang.ttf",
    SONG: r"C:\Users\camel\AppData\Local\Microsoft\Windows\Fonts\STZHONGS.TTF",
    TNR: r"C:\Windows\Fonts\times.ttf",
    TNR + "B": r"C:\Windows\Fonts\timesbd.ttf",
}

OPEN_PUNCT = "“‘「『（《〈【〔"
CLOSE_PUNCT = "，。；：、！？）》」』】”’…·,.;:!?)]}"

LATIN_RE = re.compile(r"[A-Za-z0-9]+(?:[.'’\-–/%:][A-Za-z0-9]+)*")

prs = Presentation()
prs.slide_width = Emu(W * PX)
prs.slide_height = Emu(H * PX)
BLANK = prs.slide_layouts[6]


def rgb(h):
    return RGBColor.from_string(h.lstrip("#"))


def blend(h1, h2, t):
    """Mix two sRGB hex colours; t=0 returns h1, t=1 returns h2."""
    a = [int(h1.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(h2.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(a[i] * (1 - t) + b[i] * t):02X}" for i in range(3))


@lru_cache(maxsize=None)
def pil_font(font, size, bold=False):
    path = FONT_FILES[(font + "B") if bold and font == TNR else font]
    return ImageFont.truetype(path, int(round(size)))


def measure(s, size, font=SONG, bold=False):
    f = pil_font(font, size, bold)
    # PIL measures Latin well; CJK advance is uniform for this face.
    return float(f.getlength(s))


def tokenize(s):
    out = []
    i = 0
    while i < len(s):
        m = LATIN_RE.match(s, i)
        if m:
            out.append((m.group(0), "latin"))
            i = m.end()
            continue
        ch = s[i]
        out.append((ch, "space" if ch in " \u00A0" else "cjk"))
        i += 1
    return out


def clean(tokens):
    while tokens and tokens[0][1] == "space":
        tokens.pop(0)
    while tokens and tokens[-1][1] == "space":
        tokens.pop()
    return tokens


def wrap_tokens(s, size, width, font=SONG, bold=False):
    toks = tokenize(s)
    lines, cur, cur_w = [], [], 0.0

    def tw(t):
        return 0 if t[1] == "space" else measure(t[0], size, TNR if t[1] == "latin" else font, bold)

    for tok in toks:
        w = tw(tok)
        if tok[1] == "space":
            if cur and cur_w + w <= width:
                cur.append(tok)
                cur_w += w
            continue
        if not cur:
            cur, cur_w = [tok], w
            continue
        if cur_w + w <= width + 0.01:
            cur.append(tok)
            cur_w += w
            continue
        if tok[1] == "cjk" and tok[0] in CLOSE_PUNCT:
            cur.append(tok)
            cur_w += w
            continue
        if cur[-1][1] == "cjk" and cur[-1][0] in OPEN_PUNCT and len(cur) > 1:
            moved = cur.pop()
            cur_w -= tw(moved)
            cur = clean(cur)
            if cur:
                lines.append(cur)
            cur, cur_w = [moved, tok], tw(moved) + w
            continue
        cur = clean(cur)
        if cur:
            lines.append(cur)
        cur, cur_w = [tok], w
    cur = clean(cur)
    if cur:
        lines.append(cur)

    # no closing punctuation at line start
    for i in range(1, len(lines)):
        while lines[i] and lines[i][0][1] == "cjk" and lines[i][0][0] in CLOSE_PUNCT:
            lines[i - 1].append(lines[i].pop(0))
            lines[i] = clean(lines[i])
    # no opening punctuation at line end
    for i in range(len(lines) - 1):
        while lines[i] and lines[i][-1][1] == "cjk" and lines[i][-1][0] in OPEN_PUNCT:
            lines[i + 1].insert(0, lines[i].pop())
            lines[i] = clean(lines[i])
    return [clean(x) for x in lines if clean(x)]


def _set_font(run, latin, ea, size, color, bold=False, spacing=0):
    run.font.size = Pt(size * 0.75)
    run.font.name = latin
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    rpr = run._r.get_or_add_rPr()
    for tag, face in (("a:latin", latin), ("a:ea", ea), ("a:cs", latin)):
        el = rpr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rpr, qn(tag))
        el.set("typeface", face)
    if spacing:
        rpr.set("spc", str(int(spacing * 0.75 * 100)))


def textbox(slide, x, y, w, h, wrap=True):
    tb = slide.shapes.add_textbox(Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX)))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    return tb


LATIN_SEG = re.compile(r"[A-Za-z0-9]+(?:[.'’%/:\-–][A-Za-z0-9]+)*")


def _script_segments(s):
    out = []
    i = 0
    while i < len(s):
        m = LATIN_SEG.match(s, i)
        if m:
            out.append((m.group(0), True))
            i = m.end()
            continue
        j = i
        while j < len(s) and not LATIN_SEG.match(s, j):
            j += 1
        out.append((s[i:j], False))
        i = j
    return out


def _emit_runs(p, s, size, base, color, bold=False, spacing=0):
    for seg, is_latin in _script_segments(s):
        if not seg:
            continue
        r = p.add_run()
        r.text = seg
        if is_latin:
            _set_font(r, TNR, base, size, color, bold, spacing)
        else:
            _set_font(r, TNR, base, size, color, bold, spacing)


def draw_line(slide, x, y, line, size, font, color, align=PP_ALIGN.LEFT, width=1000,
              bold=False, spacing=0, box_w=None):
    toks = line
    s = "".join(t for t, _ in toks)
    lw = sum(measure(t, size, TNR if k == "latin" else font, bold) for t, k in toks if k != "space")
    left = x
    if align == PP_ALIGN.CENTER:
        left = x + ((box_w or width) - lw) / 2.0
    elif align == PP_ALIGN.RIGHT:
        left = x + (box_w or width) - lw
    if s and s[0] in OPEN_PUNCT:
        left -= 0.40 * size
    tb = textbox(slide, left, y, max((box_w or width) + size, lw + size * 0.3), size * 1.28, wrap=False)
    tb.name = ("body_" if font == SONG else "heading_") + s[:24]
    p = tb.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    _emit_runs(p, s, size, font, color, bold, spacing)
    return lw


def text(slide, x, y, w, s, size=18, font=SONG, color=INK, leading=1.65,
         align=PP_ALIGN.LEFT, spacing=0, bold=False, max_lines=None, label=None):
    if "\n" in s:
        yy = y
        parts = [part for part in s.split("\n") if part.strip()]
        for i, part in enumerate(parts):
            yy = text(slide, x, yy, w, part, size, font, color, leading,
                      align, spacing, bold, max_lines, label)
            if i < len(parts) - 1:
                yy += size * 0.18
        return yy
    lines = wrap_tokens(s, size, w, font, bold)
    if max_lines and len(lines) > max_lines:
        print(f"[OVERFLOW] {label or s[:18]}: {len(lines)} lines > {max_lines}")
    for i, line in enumerate(lines):
        draw_line(slide, x, y + i * size * leading, line, size, font, color, align, w, bold, spacing)
    return y + len(lines) * size * leading


def rich(slide, x, y, w, segments, size=18, font=SONG, leading=1.65, align=PP_ALIGN.LEFT):
    """segments: [(text, color)] or [(text, color, bold)]."""
    toks = []
    for seg in segments:
        t, c = seg[0], seg[1]
        b = seg[2] if len(seg) > 2 else False
        for tok in tokenize(t):
            toks.append((tok[0], tok[1], c, b))
    lines, cur, cur_w = [], [], 0.0

    def tw(t):
        return 0 if t[1] == "space" else measure(t[0], size, TNR if t[1] == "latin" else font, t[3])

    for tok in toks:
        ww = tw(tok)
        if tok[1] == "space":
            if cur and cur_w + ww <= w:
                cur.append(tok)
                cur_w += ww
            continue
        if not cur:
            cur, cur_w = [tok], ww
        elif cur_w + ww <= w + 0.01 or (tok[1] == "cjk" and tok[0] in CLOSE_PUNCT):
            cur.append(tok)
            cur_w += ww
        else:
            cur = clean(cur)
            if cur:
                lines.append(cur)
            cur, cur_w = [tok], ww
    cur = clean(cur)
    if cur:
        lines.append(cur)

    for i, line in enumerate(lines):
        lw = sum(tw(t) for t in line)
        left = x
        if align == PP_ALIGN.CENTER:
            left = x + (w - lw) / 2.0
        elif align == PP_ALIGN.RIGHT:
            left = x + w - lw
        tb = textbox(slide, left, y + i * size * leading, max(w, lw + size), size * 1.28, wrap=False)
        tb.name = ("body_" if font == SONG else "heading_") + "".join(t[0] for t in line)[:24]
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        j = 0
        while j < len(line):
            color = line[j][2]
            bold = line[j][3]
            buf = []
            while j < len(line) and line[j][2] == color and line[j][3] == bold:
                buf.append(line[j][0])
                j += 1
            _emit_runs(p, "".join(buf), size, font, color, bold)
    return y + len(lines) * size * leading


def vline(slide, x, y, s, size=28, color=INK, font=SONG, leading=1.18, bottom=None, spacing=1.5):
    """One vertical column of upright characters, drawn glyph by glyph.

    The artifact renderer ignores bodyPr/@vert, so each character gets its own
    text box. This also gives exact control over the column rhythm.
    """
    chars = list(s)
    step = size * leading
    top = y if bottom is None else bottom - len(chars) * step
    for i, ch in enumerate(chars):
        if ch in " \u00A0":
            continue
        draw_line(slide, x, top + i * step, [(ch, "latin" if ch.isascii() else "cjk")],
                  size, font, color, PP_ALIGN.LEFT, size * 1.4, False, spacing)
    return top


def rect(slide, x, y, w, h, fill=INK, line=None, line_w=1.0):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(x * PX)), Emu(int(y * PX)),
                                Emu(int(w * PX)), Emu(int(h * PX)))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = rgb(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = rgb(line)
        sp.line.width = Emu(int(max(line_w, 2.0) * PX))
    sp.shadow.inherit = False
    st = sp._element.find(qn("p:style"))
    if st is not None:
        sp._element.remove(st)
    return sp


def hline(slide, x1, x2, y, color=RULE, w=3.0):
    return rect(slide, x1, y, x2 - x1, max(w, 3.0), fill=color)


def vrule(slide, x, y1, y2, color=RULE, w=3.0):
    return rect(slide, x, y1, max(w, 3.0), y2 - y1, fill=color)


def freeform(slide, pts, color=INK, width=3.0, close=False, fill=None):
    b = slide.shapes.build_freeform(pts[0][0], pts[0][1], scale=PX)
    b.add_line_segments(pts[1:], close=close)
    sp = b.convert_to_shape()
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = rgb(fill)
    if color is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = rgb(color)
        sp.line.width = Emu(int(max(width, 3.0) * PX))
    sp.shadow.inherit = False
    st = sp._element.find(qn("p:style"))
    if st is not None:
        sp._element.remove(st)
    return sp


def arc(slide, cx, cy, rx, ry, a0, a1, color=GOLD, width=3.0, steps=72):
    pts = []
    for i in range(steps + 1):
        t = math.radians(a0 + (a1 - a0) * i / steps)
        pts.append((cx + rx * math.cos(t), cy + ry * math.sin(t)))
    return freeform(slide, pts, color=color, width=width)


def contour(slide, x0, x1, base, amp=12, freq=2.0, phase=0.0, color=RULE, width=2.5,
            slope=0.0, steps=120, wobble=0.0):
    pts = []
    for i in range(steps + 1):
        x = x0 + (x1 - x0) * i / steps
        y = base + slope * (x - x0) + amp * math.sin(freq * 2 * math.pi * (x - x0) / max(1, x1 - x0) + phase)
        if wobble:
            y += wobble * math.sin(7.3 * (x - x0) / max(1, x1 - x0) + phase * 1.7)
        pts.append((x, y))
    return freeform(slide, pts, color=color, width=width)



def disc(slide, cx, cy, r, fill=TEAL):
    sp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(int((cx - r) * PX)), Emu(int((cy - r) * PX)),
                                Emu(int(2 * r * PX)), Emu(int(2 * r * PX)))
    sp.fill.solid()
    sp.fill.fore_color.rgb = rgb(fill)
    sp.line.fill.background()
    sp.shadow.inherit = False
    st = sp._element.find(qn("p:style"))
    if st is not None:
        sp._element.remove(st)
    return sp


def arc_block(slide, cx, cy, r, a0, a1, thick=22000, fill=TEAL):
    """A filled annular sector. Angles in degrees, 0 = 3 o'clock, clockwise.

    thick is the ring thickness in 1/1000 of the radius (22000 = 22%).
    """
    sp = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, Emu(int((cx - r) * PX)),
                                Emu(int((cy - r) * PX)), Emu(int(2 * r * PX)),
                                Emu(int(2 * r * PX)))
    sp.fill.solid()
    sp.fill.fore_color.rgb = rgb(fill)
    sp.line.fill.background()
    sp.shadow.inherit = False
    st = sp._element.find(qn("p:style"))
    if st is not None:
        sp._element.remove(st)
    geom = sp._element.spPr.find(qn("a:prstGeom"))
    av = geom.find(qn("a:avLst"))
    for g in list(av):
        av.remove(g)
    for name, val in (("adj1", int(a0 * 60000)), ("adj2", int(a1 * 60000)), ("adj3", int(thick))):
        gd = etree.SubElement(av, qn("a:gd"))
        gd.set("name", name)
        gd.set("fmla", f"val {val}")
    return sp


def new_slide(bg=PAPER):
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = rgb(bg)
    add_page_layers(s, bg, len(prs.slides))
    return s


def meta(slide, page, label, dark=False, total=95):
    fg = GOLD_LIGHT if dark else MUTED
    text(slide, 64, 30, 720, label, 9, FANG, fg, 1.2, spacing=2.2)
    text(slide, 1080, 30, 136, f"{page:02d} / {total}", 9, TNR, fg, 1.2, PP_ALIGN.RIGHT)


def footer(slide, page, dark=False, color=None, total=95):
    fg = color or (GOLD_LIGHT if dark else MUTED)
    text(slide, 64, 690, 640, "《蛭文》教学读本 / 文本细读与写作", 8.5, FANG, fg, 1.2, spacing=1.2)
    text(slide, 1120, 690, 96, f"{page:02d}", 8.5, TNR, fg, 1.2, PP_ALIGN.RIGHT)


# Low-chroma trio sampled from the supplied cover. Every page carries this
# three-colour index at the foot so no page falls back to a one-accent scheme.
MORANDI = ("B56F63", "71877D", "B39A78")
MORANDI_DARK = ("D49286", "8FA79C", "D6BE9A")


def color_chord(slide, dark=False, x=64, y=676, w=42, h=8, gap=4):
    cols = MORANDI_DARK if dark else MORANDI
    for i, col in enumerate(cols):
        rect(slide, x + i * (w + gap), y, w, h, fill=col)


def _line_alpha(sp, pct):
    ln = sp.line._get_or_add_ln()
    solid = ln.find(qn("a:solidFill"))
    if solid is None:
        return
    clr = solid.find(qn("a:srgbClr"))
    if clr is None:
        return
    alpha = etree.SubElement(clr, qn("a:alpha"))
    alpha.set("val", str(int(pct * 1000)))


def topo_rings(slide, cx, cy, d, color=GOLD, rings=3, alpha=24, width=0.9):
    """Thin offset rings in the manner of the Duoyuhua deck.

    The rings are deliberately asymmetric and slightly rotated; they never
    form a concentric target. They sit behind content and stay low-contrast.
    """
    for i in range(rings):
        k = 0.88 - i * (0.44 / max(rings - 1, 1))
        w, h = d * k * 1.06, d * k * 0.92
        ox, oy = (i + 1) * 5, -(i + 1) * 4
        sp = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Emu(int((cx - w / 2 + ox) * PX)), Emu(int((cy - h / 2 + oy) * PX)),
            Emu(int(w * PX)), Emu(int(h * PX)))
        sp.fill.background()
        sp.line.color.rgb = rgb(color)
        sp.line.width = Pt(width * 0.75)
        sp.shadow.inherit = False
        sp.rotation = -6 + i * 5
        sp.name = "bgcurve_contour"
        st = sp._element.find(qn("p:style"))
        if st is not None:
            sp._element.remove(st)
        _line_alpha(sp, alpha)


def add_page_layers(slide, bg, index):
    """Background contours plus a two-stage colour transition at page edge."""
    dark = bg in (NAVY, INDIGO_DEEP, TEAL_DARK) or bg == "0B1C2E"
    contour_cols = MORANDI_DARK if dark else MORANDI
    specs = [
        (1168, 238, 430, 3),
        (214, 642, 390, 2),
        (1080, 654, 320, 3),
        (532, 110, 280, 2),
    ]
    cx, cy, d, rings = specs[(index - 1) % 4]
    topo_rings(slide, cx, cy, d, contour_cols[(index - 1) % 3], rings,
               alpha=20 if dark else 17, width=0.9)

    # A hard-edged two-stage transition, not a gradient: muted mid-tone,
    # then a narrow deeper edge. It stays outside the 1216px text field.
    y = 118 + ((index - 1) % 3) * 58
    mid = contour_cols[(index + 1) % 3]
    deep = contour_cols[(index + 2) % 3]
    mid_block = rect(slide, 1252, y, 28, 286, fill=mid)
    mid_block.name = "bgtransition_mid"
    deep_block = rect(slide, 1244, y + 34, 8, 218, fill=deep)
    deep_block.name = "bgtransition_edge"


def add_color_chords():
    for idx, slide in enumerate(prs.slides, 1):
        bg = slide.background.fill.fore_color.rgb
        dark = str(bg).upper() in (NAVY, INDIGO_DEEP, "0B1C2E")
        color_chord(slide, dark=dark)


def page_title(slide, page, kicker, title, dark=False, color=None, rule=True):
    fg = WHITE if dark else INK
    meta(slide, page, kicker, dark)
    text(slide, 64, 64, 1000, title, 34, FANG, color or fg, 1.2)
    if rule:
        hline(slide, 64, 1216, 118, GOLD_LIGHT if dark else RULE, 3.0)


def tag(slide, x, y, s, color=CORAL, fg=WHITE, size=9, padx=9, h=22):
    w = measure(s, size, SONG) + padx * 2
    rect(slide, x, y, w, h, fill=color)
    text(slide, x + padx, y + 5, w - padx, s, size, SONG, fg, 1.1)
    return w


def unit_no(slide, x, y, s, size=220, color=GOLD, outline=False):
    if outline:
        tb = textbox(slide, x, y, size * 1.15, size * 1.2, wrap=False)
        p = tb.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = s
        _set_font(r, TNR, TNR, size, color)
        r.font.fill.background()
        r.font.line.color.rgb = rgb(color)
        r.font.line.width = Pt(1.2)
        return tb
    return text(slide, x, y, size * 1.15, s, size, FANG, color, 1.0)


