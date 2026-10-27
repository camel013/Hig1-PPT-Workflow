# -*- coding: utf-8 -*-
"""Build the native, editable 12-slide deck.

Design system
-------------
Morandi palette by default. Strong colour clashes (CLASH_*) are reserved for
the farewell spread and the closing page only.

Background curves are composed from three layers:
    bgcurve_rim       intermediate-colour arc that softens a hard colour edge
    bgcurve           the solid / tinted mass
    bgcurve_contour   thin topographic rings, never animated

No standalone emotional curves are added. All curve geometry belongs to the
background colour masses; text pages stay strictly rectilinear.

Picture pipeline (for later use; no placeholder is inserted now)
----------------------------------------------------------------
    add_picture_cover(...)      real embedded picture, cover-cropped, no stretch
    golden_frames(...)          38.2 / 61.8 split inside the 80 px grid
    add_image_mask(...)         solid or translucent overlap for layered masks

Example
-------
    text_box, image_box = golden_frames(side="right")
    add_picture_cover(s, r"images/portrait.jpg", *image_box)
    add_rect(s, image_box[0] - 28, image_box[1] + 28, 180, 120, SLATE)

All geometry stays on the 1280 x 720 canvas with 80 / 1200 margins.

Build pipeline
--------------
    python build_pptx.py                    -> _deck_build.pptx
    python merge_legacy.py _deck_build.pptx <audited-final.pptx> _deck_final.pptx
    copy _deck_final.pptx over the Chinese delivery name

The raw build intentionally does not overwrite the delivered 39 MB file.
"""
import pathlib
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

P = pathlib.Path(r"C:\Users\camel\Documents\ChatGPT\Hig1\projects\duoyuhua_ppt169_20261021")
OUT = P / "_deck_build.pptx"
PX = 9525
W, H = 1280, 720
M, RIGHT = 80, 1200
IMGDIR = P / "images"

PAPER = "#F4F1E9"      # warm paper ground
INK = "#232323"
MUTED = "#6E6A64"
RULE = "#D3CEC3"

# Morandi range: greyed, low-chroma, all readable with INK or PALE.
CLAY = "#8E5750"       # terracotta
SLATE = "#5A707E"      # grey blue
OCHRE = "#B49360"      # muted ochre
SAGE = "#7D8D7B"       # sage
PLUM = "#7D6B72"       # grey mauve
PALE = "#F7F3EA"
DARK = "#212426"

# 10-12 percent tints for background masses.
PALE_CLAY = "#E7D8D2"
PALE_SLATE = "#DCE3E6"
PALE_OCHRE = "#E8DECA"
PALE_SAGE = "#DFE4DB"
PALE_TEAL = "#D9E2DF"
GROUND = "#ECE6D8"

# Clash colours appear only on the emotional pages (11, 12).
CLASH_RED = "#873A33"
CLASH_TEAL = "#315553"
CLASH_LIGHT = "#C46A5E"
SLATE_MID = "#9FAAAE"
CLASH_RED_MID = "#542F2C"
CLASH_TEAL_MID = "#293C3C"

WHITE = PALE

KAI, DENG, SONG = "楷体", "等线", "华文中宋"
NUM, LAT = "Trajan Pro Deck", "Times Newer Roman"
FAM = {"title": KAI, "sub": DENG, "body": SONG, "num": NUM, "latin": LAT}

# 一套行距只服务一种文本类型；同类型跨页必须复用同一组数值。
ROLE = {
    "display":  (132, 152, "title"),
    "quote":    (108, 124, "title"),
    "hero":     (76, 90, "title"),
    "hero2":    (52, 66, "sub"),
    "h1":       (54, 68, "title"),
    "h1s":      (44, 58, "title"),
    "h2":       (42, 56, "title"),
    "h3":       (34, 48, "title"),
    "subhead":  (30, 44, "sub"),
    "name":     (40, 52, "sub"),
    "lead":     (28, 40, "sub"),
    "body":     (20, 34, "body"),
    "small":    (17, 28, "body"),
    "note":     (16, 26, "body"),
    "caption":  (13, 20, "latin"),
    "quote-s":  (58, 74, "title"),
}

prs = Presentation()
prs.slide_width = Emu(W * PX)
prs.slide_height = Emu(H * PX)
blank = prs.slide_layouts[6]

def width_est(s, size, ls=0):
    w = 0.0
    for ch in s:
        w += 1.0 if ord(ch) > 127 else 0.56
    return w * size + max(len(s) - 1, 0) * ls

OPEN_PUNCT = "“‘「『（《〈【〔"

def optical_shift(s, size):
    if s and s[0] in OPEN_PUNCT:
        return -0.42 * size
    return 0.0

def _set_run(r, s, size, fkey, color, ls):
    fam = FAM[fkey]
    r.text = s
    r.font.size = Pt(size * 0.75)
    r.font.name = fam
    r.font.bold = fkey == "sub"
    r.font.color.rgb = RGBColor.from_string(color.lstrip("#"))
    rPr = r._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(rPr, qn("a:ea"))
    ea.set("typeface", fam)
    if ls:
        rPr.set("spc", str(int(ls * 0.75 * 100)))

def add_text(slide, x, y, s, size=20, font="body", color=INK, anchor="start", ls=0):
    fam = FAM[font]
    w = width_est(s, size, ls) + 16
    left = (x - w) if anchor == "end" else (x + optical_shift(s, size))
    tb = slide.shapes.add_textbox(Emu(int(left * PX)), Emu(int((y - size * 0.86) * PX)),
                                  Emu(int(w * PX)), Emu(int(size * 1.35 * PX)))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = Pt(size * 1.5 * 0.75)
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    r = p.add_run()
    _set_run(r, s, size, font, color, ls)

def add_block(slide, x, y, lines, role="body", color=INK, anchor="start", ls=0):
    if isinstance(lines, str):
        lines = [lines]
    size, lead, fkey = ROLE[role]
    w = max(width_est(s, size, ls) for s in lines) + 16
    left = (x - w) if anchor == "end" else (x + optical_shift(lines[0], size))
    h = lead * (len(lines) - 1) + size * 1.3
    tb = slide.shapes.add_textbox(Emu(int(left * PX)), Emu(int((y - size * 0.86) * PX)),
                                  Emu(int(w * PX)), Emu(int(h * PX)))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    for i, s in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = Pt(lead * 0.75)
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        r = p.add_run()
        _set_run(r, s, size, fkey, color, ls)

def add_rect(slide, x, y, w, h, color=INK):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(x * PX)), Emu(int(y * PX)),
                                Emu(int(w * PX)), Emu(int(h * PX)))
    sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(color.lstrip("#"))
    sp.line.fill.background(); sp.shadow.inherit = False
    return sp

def _solid_rgb(sp):
    fill = sp._element.spPr.find(qn("a:solidFill"))
    return None if fill is None else fill.find(qn("a:srgbClr"))

def set_fill_alpha(sp, opacity=100):
    """opacity is 0-100; DrawingML stores thousandths of a percent."""
    clr = _solid_rgb(sp)
    if clr is None:
        return
    for old in clr.findall(qn("a:alpha")):
        clr.remove(old)
    alpha = etree.SubElement(clr, qn("a:alpha"))
    alpha.set("val", str(int(max(0, min(100, opacity)) * 1000)))

def set_line(sp, color, width=1, opacity=100):
    sp.line.color.rgb = RGBColor.from_string(color.lstrip("#"))
    sp.line.width = Emu(int(width * PX))
    ln = sp._element.spPr.find(qn("a:ln"))
    fill = ln.find(qn("a:solidFill"))
    clr = fill.find(qn("a:srgbClr"))
    alpha = etree.SubElement(clr, qn("a:alpha"))
    alpha.set("val", str(int(max(0, min(100, opacity)) * 1000)))

def set_soft_edge(sp, radius_px):
    """Large radii produce the misty edge; harmless if a viewer ignores it."""
    sp.shadow.inherit = False
    spPr = sp._element.spPr
    effect = spPr.find(qn("a:effectLst"))
    if effect is None:
        effect = etree.SubElement(spPr, qn("a:effectLst"))
    for old in effect.findall(qn("a:softEdge")):
        effect.remove(old)
    soft = etree.SubElement(effect, qn("a:softEdge"))
    soft.set("rad", str(int(radius_px * PX)))

def _luminance(color):
    h = color.lstrip("#")
    vals = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    vals = [v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4 for v in vals]
    return 0.2126 * vals[0] + 0.7152 * vals[1] + 0.0722 * vals[2]

def mix_hex(a, b, t):
    t = max(0.0, min(1.0, t))
    av = [int(a.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    bv = [int(b.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    return "#" + "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(av, bv))

def add_line(slide, x1, y1, x2, y2, color=RULE, w=1):
    add_rect(slide, x1, y1 - w / 2, max(x2 - x1, 0.5), w, color)

def add_fade_rect(slide, x, y, w, h, color=INK, fade_from=0.55, fade_end=None):
    """Solid rule whose right end dissolves instead of stopping abruptly.

    The shape keeps its full width (so animation targets stay stable); the
    dissolve is produced by alpha stops, not by shortening the geometry.
    """
    sp = add_rect(slide, x, y, w, h, color)
    spPr = sp._element.spPr
    solid = spPr.find(qn("a:solidFill"))
    idx = list(spPr).index(solid)
    spPr.remove(solid)
    hexc = color.lstrip("#").upper()
    p1 = int(round(fade_from * 100000))
    p2 = p1 if fade_end is None else int(round(min(1.0, (fade_end - x) / w) * 100000))
    p2 = max(p2, p1)
    xml = (
        '<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'rotWithShape="1"><a:gsLst>'
        f'<a:gs pos="0"><a:srgbClr val="{hexc}"><a:alpha val="100000"/></a:srgbClr></a:gs>'
        f'<a:gs pos="{p1}"><a:srgbClr val="{hexc}"><a:alpha val="100000"/></a:srgbClr></a:gs>'
        f'<a:gs pos="{p2}"><a:srgbClr val="{hexc}"><a:alpha val="0"/></a:srgbClr></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="{hexc}"><a:alpha val="0"/></a:srgbClr></a:gs>'
        '</a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>'
    )
    spPr.insert(idx, etree.fromstring(xml))
    return sp


def add_dot(slide, cx, cy, r, color=INK):
    sp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(int((cx - r) * PX)), Emu(int((cy - r) * PX)),
                                Emu(int(2 * r * PX)), Emu(int(2 * r * PX)))
    sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(color.lstrip("#"))
    sp.line.fill.background(); sp.shadow.inherit = False

def _add_oval(slide, x, y, w, h, color):
    sp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(x * PX)), Emu(int(y * PX)),
                                Emu(int(w * PX)), Emu(int(h * PX)))
    sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(color.lstrip("#"))
    sp.line.fill.background(); sp.shadow.inherit = False
    return sp

def strip_theme_style(sp):
    """Remove inherited theme effects; an empty effectLst is not enough everywhere."""
    style = sp._element.find(qn("p:style"))
    if style is not None:
        sp._element.remove(style)

def add_contour(slide, cx, cy, d, base, rings=3):
    """Thin, offset rings suggest a topographic section without extra clutter."""
    deep = _luminance(base) > 0.42
    ink = mix_hex(base, INK, 0.34) if deep else mix_hex(base, PALE, 0.38)
    for i in range(rings):
        s = 0.88 - i * (0.44 / max(rings - 1, 1))
        w, h = d * s * 1.06, d * s * 0.92
        ox, oy = (i + 1) * 5, -(i + 1) * 4
        sp = _add_oval(slide, cx - w / 2 + ox, cy - h / 2 + oy, w, h, base)
        sp.fill.background()
        set_line(sp, ink, 0.9, 24)
        sp.rotation = -6 + i * 5
        sp.name = "bgcurve_contour"

def add_curve(slide, x, y, d, color, rings=2, haze=0, soft=0, opacity=100,
              rim=None, rim_pad=24, rim_dx=0, rim_dy=0):
    cx, cy = x + d / 2, y + d / 2
    if rim:
        band = _add_oval(slide, x - rim_pad + rim_dx, y - rim_pad + rim_dy,
                         d + 2 * rim_pad, d + 2 * rim_pad, rim)
        band.name = "bgcurve_rim"
    if haze:
        pad = max(26, d * 0.07)
        halo = _add_oval(slide, x - pad, y - pad, d + 2 * pad, d + 2 * pad, color)
        set_fill_alpha(halo, haze)
        set_soft_edge(halo, pad * 0.9)
        halo.name = "bgcurve_haze"
    sp = _add_oval(slide, x, y, d, d, color)
    if opacity < 100:
        set_fill_alpha(sp, opacity)
    if soft:
        set_soft_edge(sp, soft)
    sp.name = "bgcurve"
    if rings:
        add_contour(slide, cx, cy, d, color, rings)
    return sp

def add_bg_rect(slide, x, y, w, h, color):
    sp = add_rect(slide, x, y, w, h, color)
    sp.name = "bgblock"
    return sp

def head(s, no, kicker, title=None, title_role="h1", ty=176, x=M, x2=RIGHT):
    add_text(s, x, 74, f"{no} / 12", 13, "latin", MUTED, ls=2)
    add_text(s, x2, 74, kicker, 13, "latin", MUTED, anchor="end", ls=2)
    add_line(s, x, 96, x2, 96)
    if title:
        add_block(s, x, ty, title, title_role)

def foot(s, no, dark=False, x1=M, x2=RIGHT, page_color=None):
    c = "#3A3A38" if dark else RULE
    tcol = "#C2C2BC" if dark else MUTED
    add_line(s, x1, 664, x2, 664, c)
    add_text(s, x1, 692, "多余的话", 13, "body", tcol, ls=1)
    add_text(s, x2, 692, f"{no:02d} / 12", 13, "latin", page_color or tcol, anchor="end", ls=2)

PHI = 1.61803398875
GOLDEN_MAJOR = 0.61803398875
GOLDEN_MINOR = 1 - GOLDEN_MAJOR

def golden_frames(x=M, y=176, w=1120, h=430, side="right", gutter=56):
    """Return (text_frame, image_frame) as (x, y, w, h) tuples."""
    usable = w - gutter
    major = usable * GOLDEN_MAJOR
    minor = usable * GOLDEN_MINOR
    if side == "right":
        return (x, y, minor, h), (x + minor + gutter, y, major, h)
    return (x + major + gutter, y, minor, h), (x, y, major, h)

def add_picture_cover(slide, path, x, y, w, h, name="image_content",
                      focus_x=0.5, focus_y=0.5):
    """Embed a real picture and crop it to the frame without distortion."""
    from PIL import Image
    with Image.open(path) as im:
        iw, ih = im.size
    src_ratio, box_ratio = iw / ih, w / h
    pic = slide.shapes.add_picture(str(path), Emu(int(x * PX)), Emu(int(y * PX)),
                                   Emu(int(w * PX)), Emu(int(h * PX)))
    if src_ratio > box_ratio:
        excess = 1 - box_ratio / src_ratio
        left = excess * max(0.0, min(1.0, focus_x))
        pic.crop_left, pic.crop_right = left, excess - left
    elif src_ratio < box_ratio:
        excess = 1 - src_ratio / box_ratio
        top = excess * max(0.0, min(1.0, focus_y))
        pic.crop_top, pic.crop_bottom = top, excess - top
    pic.name = name
    return pic

def add_image_mask(slide, x, y, w, h, color, opacity=100, name="image_mask"):
    """Hard-edged overlap that interleaves picture, colour block and type."""
    sp = add_rect(slide, x, y, w, h, color)
    if opacity < 100:
        set_fill_alpha(sp, opacity)
    sp.name = name
    return sp

def new_slide(blocks=(), curves=(), dark=False):
    s = prs.slides.add_slide(blank)
    if dark:
        base = add_rect(s, 0, 0, W, H, DARK)
        base.name = "bgbase"
    for x, y, w, h, col in blocks:
        add_bg_rect(s, x, y, w, h, col)
    for spec in curves:
        x, y, d, col = spec[:4]
        opts = spec[4] if len(spec) > 4 else {}
        add_curve(s, x, y, d, col, **opts)
    return s

# 01 封面：陶土／鼠尾草双边栏 + 灰蓝大圆
s = new_slide(
    blocks=[(0, 0, 44, H, CLAY), (1240, 0, 40, H, SAGE)],
    curves=[
        (925, -215, 470, SLATE, {"rings": 3, "haze": 0, "soft": 0,
                                 "rim": SLATE_MID, "rim_pad": 24, "rim_dx": 6, "rim_dy": -6}),
        (-380, 520, 680, GROUND, {"rings": 2, "haze": 0, "soft": 0}),
    ],
)
add_text(s, 100, 78, "文献 / 01", 13, "sub", MUTED)
add_text(s, 1140, 78, "1935", 13, "latin", WHITE, anchor="end", ls=2)
add_line(s, 100, 96, 520, 96, INK, 2)
add_text(s, 100, 268, "多余的话", 132, "title")
add_fade_rect(s, 100, 366, 1040, 4, INK, fade_from=0.55, fade_end=778)
add_picture_cover(s, IMGDIR / "qu_portrait_blur.png", 800, 190, 289, 361,
                  focus_x=0.5, focus_y=0.40)
add_line(s, 778, 190, 778, 551, CLAY, 4)
add_text(s, 800, 571, "瞿秋白像", 12, "body", MUTED)
add_text(s, 100, 416, "瞿秋白", 40, "sub")
add_text(s, 100, 470, "一九三五年五月 · 汀州狱中", 20, "body", MUTED)
add_text(s, 100, 600, "SUPERFLUOUS WORDS", 13, "latin", MUTED, ls=3)
add_text(s, 100, 628, "QU QIUBAI   1899—1935", 13, "latin", MUTED, ls=3)
add_text(s, 1140, 628, "最后的坦白", 14, "sub", CLAY, anchor="end", ls=1)

# 02 题记：赭石大块 + 淡陶土圆
s = new_slide(
    blocks=[(760, 470, 440, 250, OCHRE)],
    curves=[(-420, 380, 720, PALE_CLAY, {"rings": 3, "haze": 0, "soft": 0})],
)
add_text(s, 80, 78, "02 / 12", 13, "latin", MUTED, ls=2)
add_text(s, 1200, 78, "PREFACE", 13, "latin", MUTED, anchor="end", ls=2)
add_line(s, 80, 96, 1200, 96)
add_block(s, 80, 230, ["知我者，谓我心忧；", "不知我者，谓我何求。"], "quote-s")
add_text(s, 80, 392, "——《诗经·黍离》", 17, "body", MUTED)
add_line(s, 80, 462, 420, 462, CLAY, 4)
add_text(s, 80, 520, "何必说？", 76, "title")
add_text(s, 80, 600, "代序", 28, "sub", MUTED)
add_text(s, 800, 590, "02", 132, "num", INK)

# 03 代序：整幅灰蓝左栏
s = new_slide(blocks=[(0, 0, 380, H, SLATE), (380, 0, 14, H, OCHRE)])
add_text(s, 80, 74, "03 / 12", 13, "latin", WHITE, ls=2)
add_text(s, 1200, 74, "PREFACE / 代序", 13, "latin", MUTED, anchor="end", ls=2)
add_line(s, 450, 96, 1200, 96)
add_text(s, 16, 430, "03", 238, "num", "#8FA5AF")
add_block(s, 450, 180, ["话既然是多余的，", "又何必说呢？"], "h1s")
add_line(s, 450, 320, 760, 320, CLAY, 4)
add_block(s, 450, 380, ["已经是走到了生命的尽期。余剩的日子，", "不但不能按照年份来算，", "甚至不能按星期来算了。"], "body")
add_text(s, 450, 516, "这是我最后“谈天”的机会呢！", 28, "sub")
add_text(s, 450, 566, "（一九三五·五·一七于汀州狱中）", 13, "latin", MUTED)
foot(s, 3, x1=450)

# 04 历史的误会：双色统计块
s = new_slide(
    blocks=[(80, 460, 520, 190, CLAY), (600, 460, 20, 190, OCHRE), (620, 460, 580, 190, SLATE)],
    curves=[(-280, -300, 620, PALE_SLATE, {"rings": 3, "haze": 0, "soft": 0,
                                           "rim": "#AAB9BE", "rim_pad": 14,
                                           "rim_dx": -5, "rim_dy": 5})],
)
head(s, "04", "THE HISTORICAL MISUNDERSTANDING", "历史的误会")
add_block(s, 80, 246, ["我其实是一个很平凡的文人，", "竟虚负了某某党的领袖的名声十来年。", "这不是“历史的误会”，是什么呢？"], "h2")
add_text(s, 110, 540, "15", 84, "num", WHITE)
add_text(s, 240, 520, "年", 24, "sub", WHITE)
add_text(s, 240, 558, "勉强做着政治工作", 16, "body", WHITE)
add_text(s, 650, 540, "10", 84, "num", WHITE)
add_text(s, 780, 520, "余年", 24, "sub", WHITE)
add_text(s, 780, 558, "虚负领袖之名", 16, "body", WHITE)
foot(s, 4)

# 05 年表：灰蓝年份栏
rows = [
    ("1899", "出生于江苏武进（今常州）"),
    ("1919", "参加五四运动，创办《新社会》"),
    ("1920", "任《晨报》特派记者，赴苏俄"),
    ("1927", "主持中央政治局，担负直接政治领导"),
    ("1931", "被解除中央领导职务，转入左翼文化"),
    ("1935", "六月十八日，就义于福建长汀中山公园"),
]
s = new_slide(blocks=[(80, 210, 220, 430, SLATE), (300, 210, 8, 430, OCHRE)])
head(s, "05", "CHRONOLOGY", "年表")
y = 230
for yr, ev in rows:
    col = WHITE if yr != "1935" else PALE
    add_text(s, 270, y, yr, 20, "latin", col, anchor="end")
    add_dot(s, 304, y - 7, 5 if yr == "1935" else 3, INK if yr == "1935" else PALE)
    add_text(s, 334, y, ev, 20, "body", CLAY if yr == "1935" else INK)
    y += 68
foot(s, 5)

# 06 脆弱的二元人物：陶土数字块
s = new_slide(
    blocks=[(80, 460, 520, 190, CLAY), (600, 460, 20, 190, SLATE)],
    curves=[(860, -240, 520, PALE_OCHRE, {"rings": 2, "haze": 0, "soft": 0})],
)
add_picture_cover(s, IMGDIR / "qu_1935_print.jpg", 950, 100, 220, 330,
                  focus_x=0.5, focus_y=0.50)
head(s, "06", "THE FRAIL DUAL SELF", "脆弱的二元人物")
add_rect(s, 80, 226, 6, 150, CLAY)
add_block(s, 110, 250, ["一只羸弱的马拖着几千斤的辎重车，", "走上了险峻的山坡，一步步地往上爬。"], "h3")
add_text(s, 110, 372, "要往后退是不可能，要再往前去是实在不能胜任了。", 20, "body", MUTED)
add_line(s, 80, 440, 1200, 440)
add_text(s, 110, 540, "36", 96, "num", WHITE)
add_text(s, 250, 520, "岁", 28, "sub", WHITE)
add_block(s, 250, 560, ["自觉已经非常地衰惫，", "丝毫青年壮年的兴趣都没有了"], "small", WHITE)
add_text(s, 950, 448, "一九三五年六月十八日 · 汀州", 12, "body", MUTED)
foot(s, 6)

# 07 我和马克思主义：右下灰蓝块
s = new_slide(
    blocks=[(1000, 430, 280, 210, SLATE), (1260, 430, 20, 210, OCHRE)],
    curves=[(900, -220, 520, PALE_SAGE, {"rings": 2, "haze": 0, "soft": 0,
                                         "rim": "#AEB9A8", "rim_pad": 14,
                                         "rim_dx": 5, "rim_dy": 5})],
)
head(s, "07", "MARXISM / 1923—1927", "我和马克思主义")
add_block(s, 80, 246, ["我对这些学问，", "的确只知道一点皮毛。"], "h2")
add_text(s, 80, 370, "当时我只是根据几本外国文的书籍转译一下，编了一些讲义。", 20, "body", MUTED)
add_line(s, 80, 442, 1200, 442)
add_text(s, 80, 500, "从那时候起，我没有自己的思想。", 30, "sub")
add_text(s, 80, 556, "（我以中央的思想为思想。）", 34, "title", CLAY)
add_text(s, 1030, 540, "07", 80, "num", WHITE)
foot(s, 7)

# 08 五年：赭石顶块 + 陶土五年块
s = new_slide(blocks=[(1040, 0, 240, 230, OCHRE), (980, 400, 300, 260, CLAY)])
head(s, "08", "POLITICAL LEADERSHIP", "五年")
add_text(s, 80, 280, "1927", 28, "latin", MUTED)
add_text(s, 1200, 280, "1931", 28, "latin", MUTED, anchor="end")
add_rect(s, 80, 318, 1120, 10, SLATE)
for i in range(5):
    add_rect(s, 80 + i * 280 - 3, 306, 6, 34, SLATE)
add_block(s, 80, 390, ["整整五年，我居然当了中国共产党的领袖之一，", "最后三年甚至仿佛是最主要的领袖。"], "subhead")
add_block(s, 80, 520, ["“你们饶了我吧，我实在没有兴趣和能力负担这个领导工作了。”", "但是，终于没有说出口。"], "body", MUTED)
add_text(s, 1070, 120, "08", 64, "num", INK)
add_text(s, 1010, 535, "5", 132, "num", WHITE)
add_text(s, 1010, 625, "YEARS", 13, "latin", WHITE, ls=3)
foot(s, 8)

# 09 盲动主义和立三路线：三条色块目录
s = new_slide(curves=[(-260, -260, 520, PALE_SLATE, {"rings": 3, "haze": 0, "soft": 0,
                                                     "rim": "#AAB9BE", "rim_pad": 14,
                                                     "rim_dx": -5, "rim_dy": 5})])
head(s, "09", "BLIND ACTIONISM / LI LISAN LINE", "盲动主义和立三路线")
add_block(s, 80, 250, ["立三路线是我的许多错误观点", "——有人说是瞿秋白主义——逻辑的发展。"], "h3")
rows = [
    ("01", "盲动主义", "一九二八年初的判断错误", CLAY),
    ("02", "立三路线", "左倾空谈掩盖右倾实质", SLATE),
    ("03", "瞿秋白主义", "他人加之于我的命名", SAGE),
]
y = 380
for num, name, desc, col in rows:
    add_rect(s, 80, y, 64, 64, col)
    add_text(s, 96, y + 14, num, 24, "num", INK if col == SAGE else WHITE)
    add_text(s, 180, y + 12, name, 28, "sub")
    add_text(s, 560, y + 18, desc, 17, "body", MUTED)
    if num == "03":
        add_text(s, 560, y + 44, "我当然间接地负着立三路线的责任。", 15, "body", CLAY)
    y += 88
foot(s, 9)

# 10 “文人”：赭石巨型标题块 + 灰蓝圆
s = new_slide(
    blocks=[(30, 140, 660, 260, OCHRE)],
    curves=[(900, 430, 560, SLATE, {"rings": 0, "haze": 0, "soft": 0,
                                    "rim": SLATE_MID, "rim_pad": 26, "rim_dx": 6, "rim_dy": 6})],
)
head(s, "10", "THE LITERATI")
add_text(s, 80, 230, "“文人”", 108, "title")
add_line(s, 80, 420, 420, 420, CLAY, 6)
add_block(s, 80, 450, ["文人积习未除。", "至多不过是个“读者”。"], "lead")
add_block(s, 80, 540, ["极忠实地翻译几部俄国文学名著，每字每句地斟酌，也许不会“误人子弟”。"], "small", MUTED)
add_line(s, 80, 580, 760, 580)
add_text(s, 80, 600, "这一个最愉快的梦想，比创作和评论方面再求得什么成就，要实际得多。", 16, "body")
add_text(s, 80, 628, "可惜，恐怕现在这个可能已经“过时”了！", 16, "body", MUTED)
foot(s, 10, page_color=WHITE)

# 11 告别：唯一允许强撞色的内页
s = new_slide(
    blocks=[(0, 450, 1280, 170, CLASH_RED), (0, 0, 24, H, CLASH_TEAL)],
)
head(s, "11", "FAREWELL", "告别")
add_text(s, 80, 250, "一出滑稽剧就此闭幕了！", 30, "sub")
add_text(s, 80, 330, "我只要休息。", 52, "sub")
add_text(s, 100, 505, "永别了，美丽的世界！", 64, "title", WHITE)
add_text(s, 80, 630, "一九三五·五·二三", 13, "latin", MUTED, ls=2)
foot(s, 11)

# 12 终页：唯一保留大面积撞色的收束页
s = new_slide(
    curves=[
        (-220, -260, 600, CLASH_TEAL, {"rings": 0, "haze": 0, "soft": 0,
                                       "rim": CLASH_TEAL_MID, "rim_pad": 20, "rim_dx": -6, "rim_dy": -6}),
        (760, 420, 600, CLASH_RED, {"rings": 0, "haze": 0, "soft": 0,
                                    "rim": CLASH_RED_MID, "rim_pad": 20, "rim_dx": 6, "rim_dy": 6}),
    ],
    dark=True,
)
add_picture_cover(s, IMGDIR / "monument_teal_gradient.png", 780, 130, 340, 470,
                  focus_x=0.5, focus_y=0.50)
add_text(s, 80, 74, "12 / 12", 13, "latin", "#D8D8D2", ls=2)
add_text(s, 1200, 74, "THE END", 13, "latin", "#9A9A96", anchor="end", ls=2)
add_line(s, 80, 96, 1200, 96, "#3A3A38")
add_block(s, 80, 280, ["中国的豆腐也是很好吃的东西，", "世界第一。"], "h2", WHITE)
add_text(s, 80, 450, "永别了！", 96, "title", CLASH_LIGHT)
add_line(s, 80, 540, 760, 540, "#3A3A38")
add_text(s, 80, 580, "瞿秋白 · 一八九九—一九三五", 18, "body", "#D0D0CC")
add_text(s, 80, 620, "一九三五年六月十八日 · 福建长汀中山公园", 13, "latin", "#9A9A96")
add_text(s, 780, 626, "瞿秋白烈士纪念碑 · 福建长汀", 12, "body", "#D8CBC6")
foot(s, 12, dark=True, page_color=WHITE)

for slide in prs.slides:
    for sp in slide.shapes:
        strip_theme_style(sp)

prs.save(str(OUT))
print("saved 12 slides")
