from __future__ import annotations

import json
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
SRC = ROOT / "images"
OPT = ROOT / "images_opt"
POSTERS = ROOT / "posters"
OUT = ROOT / "_deck_build.pptx"

BODY = "华文中宋"
TITLE = "Trajan Pro Deck"
QUOTE = "楷体"
OUTLINE_TITLES = True
TITLE_ASSETS = ROOT / "title_assets"
TITLE_FONT_PATH = Path(r"C:\Users\camel\AppData\Local\Microsoft\Windows\Fonts\TrajanPro-Regular.ttf")
TITLE_COUNTER = 0

PAPER = RGBColor(0xF3, 0xF0, 0xE7)
INK = RGBColor(0x1F, 0x2D, 0x2D)
INK_SOFT = RGBColor(0x4E, 0x59, 0x57)
BLUE = RGBColor(0x2D, 0x43, 0x60)
BLUE_MID = RGBColor(0x5C, 0x71, 0x86)
BLUE_PALE = RGBColor(0xC9, 0xD4, 0xD8)
GREEN = RGBColor(0x73, 0x83, 0x6B)
GREEN_DARK = RGBColor(0x42, 0x58, 0x49)
GREEN_PALE = RGBColor(0xB7, 0xC2, 0xA9)
MOSS = RGBColor(0x27, 0x3A, 0x33)
DEEP_BLUE = RGBColor(0x20, 0x31, 0x44)
DEEP_BLUE_2 = RGBColor(0x2A, 0x43, 0x52)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W = 13.333
SLIDE_H = 7.5
LEFT = 0.62
RIGHT = 12.70
VIDEO_SLOTS: list[dict] = []


def hex_color(color: RGBColor) -> str:
    return str(color)


def prepare_images() -> None:
    OPT.mkdir(parents=True, exist_ok=True)
    for n in range(1, 15):
        src = SRC / f"{n:02d}.jpg"
        dst = OPT / f"{n:02d}.jpg"
        im = Image.open(src)
        im = ImageOps.exif_transpose(im).convert("RGB")
        long_edge = max(im.size)
        if long_edge > 1800:
            scale = 1800 / long_edge
            im = im.resize(
                (max(1, round(im.width * scale)), max(1, round(im.height * scale))),
                Image.Resampling.LANCZOS,
            )
        im.save(dst, "JPEG", quality=86, optimize=True, progressive=True)


def set_font(run, size: float, color: RGBColor, font: str = BODY, bold: bool = False, italic: bool = False, tracking: float = 0.0) -> None:
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rpr.append(el)
        el.set("typeface", font)
    rpr.set("lang", "en-US")
    if tracking:
        rpr.set("spc", str(int(tracking * 100)))


def rgb_tuple(color: RGBColor) -> tuple[int, int, int]:
    value = str(color)
    return (int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def add_title_image(
    slide,
    x: float,
    y: float,
    w: float,
    text: str,
    size: float,
    color: RGBColor,
    align=PP_ALIGN.LEFT,
    line_spacing: float = 1.03,
    name: str | None = None,
):
    global TITLE_COUNTER
    TITLE_COUNTER += 1
    TITLE_ASSETS.mkdir(exist_ok=True)
    px_per_pt = 4
    px_per_in = 288
    font_px = max(1, round(size * px_per_pt))
    font = ImageFont.truetype(str(TITLE_FONT_PATH), font_px)
    max_width = max(1, round(w * px_per_in))
    measure = Image.new("RGBA", (1, 1))
    measure_draw = ImageDraw.Draw(measure)

    lines: list[str] = []
    for paragraph in str(text).split("\n"):
        current = ""
        for word in paragraph.split(" "):
            candidate = word if not current else f"{current} {word}"
            if not current or measure_draw.textlength(candidate, font=font) <= max_width:
                current = candidate
            else:
                lines.append(current)
                current = word
        lines.append(current)

    line_height = max(1, round(size * line_spacing * px_per_pt))
    canvas = Image.new("RGBA", (max_width, line_height * len(lines) + font_px), (0, 0, 0, 0))
    draw = ImageDraw.Draw(canvas)
    fill = (*rgb_tuple(color), 255)
    for index, line in enumerate(lines):
        line_width = draw.textlength(line, font=font)
        if align == PP_ALIGN.RIGHT:
            draw_x = max_width - line_width
        elif align == PP_ALIGN.CENTER:
            draw_x = (max_width - line_width) / 2
        else:
            draw_x = 0
        draw.text((draw_x, index * line_height), line, font=font, fill=fill, anchor="lt")

    bbox = canvas.getbbox()
    if bbox:
        canvas = canvas.crop((0, bbox[1], max_width, bbox[3]))
    path = TITLE_ASSETS / f"title_{TITLE_COUNTER:02d}.png"
    canvas.save(path)
    pic = slide.shapes.add_picture(
        str(path), Inches(x), Inches(y), width=Inches(w), height=Inches(canvas.height / px_per_in)
    )
    pic.name = name or f"TITLE_{TITLE_COUNTER:02d}"
    pic._element.nvPicPr.cNvPr.set("descr", str(text))
    return pic


def add_text(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    text,
    size: float = 18,
    color: RGBColor = INK,
    font: str = BODY,
    bold: bool = False,
    italic: bool = False,
    align=PP_ALIGN.LEFT,
    valign=MSO_ANCHOR.TOP,
    line_spacing: float = 1.3,
    tracking: float = 0.0,
    margins: float = 0.0,
    name: str | None = None,
):
    if OUTLINE_TITLES and font == TITLE:
        return add_title_image(
            slide,
            x,
            y,
            w,
            text,
            size,
            color,
            align=align,
            line_spacing=line_spacing,
            name=name,
        )
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(margins)
    tf.margin_right = Inches(margins)
    tf.margin_top = Inches(margins)
    tf.margin_bottom = Inches(margins)
    tf.vertical_anchor = valign
    paragraphs = text.split("\n") if isinstance(text, str) else list(text)
    for index, item in enumerate(paragraphs):
        p = tf.paragraphs[0] if index == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        runs = item if isinstance(item, list) else [(str(item), color, size, font, bold, italic, tracking)]
        for run_spec in runs:
            run_text, run_color, run_size, run_font, run_bold, run_italic, run_tracking = run_spec
            run = p.add_run()
            run.text = run_text
            set_font(run, run_size, run_color, run_font, run_bold, run_italic, run_tracking)
    return box


def add_rule(slide, x1: float, y: float, x2: float, color: RGBColor = BLUE_PALE, width: float = 0.8) -> None:
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x1), Inches(y), Inches(x2 - x1), Inches(width / 72))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_vrule(slide, x: float, y1: float, y2: float, color: RGBColor = BLUE_PALE, width: float = 0.8) -> None:
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y1), Inches(width / 72), Inches(y2 - y1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_rect(slide, x: float, y: float, w: float, h: float, color: RGBColor, line: RGBColor | None = None, line_width: float = 0.8):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(line_width)
    return shape


def add_round_rect(slide, x: float, y: float, w: float, h: float, color: RGBColor, radius: float = 0.16, line: RGBColor | None = None, line_width: float = 0.8):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.adjustments[0] = max(0.01, min(0.5, radius / min(w, h)))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(line_width)
    return shape


def apply_gradient(shape, stops: list[tuple[float, RGBColor, int]], angle: float = 0.0) -> None:
    sp_pr = shape._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:pattFill", "a:blipFill", "a:grpFill"):
        for el in sp_pr.findall(qn(tag)):
            sp_pr.remove(el)
    grad = OxmlElement("a:gradFill")
    grad.set("rotWithShape", "1")
    gs_lst = OxmlElement("a:gsLst")
    for pos, color, alpha in stops:
        gs = OxmlElement("a:gs")
        gs.set("pos", str(int(pos * 1000)))
        srgb = OxmlElement("a:srgbClr")
        srgb.set("val", hex_color(color))
        if alpha < 100:
            a = OxmlElement("a:alpha")
            a.set("val", str(int(alpha * 1000)))
            srgb.append(a)
        gs.append(srgb)
        gs_lst.append(gs)
    grad.append(gs_lst)
    lin = OxmlElement("a:lin")
    lin.set("ang", str(int(angle * 60000)))
    lin.set("scaled", "1")
    grad.append(lin)
    ln = sp_pr.find(qn("a:ln"))
    if ln is not None:
        ln.addprevious(grad)
    else:
        sp_pr.append(grad)


def add_gradient_rect(slide, x: float, y: float, w: float, h: float, stops, angle: float = 0.0, name: str | None = None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        shape.name = name
    shape.line.fill.background()
    apply_gradient(shape, stops, angle)
    return shape


def add_arc(slide, x: float, y: float, w: float, h: float, color: RGBColor, width: float = 1.2):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.background()
    shape.line.color.rgb = color
    shape.line.width = Pt(width)
    return shape


def add_dot(slide, x: float, y: float, d: float, color: RGBColor):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_picture_cover(slide, path: Path, x: float, y: float, w: float, h: float, name: str | None = None):
    pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        pic.name = name
    with Image.open(path) as im:
        iw, ih = im.size
    target = w / h
    source = iw / ih
    if source > target:
        crop = (1 - target / source) / 2
        pic.crop_left = crop
        pic.crop_right = crop
    elif source < target:
        crop = (1 - source / target) / 2
        pic.crop_top = crop
        pic.crop_bottom = crop
    pic.line.fill.background()
    return pic


def add_caption(slide, x: float, y: float, w: float, text: str, color: RGBColor) -> None:
    add_text(slide, x, y, w, 0.22, text, 8.5, color, tracking=0.65, line_spacing=1.0)


def add_header(slide, number: int, label: str, chinese: str | None = None, dark: bool = False) -> None:
    fg = RGBColor(0xD9, 0xE1, 0xDB) if dark else BLUE
    sub = RGBColor(0xA9, 0xB9, 0xB0) if dark else BLUE_MID
    add_text(slide, LEFT, 0.38, 8.9, 0.24, f"{number:02d} / 10    {label}", 9.5, fg, tracking=0.95, line_spacing=1.0)
    if chinese:
        add_text(slide, 9.55, 0.38, 3.15, 0.24, chinese, 9.5, sub, align=PP_ALIGN.RIGHT, line_spacing=1.0)
    add_rule(slide, LEFT, 0.72, RIGHT, RGBColor(0x55, 0x6B, 0x70) if dark else RGBColor(0xC8, 0xCE, 0xC5), 0.65)


def add_page_number(slide, number: int, dark: bool = False) -> None:
    add_text(slide, 12.20, 7.08, 0.48, 0.2, f"{number:02d}", 9.5, RGBColor(0xC9, 0xD4, 0xD8) if dark else BLUE, align=PP_ALIGN.RIGHT, line_spacing=1.0)


def new_slide(prs, background: RGBColor = PAPER):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = background
    return slide


def add_video_placeholder(slide, number: int, slide_index: int, video_name: str, poster: Path, x: float, y: float, w: float, h: float, label: str) -> None:
    movie = slide.shapes.add_movie(
        str(SRC / video_name),
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
        poster_frame_image=str(poster),
        mime_type="video/mp4",
    )
    movie.name = f"FIELD_VIDEO_{number}"
    add_gradient_rect(slide, x, y + h - min(0.52, h * 0.20), w, min(0.52, h * 0.20), [(0, MOSS, 0), (100, MOSS, 74)], 90, name=f"VIDEO_LABEL_BED_{number}")
    add_text(slide, x + 0.12, y + h - 0.29, w - 0.24, 0.18, label, 7.8, WHITE, tracking=0.55, line_spacing=1.0)
    VIDEO_SLOTS.append({"number": number, "slide": slide_index, "video": str((SRC / video_name).resolve()), "x": x, "y": y, "w": w, "h": h})


def add_bullet_item(slide, x: float, y: float, w: float, text: str, size: float = 15.5, color: RGBColor = INK_SOFT, marker: RGBColor = GREEN) -> None:
    add_rect(slide, x, y + 0.09, 0.065, 0.065, marker)
    add_text(slide, x + 0.18, y, w - 0.18, 0.75, text, size, color, line_spacing=1.32)


def cubic_points(p0, p1, p2, p3, count=90):
    pts = []
    for i in range(count):
        t = i / (count - 1)
        u = 1 - t
        x = u**3 * p0[0] + 3 * u**2 * t * p1[0] + 3 * u * t**2 * p2[0] + t**3 * p3[0]
        y = u**3 * p0[1] + 3 * u**2 * t * p1[1] + 3 * u * t**2 * p2[1] + t**3 * p3[1]
        pts.append((x, y))
    return pts


def add_polyline(slide, points, color: RGBColor, width: float = 1.4):
    scale = 914400
    builder = slide.shapes.build_freeform(points[0][0] * scale, points[0][1] * scale, scale=1.0)
    builder.add_line_segments([(x * scale, y * scale) for x, y in points[1:]], close=False)
    shape = builder.convert_to_shape()
    shape.fill.background()
    shape.line.color.rgb = color
    shape.line.width = Pt(width)
    return shape


def add_arc_block(slide, cx: float, cy: float, r_outer: float, r_inner: float, start: float, end: float, color: RGBColor):
    outer_count = max(32, int(abs(end - start) * 2.2))
    inner_count = outer_count
    outer = []
    inner = []
    for i in range(outer_count):
        a = math.radians(start + (end - start) * i / (outer_count - 1))
        outer.append((cx + math.cos(a) * r_outer, cy + math.sin(a) * r_outer))
    for i in range(inner_count):
        a = math.radians(end + (start - end) * i / (inner_count - 1))
        inner.append((cx + math.cos(a) * r_inner, cy + math.sin(a) * r_inner))
    points = outer + inner
    scale = 914400
    builder = slide.shapes.build_freeform(points[0][0] * scale, points[0][1] * scale, scale=1.0)
    builder.add_line_segments([(x * scale, y * scale) for x, y in points[1:]], close=True)
    shape = builder.convert_to_shape()
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_partial_contour(slide, cx: float, cy: float, rx: float, ry: float, start: float, end: float, color: RGBColor, width: float):
    count = max(20, int(abs(end - start) * 1.7))
    pts = []
    for i in range(count):
        a = math.radians(start + (end - start) * i / (count - 1))
        pts.append((cx + math.cos(a) * rx, cy + math.sin(a) * ry))
    return add_polyline(slide, pts, color, width)


def strip_style_and_shadow(slide) -> None:
    for shape in slide.shapes:
        style = shape._element.find(qn("p:style"))
        if style is not None:
            shape._element.remove(style)
        try:
            shape.shadow.inherit = False
        except Exception:
            pass


def slide_01(prs):
    s = new_slide(prs, DEEP_BLUE)
    add_gradient_rect(s, 0, 0, SLIDE_W, SLIDE_H, [(0, DEEP_BLUE, 100), (62, DEEP_BLUE_2, 100), (100, MOSS, 100)], 18, "bg_gradient")
    add_arc_block(s, 11.50, 4.10, 5.00, 4.54, 225, 315, RGBColor(0x36, 0x4E, 0x4A))
    add_arc(s, 6.25, -2.8, 9.9, 9.9, RGBColor(0x7D, 0x96, 0x8F), 1.35)
    add_arc(s, 8.52, 0.18, 7.1, 7.1, RGBColor(0x4C, 0x6A, 0x67), 0.75)
    add_text(s, LEFT, 0.53, 7.0, 0.28, "ENGLISH GROUP SHARE  /  ONE-WEEK FIELD WALK", 9.5, GREEN_PALE, tracking=1.15, line_spacing=1.0)
    add_text(s, LEFT, 1.34, 7.0, 0.82, "Campus", 58, PAPER, font=TITLE, line_spacing=0.98)
    add_text(s, LEFT, 2.12, 7.0, 0.82, "Environment", 58, PAPER, font=TITLE, line_spacing=0.98)
    add_text(s, LEFT, 3.10, 7.0, 0.55, "Three corners. One problem.", 25, GREEN_PALE, font=QUOTE, line_spacing=1.05)
    add_rule(s, LEFT, 3.88, 7.24, PAPER, 1.15)
    add_text(s, LEFT, 4.18, 6.58, 1.12, "We walked Huxi and B Campus for a week. What we found was not only litter.", 17, RGBColor(0xD8, 0xDF, 0xD5), line_spacing=1.45)
    add_text(s, LEFT, 5.53, 6.2, 0.34, "HUXI  ·  B CAMPUS", 11, GREEN_PALE, tracking=0.8, line_spacing=1.0)
    add_text(s, LEFT, 6.80, 5.0, 0.3, "CAMPUS FIELD STUDY  /  2026", 9.5, RGBColor(0xA6, 0xB5, 0xAD), tracking=0.8, line_spacing=1.0)
    cover_imgs = [(8.52, 0.38, 1.25, 5.28, "01.jpg"), (9.94, 1.02, 1.25, 5.50, "07.jpg"), (11.36, 0.62, 1.25, 6.34, "13.jpg")]
    for i, (x, y, w, h, file) in enumerate(cover_imgs, 1):
        add_picture_cover(s, OPT / file, x, y, w, h)
        add_text(s, x, 7.06, w, 0.18, f"0{i}", 8.5, RGBColor(0xC9, 0xD4, 0xD8), tracking=0.8, line_spacing=1.0)
    add_page_number(s, 1, dark=True)
    return s


def slide_02(prs):
    s = new_slide(prs)
    add_gradient_rect(s, 0, 0, SLIDE_W, SLIDE_H, [(0, PAPER, 100), (58, PAPER, 100), (100, RGBColor(0xE5, 0xEB, 0xE2), 100)], 0, "bg_gradient")
    add_header(s, 2, "HUXI / HEYUAN EDGE", "虎溪 · 荷园边缘")
    add_round_rect(s, 0, 0.80, 5.72, 1.56, BLUE, radius=0.20, line=RGBColor(0x4B, 0x62, 0x78), line_width=0.6)
    add_round_rect(s, 5.72, 0.80, 0.18, 1.56, GREEN, radius=0.09, line=GREEN_PALE, line_width=0.5)
    add_text(s, 0.82, 0.91, 4.78, 1.38, "Litter hides\nin the grass.", 36, PAPER, font=TITLE, line_spacing=1.03)
    add_text(s, 6.50, 1.02, 4.95, 0.92, "Plastic bags, bottles, wrappers and cigarette butts. Some are old and half-buried; some are new.", 15.5, INK_SOFT, line_spacing=1.42)
    add_text(s, 6.50, 2.12, 4.95, 0.42, "There is no bin on this path.", 18, GREEN_DARK, line_spacing=1.1)
    add_picture_cover(s, OPT / "13.jpg", 0.62, 2.68, 3.55, 4.14)
    add_picture_cover(s, OPT / "10.jpg", 4.36, 2.68, 2.52, 1.64)
    add_picture_cover(s, OPT / "12.jpg", 7.06, 2.68, 2.52, 1.64)
    add_picture_cover(s, OPT / "09.jpg", 4.36, 4.50, 2.04, 2.32)
    add_picture_cover(s, OPT / "11.jpg", 6.58, 4.50, 1.55, 2.32)
    add_picture_cover(s, OPT / "14.jpg", 8.31, 4.50, 1.28, 2.32)
    add_vrule(s, 9.72, 2.55, 6.90, GREEN_PALE, 0.7)
    add_video_placeholder(s, 1, 2, "video3.mp4", POSTERS / "video3.jpg", 9.92, 2.63, 2.34, 4.16, "SHRUBS / CLOSE-UP")
    add_caption(s, 0.62, 6.88, 3.3, "HEYUAN SLOPE / ONE WEEK IN", BLUE_MID)
    add_page_number(s, 2)
    return s


def slide_03(prs):
    s = new_slide(prs, PAPER)
    add_gradient_rect(s, 0, 0, SLIDE_W, SLIDE_H, [(0, RGBColor(0xE8, 0xEC, 0xE0), 100), (58, PAPER, 100), (100, RGBColor(0xDD, 0xE4, 0xDF), 100)], 0, "bg_gradient")
    add_header(s, 3, "GREEN BELT", "校园绿化带")
    belt_block = add_round_rect(s, 0, 0.80, 8.70, 1.23, GREEN, radius=0.20)
    apply_gradient(belt_block, [(0, GREEN, 100), (50, GREEN, 100), (100, GREEN, 0)], 90)
    add_text(s, 0.82, 0.94, 7.65, 0.72, "Clean-looking is not clean.", 32, PAPER, font=TITLE, line_spacing=1.02)
    add_text(s, 9.22, 1.02, 3.30, 0.74, "The gap is under the hedge.", 16, GREEN_DARK, line_spacing=1.28)
    add_picture_cover(s, OPT / "05.jpg", 0.62, 2.03, 5.12, 2.97)
    add_picture_cover(s, OPT / "06.jpg", 5.92, 2.03, 2.38, 1.44)
    add_picture_cover(s, OPT / "07.jpg", 8.48, 2.03, 1.92, 1.44)
    add_picture_cover(s, OPT / "08.jpg", 5.92, 3.56, 2.38, 1.44)
    add_video_placeholder(s, 2, 3, "video2.mp4", POSTERS / "video2.jpg", 8.48, 3.56, 4.22, 2.37, "FIELD FOOTAGE / SLOPE")
    add_caption(s, 0.62, 5.08, 4.65, "SURFACE / SWEPT", BLUE_MID)
    add_caption(s, 5.92, 5.08, 2.2, "UNDER THE HEDGE", GREEN_DARK)
    add_caption(s, 8.48, 6.04, 4.2, "THE WORK THAT IS NOT ON THE LIST", BLUE_MID)
    add_rule(s, 0.62, 5.47, 5.12, GREEN_PALE, 0.65)
    add_text(s, 0.62, 5.64, 4.65, 0.76, "The gardener cuts the grass.\nThe cleaner sweeps the path.", 17, INK_SOFT, line_spacing=1.38)
    add_text(s, 0.62, 6.48, 4.85, 0.46, "Nobody cleans under the bushes.", 20, GREEN_DARK, line_spacing=1.08)
    add_page_number(s, 3)
    return s


def slide_04(prs):
    s = new_slide(prs, DEEP_BLUE)
    add_gradient_rect(s, 0, 0, SLIDE_W, SLIDE_H, [(0, DEEP_BLUE, 100), (62, DEEP_BLUE_2, 100), (100, MOSS, 100)], 180, "bg_gradient")
    add_arc(s, 4.8, -4.9, 10.1, 10.1, RGBColor(0x58, 0x72, 0x68), 0.85)
    add_header(s, 4, "B CAMPUS", "B 区老旧建筑", dark=True)
    add_text(s, LEFT, 0.92, 10.25, 0.83, "Heyuan is dirty.  B is forgotten.", 34, PAPER, font=TITLE, line_spacing=1.03)
    add_text(s, 10.55, 1.05, 2.0, 0.6, "A different scale of neglect.", 14.5, GREEN_PALE, line_spacing=1.3, align=PP_ALIGN.RIGHT)
    add_video_placeholder(s, 3, 4, "video1.mp4", POSTERS / "video1.jpg", 0.62, 2.29, 6.20, 3.49, "B CAMPUS / BUILDING WASTE")
    add_picture_cover(s, OPT / "01.jpg", 7.04, 1.93, 2.62, 1.95)
    add_picture_cover(s, OPT / "02.jpg", 9.84, 1.93, 2.62, 1.95)
    add_picture_cover(s, OPT / "03.jpg", 7.04, 4.06, 2.62, 1.95)
    add_picture_cover(s, OPT / "04.jpg", 9.84, 4.06, 2.62, 1.95)
    add_text(s, 0.62, 5.94, 5.95, 0.92, "Some of it has been there so long that grass is growing through it.", 21, GREEN_PALE, font=QUOTE, line_spacing=1.25)
    add_caption(s, 9.84, 6.10, 2.65, "FACADES / WASTE / YEARS", RGBColor(0xC9, 0xD4, 0xD8))
    add_page_number(s, 4, dark=True)
    return s


def slide_05(prs):
    s = new_slide(prs, PAPER)
    add_gradient_rect(s, 0, 0, SLIDE_W, SLIDE_H, [(0, PAPER, 100), (55, RGBColor(0xF1, 0xF0, 0xE8), 100), (100, RGBColor(0xDC, 0xE5, 0xDD), 100)], 0, "bg_gradient")
    add_header(s, 5, "WHY CORNERS STAY DIRTY", "三个原因")
    add_text(s, LEFT, 0.91, 8.2, 0.76, "Why a corner stays dirty.", 32, INK, font=TITLE, line_spacing=1.03)
    reasons = [
        (0.62, 1.78, "NO", "OWNER", "Between offices, no one feels responsible.", BLUE, GREEN_PALE),
        (0.96, 3.15, "NO", "SYSTEM", "No rule says who moves the waste, and when.", GREEN_DARK, BLUE_PALE),
        (1.30, 4.52, "NO", "COST", "Leaving it is free, so the waste stays.", GREEN, GREEN_PALE),
    ]
    for idx, (x, y, small, big, copy, fill, accent) in enumerate(reasons, 1):
        add_round_rect(s, x, y, 5.24, 0.92, fill, radius=0.18, line=accent, line_width=0.6)
        add_text(s, x + 0.16, y + 0.12, 0.42, 0.20, f"0{idx}", 9.0, accent, tracking=0.7, line_spacing=1.0)
        add_text(s, x + 0.59, y + 0.11, 0.48, 0.20, small, 9.5, accent, tracking=1.0, line_spacing=1.0)
        add_text(s, x + 0.57, y + 0.30, 4.35, 0.54, big, 37, PAPER, font=TITLE, line_spacing=0.95)
        add_text(s, x + 0.16, y + 0.98, 4.72, 0.28, copy, 13.8, INK_SOFT, line_spacing=1.2)
    cx, cy = 9.55, 3.72
    add_arc_block(s, 10.45, 3.55, 3.30, 2.78, 195, 335, RGBColor(0xDD, 0xE4, 0xDA))
    rings = [
        (3.12, 2.36, 7.0, 147.0, BLUE_PALE, 0.75),
        (2.78, 2.05, -28.0, 172.0, BLUE_MID, 0.85),
        (2.38, 1.76, 18.0, 151.0, GREEN, 1.05),
        (1.98, 1.46, 73.0, 262.0, GREEN_DARK, 1.25),
        (1.53, 1.10, 142.0, 318.0, BLUE, 1.45),
        (1.08, 0.76, 35.0, 210.0, BLUE_MID, 0.85),
    ]
    for rx, ry, start, end, color, width in rings:
        add_partial_contour(s, cx, cy, rx, ry, start, end, color, width)
    add_partial_contour(s, cx, cy, 0.97, 0.63, 210, 345, GREEN, 0.7)
    add_dot(s, cx - 0.045, cy - 0.045, 0.09, GREEN_DARK)
    add_text(s, cx + 0.14, cy - 0.12, 1.55, 0.25, "temporary pile", 9.5, BLUE_MID, tracking=0.3, line_spacing=1.0)
    add_text(s, 6.92, 5.83, 5.4, 0.82, "A temporary pile\nbecomes permanent.", 28, GREEN_DARK, font=QUOTE, line_spacing=1.12)
    add_page_number(s, 5)
    return s


def slide_06(prs):
    s = new_slide(prs, PAPER)
    add_picture_cover(s, OPT / "02.jpg", 8.50, 0, 4.833, SLIDE_H)
    add_gradient_rect(s, 7.20, 0, 2.22, SLIDE_H, [(0, PAPER, 100), (100, PAPER, 0)], 0, "image_fade")
    add_gradient_rect(s, 8.50, 6.45, 4.833, 1.05, [(0, MOSS, 0), (100, MOSS, 78)], 90, "image_caption_fade")
    add_header(s, 6, "RISK", "危害")
    add_text(s, LEFT, 0.92, 7.25, 0.78, "Ugly is the least of it.", 34, INK, font=TITLE, line_spacing=1.03)
    add_text(s, LEFT, 1.74, 6.30, 0.54, "The same corner holds four risks.", 16, GREEN_DARK, line_spacing=1.2)
    add_round_rect(s, 0, 2.20, 8.50, 4.42, GREEN_DARK, radius=0.26, line=RGBColor(0x5F, 0x74, 0x63), line_width=0.8)
    rows = [
        ("01", "SAFETY", "Loose tiles and broken glass can fall.", GREEN_PALE),
        ("02", "FIRE", "Dry grass and old doors sit against a building.", BLUE_PALE),
        ("03", "HEALTH", "Plastic and standing water breed mosquitoes.", GREEN_PALE),
        ("04", "IMAGE", "Every visitor sees a campus that neglects its corners.", BLUE_PALE),
    ]
    for i, (num, label, copy, accent) in enumerate(rows):
        y = 2.55 + i * 1.03
        add_text(s, 0.64, y + 0.02, 0.38, 0.2, num, 9.5, accent, tracking=0.65, line_spacing=1.0)
        add_text(s, 1.08, y - 0.01, 1.08, 0.24, label, 11.5, accent, tracking=1.05, line_spacing=1.0)
        add_text(s, 2.15, y - 0.06, 5.62, 0.52, copy, 18.5, PAPER, line_spacing=1.22)
        add_rule(s, 0.64, y + 0.68, 7.78, RGBColor(0x9E, 0xAE, 0x9E), 0.65)
    add_text(s, 8.72, 6.61, 4.15, 0.24, "B CAMPUS / BROKEN FACADE", 8.5, WHITE, tracking=0.7, line_spacing=1.0)
    add_page_number(s, 6)
    return s


def add_bin_icon(slide, x: float, y: float, scale: float = 1.0):
    add_rect(slide, x, y, 0.34 * scale, 0.50 * scale, PAPER, BLUE, 1.25)
    add_rule(slide, x - 0.04 * scale, y - 0.06 * scale, x + 0.38 * scale, BLUE, 1.25)
    add_rule(slide, x + 0.11 * scale, y - 0.11 * scale, x + 0.23 * scale, BLUE, 1.1)
    add_vrule(slide, x + 0.17 * scale, y + 0.06 * scale, y + 0.43 * scale, GREEN, 0.8)


def add_sign_icon(slide, x: float, y: float):
    add_vrule(slide, x + 0.26, y + 0.25, y + 0.82, BLUE_MID, 1.1)
    add_rect(slide, x, y, 0.58, 0.28, PAPER, GREEN_DARK, 1.0)
    add_text(slide, x + 0.06, y + 0.055, 0.46, 0.16, "BIN 15 m", 7.2, GREEN_DARK, align=PP_ALIGN.CENTER, line_spacing=1.0)


def add_hedge_icon(slide, x: float, y: float):
    add_rect(slide, x, y + 0.22, 1.16, 0.13, GREEN)
    for i in range(7):
        add_vrule(slide, x + 0.07 + i * 0.16, y + 0.05, y + 0.25, GREEN_DARK, 0.9)


def add_light_icon(slide, x: float, y: float):
    add_vrule(slide, x + 0.22, y + 0.25, y + 0.76, BLUE_MID, 1.2)
    add_dot(slide, x + 0.09, y, 0.26, GREEN_PALE)
    for dx, dy in [(-0.15, 0.02), (0.41, 0.02), (-0.12, 0.28), (0.38, 0.28)]:
        add_vrule(slide, x + 0.22 + dx, y + 0.05 + dy, y + 0.17 + dy, GREEN_DARK, 0.7)


def slide_07(prs):
    s = new_slide(prs, PAPER)
    add_gradient_rect(s, 0, 0, SLIDE_W, SLIDE_H, [(0, RGBColor(0xDF, 0xE7, 0xE5), 100), (57, PAPER, 100), (100, RGBColor(0xE9, 0xEE, 0xE1), 100)], 0, "bg_gradient")
    add_header(s, 7, "FIX / HUXI", "虎溪 · 设施")
    add_round_rect(s, 0, 0.80, 9.32, 2.12, BLUE, radius=0.28, line=RGBColor(0x4A, 0x60, 0x76), line_width=0.8)
    add_round_rect(s, 9.32, 0.80, 4.013, 2.12, GREEN_PALE, radius=0.28, line=RGBColor(0xC8, 0xD1, 0xC0), line_width=0.8)
    add_text(s, 0.82, 0.93, 8.15, 1.92, "Fix the place.\nMake the easy choice\nthe clean one.", 35, PAPER, font=TITLE, line_spacing=1.04)
    add_text(s, 9.60, 1.00, 3.28, 1.18, "Four small changes, visible at the point of use.", 15.5, GREEN_DARK, line_spacing=1.42)
    path = cubic_points((0.88, 5.62), (4.35, 5.44), (7.05, 3.58), (12.48, 3.82), 110)
    add_polyline(s, path, GREEN, 1.8)
    for x, y in [(1.48, 5.14), (4.23, 5.02), (7.45, 4.20), (10.44, 3.79)]:
        add_dot(s, x - 0.055, y - 0.055, 0.11, GREEN_DARK)
    add_bin_icon(s, 1.12, 4.56, 1.0)
    add_sign_icon(s, 3.83, 4.35)
    add_hedge_icon(s, 7.07, 3.86)
    add_light_icon(s, 10.05, 3.18)
    add_text(s, 0.72, 5.83, 2.20, 0.22, "01 / TWO BINS", 10.5, BLUE, tracking=0.9, line_spacing=1.0)
    add_text(s, 0.72, 6.06, 2.38, 0.40, "At each end of the path", 14.5, INK_SOFT, line_spacing=1.22)
    add_text(s, 3.42, 5.24, 2.15, 0.22, "02 / ONE SIGN", 10.5, BLUE, tracking=0.9, line_spacing=1.0)
    add_text(s, 3.42, 5.47, 2.28, 0.46, "Show the distance to the nearest bin", 14.5, INK_SOFT, line_spacing=1.22)
    add_text(s, 6.61, 4.48, 2.18, 0.22, "03 / LOW HEDGE", 10.5, GREEN_DARK, tracking=0.9, line_spacing=1.0)
    add_text(s, 6.61, 4.71, 2.34, 0.40, "Stop the shortcut through the grass", 14.5, INK_SOFT, line_spacing=1.22)
    add_text(s, 9.40, 4.08, 2.48, 0.22, "04 / SOLAR LIGHT", 10.5, GREEN_DARK, tracking=0.9, line_spacing=1.0)
    add_text(s, 9.40, 4.31, 2.55, 0.40, "Litter less where it is dark", 14.5, INK_SOFT, line_spacing=1.22)
    add_picture_cover(s, OPT / "14.jpg", 0, 6.45, SLIDE_W, 1.05)
    add_gradient_rect(s, 0, 6.45, SLIDE_W, 1.05, [(0, DEEP_BLUE, 76), (75, DEEP_BLUE, 24), (100, DEEP_BLUE, 0)], 0, "strip_fade")
    add_text(s, 0.66, 6.87, 6.25, 0.22, "HUXI / THE SAME SLOPE, MADE EASIER TO KEEP CLEAN", 8.5, PAPER, tracking=0.7, line_spacing=1.0)
    add_page_number(s, 7)
    return s


def slide_08(prs):
    s = new_slide(prs, PAPER)
    add_gradient_rect(s, 0, 0, SLIDE_W, SLIDE_H, [(0, PAPER, 100), (58, PAPER, 100), (100, RGBColor(0xE4, 0xEA, 0xDF), 100)], 0, "bg_gradient")
    add_header(s, 8, "FIX / WIDER", "绿化带 · B 区")
    add_text(s, LEFT, 0.91, 5.78, 1.28, "One checklist.\nTwo different places.", 29.5, INK, font=TITLE, line_spacing=1.05)
    add_picture_cover(s, OPT / "03.jpg", 0.62, 2.34, 4.62, 4.13)
    add_caption(s, 0.62, 6.49, 4.4, "B CAMPUS / COVERED, BUT OPEN", BLUE_MID)
    add_round_rect(s, 6.02, 1.80, 6.68, 1.82, GREEN, radius=0.22, line=RGBColor(0x8B, 0x9A, 0x80), line_width=0.7)
    add_text(s, 6.32, 1.97, 0.45, 0.20, "01", 9.0, GREEN_PALE, tracking=0.7, line_spacing=1.0)
    add_text(s, 6.86, 1.94, 1.55, 0.23, "GREEN BELT", 10.5, PAPER, tracking=1.0, line_spacing=1.0)
    add_text(s, 6.32, 2.21, 6.02, 0.84, "Add one line to the weekly checklist.", 22.5, PAPER, font=TITLE, line_spacing=1.10)
    add_text(s, 6.32, 3.12, 5.95, 0.40, "Clean under the bushes. One line. No extra budget.", 15.3, PAPER, line_spacing=1.30)
    add_round_rect(s, 6.02, 3.78, 6.68, 2.46, BLUE, radius=0.22, line=RGBColor(0x4B, 0x61, 0x78), line_width=0.7)
    add_text(s, 6.32, 3.97, 0.45, 0.20, "02", 9.0, GREEN_PALE, tracking=0.7, line_spacing=1.0)
    add_text(s, 6.86, 3.94, 1.55, 0.23, "B CAMPUS", 10.5, BLUE_PALE, tracking=1.0, line_spacing=1.0)
    add_text(s, 6.32, 4.22, 5.92, 0.55, "Safety first. Then a system.", 25, PAPER, font=TITLE, line_spacing=1.10)
    add_text(s, 6.32, 4.94, 5.90, 0.78, "Remove loose tiles and glass. Fence the risky zone. One covered waste point, scheduled pick-up, a written log.", 15.0, PAPER, line_spacing=1.34)
    add_text(s, 6.32, 5.89, 5.65, 0.30, "Make the invisible job visible.", 17.2, GREEN_PALE, line_spacing=1.08)
    add_page_number(s, 8)
    return s


def slide_09(prs):
    s = new_slide(prs, PAPER)
    add_gradient_rect(s, 0, 0, SLIDE_W, SLIDE_H, [(0, RGBColor(0xDF, 0xE6, 0xE7), 100), (48, PAPER, 100), (100, RGBColor(0xE5, 0xEA, 0xDE), 100)], 0, "bg_gradient")
    add_header(s, 9, "OWN / TEACH / TRACK", "责任闭环")
    add_text(s, LEFT, 0.91, 8.4, 0.75, "Give every corner a name.", 35.5, INK, font=TITLE, line_spacing=1.03)
    add_text(s, LEFT, 1.63, 7.0, 0.38, "A closed loop needs an owner, a habit, and a record.", 16, GREEN_DARK, line_spacing=1.2)
    columns = [
        (0.62, 3.68, BLUE, GREEN_PALE, "01", "OWN", [
            "One named team: logistics, dorm manager, students.",
            "A weekly inspection and photo report.",
            "Written into the cleaning contract.",
        ]),
        (4.46, 3.68, GREEN, BLUE_PALE, "02", "TEACH", [
            "Use these photographs in orientation.",
            "Bilingual dorm posters and green guards.",
            "Monthly plogging. One hedge day each term.",
        ]),
        (8.30, 3.90, BLUE_MID, GREEN_PALE, "03", "TRACK", [
            "A QR code for photo reports.",
            "A hot-spot map and monthly data.",
            "A small reward for the cleanest corner.",
        ]),
    ]
    for x, w, fill, accent, num, title, items in columns:
        add_round_rect(s, x, 2.20, w, 4.00, fill, radius=0.22, line=accent, line_width=0.55)
        add_text(s, x + 0.24, 2.37, 0.55, 0.22, num, 9.5, accent, tracking=0.7, line_spacing=1.0)
        add_text(s, x + 0.24, 2.60, w - 0.44, 0.60, title, 33, PAPER, font=TITLE, line_spacing=0.98)
        add_rule(s, x + 0.24, 3.28, x + w - 0.22, accent, 0.7)
        for j, item in enumerate(items):
            add_bullet_item(s, x + 0.24, 3.55 + j * 0.82, w - 0.42, item, 14.2, PAPER, accent)
    add_rule(s, 0.62, 6.34, 12.20, RGBColor(0xBF, 0xC8, 0xBC), 0.65)
    add_text(s, 0.62, 6.53, 1.35, 0.21, "MATERIAL BANK", 9.5, GREEN_DARK, tracking=0.9, line_spacing=1.0)
    add_text(s, 2.02, 6.48, 10.15, 0.39, "Keep reusable doors, tiles and buckets for student design projects.", 15.5, INK_SOFT, line_spacing=1.15)
    add_page_number(s, 9)
    return s


def slide_10(prs):
    s = new_slide(prs, DEEP_BLUE)
    add_gradient_rect(s, 0, 0, SLIDE_W, SLIDE_H, [(0, DEEP_BLUE, 100), (64, DEEP_BLUE_2, 100), (100, MOSS, 100)], 16, "bg_gradient")
    add_arc_block(s, 11.80, 2.20, 4.80, 4.24, 70, 170, RGBColor(0x3D, 0x56, 0x4C))
    add_arc(s, -2.45, -2.40, 10.1, 10.1, RGBColor(0x6E, 0x8A, 0x7E), 1.35)
    add_arc(s, 7.90, -4.55, 8.3, 8.3, RGBColor(0x50, 0x6D, 0x67), 0.75)
    add_text(s, LEFT, 0.53, 5.2, 0.28, "THE ASK", 9.5, GREEN_PALE, tracking=1.15, line_spacing=1.0)
    add_text(s, LEFT, 1.18, 7.8, 0.88, "Two bins.", 63, PAPER, font=TITLE, line_spacing=0.98)
    add_text(s, LEFT, 2.02, 7.8, 0.88, "One owner.", 63, GREEN_PALE, font=TITLE, line_spacing=0.98)
    add_text(s, LEFT, 2.86, 7.8, 0.88, "One hour.", 63, PAPER, font=TITLE, line_spacing=0.98)
    add_rule(s, LEFT, 3.93, 7.02, PAPER, 1.1)
    add_text(s, LEFT, 4.20, 5.85, 1.10, "One sign. A shared habit. If we take care of the corners, the whole campus stays clean.", 17, RGBColor(0xD9, 0xE0, 0xD6), line_spacing=1.44)
    add_picture_cover(s, OPT / "14.jpg", 0, 5.55, SLIDE_W, 1.95)
    add_gradient_rect(s, 0, 5.55, SLIDE_W, 1.95, [(0, DEEP_BLUE, 100), (30, DEEP_BLUE, 70), (100, DEEP_BLUE, 0)], 90, "footer_fade")
    add_text(s, LEFT, 5.73, 11.2, 0.24, "TWO BINS  /  ONE SIGN  /  ONE OWNER  /  ONE HOUR", 9.5, PAPER, tracking=1.0, line_spacing=1.0)
    add_text(s, 11.38, 0.52, 1.30, 0.22, "10 / 10", 9.5, BLUE_PALE, tracking=0.9, align=PP_ALIGN.RIGHT, line_spacing=1.0)
    return s


def build() -> Path:
    prepare_images()
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    for make_slide in (slide_01, slide_02, slide_03, slide_04, slide_05, slide_06, slide_07, slide_08, slide_09, slide_10):
        make_slide(prs)
    for slide in prs.slides:
        strip_style_and_shadow(slide)
    props = prs.core_properties
    props.title = "Campus Environment: Three Corners, One Problem"
    props.subject = "English group share / campus environment field study"
    props.author = "Campus Environment Group"
    props.keywords = "campus, environment, Huxi, Heyuan, B Campus, waste, maintenance"
    props.comments = "10 pages; native editable text; blue-green Morandi palette; all 14 photographs and 3 videos included."
    prs.save(OUT)
    (ROOT / "video_slots.json").write_text(json.dumps(VIDEO_SLOTS, ensure_ascii=False, indent=2), encoding="utf-8")
    return OUT


if __name__ == "__main__":
    print(build())
