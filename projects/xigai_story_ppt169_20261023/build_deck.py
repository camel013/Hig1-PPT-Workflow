# -*- coding: utf-8 -*-
"""Build the 12-page native, editable story deck.

Palette: paper, ink, deep red, wine purple.
Type: 楷体 for section headings; 华文中宋 for all other text.
"""
import pathlib
import re
from functools import lru_cache

from lxml import etree
from PIL import ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt


P = pathlib.Path(r"C:\Users\camel\Documents\ChatGPT\Hig1\projects\xigai_story_ppt169_20261023")
OUT = P / "_deck_build.pptx"

PX = 9525
W, H = 1280, 720

PAPER = "#F6F1EB"
INK = "#1C1417"
MUTED = "#6F6665"
RULE = "#D8CACA"
WHITE = "#FFFFFF"

RED = "#9E1B32"
RED_DARK = "#7A1426"
RED_LIGHT = "#C23A4A"
PURPLE = "#4C1D4F"
PURPLE_DARK = "#351335"
PURPLE_LIGHT = "#7A3C7E"
WINE = "#5A1E3A"
ROSE = "#EAD7D9"
LILAC = "#E2D6E5"

KAI = "楷体"
SONG = "华文中宋"

prs = Presentation()
prs.slide_width = Emu(W * PX)
prs.slide_height = Emu(H * PX)
BLANK = prs.slide_layouts[6]


def _rgb(h):
    return RGBColor.from_string(h.lstrip("#"))


FONT_FILES = {
    KAI: r"C:\Windows\Fonts\simkai.ttf",
    SONG: r"C:\Users\camel\AppData\Local\Microsoft\Windows\Fonts\STZHONGS.TTF",
}

OPEN_PUNCT = "“‘「『（《〈【〔"
CLOSE_PUNCT = "，。；：、！？）》」』】”’…·,.;:!?)]}"

PROTECTED = sorted(
    [
        "社会主义现代化建设", "教育、科技、人才", "继续教育", "职业教育", "社区教育",
        "教育强国", "立德树人", "五育并举", "铸魂育人", "社会主义建设者和接班人",
        "为党育人", "为国育才", "服务民族复兴", "实时数据", "慢镜头回放",
        "教育技术", "自惭形秽", "社保材料", "人才引进", "持续成长", "战略人才",
        "高水平科技自立自强", "科技自立自强", "重中之重", "基础性、战略性支撑",
        "国之利器", "人才井喷", "培养什么人", "怎样培养人", "为谁培养人",
        "社会主义建设", "现代化国家", "继续教育中心", "社区小馆", "名校背景",
        "社保材料", "各种考核", "一条卷子", "一张卷子", "基本记录", "判决书",
        "战略科技任务", "国际科技前沿", "自主创新", "真问题", "追风口",
    ],
    key=len,
    reverse=True,
)

TOKEN_RE = re.compile(
    "|".join(
        [re.escape(x) for x in PROTECTED]
        + [
            r"[A-Za-z0-9]+(?:[./\-–—%][A-Za-z0-9]+)*",
            r"[ \u00A0]+",
            r"[\u4e00-\u9fff]",
            r"[^\s]",
        ]
    )
)


@lru_cache(maxsize=None)
def _pil_font(font, size):
    return ImageFont.truetype(FONT_FILES[font], int(round(size)))


def _measure(s, size, font):
    return float(_pil_font(font, size).getlength(s))


def _clean_line(tokens):
    while tokens and tokens[0].isspace():
        tokens.pop(0)
    while tokens and tokens[-1].isspace():
        tokens.pop()
    return "".join(tokens)


def _layout_lines(s, size, width, font):
    tokens = [t for t in TOKEN_RE.findall(s) if t]
    lines, cur, cur_w = [], [], 0.0
    for tok in tokens:
        tw = _measure(tok, size, font)
        if tok.isspace():
            if cur and cur_w + tw <= width:
                cur.append(tok)
                cur_w += tw
            continue
        if not cur:
            cur, cur_w = [tok], tw
            continue
        if cur_w + tw <= width + 0.01:
            cur.append(tok)
            cur_w += tw
            continue
        if tok and tok[0] in CLOSE_PUNCT:
            cur.append(tok)
            cur_w += tw
            continue
        if cur[-1] and cur[-1][-1] in OPEN_PUNCT and len(cur) > 1:
            moved = cur.pop()
            cur_w -= _measure(moved, size, font)
            line = _clean_line(cur)
            if line:
                lines.append(line)
            cur, cur_w = [moved, tok], _measure(moved, size, font) + tw
            continue
        line = _clean_line(cur)
        if line:
            lines.append(line)
        cur, cur_w = [tok], tw
    line = _clean_line(cur)
    if line:
        lines.append(line)
    if not lines:
        return [""]

    for i in range(1, len(lines)):
        while lines[i] and lines[i][0] in CLOSE_PUNCT:
            lines[i - 1] += lines[i][0]
            lines[i] = lines[i][1:]
    lines = [x for x in lines if x]

    for i in range(len(lines) - 1):
        while lines[i] and lines[i][-1] in OPEN_PUNCT:
            lines[i + 1] = lines[i][-1] + lines[i + 1]
            lines[i] = lines[i][:-1]
        if not lines[i]:
            lines[i] = lines[i + 1]
            lines[i + 1] = ""
    lines = [x for x in lines if x]

    while len(lines) > 1 and len(lines[-1]) <= 1:
        if len(lines[-2]) > 1:
            lines[-1] = lines[-2][-1] + lines[-1]
            lines[-2] = lines[-2][:-1]
        else:
            break
    return lines


def _set_run(run, s, size, font, color, spacing=0, bold=False):
    run.text = s
    run.font.size = Pt(size * 0.75)
    run.font.name = font
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)
    rpr = run._r.get_or_add_rPr()
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(rpr, qn("a:ea"))
    ea.set("typeface", font)
    if spacing:
        rpr.set("spc", str(int(spacing * 0.75 * 100)))


def _textbox(slide, x, y, w, h, wrap=True):
    tb = slide.shapes.add_textbox(
        Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX))
    )
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    return tb


def _draw_line(slide, x, y, w, line, size, font, color, align, spacing, bold):
    lw = _measure(line, size, font)
    left = x - 0.42 * size if line and line[0] in OPEN_PUNCT else x
    if align == PP_ALIGN.CENTER:
        left = x + (w - lw) / 2.0
    elif align == PP_ALIGN.RIGHT:
        left = x + w - lw
    tb = _textbox(slide, left, y, max(w + size, lw + size * 0.2), size * 1.28, wrap=False)
    p = tb.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = Pt(size * 1.2 * 0.75)
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    r = p.add_run()
    _set_run(r, line, size, font, color, spacing, bold)


def text(slide, x, y, w, s, size=20, font=SONG, color=INK, leading=1.62,
         align=PP_ALIGN.LEFT, spacing=0, bold=False):
    lines = _layout_lines(s, size, w, font)
    for i, line in enumerate(lines):
        _draw_line(slide, x, y + i * size * leading, w, line, size, font, color, align, spacing, bold)
    return len(lines)


def paragraphs(slide, x, y, w, items, size=20, font=SONG, color=INK,
               leading=1.62, gap=10, align=PP_ALIGN.LEFT):
    for i, s in enumerate(items):
        lines = _layout_lines(s, size, w, font)
        for line in lines:
            _draw_line(slide, x, y, w, line, size, font, color, align, 0, False)
            y += size * leading
        if i < len(items) - 1:
            y += gap
    return y


def rect(slide, x, y, w, h, fill=INK, line=None, line_w=1):
    sp = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Emu(int(x * PX)), Emu(int(y * PX)),
        Emu(int(w * PX)), Emu(int(h * PX))
    )
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = _rgb(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = _rgb(line)
        sp.line.width = Emu(int(line_w * PX))
    sp.shadow.inherit = False
    return sp


def hline(slide, x1, y, x2, color=RULE, w=1):
    return rect(slide, x1, y, x2 - x1, w, fill=color)


def new_slide(bg=PAPER):
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = _rgb(bg)
    return s


def top_meta(slide, page, label, dark=False):
    fg = ROSE if dark else MUTED
    text(slide, 72, 34, 760, label, 10, SONG, fg, 1.2, spacing=2)
    text(slide, 1088, 34, 120, f"{page:02d} / 12", 10, SONG, fg, 1.2, PP_ALIGN.RIGHT)


def section_head(slide, page, part, title, dark=False):
    fg = WHITE if dark else INK
    rule = "#4A2A36" if dark else INK
    top_meta(slide, page, part, dark)
    text(slide, 72, 60, 960, title, 40, KAI, fg, 1.16)
    hline(slide, 72, 122, 1208, rule, 1)
    if not dark:
        rect(slide, 0, 0, 8, H, RED)


def footer(slide, page, dark=False, color=None):
    fg = color or (ROSE if dark else MUTED)
    text(slide, 72, 688, 620, "《留沪之前》 / 教育 · 科技 · 人才", 9, SONG, fg, 1.2, spacing=1)
    text(slide, 1088, 688, 120, f"{page:02d}", 9, SONG, fg, 1.2, PP_ALIGN.RIGHT)


def quote_page(page, kicker, quote, source, takeaway, bg, accent, big_char):
    s = new_slide(bg)
    top_meta(s, page, kicker, dark=True)
    text(s, 1050, 300, 190, big_char, 190, KAI, accent, 1.0, PP_ALIGN.RIGHT)
    lines = _layout_lines(quote, 36, 940, SONG)
    y = 176
    for line in lines:
        _draw_line(s, 90, y, 940, line, 36, SONG, WHITE, PP_ALIGN.LEFT, 0, False)
        y += 36 * 1.55
    hline(s, 90, 540, 1120, "#7A5A66", 1)
    text(s, 90, 560, 1000, source, 11, SONG, ROSE, 1.6)
    text(s, 90, 626, 1040, takeaway, 16, SONG, ROSE, 1.4)
    footer(s, page, dark=True, color=ROSE)
    return s


# ---------------------------------------------------------------------------
# 01 Cover
# ---------------------------------------------------------------------------
s = new_slide(PAPER)
rect(s, 0, 0, 520, H, RED)
rect(s, 520, 0, 14, H, PURPLE)
text(s, 60, 58, 380, "人生剧本 / LIFE SCRIPT", 9, SONG, ROSE, 1.2, spacing=3)
text(s, 60, 174, 410, "留沪之前", 76, KAI, WHITE, 1.12)
text(s, 60, 316, 400, "教育、科技、人才战略", 28, SONG, WHITE, 1.35)
text(s, 60, 374, 400, "2023版习概 · 第七专题", 10, SONG, ROSE, 1.4, spacing=1)
text(s, 60, 628, 400, "李华 / Charis / 李大爷", 12, SONG, ROSE, 1.4)

text(s, 600, 80, 570, "从社区夜课，到留在上海", 11, SONG, MUTED, 1.3, spacing=2)
text(s, 600, 156, 580, "教育是根本", 40, SONG, RED, 1.35)
text(s, 600, 214, 580, "科技是关键", 40, SONG, PURPLE, 1.35)
text(s, 600, 272, 580, "人才是基础", 40, SONG, INK, 1.35)
hline(s, 600, 352, 1160, RULE, 1)
text(s, 600, 382, 560, "留沪不是求一个位置，而是让自己成为这座城市需要的人。", 18, SONG, INK, 1.75)
rect(s, 948, 520, 260, 108, PURPLE)
text(s, 970, 542, 216, "从继续教育开始", 15, SONG, WHITE, 1.4)
text(s, 970, 568, 216, "从把手上的事做深开始", 15, SONG, ROSE, 1.4)
text(s, 980, 656, 220, "2026.10 / 12 PAGES", 9, SONG, MUTED, 1.2, PP_ALIGN.RIGHT, spacing=2)


# ---------------------------------------------------------------------------
# 02 人物与三幕
# ---------------------------------------------------------------------------
s = new_slide(PAPER)
section_head(s, 2, "第七专题 / 社会主义现代化建设的教育、科技、人才战略", "三个人，一条留沪的路")
people = [
    ("01", "李华", "归侨 / 准备留沪 / 继续教育学员", RED),
    ("02", "Charis", "Mr. C / 刚回国的朋友", PURPLE),
    ("03", "李大爷", "本名李明 / 退休工程师 / 夜课老学员", INK),
]
for i, (num, name, role, col) in enumerate(people):
    x = 72 + i * 376
    rect(s, x, 158, 336, 166, col)
    text(s, x + 22, 178, 60, num, 12, SONG, ROSE, 1.2, spacing=1)
    text(s, x + 22, 212, 280, name, 28, SONG, WHITE, 1.3)
    text(s, x + 22, 264, 292, role, 13, SONG, ROSE, 1.5)

acts = [
    ("第一幕 / 夜课", "教育政策的疑惑", "继续教育算不算教育？", RED),
    ("第二幕 / 球场", "技术变现的迷茫", "这些技术怎么变现？", PURPLE),
    ("第三幕 / 留沪", "人才引进的误会", "我是不是要先证明自己？", WINE),
]
for i, (head, sub, body, col) in enumerate(acts):
    x = 72 + i * 376
    rect(s, x, 376, 6, 6, col)
    text(s, x + 18, 366, 320, head, 17, KAI, col, 1.3)
    hline(s, x, 398, x + 336, RULE, 1)
    text(s, x, 416, 336, sub, 18, SONG, INK, 1.45)
    text(s, x, 452, 336, body, 13, SONG, MUTED, 1.6)

text(s, 72, 610, 1100, "一个社区夜课，把三个人的问题串成一条线。", 15, SONG, INK, 1.5)
footer(s, 2)


# ---------------------------------------------------------------------------
# 03 第一幕 夜课
# ---------------------------------------------------------------------------
s = new_slide(PAPER)
section_head(s, 3, "PART 01 / 第一幕", "夜课：继续教育算不算教育？")
p1 = [
    "社区继续教育中心的夜课，墙上贴着一张红色课程表，窗外的天色已经发紫。老师布置作业：把“社会主义现代化建设的教育、科技、人才战略”这一章，写成自己的故事。",
    "Charis 刚回国，坐在最后一排，低声问李华：“继续教育算不算教育？成人培训是不是低人一等？我回来以后，连政策都看不懂。”",
    "李华一时答不上来。",
]
paragraphs(s, 72, 156, 620, p1, 20, SONG, INK, 1.72, 12)

rect(s, 760, 156, 448, 430, RED)
text(s, 784, 178, 360, "NIGHT CLASS / 夜课", 10, SONG, ROSE, 1.2, spacing=2)
text(s, 784, 224, 380, "继续教育", 36, KAI, WHITE, 1.3)
text(s, 784, 276, 380, "算不算教育？", 36, KAI, WHITE, 1.3)
text(s, 784, 356, 380, "成人培训", 28, SONG, ROSE, 1.4)
text(s, 784, 392, 380, "是不是低人一等？", 28, SONG, ROSE, 1.4)
text(s, 1040, 470, 140, "？", 120, KAI, "#D9A3AC", 1.0, PP_ALIGN.RIGHT)

text(s, 72, 620, 1000, "问题没有答案，故事从这里开始。", 16, SONG, INK, 1.4)
footer(s, 3)


# ---------------------------------------------------------------------------
# 04 教育的根本问题
# ---------------------------------------------------------------------------
s = new_slide(PAPER)
section_head(s, 4, "PART 01 / 教育", "李华：这三个就是教育的根本问题")
cards = [
    ("01 培养什么人：", "德智体美劳全面发展的社会主义建设者和接班人", RED, WHITE, ROSE),
    ("02 怎样培养人：", "落实立德树人，五育并举，铸魂育人", PAPER, INK, PURPLE),
    ("03 为谁培养人：", "为党育人，为国育才，服务民族复兴", PURPLE, WHITE, LILAC),
]
for i, (head, body, bg, fg, accent) in enumerate(cards):
    x = 72 + i * 376
    rect(s, x, 170, 336, 334, bg)
    if i == 1:
        rect(s, x, 170, 336, 8, PURPLE)
    text(s, x + 22, 196, 284, head, 16, SONG, accent, 1.4)
    text(s, x + 22, 238, 292, body, 23, SONG, fg, 1.75)

text(s, 72, 560, 1100, "继续教育、职业教育、社区教育，都是把人培养起来。", 16, SONG, INK, 1.5)
rect(s, 72, 614, 1136, 48, INK)
text(s, 96, 628, 1040, "教育是根本：先把人培养起来，人才才会成批出现。", 15, SONG, WHITE, 1.3)
footer(s, 4)


# ---------------------------------------------------------------------------
# 05 引文·教育
# ---------------------------------------------------------------------------
quote_page(
    5,
    "PART 02 / 引文 · 教育",
    "中国这么多人，教育上去了，将来人才就会像井喷一样涌现出来。这是最有竞争力的。",
    "——习近平在广东考察工作时的讲话（2012年12月7—11日）",
    "教育是根本：把人培养起来，人才才会成批出现。",
    RED,
    "#7A1426",
    "教",
)


# ---------------------------------------------------------------------------
# 06 第二幕 球场
# ---------------------------------------------------------------------------
s = new_slide(PAPER)
section_head(s, 6, "PART 02 / 第二幕", "球场：技术怎么变成真正的价值")
p2 = [
    "周末，三个人在社区小馆看世预赛。屏幕上有实时数据、慢镜头回放。Charis 盯着看，忽然问：“这些技术是很先进，可它们怎么变现？我学的教育技术，回国以后能做什么？会不会一毕业就落后？”",
    "李华说：“技术如果只在屏幕里炫，就只是热闹。”",
    "李大爷指了指屏幕：“科技是关键，但关键不是追风口，是解决问题。”",
]
paragraphs(s, 72, 156, 620, p2, 20, SONG, INK, 1.72, 12)

rect(s, 760, 156, 448, 430, PURPLE)
text(s, 784, 178, 360, "LIVE DATA / 球场", 10, SONG, LILAC, 1.2, spacing=2)
for i, item in enumerate(["实时数据", "慢镜头回放", "教育技术"]):
    y = 232 + i * 70
    rect(s, 784, y, 8, 8, RED_LIGHT)
    text(s, 806, y - 8, 340, item, 24, SONG, WHITE, 1.4)
text(s, 1020, 470, 160, "？", 120, KAI, "#C9A4CB", 1.0, PP_ALIGN.RIGHT)
text(s, 784, 520, 360, "技术只有落到课堂、工厂和生活里，才叫真本事。", 13, SONG, LILAC, 1.65)

text(s, 72, 620, 1000, "技术如果只在屏幕里炫，就只是热闹。", 16, SONG, INK, 1.4)
footer(s, 6)


# ---------------------------------------------------------------------------
# 07 引文·科技
# ---------------------------------------------------------------------------
quote_page(
    7,
    "PART 02 / 引文 · 科技",
    "科技是国之利器，国家赖之以强，企业赖之以赢，人民生活赖之以好。中国要强，中国人民生活要好，必须有强大科技。",
    "——习近平在全国科技创新大会、两院院士大会、中国科协第九次全国代表大会上的讲话（2016年5月30日）",
    "科技是关键：不是追风口，是解决真问题。",
    PURPLE,
    "#2E1231",
    "科",
)


# ---------------------------------------------------------------------------
# 08 第三幕 留沪
# ---------------------------------------------------------------------------
s = new_slide(PAPER)
section_head(s, 8, "PART 03 / 第三幕", "留沪：我是不是要先证明自己？")
p3 = [
    "毕业后，李华想留在上海。他没有名校背景，也没有户口，一边准备社保材料，一边打听各种考核，越打听越觉得自己“自惭形秽”。",
    "他对李大爷说：“我是不是得先证明自己配得上留在这里？”",
]
paragraphs(s, 72, 156, 620, p3, 21, SONG, INK, 1.74, 14)

rect(s, 760, 156, 448, 430, ROSE)
text(s, 784, 178, 360, "ANXIETY / 留沪焦虑", 10, SONG, RED, 1.2, spacing=2)
items = ["名校背景", "户口", "社保材料", "各种考核"]
for i, item in enumerate(items):
    y = 228 + i * 66
    text(s, 784, y, 40, f"0{i+1}", 10, SONG, RED, 1.2)
    text(s, 830, y - 8, 320, item, 24, SONG, INK, 1.4)
    hline(s, 784, y + 42, 1180, "#D4B7BB", 1)
text(s, 784, 520, 360, "自惭形秽", 30, KAI, RED, 1.3)

rect(s, 72, 614, 1136, 50, INK)
text(s, 96, 628, 1040, "越打听，越觉得自己“自惭形秽”。", 16, SONG, WHITE, 1.3)
footer(s, 8)


# ---------------------------------------------------------------------------
# 09 人才引进的纠正
# ---------------------------------------------------------------------------
s = new_slide(PAPER)
section_head(s, 9, "PART 03 / 人才", "人才引进，不是一张卷子")
rect(s, 72, 170, 480, 360, ROSE)
text(s, 98, 196, 420, "误解 / MISREAD", 10, SONG, RED, 1.2, spacing=2)
text(s, 98, 234, 420, "留沪＝名校＋户口＋社保＋考核", 22, SONG, INK, 1.65)
text(s, 98, 330, 420, "先证明自己配得上留在这里。", 17, SONG, MUTED, 1.65)

rect(s, 600, 170, 608, 360, PURPLE)
text(s, 626, 196, 540, "纠正 / CORRECTION", 10, SONG, LILAC, 1.2, spacing=2)
text(s, 626, 234, 548, "社保是基本记录，不是判决书。", 22, SONG, WHITE, 1.6)
text(s, 626, 288, 548, "人才看能力、贡献、持续成长。", 20, SONG, WHITE, 1.6)
text(s, 626, 340, 548, "城市要的是能解决问题的人。", 20, SONG, ROSE, 1.6)

rect(s, 72, 580, 1136, 78, INK)
text(s, 96, 600, 1080, "李华：从继续教育开始，从把手上的事做深开始。", 18, SONG, WHITE, 1.3)
footer(s, 9)


# ---------------------------------------------------------------------------
# 10 引文·人才
# ---------------------------------------------------------------------------
quote_page(
    10,
    "PART 03 / 引文 · 人才",
    "战略人才站在国际科技前沿、引领科技自主创新、承担国家战略科技任务，是支撑我国高水平科技自立自强的重要力量，要把建设战略人才力量作为重中之重来抓。",
    "——习近平在中央人才工作会议上的讲话（2021年9月27日）",
    "人才是基础：战略人才力量，重中之重。",
    WINE,
    "#3A1228",
    "人",
)


# ---------------------------------------------------------------------------
# 11 三根支柱
# ---------------------------------------------------------------------------
s = new_slide(PAPER)
section_head(s, 11, "PART 04 / 合拢", "教育、科技、人才，三根支柱")
text(s, 72, 150, 1100, "教育、科技、人才是全面建设社会主义现代化国家的基础性、战略性支撑。", 22, SONG, INK, 1.55)
pillars = [
    ("教育是根本", "立德树人 / 为党育人、为国育才", RED),
    ("科技是关键", "自立自强 / 解决真问题", PURPLE),
    ("人才是基础", "战略人才力量 / 重中之重", INK),
]
for i, (head, body, col) in enumerate(pillars):
    x = 72 + i * 376
    rect(s, x, 244, 336, 286, col)
    text(s, x + 22, 272, 292, head, 28, KAI, WHITE, 1.35)
    text(s, x + 22, 328, 292, body, 17, SONG, ROSE, 1.7)

text(s, 72, 580, 1100, "教育是根，科技是钥匙，人才是走进去的那双脚。", 20, SONG, INK, 1.5)
rect(s, 72, 630, 1136, 34, RED)
text(s, 96, 638, 1080, "全面建设社会主义现代化国家，教育是根本，科技是关键，人才是基础。", 14, SONG, WHITE, 1.3)
footer(s, 11)


# ---------------------------------------------------------------------------
# 12 收束
# ---------------------------------------------------------------------------
s = new_slide(PURPLE)
rect(s, 0, 0, 260, H, RED)
text(s, 44, 52, 180, "END / 12", 9, SONG, ROSE, 1.2, spacing=2)
text(s, 44, 588, 180, "留沪之前", 26, KAI, WHITE, 1.3)
text(s, 44, 630, 180, "2023版习概 · 第七专题", 8, SONG, ROSE, 1.4)

text(s, 320, 116, 880, "留沪不是求一个位置，", 34, SONG, WHITE, 1.55)
text(s, 320, 176, 880, "而是让自己成为这座城市需要的人。", 34, SONG, WHITE, 1.55)
hline(s, 320, 300, 1160, "#7A5A82", 1)
text(s, 320, 338, 840, "教育是根，科技是钥匙，", 22, SONG, ROSE, 1.5)
text(s, 320, 378, 840, "人才是走进去的那双脚。", 22, SONG, ROSE, 1.5)
text(s, 320, 512, 800, "从继续教育开始，从把手上的事做深开始。", 16, SONG, LILAC, 1.5)
text(s, 880, 648, 280, "《留沪之前》 / 李华 · Charis · 李大爷", 9, SONG, LILAC, 1.2, PP_ALIGN.RIGHT)
footer(s, 12, dark=True, color=LILAC)


# ---------------------------------------------------------------------------
# Notes, flatness, save
# ---------------------------------------------------------------------------
notes = {
    1: "封面。用《留沪之前》建立人物和问题。",
    2: "人物与三幕。交代李华、Charis、李大爷，以及夜课、球场、留沪三个场景。",
    3: "第一幕。Charis 问继续教育算不算教育，李华一时答不上来。",
    4: "教育的根本问题。培养什么人、怎样培养人、为谁培养人。",
    5: "引文·教育。2012 年广东考察讲话，落在人才井喷。",
    6: "第二幕。世预赛转播里的技术，引出变现与解决问题的区别。",
    7: "引文·科技。2016 年科技创新大会讲话，科技是国之利器。",
    8: "第三幕。李华留沪焦虑，自惭形秽，准备社保材料和各种考核。",
    9: "李大爷纠正。人才引进不是一张卷子，看能力、贡献、持续成长。",
    10: "引文·人才。2021 年中央人才工作会议讲话，战略人才力量是重中之重。",
    11: "三根支柱。教育是根本，科技是关键，人才是基础。",
    12: "收束。留沪不是求一个位置，而是成为城市需要的人。",
}
for idx, slide in enumerate(prs.slides, start=1):
    slide.notes_slide.notes_text_frame.text = notes[idx]

for slide in prs.slides:
    for shape in slide.shapes:
        style = shape._element.find(qn("p:style"))
        if style is not None:
            shape._element.remove(style)
        sp_pr = shape._element.find(qn("p:spPr"))
        if sp_pr is not None and sp_pr.find(qn("a:effectLst")) is None:
            etree.SubElement(sp_pr, qn("a:effectLst"))

prs.save(str(OUT))
print(f"saved {len(prs.slides)} slides -> {OUT}")

