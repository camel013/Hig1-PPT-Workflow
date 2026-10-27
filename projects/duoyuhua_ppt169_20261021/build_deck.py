# -*- coding: utf-8 -*-
import pathlib
from xml.sax.saxutils import escape

OUT = pathlib.Path(r"C:\Users\camel\Documents\ChatGPT\Hig1\projects\duoyuhua_ppt169_20261021\svg_output")
OUT.mkdir(parents=True, exist_ok=True)

W, H = 1280, 720
M = 80
RIGHT = W - M
INK = "#111111"
MUTED = "#767674"
RULE = "#D9D9D4"
ACCENT = "#E2231A"
BG = "#FFFFFF"

KAI = "华文楷体"
DENG = "等线"
SONG = "华文中宋"
EN = "Arial"

def esc(s):
    return escape(str(s))

def t(x, y, s, size=22, font="body", fill=INK, anchor="start", ls=0, bold=False, opacity=1.0):
    if font == "title":
        family, weight = KAI, 400
    elif font == "sub":
        family, weight = DENG, 700
    elif font == "body":
        family, weight = SONG, 400
    else:
        family, weight = EN, 400
    return (f'<text x="{x}" y="{y}" font-family="{family}" font-size="{size}" '
            f'font-weight="{weight}" fill="{fill}" text-anchor="{anchor}" '
            f'letter-spacing="{ls}" opacity="{opacity}">{esc(s)}</text>')

def rule(x1, y1, x2, y2, color=RULE, w=1):
    return f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{color}" stroke-width="{w}"/>'

def rect(x, y, w_, h_, fill=INK, opacity=1.0):
    return f'<rect x="{x}" y="{y}" width="{w_}" height="{h_}" fill="{fill}" opacity="{opacity}"/>'

def _flat(items):
    for it in items:
        if isinstance(it, list):
            yield from _flat(it)
        else:
            yield it

def page(inner, bg=BG):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
            f'viewBox="0 0 {W} {H}">\n' + rect(0, 0, W, H, bg) + "\n" + "\n".join(_flat(inner)) + "\n</svg>\n")

def head(no, kicker, title=None, title_size=54, title_y=176):
    out = []
    out.append(t(M, 74, f"{no} / 12", 13, "en", MUTED, ls=2))
    out.append(t(RIGHT, 74, kicker, 13, "en", MUTED, anchor="end", ls=2))
    out.append(rule(M, 96, RIGHT, 96))
    if title:
        out.append(t(M, title_y, title, title_size, "title"))
    return out

def foot(no, label="多余的话"):
    out = [rule(M, 664, RIGHT, 664)]
    out.append(t(M, 692, label, 13, "body", MUTED, ls=1))
    out.append(t(RIGHT, 692, f"{no:02d} / 12", 13, "en", MUTED, anchor="end", ls=2))
    return out

slides = []

# 01 COVER
s = []
s.append(rect(0, 0, 10, H, ACCENT))
s.append(t(M, 78, "文献 / 01", 13, "sub", MUTED, ls=2))
s.append(t(RIGHT, 78, "1935", 13, "en", MUTED, anchor="end", ls=2))
s.append(rule(M, 96, RIGHT, 96))
s.append(t(M, 300, "多余的话", 132, "title"))
s.append(rule(M, 348, RIGHT, 348, INK, 2))
s.append(t(M, 420, "瞿秋白", 40, "sub"))
s.append(t(M, 466, "一九三五年五月 · 汀州狱中", 22, "body", MUTED))
s.append(t(M, 604, "SUPERFLUOUS WORDS", 13, "en", MUTED, ls=3))
s.append(t(M, 630, "QU QIUBAI   1899—1935", 13, "en", MUTED, ls=3))
s.append(t(RIGHT, 630, "最后的坦白", 14, "sub", ACCENT, anchor="end", ls=1))
slides.append(("slide_01_cover", s))

# 02 EPIGRAPH
s = []
s.append(t(M, 78, "02 / 12", 13, "en", MUTED, ls=2))
s.append(t(RIGHT, 78, "PREFACE", 13, "en", MUTED, anchor="end", ls=2))
s.append(rule(M, 96, RIGHT, 96))
s.append(t(M, 268, "知我者，谓我心忧；", 58, "title"))
s.append(t(M, 348, "不知我者，谓我何求。", 58, "title"))
s.append(t(M, 404, "——《诗经·黍离》", 18, "body", MUTED))
s.append(rule(M, 470, 620, 470, ACCENT, 3))
s.append(t(M, 552, "何必说？", 76, "title"))
s.append(t(M, 596, "代序", 20, "sub", MUTED))
slides.append(("slide_02_epigraph", s))

# 03 PREFACE
s = []
s.append(head("03", "PREFACE / 代序"))
s.append(t(62, 542, "03", 238, "en", "#ECECE8"))
s.append(t(430, 210, "话既然是多余的，", 44, "title"))
s.append(t(430, 268, "又何必说呢？", 44, "title"))
s.append(rule(430, 300, 700, 300, ACCENT, 3))
s.append(t(430, 368, "已经是走到了生命的尽期。余剩的日子，", 22, "body"))
s.append(t(430, 404, "不但不能按照年份来算，", 22, "body"))
s.append(t(430, 440, "甚至不能按星期来算了。", 22, "body"))
s.append(t(430, 512, "这是我最后“谈天”的机会呢！", 20, "body", MUTED))
s.append(t(430, 544, "（一九三五·五·一七于汀州狱中）", 15, "body", MUTED))
s.append(foot(3))
slides.append(("slide_03_preface", s))

# 04 HISTORICAL MISUNDERSTANDING
s = []
s.append(head("04", "THE HISTORICAL MISUNDERSTANDING", "历史的误会"))
s.append(t(M, 272, "我其实是一个很平凡的文人，", 40, "title"))
s.append(t(M, 336, "竟虚负了某某党的领袖的名声十来年。", 40, "title"))
s.append(t(M, 418, "这不是“历史的误会”，是什么呢？", 40, "title"))
s.append(rule(M, 470, RIGHT, 470))
s.append(t(M, 540, "15", 84, "en", ACCENT))
s.append(t(M + 118, 528, "年", 20, "sub", MUTED))
s.append(t(M + 118, 556, "勉强做着政治工作", 16, "body", MUTED))
s.append(t(680, 540, "10", 84, "en", ACCENT))
s.append(t(798, 528, "余年", 20, "sub", MUTED))
s.append(t(798, 556, "虚负领袖之名", 16, "body", MUTED))
s.append(foot(4))
slides.append(("slide_04_misunderstanding", s))

# 05 TIMELINE
rows = [
    ("1899", "出生于江苏武进（今常州）"),
    ("1917", "入北京俄文专修馆学习"),
    ("1919", "参加五四运动，与郑振铎等创办《新社会》"),
    ("1920", "任《晨报》特派记者，赴苏俄"),
    ("1922", "在莫斯科加入中国共产党"),
    ("1927", "主持中央政治局，担负直接政治领导"),
    ("1931", "被解除中央领导职务，转入左翼文化运动"),
    ("1935", "六月十八日，就义于福建长汀中山公园"),
]
s = []
s.append(head("05", "CHRONOLOGY", "年表"))
s.append(rule(300, 200, 300, 620))
y = 232
for i, (yr, ev) in enumerate(rows):
    color = ACCENT if yr == "1935" else INK
    s.append(t(268, y, yr, 20, "en", color, anchor="end"))
    s.append(f'<circle cx="300" cy="{y-7}" r="{5 if yr=="1935" else 3}" fill="{color}"/>')
    s.append(t(340, y, ev, 20, "body", INK if yr != "1935" else ACCENT))
    y += 52
s.append(foot(5))
slides.append(("slide_05_chronology", s))

# 06 THE TWO SELVES
s = []
s.append(head("06", "THE FRAIL DUAL SELF", "脆弱的二元人物"))
s.append(rect(M, 226, 4, 150, ACCENT))
s.append(t(M + 28, 268, "一只羸弱的马拖着几千斤的辎重车，", 34, "title"))
s.append(t(M + 28, 316, "走上了险峻的山坡，一步步地往上爬。", 34, "title"))
s.append(t(M + 28, 364, "要往后退是不可能，要再往前去是实在不能胜任了。", 22, "body", MUTED))
s.append(rule(M, 470, RIGHT, 470))
s.append(t(M, 590, "36", 110, "en", ACCENT))
s.append(t(M + 150, 566, "岁", 22, "sub", MUTED))
s.append(t(M + 150, 596, "自觉已经非常地衰惫，丝毫青年壮年的兴趣都没有了", 18, "body", MUTED))
s.append(foot(6))
slides.append(("slide_06_dual", s))

# 07 MARXISM
s = []
s.append(head("07", "MARXISM / 1923—1927", "我和马克思主义"))
s.append(t(M, 262, "我对这些学问，", 40, "title"))
s.append(t(M, 322, "的确只知道一点皮毛。", 40, "title"))
s.append(t(M, 396, "当时我只是根据几本外国文的书籍转译一下，编了一些讲义。", 18, "body", MUTED))
s.append(rule(M, 452, RIGHT, 452))
s.append(t(M, 524, "从那时候起，我没有自己的思想。", 32, "sub"))
s.append(t(M, 584, "（我以中央的思想为思想。）", 34, "sub", ACCENT))
s.append(foot(7))
slides.append(("slide_07_marxism", s))

# 08 LEADERSHIP
s = []
s.append(head("08", "POLITICAL LEADERSHIP", "五年"))
s.append(t(M, 300, "1927", 28, "en", MUTED))
s.append(t(RIGHT, 300, "1931", 28, "en", MUTED, anchor="end"))
s.append(rule(M, 318, RIGHT, 318, INK, 2))
for i in range(5):
    x = M + i * (RIGHT - M) / 4
    s.append(rule(x, 312, x, 326, INK, 2))
s.append(t(M, 426, "整整五年，我居然当了中国共产党的领袖之一，", 30, "sub"))
s.append(t(M, 472, "最后三年甚至仿佛是最主要的领袖。", 30, "sub"))
s.append(t(M, 552, "“你们饶了我吧，我实在没有兴趣和能力负担这个领导工作了。”", 19, "body", MUTED))
s.append(t(M, 584, "但是，终于没有说出口。", 19, "body", MUTED))
s.append(foot(8))
slides.append(("slide_08_leadership", s))

# 09 BLIND ACTIONISM
s = []
s.append(head("09", "BLIND ACTIONISM / LI LISAN LINE", "盲动主义和立三路线"))
s.append(t(M, 266, "立三路线是我的许多错误观点", 34, "title"))
s.append(t(M, 310, "——有人说是瞿秋白主义——逻辑的发展。", 34, "title"))
terms = [("盲动主义", "一九二八年初的判断错误"), ("立三路线", "左倾空谈掩盖右倾实质"), ("瞿秋白主义", "他人加之于我的命名")]
y = 424
for name, desc in terms:
    s.append(rule(M, y - 42, RIGHT, y - 42))
    s.append(t(M, y, name, 26, "sub"))
    s.append(t(520, y, desc, 18, "body", MUTED))
    y += 78
s.append(t(M, 636, "我当然间接地负着立三路线的责任。", 18, "sub", ACCENT))
s.append(foot(9))
slides.append(("slide_09_blind", s))

# 10 LITERATI
s = []
s.append(head("10", "THE LITERATI", None))
s.append(t(M, 210, "“文人”", 108, "title"))
s.append(rect(M, 266, 330, 5, ACCENT))
s.append(t(M, 352, "文人积习未除。", 28, "sub"))
s.append(t(M, 392, "至多不过是个“读者”。", 30, "sub"))
s.append(t(M, 446, "假使能够仔细而郑重地，极忠实地翻译几部俄国文学名著，", 18, "body", MUTED))
s.append(t(M, 474, "在汉字方面每字每句地斟酌着，也许不会“误人子弟”的。", 18, "body", MUTED))
s.append(rule(M, 532, RIGHT, 532))
s.append(t(M, 580, "这一个最愉快的梦想，比创作和评论方面再求得什么成就，要实际得多。", 18, "body"))
s.append(t(M, 614, "可惜，恐怕现在这个可能已经“过时”了！", 18, "body", MUTED))
s.append(foot(10))
slides.append(("slide_10_literati", s))

# 11 FAREWELL
s = []
s.append(head("11", "FAREWELL", "告别"))
items = [("一出滑稽剧就此闭幕了！", 30, "body", INK), ("我只要休息。", 52, "sub", INK), ("永别了，美丽的世界！", 64, "title", ACCENT)]
y = 262
for txt, size, font, color in items:
    s.append(t(M, y, txt, size, font, color))
    y += 120
    s.append(rule(M, y - 62, RIGHT, y - 62))
s.append(t(M, 640, "一九三五·五·二三", 15, "body", MUTED, ls=2))
s.append(foot(11))
slides.append(("slide_11_farewell", s))

# 12 END (inverted)
s = []
s.append(t(M, 90, "12 / 12", 13, "en", "#8A8A86", ls=2))
s.append(t(RIGHT, 90, "THE END", 13, "en", "#8A8A86", anchor="end", ls=2))
s.append(rule(M, 112, RIGHT, 112, "#3A3A38"))
s.append(t(M, 300, "中国的豆腐也是很好吃的东西，", 40, "title", "#F2F2EE"))
s.append(t(M, 352, "世界第一。", 40, "title", "#F2F2EE"))
s.append(t(M, 470, "永别了！", 96, "title", ACCENT))
s.append(rule(M, 540, RIGHT, 540, "#3A3A38"))
s.append(t(M, 586, "瞿秋白 · 一八九九—一九三五", 18, "body", "#C8C8C4"))
s.append(t(M, 620, "一九三五年六月十八日 · 福建长汀中山公园", 15, "body", "#8A8A86"))
s.append(foot(12))
slides.append(("slide_12_end", s))

for name, body in slides:
    svg = page(body, "#111111") if name == "slide_12_end" else page(body)
    (OUT / f"{name}.svg").write_text(svg, encoding="utf-8")
print(f"wrote {len(slides)} svg files")