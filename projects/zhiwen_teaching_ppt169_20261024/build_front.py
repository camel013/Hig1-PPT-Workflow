# -*- coding: utf-8 -*-
"""Build the 90-page Zhiwen teaching deck (native, editable PPTX)."""
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).parent))
import deck_lib as D
from deck_lib import (
                      FANG, KAI, PAPER, PAPER_WARM, WHITE, INK, MUTED, RULE, CORAL, CORAL_DARK,
                      CORAL_LIGHT, INDIGO, INDIGO_DEEP, NAVY, TEAL, TEAL_DARK, GOLD,
                      GOLD_LIGHT, SONG, TNR, PP_ALIGN)

HEAD = FANG
P = pathlib.Path(__file__).parent
TOTAL = 90
page = 0


def nxt():
    global page
    page += 1
    return page


def cover():
    p = nxt()
    s = D.new_slide(PAPER)
    D.rect(s, 0, 0, 760, D.H, fill=CORAL)
    D.rect(s, 760, 0, 520, D.H, fill=INDIGO)
    D.rect(s, 860, 0, 420, D.H, fill=INDIGO_DEEP)
    D.rect(s, 760, 0, 100, D.H, fill=NAVY)
    for k in range(3):
        pts = []
        y0 = 42 + k * 15
        for i in range(0, 220, 20):
            pts.append((40 + i, y0 + (0 if (i // 20) % 2 == 0 else -9)))
            pts.append((40 + i + 10, y0 + (-9 if (i // 20) % 2 == 0 else 0)))
        D.freeform(s, pts, color=CORAL_LIGHT, width=1.0)
    D.text(s, 64, 94, 520, "高中语文 · 文本细读与写作", 12, SONG, WHITE, 1.4, spacing=2)
    D.vline(s, 64, 156, "一本假课本的读法", 34, WHITE, HEAD, 1.3, spacing=2)
    D.text(s, 64, 596, 540, "从戏仿到见证：语言如何成为证词", 18, SONG, WHITE, 1.4)
    D.text(s, 64, 640, 540, "4 个单元 · 24 篇课文 · 3 首古诗词 · 1 篇跋", 11, TNR, CORAL_LIGHT, 1.3)
    D.rect(s, 930, 70, 286, 360, fill=WHITE)
    D.text(s, 958, 84, 240, "蛭", 96, KAI, INK, 1.0)
    D.text(s, 958, 186, 240, "文", 96, KAI, INK, 1.0)
    D.text(s, 962, 302, 220, "ZHIWEN", 11, TNR, MUTED, 1.3, spacing=3)
    D.text(s, 962, 336, 220, "特殊高中教科书", 9, SONG, MUTED, 1.3, spacing=1.5)
    D.rect(s, 930, 430, 286, 58, fill=GOLD)
    D.text(s, 958, 444, 230, "不必修 · 中册", 17, SONG, WHITE, 1.2, spacing=1.2)
    D.text(s, 1128, 110, 64, "2026", 10, TNR, CORAL_DARK, 1.0, PP_ALIGN.CENTER)
    D.text(s, 1120, 690, 120, "01", 9, TNR, CORAL_LIGHT, 1.2, PP_ALIGN.RIGHT)


def contract():
    p = nxt()
    s = D.new_slide(PAPER)
    D.page_title(s, p, "导读 / 阅读契约", "先把“假”放下")
    D.text(s, 64, 160, 600, "这不是一份需要背诵的教材，而是一部由学生写成的戏仿文集。它借课本的体例、经典的句法、庄重的文言，记录一所学校里被反复讲述的日常。", 17, SONG, INK, 1.8)
    rules = [
        ("01", "不作人身指认", "文中的“徐规尔”“水蛭”是叙述者建构的讽刺形象。课堂讨论针对写法与经验，不针对现实个人。"),
        ("02", "不把戏谑当事实", "夸张、传闻、诅咒是文本策略，不等于证据。读它，要同时读“说了什么”和“为何这样说”。"),
        ("03", "不只看热闹", "笑过之后，追问它为什么笑、笑给谁看、笑完剩下什么。"),
    ]
    y = 310
    for no, head, body in rules:
        D.text(s, 64, y - 4, 44, no, 23, TNR, CORAL, 1.0)
        D.text(s, 118, y - 6, 250, head, 20, HEAD, INK, 1.2)
        D.text(s, 376, y - 2, 410, body, 13.5, SONG, INK, 1.65)
        D.hline(s, 64, 800, y + 66, RULE, 0.8)
        y += 94
    D.rect(s, 860, 150, 356, 430, fill=CORAL)
    D.text(s, 892, 190, 290, "“语言是思维的边界，语言是创造的源头。”", 28, HEAD, WHITE, 1.55)
    D.hline(s, 892, 1180, 388, CORAL_LIGHT, 1)
    D.text(s, 892, 414, 290, "—— 编者的话", 11, SONG, CORAL_LIGHT, 1.4)
    D.text(s, 892, 470, 290, "把这句话放在第一页，是因为整本书的褒贬都从这里出发：语言究竟用来连接人，还是用来设门槛。", 13.5, SONG, WHITE, 1.75)
    D.footer(s, p)


def book_map():
    p = nxt()
    s = D.new_slide(PAPER)
    D.page_title(s, p, "导读 / 全书结构", "一本课本的四个声部")
    D.text(s, 64, 156, 740, "四个单元不是四种题材，而是四种“借壳”的方法：借古诗的腔、借小说的镜子、借古文的礼、借演讲的火。", 17, SONG, INK, 1.75)
    data = [
        ("壹", "古典诗歌的借壳", "4 篇", "歌行 · 散曲 · 新月抒情 · 现代短诗", GOLD, "pp. 1—10"),
        ("贰", "现当代叙事", "8 篇", "小说 · 书信 · 情诗 · 片段 · 寓言 · 戏剧", INDIGO, "pp. 11—44"),
        ("叁", "文言戏仿", "6 篇", "史传 · 小品 · 语录 · 侍坐 · 列传", TEAL, "pp. 45—62"),
        ("肆", "见证文体", "8 篇", "杂文 · 散文诗 · 回忆 · 讲话 · 纪实", CORAL, "pp. 63—109"),
    ]
    x = 64
    for i, (no, title, count, genres, col, rng) in enumerate(data):
        D.rect(s, x, 280, 272, 300, fill=col)
        D.text(s, x + 24, 300, 150, no, 70, HEAD, WHITE, 1.0)
        D.text(s, x + 24, 390, 220, title, 21, HEAD, WHITE, 1.35)
        D.text(s, x + 24, 444, 220, genres, 11.5, SONG, WHITE, 1.65)
        D.hline(s, x + 24, x + 248, 508, WHITE, 0.8)
        D.text(s, x + 24, 522, 100, count, 15, SONG, WHITE, 1.2)
        D.text(s, x + 168, 524, 100, rng, 9.5, TNR, WHITE, 1.2, PP_ALIGN.RIGHT)
        x += 288
    D.text(s, 64, 610, 500, "另：古诗词诵读 3 首 · 跋 1 篇 · 制度时间线 1 页", 12, SONG, MUTED, 1.4)
    D.text(s, 900, 610, 316, "24 篇 + 3 首 + 1 跋", 18, TNR, CORAL, 1.2, PP_ALIGN.RIGHT)
    D.footer(s, p)


def editor_note():
    p = nxt()
    s = D.new_slide(NAVY)
    D.meta(s, p, "导读 / 编者的话", dark=True)
    D.text(s, 64, 150, 920, "“当语言失掉了交流与创造的灵魂，", 36, HEAD, WHITE, 1.5)
    D.text(s, 64, 222, 920, "成为了设置门槛的工具，那么语言就死了。”", 36, HEAD, WHITE, 1.5)
    D.hline(s, 64, 1010, 360, CORAL, 2)
    D.text(s, 64, 392, 880, "编者把“语言”放在两种命运之间：交流与创造，或者门槛与工具。整本书的褒贬由此展开——它反对的不是某一个人，而是一种让语言失去生命的使用方式。", 17, SONG, GOLD_LIGHT, 1.8)
    for i, (h, b) in enumerate([("活尸", "苍白、浮肿、只剩腔调"), ("荒诞", "与日常生活互为镜像"), ("生命力量", "嬉笑怒骂，皆成文章")]):
        x = 64 + i * 300
        D.text(s, x, 520, 240, h, 22, HEAD, CORAL_LIGHT, 1.2)
        D.text(s, x, 558, 240, b, 12.5, SONG, WHITE, 1.5)
    D.text(s, 64, 648, 900, "法农：人是宇宙和声中一个响亮有力的“是”字。", 13, SONG, TEAL, 1.4)
    D.footer(s, p, dark=True)


def name_decode():
    p = nxt()
    s = D.new_slide(PAPER)
    D.meta(s, p, "导读 / 书名", )
    D.text(s, 64, 46, 600, "三个字，三种读法", 34, HEAD, INK, 1.2)
    D.text(s, 64, 118, 1160, "书名不是标签，而是第一个讽刺动作：把最庄重的“文”，安在最软、最贪的“蛭”后面。", 16, SONG, MUTED, 1.4)
    # giant characters
    D.text(s, 80, 190, 300, "蛭", 250, KAI, CORAL, 1.0)
    D.text(s, 380, 190, 300, "文", 250, KAI, INDIGO, 1.0)
    readings = [
        ("读法一", "水蛭写的文", "叙述者模拟“加害者”的口吻，让权力自己留下证词。"),
        ("读法二", "写水蛭的文", "学生把同一个人反复写进不同文体，形成群像与档案。"),
        ("读法三", "被水蛭吸过的文", "语言被公文、训斥、监控词侵蚀，留下疲惫的痕迹。"),
    ]
    y = 480
    for i, (k, h, b) in enumerate(readings):
        x = 80 + i * 390
        D.hline(s, x, x + 330, y, RULE, 0.9)
        D.text(s, x, y + 10, 330, k, 10, SONG, MUTED, 1.2, spacing=1.5)
        D.text(s, x, y + 38, 330, h, 21, HEAD, INK, 1.2)
        D.text(s, x, y + 78, 330, b, 12.5, SONG, INK, 1.6)
    D.footer(s, p)


def palette_page():
    p = nxt()
    s = D.new_slide(PAPER)
    D.page_title(s, p, "导读 / 视觉系统", "从封面上取下来的六种情绪")
    D.text(s, 64, 156, 720, "封面用珊瑚红与靛蓝对半切开，中间夹着一条近黑的竖缝。它不是“配色好看”，而是先把“人”和“制度”分成两个区。", 16, SONG, INK, 1.75)
    colors = [
        ("D56A58", "珊瑚砖红", "体温 · 愤怒 · 集体证词"),
        ("1B4F8B", "靛蓝", "制度 · 监控 · 夜"),
        ("649788", "灰绿", "古典 · 距离 · 反讽"),
        ("B09263", "陈金", "记忆 · 尊严 · 旧纸"),
        ("F3EDE4", "纸白", "留白 · 阅读 · 呼吸"),
        ("0C182B", "墨蓝", "不可见的权力"),
    ]
    for i, (h, name, mean) in enumerate(colors):
        x = 64 + (i % 3) * 390
        y = 290 + (i // 3) * 170
        D.rect(s, x, y, 110, 110, fill=h, line=(RULE if h == "F3EDE4" else None))
        D.text(s, x + 128, y + 6, 220, name, 19, HEAD, INK, 1.2)
        D.text(s, x + 128, y + 42, 240, mean, 12, SONG, MUTED, 1.5)
        D.text(s, x + 128, y + 72, 200, "#" + h, 9.5, TNR, MUTED, 1.2)
    D.hline(s, 64, 1216, 636, RULE, 0.8)
    D.text(s, 64, 652, 900, "用法：单元一陈金，单元二靛蓝，单元三灰绿，单元四珊瑚。黑色只做结构与句号。", 11.5, SONG, INK, 1.4)
    D.footer(s, p)


def five_layers():
    p = nxt()
    s = D.new_slide(PAPER)
    D.page_title(s, p, "导读 / 方法", "戏仿的五个层次")
    D.text(s, 64, 154, 820, "戏仿不是模仿。它借来一个熟悉的壳，把里面的心换成此刻的现实；读者先认出壳，再被里面的东西刺一下。", 16.5, SONG, INK, 1.75)
    items = [
        ("01", "体式", "歌行、散曲、史传、语录、演讲……"),
        ("02", "语体", "文言、口语、公文、网络流行语互相摩擦。"),
        ("03", "人物", "真名、诨名、类型名，在文本里被重新分配。"),
        ("04", "场景", "教室、办公室、操场、寝室、教务处。"),
        ("05", "语气", "庄严、戏谑、悲愤、冷漠，随时换挡。"),
    ]
    for i, (no, h, b) in enumerate(items):
        x = 64 + i * 232
        D.text(s, x, 300, 60, no, 28, TNR, CORAL, 1.0)
        D.hline(s, x, x + 196, 342, RULE, 0.9)
        D.text(s, x, 358, 196, h, 24, HEAD, INK, 1.2)
        D.text(s, x, 402, 196, b, 12.5, SONG, INK, 1.65)
    D.rect(s, 64, 520, 1152, 92, fill=INK)
    D.text(s, 92, 542, 1090, "借壳换心：壳越像，心越疼。", 30, HEAD, WHITE, 1.2)
    D.text(s, 900, 558, 290, "这就是本书的基本语法", 11, SONG, GOLD_LIGHT, 1.2, PP_ALIGN.RIGHT)
    D.footer(s, p)


def toolbox():
    p = nxt()
    s = D.new_slide(PAPER_WARM)
    D.page_title(s, p, "导读 / 工具箱", "读这本书，需要五把刀")
    tools = [
        ("互文", "先找到“底文”，才知道它改了什么。", "《琵琶行》→《蛭歌行》"),
        ("反讽", "字面与真意相反，语气越正经越可疑。", "“务必落实”"),
        ("不可靠叙述", "说话的人未必可信，他的漏洞就是证据。", "“我”是水蛭"),
        ("象征", "一个词从具体人物，变成一套制度。", "蛭 / 笼 / 监控"),
        ("复调", "多种声音互相拆台，没有唯一结论。", "学生、老师、校长"),
    ]
    for i, (h, b, ex) in enumerate(tools):
        y = 160 + i * 102
        D.text(s, 64, y, 44, f"0{i+1}", 18, TNR, CORAL, 1.0)
        D.text(s, 126, y - 3, 180, h, 22, HEAD, INK, 1.2)
        D.text(s, 330, y, 520, b, 14, SONG, INK, 1.55)
        D.text(s, 900, y + 2, 310, ex, 11.5, SONG, MUTED, 1.4, PP_ALIGN.RIGHT)
        D.hline(s, 64, 1210, y + 78, RULE, 0.7)
    D.text(s, 64, 620, 900, "读法：每读一篇，至少指出一次“它借了谁”，一次“它把谁刺痛了”。", 11.5, SONG, MUTED, 1.3)
    D.footer(s, p)


def emotion_curve():
    p = nxt()
    s = D.new_slide(PAPER)
    D.page_title(s, p, "导读 / 情绪结构", "越往后，笑越少，痛越具体")
    # axes
    D.hline(s, 100, 1160, 560, INK, 1.0)
    D.vrule(s, 100, 210, 560, INK, 1.0)
    D.text(s, 64, 548, 40, "轻", 11, SONG, MUTED, 1.0)
    D.text(s, 64, 202, 40, "重", 11, SONG, MUTED, 1.0)
    pts = [(180, 340), (420, 310), (660, 380), (900, 300), (1090, 230)]
    D.freeform(s, pts, color=CORAL, width=2.2)
    labels = [("01", "讽喻", "金色"), ("02", "冷峻", "靛蓝"), ("03", "荒诞", "灰绿"), ("04", "灼热", "珊瑚"), ("跋", "克制", "墨蓝")]
    for (x, y), (no, name, col) in zip(pts, labels):
        D.rect(s, x - 5, y - 5, 10, 10, fill=CORAL)
        D.text(s, x - 20, y - 44, 80, no, 11, TNR, CORAL, 1.0)
        D.text(s, x - 30, y + 18, 100, name, 14, HEAD, INK, 1.2)
        D.text(s, x - 30, y + 42, 100, col, 9, SONG, MUTED, 1.2)
    D.text(s, 140, 610, 600, "四个单元的情绪不是平铺，而是逐级加压；到第四单元，讽刺变成证词。", 12.5, SONG, MUTED, 1.4)
    D.footer(s, p)


def toc():
    p = nxt()
    s = D.new_slide(NAVY)
    D.meta(s, p, "导读 / 目录", dark=True)
    D.text(s, 64, 56, 400, "目录", 44, HEAD, WHITE, 1.2)
    D.text(s, 64, 116, 600, "从“怎么读”到“怎么写”", 13, SONG, GOLD_LIGHT, 1.4)
    cols = [
        ("导读", GOLD_LIGHT, [("阅读契约", "02"), ("全书结构", "03"), ("编者的话", "04"), ("书名与色板", "05—06"), ("方法与情绪", "07—09")]),
        ("壹 · 以古写今", GOLD, [("古典诗歌的借壳", "11—12"), ("蛭歌行", "13—16"), ("朝天子", "17—18"), ("再别三班", "19—20"), ("偶然 / 小结", "21—23")]),
        ("贰 · 规则的冷光", INDIGO, [("现当代叙事", "24—25"), ("装在规则里的蛭", "26—27"), ("我的辞蛭", "28—29"), ("情诗 / 东区", "30—33"), ("化学教员 / 寻找", "34—37"), ("蛭影 / 三班内外", "38—41"), ("小结", "42")]),
        ("叁 · 庄辞写荒诞", TEAL, [("文言戏仿", "43—44"), ("徐规尔传", "45—46"), ("朱夏风致记", "47—48"), ("蛭语 / 侍坐", "49—52"), ("随之 / 蛭", "53—54"), ("小结与任务", "55—56")]),
        ("肆 · 字字泣血", CORAL, [("见证文体", "57—58"), ("反对蛭八股", "59—61"), ("岁月 / 蛭事", "62—67"), ("小传 / 讲演", "68—74"), ("大地 / 高三牲", "75—78"), ("小结", "79")]),
        ("尾声", TEAL, [("古诗词诵读", "80—81"), ("跋 / 时间线", "82—83"), ("复盘 / 母题", "84—85"), ("戏仿伦理", "86"), ("写作与讨论", "87—88"), ("结语", "89—90")]),
    ]
    for i, (head, col, items) in enumerate(cols):
        x = 64 + (i % 3) * 390
        y = 180 + (i // 3) * 250
        D.text(s, x, y, 330, head, 19, HEAD, col, 1.2)
        D.hline(s, x, x + 330, y + 32, col, 1.0)
        yy = y + 46
        for name, pg in items:
            D.text(s, x, yy, 260, name, 11.5, SONG, WHITE, 1.3)
            D.text(s, x + 286, yy, 44, pg, 9.5, TNR, GOLD_LIGHT, 1.2, PP_ALIGN.RIGHT)
            yy += 25
    D.text(s, 1090, 688, 90, "10", 9, TNR, GOLD_LIGHT, 1.2, PP_ALIGN.RIGHT)
    D.footer(s, p, dark=True)


cover(); contract(); book_map(); editor_note(); name_decode(); palette_page(); five_layers(); toolbox(); emotion_curve(); toc()
D.prs.save(P / "_deck_front.pptx")
print("pages", page)
