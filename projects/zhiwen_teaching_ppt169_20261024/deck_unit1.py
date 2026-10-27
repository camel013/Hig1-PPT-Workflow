# -*- coding: utf-8 -*-
"""Units 1-4: pages 13-81."""
import deck_lib as D
from deck_lib import (PAPER, PAPER_WARM, WHITE, INK, MUTED, RULE, CORAL, CORAL_DARK,
                      CORAL_LIGHT, INDIGO, INDIGO_DEEP, NAVY, TEAL, TEAL_DARK, GOLD,
                      GOLD_LIGHT, FANG, SONG, TNR, PP_ALIGN)
from state import nxt

HEAD = FANG


def edit(kicker, title, body, quote, qsrc, accent=CORAL, bg=PAPER, dark=False,
         note=None, tags=None, quote_size=25):
    p = nxt()
    s = D.new_slide(bg)
    D.page_title(s, p, kicker, title, dark=dark)
    fg = WHITE if dark else INK
    mu = GOLD_LIGHT if dark else MUTED
    D.text(s, 64, 160, 530, body, 16.5, SONG, fg, 1.82)
    D.vrule(s, 620, 160, 560, accent, 2.0)
    D.text(s, 654, 166, 540, quote, min(quote_size, 23), SONG, accent if not dark else GOLD_LIGHT, 1.66)
    D.text(s, 654, 470, 500, qsrc, 10.5, HEAD, mu, 1.4)
    if tags:
        x = 64
        for t in tags:
            w = D.measure(t, 10, SONG) + 22
            D.rect(s, x, 610, w, 24, fill=accent)
            D.text(s, x + 11, 616, w - 18, t, 9.5, HEAD, WHITE, 1.1)
            x += w + 10
    if note:
        D.text(s, 64, 652, 1080, note, 10.5, SONG, mu, 1.4)
    D.footer(s, p, dark=dark)
    return s


def text_intro(no, text_title, author, headline, body, quote, qsrc, accent=CORAL, tags=None):
    p = nxt()
    s = D.new_slide(PAPER)
    D.page_title(s, p, f"第一单元 / {no}", headline)
    D.rect(s, 64, 140, 132, 210, fill=accent)
    D.text(s, 78, 156, 112, no, 84, TNR, WHITE, 1.0)
    D.text(s, 220, 150, 500, text_title, 32, HEAD, INK, 1.3)
    D.text(s, 220, 204, 500, author, 12.5, HEAD, MUTED, 1.4, spacing=1.2)
    D.text(s, 64, 392, 620, body, 15, SONG, INK, 1.78)
    D.rect(s, 760, 150, 456, 400, fill=accent)
    D.text(s, 792, 178, 392, quote, 21, SONG, WHITE, 1.62)
    D.hline(s, 792, 1184, 404, WHITE, 3.0)
    D.text(s, 792, 424, 380, qsrc, 10, HEAD, WHITE, 1.4)
    D.text(s, 792, 458, 380, "先把它放回它借来的那个“壳”里，再看它换了什么心。", 10.5, SONG, WHITE, 1.55)
    D.hline(s, 64, 584, 520, accent, 4.0)
    D.text(s, 64, 536, 120, "读法", 11, HEAD, accent, 1.2)
    D.text(s, 132, 534, 470, "先听声音，再看人；先看铺陈，再看判断。", 13.5, SONG, INK, 1.5)
    if tags:
        x = 64
        for t in tags:
            w = D.measure(t, 9.5, SONG) + 20
            D.rect(s, x, 632, w, 23, fill=accent)
            D.text(s, x + 10, 638, w - 16, t, 9, HEAD, WHITE, 1.1)
            x += w + 9
    D.footer(s, p)
    return s


def three_col(kicker, title, lead, items, accent=CORAL, note=None, bg=PAPER, dark=False, bottom=None):
    p = nxt()
    s = D.new_slide(bg)
    D.page_title(s, p, kicker, title, dark=dark)
    fg = WHITE if dark else INK
    mu = GOLD_LIGHT if dark else MUTED
    D.text(s, 64, 154, 1050, lead, 16.5, SONG, fg, 1.75)
    for i, (h, b) in enumerate(items):
        x = 64 + i * 390
        D.text(s, x, 300, 60, f"0{i+1}", 26, TNR, accent, 1.0)
        D.hline(s, x, x + 330, 342, accent if dark else RULE, 1.0)
        D.text(s, x, 360, 330, h, 22, HEAD, fg, 1.25)
        D.text(s, x, 404, 330, b, 13.5, SONG, fg, 1.7)
    if bottom:
        D.rect(s, 64, 500, 1152, 78, fill=accent)
        D.text(s, 90, 520, 1100, bottom, 17, SONG, WHITE, 1.42)
    if note:
        D.text(s, 64, 604, 1120, note, 11, SONG, mu, 1.45)
    D.footer(s, p, dark=dark)


def grid4(kicker, title, lead, items, accent=CORAL, note=None, bg=PAPER_WARM):
    p = nxt()
    s = D.new_slide(bg)
    D.page_title(s, p, kicker, title)
    D.text(s, 64, 154, 1080, lead, 15.5, SONG, INK, 1.7)
    for i, (h, b) in enumerate(items):
        x = 64 + (i % 2) * 584
        y = 260 + (i // 2) * 186
        D.rect(s, x, y, 4, 150, fill=accent)
        D.text(s, x + 24, y - 2, 44, f"0{i+1}", 17, TNR, accent, 1.0)
        D.text(s, x + 82, y - 4, 430, h, 21, HEAD, INK, 1.2)
        D.text(s, x + 82, y + 34, 450, b, 12.5, SONG, INK, 1.65)
    if note:
        D.text(s, 64, 646, 1120, note, 10.5, SONG, MUTED, 1.4)
    D.footer(s, p)


def steps(kicker, title, lead, items, accent=CORAL, note=None, bg=PAPER):
    p = nxt()
    s = D.new_slide(bg)
    D.page_title(s, p, kicker, title)
    D.text(s, 64, 154, 1080, lead, 16, SONG, INK, 1.7)
    n = len(items)
    start, end, labw = 80, 1216, 180
    gap = (end - labw - start) / (n - 1)
    D.hline(s, 80, 1216, 330, RULE, 1.0)
    for i, (h, b) in enumerate(items):
        x = start + i * gap
        D.rect(s, x - 4, 324, 8, 12, fill=accent)
        D.text(s, x - 6, 270, 80, f"0{i+1}", 20, TNR, accent, 1.0)
        D.text(s, x - 8, 350, 190, h, 19, HEAD, INK, 1.25)
        D.text(s, x - 8, 392, 190, b, 11.5, SONG, INK, 1.6)
    if note:
        D.text(s, 64, 590, 1120, note, 12, SONG, MUTED, 1.5)
    D.footer(s, p)


def table_page(kicker, title, lead, headers, rows, widths, accent=CORAL, note=None, row_h=70):
    p = nxt()
    s = D.new_slide(PAPER)
    D.page_title(s, p, kicker, title)
    if lead:
        D.text(s, 64, 150, 1100, lead, 15, SONG, INK, 1.65)
    x0, y0 = 64, 210
    x = x0
    D.rect(s, x0, y0, sum(widths), 40, fill=accent)
    for h, w in zip(headers, widths):
        D.text(s, x + 10, y0 + 10, w - 14, h, 10.5, HEAD, WHITE, 1.2, spacing=1.2)
        x += w
    y = y0 + 40
    available = 626 - y
    if rows:
        row_h = min(row_h, available / len(rows))
    for r, row in enumerate(rows):
        if r % 2 == 1:
            D.rect(s, x0, y, sum(widths), row_h, fill=PAPER_WARM)
        x = x0
        for c, (cell, w) in enumerate(zip(row, widths)):
            D.text(s, x + 10, y + 12, w - 18, cell, 12.5 if c == 0 else 11.8, SONG, INK, 1.45)
            x += w
        D.hline(s, x0, x0 + sum(widths), y + row_h, RULE, 0.7)
        y += row_h
    if note:
        D.text(s, 64, min(y + 10, 634), 1120, note, 9.8, SONG, MUTED, 1.35)
    D.footer(s, p)


def poem_vertical(kicker, title, lines, source, notes, accent=CORAL, note=None):
    p = nxt()
    s = D.new_slide(PAPER)
    D.page_title(s, p, kicker, title)
    for i, line in enumerate(lines):
        D.vline(s, 1170 - i * 62, 156, line, 24, INK, SONG, 1.22, spacing=1.5)
    y = 176
    for h, b in notes:
        D.text(s, 64, y, 300, h, 14, HEAD, accent, 1.2)
        D.text(s, 64, y + 26, 480, b, 11.5, SONG, INK, 1.55)
        y += 104
    D.hline(s, 64, 560, 620, accent, 1.0)
    D.text(s, 64, 632, 500, source, 9.5, HEAD, MUTED, 1.2)
    if note:
        D.text(s, 600, 632, 580, note, 10, SONG, MUTED, 1.3)
    D.footer(s, p)


def stagger(kicker, title, stanzas, source, accent=CORAL, ghost=None, note=None, ghost_lines=None):
    p = nxt()
    s = D.new_slide(PAPER)
    D.page_title(s, p, kicker, title)
    if ghost_lines:
        for i, line in enumerate(ghost_lines):
            D.text(s, 820, 150 + i * 28, 400, line, 16.5, SONG, RULE, 1.5)
    positions = [(64, 160), (430, 210), (64, 430), (430, 480)]
    for i, stanza in enumerate(stanzas[:4]):
        x, y = positions[i]
        D.hline(s, x, x + 300, y - 12, accent, 1.2)
        yy = y
        for line in stanza:
            D.text(s, x, yy, 320, line, 17, SONG, INK, 1.55)
            yy += 30
    D.text(s, 700, 626, 480, source, 9.5, HEAD, MUTED, 1.2, PP_ALIGN.RIGHT)
    if note:
        D.text(s, 64, 626, 600, note, 10, SONG, MUTED, 1.3)
    D.footer(s, p)


def unit_divider(no, title, sub, quote, color, count, rng, contrast=NAVY, arc=(1330, 740, 430, 190, 280)):
    p = nxt()
    s = D.new_slide(color)
    cx, cy, r, a0, a1 = arc
    D.arc_block(s, cx, cy, r, a0, a1, 26000, contrast)
    D.rect(s, 0, 0, 1280, 14, fill=INK)
    D.unit_no(s, 56, 150, no, 250, PAPER)
    D.text(s, 330, 200, 700, title, 64, HEAD, PAPER, 1.2)
    D.text(s, 334, 300, 760, sub, 18, SONG, INK, 1.4, spacing=1.4)
    D.hline(s, 334, 1120, 352, INK, 1.0)
    D.text(s, 334, 386, 700, quote, 22, SONG, INK, 1.6)
    D.rect(s, 1000, 560, 216, 122, fill=contrast)
    D.text(s, 1022, 576, 176, count, 34, HEAD, WHITE, 1.0)
    D.text(s, 1022, 630, 176, rng, 11, HEAD, WHITE, 1.2)
    D.footer(s, p, color=INK)


def unit1():
    unit_divider("壹", "以古写今", "第一单元 · 古典诗歌的戏仿", "学其体式而活其精神，以古写今。", GOLD, "4 篇", "pp. 1—10", NAVY, (1330, 740, 430, 190, 280))
    # opener
    p = nxt(); s = D.new_slide(PAPER)
    D.page_title(s, p, "第一单元 / 导览", "四种借来的声音")
    D.text(s, 64, 154, 1050, "本单元不教“古诗知识”，而教“古诗如何继续说话”。四篇作品各借一种体式，把同一种压迫写成四种音色。", 16.5, SONG, INK, 1.75)
    items = [
        ("歌行", "长篇叙事，适合铺陈一场持续的声音暴力。", "《蛭歌行》"),
        ("散曲", "短章俚语，像一记响亮的耳光。", "《朝天子·咏水蛭》"),
        ("新月抒情", "越柔的调子，越能显出“落实”的粗暴。", "《再别三班》"),
        ("现代短诗", "第一人称反转，让加害者自己开口。", "《偶然》"),
    ]
    for i, (h, b, ex) in enumerate(items):
        x = 64 + (i % 2) * 584
        y = 270 + (i // 2) * 180
        D.rect(s, x, y, 4, 142, fill=GOLD)
        D.text(s, x + 24, y - 2, 420, h, 22, HEAD, INK, 1.2)
        D.text(s, x + 24, y + 38, 480, b, 13, SONG, INK, 1.65)
        D.text(s, x + 24, y + 106, 460, ex, 11, SONG, GOLD, 1.3)
    D.text(s, 64, 640, 1100, "共同点：都不是“写古诗”，而是让古诗的体式承担今天的经验。", 11.5, SONG, MUTED, 1.3)
    D.footer(s, p)

    text_intro("01", "蛭歌行并序", "佚名 · 616 言 · 歌行体", "把一次骂人，写成一场漫长的演奏",
               "全诗以“夜学”起，以“击鼓”承，以“自述”转，以“收束”合。它把课堂冲突写成声音事件：人物还没出场，鼓声已经占满教室。",
               "“嘉陵江头夜学文，\n熏风热流夏恼人。\n后生休憩早已归，\n举笔欲题无文思。\n忽闻台上伐鼓声，\n书生顿首笔不行。”",
               "《蛭歌行并序》开篇",
               GOLD, ["歌行体", "仿《琵琶行》", "声音政治"])
    # fix newline quote by drawing separately in custom page? text() does not split \n; replace after creation
    # structure
    steps("第一单元 / 01", "五段推进：从听见，到看清，再到议论",
          "长篇歌行的力量在“铺陈”：它不急着下判断，而是让声音、动作、外貌一层层堆起来。",
          [("闻", "夜学困顿，忽闻台上伐鼓。"), ("写", "外貌、动作、声音，层层铺陈。"), ("转", "从写一人，转向写满座。"), ("议", "以“蛭”收束，点明讽喻。"), ("合", "声音落下，愤怒没有出口。")],
          GOLD, "读法：把“声音”当作人物来读。谁有权发出声音？谁只能听见？谁被迫沉默？")
    # sound close reading
    p = nxt(); s = D.new_slide(PAPER)
    D.page_title(s, p, "第一单元 / 01", "声音的四种暴力")
    quotes = [
        ("忽闻台上伐鼓声，书生顿首笔不行。", "声音先于人物出现，制造压迫。"),
        ("星落石陨三两声，未成曲调先有愤。", "以乐写怒，把骂人写成演奏。"),
        ("大鼓嘈嘈如滚雷，小鼓切切如夯垒。", "大小、轻重、快慢，全是权力节奏。"),
        ("此时有声胜无声。", "沉默不是安静，是被压住。"),
    ]
    y = 160
    for i, (q, a) in enumerate(quotes):
        D.text(s, 64, y, 44, f"0{i+1}", 16, TNR, GOLD, 1.0)
        D.text(s, 122, y - 4, 620, q, 19, SONG, INK, 1.55)
        D.text(s, 122, y + 42, 600, a, 11.5, SONG, MUTED, 1.4)
        D.hline(s, 64, 800, y + 88, RULE, 0.7)
        y += 112
    # color-block arc behind the waveform
    D.arc_block(s, 1130, 720, 270, 185, 295, 26000, NAVY)
    # bar waveform
    heights = [40, 78, 56, 104, 70, 132, 92, 150, 110, 76, 120, 58]
    for i, h in enumerate(heights):
        D.rect(s, 900 + i * 22, 540 - h, 10, h, fill=GOLD)
    D.text(s, 756, 606, 220, "声音的强度曲线", 9.5, FANG, MUTED, 1.2)
    D.text(s, 900, 166, 320, "“有声”不是热闹，", 17, SONG, GOLD, 1.55)
    D.text(s, 900, 204, 320, "是另一种不能退出。", 17, SONG, GOLD, 1.55)
    D.footer(s, p)
    # compare table
    table_page("第一单元 / 01", "同一条江，两种命运", "《蛭歌行》把《琵琶行》的“同是天涯沦落人”改成“座中愠恨谁最多”。同情变成了对峙。",
               ["比较项", "《琵琶行》", "《蛭歌行》"],
               [["起点", "浔阳江头夜送客", "嘉陵江头夜学文"],
                ["声音", "琵琶声，抑扬顿挫", "击鼓声，聒噪压迫"],
                ["人物", "琵琶女，身世飘零", "徐规尔，权力在場"],
                ["转折", "同是天涯沦落人", "座中愠恨谁最多"],
                ["结尾", "江州司马青衫湿", "唯留规尔高声闹"]],
               [140, 470, 530], GOLD, "白居易把音乐写成命运，佚名把噪音写成权力。", 78)

    poem_vertical("第一单元 / 02", "短章如刀：一支散曲的全部力气",
                  ["水蛭，蚂蟥，脑瓜小肚儿大。", "化学测试乱如麻，全仗你叫声哇。", "生听了愁眉，徒听了甚怕。", "哪里去辨什么真共假？", "眼见得收翻了这家，", "没抵住那家，", "只吸得血尽人逃罢！"],
                  "《朝天子·咏水蛭》/ 石遂志", [("曲牌", "《朝天子》本是多用于抒情与讽刺的小令。"), ("俚语", "“脑瓜小肚儿大”把人物压成一个可笑的轮廓。"), ("侧写", "不写老师多凶，只写学生“愁眉”“甚怕”。"), ("收束", "“血尽人逃”把管理写成捕食。")],
                  GOLD, "短，不是简单；是把所有铺垫压进一口气。")
    three_col("第一单元 / 02", "以轻写重：小令为什么比长诗更狠",
              "长诗让你看见过程，小令只给你结果。它不解释，直接把“吸血”这个结论拍在桌上。",
              [("轻", "曲牌、俚语、口语，读起来甚至有点好笑。"), ("重", "笑完才发现，被写的是每天要面对的人。"), ("快", "没有转折，没有辩解，一刀到底。")],
              GOLD, "对照《蛭歌行》：长篇靠铺陈，短章靠命中。", bottom="短章不解释，只判决。")

    stagger("第一单元 / 03", "把“再别”改成“砰砰”",
            [["砰砰的我走了，", "正如我砰砰的来；", "我重重地挥手，", "手机绊在讲台。"],
             ["那窗内的人儿，", "是灿烂的金银；", "升迁的艳影，", "在我心头荡漾。"],
             ["但我不能经受诱惑", "人多是没把持住的理由", "夏虫也务必落实", "落实是今晚的三班！"],
             ["哇哇的我走了，", "正如我愿赌服输地来；", "我挥一挥衣袖，", "捶捶害羞的讲台。"]],
            "《再别三班并序》/ 徐蛭魔",
            GOLD, ghost_lines=["轻轻的我走了，", "正如我轻轻的来；", "我轻轻的招手，", "作别西天的云彩。"],
            note="底文越柔，改写越粗；这不是“致敬”，是把抒情诗推到一个它原本不属于的现场。")
    grid4("第一单元 / 03", "四个替换，四次变味",
          "戏仿的力量不在“像”，而在“错位”。每一处替换，都把原诗的优雅挪走一寸。",
          [("轻轻 → 砰砰", "动作变粗，节奏变重，离别成了碰撞。"), ("金柳 → 金银", "景物变利益，康桥的倒影变成升迁的账本。"), ("悄悄 → 哇哇", "沉默变喧闹，私密的告别变成公开的训斥。"), ("作别 → 落实", "公文词闯进抒情句，像一块石头砸进水里。")],
          GOLD, "读法：先背原诗，再读改诗；不适感越强，讽刺越成立。")

    stagger("第一单元 / 04", "当“我”是一条水蛭",
            [["我是三班里的一条水蛭，", "偶尔蠕动在你的窗边——", "你不必抬头，", "更无需报信：“嘘——”"],
             ["在转瞬间抓住了把柄。", "你我对视在无窗帘的内外", "你骂你的，我当我的，水蛭；"],
             ["你知道也好，最好不成事就不惹事", "在我蠕动到你的面前，", "吸干血之后。"]],
            "《偶然》/ 铁骑",
            GOLD, ghost_lines=["我是天空里的一片云，", "偶尔投影在你的波心——", "你不必讶异，", "更无须欢喜——"],
            note="徐志摩的“偶然”是相遇的轻盈；铁骑的“偶然”是捕食的必然。")
    three_col("第一单元 / 04", "视角一翻，读者就失去了安全位置",
              "当“我”是水蛭，读者被迫站在加害者的视角里。这不是替谁辩护，而是让你看见权力如何自我叙述。",
              [("自述", "加害者开口，反而暴露其逻辑：把监视说成“偶遇”。"), ("对视", "“你骂你的，我当我的”——两种语言互不相通。"), ("时间", "“吸干血之后”把抒情的时间变成捕食的时间。")],
              GOLD, "最不舒服的阅读，往往最接近真相。", bottom="视角不是技巧，是伦理位置。")
    table_page("第一单元 / 小结", "四篇作品，四种“借壳”", "它们共同证明：古典体式不是博物馆里的标本，而是可以继续承担现实重量的容器。",
               ["篇目", "体式", "底文", "改写对象", "情绪"],
               [["蛭歌行", "歌行", "《琵琶行》", "声音 / 权力", "愤"],
                ["朝天子", "散曲", "咏物小令", "贪婪 / 吸血", "讽"],
                ["再别三班", "新月抒情", "《再别康桥》", "利益 / 落实", "苦笑"],
                ["偶然", "现代短诗", "《偶然》", "视角 / 捕食", "冷"]],
               [150, 150, 190, 330, 110], GOLD, "单元任务：选两篇，比较“借壳”的方法；写一篇 800 字鉴赏。", 74)
