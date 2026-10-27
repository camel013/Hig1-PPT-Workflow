from pathlib import Path


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "svg_output"
OUT.mkdir(parents=True, exist_ok=True)

PAPER = "#F2EEE6"
INK = "#171817"
NAVY = "#20374D"
RUST = "#963B31"
PALE_BLUE = "#C9D1D2"
WARM_GRAY = "#BEB8AD"
WHITE = "#FBFAF7"
SOFT_NAVY = "#DAE0E1"

TITLE_FONT = "KaiTi"
BODY_FONT = "STZhongsong"


def svg_start() -> list[str]:
    return [
        '<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">',
    ]


def finish(lines: list[str], number: int) -> str:
    lines.append(f'<text x="1206" y="691" text-anchor="end" font-family="{BODY_FONT}" font-size="15" fill="{INK}">{number:02d} / 12</text>')
    lines.append("</svg>")
    return "\n".join(lines) + "\n"


def write(name: str, lines: list[str], number: int) -> None:
    (OUT / name).write_text(finish(lines, number), encoding="utf-8")


def footer(lines: list[str], source: str, light: bool = False) -> None:
    line_color = "#5E7687" if light else WARM_GRAY
    text_color = "#B9C4C9" if light else "#4D4B46"
    lines.append(f'<line x1="74" y1="654" x2="1206" y2="654" stroke="{line_color}" stroke-width="1"/>')
    lines.append(f'<text x="74" y="691" font-family="{BODY_FONT}" font-size="15" fill="{text_color}">{source}</text>')


def text(
    lines: list[str],
    x: float,
    y: float,
    value: str,
    size: float,
    color: str = INK,
    font: str = BODY_FONT,
    weight: str = "normal",
    anchor: str = "start",
    spacing: float | None = None,
) -> None:
    extra = f' letter-spacing="{spacing}"' if spacing is not None else ""
    lines.append(
        f'<text x="{x}" y="{y}" text-anchor="{anchor}" font-family="{font}" '
        f'font-size="{size}" font-weight="{weight}"{extra} fill="{color}">{value}</text>'
    )


def multi_text(
    lines: list[str],
    x: float,
    y: float,
    values: list[str],
    size: float,
    color: str = INK,
    font: str = BODY_FONT,
    weight: str = "normal",
    line_gap: float = 1.45,
    anchor: str = "start",
) -> None:
    spans = []
    for idx, value in enumerate(values):
        dy = 0 if idx == 0 else round(size * line_gap, 1)
        spans.append(f'<tspan x="{x}" dy="{dy}">{value}</tspan>')
    lines.append(
        f'<text x="{x}" y="{y}" text-anchor="{anchor}" font-family="{font}" '
        f'font-size="{size}" font-weight="{weight}" fill="{color}">{"".join(spans)}</text>'
    )


def rule(lines: list[str], x1: float, y1: float, x2: float, y2: float, color: str = WARM_GRAY, width: float = 1) -> None:
    lines.append(
        f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{color}" stroke-width="{width}"/>'
    )


def dot(lines: list[str], x: float, y: float, radius: float, color: str) -> None:
    lines.append(f'<circle cx="{x}" cy="{y}" r="{radius}" fill="{color}"/>')


def arrow_right(lines: list[str], x: float, y: float, length: float, color: str) -> None:
    rule(lines, x, y, x + length - 14, y, color, 2)
    lines.append(f'<polygon points="{x + length - 18},{y - 7} {x + length},{y} {x + length - 18},{y + 7}" fill="{color}"/>')


def slide_01() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{NAVY}"/>')
    lines.append(f'<path d="M 690 560 C 830 475, 1040 405, 1370 318" fill="none" stroke="#667D8B" stroke-width="2"/>')
    lines.append(f'<path d="M 742 586 C 906 493, 1078 422, 1362 352" fill="none" stroke="{RUST}" stroke-width="4"/>')
    lines.append(f'<rect x="74" y="72" width="7" height="96" fill="{RUST}"/>')
    text(lines, 102, 111, "历史中的文化回声", 21, "#C7D0D3", BODY_FONT, spacing="2")
    text(lines, 68, 311, "蓝调背后", 102, WHITE, TITLE_FONT)
    text(lines, 74, 375, "影子中的人们", 44, WHITE, TITLE_FONT)
    text(lines, 76, 458, "一棵棉花、一部机器、一道法律，共同写进了声音。", 25, "#D7DEE0")
    text(lines, 720, 138, "“", 104, RUST, BODY_FONT)
    multi_text(
        lines,
        770,
        204,
        [
            "我的一生，我一直在创造它；",
            "我的一生，白种人拿走了它。",
        ],
        29,
        WHITE,
        BODY_FONT,
        line_gap=1.65,
    )
    text(lines, 770, 305, "Bessie Brown · Song From A Cotton Field", 18, "#C7D0D3")
    text(lines, 76, 612, "依据《“蓝调”背后——影子中的人们》精选", 16, "#B9C4C9")
    write("01_封面.svg", lines, 1)


def slide_02() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{PAPER}"/>')
    text(lines, 74, 92, "核心判断", 20, RUST)
    text(lines, 72, 158, "蓝调的历史底色", 52, INK, TITLE_FONT)
    multi_text(
        lines,
        74,
        232,
        [
            "“整个美国现代社会（政治、经济、文化、艺术……）",
            "的形成与奴隶制密不可分。”",
        ],
        29,
        INK,
        BODY_FONT,
        line_gap=1.55,
    )
    rule(lines, 88, 405, 1192, 585, PALE_BLUE, 2)
    rule(lines, 88, 585, 1192, 405, PALE_BLUE, 2)
    dot(lines, 640, 495, 13, RUST)
    text(lines, 640, 536, "相互缠结", 25, RUST, BODY_FONT, anchor="middle")

    text(lines, 90, 390, "种植园经济", 29, NAVY, TITLE_FONT)
    multi_text(lines, 90, 510, ["土地", "劳动力", "作物利润"], 20, "#4F4E49", BODY_FONT, line_gap=1.5)
    text(lines, 1188, 390, "工业资本主义", 29, NAVY, TITLE_FONT, anchor="end")
    multi_text(lines, 1188, 510, ["工厂", "铁路", "市场扩张"], 20, "#4F4E49", BODY_FONT, line_gap=1.5, anchor="end")

    footer(lines, "原文 p.4")
    write("02_核心判断.svg", lines, 2)


def slide_03() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{PAPER}"/>')
    text(lines, 74, 84, "种植园经济", 20, RUST)
    text(lines, 72, 150, "种植园不是自然产物", 50, INK, TITLE_FONT)
    text(lines, 74, 194, "烟草、奴隶贸易与法律，共同把南方推进种植园时代。", 24, "#4F4E49")

    lines.append(
        f'<path d="M 88 360 C 205 332, 260 373, 327 288 '
        f'S 454 315, 558 342 S 705 250, 810 286 S 978 338, 1107 219" '
        f'fill="none" stroke="{NAVY}" stroke-width="3.5"/>'
    )
    nodes = [
        (92, 359, "1612", "首批烟草试种成功", "above"),
        (327, 288, "1619", "约 20 名非洲人被荷兰人贩入", "below"),
        (558, 342, "1661–70", "终身与世袭奴役写入殖民地法律", "above"),
        (810, 286, "1787", "宪法拖延废奴，并保护奴隶主追索逃奴", "below"),
        (1062, 219, "18 世纪末", "南北经济", "above"),
    ]
    for idx, (x, y, year, label, side) in enumerate(nodes):
        dot(lines, x, y, 7, RUST)
        rule(lines, x, y - 12, x, y + 18, NAVY, 1.4)
        if side == "above":
            text(lines, x, y - 72, year, 31, RUST)
            if idx == len(nodes) - 1:
                multi_text(lines, x, y - 42, ["南北经济", "联系日益紧密"], 18, INK, BODY_FONT, line_gap=1.4, anchor="middle")
            else:
                multi_text(lines, x, y - 42, label.split("，") if len(label) < 17 else [label[:12], label[12:]], 18, INK, BODY_FONT, line_gap=1.2, anchor="middle")
        else:
            text(lines, x, y + 38, year, 31, RUST)
            if len(label) < 18:
                text(lines, x, y + 70, label, 18)
            else:
                cut = 15
                text(lines, x, y + 70, label[:cut], 18)
                text(lines, x, y + 95, label[cut:], 18)

    lines.append(f'<rect x="74" y="552" width="9" height="93" fill="{RUST}"/>')
    multi_text(
        lines,
        107,
        587,
        [
            "气候决定什么样的作物能够生长，",
            "资本、法律和暴力决定谁被迫劳动。",
        ],
        27,
        INK,
        TITLE_FONT,
        line_gap=1.6,
    )
    footer(lines, "原文 p.1–3")
    write("03_种植园起源.svg", lines, 3)


def slide_04() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{PAPER}"/>')
    text(lines, 74, 84, "劳动力结构", 20, RUST)
    text(lines, 72, 150, "终身奴役如何被制度化", 50, INK, TITLE_FONT)

    stages = [
        (74, 264, "01", "劳力缺口", "烟草与水稻种植迅速扩张"),
        (360, 312, "02", "契约奴到期离开", "旅费债务清偿后成为自由人"),
        (646, 264, "03", "黑人终身服役", "母系身份继承，被法律当作财产"),
        (932, 312, "04", "种族主义提供合法性", "把“他者”排除在平等之外"),
    ]
    for idx, (x, y, num, head, body) in enumerate(stages):
        text(lines, x, y - 44, num, 21, RUST)
        text(lines, x, y, head, 30, INK, TITLE_FONT)
        text(lines, x, y + 39, body, 19, "#4F4E49")
        if idx < len(stages) - 1:
            arrow_right(lines, x + 235, y - 11, 66 if idx % 2 == 0 else 72, NAVY)

    lines.append(f'<path d="M 74 452 C 268 430, 412 468, 590 451 S 874 424, 1115 458" fill="none" stroke="{PALE_BLUE}" stroke-width="2"/>')
    text(lines, 74, 530, "1661–1670", 70, RUST)
    multi_text(
        lines,
        610,
        493,
        [
            "这一阶段的关键变化不是“劳动力更多”，",
            "而是“劳动者被永久固定在身份之中”。",
        ],
        24,
        INK,
        TITLE_FONT,
        line_gap=1.6,
    )
    text(lines, 610, 577, "原文写到：黑人奴隶制是殖民者残暴压迫的表现和手段。", 18, "#4F4E49")
    footer(lines, "原文 p.1–3 · 脚注 8–10")
    write("04_制度化.svg", lines, 4)


def slide_05() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{PAPER}"/>')
    text(lines, 74, 84, "棉花帝国", 20, RUST)
    text(lines, 72, 150, "棉花把南北锁在一起", 50, INK, TITLE_FONT)
    text(lines, 1006, 166, "60%+", 78, RUST, BODY_FONT, anchor="end")
    text(lines, 1006, 230, "1807 年，英国主要港口棉花来源中，美国占六成以上", 16, "#4F4E49", anchor="end")

    lines.append(f'<ellipse cx="620" cy="443" rx="447" ry="205" fill="none" stroke="{SOFT_NAVY}" stroke-width="3"/>')
    lines.append(f'<path d="M 712 245 C 930 247, 1072 321, 1060 425" fill="none" stroke="{NAVY}" stroke-width="4"/>')
    lines.append(f'<path d="M 173 443 C 168 544, 292 632, 500 643" fill="none" stroke="{RUST}" stroke-width="4"/>')
    lines.append(f'<polygon points="1060,437 1047,418 1075,422" fill="{NAVY}"/>')
    lines.append(f'<polygon points="490,631 507,620 506,650" fill="{RUST}"/>')

    text(lines, 630, 421, "资本回路", 29, INK, TITLE_FONT, anchor="middle")
    text(lines, 630, 470, "强制劳动与全球市场", 19, "#4F4E49", anchor="middle")
    text(lines, 223, 275, "黑人奴隶劳动", 27, INK, TITLE_FONT, anchor="middle")
    text(lines, 1002, 275, "原棉出口", 27, INK, TITLE_FONT, anchor="middle")
    text(lines, 1030, 589, "英国机器与北方工业", 27, INK, TITLE_FONT, anchor="middle")
    text(lines, 245, 604, "利润回流", 27, INK, TITLE_FONT, anchor="middle")
    text(lines, 245, 633, "再次购买土地与奴隶", 18, "#4F4E49", anchor="middle")

    footer(lines, "原文 p.4–5")
    write("05_棉花帝国.svg", lines, 5)


def slide_06() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{PAPER}"/>')
    text(lines, 74, 84, "工业化背景", 20, RUST)
    text(lines, 72, 150, "工业资本与奴隶制共同增长", 50, INK, TITLE_FONT)
    text(lines, 74, 194, "北方制造业扩张，并未与南方奴隶制共同体的增长彼此隔绝。", 23, "#4F4E49")

    columns = [
        (74, "制造业就业", "约 50 万", "1840", "120 万", "1850"),
        (470, "制造业贡献占比", "17%", "1839", "29%", "1849"),
        (866, "农业部门占比", "72%", "1839", "60%", "1849"),
    ]
    for idx, (x, label, old, old_year, new, new_year) in enumerate(columns):
        if idx:
            rule(lines, x - 32, 258, x - 32, 522, WARM_GRAY, 1)
        text(lines, x, 266, label, 24, NAVY, TITLE_FONT)
        text(lines, x, 330, old, 56, "#67645D")
        text(lines, x, 365, old_year, 17, "#4F4E49")
        rule(lines, x, 398, x + 260, 398, PALE_BLUE, 2)
        text(lines, x, 466, new, 66, RUST)
        text(lines, x, 506, new_year, 17, "#4F4E49")

    lines.append(f'<rect x="74" y="552" width="9" height="77" fill="{NAVY}"/>')
    multi_text(
        lines,
        108,
        583,
        [
            "“工业资本主义的发展和奴隶制棉花种植园的发展",
            "相互缠结。”",
        ],
        25,
        INK,
        BODY_FONT,
        line_gap=1.55,
    )
    footer(lines, "数据见原文脚注 33")
    write("06_共同增长.svg", lines, 6)


def slide_07() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{PAPER}"/>')
    text(lines, 74, 84, "政治危机", 20, RUST)
    text(lines, 72, 150, "西进把经济冲突推向国家政治", 50, INK, TITLE_FONT)

    points = [
        (100, 420, "1820", "密苏里妥协", "36°30′ 划出自由州与蓄奴州界线"),
        (335, 387, "1850", "逃奴追缉法", "联邦政府参与追索逃亡奴隶"),
        (575, 340, "1854", "堪萨斯－内布拉斯加法案", "地域限制被取消，共和党成立"),
        (820, 290, "1854–59", "流血的堪萨斯", "两个政府与武装垦殖者对抗"),
        (1082, 230, "1859", "约翰·布朗起义", "战争倒计时开始"),
    ]
    lines.append(f'<path d="M 100 447 L 335 414 L 575 367 L 820 317 L 1082 257" fill="none" stroke="{RUST}" stroke-width="5"/>')
    for x, y, year, title, body in points:
        dot(lines, x, y + 27, 8, RUST)
        text(lines, x, y + 74, year, 31, NAVY)
        text(lines, x, y + 104, title, 21, INK, TITLE_FONT)
        if len(body) > 20:
            text(lines, x, y + 130, body[:17], 17, "#4F4E49")
            text(lines, x, y + 153, body[17:], 17, "#4F4E49")
        else:
            text(lines, x, y + 130, body, 17, "#4F4E49")

    text(lines, 74, 602, "“联邦事实上已成了统治南部的 30 万奴隶主的奴隶。”", 27, RUST, TITLE_FONT)
    text(lines, 74, 632, "马克思语，转引自原文 p.6", 16, "#4F4E49")
    footer(lines, "原文 p.5–6")
    write("07_政治危机.svg", lines, 7)


def slide_08() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{NAVY}"/>')
    text(lines, 74, 84, "1861—1865 · 内战", 20, "#C7D0D3", spacing="2")
    text(lines, 72, 166, "拯救联邦，还是摧毁奴隶制？", 49, WHITE, TITLE_FONT)

    lines.append(f'<rect x="74" y="224" width="9" height="281" fill="{RUST}"/>')
    multi_text(
        lines,
        116,
        272,
        [
            "“我在这场战争中的最高目标是拯救联邦。",
            "如果我能拯救联邦而不解放任何一个奴隶，",
            "我愿意这样做；如果为了拯救联邦",
            "需要解放一部分奴隶而保留另一部分，",
            "我也愿意这样做。”",
        ],
        33,
        WHITE,
        BODY_FONT,
        line_gap=1.52,
    )
    text(lines, 116, 539, "亚伯拉罕·林肯，1862 年书信节选", 18, "#C7D0D3")

    text(lines, 944, 287, "1861", 59, "#8DA0AB", BODY_FONT, anchor="middle")
    text(lines, 944, 326, "南部同盟成立，战争开始", 18, "#D7DEE0", anchor="middle")
    text(lines, 1084, 476, "1863", 86, WHITE, BODY_FONT, anchor="middle")
    text(lines, 1084, 536, "《解放黑人奴隶宣言》", 20, WHITE, anchor="middle")
    arrow_right(lines, 944, 374, 83, RUST)

    footer(lines, "原文 p.6–7", light=True)
    write("08_内战目标.svg", lines, 8)


def slide_09() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{PAPER}"/>')
    text(lines, 570, 84, "黑人的主体行动", 20, RUST)
    text(lines, 568, 150, "解放来自持续斗争", 50, INK, TITLE_FONT)
    text(lines, 74, 300, "数十万", 105, RUST)
    text(lines, 78, 360, "黑人士兵参战", 27, INK, TITLE_FONT)
    text(lines, 78, 397, "把战争推进到黑人解放的目标", 18, "#4F4E49")

    action_items = [
        ("01", "组织废奴", "道格拉斯等黑人知识分子推动废奴组织与集会。"),
        ("02", "参与战争", "黑人的逃亡和反抗，成为战争转向解放的重要力量。"),
        ("03", "切断棉花", "英国工人顶住“棉业危机”，打击南方棉花出口。"),
    ]
    for idx, (num, head, body) in enumerate(action_items):
        y = 238 + idx * 118
        text(lines, 568, y, num, 20, RUST)
        text(lines, 618, y, head, 27, NAVY, TITLE_FONT)
        text(lines, 618, y + 33, body, 19, "#4F4E49")
        rule(lines, 568, y + 70, 1206, y + 70, PALE_BLUE, 1)

    text(lines, 74, 565, "“为了自由和尊严，和联邦政府共同战斗！”", 31, NAVY, TITLE_FONT)
    footer(lines, "原文 p.7")
    write("09_主体行动.svg", lines, 9)


def slide_10() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{PAPER}"/>')
    text(lines, 74, 84, "重建时期", 20, RUST)
    text(lines, 72, 150, "军队短暂改写了南方秩序", 50, INK, TITLE_FONT)

    text(lines, 75, 244, "1866", 43, RUST)
    multi_text(lines, 75, 278, ["共和党取得国会", "三分之二议席"], 21, INK, TITLE_FONT, line_gap=1.35)
    text(lines, 75, 382, "10 / 5", 50, NAVY)
    multi_text(lines, 78, 430, ["10 个叛乱州划为 5 个军区", "军队代行地方行政权"], 19, INK, BODY_FONT, line_gap=1.35)
    text(lines, 75, 522, "黑人社区", 25, INK, TITLE_FONT)
    text(lines, 75, 553, "多由留在本地谋生的前奴隶形成。", 18, "#4F4E49")
    text(lines, 75, 582, "黑人士兵把驻军视为社会改造力量。", 18, "#4F4E49")

    chart_x = 722
    chart_y_top = 222
    chart_y_bottom = 565
    rule(lines, chart_x, chart_y_top, chart_x, chart_y_bottom, WARM_GRAY, 1.5)
    rule(lines, chart_x, chart_y_bottom, 1207, chart_y_bottom, WARM_GRAY, 1.5)
    bars = [
        (788, 180000, 280, "1865", "约 18 万"),
        (948, 20000, 38, "1867", "约 2 万"),
        (1108, 9300, 18, "1870", "9 千余"),
    ]
    for x, value, height, year, label in bars:
        lines.append(f'<rect x="{x}" y="{chart_y_bottom - height}" width="82" height="{height}" fill="{NAVY if value > 30000 else RUST}"/>')
        text(lines, x + 41, chart_y_bottom + 29, year, 19, INK, anchor="middle")
        text(lines, x + 41, chart_y_bottom - height - 15, label, 20, INK, anchor="middle")

    text(lines, 948, 200, "联邦军队在南方驻扎人数", 18, NAVY, anchor="middle")
    footer(lines, "原文 p.8–9")
    write("10_重建空间.svg", lines, 10)


def slide_11() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{INK}"/>')
    text(lines, 74, 84, "重建失败", 20, RUST)
    text(lines, 72, 150, "黑人再次成为“看不见的人”", 50, WHITE, TITLE_FONT)

    lines.append(f'<path d="M 835 227 A 250 250 0 1 1 835 617" fill="none" stroke="{NAVY}" stroke-width="28"/>')
    lines.append(f'<path d="M 902 270 A 185 185 0 1 1 902 574" fill="none" stroke="{RUST}" stroke-width="8"/>')
    lines.append(f'<path d="M 958 332 A 115 115 0 0 0 958 512" fill="none" stroke="#59615F" stroke-width="2"/>')
    text(lines, 1000, 380, "承诺", 26, "#C7D0D3", TITLE_FONT, anchor="middle")
    text(lines, 1000, 415, "撤离", 26, "#C7D0D3", TITLE_FONT, anchor="middle")
    text(lines, 1000, 450, "回潮", 26, "#C7D0D3", TITLE_FONT, anchor="middle")

    multi_text(
        lines,
        76,
        224,
        [
            "军队撤离后，旧种植园主重新掌握州政府。",
            "三 K 党、种族隔离和种族暴力卷土重来。",
            "“民主”和“平等”停留在选票的语言里，",
            "没有兑现为活生生的社会现实。",
        ],
        24,
        "#E7E3DB",
        BODY_FONT,
        line_gap=1.55,
    )
    text(lines, 76, 544, "“在黑人的劳动打上屈辱烙印的地方，", 31, WHITE, TITLE_FONT)
    text(lines, 76, 582, "白人的劳动也不能得到解放。”", 31, WHITE, TITLE_FONT)
    text(lines, 76, 616, "马克思语，转引自原文 p.9", 16, "#B9C4C9")

    footer(lines, "原文 p.9", light=True)
    write("11_重建失败.svg", lines, 11)


def slide_12() -> None:
    lines = svg_start()
    lines.append(f'<rect x="0" y="0" width="1280" height="720" fill="{PAPER}"/>')
    lines.append(f'<path d="M 72 651 C 300 578, 778 554, 1225 337" fill="none" stroke="{SOFT_NAVY}" stroke-width="17"/>')
    lines.append(f'<path d="M 104 667 C 326 615, 760 598, 1210 402" fill="none" stroke="{RUST}" stroke-width="3"/>')
    text(lines, 74, 90, "尾声", 20, RUST)
    text(lines, 72, 175, "谁在歌唱", 72, INK, TITLE_FONT)
    multi_text(
        lines,
        74,
        287,
        [
            "人是宇宙和声中",
            "一个响亮有力的“是”字。",
        ],
        55,
        INK,
        TITLE_FONT,
        line_gap=1.48,
    )
    text(lines, 76, 410, "易卜拉欣·弗朗茨·法农", 18, "#4F4E49")

    multi_text(
        lines,
        74,
        491,
        [
            "蓝调承接种植园里的圣歌与民歌，",
            "把共同的哀愁、劳动与愤怒留在声音里。",
        ],
        25,
        NAVY,
        BODY_FONT,
        line_gap=1.55,
    )
    footer(lines, "原文 p.1、p.9 · 12 页精选")
    write("12_尾声.svg", lines, 12)


if __name__ == "__main__":
    slide_01()
    slide_02()
    slide_03()
    slide_04()
    slide_05()
    slide_06()
    slide_07()
    slide_08()
    slide_09()
    slide_10()
    slide_11()
    slide_12()
    print(f"Generated 12 slides in {OUT}")
