# -*- coding: utf-8 -*-
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).parent))
import deck_lib as D

s = D.new_slide(D.PAPER)
D.rect(s, 0, 0, 760, D.H, fill=D.CORAL)
D.rect(s, 760, 0, 520, D.H, fill=D.INDIGO)
D.rect(s, 860, 0, 420, D.H, fill=D.INDIGO_DEEP)
D.rect(s, 760, 0, 100, D.H, fill=D.NAVY)
# chevron band
for k in range(3):
    pts=[]
    y0=42+k*15
    for i in range(0, 220, 20):
        pts.append((40+i, y0 + (0 if (i//20)%2==0 else -9)))
        pts.append((40+i+10, y0 + (-9 if (i//20)%2==0 else 0)))
    D.freeform(s, pts, color=D.CORAL_LIGHT, width=1.0)
# white title card
D.rect(s, 930, 70, 286, 360, fill=D.WHITE)
D.text(s, 958, 84, 240, "蛭", 96, D.KAI, D.INK, 1.0)
D.text(s, 958, 186, 240, "文", 96, D.KAI, D.INK, 1.0)
D.text(s, 962, 302, 220, "ZHIWEN", 11, D.TNR, D.MUTED, 1.3, spacing=3)
D.text(s, 962, 336, 220, "特殊高中教科书", 9, D.SONG, D.MUTED, 1.3, spacing=1.5)
D.rect(s, 930, 430, 286, 58, fill=D.GOLD)
D.text(s, 958, 444, 230, "不必修 · 中册", 17, D.SONG, D.WHITE, 1.2, spacing=1.2)
# left text
D.text(s, 64, 94, 520, "高中语文 · 文本细读与写作", 12, D.SONG, D.WHITE, 1.4, spacing=2)
D.vline(s, 64, 156, "一本假课本的读法", 34, D.WHITE, D.KAI, 1.3, spacing=2)
D.text(s, 64, 590, 520, "从戏仿到见证：语言如何成为证词", 18, D.SONG, D.WHITE, 1.4)
D.text(s, 64, 636, 520, "4 个单元 · 24 篇课文 · 3 首古诗词 · 1 篇跋", 11, D.TNR, D.CORAL_LIGHT, 1.3)
# arcs
D.arc(s, 690, 720, 420, 300, -160, -20, D.TEAL, 1.4)
D.arc(s, 690, 720, 360, 250, -155, -30, D.CORAL_LIGHT, 0.8)
# seal
D.arc(s, 1150, 122, 48, 48, 0, 360, D.CORAL_DARK, 1.1)
D.text(s, 1128, 110, 64, "2026", 10, D.TNR, D.CORAL_DARK, 1.0, D.PP_ALIGN.CENTER)
D.text(s, 1120, 690, 120, "01", 9, D.TNR, D.CORAL_LIGHT, 1.2, D.PP_ALIGN.RIGHT)

# divider
s2 = D.new_slide(D.GOLD)
D.rect(s2, 0, 0, 1280, 14, fill=D.INK)
D.unit_no(s2, 56, 150, "壹", 250, D.PAPER)
D.text(s2, 330, 200, 700, "以古写今", 64, D.KAI, D.PAPER, 1.2)
D.text(s2, 334, 300, 760, "第一单元 · 古典诗歌的戏仿", 18, D.SONG, D.INK, 1.4, spacing=1.4)
D.hline(s2, 334, 1120, 352, D.INK, 1.0)
D.text(s2, 334, 386, 700, "学其体式而活其精神，以古写今。", 22, D.SONG, D.INK, 1.6)
D.arc(s2, 1160, 720, 460, 340, -175, -55, D.PAPER, 1.2)
D.text(s2, 1060, 620, 160, "04 篇", 38, D.TNR, D.INK, 1.0, D.PP_ALIGN.RIGHT)
D.text(s2, 1060, 674, 160, "pp. 1—10", 11, D.TNR, D.INK, 1.2, D.PP_ALIGN.RIGHT)

# poem page
s3 = D.new_slide(D.PAPER)
D.page_title(s3, 17, "第一单元 / 01 蛭歌行并序", "声音先于人物到场")
D.text(s3, 64, 170, 500, "全诗以“夜学”起，以“击鼓”承。人物尚未出场，声音已经占满教室；等读者看见他时，压迫感已经完成。", 18, D.SONG, D.INK, 1.75)
D.text(s3, 64, 300, 470, "“忽闻台上伐鼓声，书生顿首笔不行。”", 28, D.KAI, D.CORAL_DARK, 1.5)
D.text(s3, 64, 390, 470, "鼓声不是背景，是秩序的开场。它先把人按住，再让理由进场。", 15, D.SONG, D.MUTED, 1.7)
D.rect(s3, 64, 470, 3, 150, fill=D.GOLD)
D.text(s3, 84, 470, 440, "读法：把“声音”当作人物来读。谁有权发出声音？谁只能听见？谁被迫沉默？", 14, D.SONG, D.INK, 1.7)
lines = ["嘉陵江头夜学文", "熏风热流夏恼人", "忽闻台上伐鼓声", "未成曲调先有愤"]
for i, line in enumerate(lines):
    D.vline(s3, 1134 - i * 58, 150, line, 30, D.INK, D.KAI, 1.22, spacing=2)
D.hline(s3, 900, 1200, 640, D.GOLD, 1.2)
D.text(s3, 900, 652, 300, "《蛭歌行并序》节选", 10, D.TNR, D.MUTED, 1.2)
D.footer(s3, 15)

D.prs.save(r"_test.pptx")
print("saved", len(D.prs.slides.__iter__.__self__._sldIdLst))


