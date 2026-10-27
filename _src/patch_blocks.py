import io
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_front.py'
s=io.open(p,encoding='utf-8').read()
# remove thin chevrons
old='''    for k in range(3):
        pts = []
        y0 = 42 + k * 15
        for i in range(0, 220, 20):
            pts.append((40 + i, y0 + (0 if (i // 20) % 2 == 0 else -9)))
            pts.append((40 + i + 10, y0 + (-9 if (i // 20) % 2 == 0 else 0)))
        D.freeform(s, pts, color=CORAL_LIGHT, width=1.0)
'''
s=s.replace(old,'''    D.arc_block(s, -40, 760, 430, 300, 28, 24000, TEAL)
    D.arc_block(s, 1330, 700, 380, 180, 268, 26000, CORAL)
''')
s=s.replace('D.text(s, 64, 640, 540, "4 个单元 · 24 篇课文 · 3 首古诗词 · 序与跋", 11, HEAD, CORAL_LIGHT, 1.3)','D.text(s, 64, 640, 540, "4 个单元 · 24 篇课文 · 3 首古诗词 · 序与跋", 11, HEAD, WHITE, 1.3)')
io.open(p,'w',encoding='utf-8').write(s)

p=r'projects\zhiwen_teaching_ppt169_20261024\deck_unit1.py'
s=io.open(p,encoding='utf-8').read()
# text intro: number block
s=s.replace('''    D.text(s, 64, 150, 160, no, 96, TNR, accent, 1.0)
    D.text(s, 64, 270, 130, "篇目", 9, HEAD, MUTED, 1.2, spacing=2)
    D.text(s, 200, 150, 520, text_title, 32, HEAD, INK, 1.3)
    D.text(s, 200, 204, 520, author, 12.5, HEAD, MUTED, 1.4, spacing=1.2)
    D.text(s, 64, 330, 600, body, 15.5, SONG, INK, 1.85)''','''    D.rect(s, 64, 140, 132, 210, fill=accent)
    D.text(s, 78, 156, 112, no, 84, TNR, WHITE, 1.0)
    D.text(s, 220, 150, 500, text_title, 32, HEAD, INK, 1.3)
    D.text(s, 220, 204, 500, author, 12.5, HEAD, MUTED, 1.4, spacing=1.2)
    D.text(s, 64, 392, 620, body, 15, SONG, INK, 1.78)''')
s=s.replace('D.rect(s, x, 620, w, 23, fill=accent)','D.rect(s, x, 632, w, 23, fill=accent)')
s=s.replace('D.text(s, x + 10, 626, w - 16, t, 9, HEAD, WHITE, 1.1)','D.text(s, x + 10, 638, w - 16, t, 9, HEAD, WHITE, 1.1)')
# unit divider with arc + count block
s=s.replace('def unit_divider(no, title, sub, quote, color, count, rng):','def unit_divider(no, title, sub, quote, color, count, rng, contrast=NAVY, arc=(1330, 740, 430, 190, 280)):')
s=s.replace('''    s = D.new_slide(color)
    D.rect(s, 0, 0, 1280, 14, fill=INK)''','''    s = D.new_slide(color)
    cx, cy, r, a0, a1 = arc
    D.arc_block(s, cx, cy, r, a0, a1, 26000, contrast)
    D.rect(s, 0, 0, 1280, 14, fill=INK)''')
s=s.replace('''    D.text(s, 1040, 620, 176, count, 34, HEAD, INK, 1.0, PP_ALIGN.RIGHT)
    D.text(s, 1040, 674, 176, rng, 11, HEAD, INK, 1.2, PP_ALIGN.RIGHT)''','''    D.rect(s, 1000, 560, 216, 122, fill=contrast)
    D.text(s, 1022, 576, 176, count, 34, HEAD, WHITE, 1.0)
    D.text(s, 1022, 630, 176, rng, 11, HEAD, WHITE, 1.2)''')
# call updates with contrast arcs
s=s.replace('unit_divider("壹", "以古写今", "第一单元 · 古典诗歌的戏仿", "学其体式而活其精神，以古写今。", GOLD, "4 篇", "pp. 1—10")','unit_divider("壹", "以古写今", "第一单元 · 古典诗歌的戏仿", "学其体式而活其精神，以古写今。", GOLD, "4 篇", "pp. 1—10", NAVY, (1330, 740, 430, 190, 280))')
# page 17 arc behind bars
s=s.replace('''    # bar waveform
    heights''','''    # color-block arc behind the waveform
    D.arc_block(s, 1130, 720, 270, 185, 295, 26000, NAVY)
    # bar waveform
    heights''')
io.open(p,'w',encoding='utf-8').write(s)
