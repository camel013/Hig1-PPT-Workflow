import io
# unit1
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_unit1.py'
s=io.open(p,encoding='utf-8').read()
s=s.replace('D.text(s, 654, 470, 500, qsrc, 10.5, SONG, mu, 1.4)','D.text(s, 654, 470, 500, qsrc, 10.5, HEAD, mu, 1.4)')
s=s.replace('D.text(s, x + 11, 616, w - 18, t, 9.5, SONG, WHITE, 1.1)','D.text(s, x + 11, 616, w - 18, t, 9.5, HEAD, WHITE, 1.1)')
s=s.replace('D.text(s, 64, 270, 130, "篇目", 9, SONG, MUTED, 1.2, spacing=2)','D.text(s, 64, 270, 130, "篇目", 9, HEAD, MUTED, 1.2, spacing=2)')
s=s.replace('D.text(s, 200, 204, 520, author, 12.5, SONG, MUTED, 1.4, spacing=1.2)','D.text(s, 200, 204, 520, author, 12.5, HEAD, MUTED, 1.4, spacing=1.2)')
s=s.replace('D.text(s, 792, 476, 380, qsrc, 10, SONG, WHITE, 1.4)','D.text(s, 792, 476, 380, qsrc, 10, HEAD, WHITE, 1.4)')
s=s.replace('D.text(s, x + 10, 626, w - 16, t, 9, SONG, WHITE, 1.1)','D.text(s, x + 10, 626, w - 16, t, 9, HEAD, WHITE, 1.1)')
s=s.replace('D.text(s, x + 10, y0 + 10, w - 14, h, 10.5, SONG, WHITE, 1.2, spacing=1.2)','D.text(s, x + 10, y0 + 10, w - 14, h, 10.5, HEAD, WHITE, 1.2, spacing=1.2)')
s=s.replace('D.text(s, 64, 632, 500, source, 9.5, SONG, MUTED, 1.2)','D.text(s, 64, 632, 500, source, 9.5, HEAD, MUTED, 1.2)')
s=s.replace('D.text(s, 700, 626, 480, source, 9.5, SONG, MUTED, 1.2, PP_ALIGN.RIGHT)','D.text(s, 700, 626, 480, source, 9.5, HEAD, MUTED, 1.2, PP_ALIGN.RIGHT)')
s=s.replace('D.text(s, x, yy, 320, line, 17, HEAD, INK, 1.55)','D.text(s, x, yy, 320, line, 17, SONG, INK, 1.55)')
s=s.replace('D.text(s, 1040, 674, 176, rng, 11, TNR, INK, 1.2, PP_ALIGN.RIGHT)','D.text(s, 1040, 674, 176, rng, 11, HEAD, INK, 1.2, PP_ALIGN.RIGHT)')
s=s.replace('D.text(s, 1040, 620, 176, count, 34, TNR, INK, 1.0, PP_ALIGN.RIGHT)','D.text(s, 1040, 620, 176, count, 34, HEAD, INK, 1.0, PP_ALIGN.RIGHT)')
io.open(p,'w',encoding='utf-8').write(s)
# front
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_front.py'
s=io.open(p,encoding='utf-8').read()
s=s.replace('D.text(s, 64, 640, 540, "4 个单元 · 24 篇课文 · 3 首古诗词 · 序与跋", 11, TNR, CORAL_LIGHT, 1.3)','D.text(s, 64, 640, 540, "4 个单元 · 24 篇课文 · 3 首古诗词 · 序与跋", 11, HEAD, CORAL_LIGHT, 1.3)')
s=s.replace('D.text(s, 900, 610, 316, "24 篇 + 3 首 + 序跋", 18, TNR, CORAL, 1.2, PP_ALIGN.RIGHT)','D.text(s, 900, 610, 316, "24 篇 + 3 首 + 序跋", 18, HEAD, CORAL, 1.2, PP_ALIGN.RIGHT)')
# captions in front
s=s.replace('D.text(s, 892, 414, 290, "—— 编者的话", 11, SONG, CORAL_LIGHT, 1.4)','D.text(s, 892, 414, 290, "—— 编者的话", 11, HEAD, CORAL_LIGHT, 1.4)')
s=s.replace('D.text(s, x, y + 10, 330, k, 10, SONG, MUTED, 1.2, spacing=1.5)','D.text(s, x, y + 10, 330, k, 10, HEAD, MUTED, 1.2, spacing=1.5)')
io.open(p,'w',encoding='utf-8').write(s)
