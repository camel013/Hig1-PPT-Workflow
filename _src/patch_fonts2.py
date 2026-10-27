import io
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_lib.py'
s=io.open(p,encoding='utf-8').read()
s=s.replace('Chinese: KaiTi for section headings, STZhongsong for body.','Chinese: FangSong for headings and labels, STZhongsong for body.')
s=s.replace('def vline(slide, x, y, s, size=28, color=INK, font=KAI, leading=1.18, bottom=None, spacing=1.5):','def vline(slide, x, y, s, size=28, color=INK, font=SONG, leading=1.18, bottom=None, spacing=1.5):')
io.open(p,'w',encoding='utf-8').write(s)

for p in [r'projects\zhiwen_teaching_ppt169_20261024\deck_front.py', r'projects\zhiwen_teaching_ppt169_20261024\deck_unit1.py']:
    s=io.open(p,encoding='utf-8').read()
    s=s.replace('KAI, FANG, SONG','FANG, SONG').replace('FANG, KAI, PAPER','FANG, PAPER')
    s=s.replace('KAI, INK','HEAD, INK').replace('KAI, CORAL','HEAD, CORAL').replace('KAI, INDIGO','HEAD, INDIGO')
    io.open(p,'w',encoding='utf-8').write(s)
