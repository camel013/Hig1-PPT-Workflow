import io
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_unit1.py'
s=io.open(p,encoding='utf-8').read()
s=s.replace('def three_col(kicker, title, lead, items, accent=CORAL, note=None, bg=PAPER, dark=False):','def three_col(kicker, title, lead, items, accent=CORAL, note=None, bg=PAPER, dark=False, bottom=None):')
s=s.replace('    if note:\n        D.text(s, 64, 620, 1120, note, 11, SONG, mu, 1.45)\n    D.footer(s, p, dark=dark)','    if bottom:\n        D.rect(s, 64, 500, 1152, 78, fill=accent)\n        D.text(s, 90, 520, 1100, bottom, 18, HEAD, WHITE, 1.3)\n    if note:\n        D.text(s, 64, 604, 1120, note, 11, SONG, mu, 1.45)\n    D.footer(s, p, dark=dark)')
s=s.replace('def stagger(kicker, title, stanzas, source, accent=CORAL, ghost=None, note=None):','def stagger(kicker, title, stanzas, source, accent=CORAL, ghost=None, note=None, ghost_lines=None):')
s=s.replace('    if ghost:\n        D.text(s, 820, 150, 390, ghost, 17, SONG, RULE, 1.8)','    if ghost_lines:\n        for i, line in enumerate(ghost_lines):\n            D.text(s, 820, 150 + i * 28, 400, line, 16.5, SONG, RULE, 1.5)')
s=s.replace('    span = 1140\n    gap = span / (n - 1)\n    D.hline(s, 70, 1210, 330, RULE, 1.0)\n    for i, (h, b) in enumerate(items):\n        x = 70 + i * gap','    start, end, labw = 80, 1216, 180\n    gap = (end - labw - start) / (n - 1)\n    D.hline(s, 80, 1216, 330, RULE, 1.0)\n    for i, (h, b) in enumerate(items):\n        x = start + i * gap')
s=s.replace('D.text(s, 900, 168, 300, "“有声”不是热闹，\\n是另一种不能退出。", 21, HEAD, GOLD, 1.5)','D.text(s, 900, 166, 320, "“有声”不是热闹，", 18, HEAD, GOLD, 1.5)\n    D.text(s, 900, 204, 320, "是另一种不能退出。", 18, HEAD, GOLD, 1.5)')
s=s.replace('D.vline(s, 1170 - i * 62, 156, line, 27, INK, SONG, 1.26, spacing=1.5)','D.vline(s, 1170 - i * 62, 156, line, 24, INK, SONG, 1.22, spacing=1.5)')
# ghost call for 再别
s=s.replace('GOLD, "轻轻的我走了，\\n正如我轻轻的来；\\n我轻轻的招手，\\n作别西天的云彩。",','GOLD, ghost_lines=["轻轻的我走了，", "正如我轻轻的来；", "我轻轻的招手，", "作别西天的云彩。"],')
s=s.replace('GOLD, "我是天空里的一片云，\\n偶尔投影在你的波心——\\n你不必讶异，\\n更无须欢喜——",','GOLD, ghost_lines=["我是天空里的一片云，", "偶尔投影在你的波心——", "你不必讶异，", "更无须欢喜——"],')
io.open(p,'w',encoding='utf-8').write(s)
