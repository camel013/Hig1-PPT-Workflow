import io, re
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_lib.py'
s=io.open(p,encoding='utf-8').read()
s=s.replace('KAI = "楷体"\nSONG = "华文中宋"', 'KAI = "楷体"\nFANG = "仿宋"\nSONG = "华文中宋"')
s=s.replace('    KAI: r"C:\\Windows\\Fonts\\simkai.ttf",', '    KAI: r"C:\\Windows\\Fonts\\simkai.ttf",\n    FANG: r"C:\\Windows\\Fonts\\simfang.ttf",')
s=s.replace('def page_title(slide, page, kicker, title, dark=False, color=None, rule=True):\n    fg = WHITE if dark else INK\n    meta(slide, page, kicker, dark)\n    text(slide, 64, 64, 1000, title, 34, KAI, color or fg, 1.2)', 'def page_title(slide, page, kicker, title, dark=False, color=None, rule=True):\n    fg = WHITE if dark else INK\n    meta(slide, page, kicker, dark)\n    text(slide, 64, 64, 1000, title, 34, FANG, color or fg, 1.2)')
s=s.replace('def unit_no(slide, x, y, s, size=220, color=GOLD, outline=False):', 'def unit_no(slide, x, y, s, size=220, color=GOLD, outline=False):')
s=s.replace('    return text(slide, x, y, size * 1.15, s, size, KAI, color, 1.0)', '    return text(slide, x, y, size * 1.15, s, size, FANG, color, 1.0)')
io.open(p,'w',encoding='utf-8').write(s)

p=r'projects\zhiwen_teaching_ppt169_20261024\build_front.py'
s=io.open(p,encoding='utf-8').read()
# drop decorative arcs / seal
s='\n'.join(l for l in s.splitlines() if 'D.arc(' not in l)
# headings to FANG
s=s.replace('KAI', 'HEAD')
s=s.replace('from deck_lib import (', 'from deck_lib import (\n                      FANG,')
s=s.replace('P = pathlib.Path(__file__).parent', 'HEAD = FANG\nP = pathlib.Path(__file__).parent')
# restore book-title display to KaiTi
s=s.replace('D.text(s, 958, 84, 240, "蛭", 96, HEAD, INK, 1.0)', 'D.text(s, 958, 84, 240, "蛭", 96, KAI, INK, 1.0)')
s=s.replace('D.text(s, 958, 186, 240, "文", 96, HEAD, INK, 1.0)', 'D.text(s, 958, 186, 240, "文", 96, KAI, INK, 1.0)')
s=s.replace('D.text(s, 80, 190, 300, "蛭", 250, HEAD, CORAL, 1.0)', 'D.text(s, 80, 190, 300, "蛭", 250, KAI, CORAL, 1.0)')
s=s.replace('D.text(s, 380, 190, 300, "文", 250, HEAD, INDIGO, 1.0)', 'D.text(s, 380, 190, 300, "文", 250, KAI, INDIGO, 1.0)')
io.open(p,'w',encoding='utf-8').write(s)
