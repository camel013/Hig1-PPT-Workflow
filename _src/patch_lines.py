import io
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_lib.py'
s=io.open(p,encoding='utf-8').read()
s=s.replace('def hline(slide, x1, x2, y, color=RULE, w=1.0):\n    return rect(slide, x1, y, x2 - x1, w, fill=color)','def hline(slide, x1, x2, y, color=RULE, w=2.0):\n    return rect(slide, x1, y, x2 - x1, max(w, 2.0), fill=color)')
s=s.replace('def vrule(slide, x, y1, y2, color=RULE, w=1.0):\n    return rect(slide, x, y1, w, y2 - y1, fill=color)','def vrule(slide, x, y1, y2, color=RULE, w=2.0):\n    return rect(slide, x, y1, max(w, 2.0), y2 - y1, fill=color)')
s=s.replace('        sp.line.width = Emu(int(line_w * PX))','        sp.line.width = Emu(int(max(line_w, 2.0) * PX))')
s=s.replace('        sp.line.width = Emu(int(width * PX))','        sp.line.width = Emu(int(max(width, 2.0) * PX))')
io.open(p,'w',encoding='utf-8').write(s)
