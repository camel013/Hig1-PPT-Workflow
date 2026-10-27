import io
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_lib.py'
s=io.open(p,encoding='utf-8').read()
# remove KaiTi family entirely
s=s.replace('    KAI: r"C:\\Windows\\Fonts\\simkai.ttf",\n','')
s=s.replace('KAI = "楷体"\n','')
# add mixed-run emitter before draw_line
marker='def draw_line(slide, x, y, line, size, font, color, align=PP_ALIGN.LEFT, width=1000,\n'
emitter='''LATIN_SEG = re.compile(r"[A-Za-z0-9]+(?:[.'’%/:\\-–][A-Za-z0-9]+)*")


def _script_segments(s):
    out = []
    i = 0
    while i < len(s):
        m = LATIN_SEG.match(s, i)
        if m:
            out.append((m.group(0), True))
            i = m.end()
            continue
        j = i
        while j < len(s) and not LATIN_SEG.match(s, j):
            j += 1
        out.append((s[i:j], False))
        i = j
    return out


def _emit_runs(p, s, size, base, color, bold=False, spacing=0):
    for seg, is_latin in _script_segments(s):
        if not seg:
            continue
        r = p.add_run()
        r.text = seg
        if is_latin:
            _set_font(r, TNR, base, size, color, bold, spacing)
        else:
            _set_font(r, base, base, size, color, bold, spacing)


'''
s=s.replace(marker, emitter+marker)
# replace single run in draw_line
s=s.replace('''    r = p.add_run()
    r.text = s
    _set_font(r, TNR if all(k == "latin" or k == "space" for _, k in toks) else font,
              font, size, color, bold, spacing)
    return lw''','''    _emit_runs(p, s, size, font, color, bold, spacing)
    return lw''')
# replace rich run grouping
old='''        j = 0
        while j < len(line):
            color = line[j][2]
            bold = line[j][3]
            buf = []
            while j < len(line) and line[j][2] == color and line[j][3] == bold:
                buf.append(line[j][0])
                j += 1
            r = p.add_run()
            r.text = "".join(buf)
            _set_font(r, TNR if all(ch.isascii() and (ch.isalnum() or ch in " .,:;/%-'’") for ch in "".join(buf)) else font,
                      font, size, color, bold)'''
new='''        j = 0
        while j < len(line):
            color = line[j][2]
            bold = line[j][3]
            buf = []
            while j < len(line) and line[j][2] == color and line[j][3] == bold:
                buf.append(line[j][0])
                j += 1
            _emit_runs(p, "".join(buf), size, font, color, bold)'''
s=s.replace(old,new)
# meta/footer labels -> FangSong
s=s.replace('text(slide, 64, 30, 720, label, 9, SONG, fg, 1.2, spacing=2.2)','text(slide, 64, 30, 720, label, 9, FANG, fg, 1.2, spacing=2.2)')
s=s.replace('text(slide, 64, 690, 640, "《蛭文》教学读本 / 文本细读与写作", 8.5, SONG, fg, 1.2, spacing=1.2)','text(slide, 64, 690, 640, "《蛭文》教学读本 / 文本细读与写作", 8.5, FANG, fg, 1.2, spacing=1.2)')
io.open(p,'w',encoding='utf-8').write(s)
