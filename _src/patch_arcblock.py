import io
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_lib.py'
s=io.open(p,encoding='utf-8').read()
add='''

def disc(slide, cx, cy, r, fill=TEAL):
    sp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(int((cx - r) * PX)), Emu(int((cy - r) * PX)),
                                Emu(int(2 * r * PX)), Emu(int(2 * r * PX)))
    sp.fill.solid()
    sp.fill.fore_color.rgb = rgb(fill)
    sp.line.fill.background()
    sp.shadow.inherit = False
    st = sp._element.find(qn("p:style"))
    if st is not None:
        sp._element.remove(st)
    return sp


def arc_block(slide, cx, cy, r, a0, a1, thick=22000, fill=TEAL):
    """A filled annular sector. Angles in degrees, 0 = 3 o'clock, clockwise.

    thick is the ring thickness in 1/1000 of the radius (22000 = 22%).
    """
    sp = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, Emu(int((cx - r) * PX)),
                                Emu(int((cy - r) * PX)), Emu(int(2 * r * PX)),
                                Emu(int(2 * r * PX)))
    sp.fill.solid()
    sp.fill.fore_color.rgb = rgb(fill)
    sp.line.fill.background()
    sp.shadow.inherit = False
    st = sp._element.find(qn("p:style"))
    if st is not None:
        sp._element.remove(st)
    geom = sp._element.spPr.find(qn("a:prstGeom"))
    av = geom.find(qn("a:avLst"))
    for g in list(av):
        av.remove(g)
    for name, val in (("adj1", int(a0 * 60000)), ("adj2", int(a1 * 60000)), ("adj3", int(thick))):
        gd = etree.SubElement(av, qn("a:gd"))
        gd.set("name", name)
        gd.set("fmla", f"val {val}")
    return sp
'''
s=s.replace('\ndef new_slide(bg=PAPER):', add+'\n\ndef new_slide(bg=PAPER):')
io.open(p,'w',encoding='utf-8').write(s)
