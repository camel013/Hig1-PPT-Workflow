# -*- coding: utf-8 -*-
import sys, pathlib
sys.path.insert(0, r'projects\zhiwen_teaching_ppt169_20261024')
import deck_lib as D
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu
from pptx.oxml.ns import qn
from lxml import etree

def block_arc(s, x, y, w, h, a0, a1, thick, fill):
    sp = s.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, Emu(int(x*D.PX)), Emu(int(y*D.PX)), Emu(int(w*D.PX)), Emu(int(h*D.PX)))
    sp.fill.solid(); sp.fill.fore_color.rgb = D.rgb(fill)
    sp.line.fill.background(); sp.shadow.inherit=False
    st=sp._element.find(qn('p:style'))
    if st is not None: sp._element.remove(st)
    av = sp._element.spPr.find(qn('a:prstGeom')).find(qn('a:avLst'))
    for g in list(av): av.remove(g)
    for name,val in [('adj1',a0*60000),('adj2',a1*60000),('adj3',thick)]:
        gd=etree.SubElement(av, qn('a:gd')); gd.set('name',name); gd.set('fmla','val %d'%int(val))
    return sp

s=D.new_slide('F3EDE4')
D.rect(s,0,0,640,720,fill='D56A58')
D.rect(s,640,0,640,720,fill='1B4F8B')
block_arc(s, 400, 300, 600, 600, 90, 180, 28000, '649788')
block_arc(s, 500, 380, 520, 520, 180, 270, 18000, 'B09263')
block_arc(s, 100, 420, 500, 500, 270, 360, 22000, '0C182B')
D.arc(s, 1000, 200, 240, 240, 0, 360, 'D56A58', 1.0)
D.prs.save(r'projects\zhiwen_teaching_ppt169_20261024\_arc_test.pptx')
