# -*- coding: utf-8 -*-
import zipfile, re
from lxml import etree
P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
path=r'projects\zhiwen_teaching_ppt169_20261024\_grouped_anim_tmp.pptx'
short=[]; counts=[]; bld=[]; totals=0
with zipfile.ZipFile(path) as z:
    bad=z.testzip(); print('zip',bad)
    for name in sorted([n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)], key=lambda x:int(re.search(r'(\d+)',x).group(1))):
        root=etree.fromstring(z.read(name)); pg=int(re.search(r'(\d+)',name).group(1))
        tim=root.find('{%s}timing'%P); tr=root.find('{%s}transition'%P)
        if tim is None or tr is None: print('MISSING',pg); continue
        if tim.find('{%s}bldLst'%P) is not None: bld.append(pg)
        ids={e.get('id'):e for e in root.iter('{%s}cNvPr'%P)}
        tids=[e.get('spid') for e in tim.iter('{%s}spTgt'%P) if e.get('spid')]
        counts.append(len(tids)); totals+=len(tids)
        for sid in tids:
            el=ids.get(sid)
            if el is None: print('NOID',pg,sid); continue
            parent=el.getparent().getparent()
            text=''.join(t.text or '' for t in parent.iter('{%s}t'%A)).strip()
            if 0<len(text)<=2: short.append((pg,text,sid,parent.tag.split('}')[-1]))
print('slides',len(counts),'units',totals,'minmax',min(counts),max(counts),'bldLst',bld)
print('short',short)
