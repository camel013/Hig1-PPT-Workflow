# -*- coding: utf-8 -*-
import zipfile, re
from lxml import etree
P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main'
path=r'projects\zhiwen_teaching_ppt169_20261024\_grouped_anim_tmp.pptx'
effects={}; short=[]; bld=[]; trans=0; groups=0
with zipfile.ZipFile(path) as z:
    print('zip_test', z.testzip())
    names=[n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)]
    for name in sorted(names,key=lambda x:int(re.search(r'(\d+)',x).group(1))):
        pg=int(re.search(r'(\d+)',name).group(1)); root=etree.fromstring(z.read(name))
        tim=root.find('{%s}timing'%P); effects[pg]=len([e for e in tim.iter('{%s}cTn'%P) if e.get('nodeType')=='withEffect'])
        if root.find('{%s}transition/{%s}fade'%(P,P)) is not None: trans+=1
        if tim.find('{%s}bldLst'%P) is not None: bld.append(pg)
        ids={e.get('id'):e for e in root.iter('{%s}cNvPr'%P)}
        groups+=len([e for e in root.iter('{%s}grpSp'%P) if (e.find('.//{%s}cNvPr'%P) is not None and e.find('.//{%s}cNvPr'%P).get('descr')=='semantic-animation-unit')])
        for sp in tim.iter('{%s}spTgt'%P):
            sid=sp.get('spid'); el=ids.get(sid)
            if el is None: continue
            parent=el.getparent().getparent(); text=''.join(t.text or '' for t in parent.iter('{%s}t'%A)).strip()
            if 0<len(text)<=2: short.append((pg,text))
print('slides',len(effects),'effects',sum(effects.values()),'min',min(effects.values()),'max',max(effects.values()))
print('distribution',{k:list(effects.values()).count(k) for k in sorted(set(effects.values()))})
print('transitions',trans,'bldLst',bld,'semantic_groups',groups,'short_targets',short)
