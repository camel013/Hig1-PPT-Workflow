# -*- coding: utf-8 -*-
import zipfile,re
from lxml import etree
P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main'
p=r'projects\zhiwen_teaching_ppt169_20261024\蛭文_教学读本_95页_新讲解版_动效_嵌入字体_终稿.pptx'
eff={};bld=[];tr=0;short=[]
with zipfile.ZipFile(p) as z:
 print('zip',z.testzip())
 for n in sorted([n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)],key=lambda x:int(re.search(r'(\d+)',x).group(1))):
  pg=int(re.search(r'(\d+)',n).group(1));r=etree.fromstring(z.read(n));t=r.find('{%s}timing'%P)
  eff[pg]=len([e for e in t.iter('{%s}cTn'%P) if e.get('nodeType')=='withEffect']) if t is not None else 0
  if r.find('{%s}transition/{%s}fade'%(P,P)) is not None:tr+=1
  if t is not None and t.find('{%s}bldLst'%P) is not None:bld.append(pg)
  ids={e.get('id'):e for e in r.iter('{%s}cNvPr'%P)}
  for sp in t.iter('{%s}spTgt'%P):
   el=ids.get(sp.get('spid'))
   if el is None:continue
   parent=el.getparent().getparent();text=''.join(x.text or '' for x in parent.iter('{%s}t'%A)).strip()
   if 0<len(text)<=2:short.append((pg,text))
print('effects',sum(eff.values()),'pages',len(eff),'minmax',min(eff.values()),max(eff.values()),'dist',{k:list(eff.values()).count(k) for k in sorted(set(eff.values()))})
print('fade_transitions',tr,'bldLst',bld,'short',short)
