# -*- coding: utf-8 -*-
from dataclasses import dataclass, field
from pptx import Presentation
from pptx.oxml.ns import qn
import sys, math
NS='{http://schemas.openxmlformats.org/presentationml/2006/main}'
SW,SH=1280,720

@dataclass
class R:
    node: object
    idx: int
    id: int
    name: str
    descr: str
    kind: str
    x: float; y: float; w: float; h: float
    text: str
    static: bool=False
    source: str=''
    @property
    def x2(self): return self.x+self.w
    @property
    def y2(self): return self.y+self.h
    @property
    def cx(self): return self.x+self.w/2
    @property
    def cy(self): return self.y+self.h/2

def xgap(a,b): return max(0,max(a.x,b.x)-min(a.x2,b.x2))
def ygap(a,b): return max(0,max(a.y,b.y)-min(a.y2,b.y2))
def overlap(a0,a1,b0,b1): return max(0,min(a1,b1)-max(a0,b0))

def union(bs):
    changed=True
    while changed:
        changed=False
        for i,a in enumerate(bs):
            for j in range(i+1,len(bs)):
                b=bs[j]
                if max(a.y,b.y) < min(a.y2,b.y2):
                    a.y=min(a.y,a.y); b.y=min(b.y,b.y)
                a.x=min(a.x,b.x); a.y=min(a.y,b.y); a.x2=max(a.x2,b.x2); a.y2=max(a.y2,b.y2)
                a.w=a.x2-a.x; a.h=a.y2-a.y
                a.text=(a.text+' '+b.text).strip(); a.name=a.name
                bs.pop(j); changed=True; break
            if changed: break
    return bs

@dataclass
class B:
    items: list=field(default_factory=list)
    x: float=0; y: float=0; w: float=0; h: float=0
    kind: str='body'; source: str=''
    @property
    def x2(self): return self.x+self.w
    @property
    def y2(self): return self.y+self.h
    @property
    def text(self): return ' / '.join(i.text for i in self.items if i.text)[:120]
    @property
    def ntext(self): return sum(1 for i in self.items if i.text)
    @property
    def is_container(self): return self.source.startswith('container')

def make_block(items, source=''):
    xs=[i.x for i in items]; ys=[i.y for i in items]; x2s=[i.x2 for i in items]; y2s=[i.y2 for i in items]
    kinds={i.kind for i in items}
    return B(items=list(items),x=min(xs),y=min(ys),w=max(x2s)-min(xs),h=max(y2s)-min(ys),kind='heading' if kinds=={'heading'} else 'body',source=source)
def union_block(a,b,source='merge'):
    return make_block(a.items+b.items, source or a.source)

def extract(prs):
    all=[]
    for si,s in enumerate(prs.slides,1):
        recs=[]
        for idx,sh in enumerate(s.shapes):
            cn=sh._element.find('.//'+NS+'cNvPr')
            if cn is None: continue
            d=cn.get('descr','') or ''; nm=cn.get('name','') or ''
            kind=(d.rsplit(':',1)[-1] if d.startswith('animgrp:') else ('heading' if nm.startswith('heading_') else 'body' if nm.startswith('body_') else 'graphic'))
            text=getattr(sh,'text','') or ''
            x,y,w,h=[z/9600 for z in (sh.left,sh.top,sh.width,sh.height)]
            static=nm.startswith('bg') or y>=668 or (not text and (h>=500 or (w>=1050 and h>=300) or (w>=200 and h<=8) or (w<=10 and h>=40)))
            recs.append(R(sh._element,idx,int(cn.get('id')),nm,d,kind,x,y,w,h,text,static))
        all.append(recs)
    return all

def base_blocks(recs, page):
    content=[r for r in recs if not r.static]
    # header
    header=[r for r in content if r.y < 125]
    rest=[r for r in content if r not in header]
    blocks=[]
    if header:
        blocks.append(make_block(header,'header'))
    bydes={}
    for r in rest:
        key=r.descr if r.descr.startswith('animgrp:') else f'x{id(r)}'
        bydes.setdefault(key,[]).append(r)
    for rr in bydes.values():
        blocks.append(make_block(rr,'descr'))
    # container rectangles consume blocks whose centers sit inside
    containers=[r for r in rest if not r.text and r.w>=80 and r.h>=25 and r.w*r.h>=18000 and r.name.startswith(('Rectangle','Block Arc'))]
    containers.sort(key=lambda r:r.w*r.h, reverse=True)
    used=set()
    for c in containers:
        inside=[]
        for i,b in enumerate(blocks):
            if i in used or b.source=='header': continue
            if b is None: continue
            iy=max(0,min(c.y2,b.y2)-max(c.y,b.y)); ih=max(1,b.h)
            ix=max(0,min(c.x2,b.x2)-max(c.x,b.x)); iw=max(1,b.w)
            if iy/ih>=0.55 and ix/iw>=0.30:
                inside.append(i)
        if inside:
            items=[c]
            for i in inside:
                items+=blocks[i].items; used.add(i)
            blocks.append(make_block(items,'container'))
            used.add(-1)
    blocks=[b for i,b in enumerate(blocks) if i not in used]
    # Dense tables keep their column rhythm; ordinary pages keep step rows.
    if len(blocks)<=32:
        changed=True
        while changed:
            changed=False
            for i,a in enumerate(blocks):
                for j in range(i+1,len(blocks)):
                    b=blocks[j]
                    if a.is_container or b.is_container: continue
                    yov=overlap(a.y,a.y2,b.y,b.y2)
                    cya=(a.y+a.y2)/2; cyb=(b.y+b.y2)/2
                    short = min(a.w,b.w)<=110 and min(a.h,b.h)<=40
                    mixed = bool({i.kind for i in a.items} ^ {i.kind for i in b.items})
                    near = xgap(a,b)<=14 or (short and xgap(a,b)<=62) or (mixed and xgap(a,b)<=60)
                    if yov >= 0.28*min(a.h,b.h) and abs(cya-cyb)<=22 and near and (short or mixed):
                        blocks[i]=union_block(a,b); blocks.pop(j); changed=True; break
                if changed: break
    # Then merge vertically aligned copy into columns or paragraphs.
    changed=True
    while changed:
        changed=False
        for i,a in enumerate(blocks):
            for j in range(i+1,len(blocks)):
                b=blocks[j]
                wa=max(a.w,b.w)
                ca=(a.x+a.x2)/2; cb=(b.x+b.x2)/2
                dy=ygap(a,b)
                aligned = min(abs(a.x-b.x),abs(a.x2-b.x2),abs(ca-cb)) <= max(28,0.24*min(a.w,b.w))
                if aligned and dy<=78 and wa < 0.56*SW and not (a.is_container or b.is_container):
                    blocks[i]=union_block(a,b); blocks.pop(j); changed=True; break
            if changed: break
    return blocks

def absorb_fragments(blocks):
    def compact(b): return "".join(i.text for i in b.items).strip()
    frags=[b for b in blocks if compact(b) and len(compact(b))<=2]
    if len(frags)>1:
        merged=make_block([i for b in frags for i in b.items],"fragments")
        blocks=[b for b in blocks if b not in frags]+[merged]
    changed=True
    while changed:
        changed=False
        frags=[b for b in blocks if compact(b) and len(compact(b))<=2]
        for f in frags:
            others=[b for b in blocks if b is not f and compact(b) and len(compact(b))>2]
            if not others: break
            best=min(others,key=lambda b:(max(0,max(f.x,b.x)-min(f.x2,b.x2)),max(0,max(f.y,b.y)-min(f.y2,b.y2))))
            blocks=[b for b in blocks if b is not f]
            blocks=[best if b is best else b for b in blocks]
            blocks.remove(best); blocks.append(union_block(best,f))
            changed=True; break
    return blocks

def merge_to(blocks,target):
    while len(blocks)>target:
        best=None
        for i,a in enumerate(blocks):
            for j in range(i+1,len(blocks)):
                b=blocks[j]
                dx=xgap(a,b); dy=ygap(a,b)
                # vertical reading order first; horizontal only when near
                score=math.hypot(dx,dy)+ (0 if (dx==0 or dy==0) else 18)
                if best is None or score<best[0]: best=(score,i,j)
        _,i,j=best
        blocks[i]=union_block(blocks[i],blocks[j]); blocks.pop(j)
    return blocks

prs=Presentation(r'projects\zhiwen_teaching_ppt169_20261024\蛭文_教学读本_95页_新讲解版.pptx')
data=extract(prs)
pages={int(x) for x in sys.argv[1:]}
for page in sorted(pages):
    bs=absorb_fragments(base_blocks(data[page-1],page))
    print(f'\n=== {page:02d} base={len(bs)} ===')
    for b in bs:
        print(f'[{b.x:4.0f},{b.y:3.0f},{b.w:4.0f},{b.h:3.0f}] {b.source:9} {b.text[:80]}')





from collections import Counter
counts={}
nsingle={}
for pg,recs in enumerate(data,1):
    bs=merge_to(base_blocks(recs,pg), min(6,len(base_blocks(recs,pg))))
    counts[pg]=len(bs)
    nsingle[pg]=sum(1 for b in bs if ''.join(i.text for i in b.items).strip() and len(''.join(i.text for i in b.items).strip())<=2)
print('COUNTS', Counter(counts.values()))
print('SINGLE', [(k,v) for k,v in nsingle.items() if v])
print('SMALL', [(k,v) for k,v in counts.items() if v<4])



