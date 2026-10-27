# -*- coding: utf-8 -*-
from PIL import Image, ImageChops
from pathlib import Path
base=Path(r'projects\zhiwen_teaching_ppt169_20261024\_anim_compare')
diffs=[]
for a in sorted((base/'before').glob('*.png')):
    b=base/'after'/a.name
    ia=Image.open(a).convert('RGB'); ib=Image.open(b).convert('RGB')
    d=ImageChops.difference(ia,ib)
    bbox=d.getbbox()
    n=sum(1 for p in d.getdata() if max(p)>8)
    if n: diffs.append((a.name,n,bbox))
print('pages',len(list((base/'before').glob('*.png'))),'different',len(diffs))
for row in diffs[:30]: print(row)
print('maxpixels', max([d[1] for d in diffs], default=0))
