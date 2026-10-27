from PIL import Image, ImageChops
from pathlib import Path
base=Path(r'projects\zhiwen_teaching_ppt169_20261024\_anim_compare')
diffs=[]
for a in sorted((base/'before').glob('*.png')):
    b=base/'final'/a.name
    ia=Image.open(a).convert('RGB'); ib=Image.open(b).convert('RGB')
    d=ImageChops.difference(ia,ib); n=sum(1 for p in d.getdata() if max(p)>8)
    if n: diffs.append((a.name,n,d.getbbox()))
print('pages',len(list((base/'final').glob('*.png'))),'different',len(diffs),'maxpixels',max([x[1] for x in diffs],default=0))
print(diffs[:20])
