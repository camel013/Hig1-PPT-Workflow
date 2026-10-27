from PIL import Image, ImageChops, ImageEnhance
from pathlib import Path
base=Path(r'projects\zhiwen_teaching_ppt169_20261024\_anim_compare')
for name in ['slide-10.png','slide-44.png','slide-81.png']:
 a=Image.open(base/'before'/name).convert('RGB'); b=Image.open(base/'after'/name).convert('RGB')
 d=ImageChops.difference(a,b); d=ImageEnhance.Contrast(d).enhance(4)
 d.save(base/('diff-'+name))
 print(name, d.getbbox())
