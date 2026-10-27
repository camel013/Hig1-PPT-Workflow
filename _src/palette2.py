from PIL import Image
from collections import Counter
p=r'C:\Users\camel\Documents\Tencent Files\3234093593\nt_qq\nt_data\Pic\2026-10\Ori\2efcae740bffecc1bbbb1ea87d1fe9c3.jpg'
im=Image.open(p).convert('RGB')
for name,box in {'teal':(820,60,930,230),'gold':(1540,630,1880,720),'blue':(1100,360,1350,700),'coral':(120,150,400,900),'deepblue':(900,150,1050,350),'white':(1560,200,1820,420)}.items():
    reg=im.crop(box).resize((40,40))
    c=Counter(reg.getdata()).most_common(3)
    print(name, ['#%02X%02X%02X'%k for k,_ in c])
