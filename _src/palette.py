from PIL import Image
from collections import Counter
p=r'C:\Users\camel\Documents\Tencent Files\3234093593\nt_qq\nt_data\Pic\2026-10\Ori\2efcae740bffecc1bbbb1ea87d1fe9c3.jpg'
im=Image.open(p).convert('RGB')
print(im.size)
w,h=im.size
q=im.quantize(colors=12).convert('RGB')
c=Counter(q.getdata())
for col,n in c.most_common(12):
    print('#%02X%02X%02X'%col, round(n/(w*h)*100,1))
# sample specific areas
for name,box in {'left_coral':(50,80,300,500),'blue':(900,400,1100,650),'deepblue':(750,120,820,300),'gold':(1280,520,1500,610),'teal':(660,40,740,200),'white':(1290,190,1470,350)}.items():
    reg=im.crop(box).resize((30,30))
    c2=Counter(reg.getdata()).most_common(3)
    print(name, ['#%02X%02X%02X'%k for k,_ in c2])
