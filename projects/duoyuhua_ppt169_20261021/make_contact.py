from PIL import Image, ImageDraw
import pathlib
p=pathlib.Path(r"C:\Users\camel\Documents\ChatGPT\Hig1\projects\duoyuhua_ppt169_20261021")
files=sorted((p/"pp_check").glob("s*.png"))
cols,rows=4,3
tw,th=480,270
gap=14
labelh=26
W=cols*tw+(cols+1)*gap
H=rows*(th+labelh)+(rows+1)*gap
sheet=Image.new("RGB",(W,H),(238,236,231))
d=ImageDraw.Draw(sheet)
for i,f in enumerate(files):
    r,c=divmod(i,cols)
    x=gap+c*(tw+gap)
    y=gap+r*(th+labelh+gap)
    im=Image.open(f).convert("RGB").resize((tw,th),Image.LANCZOS)
    sheet.paste(im,(x,y))
    d.text((x+4,y+th+5),f"{i+1:02d}",fill=(90,90,88))
out=p/"preview"/"contact_sheet.png"
out.parent.mkdir(exist_ok=True)
sheet.save(out,quality=95)
print(out, sheet.size, len(files))
