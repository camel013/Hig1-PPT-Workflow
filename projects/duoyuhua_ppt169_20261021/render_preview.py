# -*- coding: utf-8 -*-
import pathlib, tempfile, pymupdf
p = pathlib.Path(r"C:\Users\camel\Documents\ChatGPT\Hig1\projects\duoyuhua_ppt169_20261021")
out = p/"preview"; out.mkdir(exist_ok=True)
tmp = pathlib.Path(tempfile.mkdtemp())
for svg in sorted((p/"svg_output").glob("*.svg")):
    c = svg.read_text(encoding="utf-8")
    c = c.replace("华文楷体", "KaiTi").replace("华文中宋", "SimSun").replace("等线", "DengXian")
    f = tmp/(svg.name)
    f.write_text(c, encoding="utf-8")
    doc = pymupdf.open(f)
    pix = doc[0].get_pixmap(dpi=72)
    pix.save(str(out/(svg.stem+".png")))
print("rendered", len(list(out.glob("slide_*.png"))))