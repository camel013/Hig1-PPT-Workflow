import pathlib
from fontTools.ttLib import TTFont, newTable
from fontTools.pens.cu2quPen import Cu2QuPen
from fontTools.pens.ttGlyphPen import TTGlyphPen

def glyphs_to_quadratic(glyphs, max_err=1.0, reverse_direction=True):
    quad = {}
    for name in glyphs.keys():
        pen = TTGlyphPen(glyphs)
        glyphs[name].draw(Cu2QuPen(pen, max_err, reverse_direction=reverse_direction))
        quad[name] = pen.glyph()
    return quad

def otf_to_ttf(font):
    order = font.getGlyphOrder()
    font["loca"] = newTable("loca")
    font["glyf"] = glyf = newTable("glyf")
    glyf.glyphOrder = order
    glyf.glyphs = glyphs_to_quadratic(font.getGlyphSet())
    del font["CFF "]
    glyf.compile(font)
    maxp = newTable("maxp")
    maxp.tableVersion = 0x00010000
    maxp.numGlyphs = len(order)
    maxp.maxZones = 1
    maxp.maxTwilightPoints = 0
    maxp.maxStorage = 0
    maxp.maxFunctionDefs = 0
    maxp.maxInstructionDefs = 0
    maxp.maxStackElements = 0
    maxp.maxSizeOfInstructions = 0
    maxp.maxComponentElements = 0
    maxp.compile(font)
    font["maxp"] = maxp
    post = font["post"]
    post.formatType = 2.0
    post.extraNames = []
    post.mapping = {}
    post.glyphOrder = order
    font.sfntVersion = "\000\001\000\000"

jobs = [
    (r"C:\Users\camel\Documents\ChatGPT\Hig1\.fonts_tmp\trajan\TrajanPro-Regular.otf",
     r"C:\Users\camel\Documents\ChatGPT\Hig1\.fonts_tmp\TrajanPro-Regular.ttf"),
    (r"C:\Users\camel\Documents\ChatGPT\Hig1\.fonts_tmp\tnr\Times-Newer-Roman-Regular\TimesNewerRoman-Regular-2.otf",
     r"C:\Users\camel\Documents\ChatGPT\Hig1\.fonts_tmp\TimesNewerRoman-Regular.ttf"),
]
for src, dst in jobs:
    f = TTFont(src)
    otf_to_ttf(f)
    f.save(dst)
    check = TTFont(dst)
    names = {}
    for n in check["name"].names:
        if n.nameID in (1,2) and n.platformID == 3:
            names.setdefault(n.nameID, n.toUnicode())
    print(dst, pathlib.Path(dst).stat().st_size, sorted(k for k in check.keys()), names, check["OS/2"].fsType)
