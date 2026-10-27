# -*- coding: utf-8 -*-
"""Add test font embedding to the 4-variant font sample."""
import os, re, shutil, zipfile, tempfile
import sys
from pathlib import Path
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
sys.path.insert(0, str(Path(__file__).parent))
from eotutils import make_eot

src = Path(r'_src\font_render_test.pptx')
out = Path(r'_src\font_render_test_embed.pptx')
ttf = Path(r'_src\STZHONGS_subset.ttf')

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

with zipfile.ZipFile(src) as zin:
    entries = {n: zin.read(n) for n in zin.namelist()}

# presentation.xml
pres = etree.fromstring(entries['ppt/presentation.xml'])
root = pres
for child in list(root):
    if child.tag == f'{{{P_NS}}}embeddedFontLst':
        root.remove(child)
emb = etree.Element(f'{{{P_NS}}}embeddedFontLst')
ef = etree.SubElement(emb, f'{{{P_NS}}}embeddedFont')
font = etree.SubElement(ef, f'{{{P_NS}}}font')
font.set('typeface', 'STZhongsong')
reg = etree.SubElement(ef, f'{{{P_NS}}}regular')
reg.set(f'{{{R_NS}}}id', 'rIdFont1')
# insert after notesSz
idx = next((i+1 for i,c in enumerate(list(root)) if c.tag == f'{{{P_NS}}}notesSz'), 0)
root.insert(idx, emb)
entries['ppt/presentation.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

# rels
rels = etree.fromstring(entries['ppt/_rels/presentation.xml.rels'])
rids = {e.get('Id') for e in rels}
rid = 'rIdFont1'
while rid in rids:
    rid = 'rIdFont' + str(int(rid[7:])+1) if rid[7:].isdigit() else rid + 'x'
# use the id we wrote
for e in rels:
    if e.get('Id') == 'rIdFont1':
        rels.remove(e)
rel = etree.SubElement(rels, f'{{{PKG_REL_NS}}}Relationship')
rel.set('Id', 'rIdFont1')
rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font')
rel.set('Target', 'fonts/font1.fntdata')
entries['ppt/_rels/presentation.xml.rels'] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)

# content types
ct = etree.fromstring(entries['[Content_Types].xml'])
has = any(e.get('Extension') == 'fntdata' for e in ct if e.tag == f'{{{CT_NS}}}Default')
if not has:
    d = etree.Element(f'{{{CT_NS}}}Default')
    d.set('Extension', 'fntdata')
    d.set('ContentType', 'application/x-fontdata')
    ct.insert(0, d)
entries['[Content_Types].xml'] = etree.tostring(ct, xml_declaration=True, encoding='UTF-8', standalone=True)

entries['ppt/fonts/font1.fntdata'] = make_eot(ttf)
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
    for name, data in entries.items():
        zout.writestr(name, data)
print(out)
