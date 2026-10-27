import zipfile, re
p=r'projects\zhiwen_teaching_ppt169_20261024\_deck_u1.pptx'
with zipfile.ZipFile(p) as z:
    xml=z.read('ppt/slides/slide2.xml').decode('utf-8')
# print first few runs with typefaces and text
for m in re.finditer(r'<a:rPr[^>]*>(.*?)</a:rPr>(.*?)</a:r>', xml, re.S):
    rpr, body = m.group(1), m.group(2)
    tf = re.findall(r'typeface="([^"]+)"', rpr)
    txt = re.sub('<[^>]+>','',body)
    if txt.strip():
        print(repr(txt[:20]), tf)
