import zipfile, re, json
p=r'projects\zhiwen_teaching_ppt169_20261024\_deck_u1.pptx'
out=[]
with zipfile.ZipFile(p) as z:
    xml=z.read('ppt/slides/slide2.xml').decode('utf-8')
for m in re.finditer(r'<a:rPr[^>]*>(.*?)</a:rPr>(.*?)</a:r>', xml, re.S):
    rpr, body = m.group(1), m.group(2)
    tf = re.findall(r'typeface="([^"]+)"', rpr)
    txt = re.sub('<[^>]+>','',body)
    if txt.strip():
        out.append({'text':txt[:24],'fonts':tf})
open(r'_src\fontcheck.json','w',encoding='utf-8').write(json.dumps(out,ensure_ascii=False,indent=1))
