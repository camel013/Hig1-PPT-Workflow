import re
t=open(r'_src\1111111.md',encoding='utf-8').read()
parts=re.split(r'<!-- Page (\d+) -->',t)
pages={int(parts[i]):parts[i+1] for i in range(1,len(parts),2)}
out=[]
for n in sorted(pages):
    b=pages[n]
    lines=[l.strip() for l in b.splitlines() if l.strip() and not l.strip().startswith('![')]
    head=' | '.join(lines[:4])[:150]
    out.append(f'{n:3d} {len(b):5d} {head}')
open(r'_src\index.txt','w',encoding='utf-8').write('\n'.join(out))
