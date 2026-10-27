import re
t=open(r'_src\1111111.md',encoding='utf-8').read()
parts=re.split(r'<!-- Page (\d+) -->',t)
pages={int(parts[i]):parts[i+1] for i in range(1,len(parts),2)}
for n in [72,73,74,79,85,92,103,104,112,113]:
    b=pages.get(n,'')
    lines=[l for l in b.splitlines() if not l.strip().startswith('![')]
    open(fr'_src\pp{n}.txt','w',encoding='utf-8').write('\n'.join(lines))
