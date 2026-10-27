import re
t=open(r'_src\1111111.md',encoding='utf-8').read()
parts=re.split(r'<!-- Page (\d+) -->',t)
pages={int(parts[i]):parts[i+1] for i in range(1,len(parts),2)}
for n in [97,98,99,105,106,113]:
    b=pages.get(n,'')
    # strip image lines
    lines=[l for l in b.splitlines() if not l.strip().startswith('![')]
    open(fr'_src\pg{n}.txt','w',encoding='utf-8').write('\n'.join(lines))
