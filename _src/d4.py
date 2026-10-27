import re
t=open(r'_src\1111111.md',encoding='utf-8').read()
parts=re.split(r'<!-- Page (\d+) -->',t)
pages={int(parts[i]):parts[i+1] for i in range(1,len(parts),2)}
for a,b,name in [(64,66,'p6466'),(70,74,'p7074'),(93,113,'p93113')]:
    open(fr'_src\{name}.txt','w',encoding='utf-8').write(''.join(f'\n===== P{n} =====\n'+pages.get(n,'') for n in range(a,b+1)))
