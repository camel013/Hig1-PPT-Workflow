import re
t=open(r'_src\1111111.md',encoding='utf-8').read()
parts=re.split(r'<!-- Page (\d+) -->',t)
pages={int(parts[i]):parts[i+1] for i in range(1,len(parts),2)}
for a,b,name in [(30,47,'u2b'),(69,79,'u4a'),(93,107,'u4b')]:
    s=''.join(f'\n===== P{n} =====\n'+pages.get(n,'') for n in range(a,b+1))
    open(fr'_src\{name}.txt','w',encoding='utf-8').write(s)
