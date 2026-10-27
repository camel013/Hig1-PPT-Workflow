import re
t=open(r'_src\1111111.md',encoding='utf-8').read()
parts=re.split(r'<!-- Page (\d+) -->',t)
pages={int(parts[i]):parts[i+1] for i in range(1,len(parts),2)}
s=''.join(f'\n===== P{n} =====\n'+pages.get(n,'') for n in range(7,17))
open(r'_src\u1.txt','w',encoding='utf-8').write(s)
