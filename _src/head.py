import re
t=open(r'_src\1111111.md',encoding='utf-8').read()
parts=re.split(r'<!-- Page (\d+) -->',t)
pages={int(parts[i]):parts[i+1] for i in range(1,len(parts),2)}
out=[]
for n in sorted(pages):
    if 92<=n<=115:
        b=pages[n]
        lines=[l.strip() for l in b.splitlines() if l.strip() and not l.strip().startswith('![')]
        out.append(f'===== P{n} =====\n'+'\n'.join(lines[:12]))
open(r'_src\u4b_head.txt','w',encoding='utf-8').write('\n'.join(out))
