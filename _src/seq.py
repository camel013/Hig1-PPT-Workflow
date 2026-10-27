import re
t=open(r'_src\1111111.md',encoding='utf-8').read()
parts=re.split(r'<!-- Page (\d+) -->',t)
pages={int(parts[i]):parts[i+1] for i in range(1,len(parts),2)}
for n in [5,119,120]:
    b=pages.get(n,'')
    lines=[l for l in b.splitlines() if not l.strip().startswith('![')]
    print(f'===== P{n} =====')
    print('\n'.join(lines))
