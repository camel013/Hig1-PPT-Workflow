import re
p=r'_src\1111111.md'
t=open(p,encoding='utf-8').read()
parts=re.split(r'<!-- Page (\d+) -->',t)
# parts[0] preamble, then num, body...
pages={}
for i in range(1,len(parts),2):
    pages[int(parts[i])]=parts[i+1]
print('pages',min(pages),max(pages),len(pages))
for n in sorted(pages):
    b=pages[n]
    lines=[l.strip() for l in b.splitlines() if l.strip() and not l.strip().startswith('![')]
    head=' | '.join(lines[:4])[:150]
    print(f'{n:3d} {len(b):5d} {head}')
