import re
lines=open(r'_src\1111111.md',encoding='utf-8').read().splitlines()
for i in [2280,2358,2410,2504]:
    print('\n'.join(lines[i-1:i+8]))
    print('---')
