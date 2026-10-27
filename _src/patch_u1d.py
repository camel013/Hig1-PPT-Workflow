import io
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_unit1.py'
s=io.open(p,encoding='utf-8').read()
s=s.replace('GOLD, ghost_lines=["轻轻的我走了，", "正如我轻轻的来；", "我轻轻的招手，", "作别西天的云彩。"],\n            "底文越柔', 'GOLD, ghost_lines=["轻轻的我走了，", "正如我轻轻的来；", "我轻轻的招手，", "作别西天的云彩。"],\n            note="底文越柔')
s=s.replace('GOLD, ghost_lines=["我是天空里的一片云，", "偶尔投影在你的波心——", "你不必讶异，", "更无须欢喜——"],\n            "徐志摩', 'GOLD, ghost_lines=["我是天空里的一片云，", "偶尔投影在你的波心——", "你不必讶异，", "更无须欢喜——"],\n            note="徐志摩')
io.open(p,'w',encoding='utf-8').write(s)
