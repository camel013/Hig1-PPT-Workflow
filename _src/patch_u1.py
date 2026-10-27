import io
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_unit1.py'
s=io.open(p,encoding='utf-8').read()
s=s.replace('"“嘉陵江头夜学文，\\n熏风热流夏恼人。”"','"“嘉陵江头夜学文，熏风热流夏恼人。”"')
s=s.replace('D.text(s, 700, 150, 520, ghost, 20, SONG, RULE, 1.75)','D.text(s, 820, 150, 390, ghost, 17, SONG, RULE, 1.8)')
s=s.replace('x0, y0 = 64, 228','x0, y0 = 64, 210')
s=s.replace('row_h=68):','row_h=70):')
s=s.replace('D.hline(s, 64, 780, y + 86, RULE, 0.7)','D.hline(s, 64, 800, y + 88, RULE, 0.7)')
s=s.replace('["水蛭，蚂蟥，脑瓜小肚儿大。", "化学测试乱如麻，全仗你叫声哇。", "生听了愁眉，徒听了甚怕。", "哪里去辨什么真共假？", "眼见得收翻了这家，没抵住那家，只吸得血尽人逃罢！"]','["水蛭，蚂蟥，脑瓜小肚儿大。", "化学测试乱如麻，全仗你叫声哇。", "生听了愁眉，徒听了甚怕。", "哪里去辨什么真共假？", "眼见得收翻了这家，", "没抵住那家，", "只吸得血尽人逃罢！"]')
io.open(p,'w',encoding='utf-8').write(s)
