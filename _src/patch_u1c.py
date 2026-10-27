import io
p=r'projects\zhiwen_teaching_ppt169_20261024\deck_unit1.py'
s=io.open(p,encoding='utf-8').read()
s=s.replace('GOLD, "对照《蛭歌行》：长篇靠铺陈，短章靠命中。")','GOLD, "对照《蛭歌行》：长篇靠铺陈，短章靠命中。", bottom="短章不解释，只判决。")')
s=s.replace('GOLD, "最不舒服的阅读，往往最接近真相。")','GOLD, "最不舒服的阅读，往往最接近真相。", bottom="视角不是技巧，是伦理位置。")')
io.open(p,'w',encoding='utf-8').write(s)
