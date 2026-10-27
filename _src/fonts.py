from fontTools.ttLib import TTFont
for p in [r'C:\Windows\Fonts\simkai.ttf',r'C:\Users\camel\AppData\Local\Microsoft\Windows\Fonts\STZHONGS.TTF',r'C:\Windows\Fonts\times.ttf']:
    f=TTFont(p)
    print(p, f.sfntVersion, f['name'].getDebugName(1), len(f.getBestCmap()))
