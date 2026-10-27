from PIL import Image
im=Image.open(r'_test_png\slide-3.png')
im.crop((900,130,1280,330)).resize((760,400)).save(r'_test_png\crop3.png')
