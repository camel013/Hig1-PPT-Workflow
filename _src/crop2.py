from PIL import Image
im=Image.open(r'_test_png\slide-3.png')
im.crop((930,140,1200,370)).resize((810,690)).save(r'_test_png\crop3b.png')
