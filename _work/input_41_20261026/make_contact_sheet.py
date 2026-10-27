from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


folder = Path(__file__).resolve().parent
pages = sorted(folder.glob("page-*.png"), key=lambda p: int(p.stem.split("-")[-1]))

thumb_width = 420
label_height = 34
margin = 24
images = []

for page in pages:
    image = Image.open(page).convert("RGB")
    height = round(image.height * thumb_width / image.width)
    image = image.resize((thumb_width, height), Image.Resampling.LANCZOS)
    images.append((page.name, image))

thumb_height = max(image.height for _, image in images)
sheet_width = margin * 4 + thumb_width * 3
sheet_height = margin * 4 + (thumb_height + label_height) * 3
sheet = Image.new("RGB", (sheet_width, sheet_height), "#dedbd3")
draw = ImageDraw.Draw(sheet)
font = ImageFont.load_default(size=18)

for index, (name, image) in enumerate(images):
    row, column = divmod(index, 3)
    x = margin + column * (thumb_width + margin)
    y = margin + row * (thumb_height + label_height + margin)
    sheet.paste(image, (x, y + label_height))
    draw.text((x, y + 6), name, fill="#222222", font=font)

output = folder / "contact-sheet.png"
sheet.save(output, optimize=True)
print(output)
