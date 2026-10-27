from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent.parent
PREVIEW_SOURCE = ROOT / "_native_preview"
OUTPUT = ROOT / "dist"
PREVIEW_OUTPUT = OUTPUT / "蓝调背后_影子中的人们_预览"
PPTX = OUTPUT / "蓝调背后_影子中的人们_精选12页.pptx"
ZIP = OUTPUT / "蓝调背后_影子中的人们_精选12页_交付包.zip"

NAMES = [
    "01_封面",
    "02_核心判断",
    "03_种植园起源",
    "04_制度化",
    "05_棉花帝国",
    "06_共同增长",
    "07_政治危机",
    "08_内战目标",
    "09_主体行动",
    "10_重建空间",
    "11_重建失败",
    "12_尾声",
]


def build_contact_sheet(paths: list[Path]) -> Path:
    thumb_width = 480
    thumb_height = 270
    label_height = 34
    margin = 24
    columns = 4
    rows = 3
    width = columns * thumb_width + (columns + 1) * margin
    height = rows * (thumb_height + label_height) + (rows + 1) * margin
    canvas = Image.new("RGB", (width, height), "#DEDBD3")
    draw = ImageDraw.Draw(canvas)
    font = ImageFont.load_default(size=18)

    for index, image_path in enumerate(paths):
        row, column = divmod(index, columns)
        x = margin + column * (thumb_width + margin)
        y = margin + row * (thumb_height + label_height + margin)
        image = Image.open(image_path).convert("RGB")
        image = image.resize((thumb_width, thumb_height), Image.Resampling.LANCZOS)
        canvas.paste(image, (x, y + label_height))
        draw.text((x, y + 6), NAMES[index], fill="#222222", font=font)

    target = PREVIEW_OUTPUT / "00_联系表.png"
    canvas.save(target, optimize=True)
    return target


def main() -> None:
    PREVIEW_OUTPUT.mkdir(parents=True, exist_ok=True)
    source_paths = sorted(PREVIEW_SOURCE.glob("*.png"))
    if len(source_paths) != 12:
        raise RuntimeError(f"Expected 12 previews, found {len(source_paths)}.")

    copied = []
    for name, source in zip(NAMES, source_paths):
        target = PREVIEW_OUTPUT / f"{name}.png"
        shutil.copy2(source, target)
        copied.append(target)

    contact_sheet = build_contact_sheet(copied)

    if ZIP.exists():
        ZIP.unlink()
    with zipfile.ZipFile(ZIP, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.write(PPTX, PPTX.name)
        for item in [contact_sheet, *copied]:
            archive.write(item, f"预览/{item.name}")

    print(PPTX)
    print(contact_sheet)
    print(ZIP)


if __name__ == "__main__":
    main()
