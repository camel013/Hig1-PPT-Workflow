from __future__ import annotations

import csv
import zipfile
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageOps


SRC = Path(r"C:\Users\camel\Desktop\qub")
OUT = Path(r"C:\Users\camel\Documents\ChatGPT\Hig1\qub_环境影像_统一版")
EDITS = OUT / "edits"

SERIES = [
    {"n": 1, "file": "2BDA4F6A1A063FF5C618E20B3A07B020.jpg", "tag": "建筑周边 / 堆物", "cx": 0.36, "cy": 0.42, "scale": 0.72},
    {"n": 2, "file": "1C5F177C9F89A142AACC56ABC3F56515.jpg", "tag": "楼道转角 / 杂物", "cx": 0.44, "cy": 0.48, "scale": 0.82},
    {"n": 3, "file": "CE8086A40E3553A4B544557DA9DA6AB4.jpg", "tag": "棚下空间 / 遗弃物", "cx": 0.48, "cy": 0.45, "scale": 0.86},
    {"n": 4, "file": "09F9C58243E04781F97797B4B107F942.jpg", "tag": "通道 / 木板与苫布", "cx": 0.46, "cy": 0.47, "scale": 0.80},
    {"n": 5, "file": "3DC0DDD16728E9BFC7ADE7C1702412A9.jpg", "tag": "路侧 / 围挡与垃圾", "cx": 0.49, "cy": 0.48, "scale": 0.88},
    {"n": 6, "file": "7AE7585DBE72AA7697D38BC1BA9D97AD.jpg", "tag": "路缘 / 塑料垃圾", "cx": 0.48, "cy": 0.55, "scale": 0.82},
    {"n": 7, "file": "E8EE5EE9EC758D6468611AA3F038C4D6.jpg", "tag": "绿化带 / 塑料袋", "cx": 0.32, "cy": 0.33, "scale": 0.70},
    {"n": 8, "file": "C42A1B6D5890D230DF4AF2E8B8BE1F00.jpg", "tag": "灌木下 / 散落垃圾", "cx": 0.46, "cy": 0.47, "scale": 0.84},
    {"n": 9, "file": "6C22FDE939EADF53A3E866436B09DF59.jpg", "tag": "林下 / 纸塑垃圾", "cx": 0.48, "cy": 0.58, "scale": 0.80},
    {"n": 10, "file": "D56BFC74C60BACBB7BF8E525546DCA31.jpg", "tag": "草坡 / 编织袋", "cx": 0.50, "cy": 0.56, "scale": 0.80},
    {"n": 11, "file": "E6555DE1516A17014F529E7E21B43131.jpg", "tag": "枯草 / 塑料碎片", "cx": 0.50, "cy": 0.54, "scale": 0.80},
    {"n": 12, "file": "B58FC7F51F6DB71A20AB5C89D885DF0E.jpg", "tag": "地面 / 包装垃圾", "cx": 0.48, "cy": 0.62, "scale": 0.72},
    {"n": 13, "file": "9B121A99B760FBAE6F261B604822EB64.jpg", "tag": "地面细部 / 瓶与烟盒", "cx": 0.52, "cy": 0.55, "scale": 0.78},
    {"n": 14, "file": "D86792A869D33CF3EFC596A20417AB91.jpg", "tag": "草坪 / 纸屑", "cx": 0.56, "cy": 0.52, "scale": 0.65},
]


def srgb_to_linear(x: np.ndarray) -> np.ndarray:
    return np.where(x <= 0.04045, x / 12.92, ((x + 0.055) / 1.055) ** 2.4)


def linear_to_srgb(x: np.ndarray) -> np.ndarray:
    x = np.clip(x, 0.0, 1.0)
    return np.where(x <= 0.0031308, x * 12.92, 1.055 * x ** (1 / 2.4) - 0.055)


def luminance(rgb: np.ndarray) -> np.ndarray:
    return 0.2126 * rgb[..., 0] + 0.7152 * rgb[..., 1] + 0.0722 * rgb[..., 2]


def crop_for_spec(im: Image.Image, spec: dict) -> Image.Image:
    width, height = im.size
    crop_w = width * spec["scale"]
    crop_h = height * spec["scale"]
    left = spec["cx"] * width - crop_w / 2
    top = spec["cy"] * height - crop_h / 2
    left = min(max(left, 0), width - crop_w)
    top = min(max(top, 0), height - crop_h)
    box = (round(left), round(top), round(left + crop_w), round(top + crop_h))
    return im.crop(box)


def estimate_white_balance(srgb: np.ndarray) -> np.ndarray:
    sample_im = Image.fromarray(np.uint8(np.clip(srgb, 0, 1) * 255)).resize(
        (640, 480), Image.Resampling.BILINEAR
    )
    sample = np.asarray(sample_im, dtype=np.float32) / 255.0
    lin = srgb_to_linear(sample)
    y = luminance(sample)
    mx = sample.max(2)
    mn = sample.min(2)
    sat = (mx - mn) / np.maximum(mx, 1e-6)
    mask = (sat <= np.quantile(sat, 0.20)) & (y > 0.10) & (y < 0.90)
    if int(mask.sum()) < 500:
        mask = (y > 0.10) & (y < 0.90)
    mean = lin[mask].mean(axis=0)
    mean = np.maximum(mean, 1e-5)
    gains = float(mean.mean()) / mean
    gains = np.clip(gains, 0.82, 1.22)
    gains = gains / np.exp(np.mean(np.log(gains)))
    return 1.0 + (gains - 1.0) * 0.45


def exposure_gain(srgb: np.ndarray) -> float:
    y = luminance(srgb)
    median = float(np.median(y))
    target = 0.49
    gain = (target / max(median, 0.01)) ** 0.70
    return float(np.clip(gain, 0.88, 1.16))


def apply_tone(srgb: np.ndarray) -> np.ndarray:
    x = np.clip(srgb, 0.0, 1.0)
    d = x - 0.5
    x = x + 0.22 * d * (1.0 - 4.0 * d * d)
    x = 0.006 + 0.984 * x
    return np.clip(x, 0.0, 1.0)


def adjust_color(srgb: np.ndarray) -> np.ndarray:
    rgb8 = np.uint8(np.clip(srgb, 0.0, 1.0) * 255)
    hsv = np.asarray(Image.fromarray(rgb8, "RGB").convert("HSV"), dtype=np.float32) / 255.0
    h, s, v = hsv[..., 0].copy(), hsv[..., 1].copy(), hsv[..., 2].copy()

    green_w = np.clip(1.0 - np.abs(h - 0.333) / 0.115, 0.0, 1.0)
    yellow_w = np.clip(1.0 - np.abs(h - 0.155) / 0.075, 0.0, 1.0)
    blue_w = np.clip(1.0 - np.abs(h - 0.610) / 0.100, 0.0, 1.0)
    red_w = np.maximum(
        np.clip(1.0 - np.abs(h - 0.995) / 0.040, 0.0, 1.0),
        np.clip(1.0 - h / 0.040, 0.0, 1.0),
    )

    s *= 0.90
    s *= 1.0 - 0.17 * green_w
    s *= 1.0 - 0.06 * yellow_w
    s *= 1.0 - 0.09 * blue_w
    s *= 1.0 + 0.025 * red_w
    h -= 0.010 * green_w
    v *= 1.0 + 0.003 * green_w

    hsv[..., 0] = np.mod(h, 1.0)
    hsv[..., 1] = np.clip(s, 0.0, 1.0)
    hsv[..., 2] = np.clip(v, 0.0, 1.0)
    out = np.asarray(
        Image.fromarray(np.uint8(np.clip(hsv, 0.0, 1.0) * 255), "HSV").convert("RGB"),
        dtype=np.float32,
    ) / 255.0

    y = luminance(out)
    shadow = (1.0 - y) ** 2
    highlight = y ** 2
    out[..., 0] = out[..., 0] + 0.0030 * shadow + 0.0035 * highlight
    out[..., 1] = out[..., 1] + 0.0040 * shadow + 0.0010 * highlight
    out[..., 2] = out[..., 2] + 0.0070 * shadow
    return np.clip(out, 0.0, 1.0)


def grade(im: Image.Image, spec: dict) -> tuple[Image.Image, dict]:
    cropped = crop_for_spec(im, spec)
    srgb = np.asarray(cropped, dtype=np.float32) / 255.0

    gains = estimate_white_balance(srgb)
    exp_gain = exposure_gain(srgb)
    lin = srgb_to_linear(srgb)
    lin *= gains * (exp_gain ** 2.2)
    balanced = linear_to_srgb(lin)
    toned = apply_tone(balanced)
    colored = adjust_color(toned)

    out8 = np.uint8(np.clip(colored, 0.0, 1.0) * 255 + 0.5)
    out = Image.fromarray(out8, "RGB")
    out = out.filter(ImageFilter.UnsharpMask(radius=1.0, percent=52, threshold=3))

    stats = {
        "median_before": round(float(np.median(luminance(srgb))), 3),
        "median_after": round(float(np.median(luminance(colored))), 3),
        "wb_gain_r": round(float(gains[0]), 3),
        "wb_gain_g": round(float(gains[1]), 3),
        "wb_gain_b": round(float(gains[2]), 3),
        "exposure_gain": round(exp_gain, 3),
    }
    return out, stats


def font(path: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(path, size)


def fit_contain(im: Image.Image, size: tuple[int, int], bg: tuple[int, int, int]) -> Image.Image:
    fitted = ImageOps.contain(im, size, Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", size, bg)
    canvas.paste(fitted, ((size[0] - fitted.width) // 2, (size[1] - fitted.height) // 2))
    return canvas


def make_contact_sheet(items: list[dict]) -> None:
    cols, rows = 4, 4
    margin, gap = 56, 20
    cell_w, image_h, caption_h = 500, 375, 48
    row_h = image_h + caption_h
    width = margin * 2 + cols * cell_w + (cols - 1) * gap
    height = margin * 2 + 112 + rows * row_h + (rows - 1) * gap
    sheet = Image.new("RGB", (width, height), (242, 241, 236))
    draw = ImageDraw.Draw(sheet)

    title_font = font(r"C:\Windows\Fonts\simkai.ttf", 38)
    sub_font = font(r"C:\Windows\Fonts\simsun.ttc", 18)
    num_font = font(r"C:\Windows\Fonts\simsun.ttc", 20)
    tag_font = font(r"C:\Windows\Fonts\simsun.ttc", 17)
    draw.text((margin, margin - 4), "环境问题影像 · 统一画幅与色调", fill=(24, 24, 22), font=title_font)
    draw.text((margin, margin + 58), "14 幅 / 横向 4:3 · 竖向 3:4 / 仅裁切与色彩校正", fill=(72, 72, 68), font=sub_font)
    draw.line((margin, margin + 92, width - margin, margin + 92), fill=(60, 60, 56), width=2)

    top0 = margin + 112
    for idx, item in enumerate(items):
        row, col = divmod(idx, cols)
        x = margin + col * (cell_w + gap)
        y = top0 + row * (row_h + gap)
        image = Image.open(item["output"]).convert("RGB")
        tile = fit_contain(image, (cell_w, image_h), (224, 222, 215))
        sheet.paste(tile, (x, y))
        draw.text((x, y + image_h + 9), f'{item["n"]:02d}', fill=(25, 25, 23), font=num_font)
        draw.text((x + 42, y + image_h + 10), item["tag"], fill=(60, 60, 56), font=tag_font)

    sheet.save(OUT / "00_系列总览.jpg", quality=94, subsampling=0, optimize=True, progressive=True)


def make_before_after(items: list[dict], picks: list[int]) -> None:
    margin, gap = 48, 24
    pair_w, pair_h = 760, 570
    width = margin * 2 + pair_w * 2 + gap
    height = margin * 2 + 92 + len(picks) * (pair_h + 52) + (len(picks) - 1) * 34
    sheet = Image.new("RGB", (width, height), (242, 241, 236))
    draw = ImageDraw.Draw(sheet)
    title_font = font(r"C:\Windows\Fonts\simkai.ttf", 36)
    label_font = font(r"C:\Windows\Fonts\simsun.ttc", 19)
    draw.text((margin, margin - 4), "前后对照 · 裁切与统一色调", fill=(24, 24, 22), font=title_font)
    draw.line((margin, margin + 52, width - margin, margin + 52), fill=(60, 60, 56), width=2)

    by_n = {item["n"]: item for item in items}
    y = margin + 82
    for n in picks:
        item = by_n[n]
        original = ImageOps.exif_transpose(Image.open(SRC / item["file"])).convert("RGB")
        edited = Image.open(item["output"]).convert("RGB")
        draw.text((margin, y), "BEFORE / 原始画幅", fill=(46, 46, 43), font=label_font)
        draw.text((margin + pair_w + gap, y), "AFTER / 统一版", fill=(46, 46, 43), font=label_font)
        sheet.paste(fit_contain(original, (pair_w, pair_h), (224, 222, 215)), (margin, y + 30))
        sheet.paste(fit_contain(edited, (pair_w, pair_h), (224, 222, 215)), (margin + pair_w + gap, y + 30))
        y += pair_h + 30 + 52

    sheet.save(OUT / "00_前后对照.jpg", quality=94, subsampling=0, optimize=True, progressive=True)


def main() -> None:
    EDITS.mkdir(parents=True, exist_ok=True)
    items = []
    for spec in SERIES:
        source = SRC / spec["file"]
        original = ImageOps.exif_transpose(Image.open(source)).convert("RGB")
        edited, stats = grade(original, spec)
        out_name = f'{spec["n"]:02d}_{source.stem}.jpg'
        out_path = EDITS / out_name
        exif = Image.open(source).getexif()
        exif[274] = 1
        edited.save(
            out_path,
            quality=95,
            subsampling=0,
            optimize=True,
            progressive=True,
            exif=exif.tobytes(),
        )
        items.append({**spec, "output": out_path, "size": edited.size, "stats": stats})
        print(f'{spec["n"]:02d} {source.name} -> {edited.width}x{edited.height} {stats}')

    with (OUT / "manifest.csv").open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow([
            "序号", "原始文件", "输出文件", "宽", "高", "构图标签",
            "裁切中心X", "裁切中心Y", "裁切比例", "原图亮度中位数", "成片亮度中位数",
            "白平衡R", "白平衡G", "白平衡B", "曝光校正",
        ])
        for item in items:
            s = item["stats"]
            writer.writerow([
                f'{item["n"]:02d}', item["file"], item["output"].name,
                item["size"][0], item["size"][1], item["tag"],
                item["cx"], item["cy"], item["scale"],
                s["median_before"], s["median_after"],
                s["wb_gain_r"], s["wb_gain_g"], s["wb_gain_b"], s["exposure_gain"],
            ])

    readme = """# 环境问题影像 · 统一版

## 输出

- `edits/`：14 幅成片，JPEG，95 品质，4:4:4。
- 横向照片统一为 4:3，竖向照片统一为 3:4。
- 文件名前缀为叙事顺序，后半段保留原始文件名，便于回溯。

## 处理原则

- 只做裁切、白平衡、曝光、曲线、局部色彩与轻度锐化；未移除、添加或重绘任何现场物体。
- 构图围绕垃圾、堆物、破损设施与环境关系收紧，减少空白地面和无关边缘。
- 色调采用克制的纪实风格：混凝土保持中性，草绿降饱和，保留红、黄、蓝等警示与识别色。

## 校色参数

- 灰区白平衡校正强度 45%，限制在 0.82–1.22。
- 亮度中位数目标 0.49，曝光校正限制在 0.88–1.16。
- 轻微 S 曲线，暗部提升 0.6%，高光压缩至 99%。
- 全局饱和 90%；绿色额外降低 17%，黄色降低 6%，蓝色降低 9%。
- 轻度 USM 锐化：半径 1.0，强度 52%，阈值 3。

## 对应表

见 `manifest.csv`。
"""
    (OUT / "README.md").write_text(readme, encoding="utf-8")
    make_contact_sheet(items)
    make_before_after(items, [1, 7])

    zip_path = OUT / "环境影像_统一版.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for path in sorted(EDITS.glob("*.jpg")):
            zf.write(path, arcname=f"edits/{path.name}")
        zf.write(OUT / "manifest.csv", arcname="manifest.csv")
        zf.write(OUT / "README.md", arcname="README.md")
        zf.write(OUT / "00_系列总览.jpg", arcname="00_系列总览.jpg")
        zf.write(OUT / "00_前后对照.jpg", arcname="00_前后对照.jpg")


if __name__ == "__main__":
    main()
