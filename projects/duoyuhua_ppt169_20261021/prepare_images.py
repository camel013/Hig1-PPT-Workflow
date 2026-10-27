"""Prepare source photographs as restrained print-toned assets."""
from pathlib import Path

from PIL import Image, ImageEnhance, ImageFilter, ImageOps


ROOT = Path(__file__).resolve().parent
IMAGES = ROOT / "images"

DARK_RGB = (33, 36, 38)
ARC_RGB = (135, 58, 51)


def duotone(src, dst, dark, light, contrast=1.0):
    with Image.open(src) as image:
        gray = ImageOps.autocontrast(image.convert("L"), cutoff=0.5)
        if contrast != 1.0:
            gray = ImageEnhance.Contrast(gray).enhance(contrast)
        out = ImageOps.colorize(gray, black=dark, white=light)
        out.save(dst, "JPEG", quality=92, optimize=True, progressive=True)


def _smoothstep(t):
    t = max(0.0, min(1.0, t))
    return t * t * (3.0 - 2.0 * t)


def _backdrop(x, y):
    """Page colour underneath the picture at a logical slide coordinate."""
    if (x - 1060) ** 2 + (y - 720) ** 2 <= 300 ** 2:
        return ARC_RGB
    return DARK_RGB


def blur_circle_zone(src, dst, box=(289, 361), scale=3, origin=(800, 190),
                     center=(1160, 20), radius=235, blur_px=9, feather_px=22,
                     focus=(0.5, 0.40)):
    """Gaussian soften where the photograph sits over a background circle.

    The picture stays fully opaque and keeps hard outer edges; only the zone
    that coincides with the arc is defocused, with a smooth ramp so no crisp
    blur seam appears inside the photograph.
    """
    target = (box[0] * scale, box[1] * scale)
    with Image.open(src) as image:
        rgb = ImageOps.fit(image.convert("RGB"), target, method=Image.Resampling.LANCZOS,
                           centering=focus)
    blurred = rgb.filter(ImageFilter.GaussianBlur(blur_px * scale))
    out = rgb.copy()
    src_px, blur_px_data, out_px = rgb.load(), blurred.load(), out.load()
    ox, oy = origin
    cx, cy = center
    width, height = target
    ramp = feather_px * scale
    for y in range(height):
        sy = oy + (y + 0.5) / scale
        for x in range(width):
            sx = ox + (x + 0.5) / scale
            d = ((sx - cx) ** 2 + (sy - cy) ** 2) ** 0.5
            m = 1.0 - _smoothstep((d - radius + ramp) / (2.0 * ramp))
            if m <= 0.002:
                continue
            if m >= 0.998:
                out_px[x, y] = blur_px_data[x, y]
                continue
            a, b = src_px[x, y], blur_px_data[x, y]
            out_px[x, y] = tuple(round(a[i] + (b[i] - a[i]) * m) for i in range(3))
    out.save(dst, "PNG", optimize=True)


def gradient_feather_png(src, dst, box=(340, 470), scale=2,
                         feather=(26, 34, 30, 58), curve=20,
                         origin=(780, 130), tint=1.0, color_power=1.15):
    """Continuous alpha feather that dissolves into the local page colour.

    The four feather widths are independent. The bottom boundary is a shallow
    quadratic arc so it can ride along a large background circle. Edge pixels
    are also mixed toward the colour underneath them, so a light photograph
    does not leave a grey halo on a dark ground or a dark band on a red arc.
    """
    target = (box[0] * scale, box[1] * scale)
    with Image.open(src) as image:
        rgb = ImageOps.fit(image.convert("RGB"), target, method=Image.Resampling.LANCZOS,
                           centering=(0.5, 0.5))
    width, height = target
    left, right, top, bottom = (v * scale for v in feather)
    arc = curve * scale
    out = Image.new("RGBA", target)
    src_px = rgb.load()
    out_px = out.load()
    for y in range(height):
        for x in range(width):
            u = 2.0 * x / (width - 1) - 1.0
            bottom_shift = arc * ((u + 1.0) / 2.0) ** 2
            alpha = min(
                _smoothstep(x / left),
                _smoothstep((width - 1 - x) / right),
                _smoothstep(y / top),
                _smoothstep((height - 1 - y - bottom_shift) / bottom),
            )
            r, g, b = src_px[x, y]
            bg_r, bg_g, bg_b = _backdrop(origin[0] + x / scale,
                                         origin[1] + y / scale)
            keep = 1.0 - tint * (1.0 - alpha ** color_power)
            out_px[x, y] = (
                round(bg_r + (r - bg_r) * keep),
                round(bg_g + (g - bg_g) * keep),
                round(bg_b + (b - bg_b) * keep),
                round(255 * alpha),
            )
    out.save(dst, "PNG", optimize=True)


def main():
    duotone(IMAGES / "qu_portrait_original.jpg", IMAGES / "qu_portrait_print.jpg",
            "#1D1D1C", "#F1EBDD", 1.03)
    duotone(IMAGES / "qu_1935_original.jpg", IMAGES / "qu_1935_print.jpg",
            "#242220", "#EFE8DA", 1.08)
    duotone(IMAGES / "monument_original.jpg", IMAGES / "monument_teal.jpg",
            "#162322", "#C7D0CB", 1.04)
    gradient_feather_png(IMAGES / "monument_teal.jpg", IMAGES / "monument_teal_gradient.png")
    blur_circle_zone(IMAGES / "qu_portrait_print.jpg", IMAGES / "qu_portrait_blur.png")


if __name__ == "__main__":
    main()
