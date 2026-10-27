from html import escape
from pathlib import Path


root = Path(__file__).resolve().parent
slides = sorted((root / "svg_final").glob("*.svg"))

sections = []
for slide in slides:
    svg = slide.read_text(encoding="utf-8")
    sections.append(f'<section class="slide">{svg}</section>')

html = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>蓝调背后：影子中的人们</title>
<style>
  @page { size: 1280px 720px; margin: 0; }
  html, body { margin: 0; padding: 0; background: #f2eee6; }
  .slide {
    width: 1280px;
    height: 720px;
    overflow: hidden;
    page-break-after: always;
    break-after: page;
  }
  .slide:last-child { page-break-after: auto; break-after: auto; }
  svg { display: block; width: 1280px; height: 720px; }
</style>
</head>
<body>
""" + "\n".join(sections) + """
</body>
</html>
"""

(root / "preview.html").write_text(html, encoding="utf-8")
print(root / "preview.html")
