import argparse
import collections
import io
import re
import struct
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from fontTools.ttLib import TTFont

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
STYLES = ("regular", "bold", "italic", "boldItalic")


def audit(path: Path) -> bool:
    z = zipfile.ZipFile(path)
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    targets = {
        rel.get("Id"): "ppt/" + rel.get("Target")
        for rel in rels.findall(f"{{{REL}}}Relationship")
        if rel.get("Type", "").endswith("/font")
    }

    print(f"file: {path}")
    print(f"size: {path.stat().st_size:,} bytes")

    embedded = {}
    ok = True
    print("embedded fonts:")
    for node in pres.findall(f".//{{{P}}}embeddedFont"):
        face = node.find(f"{{{P}}}font")
        family = face.get("typeface", "")
        styles = []
        for style in STYLES:
            link = node.find(f"{{{P}}}{style}")
            if link is None:
                continue
            rid = link.get(f"{{{R}}}id")
            part = targets.get(rid)
            if not part:
                print(f"  {family} {style}: missing relationship {rid}")
                ok = False
                continue
            blob = z.read(part)
            eot_size, data_size, version, flags = struct.unpack("<IIII", blob[:16])
            header = eot_size - data_size
            font = TTFont(io.BytesIO(blob[header:]), lazy=True)
            glyphs = len(font.getGlyphOrder())
            styles.append(style)
            print(
                f"  {family} / {style}: {part}, flags={flags}, "
                f"glyphs={glyphs}, bytes={data_size:,}"
            )
            if flags & 0x00000005:
                print("    ERROR: subset or compressed EOT flag remains")
                ok = False
        embedded[family] = styles

    used = collections.Counter()
    for name in z.namelist():
        if re.match(r"ppt/slides/slide\d+\.xml$", name):
            slide = ET.fromstring(z.read(name))
            for tag in ("latin", "ea", "cs"):
                for el in slide.iter(f"{{{A}}}{tag}"):
                    face = el.get("typeface", "")
                    if face and not face.startswith("+"):
                        used[face] += 1

    theme = ET.fromstring(z.read("ppt/theme/theme1.xml"))
    for scheme in ("majorFont", "minorFont"):
        node = theme.find(f".//{{{A}}}{scheme}")
        if node is not None:
            for tag in ("latin", "ea", "cs"):
                el = node.find(f"{{{A}}}{tag}")
                if el is not None and el.get("typeface"):
                    used[el.get("typeface")] += 1

    print("used families:")
    for face, count in sorted(used.items()):
        state = "embedded" if face in embedded else "MISSING"
        print(f"  {face}: {count} ({state})")
        if face not in embedded:
            ok = False

    clicks = with_effects = 0
    for name in z.namelist():
        if re.match(r"ppt/slides/slide\d+\.xml$", name):
            xml = z.read(name).decode("utf-8")
            clicks += xml.count('nodeType="clickEffect"')
            with_effects += xml.count('nodeType="withEffect"')
    print(f"animations: clickEffect={clicks}, withEffect={with_effects}")

    print("RESULT:", "PASS" if ok else "FAIL")
    return ok


if __name__ == "__main__":
    default = Path(__file__).with_name("多余的话_瑞士国际主义版式.pptx")
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx", nargs="?", type=Path, default=default)
    args = parser.parse_args()
    raise SystemExit(0 if audit(args.pptx) else 1)
