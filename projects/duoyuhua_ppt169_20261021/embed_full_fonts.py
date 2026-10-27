import argparse
import os
import struct
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from fontTools.ttLib import TTFont


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"

USER_FONTS = Path(r"C:\Users\camel\AppData\Local\Microsoft\Windows\Fonts")
WIN_FONTS = Path(r"C:\Windows\Fonts")

FONT_FILES = {
    ("楷体", "regular"): WIN_FONTS / "simkai.ttf",
    ("华文中宋", "regular"): USER_FONTS / "STZHONGS.TTF",
    ("Times Newer Roman", "regular"): USER_FONTS / "TimesNewerRoman-Regular.ttf",
    ("Trajan Pro Deck", "regular"): USER_FONTS / "TrajanProDeck.ttf",
    ("等线", "regular"): WIN_FONTS / "Deng.ttf",
    ("等线", "bold"): WIN_FONTS / "Dengb.ttf",
    ("Calibri", "regular"): WIN_FONTS / "calibri.ttf",
    ("Calibri", "bold"): WIN_FONTS / "calibrib.ttf",
    ("Calibri", "italic"): WIN_FONTS / "calibrii.ttf",
    ("Calibri", "boldItalic"): WIN_FONTS / "calibriz.ttf",
}

STYLES = ("regular", "bold", "italic", "boldItalic")


def font_part_map(pres_xml: bytes, rels_xml: bytes) -> dict[str, tuple[str, str]]:
    rel_root = ET.fromstring(rels_xml)
    targets = {
        rel.get("Id"): "ppt/" + rel.get("Target")
        for rel in rel_root.findall(f"{{{REL}}}Relationship")
        if rel.get("Type", "").endswith("/font")
    }
    result = {}
    pres_root = ET.fromstring(pres_xml)
    for embedded in pres_root.findall(f".//{{{P}}}embeddedFont"):
        face = embedded.find(f"{{{P}}}font")
        if face is None:
            continue
        family = face.get("typeface", "")
        for style in STYLES:
            node = embedded.find(f"{{{P}}}{style}")
            if node is None:
                continue
            rid = node.get(f"{{{R}}}id")
            if rid in targets:
                result[targets[rid]] = (family, style)
    return result


def make_full_eot(template: bytes, ttf_path: Path) -> bytes:
    eot_size, font_data_size, version, flags = struct.unpack("<IIII", template[:16])
    header_size = eot_size - font_data_size
    if not 16 <= header_size < len(template):
        raise ValueError(f"Unexpected EOT header size {header_size}")
    header = bytearray(template[:header_size])
    ttf = ttf_path.read_bytes()
    font = TTFont(str(ttf_path), lazy=True)

    struct.pack_into("<I", header, 0, header_size + len(ttf))
    struct.pack_into("<I", header, 4, len(ttf))
    struct.pack_into("<I", header, 12, flags & ~0x00000005)
    struct.pack_into("<H", header, 32, int(font["OS/2"].fsType))
    struct.pack_into("<I", header, 60, int(font["head"].checkSumAdjustment))
    return bytes(header) + ttf


def embed(src: Path, dst: Path) -> None:
    tmp = dst.with_name(dst.stem + ".tmp" + dst.suffix)
    with zipfile.ZipFile(src) as zin:
        pres_xml = zin.read("ppt/presentation.xml")
        rels_xml = zin.read("ppt/_rels/presentation.xml.rels")
        parts = font_part_map(pres_xml, rels_xml)
        mapped = {}
        for part, key in parts.items():
            if key not in FONT_FILES:
                raise KeyError(f"No full font mapping for {key} ({part})")
            ttf_path = FONT_FILES[key]
            if not ttf_path.exists():
                raise FileNotFoundError(ttf_path)
            mapped[part] = ttf_path

        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename in mapped:
                    before = len(data)
                    data = make_full_eot(data, mapped[info.filename])
                    zout.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
                    family, style = parts[info.filename]
                    print(f"{family} {style}: {before} -> {len(data)} bytes")
                else:
                    zout.writestr(info, data)
    os.replace(tmp, dst)
    print(f"{dst} ({dst.stat().st_size} bytes)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Replace subset EOT fonts with complete TrueType fonts.")
    parser.add_argument("src", type=Path)
    parser.add_argument("dst", type=Path)
    args = parser.parse_args()
    embed(args.src, args.dst)
