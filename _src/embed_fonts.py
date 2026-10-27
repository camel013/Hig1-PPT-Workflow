import struct
import sys
import zipfile
from pathlib import Path

from fontTools.ttLib import TTFont


REF = Path(r"C:\Users\camel\Documents\ChatGPT\Hig1\projects\duoyuhua_ppt169_20261021\多余的话_瑞士国际主义版式.pptx")
SRC = Path(sys.argv[1])
DST = Path(sys.argv[2])

FONTS = [
    ("fangsong", "ppt/fonts/font1.fntdata", r"C:\Windows\Fonts\simfang.ttf",
     "仿宋", "FangSong", -122, "49"),
    ("zhongsong", "ppt/fonts/font2.fntdata",
     r"C:\Users\camel\AppData\Local\Microsoft\Windows\Fonts\STZHONGS.TTF",
     "华文中宋", "STZhongsong", -122, "2"),
    ("times", "ppt/fonts/font3.fntdata", r"C:\Windows\Fonts\times.ttf",
     "Times New Roman", "Times New Roman", 0, "18"),
]


def read_template(part):
    with zipfile.ZipFile(REF) as z:
        data = z.read(part)
    eot_size, font_size = struct.unpack_from("<II", data, 0)
    header_size = eot_size - font_size
    header = data[:header_size]
    if not 176 <= header_size <= 512:
        raise ValueError(f"Unexpected EOT header size {header_size}")

    off = 82
    names = []
    for _ in range(4):
        size = struct.unpack_from("<H", header, off)[0]
        off += 2
        names.append(header[off:off + size].decode("utf-16le").rstrip("\x00"))
        off += size + 2
    size = struct.unpack_from("<H", header, off)[0]
    off += 2 + size
    return bytearray(header[:82]), names, header[off:]


def panose_bytes(font):
    p = font["OS/2"].panose
    return bytes([
        p.bFamilyType, p.bSerifStyle, p.bWeight, p.bProportion,
        p.bContrast, p.bStrokeVariation, p.bArmStyle, p.bLetterForm,
        p.bMidline, p.bXHeight,
    ])


def build_eot(template_part, ttf_path, family, full_name, charset):
    fixed, names, tail = read_template(template_part)
    ttf_path = Path(ttf_path)
    ttf_data = ttf_path.read_bytes()
    font = TTFont(str(ttf_path), lazy=True)
    os2 = font["OS/2"]
    version = font["name"].getDebugName(5) or "Version 1.00"

    fixed[0x10:0x1A] = panose_bytes(font)
    fixed[0x1A] = charset & 0xFF
    fixed[0x1B] = 0
    struct.pack_into("<I", fixed, 0x1C, int(os2.usWeightClass))
    struct.pack_into("<H", fixed, 0x20, int(os2.fsType))
    struct.pack_into("<I", fixed, 0x3C, int(font["head"].checkSumAdjustment))

    strings = [
        family + "\x00",
        "Regular\x00",
        version + "\x00",
        full_name + "\x00",
        "",
    ]
    header = bytearray(fixed)
    for i, value in enumerate(strings):
        encoded = value.encode("utf-16le")
        header += struct.pack("<H", len(encoded))
        header += encoded
        if i < 4:
            header += b"\x00\x00"
    header += tail

    eot = bytearray(header + ttf_data)
    struct.pack_into("<I", eot, 0, len(eot))
    struct.pack_into("<I", eot, 4, len(ttf_data))
    struct.pack_into("<I", eot, 8, 0x00020002)
    struct.pack_into("<I", eot, 12, 0)
    return bytes(eot)


def main():
    payloads = {}
    for key, part, ttf, family, full_name, charset, _pitch in FONTS:
        payloads[part] = build_eot(part, ttf, family, full_name, charset)

    with zipfile.ZipFile(SRC) as zin:
        names = zin.namelist()
        pres = zin.read("ppt/presentation.xml").decode("utf-8")
        rels = zin.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
        content = zin.read("[Content_Types].xml").decode("utf-8")

        pres = pres.replace(
            'saveSubsetFonts="1"',
            'saveSubsetFonts="1" embedTrueTypeFonts="1"',
            1,
        )

        entries = []
        for i, (key, _part, _ttf, family, _full, charset, pitch) in enumerate(FONTS):
            rid = f"rId{102 + i}"
            entries.append(
                f'<p:embeddedFont><p:font typeface="{family}" '
                f'panose="{panose_hex(key)}" pitchFamily="{pitch}" '
                f'charset="{charset if charset else 0}"/>'
                f'<p:regular r:id="{rid}"/></p:embeddedFont>'
            )
        block = "<p:embeddedFontLst>" + "".join(entries) + "</p:embeddedFontLst>"
        pres = pres.replace("<p:defaultTextStyle>", block + "<p:defaultTextStyle>", 1)

        rel_block = []
        for i, (key, part, _ttf, _family, _full, _charset, _pitch) in enumerate(FONTS):
            rid = f"rId{102 + i}"
            rel_block.append(
                f'<Relationship Id="{rid}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" '
                f'Target="{part.removeprefix("ppt/")}"/>'
            )
        rels = rels.replace("</Relationships>", "".join(rel_block) + "</Relationships>", 1)

        if 'Extension="fntdata"' not in content:
            content = content.replace(
                "</Types>",
                '<Default Extension="fntdata" ContentType="application/x-fontdata"/></Types>',
                1,
            )

        with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename in payloads:
                    continue
                data = zin.read(info.filename)
                if info.filename == "ppt/presentation.xml":
                    data = pres.encode("utf-8")
                elif info.filename == "ppt/_rels/presentation.xml.rels":
                    data = rels.encode("utf-8")
                elif info.filename == "[Content_Types].xml":
                    data = content.encode("utf-8")
                zout.writestr(info, data)
            for part, data in payloads.items():
                zout.writestr(part, data, compress_type=zipfile.ZIP_DEFLATED)

    print(f"wrote {DST} ({DST.stat().st_size} bytes)")


def panose_hex(key):
    return {
        "fangsong": "02010609060101010101",
        "zhongsong": "02010600040101010101",
        "times": "02020603050405020304",
    }[key]


if __name__ == "__main__":
    main()
