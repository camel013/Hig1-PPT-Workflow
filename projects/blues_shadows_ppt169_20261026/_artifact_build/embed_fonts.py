from __future__ import annotations

import io
import struct
import zipfile
from pathlib import Path

from fontTools.ttLib import TTFont
from lxml import etree


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
FONT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"


def parse_xml(data: bytes):
    return etree.fromstring(data)


def serialize_xml(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def add_embedded_fonts(
    source_path: Path,
    donor_path: Path,
    output_path: Path,
) -> None:
    with zipfile.ZipFile(source_path) as source, zipfile.ZipFile(donor_path) as donor:
        presentation = parse_xml(source.read("ppt/presentation.xml"))
        presentation_rels = parse_xml(source.read("ppt/_rels/presentation.xml.rels"))
        content_types = parse_xml(source.read("[Content_Types].xml"))

        if presentation.find(f"{{{P_NS}}}embeddedFontLst") is not None:
            raise RuntimeError("Target presentation already contains embedded fonts.")

        rel_ids = {rel.get("Id") for rel in presentation_rels}
        new_ids = []
        candidate = 100
        while len(new_ids) < 2:
            value = f"rId{candidate}"
            if value not in rel_ids:
                new_ids.append(value)
                rel_ids.add(value)
            candidate += 1

        fonts = [
            ("华文中宋", "ppt/fonts/font1.fntdata", new_ids[0]),
            ("楷体", "ppt/fonts/font2.fntdata", new_ids[1]),
        ]

        embedded_list = etree.Element(f"{{{P_NS}}}embeddedFontLst")
        for family, part_name, rel_id in fonts:
            embedded_font = etree.SubElement(embedded_list, f"{{{P_NS}}}embeddedFont")
            font = etree.SubElement(embedded_font, f"{{{P_NS}}}font")
            font.set("typeface", family)
            regular = etree.SubElement(embedded_font, f"{{{P_NS}}}regular")
            regular.set(f"{{{R_NS}}}id", rel_id)

            rel = etree.SubElement(presentation_rels, f"{{{PKG_REL_NS}}}Relationship")
            rel.set("Id", rel_id)
            rel.set("Type", FONT_REL_TYPE)
            rel.set("Target", part_name.removeprefix("ppt/"))

        default_text_style = presentation.find(f"{{{P_NS}}}defaultTextStyle")
        if default_text_style is None:
            presentation.append(embedded_list)
        else:
            default_text_style.addprevious(embedded_list)

        presentation.set("embedTrueTypeFonts", "1")
        presentation.set("saveSubsetFonts", "0")

        has_font_default = any(
            element.get("Extension", "").lower() == "fntdata"
            for element in content_types
        )
        if not has_font_default:
            default = etree.Element(f"{{{CT_NS}}}Default")
            default.set("Extension", "fntdata")
            default.set("ContentType", "application/x-fontdata")
            content_types.insert(0, default)

        replacements = {
            "ppt/presentation.xml": serialize_xml(presentation),
            "ppt/_rels/presentation.xml.rels": serialize_xml(presentation_rels),
            "[Content_Types].xml": serialize_xml(content_types),
        }

        for family, part_name, _ in fonts:
            donor_part = part_name
            if donor_part not in donor.namelist():
                raise RuntimeError(f"Donor font part missing: {donor_part}")
            replacements[part_name] = donor.read(donor_part)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as output:
            for info in source.infolist():
                data = replacements.get(info.filename, source.read(info.filename))
                if info.filename.endswith(".xml") or info.filename.endswith(".rels"):
                    text = data.decode("utf-8")
                    text = text.replace('typeface="STZhongsong"', 'typeface="华文中宋"')
                    text = text.replace('typeface="KaiTi"', 'typeface="楷体"')
                    data = text.encode("utf-8")
                output.writestr(info, data)
            for family, part_name, _ in fonts:
                if part_name not in source.namelist():
                    output.writestr(part_name, replacements[part_name])


def verify(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        presentation = parse_xml(archive.read("ppt/presentation.xml"))
        families = [
            font.get("typeface")
            for font in presentation.findall(
                f".//{{{P_NS}}}embeddedFontLst/{{{P_NS}}}embeddedFont/{{{P_NS}}}font"
            )
        ]
        print("families:", families)
        print("embedTrueTypeFonts:", presentation.get("embedTrueTypeFonts"))
        print("saveSubsetFonts:", presentation.get("saveSubsetFonts"))

        for part in sorted(name for name in archive.namelist() if name.startswith("ppt/fonts/")):
            blob = archive.read(part)
            eot_size, font_size, _version, flags = struct.unpack("<IIII", blob[:16])
            header_size = eot_size - font_size
            payload = blob[header_size:]
            font = TTFont(io.BytesIO(payload))
            print(
                part,
                "flags=",
                flags,
                "glyphs=",
                len(font.getGlyphOrder()),
                "family=",
                font["name"].getDebugName(1),
            )
            if flags & 0x5:
                raise RuntimeError(f"{part} is compressed or subset.")
            if not font.getGlyphOrder():
                raise RuntimeError(f"{part} has no glyphs.")


if __name__ == "__main__":
    base = Path(__file__).resolve().parent
    source = base / "final-noembed.pptx"
    donor = Path(
        r"C:\Users\camel\Documents\ChatGPT\Hig1\_quality_baseline_duoyuhua_20261026\deliverables\feedback_6page\缩短客户反馈闭环_6页_重建版.pptx"
    )
    output = base / "final-embedded.pptx"
    add_embedded_fonts(source, donor, output)
    verify(output)
    print(output)
