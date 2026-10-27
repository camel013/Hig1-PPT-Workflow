# -*- coding: utf-8 -*-
"""Embed full 楷体 / 华文中宋 fonts into the campus deck without using COM."""

from __future__ import annotations

import copy
import io
import os
import struct
import zipfile
from pathlib import Path

from fontTools.ttLib import TTFont
from lxml import etree


ROOT = Path(__file__).resolve().parent
SRC = ROOT / "_deck_build.pptx"
DST = ROOT / "_deck_embedded.pptx"
DONOR = Path(r"C:\Users\camel\Documents\ChatGPT\Hig1\projects\duoyuhua_ppt169_20261021\多余的话_瑞士国际主义版式.pptx")

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
FONT_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
FAMILIES = {"楷体", "华文中宋", "Trajan Pro Deck"}
STYLES = ("regular", "bold", "italic", "boldItalic")


def donor_fonts():
    with zipfile.ZipFile(DONOR) as z:
        pres = etree.fromstring(z.read("ppt/presentation.xml"))
        rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
        targets = {
            rel.get("Id"): rel.get("Target")
            for rel in rels.findall(f"{{{REL_NS}}}Relationship")
        }
        found = []
        for embedded in pres.findall(f".//{{{P_NS}}}embeddedFont"):
            face = embedded.find(f"{{{P_NS}}}font")
            if face is None or face.get("typeface") not in FAMILIES:
                continue
            family = face.get("typeface")
            pieces = []
            for style in STYLES:
                node = embedded.find(f"{{{P_NS}}}{style}")
                if node is None:
                    continue
                rid = node.get(f"{{{R_NS}}}id")
                target = targets[rid]
                data = z.read("ppt/" + target)
                pieces.append((style, data))
            if not pieces:
                raise RuntimeError(f"No payload for {family}")
            found.append((copy.deepcopy(face), pieces))
        return found


def verify_font_blob(blob: bytes, label: str) -> None:
    eot_size, font_size, _, flags = struct.unpack("<IIII", blob[:16])
    header_size = eot_size - font_size
    ttf = TTFont(io.BytesIO(blob[header_size:]), lazy=True)
    glyphs = len(ttf.getGlyphOrder())
    if flags & 0x5:
        raise RuntimeError(f"{label}: subset/compressed flags remain ({flags})")
    if glyphs < 400:
        raise RuntimeError(f"{label}: unexpected glyph count {glyphs}")


def main() -> None:
    fonts = donor_fonts()
    if {face.get("typeface") for face, _ in fonts} != FAMILIES:
        raise RuntimeError("Donor deck does not contain both required families")

    with zipfile.ZipFile(SRC) as zin:
        pres = etree.fromstring(zin.read("ppt/presentation.xml"))
        rels = etree.fromstring(zin.read("ppt/_rels/presentation.xml.rels"))
        cts = etree.fromstring(zin.read("[Content_Types].xml"))
        theme = etree.fromstring(zin.read("ppt/theme/theme1.xml"))
        for scheme in ("majorFont", "minorFont"):
            font_group = theme.find(f".//{{{A_NS}}}{scheme}")
            if font_group is None:
                continue
            for tag in ("latin", "ea", "cs"):
                node = font_group.find(f"{{{A_NS}}}{tag}")
                if node is not None:
                    node.set("typeface", "华文中宋")
        existing_ids = [rel.get("Id") for rel in rels.findall(f"{{{REL_NS}}}Relationship")]
        next_id = max(int(x[3:]) for x in existing_ids if x.startswith("rId")) + 1

        embedded_list = etree.Element(f"{{{P_NS}}}embeddedFontLst")
        new_parts: dict[str, bytes] = {}
        for family, pieces in fonts:
            embedded = etree.SubElement(embedded_list, f"{{{P_NS}}}embeddedFont")
            embedded.append(family)
            for style, data in pieces:
                rid = f"rId{next_id}"
                next_id += 1
                name = f"fonts/font{len(new_parts) + 1}.fntdata"
                new_parts["ppt/" + name] = data
                rel = etree.SubElement(rels, f"{{{REL_NS}}}Relationship")
                rel.set("Id", rid)
                rel.set("Type", FONT_TYPE)
                rel.set("Target", name)
                node = etree.SubElement(embedded, f"{{{P_NS}}}{style}")
                node.set(f"{{{R_NS}}}id", rid)

        anchor = pres.find(f"{{{P_NS}}}defaultTextStyle")
        if anchor is None:
            pres.append(embedded_list)
        else:
            pres.insert(list(pres).index(anchor), embedded_list)

        if not any(
            el.get("Extension", "").lower() == "fntdata"
            for el in cts.findall(f"{{{CT_NS}}}Default")
        ):
            default = etree.Element(f"{{{CT_NS}}}Default")
            default.set("Extension", "fntdata")
            default.set("ContentType", "application/x-fontdata")
            cts.insert(0, default)

        new_pres = etree.tostring(pres, xml_declaration=True, encoding="UTF-8", standalone=True)
        new_rels = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        new_cts = etree.tostring(cts, xml_declaration=True, encoding="UTF-8", standalone=True)
        new_theme = etree.tostring(theme, xml_declaration=True, encoding="UTF-8", standalone=True)

        tmp = DST.with_name("_deck_embedded.tmp.pptx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename == "ppt/presentation.xml":
                    zout.writestr(info, new_pres)
                elif info.filename == "ppt/_rels/presentation.xml.rels":
                    zout.writestr(info, new_rels)
                elif info.filename == "[Content_Types].xml":
                    zout.writestr(info, new_cts)
                elif info.filename == "ppt/theme/theme1.xml":
                    zout.writestr(info, new_theme)
                else:
                    zout.writestr(info, zin.read(info.filename))
            for name, data in new_parts.items():
                zout.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
        os.replace(tmp, DST)

    with zipfile.ZipFile(DST) as z:
        pres = etree.fromstring(z.read("ppt/presentation.xml"))
        names = [
            face.get("typeface")
            for face in pres.findall(f".//{{{P_NS}}}embeddedFont/{{{P_NS}}}font")
        ]
        parts = [n for n in z.namelist() if n.startswith("ppt/fonts/")]
        if set(names) != FAMILIES:
            raise RuntimeError(f"Embedded families mismatch: {names}")
        for part in parts:
            verify_font_blob(z.read(part), part)
        videos = [n for n in z.namelist() if n.endswith(".mp4")]
        slides = [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
        if len(videos) != 3:
            raise RuntimeError(f"Expected 3 videos, found {len(videos)}")
        if len(slides) != 10:
            raise RuntimeError(f"Expected 10 slides, found {len(slides)}")
        print("embedded families:", names)
        print("font parts:", parts)
        print("videos:", videos)
        print("slides:", len(slides))
        print(DST, DST.stat().st_size)


if __name__ == "__main__":
    main()
