# -*- coding: utf-8 -*-
"""Copy full EOT payloads for 楷体 / 华文中宋 from the audited legacy deck."""
import copy
import io
import os
import pathlib
import struct
import zipfile

from fontTools.ttLib import TTFont
from lxml import etree


P = pathlib.Path(r"C:\Users\camel\Documents\ChatGPT\Hig1\projects\pharma_career_ppt169_20261023")
SRC = P / "_deck_build.pptx"
DST = P / "_deck_final.pptx"
DONOR = pathlib.Path(r"C:\Users\camel\Documents\ChatGPT\Hig1\projects\duoyuhua_ppt169_20261021\多余的话_瑞士国际主义版式.pptx")

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
FONT_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
FAMILIES = {"楷体", "华文中宋"}
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


def main():
    fonts = donor_fonts()
    if {face.get("typeface") for face, _ in fonts} != FAMILIES:
        raise RuntimeError("Donor deck does not contain both required families")

    with zipfile.ZipFile(SRC) as zin:
        pres = etree.fromstring(zin.read("ppt/presentation.xml"))
        rels = etree.fromstring(zin.read("ppt/_rels/presentation.xml.rels"))
        cts = etree.fromstring(zin.read("[Content_Types].xml"))
        existing_ids = [rel.get("Id") for rel in rels.findall(f"{{{REL_NS}}}Relationship")]
        next_id = max(int(x[3:]) for x in existing_ids if x.startswith("rId")) + 1

        embedded_list = etree.Element(f"{{{P_NS}}}embeddedFontLst")
        new_parts = {}
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

        tmp = DST.with_name("_deck_final.tmp.pptx")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename == "ppt/presentation.xml":
                    zout.writestr(info, new_pres)
                elif info.filename == "ppt/_rels/presentation.xml.rels":
                    zout.writestr(info, new_rels)
                elif info.filename == "[Content_Types].xml":
                    zout.writestr(info, new_cts)
                else:
                    zout.writestr(info, zin.read(info.filename))
            for name, data in new_parts.items():
                zout.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
        os.replace(tmp, DST)

    with zipfile.ZipFile(DST) as z:
        pres = etree.fromstring(z.read("ppt/presentation.xml"))
        names = re_font_names(pres)
        parts = [n for n in z.namelist() if n.startswith("ppt/fonts/")]
        print("embedded families:", names)
        print("font parts:", parts)
        for part in parts:
            data = z.read(part)
            eot_size, font_size, version, flags = struct.unpack("<IIII", data[:16])
            header_size = eot_size - font_size
            ttf = TTFont(io.BytesIO(data[header_size:]))
            print(part, "flags=", flags, "glyphs=", len(ttf.getGlyphOrder()), "bytes=", len(data))
            assert flags & 0x5 == 0, "subset flags present"
            assert len(ttf.getGlyphOrder()) > 1000, "unexpected subset glyph count"
    print(DST, DST.stat().st_size)


def re_font_names(pres):
    return [
        face.get("typeface")
        for face in pres.findall(f".//{{{P_NS}}}embeddedFont/{{{P_NS}}}font")
    ]


if __name__ == "__main__":
    main()

