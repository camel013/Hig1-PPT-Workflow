from __future__ import annotations

import os
import sys
import zipfile
from pathlib import Path

from lxml import etree


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = {
    "r": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def font_parts_by_family(archive: zipfile.ZipFile) -> dict[str, list[bytes]]:
    presentation = etree.fromstring(archive.read("ppt/presentation.xml"))
    rels = etree.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
    targets = {
        rel.get("Id"): rel.get("Target")
        for rel in rels.findall("r:Relationship", REL_NS)
    }
    result = {}
    for embedded in presentation.xpath(".//p:embeddedFontLst/p:embeddedFont", namespaces=NS):
        face = embedded.find("p:font", NS)
        if face is None:
            continue
        family = face.get("typeface")
        data = []
        for child in embedded:
            rid = child.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            if rid and rid in targets:
                target = targets[rid].lstrip("/")
                part = target if target.startswith("ppt/") else "ppt/" + target.lstrip("./")
                data.append((part, archive.read(part)))
        result[family] = data
    return result


def embed(src: Path, donor: Path, families: set[str]):
    temp = src.with_name(src.stem + ".fontrestore" + src.suffix)
    with zipfile.ZipFile(donor) as source, zipfile.ZipFile(src) as target_zip:
        donor_parts = font_parts_by_family(source)
        target_parts = font_parts_by_family(target_zip)
        replacements = {}
        for family in families:
            source_items = donor_parts.get(family, [])
            target_items = target_parts.get(family, [])
            if len(source_items) != len(target_items):
                raise RuntimeError(
                    f"Font part mismatch for {family}: "
                    f"{len(source_items)} donor vs {len(target_items)} target"
                )
            for (_, donor_data), (target_part, _) in zip(source_items, target_items):
                replacements[target_part] = donor_data
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as output:
            for info in target_zip.infolist():
                output.writestr(
                    info, replacements.get(info.filename, target_zip.read(info.filename))
                )
    try:
        os.replace(temp, src)
    finally:
        if temp.exists():
            temp.unlink()
    return {
        "deck": str(src),
        "families": sorted(families),
        "parts": sorted(replacements),
        "bytes": src.stat().st_size,
    }


if __name__ == "__main__":
    donor = Path(sys.argv[1])
    for value in sys.argv[2:]:
        print(embed(Path(value), donor, {"楷体", "华文中宋"}))
