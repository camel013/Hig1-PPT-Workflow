"""Merge legacy full-font embedding and staggered fade timings into a new build.

The legacy deck was produced and audited earlier (fonts flags=0). Reusing its
font payload avoids depending on PowerPoint COM, which is unstable on this PC.
Shape ids are remapped by reading order, because the new build adds background
haze/contour shapes before the content.
"""
import argparse
import copy
import zipfile

from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
FONT_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"


def tag(ns, name):
    return f"{{{ns}}}{name}"


def shape_records(slide_xml):
    root = etree.fromstring(slide_xml)
    records = []
    for sp in root.iter(tag(P, "sp")):
        nv = sp.find(tag(P, "nvSpPr"))
        c_nv = None if nv is None else nv.find(tag(P, "cNvPr"))
        if c_nv is None:
            continue
        name = c_nv.get("name", "")
        if not name or name.startswith("bg"):
            continue
        text = "".join(t.text or "" for t in sp.iter(tag(A, "t")))
        xfrm = sp.find(f"{tag(P, 'spPr')}/{tag(A, 'xfrm')}")
        off = None if xfrm is None else xfrm.find(tag(A, "off"))
        ext = None if xfrm is None else xfrm.find(tag(A, "ext"))
        records.append({
            "id": c_nv.get("id"),
            "name": name,
            "text": text,
            "x": int(off.get("x")) if off is not None else 0,
            "y": int(off.get("y")) if off is not None else 0,
            "w": int(ext.get("cx")) if ext is not None else 0,
            "h": int(ext.get("cy")) if ext is not None else 0,
        })
    return records


def _close(a, b, tol=4):
    return abs(a - b) <= tol


def map_shapes(old_records, new_records):
    """Match by text first, then by exact geometry; tolerate removed decor blocks."""
    used = set()
    mapping = {}
    for old in old_records:
        if not old["text"]:
            continue
        candidates = [n for n in new_records if n["id"] not in used and n["text"] == old["text"]]
        if candidates:
            new = min(candidates, key=lambda n: abs(n["x"] - old["x"]) + abs(n["y"] - old["y"]))
            mapping[old["id"]] = new["id"]
            used.add(new["id"])
    for old in old_records:
        if old["id"] in mapping:
            continue
        candidates = [
            n for n in new_records
            if n["id"] not in used
            and _close(old["x"], n["x"]) and _close(old["y"], n["y"])
            and _close(old["w"], n["w"]) and _close(old["h"], n["h"])
        ]
        if candidates:
            mapping[old["id"]] = candidates[0]["id"]
            used.add(candidates[0]["id"])
    return mapping


def remap_timing(timing, mapping):
    """Drop whole effect blocks whose target no longer exists, remap the rest."""
    dropped = 0
    kept = 0
    for par in list(timing.iter(tag(P, "par"))):
        c_tn = par.find(tag(P, "cTn"))
        if c_tn is None or c_tn.get("presetID") != "10":
            continue
        targets = list(par.iter(tag(P, "spTgt")))
        if not targets:
            continue
        old_id = targets[0].get("spid")
        if old_id not in mapping:
            par.getparent().remove(par)
            dropped += 1
            continue
        for target in targets:
            target.set("spid", mapping[old_id])
        kept += 1
    return kept, dropped


def merge(new_path, legacy_path, out_path):
    new_zip = zipfile.ZipFile(new_path)
    old_zip = zipfile.ZipFile(legacy_path)

    new_pres = etree.fromstring(new_zip.read("ppt/presentation.xml"))
    old_pres = etree.fromstring(old_zip.read("ppt/presentation.xml"))
    new_rels = etree.fromstring(new_zip.read("ppt/_rels/presentation.xml.rels"))
    old_rels = etree.fromstring(old_zip.read("ppt/_rels/presentation.xml.rels"))

    old_targets = {rel.get("Id"): rel.get("Target") for rel in old_rels}
    used = {int(rel.get("Id")[3:]) for rel in new_rels if rel.get("Id", "").startswith("rId") and rel.get("Id")[3:].isdigit()}
    next_id = max(used or {0}) + 1
    font_map = {}

    old_fonts = old_pres.find(tag(P, "embeddedFontLst"))
    if old_fonts is None:
        raise RuntimeError("legacy deck has no embeddedFontLst")
    new_fonts = copy.deepcopy(old_fonts)
    for node in new_fonts.iter():
        old_id = node.get(tag(R, "id"))
        if not old_id:
            continue
        if old_id not in font_map:
            font_map[old_id] = f"rId{next_id}"
            next_id += 1
            rel = etree.SubElement(new_rels, tag(REL, "Relationship"))
            rel.set("Id", font_map[old_id])
            rel.set("Type", FONT_TYPE)
            rel.set("Target", old_targets[old_id])
        node.set(tag(R, "id"), font_map[old_id])

    old_fonts.getparent().remove(old_fonts)
    anchor = new_pres.find(tag(P, "defaultTextStyle"))
    if anchor is None:
        new_pres.append(new_fonts)
    else:
        anchor.addprevious(new_fonts)
    new_pres.set("embedTrueTypeFonts", "1")
    new_pres.set("saveSubsetFonts", "1")
    new_pres.set("autoCompressPictures", "0")

    slide_timings = {}
    for index in range(1, 13):
        old_slide = etree.fromstring(old_zip.read(f"ppt/slides/slide{index}.xml"))
        new_slide = etree.fromstring(new_zip.read(f"ppt/slides/slide{index}.xml"))
        timing = old_slide.find(tag(P, "timing"))
        if timing is None:
            continue
        old_records = shape_records(old_zip.read(f"ppt/slides/slide{index}.xml"))
        new_records = shape_records(new_zip.read(f"ppt/slides/slide{index}.xml"))
        mapping = map_shapes(old_records, new_records)
        timing = copy.deepcopy(timing)
        kept, dropped = remap_timing(timing, mapping)
        print(f"slide {index}: effects {kept} kept, {dropped} dropped")
        old_existing = new_slide.find(tag(P, "timing"))
        if old_existing is not None:
            new_slide.remove(old_existing)
        clr = new_slide.find(tag(P, "clrMapOvr"))
        if clr is not None:
            clr.addnext(timing)
        else:
            new_slide.append(timing)
        slide_timings[f"ppt/slides/slide{index}.xml"] = etree.tostring(new_slide, xml_declaration=True, encoding="UTF-8", standalone=True)

    content_types = etree.fromstring(new_zip.read("[Content_Types].xml"))
    if not any(d.get("Extension") == "fntdata" for d in content_types.findall(tag(CT, "Default"))):
        default = etree.Element(tag(CT, "Default"))
        default.set("Extension", "fntdata")
        default.set("ContentType", "application/x-fontdata")
        content_types.insert(0, default)

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as out:
        for info in new_zip.infolist():
            data = new_zip.read(info.filename)
            if info.filename == "ppt/presentation.xml":
                data = etree.tostring(new_pres, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif info.filename == "ppt/_rels/presentation.xml.rels":
                data = etree.tostring(new_rels, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif info.filename == "[Content_Types].xml":
                data = etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif info.filename in slide_timings:
                data = slide_timings[info.filename]
            out.writestr(info.filename, data)
        for name in old_zip.namelist():
            if name.startswith("ppt/fonts/") and name.endswith(".fntdata"):
                out.writestr(name, old_zip.read(name))

    print(f"fonts: {len(font_map)} parts, timings: {len(slide_timings)} slides")
    print(out_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("new")
    parser.add_argument("legacy")
    parser.add_argument("out")
    args = parser.parse_args()
    merge(args.new, args.legacy, args.out)
