# -*- coding: utf-8 -*-
"""Add restrained native PowerPoint motion to the completed deck.

The animation is intentionally limited to automatic fades with a short
stagger. Background fills, large decorative arcs and fade overlays stay
static so slide changes never flash the whole canvas.
"""

from __future__ import annotations

import os
import zipfile
from pathlib import Path

from lxml import etree


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "_deck_embedded.pptx"
OUTPUT = ROOT / "_deck_animated.pptx"
SLIDE_COUNT = 10
FADE_DURATION = 400
STAGGER_MS = 60
EMU_PER_INCH = 914400

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def q(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def child(parent, namespace: str, name: str, **attrs):
    node = etree.SubElement(parent, q(namespace, name))
    for key, value in attrs.items():
        node.set(key, str(value))
    return node


def shape_c_nv_pr(element):
    if element.tag == q(P, "sp"):
        nv = element.find(q(P, "nvSpPr"))
    elif element.tag == q(P, "pic"):
        nv = element.find(q(P, "nvPicPr"))
    else:
        return None
    if nv is None:
        return None
    return nv.find(q(P, "cNvPr"))


def shape_geometry(element) -> tuple[str | None, int, int]:
    sp_pr = element.find(q(P, "spPr"))
    if sp_pr is None:
        return None, 0, 0
    if sp_pr.find(q(A, "custGeom")) is not None:
        return "custom", 0, 0
    prst = sp_pr.find(q(A, "prstGeom"))
    xfrm = sp_pr.find(q(A, "xfrm"))
    cx = cy = 0
    if xfrm is not None:
        ext = xfrm.find(q(A, "ext"))
        if ext is not None:
            cx = int(ext.get("cx", "0"))
            cy = int(ext.get("cy", "0"))
    return (prst.get("prst") if prst is not None else None), cx, cy


def is_static_shape(element, name: str) -> bool:
    if name.startswith("bg"):
        return True
    geometry, cx, cy = shape_geometry(element)
    if geometry == "custom":
        return True
    if geometry == "ellipse" and max(cx, cy) > 4 * EMU_PER_INCH:
        return True
    return False


def animated_shapes(slide):
    result = []
    tree = slide.find(f"{q(P, 'cSld')}/{q(P, 'spTree')}")
    if tree is None:
        return result
    for element in tree:
        if element.tag not in {q(P, "sp"), q(P, "pic")}:
            continue
        c_nv_pr = shape_c_nv_pr(element)
        if c_nv_pr is None:
            continue
        name = c_nv_pr.get("name", "")
        if is_static_shape(element, name):
            continue
        result.append(
            {
                "id": c_nv_pr.get("id"),
                "name": name,
                "has_text": element.find(f".//{q(A, 't')}") is not None,
                "same_delay_as_previous": name.endswith("_fade"),
            }
        )
    return result


def add_effect(parent, shape_id: str, delay_ms: int, first_id: int):
    par = child(parent, P, "par")
    effect = child(
        par,
        P,
        "cTn",
        id=first_id,
        presetID="10",
        presetClass="entr",
        presetSubtype="0",
        fill="hold",
        grpId="0",
        nodeType="withEffect",
    )
    start = child(effect, P, "stCondLst")
    child(start, P, "cond", delay=delay_ms)
    child_effects = child(effect, P, "childTnLst")

    set_node = child(child_effects, P, "set")
    behavior = child(set_node, P, "cBhvr")
    set_timing = child(behavior, P, "cTn", id=first_id + 1, dur="1", fill="hold")
    cond = child(set_timing, P, "stCondLst")
    child(cond, P, "cond", delay="0")
    target = child(behavior, P, "tgtEl")
    child(target, P, "spTgt", spid=shape_id)
    attrs = child(behavior, P, "attrNameLst")
    child(attrs, P, "attrName").text = "style.visibility"
    to_node = child(set_node, P, "to")
    child(to_node, P, "strVal", val="visible")

    fade = child(child_effects, P, "animEffect", transition="in", filter="fade")
    fade_behavior = child(fade, P, "cBhvr")
    child(fade_behavior, P, "cTn", id=first_id + 2, dur=FADE_DURATION)
    fade_target = child(fade_behavior, P, "tgtEl")
    child(fade_target, P, "spTgt", spid=shape_id)
    return first_id + 3


def build_timing(shapes) -> etree._Element:
    timing = etree.Element(q(P, "timing"), nsmap={"p": P, "a": A})
    tn_list = child(timing, P, "tnLst")
    root_par = child(tn_list, P, "par")
    root_timing = child(
        root_par, P, "cTn", id="1", dur="indefinite", restart="never", nodeType="tmRoot"
    )
    root_children = child(root_timing, P, "childTnLst")
    sequence = child(root_children, P, "seq", concurrent="1", nextAc="seek")
    main_timing = child(sequence, P, "cTn", id="2", dur="indefinite", nodeType="mainSeq")
    main_children = child(main_timing, P, "childTnLst")

    group_par = child(main_children, P, "par")
    group_timing = child(group_par, P, "cTn", id="3", fill="hold")
    group_start = child(group_timing, P, "stCondLst")
    child(group_start, P, "cond", delay="indefinite")
    begin_cond = child(group_start, P, "cond", evt="onBegin", delay="0")
    child(begin_cond, P, "tn", val="2")
    group_children = child(group_timing, P, "childTnLst")

    item_par = child(group_children, P, "par")
    item_timing = child(item_par, P, "cTn", id="4", fill="hold")
    item_start = child(item_timing, P, "stCondLst")
    child(item_start, P, "cond", delay="0")
    effects = child(item_timing, P, "childTnLst")

    next_id = 5
    delay = 0
    for index, shape in enumerate(shapes):
        if index > 0 and not shape["same_delay_as_previous"]:
            delay += STAGGER_MS
        next_id = add_effect(effects, shape["id"], delay, next_id)

    previous = child(sequence, P, "prevCondLst")
    previous_cond = child(previous, P, "cond", evt="onPrev", delay="0")
    child(child(previous_cond, P, "tgtEl"), P, "sldTgt")
    following = child(sequence, P, "nextCondLst")
    following_cond = child(following, P, "cond", evt="onNext", delay="0")
    child(child(following_cond, P, "tgtEl"), P, "sldTgt")

    text_shapes = [shape for shape in shapes if shape["has_text"]]
    if text_shapes:
        build_list = child(timing, P, "bldLst")
        for shape in text_shapes:
            child(build_list, P, "bldP", spid=shape["id"], grpId="0")
    return timing


def add_transition_and_timing(slide, shapes, slide_number: int) -> None:
    for existing in slide.findall(q(P, "transition")) + slide.findall(q(P, "timing")):
        slide.remove(existing)

    transition = etree.Element(q(P, "transition"), spd="slow" if slide_number in {1, SLIDE_COUNT} else "med")
    transition.set("advClick", "1")
    child(transition, P, "fade")
    timing = build_timing(shapes)

    clr = slide.find(q(P, "clrMapOvr"))
    if clr is not None:
        clr.addnext(transition)
        transition.addnext(timing)
    else:
        slide.append(transition)
        slide.append(timing)


def audit_bytes(data: bytes, slide_number: int) -> tuple[int, int]:
    root = etree.fromstring(data)
    timing = root.find(q(P, "timing"))
    transition = root.find(q(P, "transition"))
    if timing is None or transition is None:
        raise RuntimeError(f"slide {slide_number}: missing timing or transition")
    effects = [node for node in timing.iter(q(P, "cTn")) if node.get("presetID") == "10"]
    clicks = [node for node in effects if node.get("nodeType") == "clickEffect"]
    targets = {node.get("spid") for node in timing.iter(q(P, "spTgt"))}
    shape_ids = {
        c_nv.get("id")
        for element in root.iter()
        if element.tag in {q(P, "sp"), q(P, "pic")}
        for c_nv in [shape_c_nv_pr(element)]
        if c_nv is not None
    }
    missing = targets - shape_ids
    if clicks or missing:
        raise RuntimeError(f"slide {slide_number}: clicks={len(clicks)}, missing={sorted(missing)}")
    return len(effects), len(targets)


def main() -> None:
    if not SOURCE.exists():
        raise FileNotFoundError(SOURCE)

    slide_data: dict[str, bytes] = {}
    total_effects = 0
    with zipfile.ZipFile(SOURCE) as source:
        for number in range(1, SLIDE_COUNT + 1):
            name = f"ppt/slides/slide{number}.xml"
            slide = etree.fromstring(source.read(name))
            shapes = animated_shapes(slide)
            if not shapes:
                raise RuntimeError(f"slide {number}: no animatable shapes")
            add_transition_and_timing(slide, shapes, number)
            data = etree.tostring(
                slide, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            effect_count, target_count = audit_bytes(data, number)
            if effect_count != target_count or effect_count != len(shapes):
                raise RuntimeError(
                    f"slide {number}: shapes={len(shapes)}, effects={effect_count}, targets={target_count}"
                )
            slide_data[name] = data
            total_effects += effect_count
            print(f"slide {number}: {effect_count} automatic fade effects")

        temporary = OUTPUT.with_name(OUTPUT.stem + ".tmp" + OUTPUT.suffix)
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as output:
            for info in source.infolist():
                output.writestr(info, slide_data.get(info.filename, source.read(info.filename)))
        os.replace(temporary, OUTPUT)

    with zipfile.ZipFile(OUTPUT) as check:
        if check.testzip() is not None:
            raise RuntimeError("animation output zip is damaged")
        if sum(name.startswith("ppt/fonts/") and name.endswith(".fntdata") for name in check.namelist()) != 3:
            raise RuntimeError("embedded font parts were not preserved")
        if sum(name.lower().endswith(".mp4") for name in check.namelist()) != 3:
            raise RuntimeError("embedded videos were not preserved")

    print(f"output: {OUTPUT}")
    print(f"total: {total_effects} effects; click-triggered effects: 0")


if __name__ == "__main__":
    main()
