# -*- coding: utf-8 -*-
"""Build sentence-level entrance groups for the 95-slide teaching deck.

The deck is drawn as many short text boxes because its layout engine wraps and
sets every line separately.  Animating those boxes one by one makes even a
paragraph look like a character-by-character build.  This pass reconstructs
visual units first: a numbered step together with its heading and copy, a card
together with its label, a paragraph, or a compact table section.  Each unit is
then wrapped in one DrawingML group and receives one entrance effect.

Rules of thumb:
* header/page furniture stays out of the animation;
* full-bleed colour fields, arcs, contour lines and hairlines stay fixed;
* a page contains at most six semantic units;
* fragments shorter than three characters are absorbed by the nearest unit;
* there is no p:bldLst, so PowerPoint cannot split the text by paragraph or by
  character.  Timing varies with the number and kind of units on the page.
"""
from __future__ import annotations

import argparse
import math
import re
import zipfile
from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

SLIDE_W = 1280.0
SLIDE_H = 720.0
PX = 9600.0
SHAPE_TAGS = {
    f"{{{P}}}sp",
    f"{{{P}}}pic",
    f"{{{P}}}graphicFrame",
    f"{{{P}}}grpSp",
}


def tag(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def _c_nv_pr(element: etree._Element) -> etree._Element | None:
    return element.find(".//" + tag(P, "cNvPr"))


def _bbox(element: etree._Element) -> tuple[float, float, float, float] | None:
    if element.tag == tag(P, "graphicFrame"):
        xfrm = element.find(tag(P, "xfrm"))
    elif element.tag == tag(P, "grpSp"):
        xfrm = element.find(f"{tag(P, 'grpSpPr')}/{tag(A, 'xfrm')}")
    else:
        xfrm = element.find(f"{tag(P, 'spPr')}/{tag(A, 'xfrm')}")
    if xfrm is None:
        return None
    off = xfrm.find(tag(A, "off"))
    ext = xfrm.find(tag(A, "ext"))
    if off is None or ext is None:
        return None
    return (
        int(off.get("x", "0")) / PX,
        int(off.get("y", "0")) / PX,
        int(ext.get("cx", "0")) / PX,
        int(ext.get("cy", "0")) / PX,
    )


def _shape_text(element: etree._Element) -> str:
    return "".join(node.text or "" for node in element.iter(tag(A, "t"))).strip()


def _x_gap(a: "Block", b: "Block") -> float:
    return max(0.0, max(a.x, b.x) - min(a.x2, b.x2))


def _y_gap(a: "Block", b: "Block") -> float:
    return max(0.0, max(a.y, b.y) - min(a.y2, b.y2))


def _overlap(a0: float, a1: float, b0: float, b1: float) -> float:
    return max(0.0, min(a1, b1) - max(a0, b0))


def _compact(block: "Block") -> str:
    return "".join(item.text for item in block.items).strip()


@dataclass
class ShapeRec:
    node: etree._Element
    index: int
    shape_id: int
    name: str
    descr: str
    kind: str
    x: float
    y: float
    w: float
    h: float
    text: str
    static: bool

    @property
    def x2(self) -> float:
        return self.x + self.w

    @property
    def y2(self) -> float:
        return self.y + self.h


@dataclass
class Block:
    items: list[ShapeRec] = field(default_factory=list)
    source: str = "text"
    x: float = 0.0
    y: float = 0.0
    w: float = 0.0
    h: float = 0.0
    kind: str = "body"

    @property
    def x2(self) -> float:
        return self.x + self.w

    @property
    def y2(self) -> float:
        return self.y + self.h

    @property
    def is_container(self) -> bool:
        return self.source == "container"

    @property
    def has_visual(self) -> bool:
        return any(not item.text and item.w * item.h >= 12000 for item in self.items)


def _make_block(items: list[ShapeRec], source: str = "text") -> Block:
    left = min(item.x for item in items)
    top = min(item.y for item in items)
    right = max(item.x2 for item in items)
    bottom = max(item.y2 for item in items)
    kinds = {item.kind for item in items if item.text}
    text_kinds = {item.kind for item in items if item.text}
    kind = "heading" if text_kinds == {"heading"} else "body"
    return Block(
        items=list(items),
        source=source,
        x=left,
        y=top,
        w=max(0.1, right - left),
        h=max(0.1, bottom - top),
        kind=kind,
    )


def _merge(a: Block, b: Block, source: str | None = None) -> Block:
    return _make_block(a.items + b.items, source or (a.source if a.source == b.source else "merged"))


def _union_nodes(blocks: list[Block]) -> Block:
    return _make_block([item for block in blocks for item in block.items], "fragments")


def collect_shapes(sp_tree: etree._Element) -> list[ShapeRec]:
    shapes: list[ShapeRec] = []
    for index, node in enumerate(list(sp_tree)):
        if node.tag not in SHAPE_TAGS:
            continue
        c_nv = _c_nv_pr(node)
        bbox = _bbox(node)
        if c_nv is None or bbox is None:
            continue
        name = c_nv.get("name", "")
        descr = c_nv.get("descr", "") or ""
        text = _shape_text(node)
        kind = (
            descr.rsplit(":", 1)[-1]
            if descr.startswith("animgrp:")
            else "heading" if name.startswith("heading_") else "body" if name.startswith("body_") else "graphic"
        )
        x, y, w, h = bbox
        if w <= 0 or h <= 0:
            continue
        static = False
        if name.lower().startswith("bg") or y >= 668:
            static = True
        elif not text:
            if h >= 500 or (w >= 1050 and h >= 300):
                static = True
        shapes.append(
            ShapeRec(
                node=node,
                index=index,
                shape_id=int(c_nv.get("id", "0")),
                name=name,
                descr=descr,
                kind=kind,
                x=x,
                y=y,
                w=w,
                h=h,
                text=text,
                static=static,
            )
        )
    return shapes


def _row_merge(blocks: list[Block]) -> list[Block]:
    # Dense tables already have a column rhythm; do not chain their cells.
    if len(blocks) > 32:
        return blocks
    changed = True
    while changed:
        changed = False
        for i, a in enumerate(blocks):
            for j in range(i + 1, len(blocks)):
                b = blocks[j]
                if a.is_container or b.is_container:
                    continue
                overlap = _overlap(a.y, a.y2, b.y, b.y2)
                cy_a = (a.y + a.y2) / 2
                cy_b = (b.y + b.y2) / 2
                short = min(a.w, b.w) <= 110 and min(a.h, b.h) <= 40
                kinds_a = {item.kind for item in a.items}
                kinds_b = {item.kind for item in b.items}
                mixed = bool(kinds_a ^ kinds_b)
                gap = _x_gap(a, b)
                near = gap <= 14 or (short and gap <= 62) or (mixed and gap <= 60)
                if overlap >= 0.28 * min(a.h, b.h) and abs(cy_a - cy_b) <= 22 and near and (short or mixed):
                    blocks[i] = _merge(a, b)
                    blocks.pop(j)
                    changed = True
                    break
            if changed:
                break
    return blocks


def _column_merge(blocks: list[Block]) -> list[Block]:
    changed = True
    while changed:
        changed = False
        for i, a in enumerate(blocks):
            for j in range(i + 1, len(blocks)):
                b = blocks[j]
                if a.is_container or b.is_container:
                    continue
                ca = (a.x + a.x2) / 2
                cb = (b.x + b.x2) / 2
                aligned = min(abs(a.x - b.x), abs(a.x2 - b.x2), abs(ca - cb)) <= max(
                    28, 0.24 * min(a.w, b.w)
                )
                if (
                    aligned
                    and _y_gap(a, b) <= 78
                    and max(a.w, b.w) < 0.56 * SLIDE_W
                ):
                    blocks[i] = _merge(a, b)
                    blocks.pop(j)
                    changed = True
                    break
            if changed:
                break
    return blocks


def _attach_containers(blocks: list[Block], content: list[ShapeRec]) -> list[Block]:
    containers = [
        item
        for item in content
        if not item.text
        and item.w >= 80
        and item.h >= 25
        and item.w * item.h >= 18000
        and item.name.startswith(("Rectangle", "Block Arc"))
    ]
    containers.sort(key=lambda item: item.w * item.h, reverse=True)
    used: set[int] = set()
    for container in containers:
        inside: list[int] = []
        for i, block in enumerate(blocks):
            if i in used or block.source == "header":
                continue
            overlap_y = _overlap(container.y, container.y2, block.y, block.y2)
            overlap_x = _overlap(container.x, container.x2, block.x, block.x2)
            if overlap_y / max(1.0, block.h) >= 0.55 and overlap_x / max(1.0, block.w) >= 0.30:
                inside.append(i)
        if inside:
            items = [container]
            for i in inside:
                items += blocks[i].items
                used.add(i)
            blocks.append(_make_block(items, "container"))
    return [block for i, block in enumerate(blocks) if i not in used]


def _attach_leftovers(blocks: list[Block], content: list[ShapeRec]) -> list[Block]:
    attached = {id(item) for block in blocks for item in block.items}
    for item in content:
        if id(item) in attached:
            continue
        if not blocks:
            blocks.append(_make_block([item], "graphic"))
            continue
        nearest = min(
            blocks,
            key=lambda block: math.hypot(_x_gap(block, _make_block([item])), _y_gap(block, _make_block([item]))),
        )
        nearest.items.append(item)
        attached.add(id(item))
    return [_make_block(block.items, block.source) for block in blocks]


def _absorb_fragments(blocks: list[Block]) -> list[Block]:
    changed = True
    while changed:
        changed = False
        for fragment in list(blocks):
            text = _compact(fragment)
            if not text or len(text) > 2:
                continue
            others = [b for b in blocks if b is not fragment and _compact(b)]
            if not others:
                continue
            non_header = [b for b in others if b.source != "header"]
            pool = non_header or others

            def score(block: Block) -> float:
                dx = _x_gap(fragment, block)
                dy = _y_gap(fragment, block)
                return dx + dy + (18 if dx and dy else 0)

            nearest = min(pool, key=score)
            merged = _merge(nearest, fragment)
            blocks = [b for b in blocks if b is not fragment and b is not nearest]
            blocks.append(merged)
            changed = True
            break
    return blocks


def _absorb_ornaments(blocks: list[Block]) -> list[Block]:
    """Keep rules, strips and markers inside the unit whose z-layer they share."""
    changed = True
    while changed:
        changed = False
        for ornament in list(blocks):
            if ornament.is_container or any(item.text for item in ornament.items):
                continue
            others = [block for block in blocks if block is not ornament and any(item.text for item in block.items)]
            if not others:
                continue

            def score(block: Block) -> float:
                dx = _x_gap(ornament, block)
                dy = _y_gap(ornament, block)
                return dx + dy + (18 if dx and dy else 0)

            nearest = min(others, key=score)
            merged = _merge(nearest, ornament)
            blocks = [block for block in blocks if block is not ornament and block is not nearest]
            blocks.append(merged)
            changed = True
            break
    return blocks


def _merge_to_limit(blocks: list[Block], limit: int) -> list[Block]:
    while len(blocks) > limit:
        best: tuple[float, int, int] | None = None
        for i, a in enumerate(blocks):
            for j in range(i + 1, len(blocks)):
                b = blocks[j]
                dx = _x_gap(a, b)
                dy = _y_gap(a, b)
                score = dx + dy + (18 if dx and dy else 0)
                if best is None or score < best[0]:
                    best = (score, i, j)
        assert best is not None
        _, i, j = best
        blocks[i] = _merge(blocks[i], blocks[j])
        blocks.pop(j)
    return blocks


def _cover_blocks(content: list[ShapeRec]) -> list[Block]:
    left = [item for item in content if item.x < 900]
    right_title = [item for item in content if item.x >= 900 and item.y < 400]
    edition = [item for item in content if item.x >= 900 and item.y >= 400]
    groups = []
    for items, source in ((left, "cover_left"), (right_title, "cover_title"), (edition, "cover_edition")):
        if items:
            groups.append(_make_block(items, source))
    return groups


def semantic_blocks(recs: list[ShapeRec], page: int) -> list[Block]:
    content = [item for item in recs if not item.static]
    if page == 1:
        blocks = _cover_blocks(content)
        return _merge_to_limit(blocks, 3)

    header = [item for item in content if item.y < 125]
    rest = [item for item in content if item not in header]
    blocks: list[Block] = []
    if header:
        blocks.append(_make_block(header, "header"))

    by_descr: OrderedDict[str, list[ShapeRec]] = OrderedDict()
    for item in rest:
        key = item.descr if item.descr.startswith("animgrp:") else f"shape-{item.shape_id}"
        by_descr.setdefault(key, []).append(item)
    for items in by_descr.values():
        blocks.append(_make_block(items, "text"))

    blocks = _attach_containers(blocks, rest)
    blocks = _row_merge(blocks)
    blocks = _column_merge(blocks)
    blocks = _attach_leftovers(blocks, rest)
    blocks = _absorb_ornaments(blocks)
    blocks = _absorb_fragments(blocks)
    return _merge_to_limit(blocks, min(6, len(blocks)))


def _make_group(block: Block, group_id: int) -> etree._Element:
    group = etree.Element(tag(P, "grpSp"))
    nv = etree.SubElement(group, tag(P, "nvGrpSpPr"))
    etree.SubElement(
        nv,
        tag(P, "cNvPr"),
        id=str(group_id),
        name=f"anim_{block.kind}_group_{group_id}",
        descr="semantic-animation-unit",
    )
    etree.SubElement(nv, tag(P, "cNvGrpSpPr"))
    etree.SubElement(nv, tag(P, "nvPr"))
    group_pr = etree.SubElement(group, tag(P, "grpSpPr"))
    xfrm = etree.SubElement(group_pr, tag(A, "xfrm"))
    x = str(int(round(block.x * PX)))
    y = str(int(round(block.y * PX)))
    cx = str(max(1, int(round(block.w * PX))))
    cy = str(max(1, int(round(block.h * PX))))
    etree.SubElement(xfrm, tag(A, "off"), x=x, y=y)
    etree.SubElement(xfrm, tag(A, "ext"), cx=cx, cy=cy)
    etree.SubElement(xfrm, tag(A, "chOff"), x=x, y=y)
    etree.SubElement(xfrm, tag(A, "chExt"), cx=cx, cy=cy)
    for item in sorted(block.items, key=lambda rec: rec.index):
        group.append(item.node)
    return group


def wrap_blocks(root: etree._Element, blocks: list[Block]) -> list[dict[str, object]]:
    sp_tree = root.find(".//" + tag(P, "spTree"))
    if sp_tree is None:
        return []
    max_id = max((int(_c_nv_pr(node).get("id", "0")) for node in root.iter() if _c_nv_pr(node) is not None), default=1)
    next_id = max_id + 1
    original = list(sp_tree)
    node_to_target: dict[etree._Element, etree._Element] = {}
    target_last_index: dict[int, int] = {}
    seen: set[int] = set()
    targets: list[dict[str, object]] = []

    for order, block in enumerate(sorted(blocks, key=lambda b: (round(b.y / 80), b.x))):
        if len(block.items) == 1:
            target = block.items[0].node
        else:
            target = _make_group(block, next_id)
            next_id += 1
        target_id = int(_c_nv_pr(target).get("id", "0"))
        target_last_index[id(target)] = max(item.index for item in block.items)
        for item in block.items:
            node_to_target[item.node] = target
        targets.append(
            {
                "id": target_id,
                "kind": block.kind,
                "source": block.source,
                "header": block.source == "header",
                "visual": block.has_visual,
                "text": _compact(block),
                "order": order,
            }
        )

    rebuilt: list[etree._Element] = []
    for node_index, node in enumerate(original):
        target = node_to_target.get(node)
        if target is None:
            rebuilt.append(node)
        elif id(target) not in seen and target_last_index.get(id(target)) == node_index:
            rebuilt.append(target)
            seen.add(id(target))

    for node in list(sp_tree):
        sp_tree.remove(node)
    for node in rebuilt:
        sp_tree.append(node)
    return targets


def _effect_ids(start: int) -> tuple[int, int, int, int, int]:
    return tuple(range(start, start + 5))  # type: ignore[return-value]


def make_effect(
    parent: etree._Element,
    shape_id: int,
    delay: int,
    duration: int,
    ids: tuple[int, int, int, int, int],
    preset_id: int,
    preset_subtype: int,
    effect_filter: str,
) -> None:
    outer_id, inner_id, effect_id, set_id, anim_id = ids
    outer = etree.SubElement(parent, tag(P, "par"))
    outer_ctn = etree.SubElement(outer, tag(P, "cTn"), id=str(outer_id), fill="hold")
    outer_cond = etree.SubElement(outer_ctn, tag(P, "stCondLst"))
    etree.SubElement(outer_cond, tag(P, "cond"), delay="indefinite")
    begin = etree.SubElement(outer_cond, tag(P, "cond"), evt="onBegin", delay="0")
    etree.SubElement(begin, tag(P, "tn"), val="2")
    outer_child = etree.SubElement(outer_ctn, tag(P, "childTnLst"))

    inner = etree.SubElement(outer_child, tag(P, "par"))
    inner_ctn = etree.SubElement(inner, tag(P, "cTn"), id=str(inner_id), fill="hold")
    inner_cond = etree.SubElement(inner_ctn, tag(P, "stCondLst"))
    etree.SubElement(inner_cond, tag(P, "cond"), delay="0")
    inner_child = etree.SubElement(inner_ctn, tag(P, "childTnLst"))

    effect = etree.SubElement(inner_child, tag(P, "par"))
    effect_ctn = etree.SubElement(
        effect,
        tag(P, "cTn"),
        id=str(effect_id),
        presetID=str(preset_id),
        presetClass="entr",
        presetSubtype=str(preset_subtype),
        fill="hold",
        grpId="0",
        nodeType="withEffect",
    )
    effect_cond = etree.SubElement(effect_ctn, tag(P, "stCondLst"))
    etree.SubElement(effect_cond, tag(P, "cond"), delay=str(delay))
    effect_child = etree.SubElement(effect_ctn, tag(P, "childTnLst"))

    set_node = etree.SubElement(effect_child, tag(P, "set"))
    set_cbhvr = etree.SubElement(set_node, tag(P, "cBhvr"))
    set_ctn = etree.SubElement(set_cbhvr, tag(P, "cTn"), id=str(set_id), dur="1", fill="hold")
    set_cond = etree.SubElement(set_ctn, tag(P, "stCondLst"))
    etree.SubElement(set_cond, tag(P, "cond"), delay="0")
    set_target = etree.SubElement(set_cbhvr, tag(P, "tgtEl"))
    etree.SubElement(set_target, tag(P, "spTgt"), spid=str(shape_id))
    attrs = etree.SubElement(set_cbhvr, tag(P, "attrNameLst"))
    etree.SubElement(attrs, tag(P, "attrName")).text = "style.visibility"
    to_node = etree.SubElement(set_node, tag(P, "to"))
    etree.SubElement(to_node, tag(P, "strVal"), val="visible")

    anim = etree.SubElement(effect_child, tag(P, "animEffect"), transition="in", filter=effect_filter)
    anim_cbhvr = etree.SubElement(anim, tag(P, "cBhvr"))
    etree.SubElement(anim_cbhvr, tag(P, "cTn"), id=str(anim_id), dur=str(duration))
    anim_target = etree.SubElement(anim_cbhvr, tag(P, "tgtEl"))
    etree.SubElement(anim_target, tag(P, "spTgt"), spid=str(shape_id))


def timing_xml(shapes: list[dict[str, object]], page: int) -> etree._Element:
    timing = etree.Element(tag(P, "timing"))
    tn_list = etree.SubElement(timing, tag(P, "tnLst"))
    root_par = etree.SubElement(tn_list, tag(P, "par"))
    root_ctn = etree.SubElement(
        root_par,
        tag(P, "cTn"),
        id="1",
        dur="indefinite",
        restart="never",
        nodeType="tmRoot",
    )
    root_child = etree.SubElement(root_ctn, tag(P, "childTnLst"))
    seq = etree.SubElement(root_child, tag(P, "seq"), concurrent="1", nextAc="seek")
    seq_ctn = etree.SubElement(seq, tag(P, "cTn"), id="2", dur="indefinite", nodeType="mainSeq")
    seq_child = etree.SubElement(seq_ctn, tag(P, "childTnLst"))

    count = max(1, len(shapes))
    step = {3: 180, 4: 150, 5: 120, 6: 92}.get(count, 110)
    delay = 0
    next_id = 3
    for index, shape in enumerate(shapes):
        if index:
            jitter = ((page * 7 + index * 5) % 5 - 2) * 11
            delay += max(55, step + jitter)
        header = bool(shape.get("header"))
        visual = bool(shape.get("visual"))
        if header:
            duration = 340 + (page % 3) * 30
            preset_id, subtype, effect_filter = 22, 4, "wipe(up)"
        elif visual:
            duration = 470 + ((page + index) % 3) * 35
            preset_id, subtype, effect_filter = 10, 0, "fade"
        else:
            duration = 380 + ((page + index) % 3) * 32
            preset_id, subtype, effect_filter = 10, 0, "fade"
        ids = _effect_ids(next_id)
        next_id += 5
        make_effect(
            seq_child,
            int(shape["id"]),
            delay,
            duration,
            ids,
            preset_id,
            subtype,
            effect_filter,
        )

    prev_cond = etree.SubElement(seq, tag(P, "prevCondLst"))
    prev = etree.SubElement(prev_cond, tag(P, "cond"), evt="onPrev", delay="0")
    etree.SubElement(etree.SubElement(prev, tag(P, "tgtEl")), tag(P, "sldTgt"))
    next_cond = etree.SubElement(seq, tag(P, "nextCondLst"))
    nxt = etree.SubElement(next_cond, tag(P, "cond"), evt="onNext", delay="0")
    etree.SubElement(etree.SubElement(nxt, tag(P, "tgtEl")), tag(P, "sldTgt"))
    return timing


def process_slide(xml: bytes, page: int) -> tuple[bytes, int]:
    root = etree.fromstring(xml)
    sp_tree = root.find(".//" + tag(P, "spTree"))
    if sp_tree is None:
        return xml, 0
    recs = collect_shapes(sp_tree)
    blocks = semantic_blocks(recs, page)
    targets = wrap_blocks(root, blocks)

    old_timing = root.find(tag(P, "timing"))
    if old_timing is not None:
        root.remove(old_timing)
    old_transition = root.find(tag(P, "transition"))
    if old_transition is not None:
        root.remove(old_transition)

    transition = etree.Element(tag(P, "transition"), spd="med")
    etree.SubElement(transition, tag(P, "fade"))
    timing = timing_xml(targets, page)

    clr = root.find(tag(P, "clrMapOvr"))
    if clr is not None:
        clr.addnext(transition)
        transition.addnext(timing)
    else:
        root.append(transition)
        root.append(timing)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), len(targets)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("src", type=Path)
    parser.add_argument("dst", type=Path)
    args = parser.parse_args()

    total = 0
    distribution: dict[int, int] = {}
    with zipfile.ZipFile(args.src) as zin, zipfile.ZipFile(
        args.dst, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            match = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", info.filename)
            if match:
                data, count = process_slide(data, int(match.group(1)))
                total += count
                distribution[count] = distribution.get(count, 0) + 1
            zout.writestr(info, data)
    print(f"units={total} distribution={dict(sorted(distribution.items()))} -> {args.dst}")


if __name__ == "__main__":
    main()
