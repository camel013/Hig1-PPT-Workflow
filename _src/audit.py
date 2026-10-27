# -*- coding: utf-8 -*-
"""Role-aware font audit for the Zhiwen deck."""
from __future__ import annotations

import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

P = Path(sys.argv[1] if len(sys.argv) > 1 else
         r"projects\zhiwen_teaching_ppt169_20261024\_deck_u1.pptx")
NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PNS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

violations: list[dict] = []
role_counts: Counter[tuple[str, str]] = Counter()
run_counts: Counter[tuple[str, str, str]] = Counter()


def text_of(node):
    return "".join(x.text or "" for x in node.iter(NS + "t"))


with zipfile.ZipFile(P) as z:
    slides = sorted(
        [n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)],
        key=lambda x: int(re.findall(r"\d+", x)[0]),
    )
    for sn in slides:
        root = ET.fromstring(z.read(sn))
        for sp in root.iter(PNS + "sp"):
            shape_text = text_of(sp)
            if not shape_text.strip():
                continue
            name_el = sp.find(f"{PNS}nvSpPr/{PNS}cNvPr")
            name = (name_el.get("name") if name_el is not None else "") or ""
            role = "body" if name.startswith("body_") else "heading"
            role_counts[(role, "all")] += 1
            if role == "heading" and (
                (len(shape_text) > 40)
                or (len(shape_text) > 16 and any(ch in shape_text for ch in "。！？"))
            ):
                violations.append({
                    "slide": sn,
                    "kind": "long_heading_or_body_in_fangsong",
                    "text": shape_text[:90],
                    "shape": name,
                })

            for r in sp.iter(NS + "r"):
                run_text = text_of(r)
                if not run_text.strip():
                    continue
                rpr = r.find(NS + "rPr")
                fonts = {}
                if rpr is not None:
                    for tag in ("latin", "ea", "cs"):
                        el = rpr.find(NS + tag)
                        fonts[tag] = el.get("typeface") if el is not None else None
                has_cjk = any("\u4e00" <= ch <= "\u9fff" for ch in run_text)
                has_alnum = any(ch.isascii() and ch.isalnum() for ch in run_text)
                run_counts[(role, fonts.get("latin", ""), fonts.get("ea", ""))] += 1

                if has_alnum and (fonts.get("latin") != "Times New Roman" or fonts.get("cs") != "Times New Roman"):
                    violations.append({"slide": sn, "kind": "latin_not_tnr", "text": run_text[:60], "fonts": fonts})
                if has_cjk and fonts.get("ea") not in ("仿宋", "华文中宋"):
                    violations.append({"slide": sn, "kind": "cjk_font_invalid", "text": run_text[:60], "fonts": fonts})
                if role == "body" and has_cjk and fonts.get("ea") != "华文中宋":
                    violations.append({"slide": sn, "kind": "body_not_stzhongsong", "text": run_text[:60], "fonts": fonts})
                if any(f in ("楷体", "KaiTi", "楷体_GB2312") for f in fonts.values() if f):
                    violations.append({"slide": sn, "kind": "kaiti_forbidden", "text": run_text[:60], "fonts": fonts})

out = {
    "file": str(P),
    "slides": len(slides),
    "role_counts": {f"{k[0]}|{k[1]}": v for k, v in role_counts.items()},
    "run_counts": {f"{k[0]}|{k[1]}|{k[2]}": v for k, v in run_counts.items()},
    "violations": violations,
    "n_viol": len(violations),
}
out_path = Path(r"_src\fontaudit.json")
out_path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
print(json.dumps({k: out[k] for k in ("file", "slides", "role_counts", "n_viol")}, ensure_ascii=False, indent=1))
