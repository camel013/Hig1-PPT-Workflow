# -*- coding: utf-8 -*-
"""Make the theme fallback match the deck's declared typography."""
from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path


def normalize_theme(xml: str) -> str:
    for family in ("majorFont", "minorFont"):
        pattern = rf"(<a:{family}>)(.*?)(</a:{family}>)"

        def fix(match: re.Match[str]) -> str:
            body = match.group(2)
            body = body.replace('<a:latin typeface="Calibri"/>',
                                '<a:latin typeface="Times New Roman"/>')
            body = body.replace('<a:ea typeface=""/>',
                                '<a:ea typeface="华文中宋"/>')
            body = body.replace('<a:cs typeface=""/>',
                                '<a:cs typeface="Times New Roman"/>')
            body = body.replace('script="Hans" typeface="宋体"',
                                'script="Hans" typeface="华文中宋"')
            return match.group(1) + body + match.group(3)

        xml = re.sub(pattern, fix, xml, flags=re.S)
    return xml


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("src", type=Path)
    parser.add_argument("dst", type=Path)
    args = parser.parse_args()

    with zipfile.ZipFile(args.src) as zin, zipfile.ZipFile(
        args.dst, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.startswith("ppt/theme/") and info.filename.endswith(".xml"):
                data = normalize_theme(data.decode("utf-8")).encode("utf-8")
            zout.writestr(info, data)
    print(args.dst)


if __name__ == "__main__":
    main()
