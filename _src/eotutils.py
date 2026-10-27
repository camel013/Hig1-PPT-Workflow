# -*- coding: utf-8 -*-
"""Minimal TTF -> EOT wrapper compatible with PowerPoint fntdata parts."""
from __future__ import annotations

import io
import struct
from pathlib import Path

from fontTools.ttLib import TTFont


def _english_name(font: TTFont, name_id: int) -> bytes:
    for record in font["name"].names:
        if (
            record.nameID == name_id
            and record.platformID == 3
            and record.platEncID == 1
            and record.langID == 0x0409
        ):
            return record.string
    record = font["name"].getName(name_id, 3, 1, 0x0409)
    if record is None:
        raise ValueError(f"font has no English name ID {name_id}")
    return record.string


def _append_name(out: bytearray, value: bytes) -> None:
    if len(value) % 2:
        raise ValueError("name strings must be UTF-16BE")
    # EOT stores its UTF-16 name strings little-endian.
    value = value.decode("utf-16-be").encode("utf-16-le")
    out += struct.pack("<H", len(value))
    out += value
    out += b"\x00\x00"


def make_eot(ttf_path: str | Path) -> bytes:
    ttf = Path(ttf_path).read_bytes()
    font = TTFont(io.BytesIO(ttf))
    os2 = font["OS/2"]
    head = font["head"]

    header = bytearray()
    header += struct.pack("<IIII", 0, len(ttf), 0x00020001, 0)
    panose = os2.panose
    header += bytes(
        [
            panose.bFamilyType,
            panose.bSerifStyle,
            panose.bWeight,
            panose.bProportion,
            panose.bContrast,
            panose.bStrokeVariation,
            panose.bArmStyle,
            panose.bLetterForm,
            panose.bMidline,
            panose.bXHeight,
        ]
    )
    header += bytes([1, int(bool(os2.fsSelection & 1))])
    header += struct.pack("<I", os2.usWeightClass)
    header += struct.pack("<H", 0)  # fsType: installable, per ttf2eot
    header += struct.pack("<H", 0x504C)
    header += struct.pack(
        "<4I",
        os2.ulUnicodeRange1,
        os2.ulUnicodeRange2,
        os2.ulUnicodeRange3,
        os2.ulUnicodeRange4,
    )
    header += struct.pack("<2I", os2.ulCodePageRange1, os2.ulCodePageRange2)
    header += struct.pack("<I", head.checkSumAdjustment)
    header += struct.pack("<4I", 0, 0, 0, 0)
    header += struct.pack("<H", 0)

    _append_name(header, _english_name(font, 1))
    _append_name(header, _english_name(font, 2))
    _append_name(header, _english_name(font, 5))
    _append_name(header, _english_name(font, 4))
    header += b"\x00\x00"

    struct.pack_into("<I", header, 0, len(header) + len(ttf))
    return bytes(header) + ttf


if __name__ == "__main__":
    import sys

    src, dst = sys.argv[1], sys.argv[2]
    Path(dst).write_bytes(make_eot(src))
    print(dst, Path(dst).stat().st_size)
