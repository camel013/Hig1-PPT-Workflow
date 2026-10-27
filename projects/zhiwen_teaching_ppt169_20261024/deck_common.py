# -*- coding: utf-8 -*-
"""Reusable page systems for units 2-4 and the afterword."""
import deck_lib as D
from deck_lib import (PAPER, PAPER_WARM, WHITE, INK, MUTED, RULE, CORAL,
                      CORAL_DARK, CORAL_LIGHT, INDIGO, INDIGO_DEEP, NAVY, TEAL,
                      TEAL_DARK, GOLD, GOLD_LIGHT, FANG, SONG, TNR, PP_ALIGN)
from state import nxt


def _source(slide, x, y, w, text, dark=False, align=PP_ALIGN.LEFT):
    D.text(slide, x, y, w, text, 9.5, FANG, GOLD_LIGHT if dark else MUTED,
           1.35, align, spacing=0.8)


def _arc_mark(slide, variant, fill, edge=None):
    specs = {
        0: (1328, 748, 338, 178, 286, 23000),
        1: (-62, 762, 330, 300, 54, 22000),
        2: (1320, -78, 348, 72, 184, 21000),
        3: (1262, 770, 274, 194, 288, 24000),
    }
    cx, cy, r, a0, a1, thick = specs[variant % 4]
    D.arc_block(slide, cx, cy, r + 20, a0, a1, thick,
                D.blend(fill, PAPER if fill != PAPER else WHITE, 0.62))
    D.arc_block(slide, cx, cy, r, a0, a1, thick, fill)
    if edge:
        D.arc(slide, cx, cy, r + 20, r + 20, a0, a1, edge, 5.0, 60)


def intro_page(kicker, no, text_title, author, headline, lead, quote, source,
               accent=INDIGO, variant=0, bg=PAPER, dark=False, tags=None,
               footer_color=None):
    """A work-opening page: author, context, and a substantial quotation."""
    p = nxt()
    s = D.new_slide(bg)
    D.page_title(s, p, kicker, headline, dark=dark)
    fg = WHITE if dark else INK
    muted = GOLD_LIGHT if dark else MUTED
    _arc_mark(s, variant, accent if not dark else TEAL_DARK)

    D.rect(s, 64, 150, 126, 224, fill=accent)
    D.rect(s, 190, 150, 12, 224, fill=D.blend(accent, bg, 0.58))
    D.text(s, 78, 174, 105, no, 78, TNR, WHITE, 1.0)
    D.text(s, 218, 154, 510, text_title, 31, FANG, fg, 1.3)
    D.text(s, 218, 212, 480, author, 12.5, FANG, muted, 1.4, spacing=0.8)
    D.hline(s, 218, 718, 246, accent if not dark else GOLD_LIGHT, 4.0)
    D.text(s, 64, 414, 616, lead, 15.5, SONG, fg, 1.75, max_lines=5,
           label=text_title)

    qx, qy, qw, qh = 790, 150, 426, 414
    D.rect(s, qx, qy, qw, qh, fill=NAVY if not dark else PAPER)
    D.rect(s, qx, qy, 12, qh, fill=GOLD if accent != GOLD else CORAL)
    qfg = WHITE if not dark else INK
    lines = [line for line in quote.split("\n") if line.strip()]
    yy = qy + 36
    for line in lines[:6]:
        yy = D.text(s, qx + 34, yy, qw - 62, line, 21, SONG, qfg, 1.62,
                    max_lines=2, label=text_title + "引文")
        yy += 8
    D.hline(s, qx + 34, qx + qw - 32, qy + 334, GOLD_LIGHT if not dark else RULE, 3.0)
    _source(s, qx + 34, qy + 352, qw - 62, source, dark=not dark)

    if tags:
        x = 64
        for t in tags:
            w = D.measure(t, 9.5, SONG) + 24
            D.rect(s, x, 586, w, 25, fill=accent)
            D.text(s, x + 12, 592, w - 18, t, 9.5, FANG, WHITE, 1.1)
            x += w + 10
    D.text(s, 64, 642, 650, "读法：先辨认他借来的体裁，再找他把哪一处现实写坏了。",
           10.5, SONG, muted, 1.4)
    D.footer(s, p, dark=dark, color=footer_color)
    return s


def read_page(kicker, title, lead, quote, source, points, accent=INDIGO,
              bg=PAPER, dark=False, side="left", note=None, variant=0):
    """Quotation-led close reading with three analytical stops."""
    p = nxt()
    s = D.new_slide(bg)
    D.page_title(s, p, kicker, title, dark=dark)
    fg = WHITE if dark else INK
    muted = GOLD_LIGHT if dark else MUTED
    D.text(s, 64, 154, 1090, lead, 15.5, SONG, fg, 1.72,
           max_lines=2, label=title)
    _arc_mark(s, variant, accent, GOLD_LIGHT if dark else None)

    band_w = 594
    bx = 64 if side == "left" else 622
    band_fill = NAVY if not dark else PAPER
    D.rect(s, bx, 236, band_w, 312, fill=band_fill)
    D.rect(s, bx + band_w, 236, 12, 312, fill=D.blend(accent, bg, 0.58))
    D.rect(s, bx, 236, 12, 312, fill=accent)
    qfg = WHITE if not dark else INK
    yy = 270
    for line in [x for x in quote.split("\n") if x.strip()][:5]:
        yy = D.text(s, bx + 38, yy, band_w - 70, line, 22, SONG, qfg, 1.64,
                    max_lines=2, label=title + "引文")
        yy += 4
    D.hline(s, bx + 38, bx + band_w - 34, 486, GOLD_LIGHT if not dark else RULE, 3.0)
    _source(s, bx + 38, 504, band_w - 70, source, dark=not dark)

    px = 734 if side == "left" else 64
    for i, (head, body) in enumerate(points[:3]):
        y = 234 + i * 108
        D.text(s, px, y, 38, f"0{i+1}", 16, TNR, accent, 1.0)
        D.text(s, px + 48, y - 3, 430, head, 18, FANG, fg, 1.2)
        D.text(s, px + 48, y + 31, 430, body, 12.3, SONG, fg, 1.62,
               max_lines=3, label=head)
        D.hline(s, px, px + 480, y + 88, RULE if not dark else TEAL_DARK, 3.0)
    if note:
        D.text(s, 64, 620, 750, note, 10.5, SONG, muted, 1.4)
    D.footer(s, p, dark=dark)
    return s


def grid_quotes(kicker, title, lead, items, accent=INDIGO, bg=PAPER_WARM,
                dark=False, note=None):
    """2x2 quotation grid. Each item is (heading, quote, comment, source)."""
    p = nxt()
    s = D.new_slide(bg)
    D.page_title(s, p, kicker, title, dark=dark)
    fg = WHITE if dark else INK
    muted = GOLD_LIGHT if dark else MUTED
    D.text(s, 64, 152, 1090, lead, 15, SONG, fg, 1.65, max_lines=2, label=title)
    for i, (head, quote, comment, source) in enumerate(items[:4]):
        x = 64 + (i % 2) * 584
        y = 238 + (i // 2) * 218
        D.rect(s, x, y, 548, 190, fill=WHITE if not dark else NAVY)
        D.rect(s, x, y, 7, 190, fill=accent)
        D.text(s, x + 22, y + 16, 220, head, 15.5, FANG, accent, 1.2)
        D.text(s, x + 22, y + 48, 478, quote, 14.2, SONG, fg, 1.55,
               max_lines=3, label=head)
        D.text(s, x + 22, y + 121, 478, comment, 11.2, SONG, muted, 1.42,
               max_lines=2, label=head)
        _source(s, x + 300, y + 164, 220, source, dark=dark, align=PP_ALIGN.RIGHT)
    if note:
        D.text(s, 64, 660, 1110, note, 9.8, SONG, muted, 1.35)
    D.footer(s, p, dark=dark)
    return s


def split_page(kicker, title, lead, left_head, left_body, right_head, right_body,
               quote, source, accent=INDIGO, bg=PAPER, dark=False, note=None,
               variant=0):
    """Two opposed arguments with a common quotation at the bottom."""
    p = nxt()
    s = D.new_slide(bg)
    D.page_title(s, p, kicker, title, dark=dark)
    fg = WHITE if dark else INK
    muted = GOLD_LIGHT if dark else MUTED
    D.text(s, 64, 152, 1090, lead, 15.5, SONG, fg, 1.68, max_lines=2, label=title)
    _arc_mark(s, variant, accent)
    for i, (head, body) in enumerate(((left_head, left_body), (right_head, right_body))):
        x = 64 + i * 568
        D.text(s, x, 242, 490, head, 22, FANG, accent, 1.2)
        D.hline(s, x, x + 500, 280, accent, 5.0)
        D.text(s, x, 300, 500, body, 13.2, SONG, fg, 1.65, max_lines=5, label=head)
    D.rect(s, 64, 472, 1152, 124, fill=NAVY if not dark else PAPER)
    D.text(s, 92, 500, 1048, quote, 20, SONG, WHITE if not dark else INK, 1.55,
           max_lines=2, label=title + "引文")
    _source(s, 94, 570, 640, source, dark=not dark)
    if note:
        D.text(s, 64, 618, 1100, note, 10.5, SONG, muted, 1.42)
    D.footer(s, p, dark=dark)
    return s


def timeline_page(kicker, title, lead, events, accent=INDIGO, bg=PAPER,
                  dark=False, note=None):
    p = nxt()
    s = D.new_slide(bg)
    D.page_title(s, p, kicker, title, dark=dark)
    fg = WHITE if dark else INK
    muted = GOLD_LIGHT if dark else MUTED
    D.text(s, 64, 150, 1090, lead, 15, SONG, fg, 1.65, max_lines=2, label=title)
    yline = 330
    D.hline(s, 82, 1202, yline, accent, 6.0)
    n = len(events)
    for i, (year, head, body) in enumerate(events):
        x = 84 + i * (1116 / max(1, n - 1))
        up = i % 2 == 0
        D.rect(s, x - 6, yline - 8, 14, 22, fill=accent)
        ty = 248 if up else 362
        D.text(s, x - 34, ty, 150, year, 16, TNR, accent, 1.1)
        D.text(s, x - 34, ty + 28, 172, head, 13.5, FANG, fg, 1.25)
        D.text(s, x - 34, ty + 58, 172, body, 10.2, SONG, fg, 1.4,
               max_lines=3, label=head)
    if note:
        D.text(s, 64, 606, 1110, note, 10.5, SONG, muted, 1.4)
    D.footer(s, p, dark=dark)
    return s


def statement_page(kicker, title, lead, quote, source, points, accent=CORAL,
                   bg=NAVY, dark=True, note=None, variant=2):
    p = nxt()
    s = D.new_slide(bg)
    D.page_title(s, p, kicker, title, dark=dark)
    fg = WHITE if dark else INK
    muted = GOLD_LIGHT if dark else MUTED
    D.text(s, 64, 152, 1090, lead, 15.5, SONG, fg, 1.68, max_lines=2, label=title)
    _arc_mark(s, variant, accent)
    D.text(s, 64, 230, 740, quote, 30, SONG, WHITE if dark else INK, 1.58,
           max_lines=5, label=title + "引文")
    _source(s, 66, 444, 620, source, dark=dark)
    D.rect(s, 880, 224, 328, 330, fill=accent)
    D.rect(s, 864, 224, 16, 330, fill=D.blend(accent, bg, 0.58))
    for i, (head, body) in enumerate(points[:3]):
        y = 252 + i * 90
        D.text(s, 908, y, 270, head, 15, FANG, WHITE, 1.2)
        D.text(s, 908, y + 28, 270, body, 10.8, SONG, WHITE, 1.48,
               max_lines=3, label=head)
    if note:
        D.text(s, 64, 590, 770, note, 10.5, SONG, muted, 1.42)
    D.footer(s, p, dark=dark)
    return s


def unit_overview(kicker, title, lead, items, accent=INDIGO, bg=PAPER_WARM,
                  note=None):
    p = nxt()
    s = D.new_slide(bg)
    D.page_title(s, p, kicker, title)
    D.text(s, 64, 150, 1100, lead, 15.5, SONG, INK, 1.7, max_lines=2, label=title)
    for i, (no, head, body, genre) in enumerate(items[:8]):
        x = 64 + (i % 4) * 292
        y = 240 + (i // 4) * 194
        D.text(s, x, y, 74, no, 17, TNR, accent, 1.0)
        D.hline(s, x, x + 258, y + 28, accent, 5.0)
        D.text(s, x, y + 48, 256, head, 16.5, FANG, INK, 1.25,
               max_lines=2, label=head)
        D.text(s, x, y + 96, 256, body, 10.6, SONG, INK, 1.45,
               max_lines=3, label=head)
        D.text(s, x, y + 164, 250, genre, 9.2, FANG, accent, 1.2)
    if note:
        D.text(s, 64, 646, 1110, note, 10.2, SONG, MUTED, 1.38)
    D.footer(s, p)
    return s


def close_page(kicker, title, quote, source, body, accent=INDIGO, bg=NAVY,
               note=None, signature=None):
    p = nxt()
    s = D.new_slide(bg)
    D.meta(s, p, kicker, dark=True)
    D.text(s, 64, 90, 980, title, 40, FANG, WHITE, 1.25)
    D.hline(s, 64, 1216, 158, accent, 6.0)
    D.text(s, 64, 210, 910, quote, 36, SONG, WHITE, 1.55, max_lines=6,
           label=title + "引文")
    _source(s, 66, 520, 700, source, dark=True)
    D.text(s, 64, 578, 900, body, 13.5, SONG, GOLD_LIGHT, 1.6,
           max_lines=2, label=title)
    if note:
        D.text(s, 990, 520, 220, note, 10.5, SONG, WHITE, 1.45)
    if signature:
        D.text(s, 980, 620, 230, signature, 10.5, FANG, GOLD_LIGHT, 1.3,
               PP_ALIGN.RIGHT)
    D.footer(s, p, dark=True)
    return s
