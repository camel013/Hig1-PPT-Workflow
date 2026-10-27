# -*- coding: utf-8 -*-
"""Build the complete 95-page native PowerPoint deck."""
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).parent))

import deck_lib as D
import deck_front
import deck_unit1
import deck_unit2
import deck_unit3
import deck_unit4
import deck_end


def main():
    deck_front.build()
    deck_unit1.unit1()
    deck_unit2.unit2()
    deck_unit3.unit3()
    deck_unit4.unit4()
    deck_end.ending()
    D.add_color_chords()
    out = pathlib.Path(__file__).parent / "蛭文_教学读本_95页.pptx"
    D.prs.save(out)
    print("slides", len(D.prs.slides), "->", out.name)


if __name__ == "__main__":
    main()
