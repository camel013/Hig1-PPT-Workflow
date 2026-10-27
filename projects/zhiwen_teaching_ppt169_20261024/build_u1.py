import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).parent))
import deck_lib as D
import deck_front, deck_unit1
deck_front.build()
deck_unit1.unit1()
D.prs.save(pathlib.Path(__file__).parent / "_deck_u1.pptx")
print("slides", len(D.prs.slides))
