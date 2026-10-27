# duoyuhua

- Canvas format: ppt169
- Created: 20261021

## Directories

- `svg_output/`: raw SVG output
- `svg_final/`: finalized SVG output
- `images/`: presentation assets
- `notes/`: speaker notes
- `templates/`: project templates
- `sources/`: source materials and normalized markdown

## Delivery

- `多余的话_瑞士国际主义版式.pptx`: native editable deck, 12 slides, full fonts embedded.
- `preview/contact_sheet.png`: deck overview.
- `pp_check/s01.png` ... `s12.png`: full-size slide previews.

The phone PDF is intentionally not delivered for this revision.

## Rebuild

```powershell
$py = 'C:\Users\camel\.codex\skills\ppt-master-skill\.venv\Scripts\python.exe'
& $py build_pptx.py
& $py merge_legacy.py _deck_build.pptx '多余的话_瑞士国际主义版式.pptx' _deck_final.pptx
& $py verify_fonts.py _deck_final.pptx
```

`merge_legacy.py` copies the audited full-font payload and timings from the
current final deck; it does not use PowerPoint COM. Render previews with the
`@oai/artifact-tool` `render_presentation.mjs` helper. Do not rebuild a PDF
unless the user asks for one.
