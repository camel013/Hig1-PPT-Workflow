# pharma_career

- Canvas: 16:9, 1280×720
- Slides: 12, native editable shapes
- Type: 楷体 for section headings; 华文中宋 for all other text, numerals and Latin
- Colors: paper white, ink black, neutral greys only
- Source: `sources/报告正文.md`
- Spec: `设计规范与内容大纲.md`

## Delivery

- `祝炜程_从药学基础到制药现场_12页.pptx`
- `preview/contact_sheet.png`
- `pp_check/slide-1.png` ... `slide-12.png`

## Rebuild

```powershell
$py = 'C:\Users\camel\.codex\skills\ppt-master-skill\.venv\Scripts\python.exe'
& $py build_deck.py
& $py embed_fonts.py
```

`embed_fonts.py` copies the audited full-font EOT payloads for 楷体 and
华文中宋 from the legacy deck. It does not use PowerPoint COM.

