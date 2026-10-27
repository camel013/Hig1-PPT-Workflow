# xigai_story

- Canvas: 16:9, 1280×720
- Slides: 12, native editable shapes
- Palette: deep red, wine purple, paper white
- Type: 楷体 for section headings; 华文中宋 for all other text, numerals and Latin
- Source: `sources/故事与引用.md`
- Spec: `设计规范与内容大纲.md`

## Delivery

- `留沪之前_教育科技人才战略_12页.pptx`
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

