# Editable PowerPoint Export

The skill produces two PowerPoint files from the same `deck.json`:

- `defense_deck.pptx`: native editable version. Text, card shapes, tables, timelines, references, and formulas are PowerPoint objects.
- `defense_deck_raster.pptx`: optional preview version. Every slide is a high-fidelity PNG. Use it only for visual comparison or delivery when native fonts are unavailable.

The native version deliberately reuses the preview renderer's theme coordinates and type scale. It does not redesign the page while exporting.

Use `--profile document` for dense handout decks. Use `--profile projection` when the deck will be projected; it raises the physical PowerPoint type scale and intentionally fails on overflow instead of silently shrinking text.

## What Is Native

- Cover title, kicker, subtitle, presenter, advisor, date, institution, and bottom line.
- Section header, eyebrow, slash marks, fixed rules, mini-directory, and footer.
- Lead claim, takeaway band, footnote, cards, card headers, output strips, and bullets.
- Comparison tables.
- Bar, histogram, column, line, and curve charts. The chart is a native PowerPoint chart object with an editable data sheet; the chart label, badge, axis labels, caption, and annotation are native text.
- Pipeline nodes, timeline cards, hub cards, references, and closing blocks.
- Semantic `concept` diagrams as native shapes and text boxes: `sequence`, `branch`, `stack`, `ruler`, `matrix`, and `cycle`.
- Formula-led pages: the equation as OMML, symbol definitions as native text, and condition/boundary panels as native shapes.
- Real title placeholders, semantic alt text, decorative markers, header-row table semantics, and document core properties.
- Speaker notes.

Only the paper texture and institutional lockup are pictures. The background picture contains no readable content and can be replaced without deleting the slide text.

Header chapter numbers, outline numbers, and their adjacent titles use the measured title size on both the preview and native PPTX paths. A long title may shrink, but its number shrinks with it; do not leave the number at a fixed decorative size.

## Font Mapping

Each theme may define an optional `pptx_fonts` object:

```json
{
  "pptx_fonts": {
    "title": "Noto Sans SC",
    "body": "Noto Sans SC",
    "display": "Source Han Serif CN SemiBold",
    "latin": "Arial",
    "formula": "Cambria Math"
  }
}
```

When the field is omitted, the exporter derives a family name from the first available theme font. This is a best-effort mapping, because a TTF/OTF path is not the same thing as a PowerPoint family name.

If a user opens the deck on a machine without the configured CJK family, PowerPoint substitutes a fallback and line breaks can shift. Keep an editable copy of the source `deck.json`; do not flatten the deck to solve a font substitution problem.

Recommended public fallback:

- Windows: `Microsoft YaHei`
- macOS: `PingFang SC`
- Linux/LibreOffice: install Noto Sans CJK and Noto Serif CJK

## Formula Pipeline

For each `formula_latex`:

1. `latex2mathml` converts LaTeX to MathML.
2. Office's `MML2OMML.XSL` converts MathML to OMML.
3. The OMML node is inserted into a native PowerPoint text frame.

The formula remains editable in PowerPoint's equation editor. Keep `formula` as the plain fallback for environments without LaTeX/OMML support.

If `latex2mathml` or the Office XSL is unavailable, the exporter writes the plain formula with Cambria Math. It reports this in `pptx_report.json` and `pptx_validation_report.json`.

For a `formula` layout, the validator counts the page equation as one expected formula. Symbol rows are text, not separate equations; use real subscript characters such as `tₑ` or `uₑ` in the visible text. Do not write `t_e` in the slide.

## Accessibility And Render QA

The exporter uses a real title placeholder on every slide. Background textures and non-text chrome are marked decorative; content pictures use `image_role`/`image_caption` as semantic alt text; comparison tables set `first_row = True`; core properties are populated from `meta`.

Run these gates after exporting:

```bash
python scripts/validate_pptx.py --pptx defense_deck.pptx --deck deck.json --out output
python scripts/check_accessibility.py --pptx defense_deck.pptx --out output
python scripts/render_pptx_preview.py --pptx defense_deck.pptx --out output/pptx_preview --compare-dir output/rendered --strict
```

`render_pptx_preview.py` prefers PowerPoint COM on Windows and falls back to LibreOffice plus PyMuPDF on other systems. A skipped render is reported explicitly; it must not be described as a successful visual QA pass.

## Validation

Run:

```bash
python scripts/export_editable_pptx.py --deck deck.json --theme theme.json --out output
python scripts/validate_pptx.py --pptx output/defense_deck.pptx --deck deck.json --out output
```

`validate_pptx.py` checks:

- slide count equals the deck count
- enough non-empty editable text shapes per slide
- speaker notes are present
- comparison layouts contain native tables
- slides with `chart` data contain a native editable PowerPoint chart
- formulas are OMML rather than a single raster image
- a `formula` layout has its page equation as native OMML
- a `concept` layout has editable diagram text and shapes, not a single raster diagram
- every slide has a non-empty title placeholder
- content pictures have semantic alt text or are marked decorative
- comparison tables have header-row semantics
- the deck is not just one full-page picture per slide

`pptx_report.json` records per-slide editable shape counts, picture counts, native tables, native charts, and formula conversion status.

## Known Boundaries

- PowerPoint cannot reflow the same way as a browser or KaTeX. Small line-break differences are expected; use the preview render for visual review and the editable PPTX for delivery.
- Native charts keep the source values and encodings editable, but PowerPoint may adjust plot padding and tick spacing by a few pixels. Keep the chart card frame, labels, and palette the same as the preview; do not replace the chart with a screenshot.
- A paper texture is still a picture. This is intentional: it keeps the background stable without rasterizing the content.
- Do not solve overflow by returning to a full-page raster. Fix the content, type scale, or theme coordinates instead.
