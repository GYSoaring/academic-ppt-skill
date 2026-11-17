# Editable PowerPoint Export

The skill produces two PowerPoint files from the same `deck.json`:

- `defense_deck.pptx`: native editable version. Text, card shapes, tables, timelines, references, and formulas are PowerPoint objects.
- `defense_deck_raster.pptx`: optional preview version. Every slide is a high-fidelity PNG. Use it only for visual comparison or delivery when native fonts are unavailable.

The native version deliberately reuses the preview renderer's theme coordinates and type scale. It does not redesign the page while exporting.

Use `--profile document` for dense handout decks, `--profile projection` or `lecture` for teaching, and `--profile defense` for a defense-room deck with 24 pt body text and 32–40 pt headings. Projected profiles raise the physical PowerPoint type scale and intentionally fail on overflow instead of silently shrinking text. Use `scripts/audit_pptx.py` to check the exported PPTX's actual text runs; a good body median does not excuse a large fraction of sub-18 pt labels.

The theme's `text.body_bold_min_px` / `body_bold_max_px` bound semantic body bolding to readable body sizes. Footnotes, captions, reference metadata, and micro-labels remain regular so the slide keeps a hierarchy instead of becoming one uniform bold wall. See [defense-readability.md](defense-readability.md) for the audit thresholds.

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

## Native Geometry Contract

The preview and editable paths share the same `(left, top, width, height)` geometry. Keep these rules when changing either renderer:

- Every geometry box uses `(left, top, width, height)`. Do not pass a parent's width where a child width is expected; this produced a negative-width condition note that PIL masked but PowerPoint exposed.
- Index tokens (`01`, `02`, chapter numbers) must use `wrap_explicit=False` in the native exporter. They are navigation tokens, not prose.
- Header/module indices and their headings share one family, measured size, top edge, and baseline. A defense module index is measured in the bold face, at least 60 px wide, and `wrap_explicit=False`; a defense module header reserves at least 76 px.
- Atomic labels — formula labels, metric values, figure badges, captions, footnotes, and source lines — also use `wrap_explicit=False`. Size their box from the measured face plus 16–24 px of padding; do not reuse a fixed 34 px strip for 38 px text.
- Bold must not change font family. `run.font.bold = True` is the only change; the family stays the body/title face.
- All visible text boxes must have positive width and height and remain inside the slide. `validate_pptx.py` treats non-positive text geometry as an error and outside-slide text as a warning.
- Sequence diagram nodes compute title offsets relative to `node_top`; `body_height` is `node_height - relative_title_offset - title_height - padding`. Never mix absolute page coordinates with node-local height.
- A concept node body must stay inside its named node shape. `validate_pptx.py` pairs `diagram_sequence_node_*` with `diagram_node_body` and fails if either edge escapes.
- A metrics condition note uses `side_box[2]` as the available width, not as an x-coordinate. The note, condition grid, and evidence image share the same lower band.
- A defense metrics condition grid is 2×2 with four label/value pairs. Each pair gets a full row band (label band above value band); do not put both into a 58 px row. Omit separator rules when they would cross descenders, and cap the slide at four conditions.
- Paired condition text is named `condition_label_N` and `condition_value_N`; the label and value share a left edge, and the label's bottom edge must stay above the value's top edge. The validator uses these names, so do not rename one side without the other.
- Defense figure cards draw `figure_label`, `figure_badge`, and `figure_caption` as native text inside the card. Label and badge share one header line; the caption sits in a measured bottom band and never wraps.
- The figure card itself is named `figure_card`. The validator checks that all three chrome shapes are single-line and fully contained by that card; if you add a new chrome shape, give it a semantic name and extend the same invariant.
- Closing cards align title and body on the same left edge; the title uses the block-title scale (`52+ px` under `defense`) and the body uses the small-text scale (`39 px`). Keep the icon at the trailing edge.

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
- text boxes have positive geometry and remain on the slide
- short numeric indices are not line-wrapped
- concept-diagram bodies remain inside their sequence nodes
- condition labels stay above their values, share the value's left edge, and use one line under `defense`
- defense figure-card labels, badges, and captions are single-line and remain inside `figure_card`

`pptx_report.json` records per-slide editable shape counts, picture counts, native tables, native charts, and formula conversion status.

## Known Boundaries

- PowerPoint cannot reflow the same way as a browser or KaTeX. Small line-break differences are expected; use the preview render for visual review and the editable PPTX for delivery.
- Native charts keep the source values and encodings editable, but PowerPoint may adjust plot padding and tick spacing by a few pixels. Keep the chart card frame, labels, and palette the same as the preview; do not replace the chart with a screenshot.
- A paper texture is still a picture. This is intentional: it keeps the background stable without rasterizing the content.
- Do not solve overflow by returning to a full-page raster. Fix the content, type scale, or theme coordinates instead.
