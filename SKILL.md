---
name: academic-ppt-skill
description: Create academic defense, proposal, midterm, course, competition, and literature-review decks with fixed institutional branding, structured slides, figure-aware speaker notes, and a natural defense script. Use when visual consistency and claim-evidence discipline matter more than one-off slide decoration.
metadata:
  short-description: 学术汇报 PPT 与答辩稿生成
---

# Academic PPT Skill（学术汇报 PPT）

Build the deck as structured content plus a deterministic rendering layer. Do not ask an image model to render exact Chinese titles, school names, seals, page numbers, or fixed headers.

## Outcome

Produce a self-contained output directory containing:

- `deck.json`: sections, slides, figure roles, claims, caveats, and speech text.
- `rendered/slide_01.png` and subsequent pages at the theme canvas size.
- `defense_deck.pptx`: native editable PowerPoint slides with editable text, shapes, tables, charts, OMML formulas, and speaker notes.
- `defense_deck_raster.pptx`: optional high-fidelity full-bleed preview PPTX, used only when native PowerPoint typography is not available.
- `contact_sheet.jpg`: visual overview for review.
- `speaker_notes.md`: one timed note block per slide.
- `defense_script.md`: continuous, natural spoken script.
- `render_report.json`: overflow, missing-asset, and layout warnings.
- `accessibility_report.json`: title, alt, table-header, reading-order, and core-property checks.
- `pptx_preview/` and `pptx_preview/pptx_render_report.json`: real PowerPoint/LibreOffice render and preview comparison.
- `agent_brief.md`: portable instructions for Doubao or another model that cannot run these scripts.
- `pptx_readability.json` / `pptx_readability.md`: optional physical-font and density audit of an arbitrary PPTX.

## Workflow

1. Read the user's paper, outline, images, school identity assets, presenter, advisor, date, and time limit. If the user supplies PDF, DOCX, PPTX, Markdown, or a folder of figures, run `scripts/extract_source.py` first and review `source.md`, `source_manifest.json`, and `figures.json` before planning slides. Prefer the user's own paper figures over external images. If the source has no usable figure, run `scripts/find_figures.py` on the topic (or known arXiv ids) and `scripts/suggest_figures.py` to rank candidates; read [references/figure-sourcing.md](references/figure-sourcing.md) before using an external figure.
2. Choose a genre and narrative from [references/genre-playbooks.md](references/genre-playbooks.md), then use [references/outline-patterns.md](references/outline-patterns.md) as a starting point. Do not copy a fixed page sequence across genres: a theory proof, a systems evaluation, and a humanities close reading need different evidence objects.
3. Choose each page's evidence role and layout from [references/layout-catalog.md](references/layout-catalog.md); use layout variants (`figure.image_right`, `quote.split`, etc.) when the same content contract needs a different composition. Read [references/visual-language.md](references/visual-language.md) and [references/design-language.md](references/design-language.md) before designing a new theme or layout.
4. Build `deck.json` using [references/deck-schema.md](references/deck-schema.md).
5. Set `meta.deck_mode` to `content_only` when the user explicitly asks to remove the cover and table of contents; otherwise keep `full`.
6. For each content page, write a short topic `title` (4–10 Chinese characters, default `title_max_chars: 10`) plus a complete-sentence `lead` claim. Do not put the full argument into the fixed header. The validator warns when a header reads like a sentence. Humanities decks may let a quotation or image carry more of the page, but keep a short title for navigation and accessibility. Only when the user explicitly asks for an assertion-driven deck, set `meta.title_mode` to `assertion` and raise `title_max_chars` deliberately.
7. Aim for 3–5 content units per content page and one meaningful visual anchor: a real figure, table, chart, genuine pipeline, metric block, formula, or a clearly labeled concept schematic. If no real image exists, use the user's paper figures first, then an open-access paper figure with a source line, then a content-derived `footer` / `output` strip, formula, table, or semantic `diagram`; never insert an unrelated placeholder image.
8. Write every slide's `speech` with [references/speech-writing.md](references/speech-writing.md), then run `scripts/build_ghost_deck.py`. `build_speech.py` now also writes `speech_audit.md/json`; read [references/speech-deai.md](references/speech-deai.md) and treat A/B hits as rewrite candidates. Do not auto-change facts, numbers, citations, or X-level items. If the title chain plus claims cannot be understood without the full speech, fix the titles and leads before styling.
9. Select or copy a theme. Use `assets/themes/academic_blue/theme.json` as the public default, `assets/themes/zju_blue/theme.json` as a Zhejiang University editorial example, `assets/themes/tju_blue/theme.json` for the Tianjin University visual system, or `assets/themes/nanjing_purple/theme.json` for the official calligraphic-wordmark example. Use `document` for dense handouts, `projection` or `lecture` for teaching, and `defense` for a projected defense deck with 24 pt body text, 32–40 pt headings, and semantic body bold. Read [references/defense-readability.md](references/defense-readability.md) before diagnosing an existing PPTX or using the `defense` profile.
10. If the user provides an existing PPTX template, run `scripts/import_pptx_theme.py` to extract its theme colors and major/minor fonts into a `theme.json` draft. Review the draft before using it. If the user provides an official seal or school-name lockup, place it in the theme folder and update `brand.lockup`. If the emblem is a local file or a public URL, run `scripts/build_brand_lockup.py` to remove the white background, crop it, compose the "emblem + name" lockup, and write the asset paths back to `theme.json`. Prefer an official horizontal lockup via `--lockup` (with `--lockup-crop` and `--recolor-lockup`) so the school name uses its official calligraphic wordmark; typeset composition is only a fallback. Use `--search-query` to look up a candidate on Wikimedia Commons; always show the selected source to the user. Do not invent or redraw an official seal.
11. Render a five-page sample first (`--sample 5`) and inspect it. After the visual direction is confirmed, render the full deck, export the native editable PowerPoint, and validate both:

```bash
python scripts/render_deck.py --deck path/to/deck.json --theme assets/themes/academic_blue/theme.json --out path/to/output --profile document
python scripts/export_editable_pptx.py --deck path/to/deck.json --theme assets/themes/academic_blue/theme.json --out path/to/output --pptx-name defense_deck.pptx --profile document
python scripts/build_speech.py --deck path/to/deck.json --out path/to/output
python scripts/build_ghost_deck.py --deck path/to/deck.json --out path/to/output
python scripts/build_qa.py --deck path/to/deck.json --out path/to/output --count 10
python scripts/validate_deck.py --deck path/to/deck.json --theme assets/themes/academic_blue/theme.json --render-dir path/to/output/rendered --out path/to/output --profile document
python scripts/validate_pptx.py --pptx path/to/output/defense_deck.pptx --deck path/to/deck.json --out path/to/output
python scripts/check_accessibility.py --pptx path/to/output/defense_deck.pptx --out path/to/output
python scripts/render_pptx_preview.py --pptx path/to/output/defense_deck.pptx --out path/to/output/pptx_preview --compare-dir path/to/output/rendered --strict
python scripts/export_agent_brief.py --deck path/to/deck.json --theme assets/themes/academic_blue/theme.json --out path/to/output/agent_brief.md
python scripts/audit_pptx.py --pptx path/to/existing_or_rebuilt.pptx --out path/to/output
```

Before building the deck:

```bash
python scripts/extract_source.py --input path/to/paper.pdf path/to/figures --out path/to/source_pack
python scripts/import_pptx_theme.py --pptx path/to/school_template.pptx --out path/to/theme/theme.json --name "学校名称"
python scripts/build_brand_lockup.py --theme path/to/theme/theme.json --name "学校名称" --latin "UNIVERSITY NAME" --emblem path/to/emblem.png
python scripts/build_brand_lockup.py --theme path/to/theme/theme.json --name "学校名称" --lockup path/to/official_lockup.png --lockup-crop "0,0,236,88" --recolor-lockup
python scripts/find_figures.py --query "vision language action robot" --out path/to/figure_pack
python scripts/suggest_figures.py --deck path/to/deck.json --figures path/to/figure_pack/figures.json --out path/to/figure_suggestions --query "vision language action robot"
```

12. Inspect the contact sheet and the PowerPoint-rendered preview if PowerPoint is available. Regenerate only pages with overflow, missing assets, or low density.

## Fixed Visual Contract

Read [references/style-contract.md](references/style-contract.md) before changing the renderer or theme.
Read [references/editable-pptx.md](references/editable-pptx.md) before changing native PowerPoint export, font mapping, or OMML formulas.

- Cover: no header bar and no mini-directory. Keep one institutional identity and one cover composition.
- The cover may be centered or editorial; choose once in the theme. The editorial variant uses a left title column and a faint watermark, but still keeps one institutional identity.
- Outline: only the five or six section titles, with no descriptions, page ranges, or mini-directory.
- Content-only mode: omit both cover and outline when explicitly requested; start directly with the first content page while keeping the fixed header, footer, identity, and mini-directory.
- Content pages: fixed header, fixed upper-right identity, fixed footer rule, and a five-item mini-directory from the third page onward.
- Every slide must have a real PowerPoint title placeholder; pictures and charts need semantic alt text; background textures, lockups used as decoration, and non-text chrome must be marked decorative; comparison tables need header-row semantics.
- Treat `source_manifest.json` as the evidence index: every figure used in a slide should be traceable to a source file and page/section.
- A theme imported from a PPTX is a draft. Review the mapped palette, fonts, and identity assets before delivery; do not assume a template's accent colors are semantically correct for a new topic.
- Chapter numbers, outline row numbers, and their adjacent titles use the same measured font size and baseline. If a long title has to shrink, the number shrinks with it.
- Keep body content inside the theme's `content` box. Never move fixed elements to compensate for crowded text.
- Use paper-like pale backgrounds, deep institutional blue, thin rules, restrained panels, and faithful rectangular figure inserts.
- Use `grid` for three to six compact mechanism blocks. Do not leave large empty cards when the source page is information-dense.
- When a grid block has no real image, prefer `footer` or `output` for a short, content-derived judgement, deliverable, or metric. Do not use `visual` as a generic placeholder.
- Use `figure` when one real figure is the argument, `figure_flow` for mechanism-plus-evidence, `pipeline` for a genuine sequence, `balance` for a real tradeoff around one central system figure, and `hub` for taxonomy or route panoramas. Do not turn every page into the same left-right card split.
- Use `concept` when a causal, structural, threshold, or feedback relation needs a labeled schematic. The six semantic kinds are `sequence`, `branch`, `stack`, `ruler`, `matrix`, and `cycle`. Every label and value comes from `deck.json`; the renderer only owns geometry.
- Use `quote` for humanities close reading, textual evidence, and passages that carry the argument. Include `attribution` and `commentary`; never use a quotation as decoration.
- Use `formula` when one equation is the page's main object. Keep `formula_latex` as the typeset source and `formula` as the portable fallback; add symbol definitions, a condition strip, and explicit boundary panels. Do not use the formula layout as a decorative code block.
- A `chart` object must become a native editable PowerPoint chart, not a screenshot. Keep the same values, highlight, axes labels, caption, and palette as the deterministic preview. If neither `image` nor `chart` is provided, show a missing-asset notice rather than an unrelated placeholder.
- Keep measured figures dominant: a hero figure should normally occupy 65%–80% of the content area, with a narrow interpretation rail rather than a wall of small cards.
- Use red only for a decisive number, risk, or boundary; use one blue family for ordinary structure. Do not use a rainbow of semantic colors.
- On deep-blue bands, map every text token to a light foreground token. Red emphasis on a dark band must be lightened for contrast.
- `**text**` is red bold emphasis and must not add an underline. Use `__text__` only when an underline is intentionally requested.
- Do not add handwritten marginal notes, stars, doodles, fake page numbers, fake references, glassmorphism, neon, or decorative 3D cards.

## Evidence Discipline

- Never redraw a paper curve, measured distribution, hardware photo, or experimental setup.
- Put uncertain or literature-derived numbers in `source_note`, not in the visual as an unqualified claim.
- State experiment conditions beside comparative claims. Use `不同实验条件，不作严格排名` when appropriate.
- Keep `claim`, `evidence`, and `boundary` separate in `deck.json`.
- A concept diagram must never be described as measured evidence.
- Put a source/citation note in `source` or `cites` when a figure, table, or borrowed number appears. Keep the deck-level `citations` map real: title plus venue/year or DOI/URL. Set `meta.source_policy` to `advisory` or `strict` to make missing sources visible. If the number is not comparable across papers, say so in the footer.
- Charts must carry `unit`, `sample_size`, `condition`, and `source` when those facts exist. Do not claim `error_bar` unless the exported chart actually shows uncertainty.
- Generated illustrations may clarify a mechanism, but never carry exact Chinese labels, numbers, axes, or citations. Those belong in the deterministic layer.
- For formulas, store the portable plain fallback in `formula` and the LaTeX source in `formula_latex`. The renderer uses LaTeX-style math typesetting when available and falls back to the plain string without failing.
- Use real subscript characters or LaTeX math for symbol rows such as `tₑ` or `uₑ`; do not leave code-like `t_e` in the visible slide.
- Treat `render_report.json` warnings as delivery blockers when they indicate overflow, missing assets, or incomplete content. Use `render_pptx_preview.py` for a real PowerPoint/LibreOffice rendering check; do not rely on the PIL preview alone.
- Use `projection` or `lecture` for teaching and `defense` for a defense-room deck. If a projected profile reports overflow, split the page or remove secondary text; do not shrink below the profile minimum to force it.
- Keep 24 pt defense body text bold, but leave footnotes, reference metadata, captions, and micro-labels regular. `text.body_bold_min_px` / `text.body_bold_max_px` define the semantic body range; do not bold every small run. Use `audit_pptx.py` to check the exported PPTX, not just the PNG preview.
- Make one evidence object dominant. For `figure` and `figure_flow`, `image_width_ratio` (0.45–0.82) controls the hero area while the source figure remains contained and unstretched; `rail_ratio` is an alias. Use `grid` `span`, `footer`, `output`, or `stat` to break equal empty cards.
- The native exporter converts `formula_latex` to OMML through `latex2mathml` and Office's `MML2OMML.XSL`. Keep the plain `formula` fallback. If OMML conversion is unavailable, the PPTX uses an editable Cambria Math text fallback rather than a raster formula.

## Public-Use Fallbacks

The skill is intended for users beyond one school or one topic. See [references/theme-authoring.md](references/theme-authoring.md) when creating a new institutional theme.

- Missing school seal or wordmark: use a text lockup or neutral monogram from the theme; never invent an official mark.
- Missing theme font: fall back through the theme's ordered font candidates; keep the same hierarchy and spacing.
- Missing real image or paper figure: use a content-derived `footer` / `output` strip, a table, a formula, or a textual structure. Do not draw a generic target, histogram, node chain, or noise plot.
- Missing a measured figure but having a real mechanism relation: use `concept` with a semantic `diagram`, label it `机制示意` or `研究接口`, and state what it does not prove.
- `visual` is only drawn when the deck explicitly sets `visual_semantic: true`; otherwise it is ignored. Never use `visual` as an unlabeled placeholder.
- Missing optional metadata: leave the field blank or write `待补充`; validation must report incomplete required fields.
- Different school or discipline: copy the theme folder, replace colors, identity assets, fonts, and background motif, then update the theme name. Do not hard-code a university or a research topic in the renderer.

## Portable Reproduction

When the target environment cannot execute Python, run `scripts/export_agent_brief.py`. Give the generated `agent_brief.md` and the user's materials to that agent. The brief contains the same section flow, slide fields, fixed branding coordinates, and speech rules in a model-readable form.

Do not claim pixel-identical output from an image-only model. The guarantee is consistent structure and branding when the renderer is available; without the renderer, the brief guarantees content and layout intent, not pixel fidelity.

## Design Basis

See [references/research-basis.md](references/research-basis.md) for the GitHub sources and the specific rules adapted from them. Use the sources as design evidence, not as content sources.
