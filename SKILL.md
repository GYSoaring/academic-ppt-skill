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
- `agent_brief.md`: portable instructions for Doubao or another model that cannot run these scripts.

## Workflow

1. Read the user's paper, outline, images, school identity assets, presenter, advisor, date, and time limit.
2. Choose a structure from [references/outline-patterns.md](references/outline-patterns.md).
3. Choose each page's evidence role and layout from [references/layout-catalog.md](references/layout-catalog.md); read [references/visual-language.md](references/visual-language.md) and [references/design-language.md](references/design-language.md) before designing a new theme or layout.
4. Build `deck.json` using [references/deck-schema.md](references/deck-schema.md).
5. Set `meta.deck_mode` to `content_only` when the user explicitly asks to remove the cover and table of contents; otherwise keep `full`.
6. For each content page, write a short topic `title` (4–10 Chinese characters) plus a complete-sentence `lead` claim. Do not put the full argument into the fixed header.
7. Aim for 3–5 content units per content page and one meaningful visual anchor: a real figure, table, chart, genuine pipeline, metric block, formula, or a clearly labeled concept schematic. If no real image exists, use a content-derived `footer` / `output` strip, a formula, a table, or a semantic `diagram`; never insert an unrelated placeholder diagram.
8. Write every slide's `speech` with [references/speech-writing.md](references/speech-writing.md).
9. Select or copy a theme. Use `assets/themes/tju_blue/theme.json` for the fixed Tianjin University visual system, `assets/themes/zju_blue/theme.json` as a Zhejiang University editorial example, or `assets/themes/academic_blue/theme.json` for a generic institutional theme.
10. If the user provides an official seal or school-name lockup, place it in the theme folder and update `brand.lockup`. Do not invent or redraw an official seal.
11. Render the deterministic preview, export the native editable PowerPoint, and validate both:

```bash
python scripts/render_deck.py --deck path/to/deck.json --theme assets/themes/tju_blue/theme.json --out path/to/output
python scripts/export_editable_pptx.py --deck path/to/deck.json --theme assets/themes/tju_blue/theme.json --out path/to/output --pptx-name defense_deck.pptx
python scripts/build_speech.py --deck path/to/deck.json --out path/to/output
python scripts/validate_deck.py --deck path/to/deck.json --theme assets/themes/tju_blue/theme.json --render-dir path/to/output/rendered --out path/to/output
python scripts/validate_pptx.py --pptx path/to/output/defense_deck.pptx --deck path/to/deck.json --out path/to/output
python scripts/export_agent_brief.py --deck path/to/deck.json --theme assets/themes/tju_blue/theme.json --out path/to/output/agent_brief.md
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
- Chapter numbers, outline row numbers, and their adjacent titles use the same measured font size and baseline. If a long title has to shrink, the number shrinks with it.
- Keep body content inside the theme's `content` box. Never move fixed elements to compensate for crowded text.
- Use paper-like pale backgrounds, deep institutional blue, thin rules, restrained panels, and faithful rectangular figure inserts.
- Use `grid` for three to six compact mechanism blocks. Do not leave large empty cards when the source page is information-dense.
- When a grid block has no real image, prefer `footer` or `output` for a short, content-derived judgement, deliverable, or metric. Do not use `visual` as a generic placeholder.
- Use `figure` when one real figure is the argument, `figure_flow` for mechanism-plus-evidence, `pipeline` for a genuine sequence, `balance` for a real tradeoff around one central system figure, and `hub` for taxonomy or route panoramas. Do not turn every page into the same left-right card split.
- Use `concept` when a causal, structural, threshold, or feedback relation needs a labeled schematic. The six semantic kinds are `sequence`, `branch`, `stack`, `ruler`, `matrix`, and `cycle`. Every label and value comes from `deck.json`; the renderer only owns geometry.
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
- Put a source/citation note in `footnote` when a figure, table, or borrowed number appears. If the number is not comparable across papers, say so in the note.
- Generated illustrations may clarify a mechanism, but never carry exact Chinese labels, numbers, axes, or citations. Those belong in the deterministic layer.
- For formulas, store the portable plain fallback in `formula` and the LaTeX source in `formula_latex`. The renderer uses LaTeX-style math typesetting when available and falls back to the plain string without failing.
- Use real subscript characters or LaTeX math for symbol rows such as `tₑ` or `uₑ`; do not leave code-like `t_e` in the visible slide.
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
