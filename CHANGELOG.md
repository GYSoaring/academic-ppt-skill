# Changelog

## Unreleased

### Section dividers from a measured reference (2026-11-16, seventh pass)

- Ported the blue-band divider from a real template (991x620 screenshot): header rule at 0.19h, band 0.46h-0.855h ending at 0.968w, 100 px white title, 38 px English label, hairline with an end dot, and a large number straddling the band with a vertical gradient.
- PNG builds the gradient with a text mask; PPTX cannot gradient text, so it renders the numeral to a transparent PNG and places it as a picture.
- `scripts/add_section_pages.py` (new): inserts one divider before each section's first slide, idempotent, and auto-derives English labels from a hint table.
- `references/section-dividers.md` (new): the measured ratios and the required/optional fields.
- `validate_deck.py`: `layout: section` is exempt from the header/footer pixel-identity checks.


### Titles describe, claims move into the body (2026-11-16, sixth pass)

- **`title_mode: summary` is now the default.** Page headings summarise what the page covers (6-18 characters); the judgment lives in `lead` and `claim`. Seven claim-titles in a row read as seven slogans.
- `validate_deck.py`: summary mode checks heading length, rejects assertion cues in the heading, and requires a `claim` on every page — the judgment may not vanish from both places.
- `qa_design.py`: `assertion_ratio` is skipped in summary/topic mode; punchy headings are no longer scored.
- **`references/titles.md` (new).** Three regimes — cover titles (six Nature patterns, constraint-as-contribution), page titles (descriptive), panel headings (two to six character nouns) — with before/after examples.


### Academic register is now enforced (2026-11-16, fifth pass)

- **`references/academic-register.md` (new).** Structure for a literature review (five moves, each page carrying claim/evidence/boundary) plus eleven categories of AI-writing tells translated into Chinese: over-claiming, significance hype, empty intensifiers, novelty padding, formulaic openers, connective overuse, contribution cliches, citation dumping, hedging-by-vagueness, boilerplate emphasis, clause stacking. Distilled from `zLanqing/codex-claude-academic-skills` (6.0k★) and `AIScientists-Dev/academic-humanizer` (5.3k★).
- **`scripts/lint_register.py` (new).** Scans visible deck text for colloquial markers, hype, unattributed attribution, number-free "significant" and vague hedges. Speech notes are exempt. Reports rule, match and suggested rewrite; `--strict` fails on warnings too.
- **First catch**: a closing line and a false positive on 及其实验变体; the 其实 rule now excludes 及/极/尤 and closing slides are exempt from the lectern-phrase rule.
- Note recorded: a review deck legitimately scores lower on `dominant_object_ratio` than a research deck, so the rubric must not be satisfied by inventing figures.


### Reference measurement, not reference guessing (2026-11-16, fourth pass)

- **`scripts/analyze_reference.py` (new).** Reads any .pptx and reports objects/page, text boxes/page, text-area ratio, title pt per inch, body size, size-step count, bold ratio, palette and fonts. `--aggregate` produces a Markdown profile across several decks.
- **`references/human-deck-profile.md` (new).** Five real Chinese academic decks, 126 pages, measured. Medians: 15.0 objects/page, 8.9 text boxes, 0.25 text-area ratio, title 1.8 pt/in, body 16 pt, 13 size steps.
- **Header type reduced**: `LECTURE_TYPE_SCALE.header_title` 68 → 58 px, putting the title at 2.04 pt/inch against a human median of 1.8-2.0. The old value was 2.39.
- **Containers named as the real gap**: generated decks ran 28 objects/page against 15, and the difference is card and node chrome (~2.5 shapes per container), not content. Documented as the next target.


### Readable diagrams, icon library, editorial cover (2026-11-16, third pass)

- **Diagram labels were 7 pt.** `design.py` hard-coded 13-16 px for diagram indexes, node titles, node bodies and notes; they are now scaled ×1.6, and the PPTX `add_text` path clamps every run to the smallest type step (18.75 pt at `lecture`). The minimum measurable size in a generated deck went from 6.1 pt to 18.8 pt.
- **`scripts/icon_library.py` (new).** Fetches any Lucide icon by name (three CDN fallbacks plus a `curl` retry), recolours it to the theme, and rasterises it with a built-in path renderer — no cairo, no Inkscape. Verified: `satellite` → 160×160 PNG.
- **`resolve_emblem.py` hardened** with a curl fallback: on Windows installs where Python's urllib resets TLS, the resolver used to report a false "not found". Wuhan University now resolves from Commons (684×738).
- **Editorial cover carries real hierarchy**: kicker line, 118 px title, enlarged institutional lockup, watermark column.
- Documented that a typeset calligraphic name is a fallback and the official wordmark must be sourced or supplied.


### Brand, type hygiene and emblem sourcing (2026-11-16, second pass)

- **Fixed header-title overflow in PowerPoint.** The exporter hard-coded a 62 px title box while the lecture title is 68 px, so titles spilled into the divider in PowerPoint (invisible in the PNG preview). Box height now derives from the measured size (`title_size × 1.28`).
- **Smallest type raised to 18.75 pt.** Lecture scale: `footnote` 36 → 40 px, `micro` 36 → 40 px, `small` 38 px. Nothing in a projected deck should be smaller.
- **Calligraphic school name.** New `fonts.brand` / `pptx_fonts.brand` slots; the brand line no longer borrows the title sans. `cn_defense` ships with 华文行楷 (`STXingkai`).
- **`scripts/resolve_emblem.py` (new).** Walks Wikimedia Commons → zh.wikipedia → school homepage → favicon, downloads the first usable candidate, and reports the trail. It never draws a seal; "not found" is a valid, reported outcome. Includes a `curl` fallback because Python's urllib resets TLS on some Windows installs (`WinError 10054`) while curl works — a false "not found" was masking real hits.
- **Pseudo-formula guard.** `validate_deck.py` now warns when visible text contains `N_s`-style ASCII subscripts; use `Nₛ`, `Eₚ`, `fₚᵣ` or keep it in `formula_latex`.
- **Covers default to `editorial`** in `cn_defense`, and the theme's header band grew to 132 px to hold a 68 px title.

### Design audit and lean themes (2026-11-16)

- **New** `scripts/qa_design.py`: ten structural proxies (dominant object, card-only ratio, emphasis density, assertion ratio, takeaway ratio, layout runs, image pages, characters per line, ink-coverage rhythm, chrome share) scored 0–100 with an A–D grade. `validate_deck.py` now surfaces failing checks as warnings with a fix hint.
- **New** `scripts/compare_human.py`: benchmarks a deck against human-made PPTX files on width-normalised geometry (type scale, font-step count, pictures, objects, characters, margins). Reports pass/fail per metric.
- **New** theme `assets/themes/cn_defense/`: palette measured from a real Chinese defense deck (primary `#1A4E9B`, accent `#A82F2E`), card header bars, lean defaults, documented in `references/cn-defense-style.md`.
- Fixed a crash that made **two of the three bundled examples unrenderable**: `draw_references` passed an extra positional argument, so `fit_text_block()` received `max_lines` twice.
- `layout.sidebar_mode` (`always` / `section` / `off`) with `layout.content_no_sidebar`; the mini directory no longer has to appear on every page. The PPTX exporter honours the same policy.
- `layout.outline_style: "lines"` renders the outline as a numbered list instead of five shadowed cards (44 → 23 objects on a five-section deck).
- Decoration switches, all default-on and backwards compatible: `panel.corner_marks`, `panel.card_index`, `panel.card_highlight`, `layout.block_icons`, `layout.lead_panel`. Together they take a 13-page deck from 33.4 to **26.2 objects per page** without touching text.
- The lecture type ladder was collapsed from 12 distinct sizes to 10 and the smallest step raised from 24 px to 34 px (≈16 pt), because human decks run on six or seven sizes and nothing below 16 pt.
- Formula variable rows are now one text object per line instead of three objects per symbol (separator, symbol, meaning); a 7-variable page drops from 65 to 51 objects.
- `references/cn-defense-style.md` documents the measured palette range (H 213–217°), the card-header-bar convention, and a per-layer object cost table.

- Defense readability now includes density-aware validation, a stricter PPTX rebuild workflow, research-backed 24/18/32 pt gates, sans-serif body enforcement, semantic body-bold enforcement, and a required paper-sourced figure with provenance metadata.
- The portable `agent_brief.md` now carries the same hard contract for other generators: bold sans-serif 24 pt body, 32–40 pt headings, 18 pt secondary chrome, and at least one provenance-tagged paper figure.
- Fixed the defense PNG renderer's clamped `fit_text` helper so it calls `fit_text_block` instead of recursing.

- Added the `defense` type profile, semantic body-bold range, Nanjing University official-wordmark theme, 16-page optics example, and `audit_pptx.py` physical PPTX readability audit.
- Made hero-figure geometry configurable through `image_width_ratio` / `rail_ratio`, and added projected-deck density, visual-anchor, and layout-rhythm warnings.
- Header titles are short topic labels by default; `title_max_chars` and the validator warn when a header becomes a sentence.
- Brand asset builder supports official horizontal lockups (`--lockup`) with crop and theme-color recoloring, so the school name uses its official calligraphic wordmark instead of a typeset substitute.
- Genre playbooks (`experiment`, `systems`, `theory`, `empirical`, `survey`, `design`, `humanities`, `clinical`) with narrative and evidence-page guidance.
- `quote` layout for humanities close reading, plus layout variants (`figure.image_left/image_right/chart`, `quote.centered/split`).
- `lecture` type profile and `validate_pptx --lecture-audit` physical font-size audit.
- `build_gallery.py` and a 20+ page layout gallery covering all six diagram kinds and layout variants.
- Figure sourcing: `extract_source.py` now captures PDF figure captions; `find_figures.py` searches arXiv and renders complete figure regions; `suggest_figures.py` ranks/attaches figures with source and review-required role/boundary fields.
- Speech de-AI audit adapted from `lieflat-less-ai-tone`: A/B/C/F grading, spoken-rhythm checks, X-level no-change list, and automatic `speech_audit.md/json` from `build_speech.py`.

## v0.1.0 — 2026-11-15

First public release.

- Deterministic academic deck renderer with cover, outline, fixed header, mini-directory, and footer.
- Native editable PPTX: text, shapes, tables, charts, and OMML formulas.
- Evidence layouts: figure, figure_flow, concept, formula, balance, pipeline, compare, metrics, timeline, references.
- Semantic diagrams: sequence, branch, stack, ruler, matrix, cycle.
- Source extraction from PDF, DOCX, PPTX, Markdown, and figure folders.
- Theme import from an existing PPTX.
- Brand asset builder: local/URL/Commons emblem lookup, white-background removal, crop, and lockup composition.
- `document` and `projection` type profiles with overflow gating.
- Accessibility checks: title placeholders, alt text, table headers, reading order.
- Real PowerPoint/LibreOffice render QA.
- Assertion titles and a five-minute ghost deck.
- Citations, source lines, DOI/arXiv validation, and chart measurement conditions.
- Sample-first rendering (`--sample`, `--pages`) and Q&A prediction.

## Earlier internal milestones

- `2094ae3` Image Gen 2.6 visual language and semantic diagrams.
- `ee53e9c` Projection profile, accessibility, and PowerPoint render QA.
- `26820b7` Source extraction and PPTX theme import.
- `454f490` Assertion titles, ghost deck, and traceable sources.
