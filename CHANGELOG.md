# Changelog

## Unreleased

### Semantic relation maps and arrow vocabulary (2026-11-17)

- Added the seventh `concept` diagram kind, `relations`: 2–8 nodes, 1–10 explicit directed edges, and optional `arrangement: free | row | two_row`.
- Added five semantic arrow styles — tapered `ribbon`, `block`, `chevron`, `line`, and bidirectional `double` — plus `straight` / `orthogonal` routes, `auto | left | right | top | bottom` anchors, and `primary | muted | alert` tones.
- Relation maps draw arrows behind nodes and labels above them, so a card cannot hide an arrowhead and a relation label cannot disappear under a gradient.
- PNG and native PowerPoint now share the same relation geometry. The PPTX uses editable freeform polygons with a real linear gradient, native connectors for orthogonal paths, and named `relation_node_N`, `relation_arrow_N`, and `relation_label_N` objects.
- Relation maps may omit interpretation panels; they then occupy the full content width and default to `diagram_frame: false`. Zero panels remain valid only for this diagram kind.
- Added edge validation for missing endpoints, self-loops, duplicate edges, unknown styles/routes/tones/sides, non-positive widths, out-of-bounds nodes, and overlapping cards. The agent brief now records `from → to` semantics and edge labels.
- Added the relation-map gallery slide and regression tests for chevron geometry, PNG gradient direction, native gradients, diagonal gradient parity, orthogonal final arrowheads, connector width, node bounds, label clamping, and zero-panel concept pages.

### Iteration protocol for weaker models (2026-11-17)

- Added a screenshot-feedback triage protocol: identify the element, defect class (family / size / weight / line height / baseline / overflow / alignment), peer reference, and shared rule before changing code; never start by replacing the global font or theme.
- Added two-scale visual review: 25% contact sheet for page rhythm, 100% crops for headers, module indices, card titles, captions, footnotes, and formula labels. Each pass must inspect the densest page, a real-figure page, a formula page, and a metrics page at full size.
- Added the PNG/PPTX parity witness set: header-number page, formula page, metrics condition grid, figure-card chrome, and closing cards must be checked in real PowerPoint after every export.
- Documented the font-size / line-height / container-height triple: raising `min_size` without recomputing containers creates larger text with worse overlap; if the three cannot fit together, cut content or split the page.
- Added defense content budgets (4 condition pairs, 22–26-character single-line caption, single-line source footnote, 2 short closing bullets) and the rule that “it fits” is not the same as “it is readable at 100%”.
- Added defect-to-geometric-invariant tests, a shared-function call-site checklist, the distinction between user feedback and internal convergence, the silent-gate rule, and a fill-in sprint card for weaker models.
- Native parity is now executable: the exporter names `condition_label_N` and `figure_card`, and `validate_pptx.py` rejects condition label/value overlap, wrapped defense condition values, non-single-line figure chrome, and chrome that escapes its figure card. Regression tests cover each failure.

### Peking University theme and defense geometry (2026-11-17)

- Added `assets/themes/pku_red/`: 北大红 `#94070A`, paper ground, official horizontal lockup and emblem, fixed sidebar, pale red gradient card headers, and the `defense` type profile. Official VI provenance is recorded in `BRAND_SOURCE.json`, `SOURCE.md`, and the new `references/pku-theme.md`; the wordmark stays an image and is never re-typeset or recolored.
- Added `references/pku-theme.md` with the palette, asset URLs, font fallbacks (Noto Serif SC / Noto Sans SC when 方正大标宋 / 方正中等线 are absent), cover/header rules, figure and closing contracts, and the non-affiliation rule for papers written at other institutions.
- Added `background.cover_arc` so an editorial cover can suppress the large background arc without disabling the subtle paper arcs on content pages.
- Formula pages now reserve height for the real condition-strip font instead of a fixed strip, so the condition line cannot collide with the formula or the symbol list.
- Native formula export now groups fenced operators so a Unicode minus stays a minus in OMML instead of becoming `m:sepChr` and painting as an arrow in PowerPoint.
- Added theme-level `soft_gradient` graphics: card headers and pale card faces now share horizontal gradient endpoints between the PIL renderer and native PPTX. The PKU theme uses Beida Red fading toward a lighter rose, with pale card faces fading to paper white.
- Figure captions and footnotes now reserve their measured glyph height; the footnote sits between the takeaway band and the footer rule instead of pressing into the rule.
- Closing pages with boundary/next-step cards use the free band above the closing line, and the line is bounded by `layout.footer_y`, so the last bullet is no longer clipped by a fixed card height. The three conclusion rows use 88 px when bottom cards are present.
- Figure-card badges now paint after the image mount, so a source tag cannot be half-covered by the white photo mount.

### Defense typography, atomic labels, and card chrome (2026-11-17)

- Made module indices and module headings share one measured font, size, top edge, and baseline in both the PNG renderer and native PPTX; defense module headers now reserve at least 76 px, and module indices are measured with a 60 px minimum width and `word_wrap=False`.
- Bold runs now change weight without switching font family. The exporter measures formula labels in the actual bold face, keeps them single-line, and starts the equation below the label.
- Added semantic CJK wrapping: closing punctuation (`，。；：、）》」』`) is bound to the previous unit in both plain and rich-text paths, so a line cannot begin with a comma or period.
- Figure cards now draw their label tab, source badge, and caption inside the card for defense, reserve measured caption height, and paint badges after the image mount so they cannot be half-covered.
- Defense metrics condition panels now use a 2×2 label/value grid with a full band per pair; label and value no longer overlap, separator rules are omitted when they would cross descenders, and decks are capped at four conditions per metrics page.
- Defense captions and card chrome were raised from 17–20 px to the 30–38 px range, matching the 18 pt secondary-text floor visible in the native PPTX.
- Added regression tests for the condition-grid bands, single-line condition values, and native figure label/badge/caption geometry.

### Reproducible iteration gate (2026-11-17)

- Added `scripts/quality_gate.py`, which aggregates render, validation, design, native-PPTX, accessibility, readability and real-render reports into one `quality_gate.json` / `quality_gate.md`.
- The gate checks report freshness against deck, theme, referenced images and generator scripts; stale reports cannot masquerade as delivery evidence.
- Added stage gates (`sample`, `deck`, `pptx`, `final`), navigation contract checks, page-count consistency, defense readability enforcement, and a page-by-page fix queue with the next repair action.
- Added `references/iteration-playbook.md`: five stages, one-defect-per-pass protocol, symptom-to-fix table, freeze/stop rules, and a weak-model mode that forbids inventing layouts or lowering the type floor.
- Documented the loop in `SKILL.md` and `README.md`; tests cover pass, overflow blocking, advisory `chars_per_line`, stale reports, `font_upgrade`, navigation conflicts and missing PPTX.

### Tsinghua defense showcase (2026-11-17)

- Added a 12-page Tsinghua-purple defense showcase preview at `docs/tsinghua_gmapd_preview.jpg`, built from the GM-APD array-lidar paper with `meta.navigation_mode: sidebar`.
- Moved the deck's tuned `defense` type profile, body-bold range, and header geometry into `assets/themes/tsinghua_purple/theme.json`, so the published preview is reproducible from the shipped theme.
- Recorded the presentation provenance in the theme's `SOURCE.md`: the visual template is Tsinghua, while the paper authors are at Tianjin University; no institutional endorsement is implied.
- Fixed `tests/test_defense_readability.py` creating `tmp*` directories in the repository root by keeping the agent-brief export inside its temporary-directory context.

### Navigation mode selection (2026-11-17)

- Added `meta.navigation_mode` with three authoring choices: `sidebar`, `dividers`, and `hybrid`. Explicit values override `layout.sidebar_mode`, while legacy decks map `always → sidebar`, `section → hybrid`, and `off → dividers`.
- Renderer, editable PPTX exporter, design audit, and validator now resolve the same navigation model, so content width, sidebars, and divider pages cannot disagree between PNG and PPTX.
- `add_section_pages.py` reads the selected mode: `sidebar` no-ops; `dividers` and `hybrid` insert chapter dividers. `--navigation-mode` records the choice in `deck.json`.
- Fixed `sidebar_mode: section` to show the mini-directory on each chapter's first content page, not on the full-page divider.
- Documented the difference between the outline page, standalone `layout: "section"` dividers, and the persistent sidebar; user prompts can now choose one of the three in plain language.

### Native geometry, metrics, and closing hierarchy (2026-11-17, tenth pass)

- **Header and cover:** chapter labels and page titles now share one measured size, top baseline, and deck-wide title column. The cover metadata row measures the advisor and date values, keeps the advisor on one line, and right-aligns the date instead of relying on fixed columns.
- **Outline:** row number and section title use one measured size and baseline; long titles shrink both together. The native PPTX path now prevents `01`-style index tokens from being wrapped into `0\n1`.
- **Formula:** symbol definitions use a readable definition list: one column up to six symbols, two columns above six, with consistent hierarchy for `K`, `β`, `Nt`, `Nd`, and `ηd`.
- **Metrics:** metric cards use a vertical value → label → note flow with computed height. The evidence image and 2×N condition grid share the lower band, and the condition note is constrained to the side panel's real width.
- **Concept diagrams:** sequence nodes occupy 66% of the internal height, scale title/body type to node width, allow three body lines, and compute the body box with node-relative coordinates so it cannot receive a negative height.
- **Closing:** conclusion slides support card-only pages without placeholder `points`; cards fill the band between the lead and closing line. Card titles render at `52+ px` over `39 px` body text, share the body's left edge, and place the icon at the trailing edge.
- **Validator contracts:** summary-mode titles now accept the documented 4–18 character range (the five-character topic “多维特征链” is valid); closing pages with boundary/next-step cards count as substantive summaries; section dividers break layout runs instead of making figure pages on either side look consecutive.
- **Regression gates:** `tests/test_typography_alignment.py` covers header/outline alignment, index wrapping, concept-node containment, metrics-note geometry, and closing hierarchy. `validate_pptx.py` now fails on non-positive text geometry, wrapped two-digit indices, and escaped diagram-node bodies.

### Header, outline, and concept type hierarchy (2026-11-16, ninth pass)

- Header eyebrow labels (`1.1`, `2.3`) now render and export at the measured title size and use the same top baseline. The `defense` profile no longer silently overrides `eyebrow` to 42 px while the title is 72 px.
- Outline row numbers now inherit the section-title measured size and share its baseline. Long titles shrink together with their numbers instead of leaving a decorative 112 px numeral beside 62 px text.
- `concept/branch` uses one explicit type ladder (`56/42/42/38`) and vertically centers the root and card copy; card bodies are no longer smaller than their headings.
- Added typography-alignment regression coverage for the native PPTX path.

### Outline page measured from a screenshot (2026-11-16, eighth pass)

- `references/outline-page.md` (new): single-column geometry derived from header and footer, matched number/title type sizes, 30 px descriptions, hairline rows.
- Fixed the two failure modes behind "the TOC is ugly": a hard-coded `top = 520` that piled five rows into the lower half, and a two-column grid that left a 3+2 hole.
- Numerals moved from `blue_pale` to the solid primary colour - pale on white is a contrast bug, not a style choice.
- PNG and PPTX outline geometry are now identical; the pptx first row sits at canvas_top x 0.625.


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
