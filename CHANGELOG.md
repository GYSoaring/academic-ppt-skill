# Changelog

## Unreleased

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
