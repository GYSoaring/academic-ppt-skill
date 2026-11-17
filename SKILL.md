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
- `quality_gate.json` / `quality_gate.md`: final PASS/FAIL gate over fresh render, validation, design, PPTX, preview, speech, and Q&A evidence, with a page-by-page fix queue.

## Workflow

1. Read the user's paper, outline, images, school identity assets, presenter, advisor, date, and time limit. If the user supplies PDF, DOCX, PPTX, Markdown, or a folder of figures, run `scripts/extract_source.py` first and review `source.md`, `source_manifest.json`, and `figures.json` before planning slides. Prefer the user's own paper figures over external images. If the source has no usable figure, run `scripts/find_figures.py` on the topic (or known arXiv ids) and `scripts/suggest_figures.py` to rank candidates; read [references/figure-sourcing.md](references/figure-sourcing.md) before using an external figure.
2. Choose a genre and narrative from [references/genre-playbooks.md](references/genre-playbooks.md), then use [references/outline-patterns.md](references/outline-patterns.md) as a starting point. Do not copy a fixed page sequence across genres: a theory proof, a systems evaluation, and a humanities close reading need different evidence objects.
3. Choose each page's evidence role and layout from [references/layout-catalog.md](references/layout-catalog.md); use layout variants (`figure.image_right`, `quote.split`, etc.) when the same content contract needs a different composition. Read [references/visual-language.md](references/visual-language.md) and [references/design-language.md](references/design-language.md) before designing a new theme or layout.
4. Build `deck.json` using [references/deck-schema.md](references/deck-schema.md). Before writing slides, choose exactly one `meta.navigation_mode`: `sidebar`, `dividers`, or `hybrid` (see "Navigation: choose one model" below). This is a separate decision from the genre and page layouts.
5. Set `meta.deck_mode` to `content_only` when the user explicitly asks to remove the cover and table of contents; otherwise keep `full`.
6. For each content page, write a short topic `title` (4–10 Chinese characters, default `title_max_chars: 10`) plus a complete-sentence `lead` claim. Do not put the full argument into the fixed header. The validator warns when a header reads like a sentence. Humanities decks may let a quotation or image carry more of the page, but keep a short title for navigation and accessibility. Only when the user explicitly asks for an assertion-driven deck, set `meta.title_mode` to `assertion` and raise `title_max_chars` deliberately.
7. Aim for 3–5 content units per content page and one meaningful visual anchor: a real figure, table, chart, genuine pipeline, metric block, formula, or a clearly labeled concept schematic. If no real image exists, use the user's paper figures first, then an open-access paper figure with a source line, then a content-derived `footer` / `output` strip, formula, table, or semantic `diagram`; never insert an unrelated placeholder image.
   **One dominant object per page.** A page needs a main thing — one figure, one table, one formula, one schematic, one number — and at most two supporting blocks. Three equal cards are not a hierarchy, they are a table of contents. If you cannot name the page's dominant object in one phrase, the page is a form.
   Prose columns must stay readable: keep Chinese body lines at 18 characters or more. A three-column `grid` at 1672 px puts 13 characters on a line and reads as chopped text, so prefer two columns and let one block span.
   For a defense-room deck, at least one `figure` / `figure_flow` page must use a paper-sourced image (`image_origin: own`, `public`, or `borrowed`) with `source`, `image_role`, and `image_boundary`; a decorative template image does not satisfy this requirement.
8. Write every slide's `speech` with [references/speech-writing.md](references/speech-writing.md), then run `scripts/build_ghost_deck.py`. `build_speech.py` now also writes `speech_audit.md/json`; read [references/speech-deai.md](references/speech-deai.md) and treat A/B hits as rewrite candidates. Do not auto-change facts, numbers, citations, or X-level items. If the title chain plus claims cannot be understood without the full speech, fix the titles and leads before styling.
9. Select or copy a theme. Use `assets/themes/academic_blue/theme.json` as the public default, `assets/themes/cn_defense/theme.json` for the Chinese defense / award-application look (measured palette, card header bars, no icon badges — see references/cn-defense-style.md), `assets/themes/zju_blue/theme.json` as a Zhejiang University editorial example, `assets/themes/tju_blue/theme.json` for the Tianjin University visual system, `assets/themes/nanjing_purple/theme.json` for the official calligraphic-wordmark example, or `assets/themes/pku_red/theme.json` for the Peking University red template (official lockup and emblem, 北大红 `#94070A`; see [references/pku-theme.md](references/pku-theme.md)). Use `document` for dense handouts, `projection` or `lecture` for teaching, and `defense` for a projected defense deck with 24 pt body text, 32–40 pt headings, and semantic body bold. Read [references/defense-readability.md](references/defense-readability.md) before diagnosing an existing PPTX or using the `defense` profile.
10. If the user provides an existing PPTX template, run `scripts/import_pptx_theme.py` to extract its theme colors and major/minor fonts into a `theme.json` draft. Review the draft before using it. If the user provides an official seal or school-name lockup, place it in the theme folder and update `brand.lockup`. If the emblem is a local file or a public URL, run `scripts/build_brand_lockup.py` to remove the white background, crop it, compose the "emblem + name" lockup, and write the asset paths back to `theme.json`. Prefer an official horizontal lockup via `--lockup` (with `--lockup-crop` and `--recolor-lockup`) so the school name uses its official calligraphic wordmark; typeset composition is only a fallback. Use `--search-query` to look up a candidate on Wikimedia Commons; always show the selected source to the user. Do not invent or redraw an official seal.
11. Render a five-page sample first (`--sample 5`) and inspect it. After the visual direction is confirmed, render the full deck, export the native editable PowerPoint, and validate both:

```bash
python scripts/render_deck.py --deck path/to/deck.json --theme assets/themes/academic_blue/theme.json --out path/to/output --profile document
python scripts/export_editable_pptx.py --deck path/to/deck.json --theme assets/themes/academic_blue/theme.json --out path/to/output --pptx-name defense_deck.pptx --profile document
python scripts/build_speech.py --deck path/to/deck.json --out path/to/output
python scripts/build_ghost_deck.py --deck path/to/deck.json --out path/to/output
python scripts/build_qa.py --deck path/to/deck.json --out path/to/output --count 10
python scripts/validate_deck.py --deck path/to/deck.json --theme assets/themes/academic_blue/theme.json --render-dir path/to/output/rendered --out path/to/output --profile document
python scripts/qa_design.py --deck path/to/deck.json --theme assets/themes/academic_blue/theme.json --render-dir path/to/output/rendered --out path/to/output
python scripts/validate_pptx.py --pptx path/to/output/defense_deck.pptx --deck path/to/deck.json --out path/to/output
python scripts/check_accessibility.py --pptx path/to/output/defense_deck.pptx --out path/to/output
python scripts/render_pptx_preview.py --pptx path/to/output/defense_deck.pptx --out path/to/output/pptx_preview --compare-dir path/to/output/rendered --strict
python scripts/export_agent_brief.py --deck path/to/deck.json --theme assets/themes/academic_blue/theme.json --out path/to/output/agent_brief.md
python scripts/audit_pptx.py --pptx path/to/existing_or_rebuilt.pptx --out path/to/output
```

## Typography And Card-Chrome Invariants

These rules are shared by the PIL renderer and the native PPTX exporter. A fix that only touches one path is not a fix.

- **One type ladder per header.** A module index (`01`, `02`, `1.2`) and its heading use the same font family, measured size, top edge, and baseline. In `defense`, a module heading is about 32 pt (68 px on the 2048 canvas) with a panel header at least 76 px tall; the index is never a decorative fixed size.
- **Bold changes weight, not family.** `bold=True` must keep the body/title family and only switch the weight. Mixing a display face into a bold body run is a defect.
- **Atomic tokens never wrap.** Page-header numbers, module indices, outline numbers, footnote/source lines, formula labels, metric values, and figure captions use `word_wrap=False` (or an explicit single line) and a box at least 16–24 px wider than the measured text. Module indices are at least 60 px wide.
- **Measure the real line count.** Two-line leads, two-line formula conditions, captions, and condition notes reserve the measured glyph height; never reserve a fixed 34–40 px band for text that is actually two lines. If a caption would wrap at the defense floor, shorten the caption instead of shrinking it.
- **Closing punctuation stays attached.** Chinese `，。；：、）》」』` must be bound to the preceding semantic unit. A wrapped line may not begin with a closing mark; both `wrap_cjk_text()` and rich-text `wrap_styled()` enforce this.
- **Figure-card chrome lives inside the card.** In `defense`, the blue label tab and the source badge sit on the card header line inside the card; the caption sits in a measured bottom band. They must not straddle the border into the metric row above or be covered by the image mount.
- **Condition grids are not stacked micro-text.** In `defense`, a metrics side panel shows at most four label/value pairs in a 2×2 grid. Each pair gets a full band; label and value do not share a 58 px row, and separator rules are omitted when they would touch glyph descenders. More than four conditions belongs in the notes, not the slide.

Before building the deck:

```bash
python scripts/extract_source.py --input path/to/paper.pdf path/to/figures --out path/to/source_pack
python scripts/import_pptx_theme.py --pptx path/to/school_template.pptx --out path/to/theme/theme.json --name "学校名称"
python scripts/build_brand_lockup.py --theme path/to/theme/theme.json --name "学校名称" --latin "UNIVERSITY NAME" --emblem path/to/emblem.png
python scripts/build_brand_lockup.py --theme path/to/theme/theme.json --name "学校名称" --lockup path/to/official_lockup.png --lockup-crop "0,0,236,88" --recolor-lockup
python scripts/find_figures.py --query "vision language action robot" --out path/to/figure_pack
python scripts/resolve_emblem.py --name "学校名称" --latin "UNIVERSITY NAME" --out path/to/emblem
python scripts/suggest_figures.py --deck path/to/deck.json --figures path/to/figure_pack/figures.json --out path/to/figure_suggestions --query "vision language action robot"
```

12. Inspect the contact sheet and the PowerPoint-rendered preview if PowerPoint is available. Regenerate only pages with overflow, missing assets, or low density. Before declaring the deck done, run the final quality gate below; a deck is not delivered while `quality_gate.md` says `FAIL`.

## Iteration loop: gate the deck, do not vibe-check it

A high-quality deck is the result of staged convergence, not one long generation. Read [references/iteration-playbook.md](references/iteration-playbook.md) before the first full render. Its short version:

| Stage | Do | Gate |
|---|---|---|
| 0. Content contract | `deck.json`: claims, evidence, boundaries, navigation mode, type profile | The title chain is understandable without the speech. |
| 1. Sample | 4–5 pages: cover, outline, densest content page, real-figure page, closing | `quality_gate.py --stage sample` |
| 2. Full deck | Full PNG render, structure validation, design audit | `--stage deck` |
| 3. Native | Editable PPTX, native validation, accessibility, readability audit | `--stage pptx` |
| 4. Delivery | Real PowerPoint render, speech timing, Q&A, portable brief | `--stage final` |

```bash
python scripts/quality_gate.py \
  --deck path/to/deck.json \
  --theme path/to/theme.json \
  --out path/to/output \
  --profile defense \
  --stage deck
```

The gate reads the existing reports, checks that they are newer than the deck/theme/images/generators, verifies page counts and navigation, classifies warnings, and writes two files:

- `quality_gate.json`: machine-readable checks, issues, and counts.
- `quality_gate.md`: a page-by-page fix queue plus the next repair action.

Rules that are not optional:

- **One pass, one defect class.** Do not fix overflow, alignment, and PPTX drift in the same pass.
- **Iterate with `--pages`; deliver with a full render.** Partial renders are not delivery evidence.
- **Two failed attempts on one page means change the strategy.** Split the page or cut content; never lower the type floor.
- **Freeze theme and geometry after stage 2.** Any theme, layout, or renderer change invalidates every report.
- **Stop after two consecutive PASS runs.** Do not chase a design score above 85 with decoration.

For weaker models, the gate is the interface: follow its fix queue, open only the one reference named by the issue, choose an existing layout from [references/layout-catalog.md](references/layout-catalog.md), and do not edit the renderer unless PNG and PPTX are both wrong. This is the loop that brought the 12-page Tsinghua GM-APD showcase to 86/A with zero native-PPTX and zero accessibility errors.

### Screenshot triage and native parity

When a screenshot looks wrong, do not change the global theme first. Name the element (index / heading / body / badge / caption / footnote), the defect class (family / size / weight / line height / baseline / overflow / alignment), the peer it must match, and the shared rule that fixes every page. A module index matches its heading in family, size, top edge, and baseline; bold changes weight only.

- A contact sheet is for rhythm only. Review the densest page, a real-figure page, a formula page, and a metrics page at 100% crops.
- PNG and PPTX are separate gates. After every export, inspect five witness pages in real PowerPoint: header number, formula label, metrics condition grid, figure-card chrome, closing card.
- When a minimum type size rises, recompute line height and container height in the same change. Never raise `min_size` alone.
- If a screenshot proves a defect but the validator reports zero, add semantic shape names and a geometric invariant test first; then fix the page. A silent gate is a missing gate.
- Defense content budgets: at most 4 condition pairs, a 22–26-character single-line caption, a single-line source footnote, and at most 2 short bullets per closing card.

Full protocol: [references/iteration-playbook.md](references/iteration-playbook.md).

## Brand and Type Hygiene

These are the defects that make a generated deck look machine-made even when the layout is sound. They are cheap to prevent and expensive to notice late.

- **Overflowing header titles.** The PowerPoint header title box must be at least `title_size × 1.28` tall. A box shorter than its font renders fine in the PNG preview (PIL does not clip) and spills into the divider in PowerPoint. The exporter now derives the height from the measured size; keep it that way when editing.
- **Nothing below 18.75 pt.** At the `lecture` profile the smallest step is 40 px (≈18.75 pt). Footnotes, captions and symbol rows included. If a page cannot fit its content at that floor, cut content or split the page.
- **No ASCII pseudo-formulas in visible text.** `N_s`, `E_p`, `f_prf` read as code, not math. Use real subscripts (`Nₛ`, `Eₚ`, `fₚᵣ`) or keep the expression in `formula_latex`. `validate_deck.py` warns on the underscore form.
- **The school name needs a hand-written face.** Set `fonts.brand` to a calligraphic CJK font (Windows: `C:/Windows/Fonts/STXINGKA.TTF`, 华文行楷) and `pptx_fonts.brand` to its family name (`STXingkai`). Never let the brand line fall back to the body sans — a typed school name is the single most obvious "this was generated" tell.

- **Graphic vocabulary beyond the built-in glyphs.** `scripts/icon_library.py` pulls any [Lucide](https://lucide.dev) icon (MIT, 1500+) by name, recolours it to the theme, and rasterises it — no cairo or Inkscape needed, because Lucide icons are 24×24 stroke paths and the script ships a small path renderer. Use it when a page needs a pictogram the renderer does not have (`satellite`, `gauge`, `network`, `shield-check`). Drop the PNG into the deck's `assets/` and reference it from a block `image`; do not hand-draw a new glyph in the renderer.
- **The school wordmark must be the institution's own calligraphy, not a font.** A typeset 行楷 is a fallback, never the goal. `resolve_emblem.py` looks for an emblem; if the wordmark lives only on the school site or a non-free wiki, ask the user for the file and wire it through `build_brand_lockup.py --lockup`. Say plainly which one is in use — "typeset approximation" versus "official wordmark" — in the theme's `SOURCE.md`.
- **The emblem must come from a real source.** Run `scripts/resolve_emblem.py --name 学校名 --latin "University Name"` first; it walks Wikimedia Commons → zh.wikipedia → the school homepage → favicon and reports what each step returned. If it finds nothing, ask the user for the official file. Drawing a seal is forbidden.
- **Covers default to `editorial`.** A left accent bar, left-aligned wordmark, a vertical rule and a quiet right column read as designed; a centred stack of title + subtitle + three info lines reads as a form.
- **Index tokens never wrap.** `01`, `02`, and chapter numbers are single-line tokens. On the native export, pass `wrap_explicit=False` for indices and other short numeric labels; never let the defense pre-wrapper turn `01` into `0\n1`.
- **Native geometry must be positive and contained.** A text box with zero/negative width or height is a layout error, even if the PNG preview happens to look correct. `validate_pptx.py` rejects non-positive text geometry, two-line index tokens, and concept-node bodies that escape their node.
- **Derive diagram boxes in the node's coordinate system.** For sequence nodes, keep the title offset relative to `node_top` and compute `body_height` from node-relative offsets. Do not subtract an absolute page `title_top` from a node height.
- **Metric and closing cards follow one vertical hierarchy.** Metric values outrank labels, and labels outrank notes. In closing cards, the title is `52+ px`, the body is `39 px`; both share the same left edge and the icon moves to the trailing edge.
- **PNG and PPTX are two outputs of one geometry.** When a native branch is rewritten, mirror the renderer's coordinates and run the typography regression tests plus `validate_pptx.py`; a successful PIL render is not evidence that the editable deck is correct.






## Outline page: one column that fills the page

`layout.outline_style: "index"` is the default for decks aimed at a Chinese audience. Measured spec in `references/outline-page.md`.

- left spine (3 px); the row number and section name share **one measured size and baseline** (62 px by default, 72 px under `defense`), in the solid primary colour
- if a long section name shrinks, its number shrinks with it; one-line description 30 px, hairline between rows
- `C O N T E N T S` top-right, letter-spaced

Two rules that matter more than the styling:

1. **Derive the geometry, never hard-code the start.** `top = header + 48`, `bottom = footer - 48`, `row_h = span / sections`. A fixed `top = 520` piles five rows into the lower half and leaves ~600 px of dead space above; that is the usual reason a TOC looks wrong.
2. **One column, not two.** Five sections in two columns becomes a 3+2 grid with a hole in the bottom-right corner.

Check it numerically: bin the rendered page into 12 horizontal bands and count dark pixels. If the first four bands are empty, the list is not filling the page. The PNG preview and the PPTX must use the same numbers — the pptx `top` should be about `canvas_top × 0.625`.

## Section dividers: the blue band and the number that straddles it

The standalone chapter divider is available as `layout: "section"` (measured spec in `references/section-dividers.md`). Use it only when `meta.navigation_mode` is `dividers` or `hybrid`:

- header without a title: emblem lockup left, motto right, hairline at `0.19h`
- a blue band from `0.46h` to `0.855h`, ending at `0.968w` so the right edge keeps a 3% margin
- Chinese title in white at 100 px, English label at 38 px, hairline with an end dot
- **a large two-digit number straddling the band** — dark above on white, pale below inside the band

The straddle is the whole point. A number placed on top of the band, or beside it, loses the effect.

PowerPoint cannot gradient text, so the PPTX path renders the same numeral to a transparent PNG (`assets/section_number_XX.png`) and places it as a picture; the PNG path builds it from a text mask with a vertical gradient. Keep the two geometries identical or the preview stops predicting the deliverable.

```bash
python scripts/add_section_pages.py --deck deck.json                         # 读取 meta.navigation_mode；sidebar 自动跳过
python scripts/add_section_pages.py --deck deck.json --navigation-mode dividers  # 显式选择并在 deck.json 中记下模式
python scripts/add_section_pages.py --deck deck.json --dry-run                # 只看会插几张
```

`validate_deck.py` exempts `layout: "section"` from the header/footer pixel-identity checks — those pages are intentionally a different full-page composition.

## Navigation: choose one model before building

Three different objects are involved. Keep their names separate:

- **目录页 / outline page**: the single `layout: "outline"` page near the start. It is always present in `deck_mode: full`; it is not a chapter divider.
- **章节分隔页 / section divider**: a standalone `layout: "section"` page inserted before each chapter.
- **侧栏目录 / sidebar mini-directory**: the small five-item directory pinned to the left of a content page under `layout.sidebar_mode`.

Ask the user once, or choose from the brief, and record it as `meta.navigation_mode`:

| `meta.navigation_mode` | What the audience sees | Best for |
|---|---|---|
| `sidebar` | outline page + persistent left mini-directory; no section dividers | 10–14 page linear talks where the audience needs constant orientation |
| `dividers` | outline page + one standalone divider before every chapter; no left sidebar | 16+ page defenses, course lectures, long reports that need chapter breathing room |
| `hybrid` | outline page + chapter dividers + the left mini-directory only on each chapter's first content page | long decks that need both chapter transitions and local orientation |

When the user does not specify: use `sidebar` for a linear deck of about 14 pages or fewer; use `dividers` for a 15+ page defense, lecture, or long report; use `hybrid` only when the user explicitly wants both chapter dividers and a recurring sidebar.

Prompt forms users can copy:

```text
导航用 sidebar：目录页 + 左侧常驻小目录，不要章节分隔页。
导航用 dividers：目录页 + 每章一张独立章节分隔页，不要左侧小目录。
导航用 hybrid：目录页 + 章节分隔页，每章第一张内容页显示侧栏目录。
```

Rules for the agent:

- Do not confuse `sidebar_mode: "section"` with `layout: "section"`. The former means "show the sidebar at section starts"; the latter means "insert a full-page chapter divider".
- `meta.navigation_mode` is the authoring field and takes precedence. `theme.layout.sidebar_mode` is the low-level renderer policy and should not be presented to users as a choice.
- If `navigation_mode` is `sidebar`, do not run `add_section_pages.py`; if invoked anyway, it reads the field and no-ops. If dividers or hybrid is selected, run it before rendering.
- If the user explicitly asks for both a persistent directory and chapter dividers, choose `hybrid`; that is what the combination means.
- Legacy decks without `meta.navigation_mode` keep their old appearance: `sidebar_mode: always` maps to `sidebar`, `section` to `hybrid`, and `off` to `dividers`.

## Titles: describe the page, argue in the body

Default `title_mode` is **`summary`**. A page title says *what this page is about* (usually 4-10 characters; the validator accepts 4-18), the way a paper's subsection heading does. The judgment goes into `lead` (first sentence) and `claim`.

Why: seven claim-titles in a row read as seven slogans, and the reader stops seeing the argument. A descriptive heading gives the map; the claim stays where the evidence can hold it up.

- `summary` — headings are descriptive; `validate_deck` requires a `claim` on every content page, rejects assertion cues in the heading, and checks length.
- `assertion` — opt-in for defense panels and executive briefings, where a conclusion must hit the reader in ten seconds. The cost is that every page then owes immediate evidence.

Cover titles are a separate regime and may carry tension (`Solving olympiad geometry without human demonstrations`); the six Nature-style patterns are catalogued in `references/titles.md`. Panel headings are a third regime: two to six characters, nouns only (现象 / 机制 / 约束), never questions like 看到什么.

`qa_design.py` skips `assertion_ratio` when the deck is in `summary` mode — a deck should not be scored for how punchy its headings are.

## Register: the slide is not the transcript

A deck fails academically before it fails visually. Two community skills — `zLanqing/codex-claude-academic-skills` (6.0k★) and `AIScientists-Dev/academic-humanizer` (5.3k★) — converge on the same rule: **claims carry evidence, and hype is a defect**. Their guidance is rewritten for Chinese decks in `references/academic-register.md`.

Non-negotiables for text that appears on a slide:

- **Review structure is five moves**: question and scope (with the search limits stated) → per-route survey with representative works and their limits → comparison on shared dimensions → consensus and disagreement → gap and next step. A review that lists papers without judging them is a bibliography.
- **No spoken-language markers.** 各位老师, 大家好, 其实, 就是说, 挺好的 belong in `speech`, never in a text box.
- **No significance hype.** 重大突破, 开辟新途径, 填补空白, 具有重要意义 — replace with the specific failure mode the work addresses.
- **No unattributed attribution.** 相关研究表明 with no work named is an error, not a hedge.
- **No number-free 显著.** If the sentence says a gain is significant, the number and the condition must be on the page.
- **Each page states a claim, its evidence, and its boundary.** `boundary` is what stops a deck from reading like a brochure; the schema already has the field — use it.

`scripts/lint_register.py` enforces the mechanical part:

```bash
python scripts/lint_register.py --deck deck.json --out out --strict
```

It scans visible text only (speech notes are exempt, because spoken language may be conversational), and reports the rule, the matched phrase, and the rewrite suggestion. Treat any `error` as blocking; `--strict` also fails on warnings.

A review deck is text-dense by nature, so `qa_design.py` scores lower on `dominant_object_ratio` than a research deck. That is expected; do not manufacture figures to satisfy the rubric.

## Match a measured deck, not a vibe

Before tuning a theme, measure. `scripts/analyze_reference.py` reads any .pptx and returns the numbers that actually make a deck look professional: objects per page, text boxes per page, text-area ratio, title size per inch, body size, size-step count, bold ratio, palette, fonts.

Run it on decks the audience would call good, and on your own output — same ruler, or the comparison means nothing.

Reference point from five real Chinese academic decks (126 pages, measured in `references/human-deck-profile.md`):

| | 真人中位 | 达标线 |
|---|---|---|
| 元素/页 | 15.0 | ≤ 18 |
| 文本框/页 | 8.9 | ≤ 11 |
| 文本框占版面 | 0.25 | ≤ 0.30 |
| 标题字号 / 版面宽 | 1.8–2.0 | 1.8–2.1 |
| 正文字号 | 16 pt | 16–19 pt |
| 字号档数 | 13 | ≥ 11 |

The single biggest structural gap between generated and human decks is **containers**: every panel card costs ~2.5 shapes, every diagram node ~2.5. Humans put text on the page and group it with whitespace and alignment, not with a border around each block. If a page needs more than two containers, the content is probably two pages or one table.

## Style Families

Two reference families pull in opposite directions, and both are legitimate:

- **Institutional cover templates** (school VI decks) are sparse: about 9 elements and 90 characters per page. Use them as the ceiling for *decoration*, not the target for content density.
- **Chinese defense / award decks** are dense: four cards, nested boxes, a conclusion band. Their palette is measurable and lives in a narrow range — see [references/cn-defense-style.md](references/cn-defense-style.md) before inventing a "deep blue".

Pick one family per deck and stay in it. Mixing a sparse cover template's whitespace with a defense deck's card stack is what makes output feel assembled rather than designed.

## Fixed Visual Contract

Read [references/style-contract.md](references/style-contract.md) before changing the renderer or theme.
Read [references/editable-pptx.md](references/editable-pptx.md) before changing native PowerPoint export, font mapping, or OMML formulas.

- Cover: no header bar and no mini-directory. Keep one institutional identity and one cover composition.
- The cover may be centered or editorial; choose once in the theme. The editorial variant uses a left title column and a faint watermark, but still keeps one institutional identity.
- Outline: only the five or six section titles, with no descriptions, page ranges, or mini-directory.
- Content-only mode: omit both cover and outline when explicitly requested; start directly with the first content page while keeping the fixed header, footer, identity, and the selected navigation model.
- Content pages: fixed header, fixed upper-right identity, and fixed footer rule. Navigation is one of the three explicit models above: `sidebar` (persistent mini-directory), `dividers` (standalone chapter pages, no sidebar), or `hybrid` (both, with the sidebar only on chapter-opening content pages). The renderer maps these to `layout.sidebar_mode`; `off` also widens the text column through `layout.content_no_sidebar`.
- The takeaway band is optional and should be rare. Write `takeaway` only when the page's conclusion differs from its `lead` and turns the argument; a takeaway that repeats the lead or the title is decoration.
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
- `scripts/qa_design.py` is the design gate. It scores ten structural proxies (dominant-object ratio, card-only ratio, emphasis density, assertion ratio, takeaway ratio, layout runs, image pages, characters per line, ink-coverage rhythm, chrome share) and grades the deck A–D. A deck below 85 is not finished, whatever the overflow report says. `validate_deck.py` already surfaces the failing checks as warnings.
- Treat `render_report.json` warnings as delivery blockers when they indicate overflow, missing assets, or incomplete content. Use `render_pptx_preview.py` for a real PowerPoint/LibreOffice rendering check; do not rely on the PIL preview alone.
- Use `projection` or `lecture` for teaching and `defense` for a defense-room deck. If a projected profile reports overflow, split the page or remove secondary text; do not shrink below the profile minimum to force it. Read [references/defense-readability.md](references/defense-readability.md) for the university accessibility and scientific-presentation evidence behind these sizes.
- Keep 24 pt defense body text sans-serif and bold, titles at least 32 pt, and every substantive visible run at least 18 pt. Leave footnotes, reference metadata, captions, and micro-labels regular; `text.body_bold_min_px` / `text.body_bold_max_px` define the semantic body range. In native PPTX output, pre-wrap Chinese runs to their text-box width and use 1.0 line spacing so PowerPoint does not repack measured lines; reserve a full line height for headings and micro-labels instead of shrinking below 18 pt. `validate_deck.py --profile defense` rejects themes whose type scale, font family, bold range, or paper-figure provenance violates these rules. Use `audit_pptx.py` and `render_pptx_preview.py` to check the exported PPTX, not just the PNG preview.
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

When the target environment cannot execute Python, run `scripts/export_agent_brief.py`. Give the generated `agent_brief.md` and the user's materials to that agent. The brief contains the same section flow, slide fields, fixed branding coordinates, speech rules, and the defense delivery gate in a model-readable form: bold sans-serif 24 pt body, 32–40 pt headings, 18 pt secondary chrome, and at least one provenance-tagged paper figure. A generated illustration, template image, or unrelated placeholder never satisfies the paper-figure requirement. Treat missing bold body text or a missing paper figure as an incomplete delivery.

Do not claim pixel-identical output from an image-only model. The guarantee is consistent structure and branding when the renderer is available; without the renderer, the brief guarantees content and layout intent, not pixel fidelity.

## Design Basis

See [references/research-basis.md](references/research-basis.md) for the GitHub sources and the specific rules adapted from them. Use the sources as design evidence, not as content sources.
