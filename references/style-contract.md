# Fixed Visual Contract

## Canvas

- Default canvas: `2048 x 1152`, 16:9.
- PPTX export: `13.333333 x 7.5` inches.
- Paper color: pale blue-white, never pure gray.
- Main ink: dark navy; body text: near-black blue-gray.
- Rules: thin blue lines, not black borders.

## Editorial Hierarchy

1. Fixed chapter label and short topic title: orientation, not the argument.
2. Lead band: one complete-sentence claim that tells the audience what this page proves.
3. Evidence object: figure, table, flow, metric, formula, or labeled concept schematic.
4. Interpretation modules: two to four short notes that explain what to look at.
5. Takeaway band: the condition, boundary, or transition that closes the page.

Do not use the fixed header for a long sentence. Header titles should normally be 4–10 Chinese characters; the action claim belongs in `lead`.

## Page Rhythm

- Alternate dense evidence pages with quieter statement or section pages. Do not make ten consecutive pages with the same card rhythm.
- A content page should normally carry 3–5 content units. If it has only one or two, merge it with an adjacent page or split the claim and give each part its own evidence.
- Each page should have one dominant visual anchor. Decorative icons do not count as evidence.
- Use `figure` for one real figure, `figure_flow` for a process plus its evidence, `pipeline` for a true sequence, `grid` for several independent mechanisms, `balance` for two opposing rails around one central system figure, `compare` for a real comparison, `metrics` for measured outcomes with conditions, and `hub` for a taxonomy or route panorama.

## Typography

- Chinese headings: bold sans or display serif according to the theme; body: regular CJK sans.
- Formulas and Latin symbols: serif or the theme's formula font; do not set a whole Chinese paragraph in a math font.
- Body text should normally be 18–22 px at the 2048 × 1152 canvas. Do not shrink below 14 px to force more text onto a page.
- Use red bold text for at most two emphasis points per page. Do not underline every claim.
- `**text**` means red bold emphasis only; it must not draw an underline. Use `__text__` only for a deliberate underline.
- Keep line length readable: 22–36 Chinese characters per line in a panel; break long claims into a lead plus two or three short notes.
- Formulas use math typesetting: italic variables, upright function names, proper superscripts, fractions, sums, and operator spacing. Store the LaTeX source in `formula_latex` and a readable fallback in `formula`.

## Color Semantics

- Use one primary blue family for ordinary structure.
- Use red only for a decisive result, contradiction, risk, or evidence boundary.
- Use green only for a positive direction that is actually supported by data.
- Use muted gray for source notes, inactive navigation, axes, and secondary labels.
- Same-family data series should vary in lightness or opacity, not hue. Do not use a rainbow palette.
- On a deep-blue or otherwise dark band, every text token must map to a light foreground token. Never leave body, emphasis, or muted text at its light-background color.

## Cover

- No top header bar.
- No left mini-directory.
- Use one cover composition chosen by the theme: centered or editorial. Do not mix both on the same page or vary the composition between pages.
- Centered cover: institutional identity is horizontally centered in the upper third; title and metadata are centered below it.
- Editorial cover: identity stays in the upper-left, the title occupies a left column, and a faint institutional watermark may occupy the right column.
- Identity order: circular seal at left, handwritten Chinese school name at right, English name below the Chinese name.
- A faint blue arc or soft band may anchor the lower edge.

## Outline

- No left mini-directory.
- Show only the section numbers and section titles. The row number and section title use the same measured font size; do not make the number visually heavier or larger than the title it labels.
- Do not add descriptions, scope notes, page ranges, or decorative arrows containing text.
- Each row must visually connect to the same section names used by the mini-directory.

## Content Header

- Top `112 px` is reserved for fixed identity.
- Left double slash, chapter label, page title, and thin divider are locked.
- Upper-right identity is locked to the theme `brand.header_box`.
- The title baseline and size must not change between pages. The chapter number and topic title use the same measured font size and baseline; if a long title shrinks, its number shrinks with it.
- Do not paste the old identity into a new page or leave two identities visible.

## Mini-Directory

- Appears from the third page onward.
- Cover and outline never show it.
- Five items, one active section, fixed width and position.
- Active state: deep blue fill with white text.
- Inactive state: white or pale fill with thin blue outline and deep blue text.
- Never move the sidebar into the body to make room for content.

## Body Layout

- Leave generous paper margins.
- Use a light-blue lead band for the page's question or claim.
- Use two or three asymmetric panels rather than a uniform grid of cards.
- When a source page contains several small mechanisms, use the `grid` layout: two or three columns, three to six compact blocks, and a short conclusion band.
- Keep block headings in a deep-blue strip and body text below; avoid one large empty card holding only two short bullets.
- If a block has no real image, use `footer` or `output` for a short content-derived judgement, deliverable, or metric. Do not place an unrelated histogram, target, node chain, or timeline in the empty space.
- When the block's content is text-only, prefer a tighter card, a wider span, or a footer strip over a decorative mini-visual.
- On a `balance` page, use three large icon-led rows per rail and a two- or three-line body. Do not turn the rail into a small bullet list beside a large figure.
- Prefer thin rules and restrained fills over heavy shadows.
- Use a bottom takeaway band only when it adds a conclusion or boundary.
- Avoid nested cards, large rounded corners, saturated multi-color blocks, and dashboard-style widgets.
- When a real figure is the main evidence, let it occupy roughly 65%–80% of the content area. Put the interpretation in a narrow rail; do not letterbox a dense plot into a small decorative card.
- Use a hub layout only when the center is a real shared object or question. Four to six surrounding branches should be parallel in rhetorical level; do not use a hub as decoration.
- Use `concept` only when the relation is genuinely spatial, causal, structural, or cyclical. The diagram must have a labeled kind, content-bearing nodes, and a note stating that it is not measured evidence.
- A `sequence` is not a decorative arrow chain: it must encode a real order. A `matrix` must have two named axes and four meaningful quadrants. A `ruler` must use numeric values from the source.
- Use `formula` when one equation is the page's central object. The page must include the equation, symbol definitions, a condition strip, and a boundary panel; a formula without these is not a complete evidence page.

## Figures

- Preserve the source aspect ratio with `contain`, never stretch or crop measured plots.
- Use a thin neutral frame and a white interior.
- Do not add `论文 Fig. x` captions below figures unless the user explicitly asks.
- Do not redraw axes, curves, legends, photographs, or hardware diagrams.
- Record the figure's role in `image_role` and its evidentiary limitation in `image_boundary`.
- Keep the figure's original aspect ratio. A dense multi-panel figure that is unreadable at slide scale should be split across pages, not shrunk.
- Annotate only the one or two regions that support the page claim. Do not cover axes, legends, scale bars, or data points.
- When `chart` data is supplied, prefer a native editable chart over a screenshot. Keep the same card, palette, highlight, axis labels, caption, and annotation as the preview; keep numerical values in the chart data, not in decorative text.

## Generated Visuals

- Use generated images for atmosphere, mechanism intuition, or concept illustration only.
- Keep exact Chinese labels, formulas, numbers, axes, citations, and source marks in the deterministic layer.
- A generated concept image must be labeled `概念示意` or `机制示意`; it must never be described as a measured result.
- Use one visual style and one palette across the deck. Do not switch illustration styles from page to page.
- Never use a generic placeholder chart simply because a card looks empty. Empty space is acceptable; unrelated evidence is not.

## Formulas

- The preferred path is LaTeX-style math rendering from `formula_latex`; the plain `formula` string is the portable fallback.
- Variables are italic; function names and multi-letter labels use upright math roman (`\mathrm{...}` or `\operatorname{...}`).
- Use `\left...\right`, `\sum`, `\varphi`, `\mathcal`, and braces deliberately; do not fake superscripts with plain text.
- If the math renderer is unavailable, keep the formula readable as plain text and emit a render warning rather than failing the deck.
- Do not use an image model to generate formulas, axes, or symbols. They belong in the deterministic layer.
- Use actual subscript characters or LaTeX math in symbol rows; `t_e` is a code identifier, not a visible academic symbol.

## Image-Generation Extraction

The accepted Image Gen 2.6 academic deck images are a style reference, not a layout source to trace. Extract:

- a fixed header with a double slash, chapter number, and short topic title;
- a fixed upper-right institutional identity;
- a five-item mini-directory that is present on content pages and absent on cover/outline;
- dense evidence pages with a dominant figure, one narrow interpretation rail, and a bottom boundary band;
- deep-blue module headings, pale blue evidence fills, thin lines, and a single red accent;
- a consistent rhythm between dense evidence pages and quieter transition pages.

Do not extract or copy:

- exact Chinese titles, seals, school names, photographs, or paper figures from a generated image;
- a full-page generated slide as a background;
- fake axes, fake numbers, fake references, or an image-model formula.

The deterministic renderer must redraw the structure and keep every text object editable. The guarantee is editorial and structural fidelity, not pixel-identical reproduction.

## Identity Assets

- Use only user-provided or institution-approved identity assets.
- A generic theme may use text or a neutral monogram, but must not fabricate an official seal.
- If no approved identity asset exists, use a text lockup or the theme's neutral monogram and state that the official mark is not supplied.
- If only a dark-background lockup exists, recolor a transparent alpha asset rather than placing a dark rectangle on the slide.
- Keep cover and content identities from the same asset family.
