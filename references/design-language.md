# Design Language

This is the visual grammar for the deterministic renderer. It adapts the fixed institutional layer from the user's accepted deck and the evidence-first rules summarized in `research-basis.md`.

## Design Position

The deck should read as a research report, not a startup pitch:

- restrained paper background;
- one institutional blue family;
- dense but legible evidence;
- large, calm typography;
- fine lines and small radii;
- real figures preserved;
- every page has a claim and a boundary.

The image-generation look is useful for atmosphere and mechanism intuition. It is not a license to invent data, labels, axes, or citations.

## Fixed Layers

These layers must not drift between pages:

1. Top rule and chapter label.
2. Short page title.
3. Upper-right institution identity.
4. Left five-item mini-directory on content pages.
5. Bottom rule and takeaway band.

The cover and outline are exceptions: cover has no header or mini-directory; outline has no mini-directory.

## Type Scale

The reference canvas is 2048 × 1152. The default scale is:

| Role | Size |
|---|---:|
| Cover title | 72–82 |
| Page title | 46–52 |
| Chapter label | 28–30 |
| Lead claim | 28–32 |
| Panel heading | 24–27 |
| Body | 19–22 |
| Bullet | 18–20 |
| Source note | 13–15 |
| Metric value | 54–68 |

Do not shrink text below 14 px to rescue a crowded page. Split the page instead.

## Panels

- Use small radii, normally 6–10 px.
- Use a thin blue-gray outline and a very soft shadow.
- Panel headings may use a restrained blue gradient in the TJU theme.
- The main claim panel can use a pale blue fill; the rest should stay white.
- Do not nest a panel inside another panel unless the inner object is a formula, image, or small data strip.
- Do not repeat the same card shape six times when a hub, table, or figure would explain the relation better.

## Lead And Takeaway

The lead band answers: “What is this page trying to establish?”

The takeaway band answers one of:

- what the evidence supports;
- what it cannot support;
- what condition the audience must remember;
- what the next page will resolve.

The lead should be a complete sentence. The takeaway should not repeat the lead verbatim.

When the takeaway uses a deep-blue band, use white body text and a light accent color. A red emphasis color from the paper background is too dark on blue unless it is explicitly lightened.

The lead's `**...**` markup is red bold emphasis only. Do not add an underline to it; underlining is a separate visual instruction.

## Figures

- Place the figure on a white mount; do not put it on a colored card.
- Keep the source aspect ratio.
- Use a top-left label such as `论文原图`, `系统图`, `实测`, or `概念示意`.
- Use a top-right badge only when it clarifies evidence status.
- Keep a short caption for what the figure shows, but do not add a fake journal-style figure number unless the source has one.
- Reserve a narrow interpretation rail for two or three notes: observation, mechanism, boundary.

## Concept Visuals

Use a concept visual only when it explains a relationship that prose cannot explain faster:

- pulse train or photon events;
- folded histogram or peak;
- target/range geometry;
- timing or pipeline;
- a taxonomy hub.

Every concept visual must be labeled as a concept or mechanism schematic. It must not look like an unlabeled measured plot.

If there is no real figure and the content is mainly prose, do not invent a chart. Use a content-derived `footer` strip instead:

- `判断` for the question the evidence must answer;
- `交付` for the concrete research deliverable;
- `输出` for a pipeline step's output;
- `指标` for a metric family;
- `消融` for an evaluation condition.

The strip is structural semantics, not decoration. It may carry a short sentence and should not repeat the body verbatim.

## Formulas

Formulas should read like a paper, not like monospaced code:

- render `formula_latex` with a math typesetter when available;
- keep variables italic, function names upright, and subscripts/superscripts properly placed;
- use a light-blue formula strip with a thin border and dark blue math;
- keep a plain `formula` fallback for environments without a math renderer;
- never ask an image model to draw the formula.

The deterministic renderer uses LaTeX-style math typesetting for the visible formula and falls back to readable text if the optional math stack is unavailable.

## Color

- `blue`: structure, headings, navigation, figure labels.
- `blue_mid`: secondary structure, connectors, gradients.
- `blue_pale` / `blue_wash`: panels and quiet fills.
- `red`: one decisive number, contradiction, or limitation.
- `green`: only a supported positive result.
- `muted`: source notes, axes, secondary explanation.

Use at most two red emphasis points on a page.

## Density

A dense academic page is not a page full of tiny text. It is a page with:

- one claim;
- one dominant evidence object;
- two or three interpretation units;
- one boundary or takeaway.

If a page has more than five modules, split it. If a page has fewer than three, add evidence or merge it.

## Anti-Patterns

Do not use:

- handwritten marginalia;
- stars, sparkles, or “AI” decorations;
- glassmorphism, neon, 3D perspective cards;
- oversized empty cards;
- generic target/histogram/node-chain placeholders unrelated to the page;
- rainbow data series;
- multiple competing accent colors;
- unlabeled generated charts;
- fake citations or fake numbers;
- a thank-you slide with no conclusions.

