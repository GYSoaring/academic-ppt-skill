# Visual Language

This document records the visual grammar the renderer is trying to reproduce. It is a design contract, not a claim that a generated image can be recreated pixel for pixel.

The style base comes from two places:

1. The accepted academic deck images produced with an image-generation workflow. Those images provide atmosphere, density, fixed-page rhythm, and evidence-card language.
2. Public academic presentation guidance and layout systems. Those sources provide the reusable rules: assertion, evidence, hierarchy, white space, alignment, reading order, and restrained color.

The renderer keeps the rules and redraws the page deterministically. It never traces a generated slide as a full-page picture.

## Design Position

The deck is a research report, not a commercial pitch.

- It should look calm, edited, and evidence-led.
- It should carry more information than a keynote, but every element must have an argumentative job.
- A reader should be able to recover the page's claim, evidence, and boundary without hearing the speaker.
- The visual language should survive a change of school, discipline, and topic.

## Fixed Zones

The content page uses five zones. Their position is owned by the theme, not by the content model.

| Zone | Job | Rule |
|---|---|---|
| Header | Orientation | Chapter number and short topic title share one measured size and baseline; upper-right identity is fixed |
| Lead band | Claim | One complete sentence; emphasis is sparse |
| Evidence field | Main object | Figure, table, chart, diagram, formula, or a structured evidence grid |
| Interpretation rail | Reading help | Two or three panels: observation, mechanism, boundary |
| Takeaway band | Closure | A condition, consequence, or transition; not a restatement of the lead |

The footer rule, page number, identity, and mini-directory are chrome. They are not content and must not move to make room for content.

## Canvas And Grid

- Reference canvas: `2048 × 1152`, 16:9.
- Outer margin: approximately `36 px` on the sides.
- Header reserve: the first `112–118 px` belongs to the fixed header.
- Content top: the lead band begins below the header, never underneath it.
- Sidebar reserve: the five-item navigation occupies its own fixed column and is never overlapped.
- Use a 12-column conceptual grid inside the content area. A major evidence object may span 7–9 columns; an interpretation rail normally spans 3–4 columns.
- Use a consistent spacing scale: group-internal `8–12 px`, between groups `18–24 px`, between major regions `28–40 px`.
- Align to edges before centering. A centered object is a deliberate exception, not the default.

## Reading Order

Every page needs an obvious path:

- horizontal for a process or a comparison;
- vertical for a hierarchy or a claim-to-evidence argument;
- left-to-right for a causal chain;
- center-to-outside only for a genuine taxonomy or shared bottleneck.

Use arrows only when direction is part of the meaning. Decorative arrows, chevrons between unrelated cards, and ornamental connectors are removed.

## Evidence Layouts

Choose the layout from the evidence relation, not from visual habit.

| Relation | Layout | What dominates |
|---|---|---|
| One paper figure or measured result | `figure` | The untouched figure, with a narrow interpretation rail |
| Mechanism plus one real figure | `figure_flow` | A short causal chain above the evidence |
| Several parallel mechanisms | `grid` | Three to six compact, non-empty evidence blocks |
| One system with benefits and costs | `balance` | The shared system object in the center |
| A real sequence | `pipeline` | Ordered stages with inputs and outputs |
| A real comparison | `compare` | A native table; no decorative ranking |
| Measured outcomes with conditions | `metrics` | Two to four headline metrics plus a chart or figure |
| A taxonomy or route panorama | `hub` | A real shared center and four to six parallel branches |
| A research schedule | `timeline` | Dates, work packages, and milestones |
| A concept with a causal or structural relation | `concept` | A semantic diagram plus two or three reading panels |
| A central equation with definitions | `formula` | The equation, its symbols, its condition, and its limits |

If no real image exists, use a content-derived `footer` / `output` strip, a formula, a table, or a semantic diagram. Do not insert a target, histogram, node chain, or noise plot merely because a card looks empty.

## Semantic Diagrams

`concept` diagrams are data-driven. The renderer owns geometry; `deck.json` owns every label and value.

| Kind | Use when | Shape |
|---|---|---|
| `sequence` | A real temporal or causal order | 2–5 numbered nodes, horizontal |
| `branch` | One shared input producing several routes | A root object and 2–4 ordered branches |
| `stack` | A hierarchy, nesting, or layered model | 2–5 indented rows |
| `ruler` | A scale, interval, threshold, or timeline | 2–8 positioned ticks with optional pairs |
| `matrix` | A two-dimensional tradeoff | Exactly four quadrants with labeled axes |
| `cycle` | A closed feedback or iteration | 3–5 nodes on a ring |

Rules:

- The diagram must be readable without the speaker.
- Every node needs a content-bearing title; body text is optional but should explain the relation.
- The highlight is one node at most.
- A `ruler` must use numeric `value` fields; do not position ticks by prose.
- A `matrix` must label both axes and all four quadrants.
- A diagram is a concept schematic. If the page contains real measured evidence, use `figure` or `figure_flow` instead.
- The native PPTX renderer uses shapes and text boxes for the same diagram. It does not flatten the diagram into an image.

## Formula Pages

The formula page has four layers:

1. A short title and a complete lead claim.
2. A large equation panel with a label.
3. A condition strip that states when the equation is interpretable.
4. Symbol definitions and two or three interpretation panels.

Use `formula_latex` for the typeset source and `formula` as the portable fallback. The preview uses LaTeX-style math typesetting; the editable PPTX uses native OMML when available and an editable Cambria Math fallback otherwise.

Typography rules:

- variables italic;
- function names upright;
- true subscripts and superscripts;
- `\sum`, `\left...\right`, `\varphi`, and `\mathcal` are preferred over plain-text approximations;
- symbol rows use math symbols, not code-like strings such as `t_e`.

Do not let an image model draw the equation. Do not add a numerical value to a symbol unless the source supplies it.

## Typography

- Header topic title: normally 4–10 Chinese characters; it is orientation, not the claim.
- Chapter number and adjacent title use the same measured font size and baseline.
- Lead claim: 28–33 px on the reference canvas.
- Body: 18–22 px; do not shrink below 14 px to rescue a crowded page.
- Panel headings: 24–27 px; use a compact dark-blue strip or a clear top label.
- Source notes and conditions: 13–15 px, muted but still contrast-checked.
- Cover title may use a display serif; content titles and body use a CJK sans.
- Latin and mathematical characters use the theme's Latin or formula family, not the CJK body font.

The 2048 × 1152 canvas maps to a 13.333-inch PowerPoint slide at 153.6 px/in. That means `1 px ≈ 0.469 pt`: a body value of 22 px is only about 10.3 pt in PowerPoint. Do not confuse the preview pixel size with the physical point size in the exported deck.

Use `--profile projection` for a lecture-room deck. It raises the physical type scale and should surface overflow rather than silently shrink text. Use `--profile document` for a dense handout and state that distinction in the delivery notes.

## Color And Contrast

Use one institutional blue family for structure:

- deep blue for headers, navigation, and primary structure;
- mid blue for connectors and secondary emphasis;
- pale blue for quiet panels;
- paper white-blue for the field.

Red is reserved for one decisive number, contradiction, risk, or boundary. Do not use it as a decorative accent. Green is used only when a result is actually supported by data.

Contrast is non-negotiable:

- on a deep-blue band, every token — body, bold, red emphasis, muted note — must switch to a light foreground;
- thin rules must remain visible on the paper background;
- muted text must not fall below readable contrast;
- do not place dark blue text on a mid-blue fill.

`**text**` means red bold emphasis only. It must not add an underline. Underlining is a separate instruction and is rarely appropriate in an academic deck.

## Density And White Space

- Aim for three to five content units per content page.
- One unit is the dominant evidence object; the other units explain, compare, or bound it.
- Do not fill every card with the same amount of text. Give the main evidence more space.
- Keep a visible quiet margin around the evidence object.
- A page with one large empty card is worse than a page with a smaller, content-led structure.
- A page with six equal cards usually means the relation has not been analyzed yet; use `hub`, `pipeline`, or `compare` instead.

## Image-Generation Boundary

The image-generation style is useful for:

- paper-like atmosphere;
- concept intuition;
- restrained scientific illustration;
- a consistent visual mood.

It is not allowed to carry:

- exact Chinese titles or school names;
- seals or official identity;
- axes, measured curves, or experimental numbers;
- citations, figure numbers, or source marks;
- formulas;
- claims that the user has not supplied.

The correct production sequence is:

1. Generate or select a visual reference for mood and composition.
2. Extract its layout grammar.
3. Encode the grammar as a theme and layout contract.
4. Render text, formulas, tables, diagrams, and figures deterministically.
5. Export native PowerPoint objects.

The guarantee is structural and editorial fidelity, not pixel-identical reproduction of an image model's output.

## Sources Behind The Rules

The following public sources were used as design evidence. No slide, logo, illustration, or paper content is copied from them.

- Harvard Catalyst, *Slides*: assertion-evidence, audience-centered visuals, and slide assessment. <https://catalyst.harvard.edu/writing-communication-center/visualize-science/slides/>
- University of Washington Design Help Desk, *The Little Book of Visual Design Principles for Scientists*: contrast, space, micro-space, unity, and grid. <https://depts.washington.edu/deshelp/toolkit/DesignGuide.pdf>
- OpenPresentation layout reference: regions in reading order, layout as a relationship, and content flowing into a declared structure. <https://www.openpresentation.org/docs/reference/layout-templates>
- OpenPresentation format card: one slide, one message; explicit layout and design tokens. <https://www.openpresentation.org/docs/reference/format-card>
- GitHub academic and medical slide systems, including `Aperivue/medsci-skills` and `Russell-cell/PPT-Design-Prompt`: restrained clinical palettes, hierarchy, evidence plates, and image/text separation. <https://github.com/Aperivue/medsci-skills> and <https://github.com/Russell-cell/PPT-Design-Prompt>

## Acceptance Checklist

Before shipping a deck, verify:

- [ ] Header number and title have the same measured size.
- [ ] Upper-right identity is identical on every content page.
- [ ] Cover and outline have no mini-directory.
- [ ] Every content page has a lead claim and a takeaway or boundary.
- [ ] No deep-blue band contains dark text.
- [ ] No `**emphasis**` is rendered with an underline.
- [ ] No figure is stretched, cropped, or redrawn.
- [ ] No concept diagram is presented as measured evidence.
- [ ] Formula variables are typeset as math, not code.
- [ ] The editable PPTX contains text, shapes, tables, charts, and OMML formulas rather than a full-page raster.
