# Genre Playbooks

The deck is not a fixed sequence of layouts. Pick a genre, then choose the narrative and the evidence object for each claim. These playbooks are defaults, not templates to fill blindly.

## Why Genre Matters

Gomez et al. compared slides across the humanities, social sciences, natural sciences, and formal sciences. They found large differences in visual conventions:

- humanities slides used images far more often and bullet text far less often;
- formal and social science slides used more bullet text and more systematic titles;
- a title on every slide was common in formal/natural/social sciences, but uncommon in the humanities.

So a humanities deck that forces every page into a three-card grid is not "consistent"; it is the wrong genre. A theory deck that avoids formulas and proof sketches is also the wrong genre.

The narrative can still be rigorous. The research on scientific storytelling recommends a clear purpose, an early central claim, and a path from context to findings to implications. Common arcs include:

- `problem-method-result`: gap → approach → evidence → boundary;
- `and-but-therefore`: accepted idea → complication → resolution;
- `result-first`: headline result → how it was obtained → why it is credible;
- `close-reading`: passage/artifact → observation → counter-reading → interpretation;
- `taxonomy`: landscape → axes → groups → gaps → agenda;
- `design-rationale`: requirements → design space → artifact → evaluation → reflection.

## Genre: Experiment

Use for hypothesis-driven natural science, engineering experiments, and measurement papers.

- Narrative: `problem-method-result` or `and-but-therefore`.
- Typical order: question → gap → hypothesis → setup → result → control/ablation → boundary.
- Evidence: measured figures, metrics, comparison tables, error bars, conditions.
- Prefer: `figure`, `figure_flow`, `metrics`, `compare`.
- Avoid: long text pages that never show a measured object.
- Gate: every result page needs sample size, condition, unit, and source.

## Genre: Systems

Use for computer systems, robotics systems, platforms, and infrastructure papers.

- Narrative: `design-rationale` plus `problem-method-result`.
- Typical order: workload/problem → constraints → architecture → implementation → evaluation → limitations.
- Evidence: architecture diagram, data flow, pipeline, latency/throughput table, failure analysis.
- Prefer: `pipeline`, `figure_flow`, `concept` (stack/sequence), `balance`, `compare`.
- Avoid: showing only a final benchmark without the system boundary.
- Gate: the architecture page must state what is inside the system and what is external.

## Genre: Theory

Use for formal sciences, algorithms, mathematics, and theory-driven CS.

- Narrative: `problem-method-result` with a proof-oriented middle.
- Typical order: question → definitions → assumptions → theorem/proof sketch → implications → boundary.
- Evidence: formula, lemma/proof sketch, counterexample, complexity or trade-off comparison.
- Prefer: `formula`, `concept` (matrix/stack), `compare`.
- Avoid: decorative screenshots and unsupported empirical claims.
- Gate: state assumptions before the theorem; separate proof sketch from measured evidence.

## Genre: Empirical / Social Science

Use for economics, management, sociology, education, and other empirical work.

- Narrative: `problem-method-result`, often with a `result-first` opening.
- Typical order: question → theory → data → identification → results → robustness → limits.
- Evidence: sample construction, identification strategy, coefficient tables, robustness checks.
- Prefer: `compare`, `metrics`, `figure`, `timeline` for data construction.
- Avoid: causal language without an identification argument.
- Gate: every estimate needs sample, specification, and uncertainty.

## Genre: Survey

Use for literature reviews, benchmarks, and field overviews.

- Narrative: `taxonomy` or `and-but-therefore`.
- Typical order: why now → axes → taxonomy → representative work → gaps → agenda.
- Evidence: taxonomy/hub, comparison matrix, timeline, coverage table.
- Prefer: `hub`, `compare`, `timeline`, `grid`.
- Avoid: claiming exhaustive coverage or ranking papers across incomparable settings.
- Gate: state the search boundary and what the taxonomy omits.

## Genre: Humanities

Use for literature, history, philosophy, art history, religion, and cultural studies.

- Narrative: `close-reading` or `and-but-therefore`.
- Typical order: question → artifact/passage → observation → counter-reading → interpretation → stakes.
- Evidence: quotation, image/plate, manuscript detail, historical sequence, close-reading annotation.
- Prefer: `quote`, `figure`, `timeline`, `bullets` used sparingly.
- Avoid: forcing every page into a three-card grid; over-titling can flatten the argument.
- Gate: quotations need attribution and edition; images need source and what to look at.

## Genre: Design / HCI

Use for design research, HCI, architecture, and artifact-driven work.

- Narrative: `design-rationale`.
- Typical order: problem → requirements → design space → artifact → study → reflection.
- Evidence: design space matrix, prototype, interaction flow, qualitative/quantitative study.
- Prefer: `concept` (matrix), `figure_flow`, `pipeline`, `balance`.
- Avoid: presenting the artifact without the design rationale or evaluation.
- Gate: separate "what was built" from "what was learned".

## Genre: Clinical / Biomedical

Use for clinical trials, biomedical studies, and translational work.

- Narrative: `problem-method-result`.
- Typical order: burden → question/PICO → methods → results → safety → limitations.
- Evidence: participant flow, endpoint table, survival/forest plot, adverse events.
- Prefer: `figure`, `compare`, `metrics`, `timeline`.
- Avoid: efficacy claims without safety and eligibility boundaries.
- Gate: report population, endpoint definition, and follow-up.

## Layout Variants

Do not create a new layout for every visual difference. Prefer a layout plus a variant.

| Layout | Variants | Use when |
|---|---|---|
| `figure` | `image_left`, `image_right`, `chart` | one dominant visual with interpretation |
| `bullets` | one/two/three columns | short parallel points |
| `compare` | `table` (current) | several routes under shared dimensions |
| `metrics` | `cards` (current) | a few measured outcomes with conditions |
| `timeline` | `horizontal` (current) | phases and milestones |
| `quote` | `centered`, `split` (current: centered) | humanities close reading |

The model chooses the variant from the content shape, not from a fixed page number. A page's layout should follow its claim and evidence, not the other way around.

## Sources

- Gomez et al., *Different Strokes for Different Folks: Visual Presentation Design between Disciplines*, 2016. https://pmc.ncbi.nlm.nih.gov/articles/PMC4729220/
- Alley, *Guidelines for Assertion–Evidence Slides*, Penn State. https://www.writing.engr.psu.edu/guidelines_AE_slides.pdf
- Kulasegaram et al., *Taking presentations seriously: Invoking narrative craft in academic talks*, 2017. https://pmc.ncbi.nlm.nih.gov/articles/PMC5542897/
- Cagan, *Research story telling: using the research journey map*, 2024. https://doi.org/10.1017/pds.2024.236
- Tahta design system: layouts and variants. https://tahta.cagdas.io/
