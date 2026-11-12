# Research Basis

Research date: 2026-11-12. The goal was not to copy a GitHub theme, but to extract repeatable rules for academic hierarchy, evidence density, typography, formula rendering, and image-generation boundaries.

## Search Scope

Queries included:

- `academic presentation template`
- `academic ppt pptx`
- `beamer academic theme`
- `slidev academic theme`
- `marp academic theme`
- `image generation presentation slides`

Selection favored repositories with an actual slide system, not a single decorative template.

## Sources And What They Contribute

| Repository | Observed idea | What we adopted |
|---|---|---|
| [Gabberflast/academic-pptx-skill](https://github.com/Gabberflast/academic-pptx-skill) | Action titles, ghost-deck test, one exhibit per results page, citation discipline, communication-first design | Keep the fixed topic title short; put the full claim in the lead band; end with conclusions and boundaries rather than a blank thank-you |
| [M1n-n9/academic-ppt-master](https://github.com/M1n-n9/academic-ppt-master) | Five-beat academic structure, dense/breathing rhythm, figure-first evidence, generated visuals with `text_policy: none` | Five-module defense flow; alternate dense and quieter pages; never let generated images carry exact labels or numbers |
| [wycmochi/cn-academic-spark](https://github.com/wycmochi/cn-academic-spark) | Academic layout library, semantic colors, hero figure 65%–80%, formula blocks, technical-route contract | Layout selection by evidence role; one accent family; figure-dominant pages; formula/mechanism panels; hub layout for route taxonomy |
| [fangyuanopus/literature-report-ppt-builder](https://github.com/fangyuanopus/literature-report-ppt-builder) | Style contract, batch image consistency, real-figure policy, short evidence annotations | `image_role` / `image_boundary`; one deck-wide visual contract; annotate only the decisive region; do not redraw measured evidence |
| [jxpeng98/slidev-theme-scholarly](https://github.com/jxpeng98/slidev-theme-scholarly) | Academic layouts for paper summary, methodology, results, limitations, defense questions | More layout variety; a closing page with boundary and next steps; separate mechanism from evidence |
| [kaisugi/marp-theme-academic](https://github.com/kaisugi/marp-theme-academic) | White field, one restrained accent, simple sans typography | Paper-like background; limited accent area; no rainbow palette |
| [JuneYaooo/gpt-image2-ppt-skills](https://github.com/JuneYaooo/gpt-image2-ppt-skills) | Thesis-defense style with a persistent thin-line grid, serif display titles, sans body, square number frames, line icons, structured layout JSON | Serif display for the cover, sans for content; small corner radii; persistent grid/rhythm; prompt structure that separates style from semantic content |
| [pmichaillat/latex-presentation](https://github.com/pmichaillat/latex-presentation) | Minimalist Beamer typography, whitespace, one idea per slide | Calmer hierarchy; fewer decorative objects; explicit whitespace around the evidence |
| [VectorSpaceLab/InfoVis](https://github.com/VectorSpaceLab/InfoVis) | Slide-image generation and editing benchmark | Confirms the separation between visual generation and exact layout/text; we keep text, axes, citations, and page chrome deterministic |
| [KaTeX/KaTeX](https://github.com/KaTeX/KaTeX) | Fast LaTeX-style math typesetting with standard operator and script spacing | Treat formula typography as its own layer rather than ordinary body text; keep math italic, upright function names, and true sub/superscripts |
| [matplotlib/matplotlib](https://github.com/matplotlib/matplotlib) | Offline mathtext renderer with STIX/Computer Modern-compatible math fonts | Render `formula_latex` to a transparent high-resolution asset when a browser or Node math renderer is unavailable; fall back to plain text rather than failing |

## Rules Adopted

1. **Short title, full claim.** The fixed header is for orientation; the lead band carries the sentence the audience should remember.
2. **One page, one job.** Do not mix a result, a method description, and a limitation on one page unless they are the same argument.
3. **One dominant exhibit.** A real figure or table should be the visual center; cards support it rather than compete with it.
4. **Density with air.** Three to five modules per content page, with one clear reading order.
5. **Figures before decoration.** A decorative icon never substitutes for evidence.
6. **Color is semantic.** Blue for structure, red for a decisive number or boundary, gray for source notes.
7. **Generated images are concept illustrations.** They may explain a mechanism but cannot carry exact Chinese labels, numbers, axes, or citations.
8. **Close on the argument.** The last page should still be useful during Q&A: conclusions, boundary, next step.
9. **Semantic fallback before decoration.** If the user supplies no image, a `footer`/`output` judgement or a textual structure is better than an unrelated target, histogram, or node chain.
10. **Math is not body text.** Formula variables, operators, and scripts need LaTeX-style typesetting; keep a plain fallback for portability.

## Deliberately Not Copied

- Generic corporate gradients, glowing cards, and 3D dashboard widgets.
- One-off hero illustrations that vary in style from page to page.
- Fake charts, invented numbers, or synthetic “measured” curves.
- A single fixed card grid for every content type.
- An image model rendering exact Chinese titles or institutional identity.

## Implementation Map

| Rule | Where it lives |
|---|---|
| Short header title + lead claim | `scripts/render_deck_next.py` → `draw_header`, `draw_lead` |
| Hero figure proportion | `draw_figure` uses a 66% image column |
| Two-rail tradeoff around one system figure | `draw_balance`, `draw_balance_rail` |
| Layout selection | `references/layout-catalog.md` |
| Visual anchor and density checks | `scripts/validate_deck.py` |
| Hub / route panorama | `draw_hub`, `draw_hub_card` |
| Closing boundary and next steps | `draw_closing` |
| Concept-image boundary | `references/style-contract.md` |
| Portable model brief | `scripts/export_agent_brief.py` |
| LaTeX-style formula rendering | `scripts/design.py` → `render_math_image`, `draw_grid_block` |
| Text-only fallback strips | `draw_output_strip`; `footer` / `output` fields in the deck schema |

## Limits

Stars and search rank are not quality guarantees. These repositories were used as design evidence only; no source slide, logo, illustration, or paper content is copied into the deck. The deterministic renderer cannot reproduce a custom 3D illustration exactly, but it can reproduce the academic hierarchy, density, fixed chrome, and evidence discipline.

