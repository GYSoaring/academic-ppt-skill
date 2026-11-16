# Defense Readability Audit

Use this reference when an existing PPTX looks too small in a defense room, or when rebuilding a deck with `--profile defense`.

## Research Basis (2026-11-16)

The 24 pt / 18 pt / 32–40 pt hierarchy is not an aesthetic guess. It triangulates university accessibility guidance and scientific-presentation guidance:

| Source | Relevant recommendation | Skill interpretation |
|---|---|---|
| University of Washington, [Presenter Resources](https://www.washington.edu/urs/presenters/presenter-resources/) | Use large sans-serif fonts and never smaller than 24 pt; use simple visuals on plain backgrounds | Defense body defaults to 24 pt; body font must resolve to a sans-serif CJK/Latin family; bold is semantic emphasis, not a wall of text |
| University of Washington, [Accessible Slides / PowerPoint](https://www.washington.edu/accessibility/checklist/powerpoint/) | 24 pt or larger for PowerPoint text | 24 pt is the body floor, not a target median that may hide small labels |
| University of Portsmouth, [Accessible presentations](https://www.port.ac.uk/study/student-life/ask-us/disability-and-neurodiversity/accessible-presentations) | At least 24 pt for text and 32 pt for headings; bold for strong emphasis; avoid long blocks and keep slides readable | Minimum visible substantive text is 18 pt only after content reduction; argument text stays at 24 pt; dense pages are split |
| Creighton University, [Presenting Your Scholarship: Visuals](https://culibraries.creighton.edu/presenting/visuals) | Use sans-serif fonts at a minimum of 24 pt | Reject Times, Cambria, Songti, SimSun, and other serif/CJK serif body faces for defense decks |
| [University of Washington, Equal Access: Universal Design of Presentations](https://www.washington.edu/accesscomputing/sites/default/files/doit-sync/files/EA_UD_Presentations.pdf) | Use at least 24 pt, simple sans-serif fonts such as Arial or Verdana | Defense policy is “split or subtract, never shrink below the floor”; `audit_pptx.py` measures every PPTX run |

The university guidance addresses low-vision and projection accessibility; the scientific-presentation guidance addresses typical conference/defense room viewing. The overlap is strong: 24 pt body, 32 pt headings, sans-serif faces, and restrained bold emphasis. This skill therefore treats those as hard `defense` profile requirements. Footnotes, captions, source lines, reference metadata, and micro-labels are secondary chrome: they remain regular weight and are not bolded wholesale, but they must still meet the 18 pt visibility floor when they are substantive.

## Target

- Body text: about 24 pt in the exported PowerPoint.
- Headings: about 32–40 pt.
- Any substantive visible text: at least 18 pt. Keep only genuinely secondary chrome below 18 pt, and never rely on such text for an argument.
- Bold 24 pt body text for emphasis and projection, but keep captions, source notes, reference metadata, and micro-labels regular. A uniform bold wall loses hierarchy.
- Use a sans-serif body family (for example Noto Sans SC / Microsoft YaHei / Arial); do not use Times New Roman, SimSun, Songti, or another serif face for defense body text.

`scripts/audit_pptx.py` reads the actual PPTX text runs, so it catches cases that the PNG preview or a body-only median can hide.

```bash
python scripts/audit_pptx.py --pptx path/to/existing.pptx --out path/to/audit
```

The report writes `pptx_readability.json` and `pptx_readability.md` and reports:

- all-run and body-candidate font medians;
- table and reference font medians;
- runs below 18 pt and 24 pt;
- overall bold ratio and the semantic body-candidate bold ratio;
- pictures, tables, and charts;
- a `rebuild`, `font_upgrade`, or `defense_ready` verdict.

`defense_ready` requires the body median to reach 24 pt, the all-run median to reach 18 pt, body runs below 24 pt ≤15%, all runs below 18 pt ≤5%, table runs below 18 pt ≤5%, explicitly secondary 18–24 pt text to stay ≤75% of all runs, and at least 90% of semantic body-candidate runs to be bold. A deck with a good body median but many unreadable labels or regular-weight body text is `font_upgrade`, not `defense_ready`. The audit treats repeated small captions, source lines, footnotes, table metadata, micro-labels, and reference entries as secondary chrome; the hard gate is that none of them fall below 18 pt and that substantive body text reaches 24 pt and is bold.

The 24 pt floor applies to semantic body text, not to every micro-label. Secondary chrome may use 18–24 pt and should stay regular weight so bold remains meaningful. The renderer and native exporter therefore distinguish body/lead/takeaway text from secondary labels, captions, footnotes, and reference metadata; they never enlarge dense micro-labels to 24 pt merely to satisfy a whole-deck average. Footer rules are redrawn after page content so a figure card cannot cover fixed chrome.

The portable brief produced by `scripts/export_agent_brief.py` carries the same contract to Doubao, image-generation models, and other environments that cannot execute Python. It states the font sizes and weights, rejects serif body faces, requires an editable text layer, and requires at least one paper-sourced figure with `image_origin: own|public|borrowed`, `source`, `image_role`, and `image_boundary`. Generated illustrations, template art, and concept diagrams do not satisfy this requirement. An external deck that has many images but no provenance-tagged paper figure, or has regular-weight body text, is incomplete.

## Rebuild Workflow

1. Audit the original PPTX and preserve its claims, evidence, boundaries, figures, and citations. Do not invent missing numbers or references.
2. Build `deck.json` using the relevant genre playbook; split dense pages rather than shrinking type.
3. For projection, use `--profile defense`; use `image_width_ratio` (or `rail_ratio`) between 0.65 and 0.80 for a hero figure, and use `span`, `footer`, `output`, or `stat` to avoid equal empty cards.
4. Export the native editable PPTX, then run `validate_deck.py`, `validate_pptx.py`, `check_accessibility.py`, and `render_pptx_preview.py`.
5. Re-run `audit_pptx.py` on the delivered PPTX. Treat `rebuild` or `font_upgrade` as a delivery blocker unless the user explicitly accepts a non-defense profile.

## Delivery Gate

`validate_deck.py --profile defense` rejects a deck when the theme does not provide the 24 pt body scale, 32 pt header scale, sans-serif body face, semantic body-bold range, 18 pt secondary-text floor, or at least one paper-sourced figure with `source`, `image_role`, and `image_boundary`. `audit_pptx.py` then checks the exported PPTX's physical runs. Both checks must pass before calling a defense deck deliverable.

## Visual Density

`validate_deck.py` warns when a `lecture` / `defense` content page has little text and no visual anchor, repeats the same layout four or more times, or has no hero figure, chart, or formula page. These are judgment signals, not pixel-perfect beauty scores. Fix them by adding real evidence or varying composition; do not add decorative icons or filler text.
