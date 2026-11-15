# Defense Readability Audit

Use this reference when an existing PPTX looks too small in a defense room, or when rebuilding a deck with `--profile defense`.

## Target

- Body text: about 24 pt in the exported PowerPoint.
- Headings: about 32–40 pt.
- Any substantive visible text: at least 18 pt. Keep only genuinely secondary chrome below 18 pt, and never rely on such text for an argument.
- Bold 24 pt body text for emphasis and projection, but keep captions, source notes, reference metadata, and micro-labels regular. A uniform bold wall loses hierarchy.

`scripts/audit_pptx.py` reads the actual PPTX text runs, so it catches cases that the PNG preview or a body-only median can hide.

```bash
python scripts/audit_pptx.py --pptx path/to/existing.pptx --out path/to/audit
```

The report writes `pptx_readability.json` and `pptx_readability.md` and reports:

- all-run and body-candidate font medians;
- table and reference font medians;
- runs below 18 pt and 24 pt;
- bold ratio;
- pictures, tables, and charts;
- a `rebuild`, `font_upgrade`, or `defense_ready` verdict.

`defense_ready` requires the body median to reach 24 pt, the all-run median to reach 18 pt, body runs below 24 pt ≤15%, all runs below 18 pt ≤5%, all runs below 24 pt ≤25%, and table runs below 18 pt ≤5%. A deck with a good body median but many unreadable labels is `font_upgrade`, not `defense_ready`.

## Rebuild Workflow

1. Audit the original PPTX and preserve its claims, evidence, boundaries, figures, and citations. Do not invent missing numbers or references.
2. Build `deck.json` using the relevant genre playbook; split dense pages rather than shrinking type.
3. For projection, use `--profile defense`; use `image_width_ratio` (or `rail_ratio`) between 0.65 and 0.80 for a hero figure, and use `span`, `footer`, `output`, or `stat` to avoid equal empty cards.
4. Export the native editable PPTX, then run `validate_deck.py`, `validate_pptx.py`, `check_accessibility.py`, and `render_pptx_preview.py`.
5. Re-run `audit_pptx.py` on the delivered PPTX. Treat `rebuild` or `font_upgrade` as a delivery blocker unless the user explicitly accepts a non-defense profile.

## Visual Density

`validate_deck.py` warns when a `lecture` / `defense` content page has little text and no visual anchor, repeats the same layout four or more times, or has no hero figure, chart, or formula page. These are judgment signals, not pixel-perfect beauty scores. Fix them by adding real evidence or varying composition; do not add decorative icons or filler text.
