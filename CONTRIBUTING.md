# Contributing

## Before You Change Layout

- Read `references/style-contract.md` and `references/visual-language.md`.
- Keep content in `deck.json`; the renderer owns geometry.
- Do not invent references, measured values, or experimental conditions.
- Do not turn a full-page screenshot into the editable deliverable.

## Local Checks

```bash
python -m compileall -q scripts
python -m pytest -q
python scripts/render_deck.py --deck examples/zju_proposal/deck.json --theme assets/themes/zju_blue/theme.json --out work/check
python scripts/export_editable_pptx.py --deck examples/zju_proposal/deck.json --theme assets/themes/zju_blue/theme.json --out work/check --pptx-name check.pptx
python scripts/validate_deck.py --deck examples/zju_proposal/deck.json --theme assets/themes/zju_blue/theme.json --render-dir work/check/rendered --out work/check
python scripts/validate_pptx.py --pptx work/check/check.pptx --deck examples/zju_proposal/deck.json --out work/check
```

Use `--profile projection` when changing type scales; the deck must either fit or fail with an explicit overflow warning.

## Adding A Layout

1. Add the layout name to `SUPPORTED_LAYOUTS` in `scripts/validate_deck.py`.
2. Implement the PIL path in `scripts/render_deck_next.py`.
3. Implement the native PPTX path in `scripts/export_editable_pptx.py`.
4. Document the schema in `references/deck-schema.md`.
5. Add a content example and a unit test.
6. Run both validation scripts and inspect the contact sheet.

## Pull Requests

- One focused change per pull request.
- Include the commands you ran and their results.
- Note any new dependency, font, or Office requirement.
