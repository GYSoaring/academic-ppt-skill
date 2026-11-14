# Zhejiang University Proposal Example

This example validates that the skill is not tied to Tianjin University or a single research topic.

- Topic: urban flood emergency response and traversable-path planning
- Theme: `assets/themes/zju_blue/theme.json`
- Report type: master's thesis proposal
- Structure: problem, status, objectives, route, methods, data, evaluation, timeline, references
- Visual policy: no generic mini-visuals; text-only blocks use `footer` / `output` strips
- Formula policy: every formula has a plain `formula` fallback and a `formula_latex` source

Render from the skill root:

```bash
python scripts/render_deck.py \
  --deck examples/zju_proposal/deck.json \
  --theme assets/themes/zju_blue/theme.json \
  --out examples/zju_proposal/rendered \
  --profile document
python scripts/export_editable_pptx.py \
  --deck examples/zju_proposal/deck.json \
  --theme assets/themes/zju_blue/theme.json \
  --out examples/zju_proposal/rendered \
  --pptx-name defense_deck.pptx \
  --profile document
python scripts/validate_pptx.py \
  --pptx examples/zju_proposal/rendered/defense_deck.pptx \
  --deck examples/zju_proposal/deck.json \
  --out examples/zju_proposal/rendered
python scripts/check_accessibility.py \
  --pptx examples/zju_proposal/rendered/defense_deck.pptx \
  --out examples/zju_proposal/rendered
python scripts/render_pptx_preview.py \
  --pptx examples/zju_proposal/rendered/defense_deck.pptx \
  --out examples/zju_proposal/rendered/pptx_preview \
  --compare-dir examples/zju_proposal/rendered/rendered \
  --strict
```

The example content is illustrative. It does not claim completed experiments or measured results.
