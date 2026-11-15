# Figure Sourcing

Use the most direct evidence available, in this order:

1. figures from the user's own paper or source pack;
2. images the user explicitly supplied;
3. figures from open-access papers on the same topic;
4. a labeled semantic diagram generated from the deck's own content;
5. a clear "figure missing" state.

Never fill a page with a random web image. A figure that does not support the slide's claim is worse than no figure.

## User-Supplied Paper

```bash
python scripts/extract_source.py --input paper.pdf --out work/source_pack
```

`source_pack/figures.json` records each extracted figure, page, size, source path, and caption when a caption can be matched. Captions are matched from blocks starting with `Figure`, `Fig.`, or `图`.

Then rank figures against the deck:

```bash
python scripts/suggest_figures.py \
  --deck deck.json \
  --figures work/source_pack/figures.json \
  --base-dir work/source_pack \
  --out work/figure_suggestions \
  --top 3 \
  --min-score 0.15
```

Review `figure_suggestions.md` before applying anything. `--apply` copies the selected image into the deck's assets and fills `image`, `image_label`, `image_caption`, `image_role`, `image_boundary`, `image_origin`, and `source`.

## No Paper Or No Figures

Search open arXiv papers and extract complete figure regions:

```bash
# Topic search
python scripts/find_figures.py \
  --query "vision language action robot" \
  --out work/figure_pack \
  --max-papers 4 \
  --max-per-paper 2

# Known papers
python scripts/find_figures.py \
  --ids "2307.15818,2406.09246" \
  --query "vision language action model" \
  --out work/figure_pack
```

The finder downloads PDFs, locates `Figure`/`Fig.` captions, renders the full figure region (including multi-panel and vector diagrams) at 200 DPI, and writes `figures.json` plus `figures.md`.

It does not choose a random embedded raster. If a caption cannot be located, the figure is skipped rather than guessed.

## Apply Suggestions

```bash
python scripts/suggest_figures.py \
  --deck deck.json \
  --figures work/figure_pack/figures.json \
  --out work/figure_suggestions \
  --query "vision language action robot" \
  --apply \
  --out-deck deck.with_figures.json
```

`--apply` is a convenience, not a truth claim. It sets `image_role` and `image_boundary` to review-required text. The author must edit them to say what the figure actually supports and what it cannot prove.

## Required Slide Fields

```json
{
  "image": "assets/figures/openvla_architecture.png",
  "image_label": "OpenVLA 架构",
  "image_caption": "视觉编码器、投影层与语言主干输出 7 维动作。",
  "image_role": "说明端到端 VLA 的组成，支持“通用策略底座”的判断。",
  "image_boundary": "只说明架构，不支持本课题的性能或执行边界。",
  "image_origin": "public",
  "source": "Kim et al., OpenVLA, arXiv:2406.09246, Figure 2"
}
```

- `image_origin`: `own`, `generated`, `template`, `borrowed`, or `public`.
- `source`: paper, figure number, and URL/arXiv id.
- `image_caption`: the figure's descriptive caption, not a fabricated result.
- `image_role` and `image_boundary`: written by the author, not the retrieval script.

## Licensing

arXiv and publisher licenses differ. Record the source and check the paper's license page before public or commercial reuse. The finder writes a license reminder but does not grant rights. For a private internal talk, citing the paper may be enough; for public distribution, verify the license.

## Quality Checks

- The figure is readable at the final slide size; do not use a 200 px thumbnail as a main figure.
- The caption and the slide claim agree.
- Multi-panel figures retain panel labels (a/b/c) and are not cropped mid-panel.
- The figure has `image_role` and `image_boundary`.
- The source line is on the slide, not only in the notes.
- A generated schematic is labeled as a schematic, never as a measured figure.
