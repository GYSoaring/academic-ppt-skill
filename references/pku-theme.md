# Peking University Red Theme

`assets/themes/pku_red/` is an institutional academic-report theme for Peking University presenters. It is a visual template: using it does not make the speaker, the paper, or the deck affiliated with Peking University, and it never implies institutional endorsement.

## Palette

| Token | Value | Use |
|---|---|---|
| 北大红 / Beida Red | `#94070A` (RGB 139, 0, 18) | primary rules, chapter labels, emphasis |
| Dark red | `#6E0507` | header titles and dense text on paper |
| Mid red | `#A91B20` | secondary marks and navigation |
| Pale red wash | `#FAF2EF` | card fills and formula strips |
| Paper | `#FBF9F7` | page ground |
| Ink | `#1D1B1A` | body text |

The official VI manual also specifies CMYK `C0 M100 Y100 K45` and Pantone `PMS DS 821C / DS 821U` for 北大红. The theme uses the sRGB web value `#94070A` so the PNG renderer, native PPTX, and ordinary projectors agree.

## Brand Assets

The official VI download page is <https://vim.pku.edu.cn/xzzq/>.

- Emblem: official `标志_红色.png` from <https://vim.pku.edu.cn/docs/20240527144201516364.zip>.
- Horizontal lockup: official `标志与中英文校名组合规范_左右.png` from <https://vim.pku.edu.cn/docs/20240527144451726806.zip>.
- The shipped `emblem_mark.png` and `brand_lockup.png` are copied from those archives. Do not redraw, recolor, crop into the seal, or re-typeset the calligraphic name.
- `BRAND_SOURCE.json` records the archive URLs and file names; keep it beside the theme when redistributing.

The emblem is drawn as a quiet watermark on the cover's right column. On content pages the horizontal lockup sits in the fixed upper-right identity box. Both are decorative brand chrome: they must stay out of the title and body reading order.

## Typography

The VI manual recommends 方正大标宋 / 方正粗宋 for display and 方正中等线 for body text. Those fonts are not installed in the build environment, so the theme falls back to:

- display / serif: Noto Serif SC
- body / title / PPTX: Noto Sans SC
- formula: Cambria Math

The fallback is recorded in `fonts` and `pptx_fonts`. If the official fonts are installed later, change the ordered candidates in `theme.json`; do not bake a font file into the theme.

## Layout Contract

- Canvas: 2048 × 1152 (16:9), `defense` type profile.
- Cover: editorial composition, official horizontal lockup at upper left, faint official emblem on the right, no header band and no sidebar.
- Content pages: fixed red header rule, chapter label and title sharing one baseline, five-item left sidebar, footer rule at `layout.footer_y = 1124`.
- Cards: horizontal red-to-rose gradient header bars, thin warm-gray borders, and pale red card faces that fade toward paper white; no corner marks, icon badges, or raised glass effects.
- Figures: paper figures are mounted whole and unstretched. Keep `image_label`; use `image_badge` only when a second, different tag is needed, otherwise it duplicates the label.
- Formula pages: one pale formula strip, a condition line, then a measured symbol list. The condition strip and footnote reserve their real glyph height.
- Closing pages: the three numbered conclusions stay above two bottom cards; the closing line sits below the cards and above the footer. Do not shrink the cards to force extra bullets; cut or merge the copy instead.

The soft gradient is part of the theme contract, not a one-off slide style. `panel.header_tone: "mist"`,
`header_gradient_from: "blue_pale"`, `header_gradient_mid: "blue_wash"`, and
`header_gradient_to: "white"` keep the dark module labels legible while making the right edge visibly
lighter; the separate `header_accent` rule stays in 北大红. `panel.fill_style: "soft_gradient"` fades
`blue_wash` and `blue_pale` card faces toward white. The PNG renderer and native PPTX exporter must use
the same horizontal direction.

## Usage

```bash
python scripts/render_deck.py \
  --deck path/to/deck.json \
  --theme assets/themes/pku_red/theme.json \
  --out path/to/output \
  --profile defense

python scripts/export_editable_pptx.py \
  --deck path/to/deck.json \
  --theme assets/themes/pku_red/theme.json \
  --out path/to/output \
  --pptx-name defense_deck.pptx \
  --profile defense
```

Set `meta.navigation_mode` as usual. The bundled GM-APD example uses `sidebar`.

## Provenance Wording

The example deck in this repository analyses a paper by Chang et al., *Infrared Physics & Technology* 150 (2025) 105930. The paper authors are at Tianjin University. The Peking University theme is used only as a presentation template. Keep that distinction in the cover metadata, the source note, and the spoken script:

> 北京大学视觉模板 · 论文作者单位：天津大学

Do not write "北京大学研究团队", "北京大学成果", or any equivalent unless the deck really reports Peking University work.
