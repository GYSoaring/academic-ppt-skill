# Theme Authoring

The renderer must stay institution-agnostic. A school, lab, or conference identity belongs in `assets/themes/<theme-name>/`, not in Python code.

## Required Files

```text
theme-name/
|-- theme.json
|-- brand_lockup.png     # optional: supplied or approved wordmark/seal lockup
|-- emblem_mark.png      # optional: transparent seal for an editorial cover
`-- SOURCE.md            # where the identity assets came from and their usage limit
```

If the user does not provide an approved identity asset, omit the images and use `brand.fallback_text` or a neutral monogram.

You may look for a public emblem when the user explicitly names the institution and expects a branded deck. Do not draw a fake seal, and do not claim official endorsement.

## Import An Existing PPTX Theme

If the user supplies a school or lab `.pptx`, start with:

```bash
python scripts/import_pptx_theme.py \
  --pptx school_template.pptx \
  --out assets/themes/imported/theme.json \
  --name "学校名称"
```

The importer reads `ppt/theme/theme*.xml`, maps theme colors to the semantic palette, reads major/minor Latin and East Asian fonts, and reuses the layout geometry of `academic_blue`. Treat the result as a draft: verify the palette, font availability, cover identity, and protected zones before rendering the full deck. The importer does not infer an official seal or logo from arbitrary slide images.

## Acquire And Compose A Brand Asset

Use `scripts/build_brand_lockup.py` after you have a local emblem, a direct image URL, or a Wikimedia Commons search result:

```bash
python scripts/build_brand_lockup.py \
  --theme assets/themes/tsinghua_purple/theme.json \
  --name "清华大学" \
  --latin "TSINGHUA UNIVERSITY" \
  --emblem path/to/emblem.png
```

With a search query:

```bash
python scripts/build_brand_lockup.py \
  --theme assets/themes/imported/theme.json \
  --name "学校名称" \
  --latin "UNIVERSITY NAME" \
  --search-query "University name logo"
```

The script:

- loads a local or remote image;
- converts near-white pixels to alpha unless `--keep-white` is set;
- crops the transparent bounding box;
- composes the emblem, school name, rule, and optional Latin subtitle;
- writes `emblem_mark.png`, `brand_lockup.png`, and `BRAND_SOURCE.json`;
- updates `brand.lockup`, `brand.emblem`, and `brand.fallback_text` in `theme.json`.

The search uses Wikimedia Commons and prints the candidate list plus the selected URL. Review the selected file before delivery. Keep the source in `SOURCE.md` or `BRAND_SOURCE.json`, and follow the institution's visual identity rules.

### Prefer The Official Wordmark

Many universities use a calligraphic or custom-drawn wordmark, not a standard font. If an official horizontal lockup exists, use it directly:

```bash
python scripts/build_brand_lockup.py \
  --theme assets/themes/tsinghua_purple/theme.json \
  --name "清华大学" \
  --lockup path/to/official_lockup.png \
  --lockup-crop "0,0,236,88" \
  --recolor-lockup
```

- `--lockup` accepts a local path or a direct image URL.
- `--lockup-crop` removes adjacent text such as a VI-system title.
- `--recolor-lockup` uses the theme primary color while preserving the alpha mask. Use it only for monochrome official artwork; skip it for a full-color logo.
- `BRAND_SOURCE.json` records `mode: official_lockup` and `typographic_fallback: false`.

The typeset path is a fallback. It records `typographic_fallback: true`; do not present it as the official school wordmark.


## Lean Mode（可选：把版面骨架的花销压下来）

同一份内容，一个页面的对象数从 ~33 压到 ~29 甚至 ~21，靠的是下面五个开关。它们都只删装饰、不动文字，
所以不需要改 deck.json：

```json
{
  "panel": {
    "corner_marks": false,      // 面板四角的小折角（每块 4 个对象）
    "card_index": false,        // 卡片左上角的 01 / 02
    "card_highlight": false,    // 标题条顶部那条高光
    "header_style": "gradient"  // gradient = 深蓝标题条；flat = 纯色，会与 PPTX 导出不一致
  },
  "layout": {
    "block_icons": "off",       // 圆形图标徽章
    "lead_panel": false,        // lead 的浅灰底板；关掉后文字直接压在纸面上
    "sidebar_mode": "off",      // 常驻小目录
    "panel_radius": 2,
    "panel_shadow_alpha": 0
  }
}
```

实测（13 页 deck，同内容）：全开 33.4 对象/页 → 关角标+序号+高光 32.4 → 再去 lead 底板 28.8，
单页最低 21。作为参照，一份真人做的同题材答辩 deck 是 15.8 对象/页。

**注意**：`lead_panel: false` 会让版面明显变松。这是审美选择，不是优化，用之前请先看渲染结果。
`header_style: "flat"` 只在 PNG 预览里生效，PPTX 导出器不读它——需要扁平卡片就得改导出器。

## Minimum Theme Contract

- `canvas.width` and `canvas.height`
- `colors` with `paper`, `white`, `ink`, `muted`, `blue`, `blue_dark`, `blue_mid`, `blue_pale`, `blue_wash`, `line`, and `red`
- `fonts.title`, `fonts.body`, `fonts.serif`, `fonts.display`, `fonts.formula`, and `fonts.latin` as ordered fallback lists
- `layout.content`, `layout.sidebar`, `layout.header_height`, `layout.footer_y`, and `layout.takeaway_top`
- `type_scale.header_title` and `type_scale.eyebrow` must describe the same header type size; `layout.title_size` and `layout.chapter_size` should also match. The renderer fits both to the same measured size so a long title shrinks together with its chapter number.
- `type_profiles` is optional. If omitted, built-in scales cover `projection`, `lecture`, and `defense`. A theme may override a profile with `type_profiles.<name>.type_scale` to match its own layout density. `text.body_bold_min_px` and `text.body_bold_max_px` constrain semantic body bolding; do not use `body_bold` as a global bold-all switch.
- `layout.sidebar_mode`: `always` | `section` | `off`. `off` also uses `layout.content_no_sidebar` as the text box, which is how a theme buys back the 260 px the directory occupied.
- `layout.panel_radius` and `layout.panel_shadow_alpha`: a designed academic theme is flat. Defaults are `8` px and `20`; `2` / `0` with `panel.header_style = "flat"` removes the dashboard feel without touching any layout code.
- `brand.lockup` or `brand.fallback_text`
- `cover.style`: `centered` or `editorial`

Every font list must contain cross-platform fallbacks. Do not assume that a font installed on the author's machine exists on the user's machine.

For native PowerPoint output, add optional `pptx_fonts` with installed family names:

```json
{
  "pptx_fonts": {
    "title": "Noto Sans SC",
    "body": "Noto Sans SC",
    "display": "Source Han Serif CN SemiBold",
    "latin": "Arial",
    "formula": "Cambria Math"
  }
}
```

The font lists above use file paths for the PNG renderer; `pptx_fonts` uses PowerPoint family names. Keep both in sync when creating a new theme.

## Visual Adaptation

1. Copy the closest existing theme, never edit the shared renderer for one school.
2. Replace colors and identity first; keep the fixed header and footer geometry stable.
3. Set `background.motif` to `contour`, `arcs`, or `campus` only when it fits the institution; generic themes should usually use `arcs` or no motif.
4. Keep `panel_radius` small and `panel_shadow_alpha` restrained. The academic look comes from hierarchy, spacing, and typography, not from dashboard effects.
5. Use the same identity asset family on cover and content pages. Do not mix a dark-background lockup with a light-background page.
6. Preview the cover and one content page before rendering the full deck.
7. Run `--profile document` and the intended projected profile (`projection`, `lecture`, or `defense`) on a representative page. A projected profile is expected to reject pages that are too dense; split them instead of reducing the font below the profile minimum. For `defense`, also run `audit_pptx.py` on the exported PPTX.

## Public-Use Checklist

- No school name, seal, topic, or discipline is hard-coded in the renderer.
- Missing lockup falls back to text or monogram.
- Missing font falls back through the ordered list.
- Missing image produces a content-derived `footer`/`output` or a clear validation error, never a generic placeholder visual.
- Required metadata is visible in validation; optional metadata can remain blank.
- Formulas have both `formula_latex` and a plain `formula` fallback.
