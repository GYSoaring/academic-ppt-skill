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

## Minimum Theme Contract

- `canvas.width` and `canvas.height`
- `colors` with `paper`, `white`, `ink`, `muted`, `blue`, `blue_dark`, `blue_mid`, `blue_pale`, `blue_wash`, `line`, and `red`
- `fonts.title`, `fonts.body`, `fonts.serif`, `fonts.display`, `fonts.formula`, and `fonts.latin` as ordered fallback lists
- `layout.content`, `layout.sidebar`, `layout.header_height`, `layout.footer_y`, and `layout.takeaway_top`
- `type_scale.header_title` and `type_scale.eyebrow` must describe the same header type size; `layout.title_size` and `layout.chapter_size` should also match. The renderer fits both to the same measured size so a long title shrinks together with its chapter number.
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

## Public-Use Checklist

- No school name, seal, topic, or discipline is hard-coded in the renderer.
- Missing lockup falls back to text or monogram.
- Missing font falls back through the ordered list.
- Missing image produces a content-derived `footer`/`output` or a clear validation error, never a generic placeholder visual.
- Required metadata is visible in validation; optional metadata can remain blank.
- Formulas have both `formula_latex` and a plain `formula` fallback.
