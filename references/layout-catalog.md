# Layout Catalog

Choose a layout by the page's evidence role, not by visual habit. A deck should not repeat the same card grid on every page.

| Layout | Use when | Required fields | Avoid when |
|---|---|---|---|
| `cover` | Title and attribution | `title`, optional `subtitle` | A content claim is the main purpose |
| `outline` | Five or six section anchors | `sections` in the deck | You need sub-arguments or page ranges |
| `section` | A deliberate rhythm break between modules | `section`, `title`, optional `summary`, `keywords` | The deck is short and needs density |
| `bullets` | Two or three grouped points, usually text-led | `columns` or `bullets`, `lead`, `takeaway` | A real figure or process would explain it faster |
| `grid` | Three to six independent mechanisms or evidence blocks | `columns`, `blocks`, `lead` | The blocks are actually one sequence |
| `balance` | Two opposing rails around one dominant system figure | `left`, `right`, `image` or `chart`, `image_role`, `image_boundary`, `lead` | There is no real tradeoff or shared system object |
| `figure` | One real paper figure is the argument | `image` or `chart`, `panels`, `lead`, `takeaway` | The figure is only decorative |
| `figure_flow` | A 3–5 step principle above one main figure | `flow`, `image` or `chart`, `panels` | The flow is not a real sequence |
| `concept` | A labeled causal, structural, threshold, or feedback diagram | `diagram`, `panels`, `lead`, `takeaway` | Real measured evidence exists and should be a `figure` |
| `formula` | One equation is the page's main object | `formula_latex` and `formula`, `variables`, `panels`, `formula_condition` | The formula is decorative or has no traceable variable definitions |
| `pipeline` | Four or five ordered method steps | `steps`, optional `footer_row` | The steps are not sequential |
| `compare` | Real comparison of methods, conditions, or cases | `columns`, `rows`, optional `notes` | The rows are not comparable |
| `metrics` | Two to four measured outcomes plus conditions | `metrics`, `image` or `chart`, optional `conditions` | The numbers have no stated condition |
| `hub` | One shared problem with 4–6 parallel branches | `center`, `items`, `lead` | The branches are sequential or unrelated |
| `timeline` | Three to five research phases with dates, work packages, and milestones | `phases`, optional `period_note`, `timeline_note` | The content is not a sequence or the dates are invented |
| `references` | Four to ten traceable papers, datasets, or official documents | `references` | The entries are decorative, incomplete, or not from the user's materials |
| `closing` | Conclusions, boundary, and next step | `points`, optional `boundary_points`, `next_points` | It is only a thank-you page |

## Content-Type Mapping

Use `content_type` as an authoring hint. It does not change rendering, but it makes layout choice auditable.

| `content_type` | Preferred layout |
|---|---|
| `background_problem` | `bullets` or `grid` |
| `mechanism` | `figure_flow` or `figure` |
| `mechanism_without_figure` | `concept` |
| `central_equation` | `formula` |
| `system_architecture` | `pipeline` or `grid` with a spanning image |
| `route_taxonomy` | `hub` |
| `route_comparison` | `compare` |
| `measured_result` | `figure` or `metrics` |
| `limitation` | `bullets` or `closing` |
| `summary` | `closing` |

## `grid`

Each block supports:

- `heading`: short label.
- `body`: one short paragraph.
- `bullets`: at most four short items.
- `formula`: one-line formula.
- `formula_latex`: optional LaTeX source for the same formula; keep `formula` as the portable plain fallback.
- `image`: one figure path.
- `diagram`: an inline semantic diagram object. Use one kind from `sequence`, `branch`, `stack`, `ruler`, `matrix`, or `cycle`; see the `concept` section below.
- `footer` / `output`: one short content-derived judgement, deliverable, or metric shown as a bottom strip. Prefer this when no real image exists.
- `footer_label`: short label such as `判断`, `交付`, `输出`, `指标`, or `消融`.
- `visual`: optional concept schematic. Use only when the relation is genuinely spatial or quantitative and label it as a concept/mechanism schematic. Do not use as a generic placeholder.
- `stat` and `stat_label`: one number and its label.
- `span`: number of columns to occupy.
- `tone`: `pale` for the block that carries the main claim.

Do not turn all blocks into equal-height empty cards. Give the system image, formula, or main mechanism a `span` and more space.

If the user supplies no figure and the page is mainly prose, omit `visual`. Fill the card with body, bullets, and a `footer` strip whose text is derived from the page argument. Empty space is preferable to a histogram, target, node chain, or timeline unrelated to the content.

## `concept`

Use `concept` for a real relation that benefits from spatial encoding: an ordered evidence chain, a shared bottleneck with several routes, a hierarchy, a scale, a two-axis tradeoff, or a feedback loop. It is not a generic empty-card replacement.

```json
{
  "layout": "concept",
  "section": "method",
  "eyebrow": "3.2",
  "title": "证据接口",
  "lead": "从观测到路径，真正需要打通的是三层**证据接口**。",
  "diagram_label": "研究接口",
  "diagram_width_ratio": 0.6,
  "diagram": {
    "kind": "sequence",
    "highlight": 1,
    "items": [
      {"title": "观测接入", "body": "统一来源、时间、坐标和可见性。"},
      {"title": "边级证据", "body": "把局部观测映射到同一条路网边。"},
      {"title": "风险边权", "body": "把冲突、缺失和误差转成风险。"},
      {"title": "路径判断", "body": "在时间、风险和不确定性之间取舍。"}
    ],
    "note": "只表达研究接口，不代表已经完成的实验结果。"
  },
  "panels": [
    {"heading": "观察", "body": "先看什么。"},
    {"heading": "机制", "body": "这说明什么。"},
    {"heading": "边界", "body": "这不能证明什么。"}
  ]
}
```

Diagram kinds:

- `sequence`: 2–5 numbered nodes for a real order.
- `branch`: a shared root and 2–4 routes.
- `stack`: 2–5 nested or layered rows.
- `ruler`: 2–8 numeric ticks; every item needs a numeric `value`.
- `matrix`: exactly four quadrants plus `x_axis` and `y_axis`.
- `cycle`: 3–5 nodes in a closed loop.

Keep `highlight` at one item or omit it. Every node needs a `title`; `body` should explain the relation rather than repeat the title. A diagram is always labeled `机制示意`, `研究接口`, or `概念示意`; never describe it as measured evidence.

## `formula`

Use `formula` when the equation is the page's central object, not when a small formula is merely one block in a grid.

```json
{
  "layout": "formula",
  "section": "method",
  "eyebrow": "4.3",
  "title": "风险代价",
  "lead": "路径是否可走，取决于时间、风险与不确定性能否被**统一校准**。",
  "formula_label": "路径目标函数",
  "formula": "J(π)=Σ_e(αt_e+βr_e+γu_e)",
  "formula_latex": "J(\\pi)=\\sum_{e\\in\\pi}\\left(\\alpha t_e+\\beta r_e+\\gamma u_e\\right)",
  "formula_condition": "成立条件：边级证据可对齐，且风险项可校准",
  "variables": [
    {"symbol": "tₑ", "meaning": "边 e 的通行时间"},
    {"symbol": "rₑ", "meaning": "边 e 的风险代价"},
    {"symbol": "uₑ", "meaning": "边 e 的不确定性"},
    {"symbol": "α, β, γ", "meaning": "经验证集校准的权重"}
  ],
  "panels": [
    {"heading": "读法", "body": "公式在比较什么。"},
    {"heading": "边界", "body": "权重未校准时不能声称什么。"},
    {"heading": "验证", "body": "用什么消融或敏感性分析来检验。"}
  ]
}
```

The page must state when the equation is interpretable. Do not invent parameter values, fitted weights, or performance numbers. Use real subscript characters or LaTeX math in the visible symbol rows; `t_e` is a code identifier, not slide typography.

## `timeline`

Use for a research schedule, not for a conceptual process. Each phase needs `period`, `title`, a short `body`, and a `milestone`. Keep all dates traceable to the user's plan. Include `timeline_note` for the condition under which the schedule can change.

## `references`

Use real, traceable sources only. Each entry needs `authors`, `title`, and `venue`; include `doi` or a stable URL when available. Do not use fake page numbers, invented journals, or decorative references. Split long titles with Latin word wrapping rather than shrinking below 14 px.

## `figure`

```json
{
  "layout": "figure",
  "image": "assets/figure.png",
  "image_side": "left",
  "image_label": "论文原图",
  "image_caption": "图中对象、坐标与关键区域。",
  "image_role": "该图在本页承担的证据角色",
  "image_boundary": "该图不能证明什么",
  "panels": [
    {"heading": "观察", "body": "先看什么。"},
    {"heading": "解释", "body": "这说明什么。"},
    {"heading": "边界", "body": "这不能证明什么。"}
  ]
}
```

The renderer keeps the source aspect ratio and uses a white mount. Never redraw a measured curve, axes, legend, or hardware photo.

For `figure`, `figure_flow`, `balance`, and `metrics`, supply either `image` or `chart`. Use `chart` only for a small dataset that you can state directly; use `image` for a user-supplied paper figure, measured plot, or real system image. The native PPTX path exports `histogram`, `bar`, `column`, `line`, and `curve` charts as editable PowerPoint chart objects. If neither field is supplied, the missing-asset notice remains visible; do not fill the space with an unrelated schematic.

## `balance`

Use for a real tradeoff around one shared system object: benefits on the left, costs on the right, and the system figure in the center. This is the preferred replacement for a five-card grid when the page has one dominant architecture image.

```json
{
  "layout": "balance",
  "lead": "收益与代价来自同一套探测机制。",
  "image": "assets/system.png",
  "image_label": "系统结构",
  "image_badge": "",
  "image_role": "系统链路图（论文原图）",
  "image_boundary": "只解释信号路径与共同代价来源，不给出性能排名。",
  "left": {
    "heading": "优势：弱回波仍可形成统计量",
    "icon": "check",
    "items": [
      {"heading": "单光子灵敏度", "body": "可在弱回波下保持有效探测，适用于低反射率目标。", "icon": "atom"},
      {"heading": "低反射率目标", "body": "提升弱目标探测概率，扩展远距离场景。", "icon": "target"},
      {"heading": "公里级探测", "body": "覆盖静态目标，但受重复频率与口径约束。", "icon": "mountain"}
    ]
  },
  "right": {
    "heading": "代价：环境与器件改变有效序列",
    "icon": "alert",
    "items": [
      {"heading": "背景计数", "body": "环境光引入额外计数，降低峰值对比度。", "icon": "signal"},
      {"heading": "暗计数", "body": "自发事件形成虚假计数，污染统计结果。", "icon": "photon"},
      {"heading": "死时间", "body": "事件后短暂失响，丢失部分回波。", "icon": "clock"}
    ]
  }
}
```

Each rail should have two to four items. Write a two- or three-line `body` for each item; avoid stacking several tiny bullets beside a dominant figure. Do not use this layout if the center object is only decorative; it must be the shared system or mechanism that explains both rails.

## `hub`

```json
{
  "layout": "hub",
  "lead": "所有路线都在为同一个缺失量补充信息。",
  "center": {
    "label": "歧义阶",
    "value": "m",
    "note": "需要恢复的整数"
  },
  "items": [
    {"heading": "时间编码", "body": "用非重复序列增加区分度。", "icon": "signal"},
    {"heading": "多重复频率", "body": "用多个余数约束候选。", "icon": "layers"},
    {"heading": "统计候选", "body": "用多假设累积选主峰。", "icon": "histogram"},
    {"heading": "空间先验", "body": "用邻域几何补足单点歧义。", "icon": "target"}
  ]
}
```

Use four to six items. The center should be a shared object, question, or bottleneck, not a decorative circle. Each branch should answer “where does the extra information come from?”

## `closing`

Keep the final page useful during Q&A:

```json
{
  "layout": "closing",
  "points": ["结论一", "结论二", "结论三"],
  "boundary": "现有证据不能覆盖的范围。",
  "boundary_points": ["边界一", "边界二"],
  "next": "下一步工作。",
  "next_points": ["下一步一", "下一步二"],
  "closing": "谢谢各位老师"
}
```

The page should be readable without the speaker. Do not replace the conclusions with a blank “Thank you” slide.

