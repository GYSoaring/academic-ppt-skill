# Deck JSON

The model writes content; the renderer owns layout. Use one slide object per page and keep every visual claim traceable to the supplied materials.

## Top Level

```json
{
  "schema_version": 1,
  "meta": {},
  "sections": [],
  "slides": [],
  "qa": []
}
```

## Meta

```json
{
  "title": "报告题目",
  "subtitle": "毕业答辩 / 中期检查 / 课程汇报",
  "presenter": "汇报人",
  "advisor": "指导教师",
  "institution": "学校或学院",
  "department": "学院或项目名称，可选",
  "report_type": "硕士学位论文开题报告",
  "date": "2026.11.13",
  "duration_min": 20,
  "deck_mode": "full",
  "language": "zh-CN",
  "cover_line": "封面底部的一句边界或判断，可选",
  "source_note": "材料范围与任何全局限制"
}
```

`deck_mode` controls whether the deck includes the standard opening pages:

- `full`: default. Slide 1 is `cover`, slide 2 is `outline`.
- `content_only`: omit both the cover and the outline. Every slide is a content or closing page, and the validator does not require the first two layouts.

Use `content_only` only when the user explicitly asks for a content-only deck or says to remove the cover and table of contents.

## Sections

`sections` controls the outline page and the mini-directory. Use exactly five sections for the default theme.

```json
[
  {"id": "background", "title": "研究背景"},
  {"id": "progress", "title": "技术进展"},
  {"id": "method", "title": "原理创新"},
  {"id": "evidence", "title": "实验验证"},
  {"id": "conclusion", "title": "总结展望"}
]
```

## Common Slide Fields

```json
{
  "id": "03",
  "layout": "bullets",
  "content_type": "background_problem",
  "section": "background",
  "eyebrow": "1.1",
  "title": "研究问题",
  "lead": "本页要回答的问题或一句核心判断。",
  "takeaway": "本页的结论、边界或向下一页的过渡。",
  "claim": "一句可检验的主张。",
  "evidence": "支持该主张的数据、图、公式或文献。",
  "boundary": "该证据不能证明什么。",
  "footnote": "可选的来源、口径或缩写说明。",
  "speech": "本页的口头讲解，45 至 90 秒。"
}
```

`section` must match a section `id`. `eyebrow` is the short chapter label in the header, such as `1.1`. It is rendered at the same measured size as the adjacent title; keep it short enough to fit the chapter label column.

`content_type` is an optional authoring label such as `mechanism`, `route_taxonomy`, `measured_result`, or `limitation`. Use it to make layout choice auditable; see [layout-catalog.md](layout-catalog.md).

## Layout: cover

```json
{
  "id": "01",
  "layout": "cover",
  "title": "报告题目",
  "title_lines": ["第一行手动断行", "第二行手动断行"],
  "subtitle": "副标题可选",
  "speech": "开场。"
}
```

Presenter, advisor, and date come from `meta`. `title_lines` is optional but recommended for long Chinese titles: each line is kept as written and the renderer reduces the font size until every line fits. The cover never shows a mini-directory or title bar.

## Layout: outline

Only the section titles are rendered.

```json
{
  "id": "02",
  "layout": "outline",
  "title": "汇报提纲",
  "speech": "用一句话说明五段之间的逻辑关系。"
}
```

## Layout: section

Use a section divider only when the deck needs a rhythm break. It still uses the fixed header.

```json
{
  "id": "03",
  "layout": "section",
  "section": "method",
  "eyebrow": "3.0",
  "title": "原理创新",
  "section_title": "从时间编码到歧义阶判决",
  "summary": "这一节只回答一个问题：怎样在保留单光子预算的同时恢复 m。",
  "keywords": ["候选结构", "统计判决", "资源边界"],
  "speech": "进入方法部分。"
}
```

## Layout: bullets

```json
{
  "id": "03",
  "layout": "bullets",
  "section": "background",
  "eyebrow": "1.1",
  "title": "核心优势与技术代价",
  "lead": "单光子灵敏度与事件损失来自同一套探测机制。",
  "columns": [
    {
      "heading": "收益",
      "bullets": ["单光子级灵敏度", "弱回波可形成统计量", "适合公里级探测"]
    },
    {
      "heading": "代价",
      "bullets": ["背景计数抬高底噪", "暗计数产生伪事件", "死时间改变有效序列"]
    }
  ],
  "takeaway": "收益与约束同源，比较方法时必须同时报告探测条件。",
  "speech": "本页讲解。"
}
```

Limits: two or three columns, at most four bullets per column, each bullet at most two rendered lines.

## Layout: grid

Use `grid` when the original page needs several small evidence blocks rather than two large cards. It supports two or three columns and three to six blocks.

```json
{
  "id": "04",
  "layout": "grid",
  "section": "background",
  "eyebrow": "1.2",
  "title": "核心优势与技术代价",
  "lead": "收益与约束来自同一套单光子探测机制。",
  "columns": 2,
  "blocks": [
    {
      "heading": "单光子灵敏度",
      "body": "弱回波仍能形成可累积的事件统计量。",
      "bullets": ["低反射率目标", "公里级弱信号"],
      "icon": "signal",
      "tone": "pale"
    },
    {
      "heading": "背景与暗计数",
      "body": "额外事件会抬高底噪并改变峰值显著性。",
      "bullets": ["强背景", "暗计数", "后脉冲"]
    },
    {
      "heading": "死时间",
      "body": "探测器在单次事件后短暂失去响应。",
      "bullets": ["有效事件率下降", "高背景时更明显"]
    },
    {
      "heading": "系统图",
      "body": "保留原图，不重绘硬件结构。",
      "image": "assets/system.png"
    },
    {
      "heading": "折叠关系",
      "formula": "R_fold = R_true mod R_0",
      "body": "公式块使用衬线字体和浅蓝底板。",
      "icon": "layers"
    }
  ],
  "takeaway": "评价方法时必须同时报告收益、探测条件和事件级损失。",
  "speech": "本页讲解。"
}
```

Each block may contain `heading`, `body`, `bullets`, `formula`, `formula_latex`, `footer`, `footer_label`, `image`, `visual`, `visual_semantic`, `stat`, `stat_label`, `span`, `icon`, `body_size`, and `tone`. Supported icon names include `signal`, `wave`, `target`, `clock`, `layers`, `chart`, `check`, `alert`, `matrix`, `atom`, `photon`, `laser`, `histogram`, `mountain`, `chip`, and `database`. `formula` is the portable plain fallback; `formula_latex` is the optional LaTeX source. `footer` renders a bottom information strip with `footer_label` such as `判断`, `交付`, `输出`, `指标`, or `消融`. Use it instead of a generic `visual` when no real figure exists. `visual` is drawn only when the deck also sets `visual_semantic: true`; it is reserved for a genuine concept schematic and must be semantically related to the block. Use `span` to give a system image or formula across columns. Use `pale` sparingly for the block that carries the main claim. Do not use a grid for more than six blocks; split the page instead.

## Layout: balance

Use for a real tradeoff with one dominant central system figure. The left rail carries benefits or enabling conditions; the right rail carries costs, failure modes, or constraints.

```json
{
  "id": "04",
  "layout": "balance",
  "section": "background",
  "eyebrow": "1.2",
  "title": "核心优势与技术代价",
  "lead": "收益与约束来自同一套探测机制。",
  "image": "assets/system.png",
  "image_label": "系统结构",
  "image_caption": "系统链路示意；不用于性能比较。",
  "image_badge": "",
  "image_role": "系统链路图（论文原图）",
  "image_boundary": "只解释信号路径与共同代价来源，不给出性能排名。",
  "left": {
    "heading": "收益：弱回波仍可形成统计量",
    "icon": "check",
    "items": [
      {"heading": "单光子级灵敏度", "body": "可在弱回波下保持有效探测，适用于低反射率目标。", "icon": "atom"},
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
  },
  "takeaway": "评价方法时必须同时报告探测条件与事件级损失。",
  "speech": "本页讲解。"
}
```

Each rail should contain two to four items. Prefer a two- or three-line `body` over several tiny bullets. The center object must be the shared system or mechanism, not a decorative image.

## Layout: figure

Use an untouched image; do not redraw measured content.

```json
{
  "id": "05",
  "layout": "figure",
  "section": "background",
  "eyebrow": "1.3",
  "title": "距离模糊的产生机制",
  "lead": "周期接收机只保留飞行时间余数。",
  "image": "assets/figure.png",
  "image_side": "left",
  "image_role": "论文原图 / 系统图 / 实验场景",
  "image_boundary": "该图证明什么、不能证明什么",
  "panels": [
    {"heading": "观测", "body": "接收端看到的是折叠后的到达时间。"},
    {"heading": "缺失", "body": "整数周期编号 m 没有被记录。"},
    {"heading": "结论", "body": "提高 TDC 精度不能补回离散的 m。"}
  ],
  "takeaway": "距离模糊是周期编号缺失，不是计时精度不足。",
  "speech": "解释图中的坐标、对象、趋势与边界。"
}
```

`image` may be absolute or relative to the deck file. Prefer two or three panels.

## Layout: figure_flow

Use when the page needs both a mechanism sequence and a main figure. The flow sits above the evidence object; the right rail explains what to see.

```json
{
  "id": "05",
  "layout": "figure_flow",
  "section": "background",
  "eyebrow": "1.1",
  "title": "单光子测距链",
  "lead": "测距依赖光子到达时间的统计峰值。",
  "flow": [
    {"title": "发射", "body": "高重频脉冲", "icon": "laser"},
    {"title": "探测", "body": "SPAD 记录事件", "icon": "photon"},
    {"title": "计时", "body": "TDC 时间戳", "icon": "clock"},
    {"title": "累积", "body": "多帧直方图", "icon": "histogram"},
    {"title": "判决", "body": "峰值对应 τ", "icon": "check"}
  ],
  "chart": {
    "type": "histogram",
    "values": [8, 17, 31, 56, 96, 68, 43, 27, 15, 8],
    "highlight": 4,
    "x_label": "到达时间 t",
    "y_label": "光子计数（概念示意）",
    "annotation": "τ"
  },
  "image_label": "时间直方图（概念示意）",
  "image_caption": "本图为机制示意，不是实测分布。",
  "panels": [
    {"heading": "事件记录", "body": "单次事件方差大。"},
    {"heading": "统计累积", "body": "多帧累积提高峰值信噪比。"},
    {"heading": "边界", "body": "这里只讨论常规单周期测距。"}
  ],
  "takeaway": "先建立统计量，再做峰值判决。",
  "speech": "本页讲解。"
}
```

`chart.type` may be `histogram`, `bar`, `column`, `line`, or `curve`. The native PPTX exporter keeps these as editable PowerPoint chart objects; `values`, `highlight`, `x_label`, `y_label`, and `annotation` must remain in `deck.json`, not baked into a screenshot. Use `image` instead when the evidence is a supplied paper figure or a measured plot that cannot be reduced to a small data table.

## Layout: pipeline

```json
{
  "id": "08",
  "layout": "pipeline",
  "section": "method",
  "eyebrow": "3.1",
  "title": "方法流程",
  "lead": "从发射编码到歧义阶判决的完整链路。",
  "steps": [
    {"title": "编码", "body": "构造非重复时间指纹", "bullets": ["单帧不重复", "帧间继续重排"]},
    {"title": "探测", "body": "记录离散光子事件", "bullets": ["事件级建模", "保留死时间"]},
    {"title": "折叠", "body": "映射全部候选", "bullets": ["覆盖全量程", "维护统一直方图"]},
    {"title": "判决", "body": "累积后选择主峰", "bullets": ["真峰累积", "假峰散开"]}
  ],
  "takeaway": "每一步都对应可验证的输入、输出与代价。",
  "speech": "本页讲解。"
}
```

Use four or five steps.

## Layout: compare

```json
{
  "id": "09",
  "layout": "compare",
  "section": "progress",
  "eyebrow": "2.3",
  "title": "技术路线比较",
  "lead": "比较维度是新增信息与系统代价，不是单一最远距离。",
  "columns": ["路线", "新增信息", "优势", "主要代价"],
  "rows": [
    ["多重复频率", "多个余数", "原理直接", "误差容限敏感"],
    ["时序编码", "长序列指纹", "可辨识周期", "存储与相关成本"],
    ["多假设折叠", "候选统计结构", "保留单光子预算", "直方图随量程增长"]
  ],
  "notes": ["比较重点：新增信息、误差容限、存储量与场景假设。"],
  "takeaway": "不同实验条件，不作严格排名。",
  "speech": "本页讲解。"
}
```

Limits: two to five columns, at most five rows. Use `highlight_rows` for the row that carries the comparison; do not highlight every row.

## Layout: hub

Use when one shared object or question has four to six parallel branches. The center must be a real shared thing, not decoration.

```json
{
  "id": "08",
  "layout": "hub",
  "section": "progress",
  "eyebrow": "2.1",
  "title": "消歧路线全景",
  "lead": "不同路线都在为歧义阶 m 引入额外信息。",
  "center": {
    "label": "歧义阶",
    "value": "m",
    "note": "需要恢复的整数"
  },
  "items": [
    {"heading": "时间编码", "body": "用非重复序列增加区分度。", "icon": "signal", "visual": "pulse"},
    {"heading": "多重复频率", "body": "用多个余数约束候选。", "icon": "layers", "visual": "nodes"},
    {"heading": "统计候选", "body": "用多假设累积选主峰。", "icon": "histogram", "visual": "peak", "tone": "pale"},
    {"heading": "空间先验", "body": "用邻域几何补足单点歧义。", "icon": "target", "visual": "target"}
  ],
  "takeaway": "共同点是把“猜 m”变成有约束的推断。",
  "speech": "本页讲解。"
}
```

## Layout: metrics

```json
{
  "id": "10",
  "layout": "metrics",
  "section": "evidence",
  "eyebrow": "4.2",
  "title": "外场结果",
  "lead": "所有指标均对应等单脉冲能量条件。",
  "metrics": [
    {"value": "100%", "label": "强光检测概率", "note": "MHF"},
    {"value": "−3.04 dB", "label": "PSLR", "note": "MHF"},
    {"value": "3.39×", "label": "SNR 比", "note": "同一能量口径"}
  ],
  "image": "assets/result.png",
  "image_label": "测量结果",
  "image_badge": "实测",
  "conditions_title": "实验条件",
  "conditions": [
    {"label": "目标距离", "value": "约 2.1 km"},
    {"label": "重复频率", "value": "500 kHz"},
    {"label": "TDC 分辨率", "value": "1 ns"}
  ],
  "conditions_note": "绝对标定、运动目标与片上资源尚未覆盖。",
  "footnote": "条件：1064 nm、500 kHz、TDC 1 ns。",
  "takeaway": "结果支持恢复正确歧义阶，不支持无条件全面胜出。",
  "speech": "本页讲解。"
}
```

Use two to four metrics. `conditions` can be a list of `{label, value}` objects or short strings. `conditions_note` is the mandatory boundary when the numbers do not generalize.

## Layout: closing

```json
{
  "id": "12",
  "layout": "closing",
  "section": "conclusion",
  "eyebrow": "5.2",
  "title": "总结与讨论",
  "lead": "结论必须带着条件和边界离开。",
  "points": [
    "距离模糊的本质是周期编号缺失",
    "多假设折叠用统计结构恢复候选",
    "外场证据覆盖静态单目标与给定能量口径"
  ],
  "boundary": "现有证据只覆盖静态单目标与给定能量口径。",
  "boundary_points": ["运动目标尚未覆盖", "绝对标定未独立完成"],
  "next": "下一步：统一口径比较多目标与资源开销。",
  "next_points": ["补运动场景", "评估 FPGA 资源"],
  "closing": "谢谢各位老师",
  "speech": "收束并邀请提问。"
}
```

## Layout: timeline

Use three to five phases. Dates must come from the user's plan; do not infer a schedule.

```json
{
  "id": "11",
  "layout": "timeline",
  "section": "plan",
  "eyebrow": "5.1",
  "title": "研究计划",
  "lead": "计划按阶段滚动评估。",
  "period_note": "2026.12 — 2028.05",
  "card_height": 480,
  "timeline_note": "若数据或算力条件变化，先回退评价口径，再调整里程碑。",
  "phases": [
    {
      "period": "2026.12–2027.03",
      "title": "数据与基线",
      "body": "完成数据规范、公开数据预处理和基线复现。",
      "bullets": ["统一坐标与时间规范", "复现基线"],
      "milestone": "数据规范 v1 + 基线复现"
    }
  ]
}
```

## Layout: references

Use real sources only. Four to ten entries are supported.

```json
{
  "id": "13",
  "layout": "references",
  "section": "plan",
  "eyebrow": "5.3",
  "title": "参考文献",
  "lead": "参考文献以公开论文、数据集与官方文档为主。",
  "reference_label": "参考文献",
  "references": [
    {
      "authors": "Author A, et al.",
      "title": "Traceable paper or dataset title",
      "venue": "Venue, Year",
      "doi": "DOI: 10.xxxx/xxxxx"
    }
  ]
}
```

The renderer keeps Latin words intact when wrapping and uses the theme's Latin font for diacritics.

## QA

`qa` is optional but recommended.

```json
[
  {
    "question": "为什么提高 TDC 精度不能解决距离模糊？",
    "answer": "TDC 改善局部时间误差，不能提供离散的周期编号。"
  }
]
```

## Writing Rules

- Keep one central claim per slide.
- Use short headings, not slogans.
- Do not invent references, measured values, baselines, or hardware resources.
- Keep figure paths valid; validation must fail on a missing image.
- `speech` is required for every slide.
