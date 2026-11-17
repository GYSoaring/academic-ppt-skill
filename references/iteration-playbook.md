# 迭代与验收：把“做得好”变成可重复流程

这份手册解决的不是“怎么再想一个漂亮版式”，而是另一个更常见的问题：同一个 skill，不同模型执行，结果差距很大。

高质量 deck 不是一次生成出来的，而是沿着固定阶段收敛出来的。模型越弱，越不能让它自由发挥；越要把阶段、检查项、停止条件和失败后的第一动作写死。

本手册适用于 8 页以上的学术 deck，答辩、开题、中期和课程汇报默认强制使用。

## 四条硬规则

1. **先样板，后全套。** 不要先生成 15 页再整体返工。先做 4–5 页样板，把字体、页眉、图版率和最密的一页压到可接受，再铺满全篇。
2. **一轮只改一类问题。** 这一轮修溢出，就只修溢出；下一轮修对齐；再下一轮修 PNG 与 PPTX 的差异。混合修改会让“变好了还是变坏了”无法判断。
3. **版面几何一旦冻结，只改内容。** 主题、页眉高度、内容区、侧栏宽度和字号阶梯冻结后，不再为了单页好看而改。任何几何改动都会让全部旧报告失效。
4. **以门禁 PASS 结束，不以“看起来差不多”结束。** 每轮结束必须留下报告；最后一轮必须运行 `scripts/quality_gate.py`，并读 `quality_gate.md`。

## 五个阶段

| 阶段 | 产物 | 进入下一阶段的条件 | 常见错误 |
|---|---|---|---|
| 0. 内容契约 | `deck.json` 骨架：封面、目录、章节、每页 claim / evidence / boundary | 标题链能独立读懂；每页能用一个短语说出“主对象”；图表来源齐全 | 先想配色，后补论证 |
| 1. 视觉样板 | 4–5 页 PNG：封面、目录、最密内容页、真实图页、结论页 | 页眉/页脚/身份层稳定；无溢出；最密页在 100% 尺寸下可读 | 只看封面就铺全篇 |
| 2. 全套版面 | 全量 `rendered/`、`contact_sheet.jpg`、`validation_report.json`、`design_report.json` | `error_count = 0`；render warnings = 0；设计分 ≥ 80 | 用缩小字号解决溢出 |
| 3. 原生 PPTX | `defense_deck.pptx`、`pptx_validation_report.json`、`accessibility_report.json`、`pptx_readability.json` | PPTX 0 错误；无障碍 0 错误；答辩档 `defense_ready` | 只信 PNG，不检查 PPTX |
| 4. 交付 | PowerPoint 实机渲染、`quality_gate.md`、讲稿与 Q&A | 实机 12/12 页；无空白/重复页；门禁 PASS | 没打开过 PPTX 就交付 |

### 阶段 1 为什么必须是 4–5 页

封面通常是最宽松的一页，最密的正文页才是字体和几何的真正压力测试。样板必须包含：

- 封面：学校身份、标题层级、汇报人信息。
- 目录：五项以内，行号与标题同一基线。
- 最密内容页：文字最多、块最多、最容易溢出的一页。
- 真实图页：论文图、caption、来源行和图注边界。
- 结论页：结论卡、边界和下一步。

样板通过后再生成全套。若最密页放不下，改内容或版式，不要先降字号。

## 每一轮的标准动作

```text
读门禁/报告 → 用一句话命名缺陷 → 选最小改动 → 只重渲染受影响页
→ 记录修改前后数字 → 再跑门禁 → 决定继续还是回滚
```

### 1. 用一句话命名缺陷

坏例子：

- “第 6 页不太好。”
- “再美化一下。”
- “感觉有点空。”

好例子：

- “第 6 页公式符号说明从 7 行溢出到页脚，正文下限被突破。”
- “第 9 页主图只占内容区 38%，观众看不清坐标轴。”
- “第 12 页结论卡标题 52 px、正文 39 px，但正文 3 行溢出卡片。”

命名不出一句话，说明还没定位到对象；此时不要改代码。

### 2. 选最小改动

按这个顺序升级，能停在低一级就不要往上走：

1. **改内容**：删一句、合并一个要点、拆成两页。
2. **换版式变体**：`figure.image_left` / `image_right`、`quote.split`、`closing` 卡片化。
3. **改主题字段**：字号、间距、页眉高度、内容区宽度。
4. **改渲染器**：只有 PNG 与 PPTX 同时错、且属于共享几何时才改。

### 3. 只重渲染受影响页

```bash
python scripts/render_deck.py --deck deck.json --theme theme.json --out out --pages 3,7
```

迭代阶段可以只出问题页；**进入阶段 4 前必须全量重渲染一次**，否则门禁会把旧页判为过期。

### 4. 记录修改前后

每轮在 `iteration_log.md` 追加一条：

```markdown
## Pass 2 — 2026-11-17

- Stage: deck
- Hypothesis: 第 6 页溢出是因为符号说明超过 6 项
- Changed: page 6 symbol list → two columns
- Before: error 1 / warning 4 / design 78
- After: error 0 / warning 2 / design 84
- Decision: continue（下一轮修第 9 页主图宽度）
```

没有这行记录，下一轮很容易把上一轮的修复又改回去。

### 5. 回滚规则

如果一次修改引入了新的 blocking issue，先回滚这次修改，再换一种修法。不要在一个已经变坏的状态上继续叠补丁。

## 症状 → 第一动作 → 禁止动作

| 症状 | 第一动作 | 禁止 |
|---|---|---|
| 正文溢出卡片或页脚 | 删 20% 文字，或拆页；换更宽版式 | 把正文降到 18 pt 以下 |
| 三张等宽卡片、没有主对象 | 选一个主图/主公式/主数字，另外两块降为支撑 | 再加一张卡片填满空白 |
| 页眉章节号与标题不齐 | 把 `eyebrow` 与 `header_title` 设为同一字号、同一基线 | 只移动其中一个 |
| 目录 `01` 被拆成两行 | 短数字 token 禁止换行；PPTX 里 `wrap_explicit=False` | 缩小目录字号 |
| 公式符号说明挤成一列 | ≤6 个符号单列，>6 个符号双列 | 用 `N_s` 这类 ASCII 伪下标 |
| 论文图看不清 | 换 `figure.image_left/right`，让图占主要宽度 | 拉伸低分辨率图片 |
| PNG 好看、PPTX 跑版 | 对照渲染器坐标修 exporter，并补回归测试 | “PPTX 打开应该差不多” |
| PPTX 结论卡溢出 | 每卡 ≤2 条短 bullet，或改成无 bullet 的结论卡 | 把 39 px 正文继续缩小 |
| 无障碍缺 alt | 语义 alt；纯装饰标 decorative | `alt = 文件名` |
| 答辩正文低于 24 pt | 拆页；把标签/脚注 token 化 | 调低 `defense` 下限 |
| 设计分低于 80 | 先看 `card_only_ratio`、`dominant_object_ratio`、`image_page_ratio` | 加装饰来“刷分” |
| `chars_per_line < 18` | 减一列或加宽正文栏 | 缩小字号换行长 |
| 同一问题在三页反复出现 | 改主题或渲染器，并补测试 | 逐页打补丁 |

## 弱模型模式

模型能力不足时，不要让它“自由设计”。按下面的约束执行：

- **只读一个参考**：门禁的修复队列会给出问题类型；只打开对应的那一份 `references/*.md`，不要一次读十份。
- **只选目录里的版式**：从 `references/layout-catalog.md` 选，不发明新布局。
- **不改共享代码**：除非 PNG 和 PPTX 同时错，否则不动渲染器。
- **两次失败就换策略**：同一页同一问题修两次仍失败，改为拆页或删内容，不再调像素。
- **没有视觉能力时**：以 `quality_gate.md`、`validation_report.json`、`pptx_readability.json` 和 `pptx_render_report.json` 为准；不要凭想象判断“应该没问题”。
- **不追求满分**：设计分 ≥85 后停止美容，剩余时间用于证据、边界和讲稿。

## 停止条件

满足以下全部条件才可以结束：

- `quality_gate.md` 状态为 `PASS`；所有 blocking 为 0。
- advisory 已逐条读过，保留项写进交付说明。
- 全量 PNG、PPTX、无障碍、字号、实机预览报告都比输入文件新。
- 联系表在 25% 缩略图下能看出页面节奏；单页在 100% 下无溢出。
- 至少打开过一次真实 PPTX 渲染，而不是只看 PIL 预览。
- 讲稿时长在目标 ±20% 内；答辩 deck 的 Q&A 已生成。

### 什么时候停手

- 连续两轮 blocking 都是 0，且 warning 集合没有新增：停止。
- 连续两轮分数没有提升：停止，记录已知非阻断问题。
- 同一阶段已经做了 3 轮视觉修改：说明问题在内容结构，不在样式；回到阶段 0。

## 冻结与失效

只有明确知道“哪些报告会失效”，才不会拿旧报告骗自己：

| 改动 | 必须重跑 |
|---|---|
| 改标题、正文、讲稿 | `validate_deck.py`、`qa_design.py`；重渲染受影响页 |
| 改主题颜色、字号、页眉、内容区 | 全量渲染、`validate_deck.py`、`qa_design.py`、PPTX 导出及全部 PPTX 检查 |
| 换图片或改图片文件 | 重渲染引用页、`validate_deck.py`、PPTX 导出 |
| 改渲染器或 exporter | 全量渲染、PPTX 导出、全部检查、相关回归测试 |
| 只改讲稿 | `build_speech.py`、`build_ghost_deck.py`，不需要重渲染 |

## 用门禁收口

```bash
python scripts/quality_gate.py \
  --deck outputs/project/deck.json \
  --theme assets/themes/tsinghua_purple/theme.json \
  --out outputs/project \
  --profile defense \
  --stage final
```

它会读取现有报告，检查新鲜度、页数一致性、错误/警告、设计分、答辩字号、无障碍和实机预览，然后写出：

- `quality_gate.json`：机器可读，给下一轮脚本或模型。
- `quality_gate.md`：按页修复队列和下一动作，给人或弱模型看。

阶段参数：

| `--stage` | 检查范围 | 用途 |
|---|---|---|
| `sample` | 样板页、render warnings、联系表 | 阶段 1 |
| `deck` | 全量渲染、结构校验、设计体检 | 阶段 2 |
| `pptx` | 再加原生 PPTX、无障碍、答辩字号 | 阶段 3 |
| `final` | 再加实机渲染、讲稿、agent brief | 阶段 4 |

如果机器没有 PowerPoint / LibreOffice，可加 `--allow-missing-preview`，把实机预览降为 advisory；交付说明必须写明“未做真实 PPTX 渲染”。

## 案例：清华 GM-APD 12 页答辩

这条流程在 `assets/themes/tsinghua_purple/` 的 12 页样例上实际跑过：

- 阶段 1：封面、目录、三段式框架、实验平台、结论页先出样板，暴露页眉高度和目录行号问题。
- 阶段 2：全量后修公式符号双列、指标卡纵向层级、概念节点 containment、结论卡标题/正文层级。
- 阶段 3：PPTX 与 PNG 对齐后，原生校验 0 错误、无障碍 0 错误，答辩字号 `defense_ready`。
- 阶段 4：PowerPoint 实机 12/12 页成功；设计 86/A；唯一 advisory 是 `chars_per_line = 15.9`，记录而不阻断交付。

这说明门禁的价值不是“追求 100 分”，而是把可接受的 86 分和不可交付的 60 分明确分开。
