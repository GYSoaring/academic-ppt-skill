# Defense Outline Patterns

Use five or six major sections. The outline page and mini-directory must use the same names.

For every content page, write two layers:

- `title`: a short topic label for the fixed header (normally 4–10 Chinese characters).
- `lead`: a complete sentence stating what the page proves or asks.

Read only the `lead` sentences in sequence. If they do not tell the argument, fix the outline before rendering.

If the deck was built from a source pack, keep the `figure.id` from `figures.json` in the slide's `source_note` or `footnote` so a reviewer can trace the visual back to its PDF/DOCX/PPTX source page.

## Thesis Defense

1. 研究背景与问题
2. 相关工作与技术路线
3. 方法与系统设计
4. 实验与结果
5. 讨论与局限
6. 总结与展望

Suitable page flow: cover, outline, background, problem definition, related work, method overview, key mechanism, system implementation, experimental setup, main results, comparison, ablation or boundary, limitations, conclusion, thanks.

## Opening Report

1. 研究背景
2. 问题与目标
3. 研究现状
4. 技术方案
5. 研究计划
6. 预期成果与风险

Emphasize feasibility, novelty, milestones, required resources, and risk controls. Do not present expected results as completed evidence.

## Midterm Review

1. 任务与目标
2. 已完成工作
3. 阶段性结果
4. 当前问题
5. 下一步计划

Separate completed work from planned work. Give each claimed result a date, dataset, or experiment condition.

## Course Presentation

1. 选题背景
2. 问题定义
3. 方法或案例
4. 结果分析
5. 结论与反思

Keep the theory load appropriate for the course. Explain unfamiliar terms at first use.

## Competition Defense

1. 需求与痛点
2. 方案与创新
3. 系统演示
4. 验证与指标
5. 应用与价值
6. 团队与计划

Do not let business claims outrun technical evidence. Separate demonstrated capability, projected capability, and assumptions.

## Literature Review

1. 研究背景
2. 问题与分类
3. 代表性路线
4. 横向比较
5. 趋势与空白
6. 结论

Use a taxonomy figure early. Compare methods by information source and cost, not only by headline performance.

## Page Roles

- `cover`: title, identity, presenter, advisor, date.
- `outline`: only section names.
- `section`: a deliberate rhythm break between modules; no new evidence.
- `bullets`: definition, criteria, or a small set of parallel points.
- `figure`: paper figure, system diagram, or experimental setup with explanation panels.
- `figure_flow`: a mechanism sequence above a main figure; use when both are necessary.
- `pipeline`: processing chain or method workflow.
- `compare`: methods, conditions, or design tradeoffs.
- `metrics`: a small number of measured values with conditions.
- `hub`: a shared problem or object with four to six parallel branches.
- `closing`: conclusion, limitations, next step, and thanks.

## Density Check

Each content page should have:

1. One short header title.
2. One complete-sentence lead claim.
3. One dominant visual anchor.
4. Two to four interpretation units.
5. One takeaway, boundary, or transition.

If a page has fewer than three content units, merge it with an adjacent page. If it has more than five, split it.
