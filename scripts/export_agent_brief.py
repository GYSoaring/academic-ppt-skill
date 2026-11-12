from __future__ import annotations

import argparse
import json
from pathlib import Path

from common import load_json, load_theme


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export a portable defense-deck brief for other agents.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--theme", required=True)
    parser.add_argument("--out", required=True)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    deck_path = Path(args.deck).resolve()
    theme = load_theme(args.theme)
    deck = load_json(deck_path)
    meta = deck.get("meta", {})
    sections = deck.get("sections", [])
    slides = deck.get("slides", [])
    colors = theme["colors"]
    layout = theme["layout"]
    brand = theme["brand"]
    title = meta.get("title", "学术答辩汇报")
    output_path = Path(args.out).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    lines = [
        f"# 可移植学术答辩 PPT 任务书：{title}",
        "",
        "请把以下要求当作硬性版式契约执行，不要改成通用商业模板。",
        "",
        "## 交付物",
        "",
        f"1. 一份 {len(slides)} 页、16:9 的学术答辩 PPT，比例与顺序按下方页面清单。",
        "2. 每页 45–90 秒的自然口播讲解词。",
        "3. 一份连续答辩稿，按页面顺序衔接，不要逐条朗读页面文字。",
        "4. 每张论文图、系统图、实验图都要说明：图中是什么、为什么放在这里、支持哪条结论、不能证明什么。",
        "",
        "## 固定视觉规则",
        "",
        f"- 画布比例：16:9，参考尺寸 {theme['canvas']['width']}×{theme['canvas']['height']}。",
        f"- 底色：{colors['paper']}；主色：{colors['blue']}；深色结论带：{colors['blue_dark']}；辅助色：{colors['blue_mid']}、{colors['blue_pale']}。",
        f"- 正文色：{colors['ink']}；弱化文字：{colors['muted']}；分隔线：{colors['line']}。",
        "- 风格：浅纸色背景、深蓝标题、细蓝线、克制的浅蓝面板；不要玻璃拟态、霓虹、3D 大卡片、彩色仪表盘、手写批注、星星或装饰性箭头。",
        "- 封面：没有顶部标题栏，没有左侧小目录。按主题选择居中或编辑式封面；身份资产位置保持全篇一致，标题、汇报人和日期不要与身份层挤在一起。",
        "- 目录页：没有左侧小目录，只放五个大标题，不写解释句，不写页码范围。",
        "- 第 3 页起：顶部标题栏、右上角校名校徽、底部蓝线、左侧五项小目录全部固定。正文不因为文字多而挤压固定区域。",
        f"- 左上双斜杠与章节号 x≈{layout['chapter_x']}，页标题 x≈{layout['title_x']}，标题字号约 {layout['title_size']}。",
        f"- 左侧小目录固定在约 x={layout['sidebar'][0]}–{layout['sidebar'][2]}、y={layout['sidebar'][1]}–{layout['sidebar'][3]}。",
        f"- 右上校名校徽固定在约 x={brand['header_box'][0]}–{brand['header_box'][2]}、y={brand['header_box'][1]}–{brand['header_box'][3]}。",
        "- 字体：中文标题用粗黑体，正文用常规黑体；英文和公式保持清晰，不装饰化。",
        "- 图片：保持原始宽高比，用矩形白底相框放入；不要裁切或重绘论文曲线、坐标轴、图例和实验照片；图片下方不要加“论文 Fig. ×”之类标注。",
        "- 层级：顶部标题只写 4–10 字的主题名；真正的完整判断句放在标题下方的 lead 条，不要用长标题挤压固定页眉。",
        "- 密度：每个内容页至少 3 个内容单元（图、表、流程、指标卡、解释面板或微型机制图），不要只放一个大标题加一张空卡片。",
        "- 视觉锚点：除纯过渡页外，每页至少有一个有信息量的视觉锚点；微型机制图只作机制辅助，不能替代实测图。",
        "- 图文比例：实测图或主图占内容区约 65%–80%，解释栏控制在 20%–35%，不要把密集论文图缩成角落小图。",
        "- 颜色语义：主色只标记重点，红色只用于关键数字、风险或边界；同类数据用同一色相的深浅变化，不要彩虹配色。",
        "- 卡片语言：小圆角、细边框、低阴影、深蓝标题条；不要堆叠多层卡片、玻璃拟态、发光或 3D 装饰。",
        "- 结论页：保留 2–4 条结论，并写出边界与下一步；不要只用“谢谢”占满一页。",
        "",
        "## 内容纪律",
        "",
        "- 每页只回答一个中心问题，先给判断，再给证据，再给边界。",
        "- 不虚构数据、基线、参考文献、硬件资源、实验条件或实时性。",
        "- 跨文献数字必须注明“不同实验条件，不作严格排名”。",
        "- 概念图不能讲成实测结果；仿真、外场实验和工程部署要分开。",
        "- 结论中保留条件，避免“彻底解决”“全面胜出”“国际领先”等口号。",
        "",
        "## 章节",
        "",
    ]
    for index, section in enumerate(sections, start=1):
        lines.append(f"{index}. {section['title']}（id: `{section['id']}`）")
    lines.extend(
        [
            "",
            "## 版式选择与密度契约",
            "",
            "- `figure`：一张实测图或论文原图占主导，右侧放 2–3 条解释；主图应占内容区约 65%–80%。",
            "- `figure_flow`：顶部 3–5 步流程，下方一张主图和侧栏解释；适合“原理 + 机制”。",
            "- `concept`：一张带标签的语义示意图 + 2–3 个解释面板；只用于真实存在的时间顺序、分支、层级、刻度、二维取舍或闭环关系。",
            "- `formula`：一个公式 + 符号表 + 适用条件 + 读法/边界/验证面板；公式页不是装饰性代码块，也不能替代实测证据。",
            "- `grid`：3–6 个证据或机制块，可用 `span` 让系统图或公式跨列；不要全部做成同样大小。",
            "- `balance`：左栏收益、右栏代价、中间一张系统主图；只有真实权衡关系时才使用。",
            "- `pipeline`：4–5 个连续步骤，每步有输入/输出或微型机制图；不要用 3 个空盒子充当流程。",
            "- `compare`：只用于真实的路线、条件或方法对比；不同实验条件必须在表下注明不可严格排名。",
            "- `metrics`：2–4 个关键指标 + 一张主图 + 参数条件表；指标必须带条件和边界。",
            "- `hub`：中心命题 + 4–6 个外围支路；适合分类、机制分叉、研究路线全景，不要把所有页都做成卡片网格。",
            "- `timeline`：3–5 个研究阶段，每个阶段含时间、标题、任务与里程碑；开题、中期和进度汇报优先使用。",
            "- `references`：4–10 条真实文献，两栏排版；作者、标题、期刊/会议与 DOI/URL 必须来自用户材料。",
            "- `closing`：2–4 条结论 + 边界 + 下一步；不要把“谢谢”作为唯一内容。",
            "- 内容页至少有 3 个内容单元；如果只有 1–2 个，合并上一页或拆成两页，不要靠放大标题填空。",
            "- 没有真实图片时，用 `footer` / `output` 写“判断、交付、输出、指标”，不要放与文字无关的 target、histogram、nodes 或 noise 占位图。",
            "- 有真实机制关系但没有实测图时，可用 `concept` 的 `diagram`，并在页面上标明“机制示意 / 研究接口 / 概念示意”；不能把它说成实测结果。",
            "- 公式同时提供 `formula` 纯文本回退和 `formula_latex`；正式渲染优先用 LaTeX 数学排版，不要让生图模型画公式。",
            "- 深色结论带上的所有文字必须切换为浅色；`**强调**` 只表示红色粗体，不自动加下划线。",
            "",
        ]
    )
    lines.extend(["", "## 页面清单与内容", ""])
    for index, slide in enumerate(slides, start=1):
        lines.append(f"### 第 {index:02d} 页｜{slide.get('title', '')}")
        lines.append("")
        lines.append(f"- 版式：`{slide.get('layout', '')}`")
        if slide.get("content_type"):
            lines.append(f"- 内容类型：`{slide['content_type']}`")
        if slide.get("section"):
            lines.append(f"- 所属章节：`{slide['section']}`")
        if slide.get("eyebrow"):
            lines.append(f"- 页眉编号：`{slide['eyebrow']}`")
        if slide.get("lead"):
            lines.append(f"- 本页问题/判断：{slide['lead']}")
        if slide.get("claim"):
            lines.append(f"- 核心主张：{slide['claim']}")
        if slide.get("evidence"):
            lines.append(f"- 证据：{slide['evidence']}")
        if slide.get("boundary"):
            lines.append(f"- 边界：{slide['boundary']}")
        if slide.get("layout") == "grid":
            lines.append(f"- 分栏数：{slide.get('columns', 2)}")
            for block_index, block in enumerate(slide.get("blocks", []), start=1):
                heading = block.get("heading", "")
                body = block.get("body", "")
                bullets = "；".join(block.get("bullets", []))
                lines.append(f"- 信息块 {block_index}「{heading}」：{body or bullets}")
                if block.get("formula"):
                    lines.append(f"- 信息块 {block_index} 公式：`{block['formula']}`")
                if block.get("formula_latex"):
                    lines.append(f"- 信息块 {block_index} LaTeX：`{block['formula_latex']}`")
                if block.get("footer") or block.get("output"):
                    lines.append(
                        f"- 信息块 {block_index} 底栏（{block.get('footer_label', '产出')}）："
                        f"{block.get('footer') or block.get('output')}"
                    )
                if block.get("icon"):
                    lines.append(f"- 信息块 {block_index} 图形元素：`{block['icon']}`")
                if block.get("image"):
                    lines.append(f"- 信息块 {block_index} 图片：`{block['image']}`")
                if block.get("visual"):
                    lines.append(f"- 信息块 {block_index} 微型机制图：`{block['visual']}`")
                if block.get("diagram"):
                    diagram = block["diagram"]
                    lines.append(
                        f"- 信息块 {block_index} 语义示意图：`{diagram.get('kind', 'sequence')}`"
                    )
                    for item_index, item in enumerate(diagram.get("items", []), start=1):
                        lines.append(
                            f"- 信息块 {block_index} 图示节点 {item_index}「{item.get('title', '')}」："
                            f"{item.get('body', '') or item.get('label', '')}"
                        )
                if block.get("stat"):
                    lines.append(f"- 信息块 {block_index} 关键数字：{block['stat']}（{block.get('stat_label', '')}）")
        if slide.get("layout") == "concept":
            diagram = slide.get("diagram") or {}
            lines.append(f"- 示意图类型：`{diagram.get('kind', 'sequence')}`")
            lines.append(f"- 示意图标签：{slide.get('diagram_label', '机制示意')}")
            for item_index, item in enumerate(diagram.get("items", []), start=1):
                lines.append(
                    f"- 示意图节点 {item_index}「{item.get('title', '') or item.get('label', '')}」："
                    f"{item.get('body', '')}"
                )
            if diagram.get("note"):
                lines.append(f"- 示意图边界：{diagram['note']}")
        if slide.get("layout") == "formula":
            lines.append(f"- 公式标签：{slide.get('formula_label', '核心关系')}")
            lines.append(f"- 公式纯文本：`{slide.get('formula', '')}`")
            lines.append(f"- 公式 LaTeX：`{slide.get('formula_latex', '')}`")
            if slide.get("formula_condition"):
                lines.append(f"- 适用条件：{slide['formula_condition']}")
            for item_index, item in enumerate(slide.get("variables", []), start=1):
                lines.append(
                    f"- 符号 {item_index}「{item.get('symbol', '')}」：{item.get('meaning', '')}"
                )
        if slide.get("layout") == "balance":
            for rail_name, rail_label in (("left", "收益/条件"), ("right", "代价/约束")):
                rail = slide.get(rail_name) or {}
                lines.append(f"- {rail_label}栏：{rail.get('heading', '')}")
                for item_index, item in enumerate(rail.get("items", []), start=1):
                    detail = item.get("body", "") or "；".join(item.get("bullets", []))
                    lines.append(f"- {rail_label} {item_index}「{item.get('heading', '')}」：{detail}")
        if slide.get("layout") == "hub":
            center = slide.get("center") or {}
            lines.append(f"- 中心命题：{center.get('label', '')}｜{center.get('value', '')}｜{center.get('note', '')}")
            for item_index, item in enumerate(slide.get("items", []), start=1):
                detail = item.get("body", "") or "；".join(item.get("bullets", []))
                lines.append(f"- 枢纽支路 {item_index}「{item.get('heading', '')}」：{detail}")
                if item.get("visual"):
                    lines.append(f"- 枢纽支路 {item_index} 微型机制图：`{item['visual']}`")
        if slide.get("layout") == "timeline":
            if slide.get("period_note"):
                lines.append(f"- 时间总跨度：{slide['period_note']}")
            for phase_index, phase in enumerate(slide.get("phases", []), start=1):
                detail = phase.get("body", "") or "；".join(phase.get("bullets", []))
                lines.append(
                    f"- 阶段 {phase_index}「{phase.get('period', '')}｜{phase.get('title', '')}」：{detail}"
                )
                if phase.get("milestone"):
                    lines.append(f"- 阶段 {phase_index} 里程碑：{phase['milestone']}")
        if slide.get("layout") == "references":
            for ref_index, reference in enumerate(slide.get("references", []), start=1):
                lines.append(
                    f"- 文献 {ref_index}：{reference.get('authors', '')}｜{reference.get('title', '')}｜"
                    f"{reference.get('venue', '')}｜{reference.get('doi', '')}"
                )
        elif slide.get("layout") not in {"grid", "hub"} and slide.get("columns"):
            for column in slide["columns"]:
                if isinstance(column, dict):
                    heading = column.get("heading", "")
                    bullets = "；".join(column.get("bullets", []))
                    lines.append(f"- 栏目「{heading}」：{bullets}")
                else:
                    lines.append(f"- 表头：{column}")
        if slide.get("image"):
            lines.append(f"- 图片：`{slide['image']}`（角色：{slide.get('image_role', '未注明')}）")
            if slide.get("image_boundary"):
                lines.append(f"- 图片边界：{slide['image_boundary']}")
        if slide.get("panels"):
            for panel in slide["panels"]:
                detail = panel.get("body", "") or "；".join(panel.get("bullets", []))
                lines.append(f"- 图旁解释「{panel.get('heading', '')}」：{detail}")
                if panel.get("visual"):
                    lines.append(f"- 图旁微型机制图：`{panel['visual']}`")
        if slide.get("flow"):
            for flow_index, step in enumerate(slide["flow"], start=1):
                lines.append(f"- 顶部流程 {flow_index}「{step.get('title', '')}」：{step.get('body', '')}")
        if slide.get("steps"):
            for step_index, step in enumerate(slide["steps"], start=1):
                lines.append(f"- 流程 {step_index}「{step.get('title', '')}」：{step.get('body', '')}")
                if step.get("output"):
                    lines.append(f"- 流程 {step_index} 输出：{step['output']}")
                if step.get("visual"):
                    lines.append(f"- 流程 {step_index} 微型机制图：`{step['visual']}`")
        if slide.get("footer_row"):
            for item in slide["footer_row"]:
                lines.append(f"- 页内脚注卡：{item.get('title', '')}｜{item.get('body', '')}")
        if slide.get("metrics"):
            for metric in slide["metrics"]:
                lines.append(f"- 指标：{metric.get('value', '')}｜{metric.get('label', '')}｜{metric.get('note', '')}")
            conditions = slide.get("conditions") or []
            if conditions:
                parts = []
                for condition in conditions:
                    if isinstance(condition, dict):
                        parts.append(f"{condition.get('label', '')}={condition.get('value', '')}")
                    else:
                        parts.append(str(condition))
                lines.append(f"- 实验条件：{'；'.join(parts)}")
            if slide.get("conditions_note"):
                lines.append(f"- 条件边界：{slide['conditions_note']}")
        if slide.get("rows"):
            for row in slide["rows"]:
                lines.append(f"- 对比行：{'｜'.join(str(value) for value in row)}")
            if slide.get("highlight_rows"):
                lines.append(f"- 强调行：{slide['highlight_rows']}")
            if slide.get("notes"):
                lines.append(f"- 表下注：{'；'.join(slide['notes'])}")
        if slide.get("chart"):
            lines.append(f"- 图表结构：`{json.dumps(slide['chart'], ensure_ascii=False)}`")
        if slide.get("image_label") or slide.get("image_caption"):
            lines.append(f"- 图版标签：{slide.get('image_label', '')}｜{slide.get('image_caption', '')}｜badge={slide.get('image_badge', '')}")
        if slide.get("boundary_points"):
            lines.append(f"- 边界分点：{'；'.join(slide['boundary_points'])}")
        if slide.get("next_points"):
            lines.append(f"- 下一步分点：{'；'.join(slide['next_points'])}")
        if slide.get("closing"):
            lines.append(f"- 收束语：{slide['closing']}")
        if slide.get("takeaway"):
            lines.append(f"- 本页结论带：{slide['takeaway']}")
        if slide.get("footnote"):
            lines.append(f"- 页脚注：{slide['footnote']}")
        if slide.get("speech"):
            lines.append(f"- 口播要求：{slide['speech']}")
        lines.append("")
    lines.extend(
        [
            "## 口播写法",
            "",
            "- 像现场汇报，不要像论文摘要。",
            "- 每页先说“这一页在回答什么”，再说“图中看到什么”，然后说“支持什么结论”，最后说“不能证明什么”并过渡到下一页。",
            "- 避免机械排比和长篇定义；首次出现的缩写要展开。",
            "- 论文图必须按图讲，不允许自己补曲线、补数值或把示意图说成实测。",
            "",
            "## 幻灯片 JSON 参考",
            "",
            "```json",
            json.dumps(deck, ensure_ascii=False, indent=2),
            "```",
            "",
        ]
    )
    output_path.write_text("\n".join(lines), encoding="utf-8")
    print(str(output_path))


if __name__ == "__main__":
    main()
