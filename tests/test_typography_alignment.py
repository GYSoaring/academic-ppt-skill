from __future__ import annotations

import json
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation

from design import DIAGRAM_BRANCH_TEXT
from common import resolve_font, slide_display_title, text_width, wrap_cjk_text
from export_editable_pptx import EditableDeckExporter
from validate_pptx import geometry_issues, minimum_editable_text_shapes


SKILL_ROOT = Path(__file__).resolve().parents[1]
THEME = SKILL_ROOT / "assets" / "themes" / "tju_blue" / "theme.json"
EMU_PER_PX = 914400.0 / 153.6


def _shape_by_name(slide, name: str):
    return next(shape for shape in slide.shapes if shape.name == name)


def _font_px(shape) -> int:
    return int(shape.text_frame.paragraphs[0].runs[0].font.size.pt * 153.6 / 72)


def _shape_px(value: float) -> float:
    return float(value) / EMU_PER_PX


def _run(shape):
    return next(
        run
        for paragraph in shape.text_frame.paragraphs
        for run in paragraph.runs
        if run.text
    )


def _deck(outline: bool) -> dict:
    slide = (
        {
            "id": "01",
            "layout": "outline",
            "title": "汇报提纲",
            "speech": "目录页。",
        }
        if outline
        else {
            "id": "01",
            "layout": "bullets",
            "section": "background",
            "eyebrow": "1.2",
            "title": "探测链与退化",
            "lead": "测试页眉字号与基线。",
            "columns": [{"heading": "机制", "bullets": ["首触发竞争"]}],
            "claim": "页眉编号与标题必须使用同一字号和基线。",
            "evidence": "版式回归测试。",
            "boundary": "仅验证版式。",
            "speech": "页眉测试。",
        }
    )
    return {
        "meta": {
            "title": "版式回归",
            "type_profile": "defense",
            "title_mode": "summary",
        },
        "sections": [{"id": "background", "title": "研究背景"}],
        "slides": [slide],
    }


def _sequence_deck() -> dict:
    return {
        "meta": {
            "title": "概念图回归",
            "type_profile": "defense",
            "title_mode": "summary",
        },
        "sections": [{"id": "method", "title": "方法"}],
        "slides": [
            {
                "id": "01",
                "layout": "concept",
                "section": "method",
                "eyebrow": "3.1",
                "title": "多尺度处理链",
                "lead": "验证原生概念图节点正文的几何位置。",
                "diagram_label": "处理链",
                "diagram": {
                    "kind": "sequence",
                    "items": [
                        {"title": "1D 时间", "body": "逐像素反算事件率，先隔离明显噪声频带。"},
                        {"title": "3D 局部", "body": "联合局部空间窗做高维傅里叶变换。"},
                        {"title": "4D 非局部", "body": "用深度与强度图搜索相似立方体。"},
                        {"title": "多尺度融合", "body": "融合多个分块尺度的深度与强度估计。"},
                    ],
                    "highlight": 2,
                    "note": "机制示意，不代表实测数据。",
                },
                "speech": "概念图几何回归。",
            }
        ],
    }


def _typography_deck() -> dict:
    return {
        "meta": {
            "title": "字体与行高回归",
            "type_profile": "defense",
            "title_mode": "summary",
        },
        "sections": [{"id": "background", "title": "研究背景"}],
        "slides": [
            {
                "id": "01",
                "layout": "bullets",
                "section": "background",
                "eyebrow": "1.2",
                "title": "字体与行高",
                "lead": "两行导语必须按真实行数预留高度，不能靠缩小字号把内容勉强塞进单行面板。",
                "columns": [{"heading": "机制", "bullets": ["首触发竞争"]}],
                "claim": "粗体只改变字重，不切换字体家族。",
                "evidence": "原生 PPTX 版式回归。",
                "boundary": "仅验证版式。",
                "footnote": "来源：版式回归。",
                "speech": "字体、行高与脚注回归。",
            }
        ],
    }


def _concept_panel_deck() -> dict:
    return {
        "meta": {
            "title": "模块标题回归",
            "type_profile": "defense",
            "title_mode": "summary",
        },
        "sections": [{"id": "mechanism", "title": "机理"}],
        "slides": [
            {
                "id": "01",
                "layout": "concept",
                "section": "mechanism",
                "eyebrow": "2.1",
                "title": "三类场景",
                "lead": "题号、标题与正文必须形成稳定的字号阶梯。",
                "diagram_label": "场景分类",
                "diagram": {
                    "kind": "sequence",
                    "items": [
                        {"title": "门外", "body": "烟雾位于选通门外。"},
                        {"title": "门内", "body": "烟雾与目标峰重叠。"},
                        {"title": "全路径", "body": "背景与回波同时污染。"},
                    ],
                    "highlight": 1,
                    "note": "机制示意。",
                },
                "diagram_width_ratio": 0.58,
                "panels": [
                    {"heading": "位置决定回波", "body": "同样是烟雾，落点不同会形成衰减、双峰或全局污染。"},
                    {"heading": "论文重点", "body": "第二类聚焦局部双峰，第三类聚焦全局噪声。"},
                ],
                "speech": "模块题号与标题层级回归。",
            }
        ],
    }


def _section_deck() -> dict:
    return {
        "meta": {
            "title": "章节页回归",
            "type_profile": "defense",
            "title_mode": "summary",
        },
        "sections": [{"id": "method", "title": "方法"}],
        "slides": [
            {
                "id": "01",
                "layout": "section",
                "section": "method",
                "title": "从回波链路到探测边界",
                "label_en": "METHOD",
                "speech": "章节页标题语义回归。",
            }
        ],
    }


def _export(work_tmp: Path, deck_data: dict) -> Path:
    deck_path = work_tmp / "deck.json"
    deck_path.write_text(json.dumps(deck_data, ensure_ascii=False), encoding="utf-8")
    theme = json.loads(THEME.read_text(encoding="utf-8"))
    theme.setdefault("layout", {})["outline_style"] = "index"
    theme_path = work_tmp / "theme.json"
    theme_path.write_text(json.dumps(theme, ensure_ascii=False), encoding="utf-8")
    out = work_tmp / "out"
    exporter = EditableDeckExporter(deck_path, theme_path, out, profile="defense")
    exporter.export("check.pptx")
    return out / "check.pptx"


def test_outline_number_and_title_share_size_and_baseline(work_tmp: Path) -> None:
    pptx_path = _export(work_tmp, _deck(outline=True))
    slide = Presentation(str(pptx_path)).slides[0]
    number = _shape_by_name(slide, "outline_number_1")
    title = _shape_by_name(slide, "outline_1")
    assert _font_px(number) == _font_px(title)
    assert number.top == title.top


def test_header_eyebrow_and_title_share_size_and_baseline(work_tmp: Path) -> None:
    pptx_path = _export(work_tmp, _deck(outline=False))
    slide = Presentation(str(pptx_path)).slides[0]
    eyebrow = _shape_by_name(slide, "chapter")
    title = _shape_by_name(slide, "header_title")
    assert _font_px(eyebrow) == _font_px(title)
    assert eyebrow.top == title.top


def test_header_eyebrow_is_atomic_and_has_horizontal_gutter(work_tmp: Path) -> None:
    pptx_path = _export(work_tmp, _deck(outline=False))
    slide = Presentation(str(pptx_path)).slides[0]
    eyebrow = _shape_by_name(slide, "chapter")
    assert eyebrow.text == "1.2"
    assert "\n" not in eyebrow.text
    assert eyebrow.text_frame.word_wrap is False
    assert _shape_px(eyebrow.width) >= 100


def test_module_index_matches_heading_and_never_clips(work_tmp: Path) -> None:
    pptx_path = _export(work_tmp, _concept_panel_deck())
    presentation = Presentation(str(pptx_path))
    slide = presentation.slides[0]
    indices = [shape for shape in slide.shapes if shape.name == "module_index"]
    headings = [shape for shape in slide.shapes if shape.name == "module_heading"]
    assert indices and len(indices) == len(headings)
    for index, heading in zip(indices, headings):
        assert index.text in {"01", "02", "03"}
        assert "\n" not in index.text
        assert index.text_frame.word_wrap is False
        assert _shape_px(index.width) >= 60
        assert _font_px(index) == _font_px(heading)
        assert index.top == heading.top
    assert geometry_issues(slide, presentation.slide_width, presentation.slide_height) == []


def test_concept_panel_heading_outranks_body(work_tmp: Path) -> None:
    pptx_path = _export(work_tmp, _concept_panel_deck())
    slide = Presentation(str(pptx_path)).slides[0]
    headings = [shape for shape in slide.shapes if shape.name == "module_heading"]
    bodies = [shape for shape in slide.shapes if shape.name == "concept_panel_body"]
    assert headings and bodies
    assert min(_font_px(shape) for shape in headings) > max(_font_px(shape) for shape in bodies)


def test_two_line_lead_and_footnote_reserve_vertical_space(work_tmp: Path) -> None:
    pptx_path = _export(work_tmp, _typography_deck())
    slide = Presentation(str(pptx_path)).slides[0]
    lead = _shape_by_name(slide, "lead_claim")
    footnote = _shape_by_name(slide, "footnote")
    assert lead.text.count("\n") == 1
    assert _shape_px(lead.height) >= 2 * _font_px(lead) * 1.2 - 2
    assert "\n" not in footnote.text
    assert footnote.text_frame.word_wrap is False
    theme = json.loads(THEME.read_text(encoding="utf-8"))
    footer_y = float(theme["layout"]["footer_y"])
    assert _shape_px(footnote.top + footnote.height) <= footer_y


def test_cjk_punctuation_hang_does_not_exceed_measured_width() -> None:
    text = "少帧时不再继续堆时间，而是把单像素直方图扩展为时空频非局部数据。"
    theme = json.loads(THEME.read_text(encoding="utf-8"))
    font = resolve_font(theme["fonts"]["title"], 56)
    draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
    max_width = 1778
    lines = wrap_cjk_text(draw, text, font, max_width)
    assert len(lines) == 2
    assert all(text_width(draw, line, font) <= max_width for line in lines)


def test_bold_text_does_not_switch_font_family(work_tmp: Path) -> None:
    pptx_path = _export(work_tmp, _typography_deck())
    slide = Presentation(str(pptx_path)).slides[0]
    bold_run = _run(_shape_by_name(slide, "lead_claim"))
    regular_run = _run(_shape_by_name(slide, "footnote"))
    assert bold_run.font.bold is True
    assert regular_run.font.bold is not True
    assert bold_run.font.name
    assert bold_run.font.name == regular_run.font.name


def test_geometry_rejects_wrapped_module_index(work_tmp: Path) -> None:
    pptx_path = _export(work_tmp, _concept_panel_deck())
    presentation = Presentation(str(pptx_path))
    slide = presentation.slides[0]
    index = next(shape for shape in slide.shapes if shape.name == "module_index")
    index.text = "0\n1"
    issues = geometry_issues(slide, presentation.slide_width, presentation.slide_height)
    assert any("Atomic index" in issue for issue in issues)


def test_section_title_uses_nonempty_title_placeholder(work_tmp: Path) -> None:
    pptx_path = _export(work_tmp, _section_deck())
    presentation = Presentation(str(pptx_path))
    slide = presentation.slides[0]
    title = slide.shapes.title
    assert title is not None
    assert title.text == "从回波链路到探测边界"
    assert title.name == "section_title"
    assert geometry_issues(slide, presentation.slide_width, presentation.slide_height) == []


def test_section_title_and_minimum_text_shapes_use_divider_contract() -> None:
    section = {
        "layout": "section",
        "title": "机理建模",
        "section_title": "从回波链路到探测边界",
    }
    assert slide_display_title(section) == "从回波链路到探测边界"
    assert minimum_editable_text_shapes("section") == 1
    assert minimum_editable_text_shapes("figure") == 5
    assert minimum_editable_text_shapes("cover") == 3


def test_branch_concept_has_a_coherent_type_ladder() -> None:
    assert DIAGRAM_BRANCH_TEXT["root_title"] > DIAGRAM_BRANCH_TEXT["root_body"]
    assert DIAGRAM_BRANCH_TEXT["item_title"] >= DIAGRAM_BRANCH_TEXT["item_body"]
    assert DIAGRAM_BRANCH_TEXT["item_body"] >= 38


def test_sequence_node_body_stays_inside_node(work_tmp: Path) -> None:
    pptx_path = _export(work_tmp, _sequence_deck())
    presentation = Presentation(str(pptx_path))
    slide = presentation.slides[0]
    nodes = [shape for shape in slide.shapes if shape.name.startswith("diagram_sequence_node_")]
    indices = [shape for shape in slide.shapes if shape.name == "diagram_index"]
    bodies = [shape for shape in slide.shapes if shape.name == "diagram_node_body"]
    assert nodes and len(nodes) == len(bodies)
    assert indices and all("\n" not in shape.text for shape in indices)
    assert geometry_issues(slide, presentation.slide_width, presentation.slide_height) == []
    for node, body in zip(nodes, bodies):
        assert body.width > 0 and body.height > 0
        assert body.top >= node.top
        assert body.top + body.height <= node.top + node.height


def test_metrics_note_has_positive_geometry(work_tmp: Path) -> None:
    deck_data = {
        "meta": {
            "title": "指标页回归",
            "type_profile": "defense",
            "title_mode": "summary",
        },
        "sections": [{"id": "result", "title": "结果"}],
        "slides": [
            {
                "id": "01",
                "layout": "metrics",
                "section": "result",
                "eyebrow": "3.2",
                "title": "指标与条件",
                "lead": "验证条件注释框的宽度为正且位于页面内。",
                "image": str(SKILL_ROOT / "assets" / "themes" / "tju_blue" / "brand_lockup.png"),
                "image_label": "结果图",
                "image_caption": "版式回归素材。",
                "metrics": [
                    {"value": "+20.9%", "label": "目标复原度", "note": "相对传统模型"},
                    {"value": "−8.2%", "label": "测距误差", "note": "同条件统计"},
                ],
                "conditions_title": "实验条件",
                "conditions": [
                    {"label": "AL", "value": "1.8"},
                    {"label": "SBR", "value": "0.02"},
                ],
                "conditions_note": "结论限于室内可重复条件，不能直接外推到城市与海雾场景。",
                "speech": "指标页几何回归。",
            }
        ],
    }
    pptx_path = _export(work_tmp, deck_data)
    presentation = Presentation(str(pptx_path))
    slide = presentation.slides[0]
    note = _shape_by_name(slide, "conditions_note")
    assert note.left >= 0 and note.top >= 0
    assert note.width > 0 and note.height > 0
    assert note.left + note.width <= presentation.slide_width


def test_closing_card_title_outranks_body(work_tmp: Path) -> None:
    deck_data = {
        "meta": {
            "title": "结论页回归",
            "type_profile": "defense",
            "title_mode": "summary",
        },
        "sections": [{"id": "conclusion", "title": "结论"}],
        "slides": [
            {
                "id": "01",
                "layout": "closing",
                "section": "conclusion",
                "eyebrow": "5.5",
                "title": "边界与展望",
                "lead": "验证结论卡片的标题层级。",
                "boundary": "现有结论有明确适用边界。",
                "boundary_points": ["单一场景", "统一口径"],
                "next": "下一步先统一评价口径。",
                "next_points": ["补复杂场景", "评估实时性"],
                "closing": "谢谢各位老师，请批评指正",
                "speech": "结论页层级回归。",
            }
        ],
    }
    pptx_path = _export(work_tmp, deck_data)
    slide = Presentation(str(pptx_path)).slides[0]
    title = _shape_by_name(slide, "closing_card_title_1")
    body = _shape_by_name(slide, "closing_card_body_1")
    assert _font_px(title) > _font_px(body)
    assert title.left == body.left


def test_defense_condition_grid_keeps_label_and_value_on_separate_bands(work_tmp: Path) -> None:
    deck_data = {
        "meta": {"title": "条件网格回归", "type_profile": "defense", "title_mode": "summary"},
        "sections": [{"id": "result", "title": "结果"}],
        "slides": [
            {
                "id": "01",
                "layout": "metrics",
                "section": "result",
                "eyebrow": "3.2",
                "title": "指标与条件",
                "lead": "验证防御版条件网格不会把标签和值叠在一起。",
                "metrics": [{"value": "+20.9%", "label": "目标复原度", "note": "相对传统模型"}],
                "conditions_title": "实验条件",
                "conditions": [
                    {"label": "衰减长度", "value": "AL = 1.8"},
                    {"label": "信背比", "value": "SBR = 0.02"},
                    {"label": "波长", "value": "1064 nm"},
                    {"label": "阵元", "value": "64 × 64"},
                ],
                "conditions_note": "仅限室内条件。",
                "speech": "条件网格回归。",
            }
        ],
    }
    pptx_path = _export(work_tmp, deck_data)
    slide = Presentation(str(pptx_path)).slides[0]
    labels = {
        shape.text: shape
        for shape in slide.shapes
        if shape.has_text_frame and shape.text in {"衰减长度", "信背比", "波长", "阵元"}
    }
    assert len(labels) == 4
    for index, label in enumerate(["衰减长度", "信背比", "波长", "阵元"], start=1):
        value = _shape_by_name(slide, f"condition_value_{index}")
        assert "\n" not in value.text
        assert labels[label].top + labels[label].height <= value.top + 1
    named_labels = [
        _shape_by_name(slide, f"condition_label_{index}")
        for index in range(1, 5)
    ]
    assert all("\n" not in shape.text for shape in named_labels)
    assert all(shape.text_frame.word_wrap is False for shape in named_labels)
    presentation = Presentation(str(pptx_path))
    assert geometry_issues(
        slide,
        presentation.slide_width,
        presentation.slide_height,
        "defense",
    ) == []


def test_figure_label_and_caption_are_native_and_inside_the_card(work_tmp: Path) -> None:
    deck_data = {
        "meta": {"title": "图卡回归", "type_profile": "defense", "title_mode": "summary"},
        "sections": [{"id": "background", "title": "背景"}],
        "slides": [
            {
                "id": "01",
                "layout": "figure",
                "section": "background",
                "eyebrow": "1.1",
                "title": "探测链",
                "lead": "验证图片卡的标签、来源角标和题注都留在卡片内。",
                "image": str(SKILL_ROOT / "assets" / "themes" / "tju_blue" / "brand_lockup.png"),
                "image_label": "GM-APD 时序",
                "image_badge": "机制图",
                "image_caption": "多周期累积形成飞行时间直方图。",
                "panels": [{"heading": "要点", "body": "标签与题注不得浮出卡片。"}],
                "speech": "图卡回归。",
            }
        ],
    }
    pptx_path = _export(work_tmp, deck_data)
    presentation = Presentation(str(pptx_path))
    slide = presentation.slides[0]
    label = _shape_by_name(slide, "figure_label")
    badge = _shape_by_name(slide, "figure_badge")
    caption = _shape_by_name(slide, "figure_caption")
    assert label.text == "GM-APD 时序"
    assert badge.text == "机制图"
    assert "\n" not in caption.text
    assert label.top > 0 and badge.top > 0
    assert caption.top + caption.height <= presentation.slide_height
    card = _shape_by_name(slide, "figure_card")
    for chrome in (label, badge, caption):
        assert chrome.left >= card.left
        assert chrome.top >= card.top
        assert chrome.left + chrome.width <= card.left + card.width
        assert chrome.top + chrome.height <= card.top + card.height
    assert geometry_issues(
        slide,
        presentation.slide_width,
        presentation.slide_height,
        "defense",
    ) == []


def test_geometry_rejects_condition_label_overlap_and_value_wrap(work_tmp: Path) -> None:
    deck_data = {
        "meta": {
            "title": "Condition geometry",
            "type_profile": "defense",
            "title_mode": "summary",
        },
        "sections": [{"id": "results", "title": "Results"}],
        "slides": [
            {
                "id": "01",
                "layout": "metrics",
                "section": "results",
                "eyebrow": "1.1",
                "title": "Metrics",
                "lead": "A geometry invariant test.",
                "metrics": [
                    {"value": "+20.9%", "label": "Recovery", "note": "Relative"}
                ],
                "conditions_title": "Conditions",
                "conditions": [
                    {"label": "Alpha", "value": "AL = 1.8"},
                    {"label": "SBR", "value": "0.02"},
                    {"label": "Wavelength", "value": "1064 nm"},
                    {"label": "Array", "value": "64 x 64"},
                ],
                "conditions_note": "Indoor conditions only.",
                "speech": "Condition geometry regression.",
            }
        ],
    }
    pptx_path = _export(work_tmp, deck_data)
    presentation = Presentation(str(pptx_path))
    slide = presentation.slides[0]
    label = _shape_by_name(slide, "condition_label_1")
    value = _shape_by_name(slide, "condition_value_1")
    label.top = value.top
    value.text = "AL\n1.8"
    issues = geometry_issues(
        slide,
        presentation.slide_width,
        presentation.slide_height,
        "defense",
    )
    assert any("Condition label overlaps its value" in issue for issue in issues)
    assert any("Condition value wrapped across lines" in issue for issue in issues)


def test_geometry_rejects_figure_chrome_outside_card(work_tmp: Path) -> None:
    deck_data = {
        "meta": {
            "title": "Figure chrome geometry",
            "type_profile": "defense",
            "title_mode": "summary",
        },
        "sections": [{"id": "background", "title": "Background"}],
        "slides": [
            {
                "id": "01",
                "layout": "figure",
                "section": "background",
                "eyebrow": "1.1",
                "title": "Detection",
                "lead": "A figure chrome invariant test.",
                "image": str(SKILL_ROOT / "assets" / "themes" / "tju_blue" / "brand_lockup.png"),
                "image_label": "Timing",
                "image_badge": "Mechanism",
                "image_caption": "A short caption.",
                "panels": [{"heading": "Point", "body": "Chrome stays in the card."}],
                "speech": "Figure chrome geometry regression.",
            }
        ],
    }
    pptx_path = _export(work_tmp, deck_data)
    presentation = Presentation(str(pptx_path))
    slide = presentation.slides[0]
    card = _shape_by_name(slide, "figure_card")
    caption = _shape_by_name(slide, "figure_caption")
    caption.top = card.top - int(8 * EMU_PER_PX)
    caption.text = "Line one\nLine two"
    issues = geometry_issues(
        slide,
        presentation.slide_width,
        presentation.slide_height,
        "defense",
    )
    assert any("Figure-card chrome must be single-line" in issue for issue in issues)
    assert any("Figure-card chrome escaped its card" in issue for issue in issues)
