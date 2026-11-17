from __future__ import annotations

import json
from pathlib import Path

from pptx import Presentation

from design import DIAGRAM_BRANCH_TEXT
from common import slide_display_title
from export_editable_pptx import EditableDeckExporter
from validate_pptx import geometry_issues, minimum_editable_text_shapes


SKILL_ROOT = Path(__file__).resolve().parents[1]
THEME = SKILL_ROOT / "assets" / "themes" / "tju_blue" / "theme.json"


def _shape_by_name(slide, name: str):
    return next(shape for shape in slide.shapes if shape.name == name)


def _font_px(shape) -> int:
    return int(shape.text_frame.paragraphs[0].runs[0].font.size.pt * 153.6 / 72)


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
            "eyebrow": "1.1",
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
