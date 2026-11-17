from __future__ import annotations

import json
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE

from design import draw_gradient_panel
from export_editable_pptx import EditableDeckExporter


SKILL_ROOT = Path(__file__).resolve().parents[1]
PKU_THEME = SKILL_ROOT / "assets" / "themes" / "pku_red" / "theme.json"
PKU_COLORS = json.loads(PKU_THEME.read_text(encoding="utf-8"))["colors"]


def _mean_rgb(value: str) -> float:
    value = value.lstrip("#")
    channels = [int(value[index : index + 2], 16) for index in (0, 2, 4)]
    return sum(channels) / len(channels)


def _deck(figure_path: Path) -> dict:
    return {
        "meta": {
            "title": "渐变图形回归",
            "type_profile": "defense",
            "title_mode": "summary",
        },
        "sections": [{"id": "background", "title": "研究背景"}],
        "slides": [
            {
                "id": "01",
                "layout": "bullets",
                "section": "background",
                "eyebrow": "1.1",
                "title": "纯色卡片",
                "lead": "正文卡片使用同色系的分级纯色，重点矩形才使用向右变浅的克制渐变。",
                "columns": [
                    {"heading": "机制", "bullets": ["首触发竞争"]},
                    {"heading": "边界", "bullets": ["只验证视觉规则"]},
                ],
                "claim": "渐变填充必须同时写入 PNG 与原生 PPTX。",
                "evidence": "颜色导出回归测试。",
                "boundary": "仅验证图形填充。",
                "takeaway": "正文卡片保持纯色，结论条使用深红纯色。",
                "speech": "图形色阶回归。",
            },
            {
                "id": "02",
                "layout": "outline",
                "section": "background",
                "title": "汇报提纲",
                "speech": "目录保持纯色。",
            },
            {
                "id": "03",
                "layout": "compare",
                "section": "background",
                "eyebrow": "1.2",
                "title": "表格保持纯色",
                "lead": "表头与单元格不得继承卡片渐变。",
                "columns": ["项目", "值"],
                "rows": [["卡片", "渐变"], ["表格", "纯色"]],
                "claim": "表格是结构化读数。",
                "evidence": "单元格使用静态色。",
                "boundary": "不验证数据内容。",
                "takeaway": "表格不参与卡片渐变。",
                "speech": "表格纯色回归。",
            },
            {
                "id": "04",
                "layout": "metrics",
                "section": "background",
                "eyebrow": "1.4",
                "title": "指标卡保持纯色",
                "lead": "指标卡、图卡与条件面板都不继承正文卡片渐变。",
                "metrics": [
                    {"value": "6.91%", "label": "相对误差", "note": "静态色"},
                    {"value": "1.16E-4", "label": "系统常数", "note": "静态色"},
                ],
                "image": str(figure_path),
                "image_label": "论文原图",
                "image_caption": "仅做等比嵌入。",
                "conditions_title": "实验条件",
                "conditions": [{"label": "距离", "value": "3.67 m"}],
                "takeaway": "结构化读数不使用渐变。",
                "speech": "指标卡纯色回归。",
            },
            {
                "id": "05",
                "layout": "figure",
                "section": "background",
                "eyebrow": "1.3",
                "title": "论文图保持原样",
                "lead": "论文图卡片与标签保持静态。",
                "image": str(figure_path),
                "image_label": "论文原图",
                "image_caption": "仅做等比嵌入。",
                "panels": [
                    {"heading": "读法", "body": "图内像素不改色。"},
                    {"heading": "边界", "body": "图卡不铺渐变。"},
                ],
                "claim": "图形语义优先于装饰。",
                "evidence": "图卡使用纯白底。",
                "boundary": "不重绘论文图。",
                "speech": "论文图静态回归。",
            },
            {
                "id": "06",
                "layout": "concept",
                "section": "background",
                "eyebrow": "1.5",
                "title": "重点节点",
                "lead": "重点矩形从纯红起始，向右逐步提亮。",
                "diagram_label": "机制示意",
                "diagram": {
                    "kind": "sequence",
                    "highlight": 1,
                    "items": [
                        {"title": "纯色", "body": "正文白底"},
                        {"title": "重点", "body": "向右变浅"},
                        {"title": "回归", "body": "层级清楚"},
                    ],
                },
                "panels": [{"heading": "规则", "body": "渐变只服务重点，不铺满正文。"}],
                "takeaway": "重点矩形从自身纯色开始，向右变浅。",
                "speech": "重点节点渐变回归。",
            },
        ],
    }


def _export_pku(work_tmp: Path) -> tuple[Path, EditableDeckExporter]:
    figure_path = work_tmp / "figure.png"
    figure = Image.new("RGB", (420, 280), "white")
    figure_draw = ImageDraw.Draw(figure)
    figure_draw.rectangle((54, 42, 366, 238), outline="black", width=6)
    figure_draw.line((54, 42, 366, 238), fill="black", width=4)
    figure.save(figure_path)
    deck_path = work_tmp / "deck.json"
    deck_path.write_text(json.dumps(_deck(figure_path), ensure_ascii=False), encoding="utf-8")
    out = work_tmp / "out"
    exporter = EditableDeckExporter(deck_path, PKU_THEME, out, profile="defense")
    exporter.export("gradient.pptx")
    return out / "gradient.pptx", exporter


def test_gradient_panel_fades_left_to_right() -> None:
    image = Image.new("RGB", (120, 40), "white")
    draw = ImageDraw.Draw(image)
    draw_gradient_panel(
        image,
        draw,
        (10, 5, 110, 35),
        (10, 0, 0),
        (250, 250, 250),
        radius=0,
        width=0,
        direction="horizontal",
    )
    assert image.getpixel((12, 20))[0] < image.getpixel((107, 20))[0]


def test_pku_body_cards_stay_flat_while_accent_gradient_starts_from_solid_red(work_tmp: Path) -> None:
    pptx_path, exporter = _export_pku(work_tmp)
    slides = list(Presentation(str(pptx_path)).slides)
    first = slides[0]
    headers = [shape for shape in first.shapes if shape.name == "module_header_fill"]
    sidebar_active = next(shape for shape in first.shapes if shape.name == "sidebar_active")
    sidebar_text = next(shape for shape in first.shapes if shape.name == "sidebar_background")

    assert len(headers) == 2
    for shape in headers:
        assert shape.fill.type == MSO_FILL.SOLID
        assert str(shape.fill.fore_color.rgb) == PKU_COLORS["blue"].lstrip("#")
    assert all(
        str(run.font.color.rgb) == PKU_COLORS["white"].lstrip("#")
        for shape in first.shapes
        if shape.name == "module_heading"
        for paragraph in shape.text_frame.paragraphs
        for run in paragraph.runs
        if run.font.color.rgb is not None
    )

    assert sidebar_active.fill.type == MSO_FILL.SOLID
    assert str(sidebar_active.fill.fore_color.rgb) == PKU_COLORS["blue"].lstrip("#")
    assert "gradFill" not in sidebar_active.fill._xPr.xml
    active_text_color = sidebar_text.text_frame.paragraphs[0].runs[0].font.color.rgb
    assert str(active_text_color) == PKU_COLORS["white"].lstrip("#")

    takeaway = next(
        shape
        for shape in first.shapes
        if shape.name == "takeaway" and shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
    )
    assert takeaway.fill.type == MSO_FILL.SOLID
    assert str(takeaway.fill.fore_color.rgb) == PKU_COLORS["blue"].lstrip("#")
    assert "gradFill" not in takeaway.fill._xPr.xml

    outline = slides[1]
    outline_rows = [shape for shape in outline.shapes if shape.name.startswith("outline_row_")]
    for shape in outline_rows:
        assert shape.fill.type == MSO_FILL.SOLID
        assert str(shape.fill.fore_color.rgb) == PKU_COLORS["white"].lstrip("#")
    assert not any(
        shape.fill.type == MSO_FILL.GRADIENT
        for shape in outline.shapes
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
    )

    table_slide = slides[2]
    table = next(shape.table for shape in table_slide.shapes if shape.has_table)
    for row in table.rows:
        for cell in row.cells:
            assert cell.fill.type == MSO_FILL.SOLID

    metrics_slide = slides[3]
    metric_cards = [
        shape for shape in metrics_slide.shapes if shape.name.startswith("metric_card_")
    ]
    assert len(metric_cards) == 2
    for shape in metric_cards:
        assert shape.fill.type == MSO_FILL.SOLID
    assert str(metric_cards[0].fill.fore_color.rgb) == PKU_COLORS["blue_wash"].lstrip("#")
    assert str(metric_cards[1].fill.fore_color.rgb) == PKU_COLORS["white"].lstrip("#")

    figure_slide = slides[4]
    figure_card = next(shape for shape in figure_slide.shapes if shape.name == "figure_card")
    figure_label = next(shape for shape in figure_slide.shapes if shape.name == "figure_label_fill")
    assert figure_card.fill.type == MSO_FILL.SOLID
    assert str(figure_card.fill.fore_color.rgb) == PKU_COLORS["white"].lstrip("#")
    assert figure_label.fill.type == MSO_FILL.SOLID
    assert str(figure_label.fill.fore_color.rgb) == PKU_COLORS["blue_pale"].lstrip("#")

    accent_slide = slides[5]
    focus_node = next(
        shape for shape in accent_slide.shapes if shape.name == "diagram_sequence_node_2"
    )
    assert focus_node.fill.type == MSO_FILL.GRADIENT
    focus_xml = focus_node.fill._xPr.xml
    assert '<a:lin scaled="0" ang="0"/>' in focus_xml
    focus_stops = focus_node.fill.gradient_stops
    assert str(focus_stops[0].color.rgb) == PKU_COLORS["blue"].lstrip("#")
    assert str(focus_stops[1].color.rgb) == PKU_COLORS["blue_mid"].lstrip("#")
    assert _mean_rgb(PKU_COLORS["blue_mid"]) > _mean_rgb(PKU_COLORS["blue"])
    for index in (1, 3):
        node = next(
            shape for shape in accent_slide.shapes
            if shape.name == f"diagram_sequence_node_{index}"
        )
        assert node.fill.type == MSO_FILL.SOLID
    assert not any(shape.name == "soft_gradient_panel" for shape in accent_slide.shapes)


def test_pku_section_numeral_is_flat(work_tmp: Path) -> None:
    _, exporter = _export_pku(work_tmp)
    number_path = exporter.asset_dir / "section_number_01.png"
    exporter._build_number_asset(number_path, "01", 420, 300)
    image = Image.open(number_path).convert("RGBA")
    pixels = np.array(image)
    opaque = pixels[pixels[..., 3] == 255][:, :3]
    opaque_colors = {tuple(int(channel) for channel in row) for row in np.unique(opaque, axis=0)}
    assert len(opaque_colors) == 1
    assert opaque_colors.pop() == tuple(
        int(PKU_COLORS["blue_bright"].lstrip("#")[index : index + 2], 16)
        for index in (0, 2, 4)
    )
