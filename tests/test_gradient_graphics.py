from __future__ import annotations

import json
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.enum.dml import MSO_FILL

from design import draw_gradient_panel
from export_editable_pptx import EditableDeckExporter


SKILL_ROOT = Path(__file__).resolve().parents[1]
PKU_THEME = SKILL_ROOT / "assets" / "themes" / "pku_red" / "theme.json"
PKU_COLORS = json.loads(PKU_THEME.read_text(encoding="utf-8"))["colors"]


def _mean_rgb(value: str) -> float:
    value = value.lstrip("#")
    channels = [int(value[index : index + 2], 16) for index in (0, 2, 4)]
    return sum(channels) / len(channels)


def _deck() -> dict:
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
                "title": "渐变卡片",
                "lead": "标题条与浅色卡片应使用同一方向的柔化渐变。",
                "columns": [
                    {"heading": "机制", "bullets": ["首触发竞争"]},
                    {"heading": "边界", "bullets": ["只验证视觉规则"]},
                ],
                "claim": "渐变填充必须同时写入 PNG 与原生 PPTX。",
                "evidence": "颜色导出回归测试。",
                "boundary": "仅验证图形填充。",
                "speech": "图形渐变回归。",
            }
        ],
    }


def _export_pku(work_tmp: Path) -> Path:
    deck_path = work_tmp / "deck.json"
    deck_path.write_text(json.dumps(_deck(), ensure_ascii=False), encoding="utf-8")
    out = work_tmp / "out"
    exporter = EditableDeckExporter(deck_path, PKU_THEME, out, profile="defense")
    exporter.export("gradient.pptx")
    return out / "gradient.pptx"


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


def test_pku_module_headers_and_pale_cards_export_as_gradients(work_tmp: Path) -> None:
    slide = Presentation(str(_export_pku(work_tmp))).slides[0]
    headers = [shape for shape in slide.shapes if shape.name == "module_header_fill"]
    pale_cards = [shape for shape in slide.shapes if shape.name == "soft_gradient_panel"]
    sidebar_active = next(shape for shape in slide.shapes if shape.name == "sidebar_active")
    sidebar_text = next(shape for shape in slide.shapes if shape.name == "sidebar_background")

    assert len(headers) == 2
    assert pale_cards
    for shape in headers + pale_cards + [sidebar_active]:
        assert shape.fill.type == MSO_FILL.GRADIENT
        xml = shape.fill._xPr.xml
        assert '<a:lin scaled="0" ang="0"/>' in xml

    header_xml = headers[0].fill._xPr.xml
    assert f'"{PKU_COLORS["blue_pale"].lstrip("#")}"' in header_xml
    assert '"FFFFFF"' in header_xml
    assert '"94070A"' not in header_xml

    sidebar_xml = sidebar_active.fill._xPr.xml
    assert f'"{PKU_COLORS["blue_pale"].lstrip("#")}"' in sidebar_xml
    assert f'"{PKU_COLORS["blue_wash"].lstrip("#")}"' in sidebar_xml
    active_text_color = sidebar_text.text_frame.paragraphs[0].runs[0].font.color.rgb
    assert str(active_text_color) == PKU_COLORS["blue_dark"].lstrip("#")

    pale_xml = pale_cards[0].fill._xPr.xml
    assert any(
        f'"{PKU_COLORS[name].lstrip("#")}"' in pale_xml
        for name in ("blue_pale", "blue_wash", "white")
    )
    gradient_colors = [
        value
        for value in ("blue_pale", "blue_wash", "white")
        if (
            f'"{PKU_COLORS[value].lstrip("#")}"' in header_xml
            or f'"{PKU_COLORS[value].lstrip("#")}"' in pale_xml
        )
    ]
    assert gradient_colors
    assert min(_mean_rgb(PKU_COLORS[name]) for name in gradient_colors) >= 235
