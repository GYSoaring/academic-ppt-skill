from __future__ import annotations

import json
import re
import subprocess
import sys
from copy import deepcopy
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.enum.dml import MSO_FILL

from common import load_theme
from design import (
    FontBook,
    draw_diagram,
    relation_arrow_polygon,
    relation_connector_geometry,
    relation_node_boxes,
)
from export_editable_pptx import EditableDeckExporter
from validate_deck import content_units, validate_diagram


SKILL_ROOT = Path(__file__).resolve().parents[1]
THEME = SKILL_ROOT / "assets" / "themes" / "academic_blue" / "theme.json"


def _diagram() -> dict:
    return {
        "kind": "relations",
        "arrangement": "free",
        "nodes": [
            {
                "id": "a",
                "title": "共同问题",
                "body": "约束与目标",
                "tone": "solid",
                "x": 0.02,
                "y": 0.40,
                "w": 0.18,
                "h": 0.20,
            },
            {
                "id": "b",
                "title": "证据接口",
                "body": "指标与边界",
                "x": 0.38,
                "y": 0.06,
                "w": 0.20,
                "h": 0.24,
            },
            {
                "id": "c",
                "title": "综合判断",
                "body": "收益与代价",
                "tone": "outline",
                "x": 0.76,
                "y": 0.40,
                "w": 0.20,
                "h": 0.20,
            },
        ],
        "relations": [
            {"from": "a", "to": "c", "label": "推进", "style": "ribbon"},
            {"from": "b", "to": "c", "label": "汇合", "style": "line", "route": "orthogonal"},
            {"from": "c", "to": "a", "label": "约束", "style": "double"},
        ],
        "note": "箭头只表示论证关系。",
    }


def _deck() -> dict:
    return {
        "meta": {
            "title": "关系箭头回归",
            "type_profile": "lecture",
            "title_mode": "topic",
            "deck_mode": "content_only",
        },
        "sections": [{"id": "method", "title": "方法"}],
        "slides": [
            {
                "id": "01",
                "layout": "concept",
                "section": "method",
                "title": "关系地图",
                "lead": "用有方向的箭头表达真实论证关系。",
                "diagram_label": "研究路径",
                "diagram_frame": False,
                "diagram": _diagram(),
                "speech": "关系地图用方向、路径和标签表达真实的论证关系。",
            }
        ],
    }


def _export_deck(deck: dict, work_tmp: Path, name: str) -> Path:
    deck_path = work_tmp / "deck.json"
    deck_path.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
    output_dir = work_tmp / "out"
    exporter = EditableDeckExporter(deck_path, THEME, output_dir, profile="lecture")
    filename = f"{name}.pptx"
    exporter.export(filename)
    return output_dir / filename


def _export(work_tmp: Path) -> Path:
    return _export_deck(_deck(), work_tmp, "relations")


def test_chevron_polygon_has_a_concave_tail_and_forward_point() -> None:
    points = relation_arrow_polygon((100, 200), (420, 200), 48, "chevron")
    assert len(points) == 6
    assert points[0] == (100, 200)
    assert points[3] == (420, 200)
    assert points[1][0] > points[0][0]
    assert points[5][0] > points[0][0]


def test_png_relation_arrow_fades_from_tail_to_head() -> None:
    theme = load_theme(THEME)
    colors = theme["_colors"]
    fonts = FontBook(theme["fonts"]["title"], theme["fonts"]["body"])
    image = Image.new("RGB", (1000, 500), colors["white"])
    draw = ImageDraw.Draw(image)
    diagram = _diagram()
    diagram["relations"] = [{"from": "a", "to": "c", "label": "", "style": "ribbon"}]
    diagram["note"] = ""
    draw_diagram(
        draw,
        (0, 0, 1000, 500),
        diagram,
        colors,
        fonts,
        frame=False,
        image=image,
    )

    boxes = relation_node_boxes(diagram, 0, 0, 1000, 500)
    start, end = relation_connector_geometry(boxes["a"], boxes["c"])
    tail = image.getpixel(
        (
            round(start[0] + (end[0] - start[0]) * 0.28),
            round(start[1] + (end[1] - start[1]) * 0.28),
        )
    )
    head = image.getpixel(
        (
            round(start[0] + (end[0] - start[0]) * 0.78),
            round(start[1] + (end[1] - start[1]) * 0.78),
        )
    )
    assert sum(tail) > sum(head)
    assert sum(head) < 740


def test_native_relation_arrows_are_editable_gradients_and_keep_the_final_arrowhead(
    work_tmp: Path,
) -> None:
    presentation = Presentation(str(_export(work_tmp)))
    slide = presentation.slides[0]
    shapes = list(slide.shapes)
    names = [shape.name for shape in shapes]

    assert sum(bool(re.fullmatch(r"relation_node_\d+", name)) for name in names) == 3
    assert sum(bool(re.fullmatch(r"relation_label_\d+", name)) for name in names) == 3
    assert "relation_arrow_1" in names
    assert "relation_arrow_3" in names

    ribbon = next(shape for shape in shapes if shape.name == "relation_arrow_1")
    assert ribbon.fill.type == MSO_FILL.GRADIENT
    assert abs(float(ribbon.fill.gradient_angle)) < 1.0
    node_shapes = [
        shape
        for shape in shapes
        if re.fullmatch(r"relation_node_\d+", shape.name)
    ]
    assert all(shape.width < presentation.slide_width * 0.40 for shape in node_shapes)
    assert all(shape.height < presentation.slide_height * 0.45 for shape in node_shapes)
    body_sizes = [
        run.font.size.pt / (72.0 / 153.6)
        for shape in shapes
        if shape.name == "relation_node_body"
        for paragraph in shape.text_frame.paragraphs
        for run in paragraph.runs
        if run.font.size is not None
    ]
    assert body_sizes
    assert max(body_sizes) < 40.0

    orthogonal_segments = [
        shape
        for shape in shapes
        if shape.name == "relation_arrow_2" or shape.name.startswith("relation_arrow_2_seg_")
    ]
    assert len(orthogonal_segments) >= 2
    final_segment = orthogonal_segments[-1]
    assert 'type="triangle"' in final_segment._element.xml
    assert final_segment.line.width.pt >= 2.5


def test_native_diagonal_gradient_follows_the_arrow_direction(work_tmp: Path) -> None:
    deck = _deck()
    diagram = deck["slides"][0]["diagram"]
    diagram["note"] = ""
    diagram["nodes"] = [
        {
            "id": "a",
            "title": "起点",
            "x": 0.03,
            "y": 0.08,
            "w": 0.20,
            "h": 0.20,
        },
        {
            "id": "b",
            "title": "终点",
            "x": 0.64,
            "y": 0.62,
            "w": 0.20,
            "h": 0.20,
        },
    ]
    diagram["relations"] = [
        {"from": "a", "to": "b", "label": "", "style": "ribbon"},
    ]

    slide = Presentation(str(_export_deck(deck, work_tmp, "diagonal"))).slides[0]
    ribbon = next(shape for shape in slide.shapes if shape.name == "relation_arrow_1")
    angle_match = re.search(r'<a:lin[^>]*ang="(\d+)"', ribbon.fill._xPr.xml)
    assert angle_match is not None
    clockwise_degrees = int(angle_match.group(1)) / 60000.0
    assert 5.0 < clockwise_degrees < 80.0


def test_relation_validation_accepts_the_contract_and_counts_nodes() -> None:
    issues: list[dict] = []
    validate_diagram(_diagram(), "01", issues, label="Concept diagram")
    assert not [issue for issue in issues if issue["severity"] == "error"]
    assert content_units({"layout": "concept", "diagram": _diagram(), "panels": []}) == 3


def test_relation_validation_rejects_bad_endpoints_duplicates_and_styles() -> None:
    diagram = deepcopy(_diagram())
    diagram["nodes"][1]["id"] = "a"
    diagram["relations"][0]["to"] = "missing"
    diagram["relations"][1]["style"] = "sales"
    diagram["relations"][2]["from"] = "c"
    diagram["relations"][2]["to"] = "c"
    issues: list[dict] = []
    validate_diagram(diagram, "01", issues, label="Concept diagram")
    messages = " ".join(issue["message"] for issue in issues if issue["severity"] == "error")
    assert "duplicated" in messages
    assert "existing node ids" in messages
    assert "style must be" in messages
    assert "self-loop" in messages


def test_relation_validation_warns_on_overlapping_nodes() -> None:
    diagram = deepcopy(_diagram())
    diagram["nodes"][2]["x"] = 0.05
    diagram["nodes"][2]["y"] = 0.42
    issues: list[dict] = []
    validate_diagram(diagram, "01", issues, label="Concept diagram")
    assert any("overlap" in issue["message"] for issue in issues if issue["severity"] == "warning")


def test_concept_relation_map_can_omit_interpretation_panels(work_tmp: Path) -> None:
    deck = _deck()
    deck["meta"]["title_mode"] = "topic"
    deck["meta"]["type_profile"] = "document"
    deck["sections"] = [
        {"id": "background", "title": "背景"},
        {"id": "method", "title": "方法"},
        {"id": "result", "title": "结果"},
        {"id": "closing", "title": "结论"},
    ]
    deck_path = work_tmp / "deck.json"
    deck_path.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
    rendered_dir = work_tmp / "out" / "rendered"
    rendered_dir.mkdir(parents=True)
    Image.new("RGB", (2048, 1152), "white").save(rendered_dir / "slide_01.png")
    output_dir = work_tmp / "out"
    result = subprocess.run(
        [
            sys.executable,
            str(SKILL_ROOT / "scripts" / "validate_deck.py"),
            "--deck",
            str(deck_path),
            "--theme",
            str(THEME),
            "--render-dir",
            str(rendered_dir),
            "--out",
            str(output_dir),
            "--profile",
            "document",
        ],
        cwd=SKILL_ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert result.returncode == 0, result.stdout + result.stderr
    report = json.loads((output_dir / "validation_report.json").read_text(encoding="utf-8"))
    assert not [
        issue
        for issue in report["issues"]
        if issue["severity"] == "error" and "panel" in issue["message"].lower()
    ]
