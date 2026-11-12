from __future__ import annotations

import argparse
import json
from pathlib import Path

import numpy as np
from PIL import Image

from common import estimate_seconds, load_json, load_theme, resolve_path, write_json


SUPPORTED_LAYOUTS = {
    "cover",
    "outline",
    "section",
    "bullets",
    "figure",
    "figure_flow",
    "concept",
    "formula",
    "pipeline",
    "compare",
    "metrics",
    "closing",
    "grid",
    "balance",
    "hub",
    "timeline",
    "references",
}

REQUIRES_IMAGE_OR_CHART = {"figure", "figure_flow", "metrics", "balance"}
DIAGRAM_KINDS = {"sequence", "branch", "stack", "ruler", "matrix", "cycle"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate a defense deck and its rendered images.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--theme", required=True)
    parser.add_argument("--render-dir", required=True)
    parser.add_argument("--out", required=True)
    return parser.parse_args()


def add_issue(issues: list[dict], severity: str, page: str, message: str) -> None:
    issues.append({"severity": severity, "page": page, "message": message})


def content_units(slide: dict) -> int:
    layout = slide.get("layout")
    if layout == "bullets":
        columns = slide.get("columns") or []
        return len(columns) if columns else len(slide.get("bullets") or [])
    if layout == "grid":
        return len(slide.get("blocks") or [])
    if layout == "balance":
        left_items = (slide.get("left") or {}).get("items") or []
        right_items = (slide.get("right") or {}).get("items") or []
        return len(left_items) + len(right_items)
    if layout == "figure":
        return len(slide.get("panels") or [])
    if layout == "figure_flow":
        return len(slide.get("flow") or []) + len(slide.get("panels") or [])
    if layout == "concept":
        return len((slide.get("diagram") or {}).get("items") or []) + len(slide.get("panels") or [])
    if layout == "formula":
        return len(slide.get("variables") or []) + len(slide.get("panels") or [])
    if layout == "pipeline":
        return len(slide.get("steps") or [])
    if layout == "compare":
        return len(slide.get("rows") or [])
    if layout == "metrics":
        return len(slide.get("metrics") or []) + len(slide.get("conditions") or [])
    if layout == "closing":
        return len(slide.get("points") or [])
    if layout == "hub":
        return len(slide.get("items") or [])
    if layout == "timeline":
        return len(slide.get("phases") or [])
    if layout == "references":
        return len(slide.get("references") or [])
    return 0


def validate_diagram(diagram: dict, page_id: str, issues: list[dict], label: str = "Diagram") -> None:
    """Validate semantic diagram data before rendering.

    Diagrams are content-bearing, not decoration. Keep their cardinality and
    required fields explicit so a model cannot silently produce an empty or
    misleading schematic.
    """
    if not isinstance(diagram, dict):
        add_issue(issues, "error", page_id, f"{label} must be an object.")
        return
    kind = str(diagram.get("kind", "sequence")).strip().lower()
    if kind not in DIAGRAM_KINDS:
        add_issue(
            issues,
            "error",
            page_id,
            f"{label} has unsupported kind: {kind}. Use one of {sorted(DIAGRAM_KINDS)}.",
        )
        return
    items = [item for item in (diagram.get("items") or []) if isinstance(item, dict)]
    limits = {
        "sequence": (2, 5),
        "branch": (2, 4),
        "stack": (2, 5),
        "ruler": (2, 8),
        "matrix": (4, 4),
        "cycle": (3, 5),
    }
    minimum, maximum = limits[kind]
    if not minimum <= len(items) <= maximum:
        add_issue(
            issues,
            "error",
            page_id,
            f"{label} needs {minimum}–{maximum} items for kind '{kind}'.",
        )
    if kind == "branch" and not isinstance(diagram.get("root"), dict):
        add_issue(issues, "warning", page_id, f"{label} should define root for branch diagrams.")
    if kind == "ruler":
        for item_index, item in enumerate(items, start=1):
            if "value" not in item:
                add_issue(issues, "error", page_id, f"{label} ruler item {item_index} needs numeric value.")
                continue
            try:
                float(item["value"])
            except (TypeError, ValueError):
                add_issue(issues, "error", page_id, f"{label} ruler item {item_index} value is not numeric.")
    for item_index, item in enumerate(items, start=1):
        if not str(item.get("title") or item.get("label") or "").strip():
            add_issue(issues, "warning", page_id, f"{label} item {item_index} has no title/label.")
    if "highlight" in diagram:
        highlight = diagram.get("highlight")
        if not isinstance(highlight, int) or not 0 <= highlight < len(items):
            add_issue(issues, "error", page_id, f"{label} highlight is outside the items range.")
    if kind == "matrix":
        if not isinstance(diagram.get("x_axis"), dict) or not isinstance(diagram.get("y_axis"), dict):
            add_issue(issues, "warning", page_id, f"{label} matrix should define x_axis and y_axis labels.")
        pairs = diagram.get("pairs") or []
        if pairs and len(pairs) != len(items):
            add_issue(issues, "warning", page_id, f"{label} pairs count should match item count.")


def main() -> None:
    args = parse_args()
    deck_path = Path(args.deck).resolve()
    theme_path = Path(args.theme).resolve()
    render_dir = Path(args.render_dir).resolve()
    output_dir = Path(args.out).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    deck = load_json(deck_path)
    theme = load_theme(theme_path)
    issues: list[dict] = []
    deck_mode = str(deck.get("meta", {}).get("deck_mode", "full")).strip().lower()
    sections = deck.get("sections", [])
    slides = deck.get("slides", [])
    section_ids = [section.get("id") for section in sections]

    if deck_mode not in {"full", "content_only"}:
        add_issue(issues, "error", "meta", "meta.deck_mode must be 'full' or 'content_only'.")
    if not deck.get("meta", {}).get("title"):
        add_issue(issues, "error", "meta", "Missing meta.title.")
    if not slides:
        add_issue(issues, "error", "deck", "Deck has no slides.")
    if len(sections) < 4:
        add_issue(issues, "error", "sections", "Use at least four major sections.")
    if len(sections) > 6:
        add_issue(issues, "warning", "sections", "More than six sections will crowd the mini-directory.")
    if len(set(section_ids)) != len(section_ids):
        add_issue(issues, "error", "sections", "Section ids must be unique.")

    if slides and deck_mode == "full":
        if slides[0].get("layout") != "cover":
            add_issue(issues, "error", "01", "The first slide must use the cover layout.")
        if len(slides) > 1 and slides[1].get("layout") != "outline":
            add_issue(issues, "error", "02", "The second slide must use the outline layout.")
    if slides and deck_mode == "content_only":
        for index, slide in enumerate(slides, start=1):
            if slide.get("layout") in {"cover", "outline"}:
                add_issue(issues, "error", str(slide.get("id", f"{index:02d}")), "content_only mode must not include cover or outline slides.")

    for index, slide in enumerate(slides, start=1):
        page_id = str(slide.get("id", f"{index:02d}"))
        layout = slide.get("layout")
        if layout not in SUPPORTED_LAYOUTS:
            add_issue(issues, "error", page_id, f"Unsupported layout: {layout}")
            continue
        if not slide.get("title"):
            add_issue(issues, "error", page_id, "Missing slide title.")
        if not str(slide.get("speech", "")).strip():
            add_issue(issues, "error", page_id, "Missing speech text.")
        if layout not in {"cover", "outline"} and slide.get("section") not in section_ids:
            add_issue(issues, "error", page_id, f"Unknown or missing section: {slide.get('section')}")
        chart = slide.get("chart")
        if chart is not None:
            if not isinstance(chart, dict):
                add_issue(issues, "error", page_id, "chart must be an object.")
            else:
                chart_type = str(chart.get("type", "histogram")).strip().lower()
                if chart_type not in {"histogram", "bar", "column", "line", "curve"}:
                    add_issue(issues, "error", page_id, f"Unsupported chart type: {chart_type}")
                values = chart.get("values") or []
                if len(values) < 3:
                    add_issue(issues, "error", page_id, "Chart needs at least three values.")
                if len(values) > 24:
                    add_issue(issues, "warning", page_id, "Chart has more than 24 values and may be unreadable.")
                for value_index, value in enumerate(values, start=1):
                    if not isinstance(value, (int, float)):
                        add_issue(issues, "error", page_id, f"Chart value {value_index} is not numeric.")
                if "highlight" in chart and values:
                    highlight = chart.get("highlight")
                    if not isinstance(highlight, int) or not 0 <= highlight < len(values):
                        add_issue(issues, "error", page_id, "Chart highlight index is outside the values range.")
        if layout in {"cover", "outline"} and slide.get("section"):
            add_issue(issues, "warning", page_id, "Cover and outline should not highlight a mini-directory section.")
        if layout == "bullets":
            columns = slide.get("columns") or []
            if not columns and slide.get("bullets"):
                columns = [{"bullets": slide["bullets"]}]
            if len(columns) not in {1, 2, 3}:
                add_issue(issues, "error", page_id, "Bullets layout supports one to three columns.")
            for column_index, column in enumerate(columns, start=1):
                if len(column.get("bullets", [])) > 4:
                    add_issue(issues, "warning", page_id, f"Column {column_index} has more than four bullets.")
        if layout == "figure":
            image_path = resolve_path(slide.get("image"), deck_path.parent)
            if not slide.get("chart") and (image_path is None or not image_path.exists()):
                add_issue(issues, "error", page_id, f"Missing figure: {slide.get('image')}")
            if not slide.get("image_role"):
                add_issue(issues, "warning", page_id, "Missing image_role.")
            if not slide.get("image_boundary"):
                add_issue(issues, "warning", page_id, "Missing image_boundary.")
            panels = slide.get("panels") or []
            if not 1 <= len(panels) <= 3:
                add_issue(issues, "error", page_id, "Figure layout needs one to three side panels.")
        if layout == "figure_flow":
            flow = slide.get("flow") or []
            panels = slide.get("panels") or []
            if not 3 <= len(flow) <= 5:
                add_issue(issues, "error", page_id, "Figure-flow layout needs three to five flow steps.")
            if not 1 <= len(panels) <= 3:
                add_issue(issues, "error", page_id, "Figure-flow layout needs one to three side panels.")
            image_path = resolve_path(slide.get("image"), deck_path.parent)
            if not slide.get("chart") and (image_path is None or not image_path.exists()):
                add_issue(issues, "error", page_id, "Figure-flow layout needs an image or chart.")
        if layout == "concept":
            diagram = slide.get("diagram")
            if not isinstance(diagram, dict):
                add_issue(issues, "error", page_id, "Concept layout needs a diagram object.")
            else:
                validate_diagram(diagram, page_id, issues, label="Concept diagram")
            panels = slide.get("panels") or []
            if not 1 <= len(panels) <= 3:
                add_issue(issues, "error", page_id, "Concept layout needs one to three interpretation panels.")
            if slide.get("image") or slide.get("chart"):
                add_issue(issues, "warning", page_id, "Concept layout is diagram-led; use figure/figure_flow for real evidence.")
        if layout == "formula":
            if not (str(slide.get("formula", "")).strip() or str(slide.get("formula_latex", "")).strip()):
                add_issue(issues, "error", page_id, "Formula layout needs formula or formula_latex.")
            variables = slide.get("variables") or []
            if not 2 <= len(variables) <= 8:
                add_issue(issues, "error", page_id, "Formula layout needs two to eight variable explanations.")
            panels = slide.get("panels") or []
            if not 1 <= len(panels) <= 3:
                add_issue(issues, "error", page_id, "Formula layout needs one to three interpretation panels.")
        if layout == "pipeline":
            step_count = len(slide.get("steps") or [])
            if not 3 <= step_count <= 5:
                add_issue(issues, "error", page_id, "Pipeline needs three to five steps.")
        if layout == "compare":
            columns = slide.get("columns") or []
            if not 2 <= len(columns) <= 5:
                add_issue(issues, "error", page_id, "Compare layout needs two to five columns.")
            for row_index, row in enumerate(slide.get("rows") or [], start=1):
                if len(row) != len(columns):
                    add_issue(issues, "error", page_id, f"Compare row {row_index} has {len(row)} cells, expected {len(columns)}.")
            if len(slide.get("rows") or []) > 5:
                add_issue(issues, "warning", page_id, "Compare table has more than five rows.")
        if layout == "metrics":
            metric_count = len(slide.get("metrics") or [])
            if not 2 <= metric_count <= 4:
                add_issue(issues, "error", page_id, "Metrics layout needs two to four metrics.")
            conditions = slide.get("conditions") or []
            if conditions and not isinstance(conditions, list):
                add_issue(issues, "error", page_id, "Metrics conditions must be a list.")
            for condition_index, condition in enumerate(conditions, start=1):
                if isinstance(condition, dict):
                    if not condition.get("label") or not condition.get("value"):
                        add_issue(issues, "warning", page_id, f"Metric condition {condition_index} needs label and value.")
                elif not str(condition).strip():
                    add_issue(issues, "warning", page_id, f"Metric condition {condition_index} is empty.")
        if layout == "grid":
            blocks = slide.get("blocks") or []
            if not 3 <= len(blocks) <= 6:
                add_issue(issues, "error", page_id, "Grid layout needs three to six blocks.")
            columns = int(slide.get("columns", 2))
            if columns not in {2, 3}:
                add_issue(issues, "error", page_id, "Grid layout supports two or three columns.")
            for block_index, block in enumerate(blocks, start=1):
                if not block.get("heading"):
                    add_issue(issues, "warning", page_id, f"Grid block {block_index} has no heading.")
                span = int(block.get("span", 1))
                if not 1 <= span <= columns:
                    add_issue(issues, "error", page_id, f"Grid block {block_index} span {span} exceeds {columns} columns.")
                if block.get("image"):
                    image_path = resolve_path(block.get("image"), deck_path.parent)
                    if image_path is None or not image_path.exists():
                        add_issue(issues, "error", page_id, f"Missing grid image: {block.get('image')}")
                if block.get("formula_latex") and not block.get("formula"):
                    add_issue(
                        issues,
                        "warning",
                        page_id,
                        f"Grid block {block_index} has formula_latex but no plain formula fallback.",
                    )
                if block.get("diagram"):
                    validate_diagram(block.get("diagram"), page_id, issues, label=f"Grid block {block_index} diagram")
                    if block.get("image"):
                        add_issue(issues, "warning", page_id, f"Grid block {block_index} has both image and diagram; image wins in rendering.")
        if layout == "balance":
            for rail_name in ("left", "right"):
                rail = slide.get(rail_name) or {}
                items = rail.get("items") or []
                if not 2 <= len(items) <= 4:
                    add_issue(issues, "error", page_id, f"Balance {rail_name} rail needs two to four items.")
                if not rail.get("heading"):
                    add_issue(issues, "warning", page_id, f"Balance {rail_name} rail has no heading.")
            image_path = resolve_path(slide.get("image"), deck_path.parent)
            if not slide.get("chart") and (image_path is None or not image_path.exists()):
                add_issue(issues, "error", page_id, "Balance layout needs a center image or chart.")
            if not slide.get("image_role"):
                add_issue(issues, "warning", page_id, "Balance layout is missing image_role.")
            if not slide.get("image_boundary"):
                add_issue(issues, "warning", page_id, "Balance layout is missing image_boundary.")
        if layout == "hub":
            items = slide.get("items") or []
            if not 4 <= len(items) <= 6:
                add_issue(issues, "error", page_id, "Hub layout needs four to six surrounding items.")
            center = slide.get("center")
            if not isinstance(center, dict) or not center.get("value"):
                add_issue(issues, "warning", page_id, "Hub layout should define center.value.")
        if layout == "timeline":
            phases = slide.get("phases") or []
            if not 3 <= len(phases) <= 5:
                add_issue(issues, "error", page_id, "Timeline layout needs three to five phases.")
            for phase_index, phase in enumerate(phases, start=1):
                if not str(phase.get("period", "")).strip():
                    add_issue(issues, "error", page_id, f"Timeline phase {phase_index} has no period.")
                if not str(phase.get("title", "")).strip():
                    add_issue(issues, "error", page_id, f"Timeline phase {phase_index} has no title.")
                if not str(phase.get("body", "")).strip() and not phase.get("bullets"):
                    add_issue(issues, "warning", page_id, f"Timeline phase {phase_index} has no body or bullets.")
                if len(phase.get("bullets") or []) > 4:
                    add_issue(issues, "warning", page_id, f"Timeline phase {phase_index} has more than four bullets.")
        if layout == "references":
            references = slide.get("references") or []
            if not 4 <= len(references) <= 10:
                add_issue(issues, "error", page_id, "References layout needs four to ten entries.")
            seen_titles = set()
            for ref_index, reference in enumerate(references, start=1):
                title = str(reference.get("title", "")).strip()
                venue = str(reference.get("venue", "")).strip()
                if not title:
                    add_issue(issues, "error", page_id, f"Reference {ref_index} has no title.")
                if not venue:
                    add_issue(issues, "error", page_id, f"Reference {ref_index} has no venue.")
                if title and title in seen_titles:
                    add_issue(issues, "warning", page_id, f"Reference {ref_index} duplicates an earlier title.")
                seen_titles.add(title)
        if layout == "closing":
            if not slide.get("points"):
                add_issue(issues, "warning", page_id, "Closing slide has no summary points.")
            if len(slide.get("points") or []) > 4:
                add_issue(issues, "warning", page_id, "Closing slide should keep at most four summary points.")
        if layout not in {"cover", "outline", "section", "closing"}:
            if not str(slide.get("lead", "")).strip():
                add_issue(issues, "warning", page_id, "Content page has no lead claim.")
            units = content_units(slide)
            if units < 2:
                add_issue(issues, "warning", page_id, "Content density is low; add a second evidence or explanation unit.")
            if layout in REQUIRES_IMAGE_OR_CHART:
                image_path = resolve_path(slide.get("image"), deck_path.parent)
                if not slide.get("chart") and (image_path is None or not image_path.exists()):
                    add_issue(issues, "error", page_id, "This layout needs a real image or chart as its visual anchor.")

    expected_size = (int(theme["canvas"]["width"]), int(theme["canvas"]["height"]))
    rendered_paths = sorted(render_dir.glob("slide_*.png"))
    if len(rendered_paths) != len(slides):
        add_issue(issues, "error", "render", f"Found {len(rendered_paths)} rendered slides for {len(slides)} deck slides.")
    for index, path in enumerate(rendered_paths, start=1):
        image = Image.open(path).convert("RGB")
        if image.size != expected_size:
            add_issue(issues, "error", f"{index:02d}", f"Rendered size {image.size} != expected {expected_size}.")
    if len(rendered_paths) > 2:
        brand_box = tuple(theme["brand"]["header_box"])
        footer_y = int(theme["layout"]["footer_y"])
        reference = np.array(Image.open(rendered_paths[1]).convert("RGB"))
        reference_brand = reference[brand_box[1]:brand_box[3], brand_box[0]:brand_box[2]]
        reference_footer = reference[footer_y:footer_y + 3, 24:expected_size[0] - 24]
        brand_mask = None
        lockup_path = theme.get("brand", {}).get("lockup_path")
        if lockup_path and Path(lockup_path).exists():
            lockup = Image.open(lockup_path).convert("RGBA")
            box_width = brand_box[2] - brand_box[0]
            box_height = brand_box[3] - brand_box[1]
            scale = min(box_width / lockup.width, box_height / lockup.height)
            size = (max(1, round(lockup.width * scale)), max(1, round(lockup.height * scale)))
            alpha = lockup.resize(size, Image.Resampling.LANCZOS).getchannel("A")
            full_mask = np.zeros((box_height, box_width), dtype=bool)
            offset_x = (box_width - size[0]) // 2
            offset_y = (box_height - size[1]) // 2
            full_mask[offset_y:offset_y + size[1], offset_x:offset_x + size[0]] = np.array(alpha) > 180
            brand_mask = full_mask
        for index, path in enumerate(rendered_paths[1:], start=2):
            current = np.array(Image.open(path).convert("RGB"))
            current_brand = current[brand_box[1]:brand_box[3], brand_box[0]:brand_box[2]]
            current_footer = current[footer_y:footer_y + 3, 24:expected_size[0] - 24]
            if brand_mask is not None:
                difference = np.abs(reference_brand.astype(int) - current_brand.astype(int))[brand_mask]
                brand_mismatch = int(difference.max()) if difference.size else 0
            else:
                brand_mismatch = int(np.abs(reference_brand.astype(int) - current_brand.astype(int)).max())
            brand_tolerance = 4 if brand_mask is not None else 18
            if brand_mismatch > brand_tolerance:
                add_issue(issues, "error", f"{index:02d}", "Upper-right identity is not pixel-fixed.")
            if not np.array_equal(reference_footer, current_footer):
                add_issue(issues, "error", f"{index:02d}", "Footer rule is not pixel-fixed.")

    estimated = sum(estimate_seconds(str(slide.get("speech", ""))) for slide in slides)
    target_seconds = int(float(deck.get("meta", {}).get("duration_min", 20)) * 60)
    if target_seconds and estimated > target_seconds * 1.15:
        add_issue(issues, "warning", "timing", f"Estimated speech {estimated}s exceeds target {target_seconds}s by more than 15%.")
    if target_seconds and estimated < target_seconds * 0.65:
        add_issue(issues, "warning", "timing", f"Estimated speech {estimated}s is much shorter than target {target_seconds}s.")

    error_count = sum(1 for issue in issues if issue["severity"] == "error")
    warning_count = sum(1 for issue in issues if issue["severity"] == "warning")
    report = {
        "deck": str(deck_path),
        "theme": theme["name"],
        "deck_mode": deck_mode,
        "slide_count": len(slides),
        "rendered_count": len(rendered_paths),
        "estimated_speech_seconds": estimated,
        "target_speech_seconds": target_seconds,
        "error_count": error_count,
        "warning_count": warning_count,
        "issues": issues,
    }
    write_json(output_dir / "validation_report.json", report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if error_count:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
