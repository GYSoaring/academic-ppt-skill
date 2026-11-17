from __future__ import annotations

import argparse
import json
import re
import statistics
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from accessibility import get_alt_text, is_decorative
from common import load_json, slide_display_title, write_json


PX_PER_INCH = 153.6
EMU_PER_PX = 914400.0 / PX_PER_INCH
ATOMIC_TEXT_NAMES = {"chapter", "module_index", "diagram_index", "footnote"}
ATOMIC_TEXT_PREFIXES = ("outline_number_", "diagram_index")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate native editable PowerPoint structure.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--deck", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--lecture-audit", action="store_true", help="Check body >=18pt, heading >=24pt, footnote >=10pt.")
    parser.add_argument("--min-body-pt", type=float, default=None)
    parser.add_argument("--min-heading-pt", type=float, default=None)
    parser.add_argument("--min-footnote-pt", type=float, default=None)
    return parser.parse_args()


def add_issue(issues: list[dict], severity: str, page: str, message: str) -> None:
    issues.append({"severity": severity, "page": page, "message": message})


def count_formula_shapes(slide) -> int:
    count = 0
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        xml = shape.text_frame._txBody.xml
        if "http://schemas.microsoft.com/office/drawing/2010/main" in xml:
            count += 1
    return count


def classify_shape(shape) -> str:
    name = str(getattr(shape, "name", "")).lower()
    if any(token in name for token in ("footnote", "footer", "source", "caption", "chart_x", "chart_y", "chart_badge")):
        return "footnote"
    if any(token in name for token in ("title", "header", "eyebrow", "kicker")):
        return "heading"
    return "body"


def summarize_sizes(values: list[float]) -> dict:
    if not values:
        return {"count": 0, "min": None, "median": None, "max": None}
    return {
        "count": len(values),
        "min": round(min(values), 1),
        "median": round(statistics.median(values), 1),
        "max": round(max(values), 1),
    }


def minimum_editable_text_shapes(layout: str) -> int:
    """Section dividers intentionally carry only their native title."""
    return {"cover": 3, "section": 1}.get(str(layout).strip().lower(), 5)


def is_atomic_text_shape(name: str) -> bool:
    normalized = str(name or "").strip().lower()
    return normalized in ATOMIC_TEXT_NAMES or normalized.startswith(ATOMIC_TEXT_PREFIXES)


def shape_font_sizes(shape) -> list[float]:
    return [
        round(run.font.size.pt, 1)
        for paragraph in shape.text_frame.paragraphs
        for run in paragraph.runs
        if run.font.size
    ]


def geometry_issues(
    slide,
    slide_width: int,
    slide_height: int,
    type_profile: str = "",
) -> list[str]:
    """Return native-layout defects that PNG-only review cannot catch."""
    issues: list[str] = []
    text_shapes = [
        shape
        for shape in slide.shapes
        if shape.has_text_frame and shape.text.strip()
    ]
    nonpositive = [
        shape.name
        for shape in text_shapes
        if shape.width <= 0 or shape.height <= 0
    ]
    if nonpositive:
        issues.append("Text shapes have non-positive geometry: " + ", ".join(nonpositive))
    outside = [
        shape.name
        for shape in text_shapes
        if shape.left < 0
        or shape.top < 0
        or shape.left + shape.width > slide_width
        or shape.top + shape.height > slide_height
    ]
    if outside:
        issues.append("Text shapes extend outside the slide: " + ", ".join(outside))
    wrapped_numbers = [
        shape.name
        for shape in text_shapes
        if re.fullmatch(r"\d\s*\n\s*\d", shape.text.strip())
    ]
    if wrapped_numbers:
        issues.append("Two-digit index text wrapped across lines: " + ", ".join(wrapped_numbers))
    wrapped_atomic = [
        shape.name
        for shape in text_shapes
        if is_atomic_text_shape(shape.name) and "\n" in shape.text.strip()
    ]
    if wrapped_atomic:
        issues.append("Atomic index or source text contains a line break: " + ", ".join(wrapped_atomic))
    wrapped_module_indices = [
        shape.name
        for shape in text_shapes
        if shape.name == "module_index" and shape.text_frame.word_wrap is not False
    ]
    if wrapped_module_indices:
        issues.append("Module index must disable word wrap: " + ", ".join(wrapped_module_indices))
    narrow_module_indices = [
        shape.name
        for shape in text_shapes
        if shape.name == "module_index" and shape.width < 60 * EMU_PER_PX
    ]
    if narrow_module_indices:
        issues.append("Module index is narrower than 60 px: " + ", ".join(narrow_module_indices))
    heading_sizes = [
        size
        for shape in text_shapes
        if shape.name == "module_heading"
        for size in shape_font_sizes(shape)
    ]
    concept_body_sizes = [
        size
        for shape in text_shapes
        if shape.name == "concept_panel_body"
        for size in shape_font_sizes(shape)
    ]
    if heading_sizes and concept_body_sizes and min(heading_sizes) <= max(concept_body_sizes):
        issues.append(
            "Concept panel heading must outrank its body text: "
            f"heading min {min(heading_sizes)} pt, body max {max(concept_body_sizes)} pt"
        )

    text_by_name = {str(shape.name): shape for shape in text_shapes}
    condition_labels = [
        shape
        for shape in text_shapes
        if shape.name.startswith("condition_label_")
    ]
    for label in condition_labels:
        suffix = label.name.removeprefix("condition_label_")
        value = text_by_name.get(f"condition_value_{suffix}")
        if value is None:
            issues.append(f"Condition label has no matching value: {label.name}")
            continue
        if abs(int(label.left) - int(value.left)) > 2 * EMU_PER_PX:
            issues.append(
                f"Condition label and value are not left-aligned: {label.name} / {value.name}"
            )
        if label.top + label.height > value.top + 2 * EMU_PER_PX:
            issues.append(
                f"Condition label overlaps its value: {label.name} / {value.name}"
            )
        if type_profile == "defense" and "\n" in label.text.strip():
            issues.append(f"Condition label contains a line break: {label.name}")
        if type_profile == "defense" and label.text_frame.word_wrap is not False:
            issues.append(f"Condition label must disable word wrap: {label.name}")
        if type_profile == "defense" and "\n" in value.text.strip():
            issues.append(f"Condition value wrapped across lines: {value.name}")

    figure_card = next(
        (shape for shape in slide.shapes if shape.name == "figure_card"),
        None,
    )
    figure_chrome = [
        text_by_name[name]
        for name in ("figure_label", "figure_badge", "figure_caption")
        if name in text_by_name
    ]
    if type_profile == "defense":
        wrapped_chrome = [
            chrome.name
            for chrome in figure_chrome
            if "\n" in chrome.text.strip()
        ]
        if wrapped_chrome:
            issues.append(
                "Figure-card chrome must be single-line in defense: "
                + ", ".join(wrapped_chrome)
            )
        if figure_card is not None:
            escaped_chrome = [
                chrome.name
                for chrome in figure_chrome
                if chrome.left < figure_card.left - EMU_PER_PX
                or chrome.top < figure_card.top - EMU_PER_PX
                or chrome.left + chrome.width > figure_card.left + figure_card.width + EMU_PER_PX
                or chrome.top + chrome.height > figure_card.top + figure_card.height + EMU_PER_PX
            ]
            if escaped_chrome:
                issues.append(
                    "Figure-card chrome escaped its card: "
                    + ", ".join(escaped_chrome)
                )

    nodes = [
        shape
        for shape in slide.shapes
        if shape.name.startswith("diagram_sequence_node_")
    ]
    bodies = [shape for shape in slide.shapes if shape.name == "diagram_node_body"]
    if nodes and len(nodes) == len(bodies):
        escaped = [
            node.name
            for node, body in zip(nodes, bodies)
            if body.top < node.top
            or body.left < node.left
            or body.top + body.height > node.top + node.height
            or body.left + body.width > node.left + node.width
        ]
        if escaped:
            issues.append("Diagram node body escaped its node: " + ", ".join(escaped))
    return issues


def main() -> None:
    args = parse_args()
    pptx_path = Path(args.pptx).resolve()
    deck_path = Path(args.deck).resolve()
    output_dir = Path(args.out).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    thresholds = {
        "body": args.min_body_pt if args.min_body_pt is not None else (18.0 if args.lecture_audit else None),
        "heading": args.min_heading_pt if args.min_heading_pt is not None else (22.0 if args.lecture_audit else None),
        "footnote": args.min_footnote_pt if args.min_footnote_pt is not None else (10.0 if args.lecture_audit else None),
    }
    deck = load_json(deck_path)
    type_profile = str((deck.get("meta") or {}).get("type_profile") or "")
    presentation = Presentation(str(pptx_path))
    slides = deck.get("slides", [])
    issues: list[dict] = []
    stats: list[dict] = []

    if len(presentation.slides) != len(slides):
        add_issue(issues, "error", "pptx", f"PPTX has {len(presentation.slides)} slides, deck has {len(slides)}.")
    for field in ("title", "author", "language"):
        if not str(getattr(presentation.core_properties, field, "") or "").strip():
            add_issue(issues, "warning", "core", f"core_properties.{field} is empty.")

    expected_formula_count = 0
    for slide in slides:
        if slide.get("layout") == "formula" and (
            str(slide.get("formula_latex", "")).strip() or str(slide.get("formula", "")).strip()
        ):
            expected_formula_count += 1
        for block in slide.get("blocks") or []:
            if block.get("formula_latex") or block.get("formula"):
                expected_formula_count += 1

    total_text = 0
    total_tables = 0
    total_charts = 0
    total_native_formula = 0
    role_sizes_all: dict[str, list[float]] = {"heading": [], "body": [], "footnote": []}
    below_threshold_pages: list[dict] = []
    role_violations: dict[str, list[dict]] = {"heading": [], "body": [], "footnote": []}
    for index, (slide, source) in enumerate(zip(presentation.slides, slides), start=1):
        page = str(source.get("id", f"{index:02d}"))
        text_shapes = [shape for shape in slide.shapes if shape.has_text_frame and shape.text.strip()]
        tables = [shape for shape in slide.shapes if shape.has_table]
        charts = [shape for shape in slide.shapes if getattr(shape, "has_chart", False)]
        pictures = [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
        formula_count = count_formula_shapes(slide)
        notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        total_text += len(text_shapes)
        total_tables += len(tables)
        total_charts += len(charts)
        total_native_formula += formula_count
        layout = str(source.get("layout", ""))
        for message in geometry_issues(
            slide,
            presentation.slide_width,
            presentation.slide_height,
            type_profile,
        ):
            severity = "warning" if message.startswith("Text shapes extend outside") else "error"
            add_issue(issues, severity, page, message)
        title_shape = slide.shapes.title
        title_text = title_shape.text.strip() if title_shape is not None else ""
        if not title_text:
            add_issue(issues, "error", page, "Slide has no non-empty title placeholder.")
        font_sizes = [
            round(run.font.size.pt, 1)
            for shape in slide.shapes
            if shape.has_text_frame
            for paragraph in shape.text_frame.paragraphs
            for run in paragraph.runs
            if run.font.size
        ]
        role_sizes: dict[str, list[float]] = {"heading": [], "body": [], "footnote": []}
        for shape in text_shapes:
            shape_role = classify_shape(shape)
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.font.size:
                        size = round(run.font.size.pt, 1)
                        role = "footnote" if size < 10 else shape_role
                        role_sizes[role].append(size)
        role_summary = {role: summarize_sizes(values) for role, values in role_sizes.items()}
        for role, values in role_sizes.items():
            role_sizes_all[role].extend(values)
            limit = thresholds.get(role)
            if limit is not None and values and min(values) < limit:
                violation = {"page": page, "role": role, "min_pt": min(values), "limit_pt": limit}
                below_threshold_pages.append(violation)
                role_violations[role].append(violation)

        min_text_shapes = minimum_editable_text_shapes(layout)
        if len(text_shapes) < min_text_shapes:
            add_issue(issues, "error", page, f"Only {len(text_shapes)} editable text shapes; page is likely still an image.")
        if layout == "cover" and len(text_shapes) < 3:
            add_issue(issues, "error", page, "Cover is missing editable title or metadata.")
        if not notes:
            add_issue(issues, "error", page, "Speaker notes are missing.")
        title = slide_display_title(source)
        normalized_title = re.sub(r"\s+", "", title)
        if title and not any(normalized_title in re.sub(r"\s+", "", shape.text) for shape in text_shapes):
            add_issue(issues, "warning", page, f"Slide title not found in editable text: {title}")
        if layout == "compare" and not tables:
            add_issue(issues, "error", page, "Compare layout should be a native PowerPoint table.")
        for table in tables:
            if not table.table.first_row:
                add_issue(issues, "error", page, f"Table '{table.name}' has no header-row semantics.")
        picture_alt_missing = []
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE or is_decorative(shape):
                continue
            alt = get_alt_text(shape)
            if not alt:
                picture_alt_missing.append(shape.name)
            elif Path(alt).suffix.lower() in {".png", ".jpg", ".jpeg", ".webp", ".emf", ".wmf"}:
                add_issue(issues, "warning", page, f"Picture '{shape.name}' alt text looks like a filename.")
        if picture_alt_missing:
            add_issue(issues, "error", page, "Pictures missing semantic alt text: " + ", ".join(picture_alt_missing))
        if source.get("chart") and not charts:
            add_issue(issues, "error", page, "Chart data is present but the slide has no native editable chart.")
        if layout == "formula":
            formula_expected = int(
                bool(
                    str(source.get("formula_latex", "")).strip()
                    or str(source.get("formula", "")).strip()
                )
            )
            if formula_expected and formula_count < formula_expected:
                add_issue(issues, "error", page, f"Only {formula_count}/{formula_expected} formulas are native OMML.")
        if layout == "grid":
            formula_expected = sum(1 for block in source.get("blocks") or [] if block.get("formula_latex") or block.get("formula"))
            if formula_expected and formula_count < formula_expected:
                add_issue(issues, "error", page, f"Only {formula_count}/{formula_expected} formulas are native OMML.")
        if len(text_shapes) <= 2 and len(pictures) == 1:
            add_issue(issues, "error", page, "Slide is a single full-page raster image.")
        stats.append(
            {
                "page": page,
                "layout": layout,
                "editable_text_shapes": len(text_shapes),
                "pictures": len(pictures),
                "native_tables": len(tables),
                "native_charts": len(charts),
                "native_formulas": formula_count,
                "notes_chars": len(notes),
                "has_title_placeholder": bool(title_text),
                "decorative_pictures": sum(
                    1
                    for shape in slide.shapes
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and is_decorative(shape)
                ),
                "font_size_min_pt": min(font_sizes) if font_sizes else None,
                "font_size_median_pt": round(statistics.median(font_sizes), 1) if font_sizes else None,
                "font_size_max_pt": max(font_sizes) if font_sizes else None,
                "text_runs_below_9pt": sum(1 for size in font_sizes if size < 9),
                "font_size_by_role": role_summary,
            }
        )

    if expected_formula_count and total_native_formula < expected_formula_count:
        add_issue(
            issues,
            "error",
            "formula",
            f"Only {total_native_formula}/{expected_formula_count} formulas are native OMML; fallback text was used.",
        )
    if total_text < len(slides) * 5:
        add_issue(issues, "error", "pptx", "Too few editable text shapes across the deck.")

    for role, violations in role_violations.items():
        if not violations:
            continue
        limit = thresholds.get(role)
        summary = summarize_sizes(role_sizes_all.get(role, []))
        values = role_sizes_all.get(role, [])
        below_ratio = sum(1 for value in values if value < limit) / max(1, len(values))
        median = summary.get("median")
        if median is not None and median >= limit and not (role == "body" and below_ratio > 0.35):
            continue
        page_count = len({item["page"] for item in violations})
        add_issue(
            issues,
            "warning",
            "font",
            f"{role.capitalize()} text median is {median} pt with {below_ratio:.0%} of runs below {limit:.1f} pt "
            f"across {page_count} page(s); see font_audit.below_threshold_pages for the detail.",
        )

    error_count = sum(1 for issue in issues if issue["severity"] == "error")
    warning_count = sum(1 for issue in issues if issue["severity"] == "warning")
    report = {
        "pptx": str(pptx_path),
        "deck": str(deck_path),
        "slide_count": len(presentation.slides),
        "editable_text_shapes": total_text,
        "native_tables": total_tables,
        "native_charts": total_charts,
        "native_formulas": total_native_formula,
        "expected_formulas": expected_formula_count,
        "error_count": error_count,
        "warning_count": warning_count,
        "font_audit": {
            "thresholds": thresholds,
            "roles": {role: summarize_sizes(values) for role, values in role_sizes_all.items()},
            "below_threshold_pages": below_threshold_pages,
        },
        "slides": stats,
        "issues": issues,
    }
    write_json(output_dir / "pptx_validation_report.json", report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if error_count:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
