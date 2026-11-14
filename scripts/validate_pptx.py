from __future__ import annotations

import argparse
import json
import re
import statistics
from pathlib import Path
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from accessibility import get_alt_text, is_decorative
from common import load_json, write_json


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate native editable PowerPoint structure.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--deck", required=True)
    parser.add_argument("--out", required=True)
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


def main() -> None:
    args = parse_args()
    pptx_path = Path(args.pptx).resolve()
    deck_path = Path(args.deck).resolve()
    output_dir = Path(args.out).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    deck = load_json(deck_path)
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

        if len(text_shapes) < (3 if layout == "cover" else 5):
            add_issue(issues, "error", page, f"Only {len(text_shapes)} editable text shapes; page is likely still an image.")
        if layout == "cover" and len(text_shapes) < 3:
            add_issue(issues, "error", page, "Cover is missing editable title or metadata.")
        if not notes:
            add_issue(issues, "error", page, "Speaker notes are missing.")
        title = str(source.get("title", "")).strip()
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
        "slides": stats,
        "issues": issues,
    }
    write_json(output_dir / "pptx_validation_report.json", report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if error_count:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
