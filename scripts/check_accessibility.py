from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from accessibility import get_alt_text, is_decorative, shape_has_visible_text, shape_needs_alt_text
from common import write_json


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Check native PowerPoint accessibility semantics.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--out", required=True)
    return parser.parse_args()


def add_issue(issues: list[dict], severity: str, slide: int | str, message: str) -> None:
    issues.append({"severity": severity, "slide": slide, "message": message})


def main() -> None:
    args = parse_args()
    pptx_path = Path(args.pptx).resolve()
    output_dir = Path(args.out).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    presentation = Presentation(str(pptx_path))
    issues: list[dict] = []
    stats: list[dict] = []
    titles: dict[str, list[int]] = {}

    core = presentation.core_properties
    for field in ("title", "author", "language"):
        if not str(getattr(core, field, "") or "").strip():
            add_issue(issues, "warning", "core", f"core_properties.{field} is empty.")

    for index, slide in enumerate(presentation.slides, start=1):
        title_shape = slide.shapes.title
        title = title_shape.text.strip() if title_shape is not None else ""
        if not title:
            add_issue(issues, "error", index, "Slide has no non-empty title placeholder.")
        else:
            titles.setdefault(re.sub(r"\s+", "", title).lower(), []).append(index)

        table_count = 0
        picture_count = 0
        alt_count = 0
        decorative_count = 0
        missing_alt: list[str] = []
        nondecorative_without_text: list[str] = []
        for shape in slide.shapes:
            if getattr(shape, "has_table", False):
                table_count += 1
                if not shape.table.first_row:
                    add_issue(issues, "error", index, f"Table '{shape.name}' has no header-row semantics.")
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                picture_count += 1
                if is_decorative(shape):
                    decorative_count += 1
                    continue
                alt = get_alt_text(shape)
                if alt:
                    alt_count += 1
                    if Path(alt).suffix.lower() in {".png", ".jpg", ".jpeg", ".webp", ".emf", ".wmf"}:
                        add_issue(issues, "warning", index, f"Picture '{shape.name}' alt text looks like a filename.")
                else:
                    missing_alt.append(shape.name)
            elif shape_needs_alt_text(shape) and not is_decorative(shape):
                nondecorative_without_text.append(shape.name)
        if missing_alt:
            add_issue(issues, "error", index, "Pictures missing semantic alt text: " + ", ".join(missing_alt))
        if nondecorative_without_text:
            add_issue(
                issues,
                "warning",
                index,
                "Non-text shapes not marked decorative: " + ", ".join(nondecorative_without_text[:12]),
            )
        readable_text_shapes = [
            shape
            for shape in slide.shapes
            if not is_decorative(shape) and shape_has_visible_text(shape)
        ]
        title_first = bool(
            readable_text_shapes
            and title_shape is not None
            and readable_text_shapes[0].name == title_shape.name
        )
        if readable_text_shapes and not title_first:
            add_issue(
                issues,
                "warning",
                index,
                "Title placeholder is not first in screen-reader order; review Selection Pane.",
            )
        stats.append(
            {
                "slide": index,
                "title": title,
                "tables": table_count,
                "pictures": picture_count,
                "pictures_with_alt": alt_count,
                "decorative_pictures": decorative_count,
                "reading_order_ok": title_first,
            }
        )

    for title, pages in titles.items():
        if len(pages) > 1:
            add_issue(issues, "warning", "titles", f"Duplicate slide title on pages {pages}: {title}")

    errors = sum(1 for issue in issues if issue["severity"] == "error")
    warnings = sum(1 for issue in issues if issue["severity"] == "warning")
    report = {
        "pptx": str(pptx_path),
        "slide_count": len(presentation.slides),
        "error_count": errors,
        "warning_count": warnings,
        "issues": issues,
        "slides": stats,
    }
    write_json(output_dir / "accessibility_report.json", report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
