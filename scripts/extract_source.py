from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
from pathlib import Path
from typing import Iterable

from PIL import Image

from common import write_json


SUPPORTED_SUFFIXES = {
    ".pdf",
    ".docx",
    ".pptx",
    ".md",
    ".markdown",
    ".txt",
    ".png",
    ".jpg",
    ".jpeg",
    ".webp",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract text and figures from academic source files.")
    parser.add_argument("--input", nargs="+", required=True, help="Files or folders to extract.")
    parser.add_argument("--out", required=True)
    parser.add_argument("--min-image-size", type=int, default=64)
    parser.add_argument("--max-figures", type=int, default=200)
    parser.add_argument("--include-notes", action="store_true", default=True)
    return parser.parse_args()


def slugify(text: str, fallback: str = "figure") -> str:
    value = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", text.strip(), flags=re.UNICODE).strip("._")
    return value[:64] or fallback


def clean_context(text: str, limit: int = 180) -> str:
    value = re.sub(r"\s+", " ", text or "").strip()
    return value[:limit]


def iter_inputs(values: Iterable[str]) -> list[Path]:
    paths: list[Path] = []
    for value in values:
        path = Path(value)
        if path.is_dir():
            paths.extend(
                child
                for child in sorted(path.rglob("*"))
                if child.is_file() and child.suffix.lower() in SUPPORTED_SUFFIXES
            )
        elif path.is_file():
            paths.append(path)
    return paths


def save_image_bytes(
    data: bytes,
    figures_dir: Path,
    figure_records: list[dict],
    seen_hashes: set[str],
    *,
    source: str,
    page: int | str | None,
    kind: str,
    context: str,
    min_size: int,
    max_figures: int,
    bbox: list[float] | None = None,
) -> dict | None:
    if len(figure_records) >= max_figures:
        return None
    digest = hashlib.sha256(data).hexdigest()
    if digest in seen_hashes:
        return None
    try:
        image = Image.open(io.BytesIO(data))
        width, height = image.size
        if width < min_size or height < min_size:
            return None
        if image.mode not in {"RGB", "RGBA"}:
            image = image.convert("RGB")
        figure_id = f"fig_{len(figure_records) + 1:03d}"
        target = figures_dir / f"{figure_id}_{slugify(context, 'figure')}.png"
        image.save(target, "PNG", optimize=True)
    except Exception:
        return None
    seen_hashes.add(digest)
    record = {
        "id": figure_id,
        "path": f"figures/{target.name}",
        "source": source,
        "page": page,
        "kind": kind,
        "width": width,
        "height": height,
        "sha256": digest,
        "alt_hint": clean_context(context),
    }
    if bbox:
        record["bbox"] = bbox
    figure_records.append(record)
    return record


def extract_pdf(
    path: Path,
    figures_dir: Path,
    figure_records: list[dict],
    seen_hashes: set[str],
    min_size: int,
    max_figures: int,
) -> dict:
    import pymupdf as fitz

    document = fitz.open(str(path))
    pages: list[dict] = []
    try:
        title = str(document.metadata.get("title") or path.stem)
        for page_number, page in enumerate(document, start=1):
            text = page.get_text("text")
            pages.append({"page": page_number, "text": text})
            for image_info in page.get_images(full=True):
                xref = image_info[0]
                extracted = document.extract_image(xref)
                data = extracted.get("image")
                if not data:
                    continue
                rects = page.get_image_rects(xref)
                bbox = [rects[0].x0, rects[0].y0, rects[0].x1, rects[0].y1] if rects else None
                save_image_bytes(
                    data,
                    figures_dir,
                    figure_records,
                    seen_hashes,
                    source=str(path),
                    page=page_number,
                    kind="pdf_page",
                    context=text or "PDF figure",
                    min_size=min_size,
                    max_figures=max_figures,
                    bbox=bbox,
                )
    finally:
        document.close()
    return {
        "title": title,
        "kind": "pdf",
        "pages": pages,
        "text_chars": sum(len(page["text"]) for page in pages),
    }


def extract_docx(
    path: Path,
    figures_dir: Path,
    figure_records: list[dict],
    seen_hashes: set[str],
    min_size: int,
    max_figures: int,
) -> dict:
    from docx import Document

    document = Document(str(path))
    blocks: list[dict] = []
    paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
    for index, paragraph in enumerate(document.paragraphs, start=1):
        text = paragraph.text.strip()
        if not text:
            continue
        blocks.append({"page": f"paragraph-{index}", "text": text})
        for blip in paragraph._p.xpath(".//a:blip"):
            rel_id = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if not rel_id or rel_id not in document.part.related_parts:
                continue
            data = document.part.related_parts[rel_id].blob
            save_image_bytes(
                data,
                figures_dir,
                figure_records,
                seen_hashes,
                source=str(path),
                page=f"paragraph-{index}",
                kind="docx_paragraph",
                context=text,
                min_size=min_size,
                max_figures=max_figures,
            )
    for table_index, table in enumerate(document.tables, start=1):
        rows = []
        for row in table.rows:
            rows.append(" | ".join(cell.text.strip().replace("\n", " ") for cell in row.cells))
        blocks.append({"page": f"table-{table_index}", "text": "\n".join(rows)})
    for shape_index, shape in enumerate(document.inline_shapes, start=1):
        blip = shape._inline.graphic.graphicData.pic.blipFill.blip
        rel_id = blip.embed
        if not rel_id or rel_id not in document.part.related_parts:
            continue
        save_image_bytes(
            document.part.related_parts[rel_id].blob,
            figures_dir,
            figure_records,
            seen_hashes,
            source=str(path),
            page=f"inline-{shape_index}",
            kind="docx_inline",
            context=paragraphs[0] if paragraphs else path.stem,
            min_size=min_size,
            max_figures=max_figures,
        )
    return {
        "title": str(document.core_properties.title or path.stem),
        "kind": "docx",
        "pages": blocks,
        "text_chars": sum(len(block["text"]) for block in blocks),
    }


def extract_pptx(
    path: Path,
    figures_dir: Path,
    figure_records: list[dict],
    seen_hashes: set[str],
    min_size: int,
    max_figures: int,
    include_notes: bool,
) -> dict:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    presentation = Presentation(str(path))
    slides: list[dict] = []
    for slide_number, slide in enumerate(presentation.slides, start=1):
        texts: list[str] = []
        notes = ""
        if include_notes and slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                texts.append(shape.text_frame.text.strip())
            if getattr(shape, "has_table", False):
                rows = []
                for row in shape.table.rows:
                    rows.append(" | ".join(cell.text.strip().replace("\n", " ") for cell in row.cells))
                if rows:
                    texts.append("[table]\n" + "\n".join(rows))
            if getattr(shape, "has_chart", False):
                try:
                    chart = shape.chart
                    series_names = [str(series.name) for series in chart.series]
                    texts.append("[chart] " + ", ".join(series_names))
                except Exception:
                    pass
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    data = shape.image.blob
                except Exception:
                    continue
                save_image_bytes(
                    data,
                    figures_dir,
                    figure_records,
                    seen_hashes,
                    source=str(path),
                    page=slide_number,
                    kind="pptx_slide",
                    context=texts[0] if texts else path.stem,
                    min_size=min_size,
                    max_figures=max_figures,
                )
        text = "\n".join(texts)
        if notes:
            text = f"{text}\n\n[notes]\n{notes}".strip()
        slides.append({"page": slide_number, "text": text})
    return {
        "title": path.stem,
        "kind": "pptx",
        "pages": slides,
        "text_chars": sum(len(slide["text"]) for slide in slides),
    }


def extract_text(path: Path) -> dict:
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    return {"title": path.stem, "kind": "text", "pages": [{"page": "document", "text": text}], "text_chars": len(text)}


def extract_image_file(
    path: Path,
    figures_dir: Path,
    figure_records: list[dict],
    seen_hashes: set[str],
    min_size: int,
    max_figures: int,
) -> dict:
    data = path.read_bytes()
    save_image_bytes(
        data,
        figures_dir,
        figure_records,
        seen_hashes,
        source=str(path),
        page=None,
        kind="image",
        context=path.stem,
        min_size=min_size,
        max_figures=max_figures,
    )
    return {"title": path.stem, "kind": "image", "pages": [], "text_chars": 0}


def build_markdown(source_records: list[dict], figures: list[dict]) -> str:
    lines = ["# Extracted Academic Source", "", "> Generated by `extract_source.py`. Review source text and figure roles before building a deck.", ""]
    for source in source_records:
        lines.extend([f"## {source['path']}", "", f"- Type: `{source['kind']}`", f"- Title: {source['title']}", ""])
        for page in source["pages"]:
            label = f"Page {page['page']}" if source["kind"] == "pdf" else f"Section {page['page']}"
            lines.extend([f"### {label}", "", page["text"].strip() or "（无文本）", ""])
    if figures:
        lines.extend(["## Figure Manifest", ""])
        for figure in figures:
            lines.append(f"- `{figure['id']}` — {figure['path']} — page {figure['page']} — {figure['alt_hint']}")
        lines.append("")
    return "\n".join(lines)


def main() -> None:
    args = parse_args()
    output_dir = Path(args.out).resolve()
    figures_dir = output_dir / "figures"
    figures_dir.mkdir(parents=True, exist_ok=True)
    inputs = iter_inputs(args.input)
    if not inputs:
        raise SystemExit("No supported input files found.")
    source_records: list[dict] = []
    figure_records: list[dict] = []
    seen_hashes: set[str] = set()
    errors: list[dict] = []
    for path in inputs:
        suffix = path.suffix.lower()
        try:
            if suffix == ".pdf":
                record = extract_pdf(path, figures_dir, figure_records, seen_hashes, args.min_image_size, args.max_figures)
            elif suffix == ".docx":
                record = extract_docx(path, figures_dir, figure_records, seen_hashes, args.min_image_size, args.max_figures)
            elif suffix == ".pptx":
                record = extract_pptx(path, figures_dir, figure_records, seen_hashes, args.min_image_size, args.max_figures, args.include_notes)
            elif suffix in {".png", ".jpg", ".jpeg", ".webp"}:
                record = extract_image_file(path, figures_dir, figure_records, seen_hashes, args.min_image_size, args.max_figures)
            else:
                record = extract_text(path)
            record["path"] = str(path)
            record["suffix"] = suffix
            source_records.append(record)
        except Exception as error:
            errors.append({"path": str(path), "error": str(error)})
    (output_dir / "source.md").write_text(build_markdown(source_records, figure_records), encoding="utf-8")
    manifest = {
        "schema_version": 1,
        "generator": "academic-ppt-skill/extract_source.py",
        "sources": source_records,
        "figures": figure_records,
        "errors": errors,
        "counts": {
            "sources": len(source_records),
            "pages_or_sections": sum(len(source["pages"]) for source in source_records),
            "figures": len(figure_records),
            "errors": len(errors),
        },
    }
    write_json(output_dir / "source_manifest.json", manifest)
    write_json(output_dir / "figures.json", {"figures": figure_records})
    print(
        json.dumps(
            {
                "out": str(output_dir),
                "sources": manifest["counts"]["sources"],
                "figures": manifest["counts"]["figures"],
                "errors": errors,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    if errors and not source_records:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
