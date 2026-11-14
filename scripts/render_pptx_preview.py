from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import tempfile
from pathlib import Path

import numpy as np
from PIL import Image

from common import write_json
from accessibility import image_looks_blank


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render a real PPTX through PowerPoint or LibreOffice for visual QA.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--width", type=int, default=1920)
    parser.add_argument("--height", type=int, default=1080)
    parser.add_argument("--compare-dir", default=None, help="Optional render_deck.py rendered/ directory.")
    parser.add_argument("--strict", action="store_true", help="Fail on blank/duplicate rendered pages.")
    return parser.parse_args()


def find_powerpoint() -> Path | None:
    candidates = [
        shutil.which("POWERPNT.EXE"),
        r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return Path(candidate)
    return None


def find_libreoffice() -> Path | None:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return Path(candidate)
    return None


def export_with_powerpoint(pptx: Path, out_dir: Path, width: int, height: int) -> int:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    presentation = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            app.Visible = False
        except Exception:
            pass
        presentation = app.Presentations.Open(str(pptx), ReadOnly=True, WithWindow=False)
        for index, slide in enumerate(presentation.Slides, start=1):
            slide.Export(str(out_dir / f"pptx_slide_{index:02d}.png"), "PNG", width, height)
        return int(presentation.Slides.Count)
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def export_with_libreoffice(pptx: Path, out_dir: Path, width: int, height: int) -> int:
    import fitz

    soffice = find_libreoffice()
    if soffice is None:
        raise FileNotFoundError("LibreOffice is not installed.")
    with tempfile.TemporaryDirectory(prefix="academic_ppt_render_") as temp:
        temp_dir = Path(temp)
        result = subprocess.run(
            [
                str(soffice),
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(temp_dir),
                str(pptx),
            ],
            capture_output=True,
            text=True,
            timeout=180,
        )
        pdf_path = temp_dir / f"{pptx.stem}.pdf"
        if result.returncode != 0 or not pdf_path.exists():
            raise RuntimeError((result.stderr or result.stdout or "LibreOffice conversion failed").strip())
        document = fitz.open(str(pdf_path))
        try:
            for index, page in enumerate(document, start=1):
                scale = min(width / page.rect.width, height / page.rect.height)
                pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                pixmap.save(str(out_dir / f"pptx_slide_{index:02d}.png"))
            return len(document)
        finally:
            document.close()


def image_digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def compare_with_preview(render_paths: list[Path], compare_dir: Path) -> list[dict]:
    comparisons = []
    for index, rendered in enumerate(render_paths, start=1):
        preview = compare_dir / f"slide_{index:02d}.png"
        if not preview.exists():
            comparisons.append({"slide": index, "status": "missing_preview"})
            continue
        with Image.open(rendered) as rendered_image, Image.open(preview) as preview_image:
            size = (640, 360)
            first = np.asarray(rendered_image.convert("RGB").resize(size, Image.Resampling.LANCZOS), dtype=np.int16)
            second = np.asarray(preview_image.convert("RGB").resize(size, Image.Resampling.LANCZOS), dtype=np.int16)
            difference = np.abs(first - second)
        comparisons.append(
            {
                "slide": index,
                "status": "compared",
                "mean_abs_diff": round(float(difference.mean()), 3),
                "p95_abs_diff": round(float(np.percentile(difference, 95)), 3),
            }
        )
    return comparisons


def main() -> None:
    args = parse_args()
    pptx_path = Path(args.pptx).resolve()
    output_dir = Path(args.out).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    for old in output_dir.glob("pptx_slide_*.png"):
        old.unlink()

    report: dict = {
        "pptx": str(pptx_path),
        "renderer": "unavailable",
        "status": "skipped",
        "slide_count": 0,
        "blank_pages": [],
        "duplicate_pages": [],
        "page_sizes": {},
        "comparisons": [],
        "message": "",
    }
    render_paths: list[Path] = []
    if find_powerpoint() is not None:
        try:
            count = export_with_powerpoint(pptx_path, output_dir, args.width, args.height)
            report["renderer"] = "powerpoint-com"
            report["status"] = "ok"
            report["slide_count"] = count
        except Exception as error:
            report["message"] = f"PowerPoint COM failed: {error}"
    if report["status"] != "ok" and find_libreoffice() is not None:
        try:
            count = export_with_libreoffice(pptx_path, output_dir, args.width, args.height)
            report["renderer"] = "libreoffice"
            report["status"] = "ok"
            report["slide_count"] = count
        except Exception as error:
            report["message"] = (report["message"] + " | " + f"LibreOffice failed: {error}").strip(" |")
    if report["status"] != "ok" and not report["message"]:
        report["message"] = "Neither PowerPoint COM nor LibreOffice is available; visual QA was skipped."

    if report["status"] == "ok":
        render_paths = sorted(output_dir.glob("pptx_slide_*.png"))
        digests: dict[str, list[int]] = {}
        for index, path in enumerate(render_paths, start=1):
            with Image.open(path) as image:
                report["page_sizes"][str(index)] = list(image.size)
            if image_looks_blank(path):
                report["blank_pages"].append(index)
            digests.setdefault(image_digest(path), []).append(index)
        report["duplicate_pages"] = [pages for pages in digests.values() if len(pages) > 1]
        if args.compare_dir:
            report["comparisons"] = compare_with_preview(render_paths, Path(args.compare_dir).resolve())

    report["rendered_files"] = [str(path) for path in render_paths]
    write_json(output_dir / "pptx_render_report.json", report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report["status"] == "ok" and args.strict and (report["blank_pages"] or report["duplicate_pages"]):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
