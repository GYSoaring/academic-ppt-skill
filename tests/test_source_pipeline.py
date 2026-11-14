from __future__ import annotations

from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Inches
from docx import Document

from extract_source import extract_docx, extract_pptx
from import_pptx_theme import build_theme


def make_png(path: Path) -> None:
    image = Image.new("RGB", (320, 180), (20, 80, 160))
    image.save(path, "PNG")


def test_extract_pptx_text_and_picture(work_tmp: Path) -> None:
    image_path = work_tmp / "figure.png"
    make_png(image_path)
    deck_path = work_tmp / "source.pptx"
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = "证据链"
    slide.shapes.add_picture(str(image_path), Inches(1), Inches(1), Inches(2), Inches(1.2))
    table = slide.shapes.add_table(2, 2, Inches(1), Inches(3), Inches(4), Inches(1)).table
    table.cell(0, 0).text = "方法"
    table.cell(0, 1).text = "证据"
    table.cell(1, 0).text = "风险传播"
    table.cell(1, 1).text = "边级不确定性"
    presentation.save(deck_path)
    figures_dir = work_tmp / "figures"
    figures_dir.mkdir()
    records: list[dict] = []
    seen: set[str] = set()
    result = extract_pptx(deck_path, figures_dir, records, seen, 32, 20, True)
    assert "证据链" in result["pages"][0]["text"]
    assert "风险传播" in result["pages"][0]["text"]
    assert len(records) == 1
    assert (figures_dir / Path(records[0]["path"]).name).exists()


def test_extract_docx_paragraph(work_tmp: Path) -> None:
    docx_path = work_tmp / "source.docx"
    document = Document()
    document.add_heading("研究目标", level=1)
    document.add_paragraph("把观测转成可复核的边级证据。")
    document.save(docx_path)
    figures_dir = work_tmp / "figures"
    figures_dir.mkdir()
    records: list[dict] = []
    seen: set[str] = set()
    result = extract_docx(docx_path, figures_dir, records, seen, 32, 20)
    assert result["kind"] == "docx"
    assert any("边级证据" in page["text"] for page in result["pages"])


def test_extract_pdf_text_and_image(work_tmp: Path) -> None:
    import pymupdf as fitz

    image_path = work_tmp / "pdf_figure.png"
    make_png(image_path)
    pdf_path = work_tmp / "source.pdf"
    document = fitz.open()
    page = document.new_page()
    page.insert_text((72, 72), "Research Status")
    page.insert_image(fitz.Rect(72, 100, 240, 190), filename=str(image_path))
    document.save(pdf_path)
    document.close()
    figures_dir = work_tmp / "pdf_figures"
    figures_dir.mkdir()
    records: list[dict] = []
    seen: set[str] = set()
    from extract_source import extract_pdf

    result = extract_pdf(pdf_path, figures_dir, records, seen, 32, 20)
    assert "Research Status" in result["pages"][0]["text"]
    assert len(records) == 1


def test_import_pptx_theme_returns_semantic_theme(work_tmp: Path) -> None:
    deck_path = work_tmp / "source.pptx"
    presentation = Presentation()
    presentation.save(deck_path)
    theme = build_theme(deck_path, Path("assets/themes/academic_blue/theme.json"), "导入主题")
    assert theme["name"] == "导入主题"
    assert theme["colors"]["blue"].startswith("#")
    assert theme["pptx_fonts"]["body"]
    assert theme["imported_from"]["theme_part"].startswith("ppt/theme/")
