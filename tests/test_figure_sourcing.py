from __future__ import annotations

import json
from pathlib import Path

import pymupdf as fitz
from PIL import Image

from extract_source import extract_pdf
from find_figures import page_figure_regions
from suggest_figures import apply_suggestions, tokens


def make_pdf_with_figure(path: Path) -> None:
    image = Image.new("RGB", (400, 260), (40, 90, 150))
    image_path = path.with_suffix(".png")
    image.save(image_path)
    document = fitz.open()
    page = document.new_page(width=595, height=842)
    page.insert_image(fitz.Rect(60, 80, 520, 380), filename=str(image_path))
    page.insert_text((60, 420), "Figure 1: A synthetic architecture for testing.", fontsize=13)
    document.save(path)
    document.close()


def test_extract_pdf_captures_figure_caption(work_tmp: Path) -> None:
    pdf_path = work_tmp / "paper.pdf"
    make_pdf_with_figure(pdf_path)
    figures_dir = work_tmp / "figures"
    figures_dir.mkdir()
    records: list[dict] = []
    seen: set[str] = set()
    result = extract_pdf(pdf_path, figures_dir, records, seen, 32, 10)
    assert result["kind"] == "pdf"
    assert records
    assert any("Figure 1" in str(record.get("caption", "")) for record in records)


def test_page_figure_regions_uses_caption(work_tmp: Path) -> None:
    pdf_path = work_tmp / "paper.pdf"
    make_pdf_with_figure(pdf_path)
    document = fitz.open(pdf_path)
    regions = page_figure_regions(document[0])
    document.close()
    assert regions
    assert regions[0][1].startswith("Figure 1")


def test_glossary_expands_chinese_tokens() -> None:
    result = tokens("视觉语言机器人技能执行")
    assert {"vision", "language", "robot", "skill", "execution"} <= result


def test_apply_suggestions_copies_figure_and_source(work_tmp: Path) -> None:
    figure_dir = work_tmp / "pack"
    figure_dir.mkdir()
    image = Image.new("RGB", (320, 200), (20, 80, 160))
    image_path = figure_dir / "figures" / "f1.png"
    image_path.parent.mkdir()
    image.save(image_path)
    deck_path = work_tmp / "deck.json"
    deck = {
        "meta": {"title": "Test"},
        "slides": [
            {
                "id": "03",
                "layout": "grid",
                "section": "content",
                "title": "测试页",
                "lead": "测试",
                "speech": "测试",
                "blocks": [
                    {"heading": "A"},
                    {"heading": "B"},
                    {"heading": "C"},
                ],
            }
        ],
    }
    deck_path.write_text(json.dumps(deck, ensure_ascii=False), encoding="utf-8")
    out_deck = work_tmp / "deck.with_figures.json"
    asset_dir = work_tmp / "assets"
    figure = {
        "id": "f1",
        "path": "figures/f1.png",
        "caption": "Figure 1: A test figure.",
        "paper_title": "Test Paper",
        "source": "https://arxiv.org/abs/0000.0000",
        "page": 2,
    }
    applied = apply_suggestions(
        deck_path,
        deck,
        {"03": [{"id": "f1", "score": 0.5}]},
        [figure],
        figure_dir,
        out_deck,
        asset_dir,
    )
    assert applied == 1
    updated = json.loads(out_deck.read_text(encoding="utf-8"))
    slide = updated["slides"][0]
    assert slide["image_origin"] == "borrowed"
    assert "arxiv.org" in slide["source"]
    assert (work_tmp / slide["image"]).exists()
