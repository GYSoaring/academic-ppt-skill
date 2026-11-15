from __future__ import annotations

from build_gallery import build_deck
from validate_deck import GENRES, LAYOUT_VARIANTS


def test_gallery_covers_layouts_and_diagram_kinds() -> None:
    deck = build_deck()
    slides = deck["slides"]
    layouts = {slide["layout"] for slide in slides}
    assert {
        "cover",
        "outline",
        "section",
        "bullets",
        "grid",
        "concept",
        "formula",
        "pipeline",
        "compare",
        "hub",
        "metrics",
        "figure",
        "figure_flow",
        "balance",
        "timeline",
        "quote",
        "closing",
        "references",
    } <= layouts
    diagram_kinds = set()
    for slide in slides:
        if isinstance(slide.get("diagram"), dict):
            diagram_kinds.add(slide["diagram"].get("kind"))
        for block in slide.get("blocks") or []:
            if isinstance(block.get("diagram"), dict):
                diagram_kinds.add(block["diagram"].get("kind"))
    assert {"sequence", "branch", "stack", "ruler", "matrix", "cycle"} <= diagram_kinds


def test_gallery_slides_have_speech_and_quote_contract() -> None:
    deck = build_deck()
    assert all(str(slide.get("speech", "")).strip() for slide in deck["slides"])
    quotes = [slide for slide in deck["slides"] if slide["layout"] == "quote"]
    assert quotes
    assert all(slide.get("quote") and slide.get("attribution") and slide.get("commentary") for slide in quotes)


def test_genre_and_variant_contracts_are_declared() -> None:
    assert "humanities" in GENRES
    assert "theory" in GENRES
    assert "image_right" in LAYOUT_VARIANTS["figure"]
    assert "split" in LAYOUT_VARIANTS["quote"]
