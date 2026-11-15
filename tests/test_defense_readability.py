from __future__ import annotations

from pathlib import Path

from common import DEFENSE_TYPE_SCALE, apply_type_profile, figure_width_ratio, load_json
from design import FontBook, TextStyle
from audit_pptx import choose_verdict


ROOT = Path(__file__).resolve().parents[1]


def test_defense_profile_sets_projection_sizes() -> None:
    theme = {"type_scale": {"body": 22, "header_title": 46}}
    result = apply_type_profile(theme, "defense")
    assert result["active_type_profile"] == "defense"
    assert result["type_scale"]["body"] == DEFENSE_TYPE_SCALE["body"]
    assert result["type_scale"]["header_title"] >= 72
    assert result["type_scale"]["eyebrow"] < result["type_scale"]["header_title"]


def test_figure_width_ratio_is_clamped_and_accepts_percent() -> None:
    assert figure_width_ratio({}, 0.66) == 0.66
    assert figure_width_ratio({"image_width_ratio": 75}, 0.66) == 0.75
    assert figure_width_ratio({"rail_ratio": 0.9}, 0.66) == 0.82
    assert figure_width_ratio({"image_width_ratio": 0.2}, 0.66) == 0.45


def test_font_book_bolds_only_semantic_body_range() -> None:
    regular = ["regular-font"]
    bold = ["bold-font"]
    fonts = FontBook(bold, regular, body_bold=True, body_bold_min_px=40, body_bold_max_px=56)
    assert fonts.title_paths == bold
    assert fonts.body_paths == regular
    assert fonts.body_bold_min_px == 40
    assert fonts.body_bold_max_px == 56
    assert fonts._cache == {}


def test_choose_verdict_rejects_hidden_small_text() -> None:
    verdict = choose_verdict(
        body_median=24.4,
        median=18.8,
        table_median=18.8,
        body_below_24_ratio=0.05,
        below_18_ratio=0.46,
        below_24_ratio=0.88,
        table_below_18_ratio=0.0,
        min_body_pt=24.0,
        min_any_pt=18.0,
    )
    assert verdict == "font_upgrade"


def test_nanjing_example_has_brand_and_figure_provenance() -> None:
    deck = load_json(ROOT / "examples" / "nanjing_optics" / "deck.json")
    assert len(deck["slides"]) == 16
    assert deck["meta"]["genre"] == "theory"
    theme = load_json(ROOT / "assets" / "themes" / "nanjing_purple" / "theme.json")
    assert theme["text"]["body_bold"] is True
    assert theme["text"]["body_bold_min_px"] <= DEFENSE_TYPE_SCALE["body"] <= theme["text"]["body_bold_max_px"]
    assert (ROOT / "assets" / "themes" / "nanjing_purple" / "SOURCE.md").exists()
    figure_pages = [
        slide for slide in deck["slides"]
        if slide.get("layout") == "figure" and slide.get("image")
    ]
    assert figure_pages
    for slide in figure_pages:
        assert slide.get("source")
        assert slide.get("image_role")
        assert slide.get("image_boundary")
        assert 0.65 <= figure_width_ratio(slide) <= 0.82
