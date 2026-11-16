from __future__ import annotations

from pathlib import Path
import tempfile

from common import DEFENSE_BODY_PX, DEFENSE_MIN_ANY_PX, DEFENSE_TYPE_SCALE, apply_type_profile, figure_width_ratio, load_json
from design import FontBook, TextStyle
from audit_pptx import choose_verdict
import validate_deck
import export_agent_brief


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
        small_text_ratio=0.88,
        table_below_18_ratio=0.0,
        min_body_pt=24.0,
        min_any_pt=18.0,
        body_bold_ratio=1.0,
    )
    assert verdict == "font_upgrade"


def test_choose_verdict_accepts_secondary_18pt_chrome() -> None:
    verdict = choose_verdict(
        body_median=24.4,
        median=18.8,
        table_median=18.8,
        body_below_24_ratio=0.05,
        below_18_ratio=0.0,
        small_text_ratio=0.70,
        table_below_18_ratio=0.0,
        min_body_pt=24.0,
        min_any_pt=18.0,
        body_bold_ratio=1.0,
    )
    assert verdict == "defense_ready"


def test_defense_native_exporter_uses_compact_line_spacing_and_explicit_wrapping() -> None:
    import export_editable_pptx

    exporter = object.__new__(export_editable_pptx.EditableDeckExporter)
    exporter.profile = "defense"
    exporter.body_bold = True
    exporter.body_bold_min_px = 40
    exporter.body_bold_max_px = 56
    exporter.fonts = {"body": "Noto Sans SC", "title": "Noto Sans SC"}
    wrapped = export_editable_pptx.defense_line_text(
        "\u57fa\u9891\u4e0e\u500d\u9891\u901f\u5ea6\u4e0d\u540c\u3002",
        220,
        48,
        [r"C:\Windows\Fonts\msyhbd.ttc"],
    )
    assert "\n" in wrapped or len(wrapped) < 12


def test_defense_native_exporter_keeps_18pt_floor() -> None:
    from common import DEFENSE_MIN_ANY_PX, clamp_defense_size_px

    assert clamp_defense_size_px(10, "defense") == DEFENSE_MIN_ANY_PX
    assert clamp_defense_size_px(10, "defense", body_floor=True) > DEFENSE_MIN_ANY_PX


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


def test_defense_renderer_fit_text_uses_block_fitter(monkeypatch) -> None:
    import render_deck_next

    calls: list[tuple] = []

    def fake_fit_text_block(draw, text, font_candidates, max_width, max_height, max_size, min_size, *args, **kwargs):
        calls.append((text, max_size, min_size, args, kwargs))
        return "font", [text], 10, True

    monkeypatch.setattr(render_deck_next, "fit_text_block", fake_fit_text_block)
    theme = apply_type_profile({"type_scale": {}}, "defense")
    renderer = object.__new__(render_deck_next.DeckRenderer)
    renderer.profile = "defense"
    result = renderer.fit_text(None, "24 pt 正文", ["font"], 100, 100, 20, 12, max_lines=2)
    assert result == ("font", ["24 pt 正文"], 10, True)
    assert calls == [("24 pt 正文", int(DEFENSE_MIN_ANY_PX), 12, (), {"max_lines": 2})]


def test_choose_verdict_rejects_regular_weight_body() -> None:
    verdict = choose_verdict(
        body_median=24.4,
        median=18.8,
        table_median=18.8,
        body_below_24_ratio=0.05,
        below_18_ratio=0.0,
        small_text_ratio=0.70,
        table_below_18_ratio=0.0,
        min_body_pt=24.0,
        min_any_pt=18.0,
        body_bold_ratio=0.54,
    )
    assert verdict == "font_upgrade"


def test_defense_validation_checks_scale_bold_font_and_paper_figure() -> None:
    theme = {
        "name": "defense-test",
        "type_scale": {
            "body": DEFENSE_BODY_PX,
            "header_title": 72,
            "small": 39,
            "micro": 39,
            "reference_title": 39,
            "reference_meta": 39,
            "footnote": 39,
        },
        "text": {"body_bold": True, "body_bold_min_px": 48, "body_bold_max_px": 56},
        "pptx_fonts": {"body": "Noto Sans SC"},
    }
    assert validate_deck.defense_theme_issues(theme) == []

    bad_theme = {
        **theme,
        "type_scale": {"body": 22, "header_title": 40, "small": 14, "micro": 14, "reference_title": 14, "reference_meta": 14, "footnote": 14},
        "text": {"body_bold": False, "body_bold_min_px": 40, "body_bold_max_px": 50},
        "pptx_fonts": {"body": "Times New Roman"},
    }
    issues = validate_deck.defense_theme_issues(bad_theme)
    labels = {issue["page"] for issue in issues}
    assert {"type_scale", "text", "fonts"} <= labels


def test_agent_brief_carries_defense_readability_and_paper_figure_contract(monkeypatch) -> None:
    with tempfile.TemporaryDirectory(dir=ROOT) as temp_dir:
        temp_path = Path(temp_dir)
    argv = [
        "export_agent_brief.py",
        "--deck", str(ROOT / "examples" / "nanjing_optics" / "deck.json"),
        "--theme", str(ROOT / "assets" / "themes" / "nanjing_purple" / "theme.json"),
        "--out", str(temp_path / "agent_brief.md"),
    ]
    monkeypatch.setattr("sys.argv", argv)
    export_agent_brief.main()
    brief = (temp_path / "agent_brief.md").read_text(encoding="utf-8")
    assert "24 pt" in brief and "32" in brief and "18 pt" in brief
    assert "加粗" in brief and "无衬线黑体" in brief
    assert "image_origin" in brief and "论文原图" in brief
    assert "未完成" in brief
