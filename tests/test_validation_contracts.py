from __future__ import annotations

from validate_deck import (
    closing_has_summary_content,
    longest_content_layout_run,
    summary_title_valid,
)


def test_summary_title_accepts_five_character_topic() -> None:
    assert summary_title_valid("多维特征链")
    assert summary_title_valid("汇报提纲")
    assert not summary_title_valid("短标题")
    assert not summary_title_valid("一" * 19)


def test_closing_accepts_points_or_boundary_and_next_cards() -> None:
    assert closing_has_summary_content({"points": ["结论"]})
    assert closing_has_summary_content({"boundary": "适用边界"})
    assert closing_has_summary_content({"boundary_points": ["边界"]})
    assert closing_has_summary_content({"next": "下一步"})
    assert closing_has_summary_content({"next_points": ["展望"]})
    assert not closing_has_summary_content({"layout": "closing"})


def test_section_divider_breaks_layout_run_and_longest_layout_is_reported() -> None:
    slides = [
        {"layout": "figure"},
        {"layout": "figure"},
        {"layout": "section"},
        {"layout": "figure"},
        {"layout": "figure"},
    ]
    assert longest_content_layout_run(slides) == (2, "figure")

    slides = [
        {"layout": "figure"},
        {"layout": "figure"},
        {"layout": "figure"},
        {"layout": "figure"},
        {"layout": "metrics"},
    ]
    assert longest_content_layout_run(slides) == (4, "figure")
