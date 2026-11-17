from __future__ import annotations

from add_section_pages import insert_section_pages
from common import (
    apply_navigation_mode,
    content_section_openers,
    navigation_mode_for,
)


def _deck(mode: str | None = None) -> dict:
    meta = {"title": "导航模式测试"}
    if mode:
        meta["navigation_mode"] = mode
    return {
        "meta": meta,
        "sections": [
            {"id": "background", "title": "研究背景"},
            {"id": "method", "title": "研究方法"},
        ],
        "slides": [
            {"id": "01", "layout": "cover", "title": "封面"},
            {"id": "02", "layout": "outline", "title": "汇报提纲"},
            {"id": "03", "layout": "section", "section": "background", "title": "研究背景"},
            {"id": "04", "layout": "bullets", "section": "background", "title": "研究问题"},
            {"id": "05", "layout": "section", "section": "method", "title": "研究方法"},
            {"id": "06", "layout": "figure", "section": "method", "title": "方法总览"},
        ],
    }


def _theme(sidebar_mode: str = "always") -> dict:
    return {"layout": {"sidebar_mode": sidebar_mode}}


def test_explicit_navigation_mode_wins_and_maps_to_sidebar_policy() -> None:
    deck = _deck("hybrid")
    theme = _theme("off")

    mode = navigation_mode_for(deck, theme)
    applied = apply_navigation_mode(theme, mode)

    assert mode == "hybrid"
    assert applied == "hybrid"
    assert theme["layout"]["sidebar_mode"] == "section"


def test_legacy_sidebar_policy_is_preserved() -> None:
    assert navigation_mode_for(_deck(), _theme("always")) == "sidebar"
    assert navigation_mode_for(_deck(), _theme("section")) == "hybrid"
    assert navigation_mode_for(_deck(), _theme("off")) == "dividers"


def test_section_openers_exclude_divider_pages() -> None:
    assert content_section_openers(_deck()) == {"04", "06"}


def test_add_section_pages_skips_sidebar_mode() -> None:
    result, report = insert_section_pages(_deck(), "sidebar", persist_mode=True)

    assert report["status"] == "skipped"
    assert report["inserted"] == 0
    assert [slide["layout"] for slide in result["slides"]].count("section") == 2
    assert result["meta"]["navigation_mode"] == "sidebar"


def test_add_section_pages_inserts_for_dividers_and_persists_mode() -> None:
    deck = {
        "meta": {"title": "导航模式测试"},
        "sections": [
            {"id": "background", "title": "研究背景"},
            {"id": "method", "title": "研究方法"},
        ],
        "slides": [
            {"id": "01", "layout": "cover", "title": "封面"},
            {"id": "02", "layout": "outline", "title": "汇报提纲"},
            {"id": "03", "layout": "bullets", "section": "background", "title": "研究问题"},
            {"id": "04", "layout": "figure", "section": "method", "title": "方法总览"},
        ],
    }

    result, report = insert_section_pages(deck, "dividers", persist_mode=True)

    assert report["status"] == "inserted"
    assert report["inserted"] == 2
    assert result["meta"]["navigation_mode"] == "dividers"
    assert [slide["layout"] for slide in result["slides"]] == [
        "cover",
        "outline",
        "section",
        "bullets",
        "section",
        "figure",
    ]
