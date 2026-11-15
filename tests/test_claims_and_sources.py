from __future__ import annotations

from build_ghost_deck import build
from common import is_assertion_title, normalize_cites, source_line_for_slide
from validate_deck import validate_citation_entry


def test_normalize_cites_dedupes_and_keeps_order() -> None:
    assert normalize_cites(["b", "a", "b", " "]) == ["b", "a"]
    assert normalize_cites("key") == ["key"]
    assert normalize_cites(None) == []


def test_source_line_prefers_explicit_source() -> None:
    deck = {"citations": {"a": {"title": "A", "venue": "V", "year": "2024"}}}
    slide = {"source": "实验室自测", "cites": ["a"]}
    assert source_line_for_slide(deck, slide) == "来源：实验室自测"


def test_source_line_resolves_cites() -> None:
    deck = {
        "citations": {
            "a": {"title": "Evidence interface", "venue": "CVPR", "year": "2024"},
            "b": {"title": "Road risk", "venue": "TRC", "year": "2023"},
        }
    }
    slide = {"cites": ["a", "b"]}
    line = source_line_for_slide(deck, slide)
    assert line.startswith("来源：")
    assert "Evidence interface" in line
    assert "Road risk" in line


def test_assertion_title_heuristic() -> None:
    assert is_assertion_title("研究背景") is False
    assert is_assertion_title("证据接口决定通行判断能否复核") is True
    assert is_assertion_title("所以本课题不再堆叠分割网络。") is True


def test_citation_entry_rejects_bad_doi() -> None:
    issues: list[dict] = []
    validate_citation_entry(
        {"title": "Paper", "venue": "CVPR", "doi": "10.1234/bad doi"},
        "citations",
        issues,
        "Citation 'x'",
    )
    assert any(issue["severity"] == "error" for issue in issues)
    assert any("doi" in issue["message"].lower() for issue in issues)


def test_citation_entry_accepts_valid_doi() -> None:
    issues: list[dict] = []
    validate_citation_entry(
        {"title": "Paper", "venue": "CVPR", "year": "2024", "doi": "10.1109/CVPR.2024.00123"},
        "citations",
        issues,
        "Citation 'x'",
    )
    assert not any(issue["severity"] == "error" for issue in issues)


def test_citation_entry_accepts_prefixed_doi() -> None:
    issues: list[dict] = []
    validate_citation_entry(
        {"title": "Paper", "venue": "Nature", "doi": "DOI: 10.1038/s41586-021-03695-w"},
        "citations",
        issues,
        "Citation 'x'",
    )
    assert not any(issue["severity"] == "error" for issue in issues)


def test_ghost_deck_flags_topic_titles_in_assertion_mode() -> None:
    deck = {
        "meta": {"title": "开题报告", "title_mode": "assertion", "report_type": "开题"},
        "slides": [
            {"id": "01", "layout": "cover", "title": "封面", "speech": "开场。"},
            {
                "id": "03",
                "layout": "grid",
                "title": "研究背景",
                "lead": "真正缺的是证据接口。",
                "speech": "讲解。",
            },
            {
                "id": "04",
                "layout": "grid",
                "title": "证据接口决定通行判断能否复核",
                "lead": "接口缺失会让路径结论不可追溯。",
                "takeaway": "先补接口，再谈模型。",
                "boundary": "不覆盖运动目标。",
                "speech": "讲解。",
            },
        ],
    }
    markdown, report = build(deck)
    assert report["content_slide_count"] == 2
    assert report["assertion_titles"] == 1
    assert report["assertion_ratio"] == 0.5
    assert any("研究背景" in warning for warning in report["warnings"])
    assert "03 研究背景" in markdown
    assert "证据接口决定通行判断能否复核" in markdown
