from __future__ import annotations

import json

from audit_speech import apply_safe, audit_deck, render_report


def make_deck() -> dict:
    return {
        "meta": {"title": "测试答辩"},
        "slides": [
            {
                "id": "01",
                "title": "问题",
                "speech": "这不是效率问题，而是相位问题。一句话总结：先算周期；再测效率。",
            },
            {
                "id": "02",
                "title": "边界",
                "speech": "当温度变化时，对于早期样品来说，在稳定性方面还需要观察。首先，我们看背景。这意味着周期会漂移。",
            },
            {
                "id": "03",
                "title": "长句",
                "speech": "这一段故意写得很长用来测试口播节奏因为一句话里塞进了太多限定条件、太多并列项和太多需要听众在脑子里同时保持的信息所以它应该被拆成两口气。",
            },
            {
                "id": "04",
                "title": "念PPT",
                "bullets": ["周期均匀性决定旁瓣位置"],
                "speech": "周期均匀性决定旁瓣位置。",
            },
        ],
    }


def test_audit_detects_graded_hits() -> None:
    report = audit_deck(make_deck(), min_sentence=50)
    ids = {hit["id"] for record in report["records"] for hit in record["hits"]}
    assert "A1" in ids
    assert "B1" in ids
    assert "C1" in ids
    assert "C2" in ids
    assert "C3" in ids
    assert "F1" in ids
    assert "R1" in ids
    assert "R2" in ids
    assert report["level_counts"]["A"] >= 1
    assert report["level_counts"]["B"] >= 1


def test_x_level_examples_are_not_flagged() -> None:
    report = audit_deck(make_deck(), min_sentence=50)
    labels = {hit["label"] for record in report["records"] for hit in record["hits"]}
    texts = {hit["text"] for record in report["records"] for hit in record["hits"]}
    assert "首先" not in labels
    assert "这意味着" not in labels
    assert not any("首先" in text or "这意味着" in text for text in texts)


def test_apply_safe_only_removes_safe_markers_and_semicolons() -> None:
    deck = make_deck()
    polished, changes = apply_safe(deck)
    assert "一句话总结：" in polished["slides"][0]["speech"]
    assert "；" not in polished["slides"][0]["speech"]
    assert any(change["kind"] == "semicolon_to_period" for change in changes)
    assert "这不是效率问题，而是相位问题。" in polished["slides"][0]["speech"]


def test_report_contains_graded_table_and_x_list() -> None:
    report = audit_deck(make_deck(), min_sentence=50)
    markdown = render_report(report)
    assert "分级命中" in markdown
    assert "X 级不改" in markdown
    assert "这意味着" in markdown
