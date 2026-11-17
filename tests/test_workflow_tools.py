from __future__ import annotations

from build_qa import build
from render_deck_next import copy_if_distinct


def test_build_qa_uses_explicit_then_generated() -> None:
    deck = {
        "meta": {"title": "示例开题"},
        "qa": [{"question": "为什么做这个课题？", "answer": "因为证据链缺口。", "page": "03"}],
        "slides": [
            {"id": "01", "layout": "cover", "title": "封面"},
            {
                "id": "03",
                "layout": "grid",
                "title": "研究缺口",
                "claim": "真正缺的是证据接口。",
                "evidence": "五类方法的能力对照。",
                "boundary": "不比较不同论文的绝对精度。",
            },
        ],
    }
    markdown, report = build(deck, count=4)
    assert report["explicit_count"] == 1
    assert report["generated_count"] == 2
    assert report["question_count"] == 3
    assert "为什么做这个课题" in markdown
    assert "直接证据" in markdown
    assert "适用边界" in markdown
    assert "（第 03 页）" in markdown


def test_build_qa_respects_count() -> None:
    deck = {
        "meta": {"title": "示例"},
        "slides": [
            {
                "id": f"{index:02d}",
                "layout": "grid",
                "title": f"页面 {index}",
                "evidence": f"证据 {index}",
                "boundary": f"边界 {index}",
            }
            for index in range(1, 6)
        ],
    }
    _, report = build(deck, count=3)
    assert report["question_count"] == 3


def test_copy_if_distinct_handles_same_source_and_target(work_tmp) -> None:
    deck_path = work_tmp / "deck.json"
    deck_path.write_text('{"slides": []}', encoding="utf-8")
    assert copy_if_distinct(deck_path, deck_path) is False
    assert deck_path.read_text(encoding="utf-8") == '{"slides": []}'

    copied_path = work_tmp / "copied.json"
    assert copy_if_distinct(deck_path, copied_path) is True
    assert copied_path.read_text(encoding="utf-8") == deck_path.read_text(encoding="utf-8")
