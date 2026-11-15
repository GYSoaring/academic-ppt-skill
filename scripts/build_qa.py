from __future__ import annotations

import argparse
import json
from pathlib import Path

from common import load_json, write_json


SKIP_LAYOUTS = {"cover", "outline", "section", "references"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a traceable Q&A sheet from slide claims, evidence, and boundaries.",
    )
    parser.add_argument("--deck", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--count", type=int, default=10)
    return parser.parse_args()


def explicit_questions(deck: dict) -> list[dict]:
    entries = deck.get("qa") or []
    result: list[dict] = []
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        question = str(entry.get("question", "")).strip()
        answer = str(entry.get("answer", "")).strip()
        if question and answer:
            result.append(
                {
                    "question": question,
                    "answer": answer,
                    "page": str(entry.get("page", "")),
                    "origin": "deck",
                }
            )
    return result


def generated_questions(slides: list[dict], limit: int) -> list[dict]:
    candidates: list[dict] = []
    for index, slide in enumerate(slides, start=1):
        layout = slide.get("layout")
        if layout in SKIP_LAYOUTS:
            continue
        page = str(slide.get("id", f"{index:02d}"))
        title = str(slide.get("title", "")).strip()
        claim = str(slide.get("claim") or slide.get("lead") or "").strip()
        evidence = str(slide.get("evidence", "")).strip()
        boundary = str(slide.get("boundary", "")).strip()
        if evidence:
            candidates.append(
                {
                    "question": f"第 {page} 页「{title}」的直接证据是什么？",
                    "answer": evidence,
                    "page": page,
                    "origin": "generated",
                }
            )
        if boundary:
            candidates.append(
                {
                    "question": f"「{title}」这一结论的适用边界在哪里？",
                    "answer": boundary,
                    "page": page,
                    "origin": "generated",
                }
            )
        if claim and not evidence:
            candidates.append(
                {
                    "question": f"请用一句话概括第 {page} 页的核心判断。",
                    "answer": claim,
                    "page": page,
                    "origin": "generated",
                }
            )
    return candidates[:limit]


def build(deck: dict, count: int = 10) -> tuple[str, dict]:
    explicit = explicit_questions(deck)
    remaining = max(0, count - len(explicit))
    generated = generated_questions(deck.get("slides", []), remaining)
    questions = explicit + generated
    report = {
        "deck": str(deck.get("meta", {}).get("title", "")),
        "question_count": len(questions),
        "explicit_count": len(explicit),
        "generated_count": len(generated),
        "questions": questions,
    }

    lines = [
        f"# 答辩 Q&A 预测：{deck.get('meta', {}).get('title', '未命名汇报')}",
        "",
        "> 每个问题都指向具体页面。回答只使用该页已有的 claim、evidence 和 boundary，不补造数据。",
        "",
    ]
    for index, item in enumerate(questions, start=1):
        page = f"（第 {item['page']} 页）" if item.get("page") else ""
        lines.append(f"{index}. **{item['question']}**{page}")
        lines.append(f"   - 回答：{item['answer']}")
        lines.append("")

    if not questions:
        lines.append("当前 deck 没有可生成的问题；请补充 claim、evidence 或 boundary。")
        lines.append("")

    return "\n".join(lines), report


def main() -> None:
    args = parse_args()
    deck = load_json(Path(args.deck).resolve())
    markdown, report = build(deck, count=max(1, args.count))
    output_dir = Path(args.out).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "qa.md").write_text(markdown, encoding="utf-8")
    write_json(output_dir / "qa.json", report)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
