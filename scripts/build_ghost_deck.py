from __future__ import annotations

import argparse
import json
from pathlib import Path

from common import is_assertion_title, load_json, write_json


READING_CHARS_PER_MINUTE = 320


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a ghost deck: a five-minute read of titles, claims, and boundaries.",
    )
    parser.add_argument("--deck", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--strict", action="store_true", help="Exit non-zero when warnings exist.")
    return parser.parse_args()


def slide_claim(slide: dict) -> str:
    for field in ("lead", "claim", "takeaway"):
        value = str(slide.get(field, "")).strip()
        if value:
            return value
    return ""


def build(deck: dict) -> tuple[str, dict]:
    meta = deck.get("meta", {})
    slides = deck.get("slides", [])
    title_mode = str(meta.get("title_mode", "topic")).strip().lower() or "topic"
    warnings: list[str] = []
    entries: list[dict] = []
    titles: list[str] = []
    seen_titles: set[str] = set()
    assertion_count = 0
    content_count = 0

    for index, slide in enumerate(slides, start=1):
        page_id = str(slide.get("id", f"{index:02d}"))
        title = str(slide.get("title", "")).strip()
        layout = str(slide.get("layout", "")).strip()
        if not title:
            warnings.append(f"{page_id}: missing title")
        if title and title in seen_titles:
            warnings.append(f"{page_id}: duplicate title '{title}'")
        seen_titles.add(title)
        titles.append(title)
        is_content = layout not in {"cover", "outline", "section", "closing", "references"}
        if is_content:
            content_count += 1
            if is_assertion_title(title):
                assertion_count += 1
            elif title_mode == "assertion":
                warnings.append(f"{page_id}: title reads as a topic label, not a claim: '{title}'")
            if not slide_claim(slide):
                warnings.append(f"{page_id}: missing lead/claim")
        entries.append(
            {
                "page": page_id,
                "layout": layout,
                "title": title,
                "claim": slide_claim(slide),
                "takeaway": str(slide.get("takeaway", "")).strip(),
                "boundary": str(slide.get("boundary", "")).strip(),
                "is_assertion": is_assertion_title(title),
            }
        )

    chain = " → ".join(f"{entry['page']} {entry['title']}" for entry in entries if entry["title"])
    read_chars = sum(len(entry["title"]) + len(entry["claim"]) + len(entry["takeaway"]) for entry in entries)
    read_minutes = max(1, round(read_chars / READING_CHARS_PER_MINUTE))
    assertion_ratio = round(assertion_count / content_count, 2) if content_count else 0.0
    report = {
        "deck": str(meta.get("title", "")),
        "title_mode": title_mode,
        "slide_count": len(slides),
        "content_slide_count": content_count,
        "assertion_titles": assertion_count,
        "assertion_ratio": assertion_ratio,
        "estimated_read_minutes": read_minutes,
        "warnings": warnings,
        "chain": chain,
        "slides": entries,
    }

    lines = [
        f"# {meta.get('title', '未命名汇报')}",
        "",
        f"> 幽灵稿：只读标题、主张、结论与边界。预计通读 {read_minutes} 分钟。",
        "",
        f"- 汇报类型：{meta.get('report_type', '—')}",
        f"- 页数：{len(slides)}",
        f"- 标题模式：`{title_mode}`",
        f"- 断言型标题：{assertion_count}/{content_count}（{round(assertion_ratio * 100)}%）",
        "",
        "## 主线",
        "",
        chain or "（无标题）",
        "",
        "## 逐页",
        "",
    ]
    for entry in entries:
        lines.append(f"### {entry['page']} · {entry['title'] or '（无标题）'}")
        lines.append("")
        if entry["claim"]:
            lines.append(f"- 主张：{entry['claim']}")
        if entry["takeaway"]:
            lines.append(f"- 结论：{entry['takeaway']}")
        if entry["boundary"]:
            lines.append(f"- 边界：{entry['boundary']}")
        lines.append("")

    if warnings:
        lines.append("## 需要确认")
        lines.append("")
        lines.extend(f"- {warning}" for warning in warnings)
        lines.append("")

    return "\n".join(lines), report


def main() -> None:
    args = parse_args()
    deck = load_json(Path(args.deck).resolve())
    markdown, report = build(deck)
    output_dir = Path(args.out).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "ghost_deck.md").write_text(markdown, encoding="utf-8")
    write_json(output_dir / "ghost_deck.json", report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.strict and report["warnings"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
