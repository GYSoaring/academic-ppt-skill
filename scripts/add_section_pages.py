#!/usr/bin/env python3
"""Insert a section divider before each section's first slide.

The divider look (blue band + number straddling it) is documented in
`references/section-dividers.md`. This script only handles the plumbing: for
every `sections[]` entry, put one `layout: "section"` slide ahead of the first
content slide that belongs to it.

Idempotent: running it twice changes nothing.

Usage:
  python scripts/add_section_pages.py --deck deck.json --out deck.json
  python scripts/add_section_pages.py --deck deck.json --dry-run
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

# 常见章节名 → 英文标签，避免每份 deck 手写
EN_HINTS = {
    "背景": "BACKGROUND",
    "问题": "PROBLEM",
    "原理": "PRINCIPLE",
    "方法": "METHOD",
    "系统": "SYSTEM",
    "信号": "SIGNAL",
    "实验": "EXPERIMENT",
    "结果": "RESULTS",
    "解码": "DECODING",
    "比较": "COMPARISON",
    "讨论": "DISCUSSION",
    "局限": "LIMITATIONS",
    "边界": "LIMITATIONS",
    "空白": "GAPS",
    "结论": "CONCLUSION",
    "展望": "OUTLOOK",
    "应用": "APPLICATION",
    "综述": "REVIEW",
}


def english_label(section: dict) -> str:
    explicit = str(section.get("label_en", "")).strip()
    if explicit:
        return explicit.upper()
    title = str(section.get("title", ""))
    for key, value in EN_HINTS.items():
        if key in title:
            return value
    return "CHAPTER"


def main() -> None:
    parser = argparse.ArgumentParser(description="Insert section divider pages.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--out", default=None, help="默认覆盖原文件")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    deck_path = Path(args.deck)
    deck = json.loads(deck_path.read_text(encoding="utf-8"))
    sections = deck.get("sections") or []
    if not sections:
        raise SystemExit("deck has no sections[] — nothing to divide")

    order = [str(section.get("id")) for section in sections]
    titles = {str(section.get("id")): section for section in sections}
    existing = {
        str(slide.get("section"))
        for slide in deck.get("slides", [])
        if str(slide.get("layout", "")).strip().lower() == "section"
    }

    rebuilt = []
    placed: set[str] = set()
    inserted = 0
    for slide in deck.get("slides", []):
        layout = str(slide.get("layout", "")).strip().lower()
        section_id = str(slide.get("section", ""))
        if (
            layout not in {"cover", "outline", "section"}
            and section_id in order
            and section_id not in placed
            and section_id not in existing
        ):
            position = order.index(section_id) + 1
            meta = titles[section_id]
            rebuilt.append(
                {
                    "id": f"S{position}",
                    "layout": "section",
                    "section": section_id,
                    "eyebrow": f"{position:02d}",
                    "title": str(meta.get("title", "")),
                    "label_en": english_label(meta),
                    "summary": str(meta.get("desc", "")),
                    "claim": f"本部分讨论{meta.get('title', '')}。",
                    "evidence": "见后续内容页。",
                    "boundary": "仅覆盖方法学层面，不含具体装置参数。",
                    "speech": f"进入第 {position} 部分：{meta.get('title', '')}。",
                }
            )
            inserted += 1
        if section_id:
            placed.add(section_id)
        rebuilt.append(slide)

    deck["slides"] = rebuilt
    print(json.dumps({"inserted": inserted, "pages": len(rebuilt)}, ensure_ascii=False))
    if args.dry_run:
        return
    target = Path(args.out) if args.out else deck_path
    target.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
