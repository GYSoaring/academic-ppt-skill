from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from common import estimate_seconds, load_json, write_json


PATTERNS = [
    ("A1", "A", "翻案腔", r"(?:不是|并非|并不是|不是要)[^。！？；\n]{1,50}(?:而是|而是要|关键在于)", "直接给正面判断；如果反面信息必要，另起一句。"),
    ("A2", "A", "翻案腔", r"与其说[^。！？；\n]{1,40}不如说", "直接说结论，不先立一个假对立。"),
    ("A3", "A", "翻案腔", r"看似[^。！？；\n]{1,30}(?:实则|其实)", "直接说实际结论，保留必要限定。"),
    ("B1", "B", "提示语冒号", r"(?:一句话总结|一句话概括|核心是|关键在于|原因如下|本质是|结论是)[：:]", "删掉提示语，直接说后面的内容。"),
    ("B2", "B", "起手式", r"(?:说白了|说穿了|先说结论|直接说结论|简单来说)[，,]", "删掉起手式，直接进入判断。"),
    ("C1", "C", "翻译腔", r"当[^。！？；\n]{1,30}时[，,]", "删掉“当……时”，让从句直接作前一分句。"),
    ("C2", "C", "翻译腔", r"(?:对于|对)[^。！？；\n]{1,24}(?:来说|而言)[，,]", "删掉“对于……来说”，让对象直接作主语。"),
    ("C3", "C", "翻译腔", r"在[^。！？；\n]{1,24}方面[，,]?", "删掉“在……方面”的话题壳。"),
    ("C4", "C", "句首路标", r"(?:^|[。！？]\s*)(?:然而|因此|此外|同时|综上所述|总而言之)[，,]", "把连接词移到主语后，或直接删掉。"),
]

X_LEVEL_EXAMPLES = [
    "首先，其次，最后",
    "这意味着",
    "一句话引出列表",
    "序号小标题",
    "精确数字",
    "空转句引列表",
    "概括替代",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Audit defense speech for AI-tone patterns and spoken rhythm.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--min-sentence", type=int, default=70)
    parser.add_argument("--apply-safe", action="store_true", help="Apply only punctuation/discourse-marker cleanup.")
    parser.add_argument("--out-deck", default=None)
    return parser.parse_args()


def split_sentences(text: str) -> list[str]:
    return [part.strip() for part in re.split(r"(?<=[。！？])", text or "") if part.strip()]


def find_hits(text: str, min_sentence: int) -> list[dict]:
    hits: list[dict] = []
    seen: set[tuple[str, int, str]] = set()
    for pattern_id, level, label, pattern, suggestion in PATTERNS:
        for match in re.finditer(pattern, text):
            snippet = match.group(0).strip()
            key = (pattern_id, match.start(), snippet)
            if key in seen:
                continue
            seen.add(key)
            hits.append(
                {
                    "id": pattern_id,
                    "level": level,
                    "label": label,
                    "text": snippet,
                    "suggestion": suggestion,
                }
            )
    for sentence in split_sentences(text):
        separators = sentence.count("、")
        if separators >= 3 or (separators == 2 and len(sentence) > 60):
            hits.append(
                {
                    "id": "D1",
                    "level": "C",
                    "label": "密集并列",
                    "text": sentence,
                    "suggestion": "这句有四项以上并列；口播里分组或概括，不要一口气念完。",
                }
            )
        if len(sentence) > min_sentence:
            hits.append(
                {
                    "id": "R1",
                    "level": "C",
                    "label": "长句",
                    "text": sentence,
                    "suggestion": f"这句 {len(sentence)} 字，超过 {min_sentence} 字；拆成两口气，别在句中换气。",
                }
            )
    semicolons = text.count("；")
    if semicolons:
        hits.append(
            {
                "id": "F1",
                "level": "F",
                "label": "分号",
                "text": f"本页 {semicolons} 个分号",
                "suggestion": "口播里把分号拆成句号，给听众换气点。",
            }
        )
    return hits


def slide_bullets(slide: dict) -> list[str]:
    values: list[str] = []
    columns = slide.get("columns")
    if isinstance(columns, list):
        for column in columns:
            if isinstance(column, dict):
                values.extend(str(item) for item in column.get("bullets") or [])
    bullets = slide.get("bullets")
    if isinstance(bullets, list):
        values.extend(str(item) for item in bullets)
    for block in slide.get("blocks") or []:
        values.extend(str(item) for item in block.get("bullets") or [])
    for panel in slide.get("panels") or []:
        values.extend(str(item) for item in panel.get("bullets") or [])
    return [value.strip() for value in values if len(value.strip()) >= 8]


def audit_deck(deck: dict, min_sentence: int = 70) -> dict:
    slides = deck.get("slides", [])
    records: list[dict] = []
    level_counts = {"A": 0, "B": 0, "C": 0, "F": 0}
    total_chars = 0
    total_seconds = 0
    for index, slide in enumerate(slides, start=1):
        speech = str(slide.get("speech", "")).strip()
        sentences = split_sentences(speech)
        hits = find_hits(speech, min_sentence)
        echoed = [bullet for bullet in slide_bullets(slide) if bullet in speech]
        if echoed:
            hits.append(
                {
                    "id": "R2",
                    "level": "C",
                    "label": "念 PPT",
                    "text": "；".join(echoed[:3]),
                    "suggestion": "讲稿里出现了 slide 上的原句；改成自己的话，先说这页在回答什么。",
                }
            )
        for hit in hits:
            level_counts[hit["level"]] = level_counts.get(hit["level"], 0) + 1
        seconds = estimate_seconds(speech)
        total_chars += len(speech)
        total_seconds += seconds
        records.append(
            {
                "page": index,
                "title": slide.get("title", ""),
                "chars": len(speech),
                "sentences": len(sentences),
                "avg_sentence_chars": round(len(speech) / max(1, len(sentences)), 1),
                "max_sentence_chars": max((len(item) for item in sentences), default=0),
                "estimated_seconds": seconds,
                "hits": hits,
            }
        )
    return {
        "deck": deck.get("meta", {}).get("title", ""),
        "slide_count": len(slides),
        "chars": total_chars,
        "estimated_seconds": total_seconds,
        "level_counts": level_counts,
        "records": records,
        "x_level_not_changed": X_LEVEL_EXAMPLES,
        "method": "Adapted from lieflat-less-ai-tone: script-detectable patterns plus spoken-rhythm checks. No authorship claim.",
    }


def render_report(report: dict) -> str:
    counts = report["level_counts"]
    lines = [
        f"# 讲解稿去 AI 味审计：{report['deck']}",
        "",
        "> 方法：基于 `lieflat-less-ai-tone` 的分级检测，适配答辩口播。只给改法，不改事实、结构、限定词和引用。",
        "",
        "## 总览",
        "",
        f"- 页数：{report['slide_count']}",
        f"- 讲稿字数：{report['chars']}",
        f"- 估算时长：{report['estimated_seconds']} 秒",
        f"- 命中：A 级 {counts.get('A', 0)}，B 级 {counts.get('B', 0)}，C 级 {counts.get('C', 0)}，F 级 {counts.get('F', 0)}",
        "",
        "## 分级命中",
        "",
        "| 页 | 级别 | 类型 | 原文 | 改法 |",
        "|---|---|---|---|---|",
    ]
    any_hits = False
    for record in report["records"]:
        for hit in record["hits"]:
            any_hits = True
            text = hit["text"].replace("|", "\\|").replace("\n", " ")
            lines.append(
                f"| {record['page']:02d} | {hit['level']} | {hit['label']} | {text} | {hit['suggestion']} |"
            )
    if not any_hits:
        lines.append("| — | — | — | 未发现 A/B/C/F 命中 | 仍需人工检查比喻贴切度和现场语气 |")
    lines.extend(["", "## 逐页节奏", ""])
    for record in report["records"]:
        lines.append(
            f"- 第 {record['page']:02d} 页「{record['title']}」：{record['chars']} 字，"
            f"{record['sentences']} 句，平均 {record['avg_sentence_chars']} 字/句，"
            f"最长 {record['max_sentence_chars']} 字，约 {record['estimated_seconds']} 秒。"
        )
        if record["hits"]:
            lines.append(f"  - 本页命中 {len(record['hits'])} 项；优先处理 A/B 级，再处理长句和分号。")
    lines.extend(
        [
            "",
            "## X 级不改",
            "",
            "以下项目在适配研究中没有稳定的人机差异，或人类反而更常用；本报告不把它们当错误：",
            "",
        ]
    )
    lines.extend(f"- {item}" for item in report["x_level_not_changed"])
    lines.extend(
        [
            "",
            "## 改写纪律",
            "",
            "- 最小改动：一处命中只改解决该问题所需的部分。",
            "- 信息守恒：观点、事实、数字、限定词、引文不改。",
            "- 结构不动：页面顺序、标题层级、列表和表格不改。",
            "- 不制造参差：不为了“像人写的”故意把句子改得忽长忽短。",
            "- 不判断作者身份：报告只描述文本特征，不推断由哪个模型写成。",
            "",
        ]
    )
    return "\n".join(lines)


def apply_safe(deck: dict) -> tuple[dict, list[dict]]:
    replacements = {
        "值得注意的是，": "",
        "需要指出的是，": "",
        "简而言之，": "",
        "综上所述，": "所以，",
        "总而言之，": "所以，",
    }
    changes: list[dict] = []
    for index, slide in enumerate(deck.get("slides", []), start=1):
        speech = str(slide.get("speech", ""))
        updated = speech
        for source, target in replacements.items():
            if source in updated:
                updated = updated.replace(source, target)
                changes.append({"page": index, "kind": "discourse_marker", "from": source, "to": target})
        if "；" in updated:
            updated = updated.replace("；", "。")
            changes.append({"page": index, "kind": "semicolon_to_period", "from": "；", "to": "。"})
        if updated != speech:
            slide["speech"] = updated
    return deck, changes


def main() -> None:
    args = parse_args()
    deck_path = Path(args.deck).resolve()
    output_dir = Path(args.out).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    deck = load_json(deck_path)
    report = audit_deck(deck, min_sentence=max(30, args.min_sentence))
    (output_dir / "speech_audit.md").write_text(render_report(report), encoding="utf-8")
    write_json(output_dir / "speech_audit.json", report)
    if args.apply_safe:
        polished, changes = apply_safe(deck)
        out_deck = Path(args.out_deck).resolve() if args.out_deck else deck_path.with_name(deck_path.stem + ".polished.json")
        write_json(out_deck, polished)
        write_json(output_dir / "speech_safe_changes.json", {"changes": changes, "out_deck": str(out_deck)})
        after = audit_deck(polished, min_sentence=max(30, args.min_sentence))
        (output_dir / "speech_audit_after_safe.md").write_text(render_report(after), encoding="utf-8")
    print(json.dumps(report["level_counts"] | {"pages": report["slide_count"]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
