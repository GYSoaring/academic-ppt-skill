from __future__ import annotations

import argparse
import json
from pathlib import Path

from common import estimate_seconds, format_duration, load_json, write_json


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build speaker notes and a continuous defense script.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--allow-missing", action="store_true")
    return parser.parse_args()


def slide_label(index: int, slide: dict) -> str:
    return f"第{index:02d}页｜{slide.get('title', '未命名')}"


def main() -> None:
    args = parse_args()
    deck_path = Path(args.deck).resolve()
    output_dir = Path(args.out).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    deck = load_json(deck_path)
    meta = deck.get("meta", {})
    slides = deck.get("slides", [])
    if not slides:
        raise ValueError("deck.json has no slides")

    missing = [index for index, slide in enumerate(slides, start=1) if not str(slide.get("speech", "")).strip()]
    if missing and not args.allow_missing:
        raise ValueError(f"Missing speech on slides: {missing}")

    note_lines = [
        f"# {meta.get('title', '答辩汇报')}：逐页讲解词",
        "",
        f"- 汇报人：{meta.get('presenter', '')}",
        f"- 指导教师：{meta.get('advisor', '')}",
        f"- 汇报时间：{meta.get('date', '')}",
        f"- 目标时长：{meta.get('duration_min', 20)} 分钟",
        "",
        "> 每页时长按中文口播速度估算，实际讲稿应结合停顿、翻页和现场节奏微调。",
        "",
    ]
    script_lines = [
        f"# {meta.get('title', '答辩汇报')}：临场答辩稿",
        "",
        f"> 汇报人：{meta.get('presenter', '')}　指导教师：{meta.get('advisor', '')}　日期：{meta.get('date', '')}",
        "",
    ]
    records = []
    total_seconds = 0
    for index, slide in enumerate(slides, start=1):
        speech = str(slide.get("speech", "")).strip()
        seconds = estimate_seconds(speech)
        total_seconds += seconds
        label = slide_label(index, slide)
        records.append(
            {
                "page": index,
                "title": slide.get("title", ""),
                "layout": slide.get("layout", ""),
                "estimated_seconds": seconds,
                "speech_chars": len(speech),
            }
        )
        note_lines.extend(
            [
                f"## {label}",
                "",
                f"- 版式：`{slide.get('layout', '')}`",
                f"- 预计：{format_duration(seconds)}",
                f"- 本页任务：{slide.get('claim', '') or slide.get('lead', '')}",
                f"- 边界：{slide.get('boundary', '') or '未单独填写'}",
                "",
                speech or "【本页讲解词待补充】",
                "",
                "---",
                "",
            ]
        )
        script_lines.extend(
            [
                f"## {label}",
                "",
                speech or "【本页讲解词待补充】",
                "",
                "---",
                "",
            ]
        )

    qa = deck.get("qa", [])
    if qa:
        note_lines.extend(["## 预设问答", ""])
        script_lines.extend(["## 预设问答", ""])
        for item in qa:
            question = str(item.get("question", "")).strip()
            answer = str(item.get("answer", "")).strip()
            note_lines.extend([f"**问：{question}**", "", answer, ""])
            script_lines.extend([f"**问：{question}**", "", answer, ""])

    target_seconds = int(float(meta.get("duration_min", 20)) * 60)
    ratio = total_seconds / target_seconds if target_seconds else 0
    script_lines.insert(
        4,
        f"> 预计正文时长：{format_duration(total_seconds)}；目标时长：{format_duration(target_seconds)}。",
    )
    note_lines.extend(
        [
            "## 时长检查",
            "",
            f"- 目标：{format_duration(target_seconds)}",
            f"- 估算：{format_duration(total_seconds)}",
            f"- 比例：{ratio:.0%}",
            "",
        ]
    )
    (output_dir / "speaker_notes.md").write_text("\n".join(note_lines), encoding="utf-8")
    (output_dir / "defense_script.md").write_text("\n".join(script_lines), encoding="utf-8")
    report = {
        "target_seconds": target_seconds,
        "estimated_seconds": total_seconds,
        "ratio": ratio,
        "missing_speech": missing,
        "slides": records,
    }
    write_json(output_dir / "speech_report.json", report)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
