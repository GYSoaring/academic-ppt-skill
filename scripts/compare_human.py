#!/usr/bin/env python3
"""Benchmark a generated deck against decks a human actually designed.

The skill can grade its own structure (`qa_design.py`), but that only measures
whether a deck follows the skill's own rules. This script answers a different
question: does the output look like something a human designer made?

It extracts a geometric fingerprint from two sets of PPTX files — a human
reference set and the generated deck — and compares them on properties that are
visible in the file format and do not depend on taste:

- `title_pt_per_in` / `body_pt_per_in`  type scale normalised by slide width
- `font_steps`                          how many distinct sizes the deck uses
- `pictures_per_slide`                  real imagery per page
- `chars_per_slide`                     how much text a page carries
- `shapes_per_slide`                    visual complexity
- `text_area_ratio`                     how much of the page text boxes claim
- `margin_left_ratio`                   how far the text block sits from the edge
- `overlap_ratio`                       boxes that collide with each other

Normalising by slide width matters: a 10 in template at 14 pt and a 13.3 in
deck at 18.75 pt are typographically the same size class.
"""

from __future__ import annotations

import argparse
import json
import statistics
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

EMU_PER_INCH = 914400
PICTURE_TYPE = 13


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compare a deck against human-made reference decks.")
    parser.add_argument("--human", nargs="+", required=True, help="Human-designed .pptx files or folders.")
    parser.add_argument("--deck", required=True, help="The generated .pptx to benchmark.")
    parser.add_argument("--out", default=None)
    parser.add_argument("--json", action="store_true", help="Print JSON instead of a table.")
    return parser.parse_args()


def expand(paths: list[str]) -> list[Path]:
    files: list[Path] = []
    for raw in paths:
        path = Path(raw)
        if path.is_dir():
            files.extend(sorted(path.rglob("*.pptx")))
        elif path.suffix.lower() == ".pptx":
            files.append(path)
    return [path for path in files if path.exists()]


def overlap_ratio(boxes: list[tuple[int, int, int, int]]) -> float:
    pairs = 0
    hits = 0
    for index, (left, top, width, height) in enumerate(boxes):
        for other_left, other_top, other_width, other_height in boxes[index + 1 :]:
            pairs += 1
            if (
                left < other_left + other_width
                and other_left < left + width
                and top < other_top + other_height
                and other_top < top + height
            ):
                hits += 1
    return hits / pairs if pairs else 0.0


def fingerprint(path: Path, skip_first: int = 0) -> dict:
    prs = Presentation(str(path))
    width = prs.slide_width or Emu(1)
    height = prs.slide_height or Emu(1)
    width_in = float(width) / EMU_PER_INCH
    area = float(width) * float(height)
    slides = list(prs.slides)[skip_first:]
    if not slides:
        slides = list(prs.slides)

    per_slide = []
    deck_sizes: set[float] = set()
    for slide in slides:
        sizes: list[float] = []
        chars = 0
        pictures = 0
        boxes: list[tuple[int, int, int, int]] = []
        text_boxes: list[tuple[int, int, int, int]] = []
        text_area = 0.0
        for shape in slide.shapes:
            try:
                box = (int(shape.left or 0), int(shape.top or 0), int(shape.width or 0), int(shape.height or 0))
            except Exception:
                box = None
            if box is not None:
                if area and float(box[2]) * float(box[3]) / area >= 0.92:
                    # full-bleed background: not part of the layout, skip it everywhere
                    continue
                boxes.append(box)
            if shape.shape_type == PICTURE_TYPE or shape.__class__.__name__ == "Picture":
                pictures += 1
            if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
                chars += len(shape.text_frame.text.strip())
                if box is not None:
                    text_boxes.append(box)
                try:
                    text_area += float((shape.width or 0) * (shape.height or 0))
                except Exception:
                    pass
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.font.size:
                            sizes.append(round(run.font.size.pt, 1))
        deck_sizes.update(sizes)
        body_sizes = sorted(sizes)
        body_pt = statistics.median(body_sizes) if body_sizes else 0.0
        left = min((box[0] for box in text_boxes), default=0)
        per_slide.append(
            {
                "shapes": len(boxes),
                "pictures": pictures,
                "chars": chars,
                "sizes": sizes,
                "max_pt": max(sizes) if sizes else 0.0,
                "min_pt": min(sizes) if sizes else 0.0,
                "body_pt": body_pt,
                "font_steps": len(set(sizes)),
                "text_area_ratio": min(1.0, text_area / area) if area else 0.0,
                "margin_left_ratio": left / float(width) if width else 0.0,
                "overlap_ratio": overlap_ratio(boxes),
            }
        )
    return {"file": path.name, "width_in": round(width_in, 2), "slides": per_slide, "deck_sizes": deck_sizes}


def summarise(fingerprints: list[dict]) -> dict:
    slides = [slide for fp in fingerprints for slide in fp["slides"]]
    sizes = set()
    for fp in fingerprints:
        sizes.update(fp["deck_sizes"])
    max_pts = [slide["max_pt"] for slide in slides if slide["max_pt"]]
    body_pts = [slide["body_pt"] for slide in slides if slide["body_pt"]]
    per_deck_steps = [len(fp["deck_sizes"]) for fp in fingerprints if fp["deck_sizes"]]
    widths = [fp["width_in"] for fp in fingerprints if fp["width_in"]]
    width_in = statistics.median(widths) if widths else 13.333
    return {
        "decks": len(fingerprints),
        "slides": len(slides),
        "width_in": round(width_in, 2),
        "font_steps": round(statistics.median(per_deck_steps), 1) if per_deck_steps else 0,
        "title_pt_per_in": round(statistics.median(max_pts) / width_in, 2) if max_pts else 0.0,
        "body_pt_per_in": round(statistics.median(body_pts) / width_in, 2) if body_pts else 0.0,
        "body_pt": round(statistics.median(body_pts), 1) if body_pts else 0.0,
        "pictures_per_slide": round(statistics.mean(slide["pictures"] for slide in slides), 2),
        "shapes_per_slide": round(statistics.mean(slide["shapes"] for slide in slides), 1),
        "chars_per_slide": round(statistics.mean(slide["chars"] for slide in slides), 0),
        "text_area_ratio": round(statistics.mean(slide["text_area_ratio"] for slide in slides), 3),
        "margin_left_ratio": round(statistics.mean(slide["margin_left_ratio"] for slide in slides), 3),
        "overlap_ratio": round(statistics.mean(slide["overlap_ratio"] for slide in slides), 3),
    }


ROWS = [
    ("title_pt_per_in", "标题字号 / 版面宽（in）", 0.15),
    ("body_pt_per_in", "正文字号 / 版面宽（in）", 0.15),
    ("font_steps", "全篇字号档数", 2),
    ("pictures_per_slide", "每页真图数", 0.6),
    ("shapes_per_slide", "每页元素数", 6.0),
    ("chars_per_slide", "每页字数", 80),
    ("text_area_ratio", "文本框占版面比", 0.10),
    ("margin_left_ratio", "左边距占版面比", 0.03),
]


def main() -> None:
    args = parse_args()
    human_files = expand(args.human)
    if not human_files:
        raise SystemExit("No human decks found.")
    human = summarise([fingerprint(path) for path in human_files])
    ours = summarise([fingerprint(Path(args.deck))])

    rows = []
    for key, label, tolerance in ROWS:
        human_value = human.get(key)
        our_value = ours.get(key)
        if isinstance(human_value, float) and isinstance(our_value, float):
            delta = round(our_value - human_value, 3)
            ok = abs(delta) <= tolerance
        else:
            delta = our_value - human_value
            ok = abs(delta) <= tolerance
        rows.append({"metric": key, "label": label, "human": human_value, "ours": our_value, "delta": delta, "ok": bool(ok)})

    passed = sum(1 for row in rows if row["ok"])
    report = {
        "human": human,
        "generated": ours,
        "human_files": [path.name for path in human_files],
        "rows": rows,
        "passed": passed,
        "total": len(rows),
        "verdict": "和真人设计同一量级" if passed >= len(rows) - 1 else ("接近，但仍有明显偏差" if passed >= len(rows) - 3 else "和真人设计不在同一量级"),
    }
    if args.out:
        out = Path(args.out)
        out.mkdir(parents=True, exist_ok=True)
        (out / "human_benchmark.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return
    print(f"真人基准：{human['decks']} 份 deck / {human['slides']} 页，版面宽 {human['width_in']} in")
    print(f"待评 deck：{ours['slides']} 页，版面宽 {ours['width_in']} in")
    print("")
    print(f"{'指标':<28}{'真人':>10}{'本次':>10}{'差':>9}  结论")
    for row in rows:
        print(f"{row['label']:<28}{row['human']:>10}{row['ours']:>10}{row['delta']:>9}  {'✅' if row['ok'] else '⚠️'}")
    print("")
    print(f"{report['passed']}/{report['total']} 项达标 —— {report['verdict']}")


if __name__ == "__main__":
    main()
