#!/usr/bin/env python3
"""Measure a human-made deck so the generator can stop guessing.

`compare_human.py` answers "is my deck close to this one". This answers the
prior question: **what does this deck actually do** — which fonts, which size
ladder, which palette, how much chrome, how the title block sits.

Run it on a handful of decks that you consider good, aggregate the numbers, and
write the result into the theme. Guessing a house style from a screenshot is how
decks end up looking AI-generated; measuring is how they stop.

Usage:
  python scripts/analyze_reference.py deck.pptx --out profile.json
  python scripts/analyze_reference.py a.pptx b.pptx --out profiles/
  python scripts/analyze_reference.py a.pptx b.pptx --aggregate --out profile.md
"""

from __future__ import annotations

import argparse
import json
import statistics
from collections import Counter
from pathlib import Path

EMU_PER_INCH = 914400
FULL_BLEED_RATIO = 0.92


def colour_of(fill) -> tuple[int, int, int] | None:
    try:
        if fill.type is None:
            return None
        rgb = fill.fore_color.rgb
        return (rgb[0], rgb[1], rgb[2])
    except Exception:
        return None


def analyse(path: Path) -> dict:
    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation(str(path))
    width_in = prs.slide_width / EMU_PER_INCH
    height_in = prs.slide_height / EMU_PER_INCH
    area = width_in * height_in

    fonts: Counter = Counter()
    sizes: Counter = Counter()
    colours: Counter = Counter()
    bold_runs = 0
    text_runs = 0
    chars = 0
    objects = 0
    pictures = 0
    tables = 0
    charts = 0
    chrome = 0
    title_sizes: list[float] = []
    title_tops: list[float] = []
    text_boxes = 0
    text_area = 0.0
    left_edges: list[float] = []

    for slide in prs.slides:
        page_objects = 0
        biggest_size = 0.0
        biggest_top = None
        for shape in slide.shapes:
            page_objects += 1
            objects += 1
            if shape.shape_type == 13 or getattr(shape, "image", None) is not None:
                pictures += 1
            if getattr(shape, "has_table", False) and shape.has_table:
                tables += 1
            if getattr(shape, "has_chart", False) and shape.has_chart:
                charts += 1

            shape_area = 0.0
            try:
                shape_area = (shape.width / EMU_PER_INCH) * (shape.height / EMU_PER_INCH)
            except Exception:
                pass
            is_full_bleed = area and shape_area / area >= FULL_BLEED_RATIO

            has_text = getattr(shape, "has_text_frame", False) and shape.has_text_frame
            text = shape.text_frame.text.strip() if has_text else ""
            if has_text and text:
                text_boxes += 1
                if not is_full_bleed:
                    text_area += shape_area
                    try:
                        left_edges.append(round(shape.left / EMU_PER_INCH, 2))
                    except Exception:
                        pass
            elif not text and not is_full_bleed:
                chrome += 1

            if not has_text:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if not run.text.strip():
                        continue
                    text_runs += 1
                    chars += len(run.text.strip())
                    if run.font.bold:
                        bold_runs += 1
                    if run.font.name:
                        fonts[run.font.name] += 1
                    if run.font.size:
                        pt = round(run.font.size.pt, 1)
                        sizes[pt] += 1
                        if pt > biggest_size:
                            biggest_size = pt
                            biggest_top = shape.top
                    try:
                        if run.font.color and run.font.color.rgb:
                            colours[str(run.font.color.rgb)] += 1
                    except Exception:
                        pass
            colour = colour_of(shape.fill) if getattr(shape, "fill", None) else None
            if colour:
                colours["#%02X%02X%02X" % colour] += 1

        if biggest_size:
            title_sizes.append(biggest_size)
        if biggest_top is not None:
            title_tops.append(round(Emu(int(biggest_top)).inches, 2))

    pages = len(prs.slides) or 1
    margins = [edge for edge in left_edges if 0 < edge < width_in * 0.25]
    return {
        "file": path.name,
        "pages": len(prs.slides),
        "canvas_in": [round(width_in, 2), round(height_in, 2)],
        "objects_per_page": round(objects / pages, 1),
        "pictures_per_page": round(pictures / pages, 2),
        "tables": tables,
        "charts": charts,
        "chars_per_page": round(chars / pages),
        "text_boxes_per_page": round(text_boxes / pages, 1),
        "text_area_ratio": round(text_area / (area * pages), 3),
        "left_margin_in": round(statistics.median(margins), 2) if margins else None,
        "left_margin_ratio": round(statistics.median(margins) / width_in, 3) if margins else None,
        "bold_ratio": round(bold_runs / text_runs, 2) if text_runs else 0,
        "chrome_share": round(chrome / (objects or 1), 2),
        "title_pt_per_inch": round(statistics.median(title_sizes) / width_in, 2) if title_sizes else None,
        "title_top_in": round(statistics.median(title_tops), 2) if title_tops else None,
        "size_steps": len(sizes),
        "size_histogram": dict(sizes.most_common(14)),
        "body_pt": round(
            statistics.median([pt for pt, count in sizes.items() for _ in range(count) if pt < 24]) if sizes else 0, 1
        ),
        "fonts": [name for name, _ in fonts.most_common(6)],
        "font_share": {name: round(count / text_runs, 2) for name, count in fonts.most_common(6)},
        "palette": [colour for colour, _ in colours.most_common(8)],
    }


def aggregate(profiles: list[dict]) -> str:
    pages = sum(profile["pages"] for profile in profiles)

    def median(key: str):
        values = [profile[key] for profile in profiles if isinstance(profile.get(key), (int, float))]
        return round(statistics.median(values), 2) if values else None

    def weighted(key: str):
        total = 0.0
        for profile in profiles:
            value = profile.get(key)
            if isinstance(value, (int, float)):
                total += value * profile["pages"]
        return round(total / pages, 2) if pages else None

    font_votes: Counter = Counter()
    for profile in profiles:
        for name in profile["fonts"][:4]:
            font_votes[name] += 1

    lines = [
        "# 参考 deck 实测画像",
        "",
        f"样本：{len(profiles)} 份，共 {pages} 页，全部 "
        f"{profiles[0]['canvas_in'][0]:.2f}×{profiles[0]['canvas_in'][1]:.2f} in。",
        "",
        "| 指标 | 中位数 | 按页加权 |",
        "|---|---|---|",
    ]
    for key, label in [
        ("objects_per_page", "每页元素数"),
        ("pictures_per_page", "每页图片数"),
        ("chars_per_page", "每页字数"),
        ("text_boxes_per_page", "每页文本框"),
        ("text_area_ratio", "文本框占版面比"),
        ("left_margin_ratio", "左边距占版面比"),
        ("bold_ratio", "加粗占比"),
        ("chrome_share", "无文字装饰占比"),
        ("title_pt_per_inch", "标题字号/版面宽"),
        ("body_pt", "正文字号 pt"),
        ("size_steps", "字号档数"),
    ]:
        lines.append(f"| {label} | {median(key)} | {weighted(key)} |")

    lines += ["", "## 字体", ""]
    lines += [f"- {name}（{count} 份 deck 使用）" for name, count in font_votes.most_common(6)]

    lines += ["", "## 逐份", "", "| 文件 | 页 | 元素/页 | 图/页 | 字/页 | 标题pt/in | 正文pt | 档数 |"]
    lines.append("|---|---|---|---|---|---|---|---|")
    for profile in profiles:
        lines.append(
            "| {file} | {pages} | {objects_per_page} | {pictures_per_page} | {chars_per_page} | "
            "{title_pt_per_inch} | {body_pt} | {size_steps} |".format(**profile)
        )
    return "\n".join(lines) + "\n"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Measure human-made decks.")
    parser.add_argument("decks", nargs="+")
    parser.add_argument("--out", default="out/reference_profile")
    parser.add_argument("--aggregate", action="store_true", help="Also write an aggregated Markdown summary")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    profiles = [analyse(Path(path)) for path in args.decks]
    out = Path(args.out)
    if len(profiles) == 1 and not args.aggregate:
        out.parent.mkdir(parents=True, exist_ok=True)
        out = out.with_suffix(".json")
        out.write_text(json.dumps(profiles[0], ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(profiles[0], ensure_ascii=False, indent=2))
        return
    out = out.with_suffix(".md") if out.suffix != ".md" else out
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(aggregate(profiles), encoding="utf-8")
    json_out = out.with_suffix(".json")
    json_out.write_text(json.dumps(profiles, ensure_ascii=False, indent=2), encoding="utf-8")
    print(aggregate(profiles))


if __name__ == "__main__":
    main()
