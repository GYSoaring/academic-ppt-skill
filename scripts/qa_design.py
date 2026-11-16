#!/usr/bin/env python3
"""Design quality audit for an academic deck.

The skill can already prove that a deck is not *broken* (no overflow, no missing
assets). This script answers the other question: does the deck read like it was
designed, or does it read like a form?

Every metric below is a proxy for something a human reviewer notices in the
first five seconds:

- `dominant_object_ratio`  每页有没有一个主角（图/图/公式/图式/表），而不是一堆等权卡片
- `card_only_ratio`        整页只有等权卡片的比例（模板感的主要来源）
- `emphasis_per_page`      红色强调的密度：没有强调 = 没有重点
- `assertion_ratio`        标题是判断还是名词标签
- `takeaway_ratio`         深蓝色结论带的出现频率（每页都有 = 装饰）
- `layout_run_max`         同一种版式连续出现多少次（连排 3 页就腻）
- `image_page_ratio`       真图页占比
- `chars_per_line`         正文一行能放多少汉字（低于 18 字，中文读起来是碎的）
- `rhythm_stdev`           各页墨迹覆盖率的标准差（越小越像模板）
- `chrome_share`           页眉+lead+结论带+目录 占版面比例

Thresholds are deliberately conservative: a deck should be able to hit them
without becoming a poster.
"""

from __future__ import annotations

import argparse
import json
import statistics
from pathlib import Path

from common import apply_type_profile, is_assertion_title, load_json, load_theme, write_json

CONTENT_LAYOUTS = {
    "bullets",
    "figure",
    "figure_flow",
    "concept",
    "formula",
    "pipeline",
    "compare",
    "metrics",
    "grid",
    "balance",
    "hub",
    "timeline",
    "quote",
}
DECOR_LAYOUTS = {"cover", "outline", "section", "closing", "references"}
# Only prose layouts are judged by characters-per-line. A table cell or a five
# step pipeline is scanned, not read as a line of text.
TEXT_LAYOUTS = {"bullets", "grid"}
# A table is a designed object, not a row of equal-weight cards; only grid and
# bullets can produce the "form of identical boxes" the metric is looking for.
CARD_LAYOUTS = {"grid", "bullets"}

TARGETS = {
    "dominant_object_ratio": 0.60,
    "card_only_ratio": 0.25,
    "emphasis_per_page": 0.60,
    "assertion_ratio": 0.60,
    "takeaway_ratio": 0.35,
    "layout_run_max": 2,
    "image_page_ratio": 0.15,
    "chars_per_line": 18.0,
    "rhythm_stdev": 0.050,
    "chrome_share": 0.32,
}

WEIGHTS = {
    "dominant_object_ratio": 16,
    "card_only_ratio": 12,
    "emphasis_per_page": 14,
    "assertion_ratio": 12,
    "takeaway_ratio": 8,
    "layout_run_max": 8,
    "image_page_ratio": 8,
    "chars_per_line": 12,
    "rhythm_stdev": 6,
    "chrome_share": 4,
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Audit an academic deck for design quality.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--theme", default=None)
    parser.add_argument("--render-dir", default=None)
    parser.add_argument("--out", default=None, help="Write design_report.json / .md here.")
    parser.add_argument(
        "--profile",
        choices=["document", "projection", "lecture", "defense"],
        default=None,
        help="Type profile the deck is rendered with; characters-per-line depends on it.",
    )
    return parser.parse_args()


def count_emphasis(value) -> int:
    text = json.dumps(value, ensure_ascii=False)
    return text.count("**") // 2


def has_dominant_object(slide: dict) -> bool:
    if slide.get("image") or slide.get("chart") or slide.get("diagram"):
        return True
    if slide.get("formula") or slide.get("formula_latex"):
        return True
    if slide.get("layout") in {"figure", "figure_flow", "concept", "formula", "hub", "balance", "metrics", "quote"}:
        return True
    for block in slide.get("blocks") or []:
        if isinstance(block, dict) and (block.get("image") or block.get("diagram") or block.get("formula")):
            return True
    return False


def column_count(slide: dict) -> int:
    layout = slide.get("layout")
    raw_columns = slide.get("columns")
    if layout == "bullets":
        # `bullets` uses a list of column objects; `grid` uses an int.
        if isinstance(raw_columns, list):
            return max(1, min(3, len(raw_columns)))
        return max(1, min(3, int(raw_columns or 3)))
    if layout == "grid":
        return max(1, min(3, int(raw_columns or 2)))
    if layout in {"compare", "timeline"}:
        return max(1, len(slide.get("columns") or []) or 1)
    if layout in {"pipeline", "figure_flow"}:
        return max(1, min(5, len(slide.get("steps") or slide.get("flow") or []) or 1))
    return 1


def chars_per_line(slide: dict, theme: dict) -> float | None:
    layout = slide.get("layout")
    if layout not in TEXT_LAYOUTS:
        return None
    layout_cfg = theme.get("layout", {}) if theme else {}
    content = layout_cfg.get("content") or [320, 186, 1992, 986]
    if str(layout_cfg.get("sidebar_mode", "always")) == "off":
        content = layout_cfg.get("content_no_sidebar") or content
    width = float(content[2]) - float(content[0])
    columns = column_count(slide)
    gap = 18.0
    cell = (width - gap * (columns - 1)) / max(1, columns)
    scale = theme.get("type_scale", {}) if theme else {}
    body_px = float(scale.get("body", 32) or 32)
    if layout == "compare":
        body_px = float(scale.get("table_cell", body_px) or body_px)
    return round(cell / max(1.0, body_px), 1)


def ink_coverage(render_dir: Path | None) -> list[float]:
    if not render_dir or not Path(render_dir).exists():
        return []
    try:
        from PIL import Image
    except Exception:
        return []
    values: list[float] = []
    for path in sorted(Path(render_dir).glob("slide_*.png")):
        image = Image.open(path).convert("L")
        width, height = image.size
        box = (
            int(width * 0.10),
            int(height * 0.16),
            int(width * 0.97),
            int(height * 0.86),
        )
        crop = image.crop(box)
        pixels = list(crop.getdata())
        if not pixels:
            continue
        values.append(sum(1 for value in pixels if value < 225) / len(pixels))
    return values


def chrome_share(deck: dict, theme: dict) -> float:
    layout_cfg = theme.get("layout", {}) if theme else {}
    height = float((theme.get("canvas", {}) or {}).get("height", 1152) or 1152)
    width = float((theme.get("canvas", {}) or {}).get("width", 2048) or 2048)
    if not height or not width:
        return 0.0
    header = float(layout_cfg.get("header_height", 112) or 112)
    lead = float(layout_cfg.get("lead_height", 88) or 88)
    takeaway = float(layout_cfg.get("takeaway_height", 68) or 68)
    sidebar = layout_cfg.get("sidebar") or [0, 0, 0, 0]
    sidebar_area = 0.0
    if str(layout_cfg.get("sidebar_mode", "always")) != "off":
        sidebar_area = (float(sidebar[2]) - float(sidebar[0])) * (float(sidebar[3]) - float(sidebar[1]))
    content_pages = [slide for slide in deck.get("slides", []) if slide.get("layout") not in {"cover", "outline"}]
    if not content_pages:
        return 0.0
    shares = []
    for slide in content_pages:
        area = header * width + lead * width + sidebar_area
        if slide.get("takeaway"):
            area += takeaway * width
        shares.append(area / (width * height))
    return round(statistics.mean(shares), 3)


def audit_design(deck: dict, theme: dict | None, render_dir: Path | None = None) -> dict:
    slides = deck.get("slides", [])
    content = [slide for slide in slides if slide.get("layout") in CONTENT_LAYOUTS]
    total_content = max(1, len(content))

    dominant = sum(1 for slide in content if has_dominant_object(slide))
    card_only = sum(
        1
        for slide in content
        if not has_dominant_object(slide) and slide.get("layout") in CARD_LAYOUTS
    )
    emphasis = sum(count_emphasis(slide) for slide in content)
    assertions = [slide for slide in content if is_assertion_title(str(slide.get("title", "")))]
    takeaways = sum(1 for slide in content if slide.get("takeaway"))
    images = sum(1 for slide in content if slide.get("image"))

    run_max = 0
    run_layout = None
    run = 0
    # A section divider interrupts the visual rhythm even though it carries no
    # content, so runs are measured across everything except cover and outline.
    for slide in [item for item in slides if item.get("layout") not in {"cover", "outline"}]:
        current = slide.get("layout")
        if current == run_layout:
            run += 1
        else:
            run_layout = current
            run = 1
        run_max = max(run_max, run)

    lines = [value for value in (chars_per_line(slide, theme or {}) for slide in content) if value is not None]
    worst_line = min(lines) if lines else 0.0
    ink_coverage_values = ink_coverage(render_dir)
    inks = ink_coverage_values
    rhythm = statistics.pstdev(inks[2:-2]) if len(inks) > 4 else 0.0

    metrics = {
        "content_pages": len(content),
        "dominant_object_ratio": round(dominant / total_content, 3),
        "card_only_ratio": round(card_only / total_content, 3),
        "emphasis_per_page": round(emphasis / total_content, 3),
        "assertion_ratio": round(len(assertions) / total_content, 3),
        "takeaway_ratio": round(takeaways / total_content, 3),
        "layout_run_max": run_max,
        "image_page_ratio": round(images / total_content, 3),
        "chars_per_line": worst_line,
        "rhythm_stdev": round(rhythm, 4),
        "chrome_share": chrome_share(deck, theme or {}),
    }

    checks = []
    have_rhythm = bool(ink_coverage_values)
    have_prose = bool([value for value in (chars_per_line(slide, theme or {}) for slide in content) if value is not None])
    for name, target in TARGETS.items():
        value = metrics.get(name)
        if value is None:
            continue
        # Without rendered pages the rhythm metric cannot be measured; skipping it
        # is honest, marking it failed would punish the caller for not rendering.
        if name == "rhythm_stdev" and not have_rhythm:
            continue
        # A deck made only of figures, formulas and tables has no prose column to
        # measure; that is a design choice, not a zero.
        if name == "chars_per_line" and not have_prose:
            continue
        if name in {"card_only_ratio", "takeaway_ratio", "layout_run_max", "chrome_share"}:
            ok = value <= target
            direction = "≤"
        else:
            ok = value >= target
            direction = "≥"
        checks.append(
            {
                "metric": name,
                "value": value,
                "target": f"{direction} {target}",
                "ok": bool(ok),
                "weight": WEIGHTS.get(name, 0),
            }
        )

    earned = sum(check["weight"] for check in checks if check["ok"])
    total_weight = sum(WEIGHTS.get(check["metric"], 0) for check in checks) or 1
    score = round(100 * earned / total_weight)

    if score >= 85:
        grade = "A｜可以直接拿去讲"
    elif score >= 70:
        grade = "B｜结构清楚了，还差气口"
    elif score >= 55:
        grade = "C｜能看，但仍是版面填空题"
    else:
        grade = "D｜模板感明显"

    return {"score": score, "grade": grade, "metrics": metrics, "checks": checks}


def render_markdown(report: dict, deck_title: str) -> str:
    lines = [f"# 设计体检：{deck_title}", "", f"**{report['score']} / 100 —— {report['grade']}**", ""]
    lines.append("| 指标 | 实测 | 目标 | 结论 |")
    lines.append("|---|---|---|---|")
    for check in report["checks"]:
        mark = "✅" if check["ok"] else "⚠️"
        lines.append(f"| `{check['metric']}` | {check['value']} | {check['target']} | {mark} |")
    lines.append("")
    failed = [check for check in report["checks"] if not check["ok"]]
    if failed:
        lines.append("## 先修这几个")
        lines.append("")
        for check in failed:
            lines.append(f"- `{check['metric']}` = {check['value']}，目标 {check['target']}")
    else:
        lines.append("全部达标。")
    return "\n".join(lines) + "\n"


def main() -> None:
    args = parse_args()
    deck_path = Path(args.deck).resolve()
    deck = load_json(deck_path)
    theme = load_theme(Path(args.theme)) if args.theme else None
    if theme is not None:
        apply_type_profile(theme, args.profile or str(deck.get("meta", {}).get("type_profile", "document")))
    render_dir = Path(args.render_dir) if args.render_dir else None
    report = audit_design(deck, theme, render_dir)
    report["deck"] = str(deck_path)
    report["title"] = str(deck.get("meta", {}).get("title", ""))
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if args.out:
        out = Path(args.out)
        out.mkdir(parents=True, exist_ok=True)
        write_json(out / "design_report.json", report)
        (out / "design_report.md").write_text(
            render_markdown(report, report["title"] or deck_path.stem), encoding="utf-8"
        )


if __name__ == "__main__":
    main()
