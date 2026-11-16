from __future__ import annotations

import argparse
import json
import statistics
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from common import is_defense_secondary_text, write_json


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Audit an arbitrary PPTX for academic-defense readability.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--min-body-pt", type=float, default=24.0)
    parser.add_argument("--min-any-pt", type=float, default=18.0)
    parser.add_argument("--min-heading-pt", type=float, default=32.0)
    return parser.parse_args()


def bucket(size: float) -> str:
    if size < 10:
        return "<10"
    if size < 14:
        return "10-14"
    if size < 18:
        return "14-18"
    if size < 24:
        return "18-24"
    if size < 32:
        return "24-32"
    return ">=32"


def choose_verdict(
    *,
    body_median: float,
    median: float,
    table_median: float,
    body_below_24_ratio: float,
    below_18_ratio: float,
    small_text_ratio: float,
    table_below_18_ratio: float,
    min_body_pt: float,
    min_any_pt: float,
    body_bold_ratio: float = 1.0,
) -> str:
    if median < min_any_pt or body_median < min_any_pt:
        return "rebuild"
    if (
        body_median < min_body_pt - 0.5
        or body_below_24_ratio > 0.15
        or (table_median and table_median < min_any_pt)
        or table_below_18_ratio > 0.05
        or below_18_ratio > 0.05
        or small_text_ratio > 0.75
        or body_bold_ratio < 0.9
    ):
        return "font_upgrade"
    return "defense_ready"


def main() -> None:
    args = parse_args()
    pptx_path = Path(args.pptx).resolve()
    output_dir = Path(args.out).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    presentation = Presentation(str(pptx_path))
    slides_report: list[dict] = []
    all_sizes: list[float] = []
    body_sizes: list[float] = []
    table_sizes: list[float] = []
    reference_sizes: list[float] = []
    buckets = {"<10": 0, "10-14": 0, "14-18": 0, "18-24": 0, "24-32": 0, ">=32": 0}
    secondary_sizes: list[float] = []
    bold_count = 0
    run_count = 0
    body_bold_count = 0
    body_run_count = 0
    total_pictures = 0
    total_tables = 0
    total_charts = 0
    for index, slide in enumerate(presentation.slides, start=1):
        title_text = ""
        try:
            if slide.shapes.title is not None:
                title_text = slide.shapes.title.text.strip()
        except Exception:
            title_text = ""
        is_reference = "参考文献" in title_text or "reference" in title_text.lower() or "bibliography" in title_text.lower()
        sizes: list[float] = []
        slide_body_sizes: list[float] = []
        slide_table_sizes: list[float] = []
        slide_runs = 0
        slide_bold = 0
        slide_body_runs = 0
        slide_body_bold = 0
        pictures = 0
        tables = 0
        charts = 0
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pictures += 1
            if getattr(shape, "has_table", False):
                tables += 1
                for row in shape.table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if run.text.strip() and run.font.size:
                                    size = round(run.font.size.pt, 1)
                                    table_sizes.append(size)
                                    slide_table_sizes.append(size)
                                    slide_runs += 1
                                    all_sizes.append(size)
                                    buckets[bucket(size)] += 1
                                    if run.font.bold:
                                        slide_bold += 1
            if getattr(shape, "has_chart", False):
                charts += 1
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if not run.text.strip():
                        continue
                    slide_runs += 1
                    if run.font.size:
                        size = round(run.font.size.pt, 1)
                        sizes.append(size)
                        all_sizes.append(size)
                        buckets[bucket(size)] += 1
                        content = run.text.strip()
                        if is_reference or is_defense_secondary_text(shape.name, content, size):
                            secondary_sizes.append(size)
                            if is_reference:
                                reference_sizes.append(size)
                        elif len(content) >= 10 and size >= args.min_any_pt - 4:
                            body_sizes.append(size)
                            slide_body_sizes.append(size)
                            body_run_count += 1
                            slide_body_runs += 1
                            if run.font.bold:
                                body_bold_count += 1
                                slide_body_bold += 1
                    if run.font.bold:
                        slide_bold += 1
        total_pictures += pictures
        total_tables += tables
        total_charts += charts
        run_count += slide_runs
        bold_count += slide_bold
        below_18 = sum(1 for size in sizes if size < args.min_any_pt)
        below_24 = sum(1 for size in sizes if size < args.min_body_pt)
        slides_report.append(
            {
                "page": index,
                "runs": slide_runs,
                "font_size_min_pt": min(sizes) if sizes else None,
                "font_size_median_pt": round(statistics.median(sizes), 1) if sizes else None,
                "font_size_max_pt": max(sizes) if sizes else None,
                "body_font_median_pt": round(statistics.median(slide_body_sizes), 1) if slide_body_sizes else None,
                "table_font_median_pt": round(statistics.median(slide_table_sizes), 1) if slide_table_sizes else None,
                "runs_below_18pt": below_18,
                "runs_below_24pt": below_24,
                "bold_ratio": round(slide_bold / max(1, slide_runs), 2),
                "body_bold_ratio": round(slide_body_bold / max(1, slide_body_runs), 2),
                "pictures": pictures,
                "tables": tables,
                "charts": charts,
            }
        )
    median = round(statistics.median(all_sizes), 1) if all_sizes else 0
    body_median = round(statistics.median(body_sizes), 1) if body_sizes else 0
    table_median = round(statistics.median(table_sizes), 1) if table_sizes else 0
    reference_median = round(statistics.median(reference_sizes), 1) if reference_sizes else 0
    below_18_ratio = sum(1 for size in all_sizes if size < args.min_any_pt) / max(1, len(all_sizes))
    below_24_ratio = sum(1 for size in all_sizes if size < args.min_body_pt) / max(1, len(all_sizes))
    small_text_ratio = sum(1 for size in secondary_sizes if size < args.min_any_pt) / max(1, len(all_sizes))
    audit_median = body_median or median
    body_below_24 = sum(1 for size in body_sizes if size < args.min_body_pt) / max(1, len(body_sizes))
    table_below_18 = sum(1 for size in table_sizes if size < args.min_any_pt) / max(1, len(table_sizes))
    body_bold_ratio = body_bold_count / max(1, body_run_count)
    verdict = choose_verdict(
        body_median=body_median,
        median=median,
        table_median=table_median,
        body_below_24_ratio=body_below_24,
        below_18_ratio=below_18_ratio,
        small_text_ratio=small_text_ratio,
        table_below_18_ratio=table_below_18,
        min_body_pt=args.min_body_pt,
        min_any_pt=args.min_any_pt,
        body_bold_ratio=body_bold_ratio,
    )
    report = {
        "pptx": str(pptx_path),
        "slide_count": len(presentation.slides),
        "runs": run_count,
        "font_size_median_pt": median,
        "body_runs": len(body_sizes),
        "body_font_median_pt": body_median,
        "body_runs_below_24pt_ratio": round(body_below_24, 2),
        "table_runs": len(table_sizes),
        "table_font_median_pt": table_median,
        "reference_runs": len(reference_sizes),
        "reference_font_median_pt": reference_median,
        "font_size_min_pt": round(min(all_sizes), 1) if all_sizes else None,
        "font_size_max_pt": round(max(all_sizes), 1) if all_sizes else None,
        "font_size_buckets": buckets,
        "runs_below_18pt_ratio": round(below_18_ratio, 2),
        "runs_below_24pt_ratio": round(below_24_ratio, 2),
        "secondary_runs": len(secondary_sizes),
        "secondary_runs_below_18pt_of_all_ratio": round(small_text_ratio, 2),
        "secondary_runs_below_24pt_ratio": round(
            sum(1 for size in secondary_sizes if size < args.min_body_pt) / max(1, len(secondary_sizes)), 2
        ),
        "bold_ratio": round(bold_count / max(1, run_count), 2),
        "body_bold_ratio": round(body_bold_ratio, 2),
        "table_runs_below_18pt_ratio": round(
            table_below_18, 2
        ),
        "runs_below_18pt_count": sum(1 for size in all_sizes if size < args.min_any_pt),
        "runs_below_24pt_count": sum(1 for size in all_sizes if size < args.min_body_pt),
        "verdict_reasons": [
            f"body median {body_median} pt",
            f"all-run median {median} pt",
            f"table median {table_median} pt",
            f"body runs below {args.min_body_pt:.0f} pt {body_below_24:.0%}",
            f"all runs below {args.min_any_pt:.0f} pt {below_18_ratio:.0%}",
            f"secondary runs below {args.min_any_pt:.0f} pt of all runs {small_text_ratio:.0%}",
            f"table runs below {args.min_any_pt:.0f} pt {table_below_18:.0%}",
            f"semantic body bold {body_bold_ratio:.0%}",
        ],
        "pictures": total_pictures,
        "tables": total_tables,
        "charts": total_charts,
        "thresholds": {
            "min_body_pt": args.min_body_pt,
            "min_any_pt": args.min_any_pt,
            "min_heading_pt": args.min_heading_pt,
        },
        "verdict": verdict,
        "slides": slides_report,
    }
    write_json(output_dir / "pptx_readability.json", report)
    lines = [
        f"# PPTX 可读性体检：{pptx_path.name}",
        "",
        "> 口径：`defense_ready` 要求正文目标 ≥24 pt、任何可见文字 ≥18 pt，正文低于 24 pt ≤15%、全篇低于 18 pt ≤5%、表格低于 18 pt ≤5%、二级文字低于 18 pt 占全部 run ≤75%、正文候选加粗率 ≥90%。字号来自 PPTX 实际 run；只把正文中位放大不能掩盖密集小字或常规字重正文。",
        "",
        "## 总览",
        "",
        f"- 页数：{report['slide_count']}",
        f"- 文字 run：{report['runs']}",
        f"- 字号中位数：{report['font_size_median_pt']} pt（最小 {report['font_size_min_pt']}，最大 {report['font_size_max_pt']}）",
        f"- 正文候选：{report['body_runs']} 个 run，中位数 {report['body_font_median_pt']} pt，低于 24 pt 占 {report['body_runs_below_24pt_ratio']:.0%}",
        f"- 表格文字：{report['table_runs']} 个 run，中位数 {report['table_font_median_pt']} pt",
        f"- 参考文献：{report['reference_runs']} 个 run，中位数 {report['reference_font_median_pt']} pt",
        f"- 低于 18 pt：{report['runs_below_18pt_ratio']:.0%}",
        f"- 低于 24 pt（全部 run）：{report['runs_below_24pt_ratio']:.0%}",
        f"- 二级小字：{report['secondary_runs']} 个 run；其中低于 24 pt 占 {report['secondary_runs_below_24pt_ratio']:.0%}，低于 18 pt 占全部 run {report['secondary_runs_below_18pt_of_all_ratio']:.0%}",
        f"- 判定依据：{'；'.join(report['verdict_reasons'])}",
        f"- 加粗比例：{report['bold_ratio']:.0%}",
        f"- 正文候选加粗比例：{report['body_bold_ratio']:.0%}",
        f"- 图片 {report['pictures']}，表格 {report['tables']}，图表 {report['charts']}",
        f"- 结论：**{verdict}**",
        "",
        "## 字号分布",
        "",
        "| 区间 | run 数 |",
        "|---|---:|",
    ]
    for name, count in buckets.items():
        lines.append(f"| {name} pt | {count} |")
    lines.extend(
        [
            "",
            "## 逐页",
            "",
            "| 页 | run | 最小 | 中位 | 正文中位 | 最大 | <18 | <24 | 加粗 | 图 | 表 |",
            "|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
        ]
    )
    for slide in slides_report:
        lines.append(
            f"| {slide['page']:02d} | {slide['runs']} | {slide['font_size_min_pt']} | "
            f"{slide['font_size_median_pt']} | {slide['body_font_median_pt']} | {slide['font_size_max_pt']} | "
            f"{slide['runs_below_18pt']} | {slide['runs_below_24pt']} | "
            f"{slide['bold_ratio']:.0%} | {slide['pictures']} | {slide['tables']} |"
        )
    lines.extend(
        [
            "",
            "## 改造建议",
            "",
            "- `rebuild`：全篇或正文中位数低于 18 pt，不要局部调大，按 `defense` 档重建；每页内容量要减半。",
            "- `font_upgrade`：正文候选未达 24 pt、存在低于 18 pt 的文字，或二级小字占全部 run 超过 75%。优先拆页并把标签、脚注、表格和参考文献元数据 token 化，不要只调正文字号。",
            "- 用 `extract_source.py --input original.pptx` 保留原文、备注和图片，再按 `defense` 档重建 `deck.json` 并用 `audit_pptx.py` 复查；不要直接把正文缩到 24 pt 以下。",
            "- 图片数量大于 0 不等于有论文图：检查每张图是否有 `image_role`、`image_boundary` 和来源。",
            "- 校名如果是普通字体而不是官方书法字标，使用 `build_brand_lockup.py --lockup` 替换。",
            "",
        ]
    )
    (output_dir / "pptx_readability.md").write_text("\n".join(lines), encoding="utf-8")
    print(json.dumps({k: report[k] for k in ("slide_count", "font_size_median_pt", "runs_below_24pt_ratio", "body_bold_ratio", "verdict")}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
