from __future__ import annotations

import argparse
import json
import re
import shutil
from pathlib import Path

from common import load_json, write_json


STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "by", "for", "from", "in", "into",
    "is", "of", "on", "or", "that", "the", "their", "this", "to", "using", "via", "with",
}
GLOSSARY = {
    "视觉": "vision",
    "语言": "language",
    "机器人": "robot",
    "具身": "embodied",
    "智能": "intelligence",
    "技能": "skill",
    "执行": "execution",
    "动作": "action",
    "风险": "risk",
    "评价": "evaluation",
    "实验": "experiment",
    "数据": "data",
    "模型": "model",
    "感知": "perception",
    "路径": "path",
    "规划": "planning",
    "学习": "learning",
    "任务": "task",
    "场景": "scene",
    "泛化": "generalization",
    "失败": "failure",
    "监控": "monitoring",
    "闭环": "closed loop",
    "方法": "method",
    "架构": "architecture",
    "系统": "system",
    "训练": "training",
    "推理": "inference",
    "多模态": "multimodal",
    "验证": "verification",
    "约束": "constraint",
    "条件": "condition",
    "证据": "evidence",
    "地图": "map",
    "导航": "navigation",
    "观测": "observation",
    "长尾": "long-tail",
    "不确定性": "uncertainty",
    "安全": "safety",
    "碰撞": "collision",
    "抓取": "grasping",
    "放置": "placement",
    "物体": "object",
    "关系": "relation",
    "可行性": "feasibility",
    "前提": "precondition",
    "语义": "semantic",
    "物理": "physical",
    "归因": "attribution",
    "更新": "update",
    "策略": "policy",
    "比较": "comparison",
    "基线": "baseline",
    "消融": "ablation",
    "外场": "field",
}
CONTENT_LAYOUTS = {
    "bullets", "figure", "figure_flow", "concept", "formula", "pipeline",
    "compare", "metrics", "grid", "balance", "hub", "timeline", "quote",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Suggest paper figures for deck slides.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--figures", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--query", default=None, help="English deck-level query; defaults to meta.figure_query.")
    parser.add_argument("--base-dir", default=None, help="Directory that figure paths are relative to.")
    parser.add_argument("--top", type=int, default=3)
    parser.add_argument("--min-score", type=float, default=0.15)
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--out-deck", default=None)
    parser.add_argument("--asset-dir", default=None)
    return parser.parse_args()


def tokens(text: str) -> set[str]:
    result = {
        token
        for token in re.findall(r"[a-z0-9]+|[\u4e00-\u9fff]", str(text).lower())
        if token not in STOPWORDS and len(token) > 1
    }
    lowered = str(text).lower()
    for chinese, english in GLOSSARY.items():
        if chinese in lowered:
            result.update(english.split())
    return result


def figure_text(figure: dict) -> str:
    return " ".join(
        str(figure.get(key, ""))
        for key in ("caption", "paper_title", "alt_hint")
    )


def score(slide_text: str, figure: dict, deck_tokens: set[str]) -> float:
    query = tokens(slide_text)
    if not query:
        query = deck_tokens
    figure_tokens = tokens(figure_text(figure))
    slide_score = len(query & figure_tokens) / max(1, len(query))
    deck_score = len(deck_tokens & figure_tokens) / max(1, len(deck_tokens))
    return round(0.85 * slide_score + 0.15 * deck_score, 3)


def load_figures(path: Path) -> list[dict]:
    data = load_json(path)
    if isinstance(data, dict):
        return list(data.get("figures") or [])
    return list(data or [])


def apply_suggestions(
    deck_path: Path,
    deck: dict,
    suggestions: dict,
    figures: list[dict],
    base_dir: Path,
    out_deck: Path,
    asset_dir: Path,
    overwrite: bool = False,
) -> int:
    by_id = {figure["id"]: figure for figure in figures}
    used: set[str] = set()
    applied = 0
    asset_dir.mkdir(parents=True, exist_ok=True)
    for slide in deck.get("slides", []):
        page = str(slide.get("id", ""))
        if str(slide.get("layout")) not in CONTENT_LAYOUTS:
            continue
        if (slide.get("image") or slide.get("chart")) and not overwrite:
            continue
        candidates = suggestions.get(page) or []
        chosen = next((item for item in candidates if item["id"] not in used), candidates[0] if candidates else None)
        if not chosen:
            continue
        figure = by_id.get(chosen["id"])
        if not figure:
            continue
        source_path = Path(figure.get("path", ""))
        if not source_path.is_absolute():
            source_path = base_dir / source_path
        if not source_path.exists():
            continue
        target = asset_dir / source_path.name
        shutil.copy2(source_path, target)
        try:
            image_value = target.relative_to(deck_path.parent).as_posix()
        except ValueError:
            image_value = str(target)
        caption = str(figure.get("caption") or figure.get("alt_hint") or "").strip()
        paper_title = str(figure.get("paper_title") or Path(str(figure.get("source", ""))).stem).strip()
        source = str(figure.get("source") or "")
        if figure.get("arxiv_id"):
            source = f"{source} ({figure['arxiv_id']})"
        slide["image"] = image_value
        slide["image_label"] = (caption or paper_title)[:48]
        slide["image_caption"] = caption[:220]
        slide["image_role"] = f"来自公开论文的图片：{paper_title}；请确认它在本页承担的论证角色"
        slide["image_boundary"] = "只支持图中展示的内容；具体边界需作者确认"
        slide["image_origin"] = "borrowed"
        slide["source"] = f"{source}，第 {figure.get('page', '?')} 页"
        used.add(figure["id"])
        applied += 1
    write_json(out_deck, deck)
    return applied


def main() -> None:
    args = parse_args()
    deck_path = Path(args.deck).resolve()
    figures_path = Path(args.figures).resolve()
    out_dir = Path(args.out).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    base_dir = Path(args.base_dir).resolve() if args.base_dir else figures_path.parent
    deck = load_json(deck_path)
    figures = load_figures(figures_path)
    deck_query = args.query or str(deck.get("meta", {}).get("figure_query", "")).strip()
    deck_tokens = tokens(deck_query)
    suggestions: dict[str, list[dict]] = {}
    for slide in deck.get("slides", []):
        if str(slide.get("layout")) not in CONTENT_LAYOUTS:
            continue
        page = str(slide.get("id", ""))
        slide_text = " ".join(
            str(slide.get(key, ""))
            for key in ("title", "lead", "claim", "evidence", "takeaway")
        )
        ranked = []
        for figure in figures:
            value = score(slide_text, figure, deck_tokens)
            if value >= args.min_score:
                ranked.append(
                    {
                        "id": figure["id"],
                        "score": round(value, 3),
                        "caption": figure.get("caption") or figure.get("alt_hint") or "",
                        "paper_title": figure.get("paper_title") or "",
                        "path": figure.get("path") or "",
                    }
                )
        ranked.sort(key=lambda item: item["score"], reverse=True)
        suggestions[page] = ranked[: max(1, args.top)]

    report = {
        "deck": str(deck_path),
        "figures": str(figures_path),
        "min_score": args.min_score,
        "deck_query": deck_query,
        "suggestions": suggestions,
        "applied": 0,
    }
    lines = [
        f"# Figure Suggestions: {deck.get('meta', {}).get('title', '')}",
        "",
        f"> 相关性阈值 {args.min_score}；只做候选，不自动声称图片支持某条结论。",
        "",
    ]
    for page, items in suggestions.items():
        lines.append(f"## 第 {page} 页")
        lines.append("")
        if not items:
            lines.append("- 没有达到阈值的候选。")
            lines.append("")
            continue
        for item in items:
            lines.append(f"- `{item['id']}` — score {item['score']} — {item['caption'][:160]}")
        lines.append("")
    (out_dir / "figure_suggestions.md").write_text("\n".join(lines), encoding="utf-8")

    if args.apply:
        out_deck = Path(args.out_deck).resolve() if args.out_deck else deck_path.with_name(deck_path.stem + ".with_figures.json")
        asset_dir = Path(args.asset_dir).resolve() if args.asset_dir else deck_path.parent / "assets" / "figures"
        report["applied"] = apply_suggestions(
            deck_path,
            deck,
            suggestions,
            figures,
            base_dir,
            out_deck,
            asset_dir,
            overwrite=args.overwrite,
        )
        report["out_deck"] = str(out_deck)
    write_json(out_dir / "figure_suggestions.json", report)
    print(json.dumps({"suggestions": len(suggestions), "applied": report["applied"]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
