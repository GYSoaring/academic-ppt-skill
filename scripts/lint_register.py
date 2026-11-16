#!/usr/bin/env python3
"""Flag colloquial, hyped or empty phrasing in a deck's visible text.

`validate_deck.py` checks structure and layout. This checks **register**: the
difference between a slide a professor would put up and one that reads like a
chat summary. Rules come from two community academic skills (see
`references/academic-register.md`).

Speech notes are exempt — spoken language is allowed to be conversational.
Only text that appears on the page is linted.

Usage:
  python scripts/lint_register.py --deck deck.json
  python scripts/lint_register.py --deck deck.json --profile lecture --strict
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

# (pattern, severity, message, suggestion)
RULES: list[tuple[str, str, str, str]] = [
    # 口语标记：页面正文里不该出现
    (r"各位老师|大家好|谢谢大家", "error", "讲台用语出现在正文", "删掉，或移到 speech"),
    (r"我们来看|我们看|咱们", "error", "第一人称口语", "改为客观陈述"),
    (r"(?<![及极尤可])其实|就是说|那么(?=，)|呃|基本上|大概齐", "warning", "口语填充词", "删除"),
    (r"挺好的|好多|特别(?!是)|超级|差不多|差不多|搞定|弄一下", "warning", "口语化表达", "改为书面语"),
    (r"非常(?![\u4e00-\u9fa5]{0,2}(大|小|高|低|快|慢))", "warning", "空洞强度词", "给出数值或删除"),
    # 意义吹捧
    (r"重大突破|颠覆性|前所未有|革命性|里程碑", "error", "意义吹捧", "写明解决了哪个具体问题"),
    (r"开辟(新|一条新)|奠定(了)?基础|填补(了)?空白|具有重要意义|意义重大", "error", "套话", "换成具体贡献"),
    (r"受到(广泛|学界)关注|引起了广泛关注", "warning", "套式开头", "直接写问题本身"),
    # 空泛归因
    (r"随着.{0,12}的(快速|飞速)?发展", "warning", "套式开头", "直接给事实或矛盾"),
    (r"(相关|已有|大量)(研究|工作)(表明|显示|指出)(?![^，。]{0,12}(等|arXiv|Nature|Science|PRL|Opt))",
     "error", "无出处的引用归因", "写出作者或文献编号"),
    (r"大量研究|广泛研究|众多工作", "warning", "引用倾倒", "只引一两篇关键工作并说明理由"),
    # 模糊限定
    (r"一定程度上|某种程度上|相对(较好|不错)|较为(良好|理想)|效果不错", "warning", "模糊限定", "量化或删除"),
    (r"显著(提升|提高|改善|优于)(?![^，。]{0,16}\d)", "error", "无数字的显著", "给出数值、基线与条件"),
    # 首创标榜
    (r"首次提出|据我们所知|国内外首创", "warning", "首创标榜", "说明与前人工作的差别"),
    # 模板强调
    (r"值得注意的是|需要强调的是|众所周知|不难发现", "warning", "模板强调", "删除，让内容自己说话"),
    # 贡献清单
    (r"主要贡献(是|为)[:：]?\s*[①②③1-9]", "warning", "贡献清单套式", "每条贡献绑定具体结果"),
]

# 这些字段是讲稿，不检查
SKIP_KEYS = {"speech", "closing_speech", "note_speech"}


def iter_text(node, path=""):
    if isinstance(node, dict):
        for key, value in node.items():
            if key in SKIP_KEYS:
                continue
            yield from iter_text(value, f"{path}.{key}" if path else key)
    elif isinstance(node, list):
        for index, value in enumerate(node):
            yield from iter_text(value, f"{path}[{index}]")
    elif isinstance(node, str):
        yield path, node


def main() -> None:
    parser = argparse.ArgumentParser(description="Lint academic register in a deck.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--out", default=None)
    parser.add_argument("--strict", action="store_true", help="Warnings also fail the run")
    args = parser.parse_args()

    deck = json.loads(Path(args.deck).read_text(encoding="utf-8"))
    findings = []
    for index, slide in enumerate(deck.get("slides", []), 1):
        slide_id = slide.get("id", str(index))
        is_closing = slide.get("layout") == "closing"
        for path, text in iter_text(slide):
            if "image_" in path or path.endswith("url"):
                continue
            for pattern, severity, message, suggestion in RULES:
                if is_closing and message == "讲台用语出现在正文":
                    continue
                for match in re.finditer(pattern, text):
                    findings.append(
                        {
                            "page": slide_id,
                            "field": path,
                            "severity": severity,
                            "rule": message,
                            "match": match.group(0),
                            "suggestion": suggestion,
                            "context": text[max(0, match.start() - 12): match.end() + 12],
                        }
                    )

    errors = [f for f in findings if f["severity"] == "error"]
    warnings = [f for f in findings if f["severity"] == "warning"]
    result = {
        "error_count": len(errors),
        "warning_count": len(warnings),
        "findings": findings,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if args.out:
        Path(args.out).mkdir(parents=True, exist_ok=True)
        (Path(args.out) / "register_report.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    failed = errors or (args.strict and warnings)
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
