from __future__ import annotations

import json
import os
import time
from pathlib import Path

from quality_gate import build_report, render_markdown


def _write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _make_output(
    root: Path,
    *,
    slide_count: int = 2,
    validation_issues: list[dict] | None = None,
    validation_warning_count: int = 0,
    design_checks: list[dict] | None = None,
    design_score: float = 86,
    pptx: bool = True,
    readability_verdict: str = "defense_ready",
    section_slide: bool = False,
) -> dict:
    root.mkdir(parents=True, exist_ok=True)
    deck_path = root / "deck.json"
    theme_path = root / "theme.json"
    old = time.time() - 120

    slides = [
        {"id": "01", "layout": "cover", "title": "示例封面"},
        {"id": "02", "layout": "grid", "title": "示例内容", "section": "1"},
    ]
    if section_slide:
        slides.insert(1, {"id": "02", "layout": "section", "title": "第一章", "section": "1"})
        slides[-1]["id"] = "03"
    deck = {
        "meta": {
            "title": "示例答辩",
            "navigation_mode": "sidebar",
            "deck_mode": "full",
        },
        "slides": slides,
    }
    _write_json(deck_path, deck)
    _write_json(
        theme_path,
        {
            "name": "test_theme",
            "colors": {"blue": "#123456"},
            "layout": {"sidebar_mode": "always"},
            "brand": {},
        },
    )
    os.utime(deck_path, (old, old))
    os.utime(theme_path, (old, old))

    png_dir = root / "rendered"
    png_dir.mkdir(exist_ok=True)
    for index in range(1, slide_count + 1):
        (png_dir / f"slide_{index:02d}.png").write_bytes(b"png")
    (root / "contact_sheet.jpg").write_bytes(b"jpg")

    _write_json(
        root / "render_report.json",
        {
            "theme": "test_theme",
            "slide_count": slide_count,
            "rendered": [str(png_dir / f"slide_{index:02d}.png") for index in range(1, slide_count + 1)],
            "warnings": [],
        },
    )
    _write_json(
        root / "validation_report.json",
        {
            "deck": str(deck_path),
            "theme": "test_theme",
            "type_profile": "defense",
            "navigation_mode": "sidebar",
            "slide_count": slide_count,
            "rendered_count": slide_count,
            "error_count": sum(1 for issue in validation_issues or [] if issue.get("severity") == "error"),
            "warning_count": validation_warning_count,
            "issues": validation_issues or [],
        },
    )
    _write_json(
        root / "design_report.json",
        {
            "score": design_score,
            "grade": "A｜可以直接拿去讲" if design_score >= 85 else "B｜结构清楚了，还差气口",
            "checks": design_checks or [],
            "deck": str(deck_path),
        },
    )

    pptx_path = root / "defense_deck.pptx"
    if pptx:
        pptx_path.write_bytes(b"pptx")
        os.utime(pptx_path, (time.time() - 5, time.time() - 5))
        _write_json(
            root / "pptx_validation_report.json",
            {
                "pptx": str(pptx_path),
                "deck": str(deck_path),
                "slide_count": slide_count,
                "error_count": 0,
                "warning_count": 0,
                "issues": [],
            },
        )
        _write_json(
            root / "accessibility_report.json",
            {
                "pptx": str(pptx_path),
                "slide_count": slide_count,
                "error_count": 0,
                "warning_count": 0,
                "issues": [],
            },
        )
        _write_json(
            root / "pptx_readability.json",
            {
                "pptx": str(pptx_path),
                "slide_count": slide_count,
                "font_size_min_pt": 18.0,
                "runs_below_18pt_ratio": 0.0,
                "verdict": readability_verdict,
                "verdict_reasons": ["body median 26 pt"],
            },
        )
        _write_json(
            root / "pptx_preview" / "pptx_render_report.json",
            {
                "pptx": str(pptx_path),
                "renderer": "powerpoint-com",
                "status": "ok",
                "slide_count": slide_count,
                "blank_pages": [],
                "duplicate_pages": [],
                "comparisons": [
                    {"slide": index, "status": "compared", "mean_abs_diff": 12.0, "p95_abs_diff": 80.0}
                    for index in range(1, slide_count + 1)
                ],
                "rendered_files": [],
            },
        )
    _write_json(
        root / "speech_report.json",
        {
            "target_seconds": 600,
            "estimated_seconds": 570,
            "missing_speech": [],
        },
    )
    _write_json(root / "qa.json", {"question_count": 10, "questions": []})
    (root / "agent_brief.md").write_text("brief", encoding="utf-8")

    return {
        "deck": deck_path,
        "theme": theme_path,
        "out": root,
        "pptx": pptx_path,
    }


def _gate(paths: dict, **kwargs) -> dict:
    return build_report(
        deck_path=paths["deck"],
        theme_path=paths["theme"],
        out_dir=paths["out"],
        pptx_path=paths["pptx"],
        profile="defense",
        stage="final",
        **kwargs,
    )


def test_quality_gate_passes_a_complete_delivery(work_tmp: Path) -> None:
    report = _gate(_make_output(work_tmp))

    assert report["status"] == "pass"
    assert report["counts"]["blocking_checks"] == 0
    assert report["counts"]["blocking_issues"] == 0
    assert report["counts"]["advisory_issues"] == 0
    assert "PASS" in render_markdown(report)


def test_quality_gate_blocks_overflow_errors(work_tmp: Path) -> None:
    report = _gate(
        _make_output(
            work_tmp,
            validation_issues=[
                {
                    "severity": "error",
                    "page": "03",
                    "message": "Body overflowed the content box on page 03.",
                }
            ],
        )
    )

    assert report["status"] == "fail"
    assert report["counts"]["blocking_issues"] >= 1
    assert any("overflow" in issue["message"].lower() for issue in report["issues"] if issue["blocking"])
    assert any("Split the page" in action for action in report["next_actions"])


def test_quality_gate_keeps_chars_per_line_as_advisory(work_tmp: Path) -> None:
    report = _gate(
        _make_output(
            work_tmp,
            validation_issues=[
                {
                    "severity": "warning",
                    "page": "design",
                    "message": "Design: chars_per_line = 15.9 (target ≥ 18.0).",
                }
            ],
            validation_warning_count=1,
        )
    )

    assert report["status"] == "pass"
    assert report["counts"]["advisory_issues"] >= 1
    assert report["counts"]["blocking_issues"] == 0


def test_quality_gate_rejects_a_stale_validation_report(work_tmp: Path) -> None:
    paths = _make_output(work_tmp)
    stale = time.time() - 3600
    os.utime(paths["out"] / "validation_report.json", (stale, stale))

    report = _gate(paths)

    assert report["status"] == "fail"
    check = next(check for check in report["checks"] if check["id"] == "deck.validation")
    assert check["status"] == "fail"
    assert "stale" in check["detail"].lower()


def test_quality_gate_blocks_font_upgrade_verdict(work_tmp: Path) -> None:
    report = _gate(_make_output(work_tmp, readability_verdict="font_upgrade"))

    assert report["status"] == "fail"
    assert any("font_upgrade" in issue["message"] for issue in report["issues"] if issue["blocking"])
    assert any("defense-readability" in action for action in report["next_actions"])


def test_quality_gate_blocks_sidebar_with_section_pages(work_tmp: Path) -> None:
    report = _gate(_make_output(work_tmp, slide_count=3, section_slide=True))

    assert report["status"] == "fail"
    check = next(check for check in report["checks"] if check["id"] == "deck.navigation")
    assert check["status"] == "fail"
    assert "sidebar" in check["detail"]


def test_quality_gate_requires_pptx_for_final_stage(work_tmp: Path) -> None:
    report = _gate(_make_output(work_tmp, pptx=False))

    assert report["status"] == "fail"
    assert any(check["id"] == "pptx.exists" and check["status"] == "fail" for check in report["checks"])

