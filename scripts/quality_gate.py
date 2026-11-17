from __future__ import annotations

import argparse
import json
import re
from datetime import datetime
from pathlib import Path

from common import load_json, navigation_mode_for, write_json


STAGES = ("sample", "deck", "pptx", "final")
STAGE_ORDER = {name: index for index, name in enumerate(STAGES)}
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}

BLOCKING_WARNING = re.compile(
    r"(overflow|overlap|outside the slide|escaped|non-positive|"
    r"missing asset|not found|mismatch|wrapped across|"
    r"no non-empty title|missing semantic alt|no header-row|"
    r"only \d+/\d+ formulas|only \d+ editable text)",
    re.IGNORECASE,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Audit a rendered academic deck and emit one prioritized quality gate."
    )
    parser.add_argument("--deck", required=True)
    parser.add_argument("--theme", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--profile", choices=["document", "projection", "lecture", "defense"], default=None)
    parser.add_argument("--stage", choices=STAGES, default="final")
    parser.add_argument("--pptx", default=None, help="Defaults to <out>/defense_deck.pptx.")
    parser.add_argument("--min-design-score", type=float, default=80.0)
    parser.add_argument("--strict-warnings", action="store_true")
    parser.add_argument(
        "--allow-missing-preview",
        action="store_true",
        help="Downgrade a missing real-PowerPoint render to advisory for environments without Office.",
    )
    parser.add_argument("--require-speech", action="store_true")
    return parser.parse_args()


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _dedupe_paths(paths: list[Path]) -> list[Path]:
    seen: set[str] = set()
    result: list[Path] = []
    for path in paths:
        try:
            key = str(path.resolve()).lower()
        except OSError:
            continue
        if key in seen:
            continue
        seen.add(key)
        result.append(path)
    return result


def _walk_strings(value) -> list[str]:
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        result: list[str] = []
        for item in value:
            result.extend(_walk_strings(item))
        return result
    if isinstance(value, dict):
        result = []
        for item in value.values():
            result.extend(_walk_strings(item))
        return result
    return []


def _script_paths(*names: str) -> list[Path]:
    scripts = Path(__file__).resolve().parent
    return [scripts / name for name in names if (scripts / name).is_file()]


def _collect_inputs(deck_path: Path, theme_path: Path, deck: dict, theme: dict) -> dict[str, list[Path]]:
    deck_dir = deck_path.parent
    theme_dir = theme_path.parent

    images: list[Path] = []
    for value in _walk_strings(deck):
        candidate = Path(value)
        if candidate.suffix.lower() not in IMAGE_SUFFIXES:
            continue
        if not candidate.is_absolute():
            candidate = deck_dir / candidate
        if candidate.is_file():
            images.append(candidate)

    brand = theme.get("brand") or {}
    for key in ("lockup", "emblem"):
        value = str(brand.get(key) or "").strip()
        if not value:
            continue
        candidate = Path(value)
        if not candidate.is_absolute():
            candidate = theme_dir / candidate
        if candidate.is_file():
            images.append(candidate)

    base = _dedupe_paths([deck_path, theme_path, *images])
    render = _dedupe_paths(
        base
        + _script_paths("render_deck.py", "render_deck_next.py", "design.py", "common.py")
    )
    validation = _dedupe_paths(render + _script_paths("validate_deck.py"))
    design = _dedupe_paths(render + _script_paths("qa_design.py"))
    pptx = _dedupe_paths(render + _script_paths("export_editable_pptx.py"))
    return {
        "base": base,
        "render": render,
        "validation": validation,
        "design": design,
        "pptx": pptx,
    }


def _max_mtime(paths: list[Path]) -> float:
    values = []
    for path in paths:
        try:
            values.append(path.stat().st_mtime)
        except OSError:
            continue
    return max(values) if values else 0.0


def _is_fresh(path: Path, inputs: list[Path]) -> bool:
    if not path.is_file():
        return False
    return path.stat().st_mtime + 0.001 >= _max_mtime(inputs)


def _mtime_label(path: Path) -> str:
    try:
        return datetime.fromtimestamp(path.stat().st_mtime).astimezone().isoformat(timespec="seconds")
    except OSError:
        return "missing"


def _same_path(value, target: Path) -> bool:
    if not value:
        return True
    try:
        return Path(str(value)).resolve() == target.resolve()
    except OSError:
        return False


def _shorten(value: str, limit: int = 180) -> str:
    text = " ".join(str(value).split())
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _pages_label(pages: list[str]) -> str:
    clean = [str(page) for page in pages if str(page) not in {"", "general"}]
    if not clean:
        return "general"
    if len(clean) <= 6:
        return ", ".join(clean)
    return ", ".join(clean[:5]) + f" … (+{len(clean) - 5})"


def _page_sort_key(page: str) -> tuple[int, str]:
    text = str(page)
    match = re.match(r"^(\d+)", text)
    if match:
        return int(match.group(1)), text
    return 999, text


def suggest_action(message: str, source: str) -> str:
    text = str(message).lower()
    if "chars_per_line" in text:
        return "Widen the prose column or drop one column; do not shrink type below the profile floor. See layout-catalog.md."
    if "header eyebrow" in text:
        return "Set type_scale.eyebrow equal to type_scale.header_title in the theme; rerender the whole deck."
    if "closing card" in text or "closing" in text and "overflow" in text:
        return "Cut each closing card to at most two short bullets, or use card-only closing; do not shrink the 39 px body."
    if "overflow" in text or "clipped" in text or "escaped" in text:
        return "Split the page or remove about 20% of the text; rerender only this page, then run the full gate."
    if "wrapped across" in text or "index text wrapped" in text:
        return "Keep 01/02-style index tokens on one line (wrap_explicit=False); do not reduce the index size."
    if "alt" in text or "decorative" in text:
        return "Add semantic alt text, or mark purely decorative chrome as decorative; rerun check_accessibility.py."
    if "header-row" in text or "table" in text:
        return "Use a native table with header-row semantics; inspect the table page in the PPTX render."
    if "formula" in text or "omml" in text:
        return "Keep the formula native OMML, then verify the symbol list and its column count; see editable-pptx.md."
    if "font_upgrade" in text or "below 24" in text or "24 pt" in text:
        return "Raise body text to 24 pt and split dense pages; tokenize labels and footnotes instead of shrinking them."
    if "readability" in text or "font" in text:
        return "Read defense-readability.md, split the named page, then rerun audit_pptx.py."
    if source == "qa_design":
        return "Read the failed metric in design_report.md; change composition, not decoration, then rerun qa_design.py."
    if source == "render_pptx_preview":
        return "Open the named PPTX preview page and fix the native layout; do not trust the PIL preview as delivery evidence."
    if source == "validate_pptx":
        return "Fix the native PPTX object on the named page, regenerate the PPTX, then rerun validate_pptx.py."
    if source == "accessibility":
        return "Fix the accessibility issue on the named slide, then rerun check_accessibility.py."
    return "Fix the issue on the named page, rerun the source check, then run quality_gate.py again."


def _add_check(
    checks: list[dict],
    *,
    check_id: str,
    label: str,
    status: str,
    detail: str,
    action: str = "",
    blocking: bool | None = None,
) -> None:
    if status not in {"pass", "warn", "fail"}:
        raise ValueError(f"Unsupported check status: {status}")
    if blocking is None:
        blocking = status == "fail"
    checks.append(
        {
            "id": check_id,
            "label": label,
            "status": status,
            "blocking": bool(blocking),
            "detail": detail,
            "action": action,
        }
    )


def _add_issue(
    issues: list[dict],
    *,
    page,
    source: str,
    severity: str,
    blocking: bool,
    message: str,
    action: str | None = None,
) -> None:
    issues.append(
        {
            "page": str(page if page not in (None, "") else "general"),
            "source": source,
            "severity": severity,
            "blocking": bool(blocking),
            "message": _shorten(message),
            "action": action or suggest_action(message, source),
        }
    )


def _command(*parts: str) -> str:
    quoted = []
    for part in parts:
        if " " in part and not (part.startswith('"') and part.endswith('"')):
            quoted.append(f'"{part}"')
        else:
            quoted.append(part)
    return " ".join(quoted)


def _load_optional_json(path: Path) -> tuple[dict | None, str | None]:
    if not path.is_file():
        return None, "missing"
    try:
        data = load_json(path)
    except Exception as exc:  # noqa: BLE001 - report the parse failure in the gate
        return None, f"{type(exc).__name__}: {exc}"
    return data if isinstance(data, dict) else None, None


def _check_render(
    *,
    checks: list[dict],
    issues: list[dict],
    out_dir: Path,
    deck_path: Path,
    theme_path: Path,
    profile: str,
    expected_slides: int,
    stage: str,
    inputs: dict[str, list[Path]],
) -> bool:
    render_dir = out_dir / "rendered"
    pages = sorted(render_dir.glob("slide_*.png"))
    render_command = _command(
        "python",
        "scripts/render_deck.py",
        "--deck",
        str(deck_path),
        "--theme",
        str(theme_path),
        "--out",
        str(out_dir),
        "--profile",
        profile,
    )
    required = 1 if stage == "sample" else expected_slides
    if stage == "sample":
        required = min(5, expected_slides)
    files_ok = True
    if not pages:
        files_ok = False
        _add_check(
            checks,
            check_id="render.files",
            label="Rendered PNG pages",
            status="fail",
            detail=f"No pages in {render_dir}.",
            action=render_command,
        )
    elif stage != "sample" and len(pages) != required:
        files_ok = False
        _add_check(
            checks,
            check_id="render.files",
            label="Rendered PNG pages",
            status="fail",
            detail=f"Found {len(pages)} pages, expected {required}.",
            action="Run a full render; partial --pages renders cannot be the delivery evidence.",
        )
    else:
        stale = [page for page in pages if page.stat().st_mtime + 0.001 < _max_mtime(inputs["render"])]
        if stage != "sample" and stale:
            files_ok = False
            _add_check(
                checks,
                check_id="render.files",
                label="Rendered PNG pages",
                status="fail",
                detail=f"{len(stale)} of {len(pages)} pages are older than the deck/theme/images.",
                action=render_command,
            )
        elif len(pages) < required:
            _add_check(
                checks,
                check_id="render.files",
                label="Rendered PNG pages",
                status="warn",
                detail=f"Sample has {len(pages)} pages; {required} are recommended.",
                action="Render cover, outline, the densest content page, a real-figure page, and closing.",
            )
        else:
            _add_check(
                checks,
                check_id="render.files",
                label="Rendered PNG pages",
                status="pass",
                detail=f"{len(pages)} page(s) present and fresh.",
            )

    report_path = out_dir / "render_report.json"
    report, error = _load_optional_json(report_path)
    if report is None:
        _add_check(
            checks,
            check_id="render.report",
            label="Render report",
            status="fail" if stage != "sample" else "warn",
            detail=f"{report_path.name}: {error}.",
            action=render_command,
        )
        return False if stage != "sample" else files_ok
    warnings = report.get("warnings") or []
    stale = not _is_fresh(report_path, inputs["render"])
    count = report.get("slide_count")
    mismatch = stage != "sample" and count != expected_slides
    report_ok = True
    if stale:
        report_ok = False
        _add_check(
            checks,
            check_id="render.report",
            label="Render report",
            status="fail" if stage != "sample" else "warn",
            detail=f"Report is older than the deck/theme/images ({_mtime_label(report_path)}).",
            action=render_command,
        )
    elif mismatch:
        report_ok = False
        _add_check(
            checks,
            check_id="render.report",
            label="Render report",
            status="fail",
            detail=f"Report says {count} pages, deck has {expected_slides}.",
            action=render_command,
        )
    elif warnings:
        report_ok = False
        _add_check(
            checks,
            check_id="render.report",
            label="Render report",
            status="fail",
            detail=f"{len(warnings)} render warning(s).",
            action="Fix every render warning on the named page; these are layout defects, not style notes.",
        )
        for warning in warnings:
            _add_issue(
                issues,
                page=warning.get("page", "general"),
                source="render",
                severity="error",
                blocking=True,
                message=str(warning.get("message", "")),
            )
    else:
        _add_check(
            checks,
            check_id="render.report",
            label="Render report",
            status="pass",
            detail="0 render warnings.",
        )
    return files_ok and report_ok


def _check_contact_sheet(
    *,
    checks: list[dict],
    out_dir: Path,
    stage: str,
    inputs: dict[str, list[Path]],
) -> None:
    path = out_dir / "contact_sheet.jpg"
    if not path.is_file():
        _add_check(
            checks,
            check_id="render.contact_sheet",
            label="Contact sheet",
            status="warn" if stage == "sample" else "fail",
            detail="contact_sheet.jpg is missing.",
            action="Run scripts/render_deck.py; inspect the contact sheet before every full-deck pass.",
        )
        return
    if stage != "sample" and not _is_fresh(path, inputs["render"]):
        _add_check(
            checks,
            check_id="render.contact_sheet",
            label="Contact sheet",
            status="fail",
            detail="contact_sheet.jpg is older than the deck.",
            action="Run a full render to refresh the contact sheet.",
        )
        return
    _add_check(
        checks,
        check_id="render.contact_sheet",
        label="Contact sheet",
        status="pass",
        detail="Present and current.",
    )


def _check_validation(
    *,
    checks: list[dict],
    issues: list[dict],
    out_dir: Path,
    deck_path: Path,
    theme_path: Path,
    expected_slides: int,
    nav_mode: str,
    profile: str,
    stage: str,
    strict_warnings: bool,
    inputs: dict[str, list[Path]],
) -> None:
    if stage == "sample":
        return
    path = out_dir / "validation_report.json"
    report, error = _load_optional_json(path)
    rerun = _command(
        "python",
        "scripts/validate_deck.py",
        "--deck",
        str(deck_path),
        "--theme",
        str(theme_path),
        "--render-dir",
        str(out_dir / "rendered"),
        "--out",
        str(out_dir),
        "--profile",
        profile,
    )
    if report is None:
        _add_check(
            checks,
            check_id="deck.validation",
            label="Deck validation",
            status="fail",
            detail=f"validation_report.json: {error}.",
            action=rerun,
        )
        return
    errors = int(report.get("error_count", 0) or 0)
    warnings = int(report.get("warning_count", 0) or 0)
    stale = not _is_fresh(path, inputs["validation"])
    wrong_deck = not _same_path(report.get("deck"), deck_path)
    count = report.get("slide_count")
    count_mismatch = count != expected_slides
    nav_mismatch = bool(report.get("navigation_mode")) and report.get("navigation_mode") != nav_mode
    blocking_warnings = [
        issue
        for issue in report.get("issues") or []
        if issue.get("severity") == "warning"
        and (strict_warnings or BLOCKING_WARNING.search(str(issue.get("message", ""))))
    ]
    if stale or wrong_deck or count_mismatch or nav_mismatch or errors or blocking_warnings:
        reasons = []
        if stale:
            reasons.append("report is stale")
        if wrong_deck:
            reasons.append("report belongs to another deck")
        if count_mismatch:
            reasons.append(f"report says {count} pages, deck has {expected_slides}")
        if nav_mismatch:
            reasons.append(f"navigation mode is {report.get('navigation_mode')}, expected {nav_mode}")
        if errors:
            reasons.append(f"{errors} error(s)")
        if blocking_warnings:
            reasons.append(f"{len(blocking_warnings)} blocking warning(s)")
        _add_check(
            checks,
            check_id="deck.validation",
            label="Deck validation",
            status="fail",
            detail="; ".join(reasons) + ".",
            action=rerun,
        )
    elif warnings:
        _add_check(
            checks,
            check_id="deck.validation",
            label="Deck validation",
            status="warn",
            detail=f"0 errors, {warnings} advisory warning(s).",
            action="Read the warnings; keep only the ones that are explicitly non-blocking and record them.",
        )
    else:
        _add_check(
            checks,
            check_id="deck.validation",
            label="Deck validation",
            status="pass",
            detail="0 errors, 0 warnings.",
        )
    for issue in report.get("issues") or []:
        severity = str(issue.get("severity", "warning"))
        blocking = severity == "error" or (
            severity == "warning" and (strict_warnings or bool(BLOCKING_WARNING.search(str(issue.get("message", "")))))
        )
        _add_issue(
            issues,
            page=issue.get("page", "general"),
            source="validate_deck",
            severity=severity,
            blocking=blocking,
            message=str(issue.get("message", "")),
        )


def _check_design(
    *,
    checks: list[dict],
    issues: list[dict],
    out_dir: Path,
    deck_path: Path,
    theme_path: Path,
    expected_slides: int,
    profile: str,
    stage: str,
    min_score: float,
    inputs: dict[str, list[Path]],
) -> None:
    if stage == "sample":
        return
    path = out_dir / "design_report.json"
    report, error = _load_optional_json(path)
    rerun = _command(
        "python",
        "scripts/qa_design.py",
        "--deck",
        str(deck_path),
        "--theme",
        str(theme_path),
        "--render-dir",
        str(out_dir / "rendered"),
        "--out",
        str(out_dir),
        "--profile",
        profile,
    )
    if report is None:
        _add_check(
            checks,
            check_id="deck.design",
            label="Design audit",
            status="fail",
            detail=f"design_report.json: {error}.",
            action=rerun,
        )
        return
    score = float(report.get("score", 0) or 0)
    stale = not _is_fresh(path, inputs["design"])
    wrong_deck = not _same_path(report.get("deck"), deck_path)
    failed = [check for check in report.get("checks") or [] if not check.get("ok")]
    if stale or wrong_deck or score < min_score:
        reasons = []
        if stale:
            reasons.append("report is stale")
        if wrong_deck:
            reasons.append("report belongs to another deck")
        if score < min_score:
            reasons.append(f"score {score:g} < {min_score:g}")
        _add_check(
            checks,
            check_id="deck.design",
            label="Design audit",
            status="fail",
            detail="; ".join(reasons) + ".",
            action=rerun + " Then fix the lowest metric, not the colors.",
        )
    elif failed:
        _add_check(
            checks,
            check_id="deck.design",
            label="Design audit",
            status="warn",
            detail=f"Score {score:g} ({report.get('grade', '')}); {len(failed)} metric(s) below target.",
            action="Read design_report.md; fix composition before decoration; do not game one metric.",
        )
    else:
        _add_check(
            checks,
            check_id="deck.design",
            label="Design audit",
            status="pass",
            detail=f"Score {score:g} ({report.get('grade', '')}).",
        )
    for check in failed:
        _add_issue(
            issues,
            page="design",
            source="qa_design",
            severity="advisory",
            blocking=False,
            message=f"{check.get('metric')} = {check.get('value')} (target {check.get('target')})",
            action=suggest_action(str(check.get("metric", "")), "qa_design"),
        )


def _check_navigation(
    *,
    checks: list[dict],
    deck: dict,
    nav_mode: str,
) -> None:
    slides = deck.get("slides") or []
    sections = [
        str(slide.get("section"))
        for slide in slides
        if slide.get("layout") not in {"cover", "outline", "section"} and slide.get("section")
    ]
    sections = list(dict.fromkeys(sections))
    section_pages = [slide for slide in slides if slide.get("layout") == "section"]
    if nav_mode == "sidebar" and section_pages:
        _add_check(
            checks,
            check_id="deck.navigation",
            label="Navigation contract",
            status="fail",
            detail=f"navigation_mode=sidebar but deck has {len(section_pages)} section page(s).",
            action="Remove layout:section pages or switch meta.navigation_mode to dividers/hybrid.",
        )
        return
    if nav_mode in {"dividers", "hybrid"} and len(sections) > 1 and not section_pages:
        _add_check(
            checks,
            check_id="deck.navigation",
            label="Navigation contract",
            status="warn",
            detail=f"navigation_mode={nav_mode} with {len(sections)} sections but no section pages.",
            action="Run scripts/add_section_pages.py --navigation-mode <mode>, or switch to sidebar.",
        )
        return
    _add_check(
        checks,
        check_id="deck.navigation",
        label="Navigation contract",
        status="pass",
        detail=f"{nav_mode}; {len(section_pages)} section page(s), {len(sections)} section(s).",
    )


def _check_pptx(
    *,
    checks: list[dict],
    issues: list[dict],
    out_dir: Path,
    pptx_path: Path,
    deck_path: Path,
    theme_path: Path,
    expected_slides: int,
    profile: str,
    strict_warnings: bool,
    inputs: dict[str, list[Path]],
) -> bool:
    if not pptx_path.is_file():
        _add_check(
            checks,
            check_id="pptx.exists",
            label="Native PPTX",
            status="fail",
            detail=f"{pptx_path.name} is missing.",
            action=_command(
                "python",
                "scripts/export_editable_pptx.py",
                "--deck",
                str(deck_path),
                "--theme",
                str(theme_path),
                "--out",
                str(out_dir),
                "--pptx-name",
                pptx_path.name,
                "--profile",
                profile,
            ),
        )
        return False
    if pptx_path.stat().st_mtime + 0.001 < _max_mtime(inputs["pptx"]):
        _add_check(
            checks,
            check_id="pptx.exists",
            label="Native PPTX",
            status="fail",
            detail="PPTX is older than the deck/theme/images or the exporter.",
            action="Regenerate the editable PPTX before validating it.",
        )
        return False
    else:
        _add_check(
            checks,
            check_id="pptx.exists",
            label="Native PPTX",
            status="pass",
            detail=f"{pptx_path.name} present and newer than inputs.",
        )

    path = out_dir / "pptx_validation_report.json"
    report, error = _load_optional_json(path)
    rerun = _command(
        "python",
        "scripts/validate_pptx.py",
        "--pptx",
        str(pptx_path),
        "--deck",
        str(deck_path),
        "--out",
        str(out_dir),
    )
    if report is None:
        _add_check(
            checks,
            check_id="pptx.validation",
            label="Native PPTX validation",
            status="fail",
            detail=f"pptx_validation_report.json: {error}.",
            action=rerun,
        )
    else:
        errors = int(report.get("error_count", 0) or 0)
        warnings = int(report.get("warning_count", 0) or 0)
        stale = not _is_fresh(path, [pptx_path, *_script_paths("validate_pptx.py")])
        wrong_deck = not _same_path(report.get("deck"), deck_path)
        wrong_pptx = not _same_path(report.get("pptx"), pptx_path)
        count = report.get("slide_count")
        if errors or stale or wrong_deck or wrong_pptx or count != expected_slides:
            reasons = []
            if errors:
                reasons.append(f"{errors} error(s)")
            if stale:
                reasons.append("report is older than the PPTX")
            if wrong_deck or wrong_pptx:
                reasons.append("report belongs to another file")
            if count != expected_slides:
                reasons.append(f"report says {count} slides, deck has {expected_slides}")
            _add_check(
                checks,
                check_id="pptx.validation",
                label="Native PPTX validation",
                status="fail",
                detail="; ".join(reasons) + ".",
                action=rerun,
            )
        elif warnings:
            _add_check(
                checks,
                check_id="pptx.validation",
                label="Native PPTX validation",
                status="warn",
                detail=f"0 errors, {warnings} advisory warning(s).",
                action="Read the warnings; fix semantic issues even when they do not block.",
            )
        else:
            _add_check(
                checks,
                check_id="pptx.validation",
                label="Native PPTX validation",
                status="pass",
                detail="0 errors, 0 warnings.",
            )
        for issue in report.get("issues") or []:
            severity = str(issue.get("severity", "warning"))
            blocking = severity == "error" or (severity == "warning" and strict_warnings)
            _add_issue(
                issues,
                page=issue.get("page", "pptx"),
                source="validate_pptx",
                severity=severity,
                blocking=blocking,
                message=str(issue.get("message", "")),
            )

    access_path = out_dir / "accessibility_report.json"
    access, access_error = _load_optional_json(access_path)
    if access is None:
        _add_check(
            checks,
            check_id="pptx.accessibility",
            label="Accessibility",
            status="fail",
            detail=f"accessibility_report.json: {access_error}.",
            action=_command(
                "python",
                "scripts/check_accessibility.py",
                "--pptx",
                str(pptx_path),
                "--out",
                str(out_dir),
            ),
        )
    else:
        errors = int(access.get("error_count", 0) or 0)
        warnings = int(access.get("warning_count", 0) or 0)
        stale = not _is_fresh(access_path, [pptx_path, *_script_paths("check_accessibility.py")])
        if errors or stale:
            _add_check(
                checks,
                check_id="pptx.accessibility",
                label="Accessibility",
                status="fail",
                detail=f"{errors} error(s); report stale={stale}.",
                action="Fix title placeholders, alt text, table headers and reading order; rerun check_accessibility.py.",
            )
        elif warnings:
            _add_check(
                checks,
                check_id="pptx.accessibility",
                label="Accessibility",
                status="warn",
                detail=f"0 errors, {warnings} warning(s).",
                action="Fix semantic warnings when cheap; record the rest as accepted advisories.",
            )
        else:
            _add_check(
                checks,
                check_id="pptx.accessibility",
                label="Accessibility",
                status="pass",
                detail="0 errors, 0 warnings.",
            )
        for issue in access.get("issues") or []:
            severity = str(issue.get("severity", "warning"))
            _add_issue(
                issues,
                page=issue.get("slide", "general"),
                source="accessibility",
                severity=severity,
                blocking=severity == "error",
                message=str(issue.get("message", "")),
            )

    read_path = out_dir / "pptx_readability.json"
    read, read_error = _load_optional_json(read_path)
    if read is None:
        _add_check(
            checks,
            check_id="pptx.readability",
            label="Defense readability",
            status="fail" if profile in {"defense", "lecture"} else "warn",
            detail=f"pptx_readability.json: {read_error}.",
            action=_command(
                "python",
                "scripts/audit_pptx.py",
                "--pptx",
                str(pptx_path),
                "--out",
                str(out_dir),
            ),
        )
    else:
        stale = not _is_fresh(read_path, [pptx_path, *_script_paths("audit_pptx.py")])
        verdict = str(read.get("verdict", "")).strip()
        min_pt = float(read.get("font_size_min_pt", 0) or 0)
        below_ratio = float(read.get("runs_below_18pt_ratio", 0) or 0)
        if profile == "defense":
            ok = verdict == "defense_ready" and not stale
            detail = f"verdict={verdict or 'missing'}; min={min_pt:g} pt; below18={below_ratio:.0%}; stale={stale}."
        elif profile == "lecture":
            ok = not stale and min_pt >= 18 and below_ratio <= 0.05
            detail = f"verdict={verdict or 'n/a'}; min={min_pt:g} pt; below18={below_ratio:.0%}; stale={stale}."
        else:
            ok = not stale and below_ratio <= 0.05 and min_pt >= 10
            detail = f"verdict={verdict or 'n/a'}; min={min_pt:g} pt; below18={below_ratio:.0%}; stale={stale}."
        if ok:
            _add_check(
                checks,
                check_id="pptx.readability",
                label="Defense readability",
                status="pass",
                detail=detail,
            )
        else:
            _add_check(
                checks,
                check_id="pptx.readability",
                label="Defense readability",
                status="fail" if profile in {"defense", "lecture"} else "warn",
                detail=detail,
                action="Read defense-readability.md; split dense pages and raise body text without shrinking labels further.",
            )
        if profile == "defense" and verdict and verdict != "defense_ready":
            _add_issue(
                issues,
                page="pptx",
                source="audit_pptx",
                severity="error",
                blocking=True,
                message=f"Readability verdict is {verdict}: " + "; ".join(read.get("verdict_reasons") or []),
            )
    return True


def _check_preview(
    *,
    checks: list[dict],
    issues: list[dict],
    out_dir: Path,
    pptx_path: Path,
    expected_slides: int,
    allow_missing: bool,
) -> None:
    path = out_dir / "pptx_preview" / "pptx_render_report.json"
    report, error = _load_optional_json(path)
    rerun = _command(
        "python",
        "scripts/render_pptx_preview.py",
        "--pptx",
        str(pptx_path),
        "--out",
        str(out_dir / "pptx_preview"),
        "--compare-dir",
        str(out_dir / "rendered"),
        "--strict",
    )
    if report is None:
        _add_check(
            checks,
            check_id="pptx.preview",
            label="Real PowerPoint render",
            status="warn" if allow_missing else "fail",
            detail=f"pptx_render_report.json: {error}. {'' if allow_missing else 'Real render evidence is required.'}",
            action=rerun if not allow_missing else "Install PowerPoint/LibreOffice and rerun before claiming visual fidelity.",
        )
        return
    status = str(report.get("status", "")).strip()
    count = report.get("slide_count")
    blank = report.get("blank_pages") or []
    duplicate = report.get("duplicate_pages") or []
    stale = not _is_fresh(path, [pptx_path, *_script_paths("render_pptx_preview.py")])
    if status != "ok" or count != expected_slides or blank or duplicate or stale:
        reasons = []
        if status != "ok":
            reasons.append(f"status={status or 'missing'}")
        if count != expected_slides:
            reasons.append(f"{count} pages vs {expected_slides}")
        if blank:
            reasons.append(f"blank pages {blank}")
        if duplicate:
            reasons.append(f"duplicate pages {duplicate}")
        if stale:
            reasons.append("report is older than the PPTX")
        _add_check(
            checks,
            check_id="pptx.preview",
            label="Real PowerPoint render",
            status="warn" if allow_missing and status != "ok" else "fail",
            detail="; ".join(reasons) + ".",
            action=rerun,
        )
    else:
        _add_check(
            checks,
            check_id="pptx.preview",
            label="Real PowerPoint render",
            status="pass",
            detail=f"{count}/{expected_slides} pages rendered; no blanks or duplicates.",
        )
    for page in blank:
        _add_issue(
            issues,
            page=page,
            source="render_pptx_preview",
            severity="error",
            blocking=True,
            message="PowerPoint rendered a blank page.",
        )
    for pages in duplicate:
        _add_issue(
            issues,
            page=", ".join(str(page) for page in pages),
            source="render_pptx_preview",
            severity="error",
            blocking=True,
            message="Duplicate rendered pages.",
        )
    for comparison in report.get("comparisons") or []:
        if float(comparison.get("p95_abs_diff", 0) or 0) > 220:
            _add_issue(
                issues,
                page=comparison.get("slide", "pptx"),
                source="render_pptx_preview",
                severity="advisory",
                blocking=False,
                message=f"PPTX/PIL preview differs at p95={comparison.get('p95_abs_diff')}.",
                action="Inspect this page in both renders; large differences usually mean font substitution or a geometry drift.",
            )


def _check_speech(
    *,
    checks: list[dict],
    issues: list[dict],
    out_dir: Path,
    require_speech: bool,
) -> None:
    path = out_dir / "speech_report.json"
    report, error = _load_optional_json(path)
    if report is None:
        _add_check(
            checks,
            check_id="delivery.speech",
            label="Speech timing",
            status="fail" if require_speech else "warn",
            detail=f"speech_report.json: {error}.",
            action="Run scripts/build_speech.py; record the duration and any A/B speech-audit hits.",
        )
        return
    missing = report.get("missing_speech") or []
    target = float(report.get("target_seconds", 0) or 0)
    estimated = float(report.get("estimated_seconds", 0) or 0)
    ratio = estimated / target if target else 0.0
    if missing:
        _add_check(
            checks,
            check_id="delivery.speech",
            label="Speech timing",
            status="fail" if require_speech else "warn",
            detail=f"{len(missing)} slide(s) missing speech.",
            action="Fill every slide's speech before delivery; run build_speech.py again.",
        )
        _add_issue(
            issues,
            page=", ".join(str(page) for page in missing[:8]),
            source="speech",
            severity="warning",
            blocking=require_speech,
            message="Slides missing speech text.",
        )
        return
    if target and not 0.8 <= ratio <= 1.2:
        _add_check(
            checks,
            check_id="delivery.speech",
            label="Speech timing",
            status="fail" if require_speech else "warn",
            detail=f"Estimated {estimated:.0f}s against target {target:.0f}s ({ratio:.0%}).",
            action="Cut or expand the spoken script, not the slides; rerun build_speech.py.",
        )
        return
    _add_check(
        checks,
        check_id="delivery.speech",
        label="Speech timing",
        status="pass",
        detail=f"Estimated {estimated:.0f}s against target {target:.0f}s ({ratio:.0%}).",
    )


def _check_delivery_files(
    *,
    checks: list[dict],
    out_dir: Path,
    deck_path: Path,
    theme_path: Path,
    inputs: dict[str, list[Path]],
) -> None:
    brief = out_dir / "agent_brief.md"
    if not brief.is_file():
        _add_check(
            checks,
            check_id="delivery.agent_brief",
            label="Portable agent brief",
            status="warn",
            detail="agent_brief.md is missing.",
            action=_command(
                "python",
                "scripts/export_agent_brief.py",
                "--deck",
                str(deck_path),
                "--theme",
                str(theme_path),
                "--out",
                str(brief),
            ),
        )
    elif not _is_fresh(brief, inputs["base"]):
        _add_check(
            checks,
            check_id="delivery.agent_brief",
            label="Portable agent brief",
            status="warn",
            detail="agent_brief.md is older than the deck/theme.",
            action="Regenerate agent_brief.md so a non-script model receives the current contract.",
        )
    else:
        _add_check(
            checks,
            check_id="delivery.agent_brief",
            label="Portable agent brief",
            status="pass",
            detail="Present and current.",
        )

    qa = out_dir / "qa.json"
    report, error = _load_optional_json(qa)
    if report is None:
        _add_check(
            checks,
            check_id="delivery.qa",
            label="Q&A forecast",
            status="warn",
            detail=f"qa.json: {error}.",
            action="Run scripts/build_qa.py --count 8-10 for a defense deck.",
        )
        return
    count = int(report.get("question_count", 0) or 0)
    if count < 8:
        _add_check(
            checks,
            check_id="delivery.qa",
            label="Q&A forecast",
            status="warn",
            detail=f"Only {count} questions.",
            action="Generate at least 8 traceable questions for a defense; do not invent facts.",
        )
    else:
        _add_check(
            checks,
            check_id="delivery.qa",
            label="Q&A forecast",
            status="pass",
            detail=f"{count} questions.",
        )


def build_report(
    *,
    deck_path: str | Path,
    theme_path: str | Path,
    out_dir: str | Path,
    profile: str | None = None,
    stage: str = "final",
    pptx_path: str | Path | None = None,
    min_design_score: float = 80.0,
    strict_warnings: bool = False,
    allow_missing_preview: bool = False,
    require_speech: bool = False,
) -> dict:
    deck_path = Path(deck_path).resolve()
    theme_path = Path(theme_path).resolve()
    out_dir = Path(out_dir).resolve()
    pptx_path = Path(pptx_path).resolve() if pptx_path else out_dir / "defense_deck.pptx"
    checks: list[dict] = []
    issues: list[dict] = []

    try:
        deck = load_json(deck_path)
    except Exception as exc:  # noqa: BLE001 - the gate must report an unreadable deck
        deck = {}
        _add_check(
            checks,
            check_id="deck.parse",
            label="Deck JSON",
            status="fail",
            detail=f"{type(exc).__name__}: {exc}",
            action="Fix deck.json before running any visual check.",
        )
    try:
        theme = load_json(theme_path)
    except Exception as exc:  # noqa: BLE001
        theme = {}
        _add_check(
            checks,
            check_id="theme.parse",
            label="Theme JSON",
            status="fail",
            detail=f"{type(exc).__name__}: {exc}",
            action="Fix theme.json before running any visual check.",
        )

    slides = deck.get("slides") or []
    expected_slides = len(slides)
    if not checks:
        _add_check(
            checks,
            check_id="deck.parse",
            label="Deck JSON",
            status="pass" if expected_slides else "fail",
            detail=f"{expected_slides} slide(s)." if expected_slides else "Deck has no slides.",
            action="" if expected_slides else "Add cover/outline/content slides to deck.json.",
        )
        _add_check(
            checks,
            check_id="theme.parse",
            label="Theme JSON",
            status="pass",
            detail=f"Theme '{theme.get('name', 'unnamed')}' loaded.",
        )

    if stage not in STAGE_ORDER:
        raise ValueError(f"Unknown stage: {stage}")
    if not profile:
        validation, _ = _load_optional_json(out_dir / "validation_report.json")
        profile = str((validation or {}).get("type_profile") or "document")
    if profile not in {"document", "projection", "lecture", "defense"}:
        profile = "document"

    nav_mode = navigation_mode_for(deck, theme) if deck else "sidebar"
    inputs = _collect_inputs(deck_path, theme_path, deck, theme)

    render_ready = _check_render(
        checks=checks,
        issues=issues,
        out_dir=out_dir,
        deck_path=deck_path,
        theme_path=theme_path,
        profile=profile,
        expected_slides=expected_slides,
        stage=stage,
        inputs=inputs,
    )
    _check_contact_sheet(checks=checks, out_dir=out_dir, stage=stage, inputs=inputs)
    _check_navigation(checks=checks, deck=deck, nav_mode=nav_mode)
    if render_ready:
        _check_validation(
            checks=checks,
            issues=issues,
            out_dir=out_dir,
            deck_path=deck_path,
            theme_path=theme_path,
            expected_slides=expected_slides,
            nav_mode=nav_mode,
            profile=profile,
            stage=stage,
            strict_warnings=strict_warnings,
            inputs=inputs,
        )
        _check_design(
            checks=checks,
            issues=issues,
            out_dir=out_dir,
            deck_path=deck_path,
            theme_path=theme_path,
            expected_slides=expected_slides,
            profile=profile,
            stage=stage,
            min_score=min_design_score,
            inputs=inputs,
        )
    elif stage != "sample":
        _add_check(
            checks,
            check_id="pipeline.evidence",
            label="Evidence chain",
            status="warn",
            detail="Validation, design, PPTX and preview checks were skipped until the rendered pages are refreshed.",
            action="Run a full render with the selected profile, then rerun quality_gate.py.",
        )

    pptx_ready = False
    if render_ready and STAGE_ORDER[stage] >= STAGE_ORDER["pptx"]:
        pptx_ready = _check_pptx(
            checks=checks,
            issues=issues,
            out_dir=out_dir,
            pptx_path=pptx_path,
            deck_path=deck_path,
            theme_path=theme_path,
            expected_slides=expected_slides,
            profile=profile,
            strict_warnings=strict_warnings,
            inputs=inputs,
        )
    if stage == "final":
        if pptx_ready:
            _check_preview(
                checks=checks,
                issues=issues,
                out_dir=out_dir,
                pptx_path=pptx_path,
                expected_slides=expected_slides,
                allow_missing=allow_missing_preview,
            )
        else:
            _add_check(
                checks,
                check_id="pptx.preview",
                label="Real PowerPoint render",
                status="warn",
                detail="Skipped because the native PPTX is missing or older than its inputs.",
                action="Regenerate the PPTX, run validate_pptx.py, then render the real PowerPoint preview.",
            )
        _check_speech(
            checks=checks,
            issues=issues,
            out_dir=out_dir,
            require_speech=require_speech,
        )
        _check_delivery_files(
            checks=checks,
            out_dir=out_dir,
            deck_path=deck_path,
            theme_path=theme_path,
            inputs=inputs,
        )

    issues = _dedupe_issues(issues)
    blocking_checks = [check for check in checks if check["blocking"]]
    advisory_checks = [check for check in checks if check["status"] == "warn"]
    passed_checks = [check for check in checks if check["status"] == "pass"]
    blocking_issues = [issue for issue in issues if issue["blocking"]]
    advisory_issues = [issue for issue in issues if not issue["blocking"]]
    status = "fail" if blocking_checks or blocking_issues else "pass"

    next_actions: list[str] = []
    for check in blocking_checks:
        if check.get("action"):
            next_actions.append(check["action"])
    for issue in blocking_issues:
        if issue.get("action"):
            next_actions.append(issue["action"])
    for issue in advisory_issues[:2]:
        if issue.get("action"):
            next_actions.append(issue["action"])
    next_actions = list(dict.fromkeys(next_actions))[:8]
    if status == "pass" and not next_actions:
        next_actions = ["No blocking issues. Review the advisory list, record accepted items, then deliver."]

    return {
        "schema_version": 1,
        "generated_at": _now(),
        "stage": stage,
        "profile": profile,
        "status": status,
        "deck": str(deck_path),
        "theme": str(theme_path),
        "out": str(out_dir),
        "pptx": str(pptx_path),
        "navigation_mode": nav_mode,
        "slide_count": expected_slides,
        "counts": {
            "blocking_checks": len(blocking_checks),
            "advisory_checks": len(advisory_checks),
            "passed_checks": len(passed_checks),
            "blocking_issues": len(blocking_issues),
            "advisory_issues": len(advisory_issues),
        },
        "checks": checks,
        "issues": sorted(
            issues,
            key=lambda issue: (not issue["blocking"], issue["source"], _page_sort_key(issue["page"])),
        ),
        "next_actions": next_actions,
        "rule": "Do not claim the deck is delivered while status is fail; two consecutive passes with no blocking issue are the stop condition.",
    }


def _dedupe_issues(issues: list[dict]) -> list[dict]:
    merged: dict[tuple[str, str], dict] = {}
    for issue in issues:
        message = str(issue.get("message", "")).replace("Renderer warning: ", "").strip()
        key = (str(issue.get("page", "general")), message.lower())
        existing = merged.get(key)
        if existing is None:
            copy = dict(issue)
            copy["message"] = message
            merged[key] = copy
            continue
        existing["blocking"] = bool(existing["blocking"] or issue["blocking"])
        if existing["severity"] == "advisory" and issue["severity"] != "advisory":
            existing["severity"] = issue["severity"]
        if not existing.get("action") and issue.get("action"):
            existing["action"] = issue["action"]
    return list(merged.values())


def _group_issues(issues: list[dict]) -> list[dict]:
    groups: dict[tuple[str, str, bool], dict] = {}
    for issue in issues:
        key = (str(issue["source"]), str(issue["message"]), bool(issue["blocking"]))
        group = groups.setdefault(
            key,
            {
                "source": issue["source"],
                "message": issue["message"],
                "blocking": issue["blocking"],
                "severity": issue["severity"],
                "action": issue.get("action", ""),
                "pages": [],
            },
        )
        group["pages"].append(str(issue["page"]))
        if issue["severity"] == "error":
            group["severity"] = "error"
    result = list(groups.values())
    for group in result:
        group["pages"] = sorted(set(group["pages"]), key=_page_sort_key)
    result.sort(key=lambda group: (not group["blocking"], group["source"], _page_sort_key(group["pages"][0]) if group["pages"] else (999, "")))
    return result


def render_markdown(report: dict) -> str:
    status = report["status"].upper()
    counts = report["counts"]
    lines = [
        f"# Quality Gate — {status}",
        "",
        f"- Stage: `{report['stage']}`",
        f"- Profile: `{report['profile']}`",
        f"- Deck: `{report['deck']}`",
        f"- PPTX: `{report['pptx']}`",
        f"- Slides: {report['slide_count']}",
        f"- Navigation: `{report['navigation_mode']}`",
        f"- Blocking checks: {counts['blocking_checks']}; blocking issues: {counts['blocking_issues']}",
        f"- Advisory checks: {counts['advisory_checks']}; advisory issues: {counts['advisory_issues']}",
        "",
        "> " + report["rule"],
        "",
        "## Checks",
        "",
        "| Check | Status | Detail |",
        "|---|---|---|",
    ]
    for check in report["checks"]:
        status_icon = {"pass": "PASS", "warn": "WARN", "fail": "FAIL"}[check["status"]]
        lines.append(
            f"| `{check['id']}` | {status_icon} | {_shorten(check['detail'], 180).replace('|', '/')} |"
        )

    groups = _group_issues(report["issues"])
    if groups:
        lines.extend(
            [
                "",
                "## Fix queue",
                "",
                "Only blocking rows must be fixed before delivery. Advisory rows are recorded, not silently ignored.",
                "",
                "| Blocking | Page(s) | Source | Issue | First action |",
                "|---|---|---|---|---|",
            ]
        )
        for group in groups[:24]:
            mark = "BLOCK" if group["blocking"] else "advisory"
            lines.append(
                "| {mark} | {pages} | `{source}` | {message} | {action} |".format(
                    mark=mark,
                    pages=_pages_label(group["pages"]),
                    source=group["source"],
                    message=_shorten(group["message"], 140).replace("|", "/"),
                    action=_shorten(group["action"], 150).replace("|", "/"),
                )
            )
        if len(groups) > 24:
            lines.append(f"| … | | | {len(groups) - 24} more groups | Read `quality_gate.json` |")

    lines.extend(["", "## Next actions", ""])
    for index, action in enumerate(report["next_actions"], start=1):
        lines.append(f"{index}. {action}")

    lines.extend(
        [
            "",
            "## Freshness rule",
            "",
            "A report older than the deck, theme, referenced image, or its generator is not evidence. "
            "Rerun the named command; do not reuse a previous pass.",
            "",
            "## Stop rule",
            "",
            "- PASS with blocking = 0 for two consecutive runs, and no new warning: stop.",
            "- Two rounds without measurable improvement: stop and record non-blocking issues.",
            "- Three visual passes on the same stage: return to the content contract; the problem is structural.",
        ]
    )
    return "\n".join(lines) + "\n"


def main() -> None:
    args = parse_args()
    report = build_report(
        deck_path=args.deck,
        theme_path=args.theme,
        out_dir=args.out,
        profile=args.profile,
        stage=args.stage,
        pptx_path=args.pptx,
        min_design_score=args.min_design_score,
        strict_warnings=args.strict_warnings,
        allow_missing_preview=args.allow_missing_preview,
        require_speech=args.require_speech,
    )
    out_dir = Path(args.out).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    json_path = out_dir / "quality_gate.json"
    md_path = out_dir / "quality_gate.md"
    write_json(json_path, report)
    md_path.write_text(render_markdown(report), encoding="utf-8")
    counts = report["counts"]
    print(
        json.dumps(
            {
                "status": report["status"],
                "stage": report["stage"],
                "profile": report["profile"],
                "blocking_checks": counts["blocking_checks"],
                "advisory_checks": counts["advisory_checks"],
                "blocking_issues": counts["blocking_issues"],
                "report": str(md_path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    for action in report["next_actions"][:3]:
        print(f"- {action}")
    if report["status"] == "fail":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
