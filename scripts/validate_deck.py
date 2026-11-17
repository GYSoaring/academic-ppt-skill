from __future__ import annotations
import re

import argparse
import json
from pathlib import Path

import numpy as np
from PIL import Image

from common import (
    ARXIV_RE,
    DOI_RE,
    DEFENSE_BODY_PX,
    DEFENSE_MIN_ANY_PX,
    NAVIGATION_MODES,
    apply_navigation_mode,
    apply_type_profile,
    estimate_seconds,
    is_assertion_title,
    load_json,
    load_theme,
    navigation_mode_for,
    normalize_cites,
    normalize_doi,
    resolve_path,
    write_json,
)


DESIGN_HINTS = {
    "dominant_object_ratio": "Give more pages one dominant object (a real figure, chart, formula or schematic) instead of equal cards.",
    "card_only_ratio": "Too many pages are only rows of equal-weight cards; that is the shape readers call a template.",
    "emphasis_per_page": "Mark the decisive number or boundary with **text**; a deck with no accent has no focus.",
    "assertion_ratio": "Turn topic-label titles into claims (see is_assertion_title in common.py).",
    "takeaway_ratio": "The takeaway band is appearing on almost every page; keep it for real turns of argument.",
    "layout_run_max": "The same layout repeats too many pages in a row; break the rhythm.",
    "image_page_ratio": "Source at least one real figure; a deck of pure text and boxes reads as generated.",
    "chars_per_line": "Text columns are too narrow for Chinese prose (under 18 characters per line reads choppy).",
    "rhythm_stdev": "Every page carries the same visual weight; alternate dense evidence pages with quiet ones.",
    "chrome_share": "Fixed chrome eats too much of the page; trim the header band, lead band or mini directory.",
}


SUPPORTED_LAYOUTS = {
    "cover",
    "outline",
    "section",
    "bullets",
    "figure",
    "figure_flow",
    "concept",
    "formula",
    "pipeline",
    "compare",
    "metrics",
    "closing",
    "grid",
    "balance",
    "hub",
    "timeline",
    "references",
    "quote",
}

REQUIRES_IMAGE_OR_CHART = {"figure", "figure_flow", "metrics", "balance"}
DIAGRAM_KINDS = {"sequence", "branch", "stack", "ruler", "matrix", "cycle", "relations"}
RELATION_NODE_TONES = {"pale", "outline", "solid"}
RELATION_ARROW_STYLES = {"ribbon", "block", "chevron", "line", "double"}
RELATION_ROUTES = {"straight", "orthogonal"}
RELATION_SIDES = {"auto", "left", "right", "top", "bottom"}
RELATION_TONES = {"primary", "muted", "alert"}
IMAGE_ORIGINS = {"own", "generated", "template", "borrowed", "public"}
SOURCE_POLICIES = {"off", "advisory", "strict"}
GENRES = {"general", "experiment", "systems", "theory", "empirical", "survey", "design", "humanities", "clinical"}
NARRATIVES = {
    "problem-method-result",
    "and-but-therefore",
    "result-first",
    "close-reading",
    "taxonomy",
    "design-rationale",
    "chronological",
}
LAYOUT_VARIANTS = {
    "figure": {"image_left", "image_right", "chart"},
    "quote": {"centered", "split"},
}


def title_units(text: str) -> int:
    return sum(2 if ord(character) > 0x2E80 else 1 for character in str(text or ""))


def summary_title_valid(text: str) -> bool:
    """Accept concise topic labels, not only six-character noun phrases."""
    length = len(re.sub(r"\s", "", str(text or "")))
    return 4 <= length <= 18


def closing_has_summary_content(slide: dict) -> bool:
    """A closing page can summarise with points or with boundary/next cards."""
    return bool(
        slide.get("points")
        or slide.get("boundary")
        or slide.get("boundary_points")
        or slide.get("next")
        or slide.get("next_points")
    )


def longest_content_layout_run(slides: list[dict]) -> tuple[int, str]:
    """Return the longest visible layout run.

    Section dividers break the visual rhythm, so the validator must not delete
    them before counting. Otherwise two figure pairs on either side of a divider
    look like four consecutive figure pages.
    """
    longest = 0
    longest_layout = ""
    current = 0
    previous = ""
    for slide in slides:
        layout = str(slide.get("layout", "")).strip().lower()
        if layout in {"cover", "outline", "section", "closing", "references"}:
            current = 0
            previous = ""
            continue
        current = current + 1 if layout == previous else 1
        if current > longest:
            longest = current
            longest_layout = layout
        previous = layout
    return longest, longest_layout


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate a defense deck and its rendered images.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--theme", required=True)
    parser.add_argument("--render-dir", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--profile", choices=["document", "projection", "lecture", "defense"], default=None)
    return parser.parse_args()


def add_issue(issues: list[dict], severity: str, page: str, message: str) -> None:
    issues.append({"severity": severity, "page": page, "message": message})


def defense_theme_issues(theme: dict) -> list[dict]:
    issues: list[dict] = []
    scale = theme.get("type_scale", {}) or {}
    if float(scale.get("body", 0)) * (72.0 / 153.6) < 24.0:
        add_issue(issues, "error", "type_scale", "Defense profile requires body >= 24 pt at 153.6 px/in.")
    if float(scale.get("header_title", 0)) * (72.0 / 153.6) < 32.0:
        add_issue(issues, "error", "type_scale", "Defense profile requires header_title >= 32 pt at 153.6 px/in.")
    text = theme.get("text", {}) or {}
    if not bool(text.get("body_bold", False)):
        add_issue(issues, "error", "text", "Defense profile requires text.body_bold for semantic body emphasis.")
    if not float(text.get("body_bold_min_px", 0)) <= float(scale.get("body", 0)) <= float(text.get("body_bold_max_px", 0)):
        add_issue(issues, "error", "text", "Defense profile body-bold range must include 24 pt body text.")
    body_family = str((theme.get("pptx_fonts", {}) or {}).get("body", "")).strip().lower()
    if any(term in body_family for term in ("times", "serif", "song", "songti", "simsun", "cambria")):
        add_issue(issues, "error", "fonts", "Defense profile requires a sans-serif body font, not a serif or CJK serif face.")
    for token_name in ("small", "micro", "reference_title", "reference_meta", "footnote"):
        if float(scale.get(token_name, 0)) * (72.0 / 153.6) < 18.0:
            add_issue(issues, "error", "type_scale", f"Defense profile requires {token_name} >= 18 pt at 153.6 px/in.")
    return issues


def content_units(slide: dict) -> int:
    layout = slide.get("layout")
    if layout == "bullets":
        columns = slide.get("columns") or []
        return len(columns) if columns else len(slide.get("bullets") or [])
    if layout == "grid":
        return len(slide.get("blocks") or [])
    if layout == "balance":
        left_items = (slide.get("left") or {}).get("items") or []
        right_items = (slide.get("right") or {}).get("items") or []
        return len(left_items) + len(right_items)
    if layout == "figure":
        return len(slide.get("panels") or [])
    if layout == "figure_flow":
        return len(slide.get("flow") or []) + len(slide.get("panels") or [])
    if layout == "concept":
        diagram = slide.get("diagram") or {}
        nodes = diagram.get("nodes") or diagram.get("items") or []
        return len(nodes) + len(slide.get("panels") or [])
    if layout == "formula":
        return len(slide.get("variables") or []) + len(slide.get("panels") or [])
    if layout == "pipeline":
        return len(slide.get("steps") or [])
    if layout == "compare":
        return len(slide.get("rows") or [])
    if layout == "metrics":
        return len(slide.get("metrics") or []) + len(slide.get("conditions") or [])
    if layout == "closing":
        return len(slide.get("points") or [])
    if layout == "hub":
        return len(slide.get("items") or [])
    if layout == "timeline":
        return len(slide.get("phases") or [])
    if layout == "references":
        return len(slide.get("references") or [])
    if layout == "quote":
        return 1 + int(bool(str(slide.get("commentary", "")).strip()))
    return 0


def slide_text_volume(slide: dict) -> int:
    total = sum(
        len(str(slide.get(key, "")).strip())
        for key in ("lead", "takeaway", "claim", "evidence", "boundary", "quote", "commentary")
    )
    for key in ("blocks", "steps", "panels", "rows", "items", "nodes", "metrics", "phases", "points"):
        for item in slide.get(key) or []:
            if isinstance(item, str):
                total += len(item.strip())
                continue
            if not isinstance(item, dict):
                continue
            for field in ("heading", "title", "body", "text", "footer", "output", "note", "period"):
                total += len(str(item.get(field, "")).strip())
            for field in ("bullets", "items", "conditions"):
                for value in item.get(field) or []:
                    total += len(str(value).strip())
    diagram = slide.get("diagram") or {}
    if isinstance(diagram, dict) and str(diagram.get("kind", "")).strip().lower() == "relations":
        for node in diagram.get("nodes") or diagram.get("items") or []:
            if isinstance(node, dict):
                for field in ("title", "label", "body"):
                    total += len(str(node.get(field, "")).strip())
        for relation in diagram.get("relations") or diagram.get("edges") or []:
            if isinstance(relation, dict):
                total += len(str(relation.get("label", "")).strip())
    return total


def visual_anchor(slide: dict) -> bool:
    return any(
        (
            slide.get("image"),
            slide.get("chart"),
            slide.get("image_width_ratio"),
            slide.get("diagram"),
            slide.get("formula"),
            slide.get("formula_latex"),
            slide.get("layout") in {"figure", "figure_flow", "hub", "concept", "formula"},
            any(str(block.get("span", 1)).isdigit() and int(block.get("span", 1)) > 1 for block in slide.get("blocks") or [] if isinstance(block, dict)),
            any(block.get("footer") or block.get("output") or block.get("stat") for block in slide.get("blocks") or [] if isinstance(block, dict)),
        )
    )


def validate_relations_diagram(
    diagram: dict,
    page_id: str,
    issues: list[dict],
    label: str,
) -> None:
    nodes = [
        node
        for node in (diagram.get("nodes") or diagram.get("items") or [])
        if isinstance(node, dict)
    ]
    if not 2 <= len(nodes) <= 8:
        add_issue(issues, "error", page_id, f"{label} needs two to eight relation nodes.")

    arrangement = str(diagram.get("arrangement", "")).strip().lower()
    has_coordinates = bool(nodes) and all(
        all(key in node for key in ("x", "y", "w", "h"))
        for node in nodes
    )
    if arrangement and arrangement not in {"free", "row", "two_row"}:
        add_issue(issues, "error", page_id, f"{label} arrangement must be free, row, or two_row.")
    if arrangement not in {"free", "row", "two_row"}:
        arrangement = "free" if has_coordinates else "row"

    node_ids: list[str] = []
    normalized_boxes: list[tuple[str, float, float, float, float]] = []
    for node_index, node in enumerate(nodes, start=1):
        node_id = str(node.get("id", "")).strip()
        if not node_id:
            add_issue(issues, "error", page_id, f"{label} node {node_index} needs a stable id.")
            continue
        if node_id in node_ids:
            add_issue(issues, "error", page_id, f"{label} node id '{node_id}' is duplicated.")
        node_ids.append(node_id)
        if not str(node.get("title") or node.get("label") or "").strip():
            add_issue(issues, "warning", page_id, f"{label} node '{node_id}' has no title/label.")
        tone = str(node.get("tone", "pale")).strip().lower()
        if tone not in RELATION_NODE_TONES:
            add_issue(
                issues,
                "error",
                page_id,
                f"{label} node '{node_id}' tone must be one of {sorted(RELATION_NODE_TONES)}.",
            )
        accent = str(node.get("accent", "")).strip().lower()
        if accent and accent not in {"blue", "red", "alert"}:
            add_issue(issues, "warning", page_id, f"{label} node '{node_id}' has an unknown accent.")

        if arrangement != "free":
            continue
        missing = [key for key in ("x", "y", "w", "h") if key not in node]
        if missing:
            add_issue(
                issues,
                "error",
                page_id,
                f"{label} free-layout node '{node_id}' is missing {', '.join(missing)}.",
            )
            continue
        try:
            x = float(node["x"])
            y = float(node["y"])
            node_width = float(node["w"])
            node_height = float(node["h"])
        except (TypeError, ValueError):
            add_issue(issues, "error", page_id, f"{label} node '{node_id}' coordinates must be numeric.")
            continue
        if not all(0.0 <= value <= 1.0 for value in (x, y, node_width, node_height)):
            add_issue(issues, "error", page_id, f"{label} node '{node_id}' coordinates must stay in 0–1.")
            continue
        if node_width <= 0.0 or node_height <= 0.0:
            add_issue(issues, "error", page_id, f"{label} node '{node_id}' needs positive width and height.")
            continue
        if x + node_width > 1.0001 or y + node_height > 1.0001:
            add_issue(issues, "error", page_id, f"{label} node '{node_id}' escapes the diagram bounds.")
            continue
        normalized_boxes.append((node_id, x, y, node_width, node_height))

    for first_index, first in enumerate(normalized_boxes):
        first_id, first_x, first_y, first_width, first_height = first
        first_area = first_width * first_height
        for second in normalized_boxes[first_index + 1 :]:
            second_id, second_x, second_y, second_width, second_height = second
            overlap_width = max(
                0.0,
                min(first_x + first_width, second_x + second_width) - max(first_x, second_x),
            )
            overlap_height = max(
                0.0,
                min(first_y + first_height, second_y + second_height) - max(first_y, second_y),
            )
            if overlap_width <= 0.0 or overlap_height <= 0.0:
                continue
            overlap_ratio = (overlap_width * overlap_height) / max(0.0001, min(first_area, second_width * second_height))
            if overlap_ratio > 0.08:
                add_issue(
                    issues,
                    "warning",
                    page_id,
                    f"{label} nodes '{first_id}' and '{second_id}' overlap; adjust x/y/w/h.",
                )

    relations = [
        relation
        for relation in (diagram.get("relations") or diagram.get("edges") or [])
        if isinstance(relation, dict)
    ]
    if not 1 <= len(relations) <= 10:
        add_issue(issues, "error", page_id, f"{label} needs one to ten explicit relations.")
    node_id_set = set(node_ids)
    known_edges: set[tuple[str, str]] = set()
    for relation_index, relation in enumerate(relations, start=1):
        source_id = str(relation.get("from", "")).strip()
        target_id = str(relation.get("to", "")).strip()
        if source_id not in node_id_set or target_id not in node_id_set:
            add_issue(
                issues,
                "error",
                page_id,
                f"{label} relation {relation_index} must reference existing node ids.",
            )
        if source_id and source_id == target_id:
            add_issue(issues, "error", page_id, f"{label} relation {relation_index} cannot be a self-loop.")
        edge = (source_id, target_id)
        if source_id and target_id and edge in known_edges:
            add_issue(issues, "warning", page_id, f"{label} relation {source_id}→{target_id} is duplicated.")
        known_edges.add(edge)
        style = str(relation.get("style", "ribbon")).strip().lower()
        if style not in RELATION_ARROW_STYLES:
            add_issue(
                issues,
                "error",
                page_id,
                f"{label} relation {relation_index} style must be one of {sorted(RELATION_ARROW_STYLES)}.",
            )
        route = str(relation.get("route", "straight")).strip().lower()
        if route not in RELATION_ROUTES:
            add_issue(
                issues,
                "error",
                page_id,
                f"{label} relation {relation_index} route must be one of {sorted(RELATION_ROUTES)}.",
            )
        if route == "orthogonal" and style != "line":
            add_issue(
                issues,
                "warning",
                page_id,
                f"{label} relation {relation_index} uses an orthogonal route; style '{style}' is ignored.",
            )
        tone = str(relation.get("tone", "primary")).strip().lower()
        if tone not in RELATION_TONES:
            add_issue(
                issues,
                "error",
                page_id,
                f"{label} relation {relation_index} tone must be one of {sorted(RELATION_TONES)}.",
            )
        for side_key in ("start_side", "end_side"):
            side = str(relation.get(side_key, "auto")).strip().lower()
            if side not in RELATION_SIDES:
                add_issue(
                    issues,
                    "error",
                    page_id,
                    f"{label} relation {relation_index} {side_key} must be one of {sorted(RELATION_SIDES)}.",
                )
        if "width" in relation:
            try:
                arrow_width = float(relation["width"])
            except (TypeError, ValueError):
                add_issue(issues, "error", page_id, f"{label} relation {relation_index} width is not numeric.")
            else:
                if arrow_width <= 0:
                    add_issue(issues, "error", page_id, f"{label} relation {relation_index} width must be positive.")


def validate_diagram(diagram: dict, page_id: str, issues: list[dict], label: str = "Diagram") -> None:
    """Validate semantic diagram data before rendering.

    Diagrams are content-bearing, not decoration. Keep their cardinality and
    required fields explicit so a model cannot silently produce an empty or
    misleading schematic.
    """
    if not isinstance(diagram, dict):
        add_issue(issues, "error", page_id, f"{label} must be an object.")
        return
    kind = str(diagram.get("kind", "sequence")).strip().lower()
    if kind not in DIAGRAM_KINDS:
        add_issue(
            issues,
            "error",
            page_id,
            f"{label} has unsupported kind: {kind}. Use one of {sorted(DIAGRAM_KINDS)}.",
        )
        return
    if kind == "relations":
        validate_relations_diagram(diagram, page_id, issues, label)
        return
    items = [item for item in (diagram.get("items") or []) if isinstance(item, dict)]
    limits = {
        "sequence": (2, 5),
        "branch": (2, 4),
        "stack": (2, 5),
        "ruler": (2, 8),
        "matrix": (4, 4),
        "cycle": (3, 5),
    }
    minimum, maximum = limits[kind]
    if not minimum <= len(items) <= maximum:
        add_issue(
            issues,
            "error",
            page_id,
            f"{label} needs {minimum}–{maximum} items for kind '{kind}'.",
        )
    if kind == "branch" and not isinstance(diagram.get("root"), dict):
        add_issue(issues, "warning", page_id, f"{label} should define root for branch diagrams.")
    if kind == "ruler":
        for item_index, item in enumerate(items, start=1):
            if "value" not in item:
                add_issue(issues, "error", page_id, f"{label} ruler item {item_index} needs numeric value.")
                continue
            try:
                float(item["value"])
            except (TypeError, ValueError):
                add_issue(issues, "error", page_id, f"{label} ruler item {item_index} value is not numeric.")
    for item_index, item in enumerate(items, start=1):
        if not str(item.get("title") or item.get("label") or "").strip():
            add_issue(issues, "warning", page_id, f"{label} item {item_index} has no title/label.")
    if "highlight" in diagram:
        highlight = diagram.get("highlight")
        if not isinstance(highlight, int) or not 0 <= highlight < len(items):
            add_issue(issues, "error", page_id, f"{label} highlight is outside the items range.")
    if kind == "matrix":
        if not isinstance(diagram.get("x_axis"), dict) or not isinstance(diagram.get("y_axis"), dict):
            add_issue(issues, "warning", page_id, f"{label} matrix should define x_axis and y_axis labels.")
        pairs = diagram.get("pairs") or []
        if pairs and len(pairs) != len(items):
            add_issue(issues, "warning", page_id, f"{label} pairs count should match item count.")


def validate_citation_entry(entry: dict, page_id: str, issues: list[dict], label: str) -> None:
    if not isinstance(entry, dict):
        add_issue(issues, "error", page_id, f"{label} must be an object.")
        return
    if not str(entry.get("title", "")).strip():
        add_issue(issues, "error", page_id, f"{label} needs a title.")
    venue = str(entry.get("venue", "")).strip()
    year = str(entry.get("year", "")).strip()
    doi = str(entry.get("doi", "")).strip()
    url = str(entry.get("url", "")).strip()
    if not (venue or year or doi or url):
        add_issue(issues, "warning", page_id, f"{label} needs venue, year, doi, or url.")
    identifier = normalize_doi(doi)
    if identifier:
        if identifier.lower().startswith("arxiv:"):
            if not ARXIV_RE.match(identifier):
                add_issue(issues, "warning", page_id, f"{label} arXiv id is malformed: {identifier}")
        elif identifier.startswith("10."):
            if not DOI_RE.match(identifier):
                add_issue(issues, "error", page_id, f"{label} doi is malformed: {identifier}")
        else:
            add_issue(
                issues,
                "warning",
                page_id,
                f"{label} doi field does not look like a DOI or arXiv id; move it to venue.",
            )
    if url and not url.lower().startswith(("http://", "https://")):
        add_issue(issues, "warning", page_id, f"{label} url should start with http:// or https://.")
    if year and len(year) != 4:
        add_issue(issues, "warning", page_id, f"{label} year should be four digits.")


def main() -> None:
    args = parse_args()
    deck_path = Path(args.deck).resolve()
    theme_path = Path(args.theme).resolve()
    render_dir = Path(args.render_dir).resolve()
    output_dir = Path(args.out).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    deck = load_json(deck_path)
    theme = load_theme(theme_path)
    profile = args.profile or str(deck.get("meta", {}).get("type_profile", "document"))
    apply_type_profile(theme, profile)
    navigation_mode = apply_navigation_mode(theme, navigation_mode_for(deck, theme))
    issues: list[dict] = []
    if profile == "defense":
        issues.extend(defense_theme_issues(theme))
    deck_mode = str(deck.get("meta", {}).get("deck_mode", "full")).strip().lower()
    sections = deck.get("sections", [])
    slides = deck.get("slides", [])
    section_ids = [section.get("id") for section in sections]
    declared_navigation_mode = str(deck.get("meta", {}).get("navigation_mode", "")).strip().lower()
    section_pages = [
        slide for slide in slides
        if str(slide.get("layout", "")).strip().lower() == "section"
    ]

    if deck_mode not in {"full", "content_only"}:
        add_issue(issues, "error", "meta", "meta.deck_mode must be 'full' or 'content_only'.")
    if declared_navigation_mode and declared_navigation_mode not in NAVIGATION_MODES:
        add_issue(
            issues,
            "error",
            "meta",
            f"meta.navigation_mode must be one of {sorted(NAVIGATION_MODES)}.",
        )
    if declared_navigation_mode == "sidebar" and section_pages:
        add_issue(
            issues,
            "warning",
            "navigation",
            "navigation_mode 'sidebar' does not include section dividers; remove layout 'section' pages or choose 'dividers'/'hybrid'.",
        )
    if declared_navigation_mode in {"dividers", "hybrid"} and not section_pages:
        add_issue(
            issues,
            "warning",
            "navigation",
            f"navigation_mode '{declared_navigation_mode}' expects section dividers; run add_section_pages.py before rendering.",
        )
    if not deck.get("meta", {}).get("title"):
        add_issue(issues, "error", "meta", "Missing meta.title.")
    title_mode = str(deck.get("meta", {}).get("title_mode", "topic")).strip().lower()
    if title_mode not in {"topic", "assertion", "summary"}:
        add_issue(issues, "error", "meta", "meta.title_mode must be 'topic', 'assertion' or 'summary'.")
    genre = str(deck.get("meta", {}).get("genre", "general")).strip().lower()
    if genre not in GENRES:
        add_issue(issues, "error", "meta", f"meta.genre must be one of {sorted(GENRES)}.")
    narrative = str(deck.get("meta", {}).get("narrative", "")).strip().lower()
    if narrative and narrative not in NARRATIVES:
        add_issue(issues, "error", "meta", f"meta.narrative must be one of {sorted(NARRATIVES)}.")
    source_policy = str(deck.get("meta", {}).get("source_policy", "off")).strip().lower()
    if source_policy not in SOURCE_POLICIES:
        add_issue(issues, "error", "meta", "meta.source_policy must be 'off', 'advisory', or 'strict'.")
    citations = deck.get("citations")
    if citations is not None and not isinstance(citations, dict):
        add_issue(issues, "error", "citations", "citations must be an object keyed by citation id.")
        citations = {}
    citations = citations if isinstance(citations, dict) else {}
    for key, entry in citations.items():
        validate_citation_entry(entry, "citations", issues, f"Citation '{key}'")
    if not slides:
        add_issue(issues, "error", "deck", "Deck has no slides.")
    if len(sections) < 4:
        add_issue(issues, "error", "sections", "Use at least four major sections.")
    if len(sections) > 6:
        add_issue(issues, "warning", "sections", "More than six sections will crowd the mini-directory.")
    if len(set(section_ids)) != len(section_ids):
        add_issue(issues, "error", "sections", "Section ids must be unique.")

    if slides and deck_mode == "full":
        if slides[0].get("layout") != "cover":
            add_issue(issues, "error", "01", "The first slide must use the cover layout.")
        if len(slides) > 1 and slides[1].get("layout") != "outline":
            add_issue(issues, "error", "02", "The second slide must use the outline layout.")
    if slides and deck_mode == "content_only":
        for index, slide in enumerate(slides, start=1):
            if slide.get("layout") in {"cover", "outline"}:
                add_issue(issues, "error", str(slide.get("id", f"{index:02d}")), "content_only mode must not include cover or outline slides.")

    for index, slide in enumerate(slides, start=1):
        page_id = str(slide.get("id", f"{index:02d}"))
        layout = slide.get("layout")
        title_text = str(slide.get("title", "")).strip()
        title_limit = int(deck.get("meta", {}).get("title_max_chars", 10))
        if title_mode == "assertion":
            title_limit += 8
        if layout not in SUPPORTED_LAYOUTS:
            add_issue(issues, "error", page_id, f"Unsupported layout: {layout}")
            continue
        if not slide.get("title"):
            add_issue(issues, "error", page_id, "Missing slide title.")
        elif layout not in {"cover", "outline", "section", "closing", "references"}:
            if title_units(title_text) > title_limit * 2:
                add_issue(
                    issues,
                    "warning",
                    page_id,
                    f"Header title is too long ({title_units(title_text)} units > {title_limit * 2}); "
                    "use a short topic label and move the claim to lead.",
                )
        if not str(slide.get("speech", "")).strip():
            add_issue(issues, "error", page_id, "Missing speech text.")
        variant = str(slide.get("variant", "")).strip().lower()
        if variant:
            allowed = LAYOUT_VARIANTS.get(str(layout), set())
            if allowed and variant not in allowed:
                add_issue(issues, "warning", page_id, f"Unknown variant '{variant}' for layout '{layout}'.")
            elif not allowed:
                add_issue(issues, "warning", page_id, f"Layout '{layout}' does not declare variants.")
        raw_cites = slide.get("cites")
        if raw_cites is not None and not isinstance(raw_cites, (str, list, tuple)):
            add_issue(issues, "error", page_id, "cites must be a string or a list of citation ids.")
        cites = normalize_cites(raw_cites)
        for key in cites:
            if key not in citations:
                add_issue(issues, "error", page_id, f"Unknown citation id: {key}")
        if len(cites) > 3:
            add_issue(issues, "warning", page_id, "More than three citations on one slide will crowd the footer.")
        if "source" in slide and not str(slide.get("source", "")).strip():
            add_issue(issues, "warning", page_id, "source is present but empty.")
        image_origin = str(slide.get("image_origin", "")).strip().lower()
        if image_origin and image_origin not in IMAGE_ORIGINS:
            add_issue(
                issues,
                "warning",
                page_id,
                f"image_origin should be one of {sorted(IMAGE_ORIGINS)}.",
            )
        if title_mode == "assertion" and layout not in {"cover", "outline", "section", "closing", "references"}:
            if not is_assertion_title(str(slide.get("title", ""))):
                add_issue(
                    issues,
                    "warning",
                    page_id,
                    "title_mode is 'assertion' but this title reads as a topic label, not a claim.",
                )
        if title_mode == "summary" and layout not in {"cover", "outline", "section", "closing", "references"}:
            title_text = str(slide.get("title", "")).strip()
            length = len(re.sub(r"\s", "", title_text))
            if not summary_title_valid(title_text):
                add_issue(
                    issues,
                    "warning",
                    page_id,
                    f"summary title should be 4-18 characters, got {length}.",
                )
            if is_assertion_title(title_text):
                add_issue(
                    issues,
                    "warning",
                    page_id,
                    "summary title reads like a claim; move the judgment into lead/claim and keep the heading descriptive.",
                )
            if not str(slide.get("claim", "")).strip():
                add_issue(
                    issues,
                    "error",
                    page_id,
                    "summary title mode requires a claim in the body - the judgment must live somewhere.",
                )
        if layout not in {"cover", "outline"} and slide.get("section") not in section_ids:
            add_issue(issues, "error", page_id, f"Unknown or missing section: {slide.get('section')}")
        chart = slide.get("chart")
        if chart is not None:
            if not isinstance(chart, dict):
                add_issue(issues, "error", page_id, "chart must be an object.")
            else:
                chart_type = str(chart.get("type", "histogram")).strip().lower()
                if chart_type not in {"histogram", "bar", "column", "line", "curve"}:
                    add_issue(issues, "error", page_id, f"Unsupported chart type: {chart_type}")
                values = chart.get("values") or []
                if len(values) < 3:
                    add_issue(issues, "error", page_id, "Chart needs at least three values.")
                if len(values) > 24:
                    add_issue(issues, "warning", page_id, "Chart has more than 24 values and may be unreadable.")
                for value_index, value in enumerate(values, start=1):
                    if not isinstance(value, (int, float)):
                        add_issue(issues, "error", page_id, f"Chart value {value_index} is not numeric.")
                if "highlight" in chart and values:
                    highlight = chart.get("highlight")
                    if not isinstance(highlight, int) or not 0 <= highlight < len(values):
                        add_issue(issues, "error", page_id, "Chart highlight index is outside the values range.")
                for field in ("unit", "condition", "source"):
                    if field in chart and not str(chart.get(field, "")).strip():
                        add_issue(issues, "warning", page_id, f"Chart {field} is present but empty.")
                if "sample_size" in chart:
                    sample_size = chart.get("sample_size")
                    try:
                        numeric_sample = float(sample_size)
                    except (TypeError, ValueError):
                        add_issue(issues, "error", page_id, "Chart sample_size must be numeric.")
                    else:
                        if numeric_sample <= 0:
                            add_issue(issues, "error", page_id, "Chart sample_size must be positive.")
                if "error_bar" in chart and not isinstance(chart.get("error_bar"), (list, tuple, dict, int, float)):
                    add_issue(issues, "warning", page_id, "Chart error_bar should be numbers or a spec object.")
        if layout in {"cover", "outline"} and slide.get("section"):
            add_issue(issues, "warning", page_id, "Cover and outline should not highlight a mini-directory section.")
        if layout == "bullets":
            columns = slide.get("columns") or []
            if not columns and slide.get("bullets"):
                columns = [{"bullets": slide["bullets"]}]
            if len(columns) not in {1, 2, 3}:
                add_issue(issues, "error", page_id, "Bullets layout supports one to three columns.")
            for column_index, column in enumerate(columns, start=1):
                if len(column.get("bullets", [])) > 4:
                    add_issue(issues, "warning", page_id, f"Column {column_index} has more than four bullets.")
        if layout == "figure":
            image_path = resolve_path(slide.get("image"), deck_path.parent)
            if not slide.get("chart") and (image_path is None or not image_path.exists()):
                add_issue(issues, "error", page_id, f"Missing figure: {slide.get('image')}")
            if not slide.get("image_role"):
                add_issue(issues, "warning", page_id, "Missing image_role.")
            if not slide.get("image_boundary"):
                add_issue(issues, "warning", page_id, "Missing image_boundary.")
            panels = slide.get("panels") or []
            if not 1 <= len(panels) <= 3:
                add_issue(issues, "error", page_id, "Figure layout needs one to three side panels.")
        if layout == "figure_flow":
            flow = slide.get("flow") or []
            panels = slide.get("panels") or []
            if not 3 <= len(flow) <= 5:
                add_issue(issues, "error", page_id, "Figure-flow layout needs three to five flow steps.")
            if not 1 <= len(panels) <= 3:
                add_issue(issues, "error", page_id, "Figure-flow layout needs one to three side panels.")
            image_path = resolve_path(slide.get("image"), deck_path.parent)
            if not slide.get("chart") and (image_path is None or not image_path.exists()):
                add_issue(issues, "error", page_id, "Figure-flow layout needs an image or chart.")
        if layout == "concept":
            diagram = slide.get("diagram")
            if not isinstance(diagram, dict):
                add_issue(issues, "error", page_id, "Concept layout needs a diagram object.")
            else:
                validate_diagram(diagram, page_id, issues, label="Concept diagram")
            panels = slide.get("panels") or []
            relation_map = isinstance(diagram, dict) and str(diagram.get("kind", "")).strip().lower() == "relations"
            if relation_map and not 0 <= len(panels) <= 3:
                add_issue(issues, "error", page_id, "Relation maps allow zero to three interpretation panels.")
            if not relation_map and not 1 <= len(panels) <= 3:
                add_issue(issues, "error", page_id, "Concept layout needs one to three interpretation panels.")
            if "diagram_frame" in slide and not isinstance(slide.get("diagram_frame"), bool):
                add_issue(issues, "warning", page_id, "diagram_frame should be true or false.")
            if slide.get("image") or slide.get("chart"):
                add_issue(issues, "warning", page_id, "Concept layout is diagram-led; use figure/figure_flow for real evidence.")
        if layout == "formula":
            if not (str(slide.get("formula", "")).strip() or str(slide.get("formula_latex", "")).strip()):
                add_issue(issues, "error", page_id, "Formula layout needs formula or formula_latex.")
            variables = slide.get("variables") or []
            if not 2 <= len(variables) <= 8:
                add_issue(issues, "error", page_id, "Formula layout needs two to eight variable explanations.")
            panels = slide.get("panels") or []
            if not 1 <= len(panels) <= 3:
                add_issue(issues, "error", page_id, "Formula layout needs one to three interpretation panels.")
        if layout == "pipeline":
            step_count = len(slide.get("steps") or [])
            if not 3 <= step_count <= 5:
                add_issue(issues, "error", page_id, "Pipeline needs three to five steps.")
        if layout == "compare":
            columns = slide.get("columns") or []
            if not 2 <= len(columns) <= 5:
                add_issue(issues, "error", page_id, "Compare layout needs two to five columns.")
            for row_index, row in enumerate(slide.get("rows") or [], start=1):
                if len(row) != len(columns):
                    add_issue(issues, "error", page_id, f"Compare row {row_index} has {len(row)} cells, expected {len(columns)}.")
            if len(slide.get("rows") or []) > 5:
                add_issue(issues, "warning", page_id, "Compare table has more than five rows.")
        if layout == "metrics":
            metric_count = len(slide.get("metrics") or [])
            if not 2 <= metric_count <= 4:
                add_issue(issues, "error", page_id, "Metrics layout needs two to four metrics.")
            conditions = slide.get("conditions") or []
            if conditions and not isinstance(conditions, list):
                add_issue(issues, "error", page_id, "Metrics conditions must be a list.")
            for condition_index, condition in enumerate(conditions, start=1):
                if isinstance(condition, dict):
                    if not condition.get("label") or not condition.get("value"):
                        add_issue(issues, "warning", page_id, f"Metric condition {condition_index} needs label and value.")
                elif not str(condition).strip():
                    add_issue(issues, "warning", page_id, f"Metric condition {condition_index} is empty.")
        if layout == "grid":
            blocks = slide.get("blocks") or []
            if not 3 <= len(blocks) <= 6:
                add_issue(issues, "error", page_id, "Grid layout needs three to six blocks.")
            columns = int(slide.get("columns", 2))
            if columns not in {2, 3}:
                add_issue(issues, "error", page_id, "Grid layout supports two or three columns.")
            for block_index, block in enumerate(blocks, start=1):
                if not block.get("heading"):
                    add_issue(issues, "warning", page_id, f"Grid block {block_index} has no heading.")
                span = int(block.get("span", 1))
                if not 1 <= span <= columns:
                    add_issue(issues, "error", page_id, f"Grid block {block_index} span {span} exceeds {columns} columns.")
                if block.get("image"):
                    image_path = resolve_path(block.get("image"), deck_path.parent)
                    if image_path is None or not image_path.exists():
                        add_issue(issues, "error", page_id, f"Missing grid image: {block.get('image')}")
                if block.get("formula_latex") and not block.get("formula"):
                    add_issue(
                        issues,
                        "warning",
                        page_id,
                        f"Grid block {block_index} has formula_latex but no plain formula fallback.",
                    )
                if block.get("diagram"):
                    validate_diagram(block.get("diagram"), page_id, issues, label=f"Grid block {block_index} diagram")
                    if block.get("image"):
                        add_issue(issues, "warning", page_id, f"Grid block {block_index} has both image and diagram; image wins in rendering.")
        if layout == "balance":
            for rail_name in ("left", "right"):
                rail = slide.get(rail_name) or {}
                items = rail.get("items") or []
                if not 2 <= len(items) <= 4:
                    add_issue(issues, "error", page_id, f"Balance {rail_name} rail needs two to four items.")
                if not rail.get("heading"):
                    add_issue(issues, "warning", page_id, f"Balance {rail_name} rail has no heading.")
            image_path = resolve_path(slide.get("image"), deck_path.parent)
            if not slide.get("chart") and (image_path is None or not image_path.exists()):
                add_issue(issues, "error", page_id, "Balance layout needs a center image or chart.")
            if not slide.get("image_role"):
                add_issue(issues, "warning", page_id, "Balance layout is missing image_role.")
            if not slide.get("image_boundary"):
                add_issue(issues, "warning", page_id, "Balance layout is missing image_boundary.")
        if layout == "hub":
            items = slide.get("items") or []
            if not 4 <= len(items) <= 6:
                add_issue(issues, "error", page_id, "Hub layout needs four to six surrounding items.")
            center = slide.get("center")
            if not isinstance(center, dict) or not center.get("value"):
                add_issue(issues, "warning", page_id, "Hub layout should define center.value.")
        if layout == "timeline":
            phases = slide.get("phases") or []
            if not 3 <= len(phases) <= 5:
                add_issue(issues, "error", page_id, "Timeline layout needs three to five phases.")
            for phase_index, phase in enumerate(phases, start=1):
                if not str(phase.get("period", "")).strip():
                    add_issue(issues, "error", page_id, f"Timeline phase {phase_index} has no period.")
                if not str(phase.get("title", "")).strip():
                    add_issue(issues, "error", page_id, f"Timeline phase {phase_index} has no title.")
                if not str(phase.get("body", "")).strip() and not phase.get("bullets"):
                    add_issue(issues, "warning", page_id, f"Timeline phase {phase_index} has no body or bullets.")
                if len(phase.get("bullets") or []) > 4:
                    add_issue(issues, "warning", page_id, f"Timeline phase {phase_index} has more than four bullets.")
        if layout == "references":
            references = slide.get("references") or []
            if not 4 <= len(references) <= 10:
                add_issue(issues, "error", page_id, "References layout needs four to ten entries.")
            seen_titles = set()
            for ref_index, reference in enumerate(references, start=1):
                title = str(reference.get("title", "")).strip()
                venue = str(reference.get("venue", "")).strip()
                if not title:
                    add_issue(issues, "error", page_id, f"Reference {ref_index} has no title.")
                if not venue:
                    add_issue(issues, "error", page_id, f"Reference {ref_index} has no venue.")
                if title and title in seen_titles:
                    add_issue(issues, "warning", page_id, f"Reference {ref_index} duplicates an earlier title.")
                seen_titles.add(title)
                identifier = normalize_doi(str(reference.get("doi", "")))
                if identifier:
                    if identifier.lower().startswith("arxiv:"):
                        if not ARXIV_RE.match(identifier):
                            add_issue(issues, "warning", page_id, f"Reference {ref_index} arXiv id is malformed: {identifier}")
                    elif identifier.startswith("10."):
                        if not DOI_RE.match(identifier):
                            add_issue(issues, "error", page_id, f"Reference {ref_index} doi is malformed: {identifier}")
                    else:
                        add_issue(
                            issues,
                            "warning",
                            page_id,
                            f"Reference {ref_index} doi field does not look like a DOI or arXiv id; move it to venue.",
                        )
                url = str(reference.get("url", "")).strip()
                if url and not url.lower().startswith(("http://", "https://")):
                    add_issue(issues, "warning", page_id, f"Reference {ref_index} url should start with http(s)://.")
        if layout == "closing":
            if not closing_has_summary_content(slide):
                add_issue(
                    issues,
                    "warning",
                    page_id,
                    "Closing slide has no summary points or boundary/next-step cards.",
                )
            if len(slide.get("points") or []) > 4:
                add_issue(issues, "warning", page_id, "Closing slide should keep at most four summary points.")
        if layout == "quote":
            if not str(slide.get("quote", "")).strip():
                add_issue(issues, "error", page_id, "Quote layout needs quote text.")
            if not str(slide.get("attribution", "")).strip():
                add_issue(issues, "warning", page_id, "Quote layout should include an attribution.")
            if not str(slide.get("commentary", "")).strip():
                add_issue(issues, "warning", page_id, "Quote layout should include commentary tying the quote to the argument.")
        if layout not in {"cover", "outline", "section", "closing"}:
            if not str(slide.get("lead", "")).strip():
                add_issue(issues, "warning", page_id, "Content page has no lead claim.")
            units = content_units(slide)
            if units < 2:
                add_issue(issues, "warning", page_id, "Content density is low; add a second evidence or explanation unit.")
            if profile in {"lecture", "defense"} and slide_text_volume(slide) < 70 and not visual_anchor(slide):
                add_issue(
                    issues,
                    "warning",
                    page_id,
                    "Projected content is sparse; add evidence, a content-derived footer, or a dominant figure.",
                )
            if layout in REQUIRES_IMAGE_OR_CHART:
                image_path = resolve_path(slide.get("image"), deck_path.parent)
                if not slide.get("chart") and (image_path is None or not image_path.exists()):
                    add_issue(issues, "error", page_id, "This layout needs a real image or chart as its visual anchor.")
        if source_policy != "off":
            severity = "error" if source_policy == "strict" else "warning"
            has_source = bool(str(slide.get("source", "")).strip() or cites)
            chart_source = bool(isinstance(chart, dict) and str(chart.get("source", "")).strip())
            if chart is not None and not (has_source or chart_source):
                add_issue(issues, severity, page_id, "Chart carries values but has no source or cites.")
            image_path = resolve_path(slide.get("image"), deck_path.parent)
            if image_path is not None and image_path.exists() and image_origin not in {"own", "generated", "template"}:
                if not has_source:
                    add_issue(
                        issues,
                        severity,
                        page_id,
                        "Image has no source. Set image_origin to own/generated/template or add source/cites.",
                    )

    content_layouts = [
        str(slide.get("layout", ""))
        for slide in slides
        if slide.get("layout") not in {"cover", "outline", "section", "closing", "references"}
    ]
    if profile in {"lecture", "defense"}:
        longest_run, longest_layout = longest_content_layout_run(slides)
        if longest_run >= 4:
            add_issue(
                issues,
                "warning",
                "rhythm",
                f"Projected deck repeats {longest_layout} for {longest_run} consecutive pages; vary the evidence composition.",
            )
        anchor_pages = [
            slide for slide in slides
            if slide.get("layout") not in {"cover", "outline", "section", "closing"}
            and (slide.get("image") or slide.get("chart") or slide.get("formula") or slide.get("formula_latex"))
        ]
        if not anchor_pages:
            add_issue(issues, "warning", "rhythm", "Projected deck has no hero figure, chart, or formula page.")
        if profile == "defense" and not any(
            str(slide.get("image_origin", "")).strip().lower() in {"own", "public", "borrowed"}
            and slide.get("image")
            and slide.get("source")
            and slide.get("image_role")
            and slide.get("image_boundary")
            for slide in slides
        ):
            add_issue(issues, "error", "figures", "Defense profile requires at least one paper-sourced figure with source, image_role, and image_boundary.")
    layout_set = set(content_layouts)
    genre_requirements = {
        "experiment": ({"figure", "figure_flow", "metrics", "compare"}, "a measured figure, metric, or comparison page"),
        "systems": ({"pipeline", "figure_flow", "concept", "balance"}, "an architecture/flow or system-boundary page"),
        "theory": ({"formula", "concept", "compare"}, "a formula, proof sketch, or formal comparison page"),
        "empirical": ({"compare", "metrics", "figure"}, "a data/identification/robustness page"),
        "survey": ({"hub", "compare", "timeline"}, "a taxonomy, comparison, or coverage page"),
        "design": ({"concept", "figure_flow", "pipeline", "balance"}, "a design-space or artifact-evaluation page"),
        "clinical": ({"figure", "metrics", "compare", "timeline"}, "an endpoint, safety, or participant-flow page"),
    }
    if genre in genre_requirements:
        required, description = genre_requirements[genre]
        if not layout_set & required:
            add_issue(issues, "warning", "genre", f"Genre '{genre}' usually needs {description}.")
    if genre == "humanities":
        if not layout_set & {"quote", "figure"}:
            add_issue(issues, "warning", "genre", "Humanities decks usually need a quotation or close-reading figure page.")
        if content_layouts:
            text_heavy = sum(1 for layout in content_layouts if layout in {"bullets", "grid", "compare"})
            if text_heavy / len(content_layouts) > 0.5:
                add_issue(
                    issues,
                    "warning",
                    "genre",
                    "More than half of the content pages are bullet/grid/table pages; humanities decks often rely more on quotations and images.",
                )

    expected_size = (int(theme["canvas"]["width"]), int(theme["canvas"]["height"]))
    rendered_paths = sorted(render_dir.glob("slide_*.png"))
    if len(rendered_paths) != len(slides):
        add_issue(issues, "error", "render", f"Found {len(rendered_paths)} rendered slides for {len(slides)} deck slides.")
    for index, path in enumerate(rendered_paths, start=1):
        image = Image.open(path).convert("RGB")
        if image.size != expected_size:
            add_issue(issues, "error", f"{index:02d}", f"Rendered size {image.size} != expected {expected_size}.")
    if len(rendered_paths) > 2:
        brand_box = tuple(theme["brand"]["header_box"])
        footer_y = int(theme["layout"]["footer_y"])
        reference = np.array(Image.open(rendered_paths[1]).convert("RGB"))
        reference_brand = reference[brand_box[1]:brand_box[3], brand_box[0]:brand_box[2]]
        reference_footer = reference[footer_y:footer_y + 3, 24:expected_size[0] - 24]
        brand_mask = None
        lockup_path = theme.get("brand", {}).get("lockup_path")
        if lockup_path and Path(lockup_path).exists():
            lockup = Image.open(lockup_path).convert("RGBA")
            box_width = brand_box[2] - brand_box[0]
            box_height = brand_box[3] - brand_box[1]
            scale = min(box_width / lockup.width, box_height / lockup.height)
            size = (max(1, round(lockup.width * scale)), max(1, round(lockup.height * scale)))
            alpha = lockup.resize(size, Image.Resampling.LANCZOS).getchannel("A")
            full_mask = np.zeros((box_height, box_width), dtype=bool)
            offset_x = (box_width - size[0]) // 2
            offset_y = (box_height - size[1]) // 2
            full_mask[offset_y:offset_y + size[1], offset_x:offset_x + size[0]] = np.array(alpha) > 180
            brand_mask = full_mask
        section_pages = {
            position for position, slide in enumerate(slides, start=1)
            if str(slide.get("layout", "")).strip().lower() == "section"
        }
        for index, path in enumerate(rendered_paths[1:], start=2):
            # 分隔页是有意的整页版式，不承载页眉与页脚，豁免这两项一致性检查
            if index in section_pages:
                continue
            current = np.array(Image.open(path).convert("RGB"))
            current_brand = current[brand_box[1]:brand_box[3], brand_box[0]:brand_box[2]]
            current_footer = current[footer_y:footer_y + 3, 24:expected_size[0] - 24]
            if brand_mask is not None:
                difference = np.abs(reference_brand.astype(int) - current_brand.astype(int))[brand_mask]
                brand_mismatch = int(difference.max()) if difference.size else 0
            else:
                brand_mismatch = int(np.abs(reference_brand.astype(int) - current_brand.astype(int)).max())
            brand_tolerance = 4 if brand_mask is not None else 18
            if brand_mismatch > brand_tolerance:
                add_issue(issues, "error", f"{index:02d}", "Upper-right identity is not pixel-fixed.")
            if not np.array_equal(reference_footer, current_footer):
                add_issue(issues, "error", f"{index:02d}", "Footer rule is not pixel-fixed.")

    render_report_path = render_dir.parent / "render_report.json"
    if render_report_path.exists():
        render_report = load_json(render_report_path)
        for warning in render_report.get("warnings", []):
            message = str(warning.get("message", "")).strip()
            page = str(warning.get("page", "render"))
            if not message:
                continue
            fatal_terms = ("overflowed", "did not fit", "missing", "has no")
            severity = "error" if any(term in message.lower() for term in fatal_terms) else "warning"
            add_issue(issues, severity, page, f"Renderer warning: {message}")
    else:
        add_issue(issues, "warning", "render", "render_report.json is missing; visual overflow was not checked.")

    estimated = sum(estimate_seconds(str(slide.get("speech", ""))) for slide in slides)
    target_seconds = int(float(deck.get("meta", {}).get("duration_min", 20)) * 60)
    if target_seconds and estimated > target_seconds * 1.15:
        add_issue(issues, "warning", "timing", f"Estimated speech {estimated}s exceeds target {target_seconds}s by more than 15%.")
    if target_seconds and estimated < target_seconds * 0.65:
        add_issue(issues, "warning", "timing", f"Estimated speech {estimated}s is much shorter than target {target_seconds}s.")

    # Visible text must not carry code-like pseudo-formulas: N_s reads as code,
    # Nₛ reads as math. LaTeX sources are exempt.
    import re as _re

    pseudo = _re.compile(r"[A-Za-z][A-Za-z0-9]{0,6}_[a-z]{1,3}")
    for slide in slides:
        page_id = str(slide.get("id", ""))
        probes = []
        for key in ("title", "lead", "takeaway", "footnote"):
            probes.append(slide.get(key))
        for block in slide.get("blocks") or []:
            if isinstance(block, dict):
                probes.extend([block.get("heading"), block.get("body"), block.get("footer")])
                diagram = block.get("diagram")
                if isinstance(diagram, dict):
                    for node in diagram.get("nodes") or diagram.get("items") or []:
                        if isinstance(node, dict):
                            probes.extend([node.get("title"), node.get("label"), node.get("body")])
                    for relation in diagram.get("relations") or diagram.get("edges") or []:
                        if isinstance(relation, dict):
                            probes.append(relation.get("label"))
        diagram = slide.get("diagram")
        if isinstance(diagram, dict):
            for node in diagram.get("nodes") or diagram.get("items") or []:
                if isinstance(node, dict):
                    probes.extend([node.get("title"), node.get("label"), node.get("body")])
            for relation in diagram.get("relations") or diagram.get("edges") or []:
                if isinstance(relation, dict):
                    probes.append(relation.get("label"))
        for panel in slide.get("panels") or []:
            if isinstance(panel, dict):
                probes.extend([panel.get("heading"), panel.get("body")])
        for variable in slide.get("variables") or []:
            if isinstance(variable, dict):
                probes.append(variable.get("symbol"))
        for probe in probes:
            if not isinstance(probe, str):
                continue
            hit = pseudo.search(probe)
            if hit:
                add_issue(
                    issues,
                    "warning",
                    page_id,
                    f"Pseudo-formula in visible text: '{hit.group(0)}'. Use Unicode subscripts (Nₛ, Eₚ, fₚᵣ) or keep it in formula_latex.",
                )
                break

    # Design audit: structure, hierarchy and emphasis. These are warnings rather
    # than errors because "looks like a form" cannot invalidate an otherwise
    # valid deck, but it should never pass silently either.
    try:
        from qa_design import audit_design

        design = audit_design(deck, theme, Path(args.render_dir))
        for check in design["checks"]:
            if check["ok"]:
                continue
            add_issue(
                issues,
                "warning",
                "design",
                f"Design: {check['metric']} = {check['value']} (target {check['target']}). {DESIGN_HINTS.get(check['metric'], '')}".strip(),
            )
    except Exception as exc:  # noqa: BLE001 - the audit must never break validation
        add_issue(issues, "warning", "design", f"Design audit skipped: {type(exc).__name__}: {exc}")

    error_count = sum(1 for issue in issues if issue["severity"] == "error")
    warning_count = sum(1 for issue in issues if issue["severity"] == "warning")
    report = {
        "deck": str(deck_path),
        "theme": theme["name"],
        "type_profile": profile,
        "genre": genre,
        "narrative": narrative,
        "deck_mode": deck_mode,
        "navigation_mode": navigation_mode,
        "slide_count": len(slides),
        "rendered_count": len(rendered_paths),
        "estimated_speech_seconds": estimated,
        "target_speech_seconds": target_seconds,
        "error_count": error_count,
        "warning_count": warning_count,
        "issues": issues,
    }
    write_json(output_dir / "validation_report.json", report)
    print(json.dumps(report, ensure_ascii=True, indent=2))
    if error_count:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
