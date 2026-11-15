from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path

from PIL import Image

from common import write_json
from extract_source import FIGURE_CAPTION_RE, clean_context, slugify


USER_AGENT = "academic-ppt-skill/0.1 (figure finder)"
ATOM = "{http://www.w3.org/2005/Atom}"
ARXIV = "{http://arxiv.org/schemas/atom}"
STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "by", "for", "from", "in", "into",
    "is", "of", "on", "or", "that", "the", "their", "this", "to", "using", "via", "with",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Find and extract figures from open arXiv papers.")
    parser.add_argument("--query", default=None, help="English search query, e.g. vision-language-action robot.")
    parser.add_argument("--ids", default=None, help="Comma-separated arXiv ids, e.g. 2406.09246,2307.15818.")
    parser.add_argument("--out", default="work/figure_pack")
    parser.add_argument("--max-papers", type=int, default=5)
    parser.add_argument("--max-figures", type=int, default=12)
    parser.add_argument("--max-per-paper", type=int, default=3)
    parser.add_argument("--min-size", type=int, default=160)
    return parser.parse_args()


def tokens(text: str) -> set[str]:
    return {
        token
        for token in re.findall(r"[a-z0-9]+|[\u4e00-\u9fff]", str(text).lower())
        if token not in STOPWORDS and len(token) > 1
    }


def fetch(url: str, timeout: int = 60, retries: int = 3) -> bytes:
    last_error: Exception | None = None
    for attempt in range(retries):
        try:
            request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return response.read()
        except Exception as error:
            last_error = error
            if attempt + 1 < retries:
                time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"Failed to fetch {url}: {last_error}")


def query_arxiv(params: dict) -> list[dict]:
    url = "https://export.arxiv.org/api/query?" + urllib.parse.urlencode(params)
    root = ET.fromstring(fetch(url))
    papers: list[dict] = []
    for entry in root.findall(f"{ATOM}entry"):
        arxiv_id = str(entry.findtext(f"{ATOM}id", default="")).rsplit("/", 1)[-1]
        title = clean_context(entry.findtext(f"{ATOM}title", default=""), 300)
        summary = clean_context(entry.findtext(f"{ATOM}summary", default=""), 600)
        authors = [
            clean_context(author.findtext(f"{ATOM}name", default=""), 80)
            for author in entry.findall(f"{ATOM}author")
        ]
        pdf_url = ""
        for link in entry.findall(f"{ATOM}link"):
            if link.get("title") == "pdf" or link.get("type") == "application/pdf":
                pdf_url = link.get("href", "")
                break
        if not pdf_url and arxiv_id:
            pdf_url = f"https://arxiv.org/pdf/{arxiv_id}.pdf"
        license_value = str(entry.findtext(f"{ARXIV}license", default=""))
        papers.append(
            {
                "arxiv_id": arxiv_id,
                "title": title,
                "summary": summary,
                "authors": authors,
                "pdf_url": pdf_url,
                "abs_url": f"https://arxiv.org/abs/{arxiv_id}" if arxiv_id else "",
                "license": license_value or "Check the arXiv abs page for the license.",
            }
        )
    return [paper for paper in papers if paper["pdf_url"]]


def search_arxiv(query: str, limit: int) -> list[dict]:
    return query_arxiv(
        {
            "search_query": f"all:{query}",
            "start": "0",
            "max_results": str(max(1, limit)),
            "sortBy": "relevance",
            "sortOrder": "descending",
        }
    )


def fetch_arxiv_ids(ids: list[str]) -> list[dict]:
    return query_arxiv({"id_list": ",".join(ids), "max_results": str(max(1, len(ids)))})


def page_figure_regions(page) -> list[tuple[object, str]]:
    import pymupdf as fitz

    blocks = page.get_text("blocks")
    image_rects = []
    for image_info in page.get_images(full=True):
        for rect in page.get_image_rects(image_info[0]):
            if rect.width >= 60 and rect.height >= 60:
                image_rects.append(rect)
    drawing_rects = []
    for drawing in page.get_drawings():
        rect = drawing.get("rect")
        if rect is not None and rect.width >= 20 and rect.height >= 20:
            drawing_rects.append(rect)
    figure_rects = image_rects + drawing_rects
    regions: list[tuple[object, str]] = []
    for block in blocks:
        if len(block) < 5:
            continue
        caption = str(block[4]).strip()
        if not caption or not FIGURE_CAPTION_RE.match(caption):
            continue
        caption_rect = fitz.Rect(block[:4])
        above = [rect for rect in figure_rects if rect.y1 <= caption_rect.y0 - 2]
        below = [rect for rect in figure_rects if rect.y0 >= caption_rect.y1 + 2]
        candidates: list = []
        if above:
            nearest = max(above, key=lambda rect: rect.y1)
            candidates = [
                rect
                for rect in above
                if rect.y1 >= nearest.y1 - 140
                and min(rect.x1, nearest.x1) > max(rect.x0, nearest.x0)
            ]
            direction = "above"
        elif below:
            nearest = min(below, key=lambda rect: rect.y0)
            candidates = [
                rect
                for rect in below
                if rect.y0 <= nearest.y0 + 140
                and min(rect.x1, nearest.x1) > max(rect.x0, nearest.x0)
            ]
            direction = "below"
        else:
            continue
        clip = fitz.Rect(
            min(rect.x0 for rect in candidates),
            min(rect.y0 for rect in candidates),
            max(rect.x1 for rect in candidates),
            max(rect.y1 for rect in candidates),
        )
        clip.x0 = max(0, clip.x0 - 6)
        clip.x1 = min(page.rect.width, clip.x1 + 6)
        if direction == "above":
            clip.y0 = max(0, clip.y0 - 6)
            clip.y1 = min(clip.y1 + 6, caption_rect.y0 - 4)
        else:
            clip.y0 = max(clip.y0 - 6, caption_rect.y1 + 4)
            clip.y1 = min(page.rect.height, clip.y1 + 6)
        if clip.width < 120 or clip.height < 80:
            continue
        regions.append((clip, caption))
    return regions


def extract_paper_figures(
    paper: dict,
    figures_dir: Path,
    max_per_paper: int,
    min_size: int,
) -> list[dict]:
    import pymupdf as fitz

    payload = fetch(paper["pdf_url"], timeout=120)
    document = fitz.open(stream=payload, filetype="pdf")
    figures: list[dict] = []
    try:
        for page_number, page in enumerate(document, start=1):
            if len(figures) >= max_per_paper:
                break
            for clip, caption in page_figure_regions(page):
                if len(figures) >= max_per_paper:
                    break
                pixmap = page.get_pixmap(clip=clip, dpi=200, alpha=False)
                data = pixmap.tobytes("png")
                image = Image.open(io.BytesIO(data)).convert("RGB")
                width, height = image.size
                if width < max(min_size, 240) or height < max(min_size, 160):
                    continue
                grayscale = image.convert("L")
                minimum, maximum = grayscale.getextrema()
                if maximum - minimum < 12:
                    continue
                digest = hashlib.sha256(data).hexdigest()
                figure_id = f"paper_{paper['arxiv_id'].replace('.', '_')}_fig_{len(figures) + 1:02d}"
                target = figures_dir / f"{figure_id}_{slugify(caption or paper['title'], 'figure')}.png"
                image.save(target, "PNG", optimize=True)
                figures.append(
                    {
                        "id": figure_id,
                        "path": f"figures/{target.name}",
                        "source": paper["abs_url"] or paper["pdf_url"],
                        "paper_title": paper["title"],
                        "authors": paper["authors"],
                        "arxiv_id": paper["arxiv_id"],
                        "pdf_url": paper["pdf_url"],
                        "license": paper["license"],
                        "page": page_number,
                        "caption": clean_context(caption, 400),
                        "alt_hint": clean_context(caption or paper["title"], 180),
                        "width": width,
                        "height": height,
                        "sha256": digest,
                        "kind": "pdf_region",
                        "bbox": [round(value, 1) for value in (clip.x0, clip.y0, clip.x1, clip.y1)],
                    }
                )
    finally:
        document.close()
    return figures


def build(
    query: str | None,
    out: Path,
    max_papers: int,
    max_figures: int,
    max_per_paper: int,
    min_size: int,
    ids: list[str] | None = None,
) -> dict:
    figures_dir = out / "figures"
    figures_dir.mkdir(parents=True, exist_ok=True)
    try:
        papers = fetch_arxiv_ids(ids) if ids else search_arxiv(query or "", max_papers)
    except Exception as error:
        raise SystemExit(f"arXiv search failed: {error}")
    query_tokens = tokens(query or " ".join(ids or []))
    figures: list[dict] = []
    errors: list[dict] = []
    for paper in papers:
        try:
            extracted = extract_paper_figures(paper, figures_dir, max_per_paper, min_size)
        except Exception as error:
            errors.append({"arxiv_id": paper.get("arxiv_id"), "error": str(error)})
            continue
        for figure in extracted:
            text = f"{paper['title']} {paper['summary']} {figure.get('caption', '')}"
            overlap = query_tokens & tokens(text)
            figure["relevance"] = round(len(overlap) / max(1, len(query_tokens)), 3)
            figures.append(figure)
    figures.sort(key=lambda item: (item.get("relevance", 0), item["width"] * item["height"]), reverse=True)
    figures = figures[:max_figures]
    report = {
        "schema_version": 1,
        "query": query or "ids: " + ", ".join(ids or []),
        "ids": ids or [],
        "papers": papers,
        "figures": figures,
        "errors": errors,
        "counts": {"papers": len(papers), "figures": len(figures), "errors": len(errors)},
        "license_note": "arXiv figures may be subject to the paper's license. Check the abs page and cite the source before public use.",
    }
    write_json(out / "figures.json", report)
    lines = [
        f"# Figure Pack: {query or ', '.join(ids or [])}",
        "",
        "> Auto-extracted from open arXiv PDFs. Check each caption, license, and whether the figure actually supports the slide claim.",
        "",
    ]
    for figure in figures:
        lines.append(f"## `{figure['id']}` — relevance {figure['relevance']}")
        lines.append("")
        lines.append(f"- File: `{figure['path']}`")
        lines.append(f"- Paper: {figure['paper_title']} ({figure['arxiv_id']})")
        lines.append(f"- Source: {figure['source']}")
        lines.append(f"- Caption: {figure.get('caption') or '（未提取到 caption）'}")
        lines.append(f"- License: {figure['license']}")
        lines.append("")
    (out / "figures.md").write_text("\n".join(lines), encoding="utf-8")
    return report


def main() -> None:
    args = parse_args()
    ids = [item.strip() for item in (args.ids or "").replace("，", ",").split(",") if item.strip()]
    if not args.query and not ids:
        raise SystemExit("Provide --query or --ids.")
    report = build(
        args.query,
        Path(args.out).resolve(),
        max(1, args.max_papers),
        max(1, args.max_figures),
        max(1, args.max_per_paper),
        max(32, args.min_size),
        ids=ids or None,
    )
    print(json.dumps(report["counts"] | {"out": args.out}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
