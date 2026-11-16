#!/usr/bin/env python3
"""Find a university emblem on the open web, or say clearly that there isn't one.

Chinese university seals are usually non-free: they live on zh.wikipedia (blocked
in some networks) or behind a school-site WAF. Wikimedia Commons only carries
them for some institutions. So instead of one hard-coded source, this script
walks a fixed chain and reports what each step said:

1. Wikimedia Commons search — several keyword variants, SVG/PNG preferred
2. Chinese Wikipedia file API — near-universal coverage for mainland schools
3. The school homepage — first ``<img>`` whose src looks like a logo
4. ``/favicon.ico`` — last resort, low resolution but real

Nothing found is a valid answer: the caller then asks the user for a file instead
of drawing a fake seal. The skill must never invent one.
"""

from __future__ import annotations

import argparse
import io
import json
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

USER_AGENT = "academic-ppt-skill/0.1 (emblem resolver; contact: local user)"
TIMEOUT = 25


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Resolve an institution emblem from public sources.")
    parser.add_argument("--name", required=True, help="Chinese name, e.g. 兰州大学")
    parser.add_argument("--latin", default="", help="Latin name, e.g. Lanzhou University")
    parser.add_argument("--out", default="emblem", help="Output directory")
    parser.add_argument("--min-size", type=int, default=120, help="Minimum width in pixels")
    return parser.parse_args()


def fetch(url: str, binary: bool = False):
    """Fetch with urllib, then fall back to curl.

    Some Windows installs reset TLS inside Python's urllib (``WinError 10054``)
    while the system curl works fine. The fallback keeps the resolver usable on
    those machines instead of reporting a false "not found".
    """
    try:
        request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
            payload = response.read()
    except Exception:
        import subprocess

        completed = subprocess.run(
            ["curl", "-sSL", "--connect-timeout", "10", "--max-time", str(TIMEOUT), "-A", USER_AGENT, url],
            capture_output=True,
        )
        if completed.returncode != 0 or not completed.stdout:
            raise RuntimeError(f"curl failed for {url}")
        payload = completed.stdout
    return payload if binary else payload.decode("utf-8", "replace")


def commons_search(queries: list[str]) -> list[dict]:
    hits: list[dict] = []
    for query in queries:
        params = urllib.parse.urlencode(
            {
                "action": "query",
                "generator": "search",
                "gsrsearch": query,
                "gsrnamespace": "6",
                "gsrlimit": "8",
                "prop": "imageinfo",
                "iiprop": "url|mime|size",
                "format": "json",
            }
        )
        try:
            data = json.loads(fetch("https://commons.wikimedia.org/w/api.php?" + params))
        except Exception:
            continue
        for page in ((data.get("query") or {}).get("pages") or {}).values():
            info = (page.get("imageinfo") or [{}])[0]
            mime = str(info.get("mime") or "")
            if mime not in {"image/png", "image/svg+xml", "image/jpeg"}:
                continue
            title = str(page.get("title") or "")
            if not re.search(r"(logo|seal|emblem|badge|校徽)", title, re.I):
                continue
            hits.append({"source": "commons", "title": title, "url": info.get("url"), "mime": mime,
                         "width": info.get("width"), "height": info.get("height")})
    return hits


def wikipedia_file(latin: str, name: str) -> list[dict]:
    titles = [f"File:{latin}_logo.png", f"File:{latin}_seal.svg", f"File:{latin}_logo.svg"]
    if not latin:
        return []
    params = urllib.parse.urlencode(
        {"action": "query", "titles": "|".join(titles), "prop": "imageinfo",
         "iiprop": "url|mime|size", "format": "json"}
    )
    hits: list[dict] = []
    try:
        data = json.loads(fetch("https://zh.wikipedia.org/w/api.php?" + params))
    except Exception:
        return hits
    for page in ((data.get("query") or {}).get("pages") or {}).values():
        if "missing" in page:
            continue
        info = (page.get("imageinfo") or [{}])[0]
        if info.get("url"):
            hits.append({"source": "zh.wikipedia", "title": page.get("title"), "url": info.get("url"),
                         "mime": str(info.get("mime") or ""), "width": info.get("width"), "height": info.get("height")})
    return hits


def homepage_logo(domain: str) -> list[dict]:
    hits: list[dict] = []
    try:
        html = fetch(f"https://{domain}/")
    except Exception:
        return hits
    for match in re.finditer(r'<img[^>]+src="([^"]+)"', html, re.I):
        src = match.group(1)
        if not re.search(r"(logo|badge|emblem|校徽|xh)", src, re.I):
            continue
        url = urllib.parse.urljoin(f"https://{domain}/", src)
        hits.append({"source": "homepage", "title": src.split("/")[-1], "url": url, "mime": "", "width": None, "height": None})
        if len(hits) >= 3:
            break
    return hits


def download(candidate: dict, out_dir: Path, min_size: int) -> dict | None:
    url = candidate.get("url")
    if not url:
        return None
    try:
        payload = fetch(url, binary=True)
    except Exception as exc:
        return None
    suffix = ".svg" if str(candidate.get("mime", "")).endswith("svg") else (".png" if payload[:4] == b"\x89PNG" else ".img")
    path = out_dir / f"emblem_candidate{suffix}"
    path.write_bytes(payload)
    width = candidate.get("width")
    if suffix != ".svg" and width is None:
        try:
            from PIL import Image

            with Image.open(io.BytesIO(payload)) as image:
                width, height = image.size
                candidate["width"], candidate["height"] = width, height
        except Exception:
            pass
    if suffix != ".svg" and candidate.get("width") and int(candidate["width"]) < min_size:
        return None
    candidate["file"] = str(path)
    candidate["bytes"] = len(payload)
    return candidate


def main() -> None:
    args = parse_args()
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    queries = [args.latin, f"{args.latin} logo", f"{args.latin} seal", args.name, f"{args.name} 校徽"]
    queries = [query for query in queries if query.strip()]

    trail: list[dict] = []
    candidates = commons_search(queries)
    trail.append({"step": "commons", "found": len(candidates)})
    if not candidates:
        candidates = wikipedia_file(args.latin, args.name)
        trail.append({"step": "zh.wikipedia", "found": len(candidates)})
    if not candidates and args.latin:
        domain = re.sub(r"[^a-z]", "", args.latin.lower().replace("university", "")) + ".edu.cn"
        candidates = homepage_logo(domain)
        trail.append({"step": "homepage", "domain": domain, "found": len(candidates)})

    saved = None
    for candidate in candidates:
        saved = download(candidate, out_dir, args.min_size)
        if saved:
            break

    report = {
        "name": args.name,
        "latin": args.latin,
        "trail": trail,
        "found": bool(saved),
        "emblem": saved,
        "next": (
            "把这张图喂给 build_brand_lockup.py --emblem 生成 lockup，再写回 theme.json 的 brand.lockup"
            if saved
            else "公开源里没有可用校徽。不要伪造，改为向用户索要官方 PNG/SVG，或用纯文字锁版（brand.fallback_text + fonts.brand 手写体）。"
        ),
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if not saved:
        sys.exit(2)


if __name__ == "__main__":
    main()
