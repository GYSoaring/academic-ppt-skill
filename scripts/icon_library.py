#!/usr/bin/env python3
"""Pull a vector icon by name and hand it to the deck as a PNG or SVG path.

The renderer's built-in glyphs (`draw_icon`) cover a dozen shapes. Real decks
need a wider vocabulary: shields, arrows, gauges, networks, satellites. Instead
of hand-drawing them, fetch from **Lucide** (MIT, 1500+ icons, 24×24 stroke SVG)
and recolor to the theme.

Sources tried in order, because CDN availability varies by network:
  1. unpkg.com/lucide-static/icons
  2. cdn.jsdelivr.net/npm/lucide-static/icons
  3. raw.githubusercontent.com/lucide-icons/lucide/main/icons

Both urllib and curl are attempted — some Windows installs reset TLS inside
Python while curl still works.

Usage:
  python scripts/icon_library.py --list --query shield
  python scripts/icon_library.py --name shield-check --color "#1A4E9B" --out icons/
  python scripts/icon_library.py --name satellite --format svg --out icons/
"""

from __future__ import annotations

import argparse
import io
import json
import re
import subprocess
import urllib.request
from pathlib import Path

USER_AGENT = "academic-ppt-skill/0.1 (icon library)"
CDN_TEMPLATES = [
    "https://unpkg.com/lucide-static@latest/icons/{name}.svg",
    "https://cdn.jsdelivr.net/npm/lucide-static@latest/icons/{name}.svg",
    "https://raw.githubusercontent.com/lucide-icons/lucide/main/icons/{name}.svg",
]
INDEX_URL = "https://unpkg.com/lucide-static@latest/icons.json"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fetch a Lucide icon and optionally rasterise it.")
    parser.add_argument("--name", default=None, help="Icon name, e.g. shield-check")
    parser.add_argument("--list", action="store_true", help="List available icons (optionally filtered)")
    parser.add_argument("--query", default="", help="Filter for --list")
    parser.add_argument("--color", default="#1A4E9B", help="Stroke colour")
    parser.add_argument("--size", type=int, default=128, help="Raster size in pixels")
    parser.add_argument("--format", choices=["svg", "png"], default="png")
    parser.add_argument("--out", default="icons")
    return parser.parse_args()


def fetch_text(url: str) -> str:
    try:
        request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(request, timeout=20) as response:
            return response.read().decode("utf-8", "replace")
    except Exception:
        completed = subprocess.run(
            ["curl", "-sSL", "--connect-timeout", "10", "--max-time", "25", "-A", USER_AGENT, url],
            capture_output=True,
        )
        if completed.returncode != 0:
            raise RuntimeError(f"cannot fetch {url}")
        return completed.stdout.decode("utf-8", "replace")


def list_icons(query: str) -> list[str]:
    try:
        payload = fetch_text(INDEX_URL)
        names = sorted(re.findall(r'"([a-z0-9-]+)"', payload))
        names = sorted({name for name in names if name and "-" in name or len(name) > 2})
    except Exception:
        names = []
    if query:
        names = [name for name in names if query in name]
    return names[:400]


def get_svg(name: str) -> str:
    errors = []
    for template in CDN_TEMPLATES:
        url = template.format(name=name)
        try:
            svg = fetch_text(url)
            if "<svg" in svg:
                return svg
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{url}: {exc}")
    raise RuntimeError("icon not found on any CDN:\n" + "\n".join(errors))


def recolour(svg: str, color: str) -> str:
    svg = re.sub(r'stroke="[^"]*"', f'stroke="{color}"', svg)
    if "stroke=" not in svg:
        svg = svg.replace("<svg", f'<svg stroke="{color}"', 1)
    return svg



def _flatten(path_d: str, scale: float, offset: float) -> list[list[tuple[float, float]]]:
    """Turn an SVG path into polylines. Covers the subset Lucide uses:
    M/m L/l H/h V/v C/c Q/q Z, absolute and relative."""
    parts = re.findall(r"[MmLlHhVvCcQqZz][^MmLlHhVvCcQqZz]*", path_d)
    polylines: list[list[tuple[float, float]]] = []
    current: list[tuple[float, float]] = []
    x = y = 0.0
    start_x = start_y = 0.0

    def point(px: float, py: float) -> tuple[float, float]:
        return (offset + px * scale, offset + py * scale)

    for part in parts:
        command = part[0]
        numbers = [float(value) for value in re.findall(r"-?\d*\.?\d+(?:e-?\d+)?", part[1:])]
        if command in "Zz":
            if current:
                current.append(point(start_x, start_y))
                polylines.append(current)
                current = []
            x, y = start_x, start_y
            continue
        step = {"M": 2, "L": 2, "H": 1, "V": 1, "C": 6, "Q": 4}[command.upper()]
        relative = command.islower()
        for index in range(0, len(numbers) - step + 1, step):
            values = numbers[index:index + step]
            upper = command.upper()
            if upper == "M":
                nx, ny = values
                if relative:
                    nx, ny = x + nx, y + ny
                if current:
                    polylines.append(current)
                current = [point(nx, ny)]
                x, y = nx, ny
                start_x, start_y = nx, ny
            elif upper == "L":
                nx, ny = values
                if relative:
                    nx, ny = x + nx, y + ny
                current.append(point(nx, ny))
                x, y = nx, ny
            elif upper == "H":
                nx = values[0] + (x if relative else 0)
                current.append(point(nx, y))
                x = nx
            elif upper == "V":
                ny = values[0] + (y if relative else 0)
                current.append(point(x, ny))
                y = ny
            elif upper == "C":
                x1, y1, x2, y2, nx, ny = values
                if relative:
                    x1, y1, x2, y2, nx, ny = x + x1, y + y1, x + x2, y + y2, x + nx, y + ny
                for tick in range(1, 9):
                    t = tick / 8.0
                    mt = 1 - t
                    bx = mt ** 3 * x + 3 * mt ** 2 * t * x1 + 3 * mt * t ** 2 * x2 + t ** 3 * nx
                    by = mt ** 3 * y + 3 * mt ** 2 * t * y1 + 3 * mt * t ** 2 * y2 + t ** 3 * ny
                    current.append(point(bx, by))
                x, y = nx, ny
            elif upper == "Q":
                x1, y1, nx, ny = values
                if relative:
                    x1, y1, nx, ny = x + x1, y + y1, x + nx, y + ny
                for tick in range(1, 9):
                    t = tick / 8.0
                    mt = 1 - t
                    bx = mt ** 2 * x + 2 * mt * t * x1 + t ** 2 * nx
                    by = mt ** 2 * y + 2 * mt * t * y1 + t ** 2 * ny
                    current.append(point(bx, by))
                x, y = nx, ny
    if current:
        polylines.append(current)
    return polylines


def rasterise(svg: str, size: int, color: str) -> bytes | None:
    """Rasterise a Lucide-style stroke icon without external SVG libraries."""
    try:
        from PIL import Image, ImageDraw
    except Exception:
        return None
    width = int(re.search(r'viewBox="0 0 ([\d.]+) ([\d.]+)"', svg).group(1)) if 'viewBox' in svg else 24
    stroke = float((re.search(r'stroke-width="([\d.]+)"', svg) or [None, "2"])[1])
    rgb = tuple(int(color.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    scale = size / float(width)
    line = max(1, int(round(stroke * scale)))
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    drawn = 0
    for match in re.finditer(r'<path[^>]*d="([^"]+)"', svg):
        for polyline in _flatten(match.group(1), scale, line * 0.5):
            if len(polyline) >= 2:
                draw.line(polyline, fill=rgb + (255,), width=line, joint="curve")
                drawn += 1
    for match in re.finditer(r'<(circle|cx=")', svg):
        break
    for circle in re.finditer(r'<circle[^>]*cx="([\d.]+)"[^>]*cy="([\d.]+)"[^>]*r="([\d.]+)"', svg):
        cx, cy, radius = (float(circle.group(index)) * scale for index in (1, 2, 3))
        draw.ellipse([cx - radius, cy - radius, cx + radius, cy + radius], outline=rgb + (255,), width=line)
        drawn += 1
    if not drawn:
        return None
    buffer = __import__("io").BytesIO()
    image.save(buffer, "PNG")
    return buffer.getvalue()

def to_png(svg: str, size: int) -> bytes | None:
    try:
        import cairosvg  # type: ignore

        return cairosvg.svg2png(bytestring=svg.encode("utf-8"), output_width=size, output_height=size)
    except Exception:
        pass
    # Fallback: many systems have rsvg-convert or Inkscape on PATH
    for command in (["rsvg-convert", "-w", str(size), "-h", str(size)], ["inkscape", "--export-type=png"]):
        try:
            completed = subprocess.run(command, input=svg.encode("utf-8"), capture_output=True)
            if completed.returncode == 0 and completed.stdout:
                return completed.stdout
        except Exception:
            continue
    return rasterise(svg, size, "#000000")


def main() -> None:
    args = parse_args()
    if args.list:
        names = list_icons(args.query)
        print(json.dumps({"count": len(names), "icons": names[:120]}, ensure_ascii=False, indent=2))
        return
    if not args.name:
        raise SystemExit("--name is required unless --list is used")

    svg = recolour(get_svg(args.name), args.color)
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    svg_path = out_dir / f"{args.name}.svg"
    svg_path.write_text(svg, encoding="utf-8")
    result = {"name": args.name, "svg": str(svg_path), "color": args.color}

    if args.format == "png":
        payload = to_png(svg, args.size)
        if payload:
            png_path = out_dir / f"{args.name}.png"
            png_path.write_bytes(payload)
            result["png"] = str(png_path)
            result["size"] = args.size
        else:
            result["warning"] = (
                "No SVG rasteriser on this machine (cairosvg / rsvg-convert / Inkscape). "
                "The SVG is still usable; pass it to PowerPoint or render it elsewhere."
            )
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
