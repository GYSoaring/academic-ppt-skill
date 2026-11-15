from __future__ import annotations

import argparse
import io
import json
import urllib.parse
import urllib.request
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont

from common import hex_rgb, load_json, write_json


USER_AGENT = "academic-ppt-skill/0.1 (brand-asset builder)"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download or load a school emblem and compose a typographic brand lockup for a theme.",
    )
    parser.add_argument("--theme", required=True)
    parser.add_argument("--name", required=True, help="Institution name shown in the lockup.")
    parser.add_argument("--latin", default="", help="Optional Latin subtitle.")
    parser.add_argument("--emblem", default=None, help="Local image path or direct image URL.")
    parser.add_argument("--search-query", default=None, help="Search Wikimedia Commons for an emblem.")
    parser.add_argument("--out-dir", default=None)
    parser.add_argument("--font", default=None, help="Override the display font path.")
    parser.add_argument("--emblem-height", type=int, default=132)
    parser.add_argument("--keep-white", action="store_true", help="Do not convert near-white pixels to alpha.")
    parser.add_argument("--no-update-theme", action="store_true")
    return parser.parse_args()


def read_image(source: str) -> tuple[Image.Image, dict]:
    if source.startswith(("http://", "https://")):
        request = urllib.request.Request(source, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(request, timeout=45) as response:
            payload = response.read()
        image = Image.open(io.BytesIO(payload))
        return image.convert("RGBA"), {"kind": "url", "url": source}
    path = Path(source).resolve()
    if not path.exists():
        raise FileNotFoundError(path)
    return Image.open(path).convert("RGBA"), {"kind": "file", "path": str(path)}


def search_commons(query: str, limit: int = 8) -> list[dict]:
    params = {
        "action": "query",
        "generator": "search",
        "gsrsearch": query,
        "gsrnamespace": "6",
        "gsrlimit": str(limit),
        "prop": "imageinfo",
        "iiprop": "url|mime|size",
        "iiurlwidth": "1024",
        "format": "json",
    }
    url = "https://commons.wikimedia.org/w/api.php?" + urllib.parse.urlencode(params)
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=45) as response:
        payload = json.loads(response.read().decode("utf-8"))
    pages = (payload.get("query") or {}).get("pages") or {}
    results: list[dict] = []
    for page in pages.values():
        info = (page.get("imageinfo") or [{}])[0]
        title = str(page.get("title", ""))
        mime = str(info.get("mime", ""))
        candidate_url = str(info.get("thumburl") or info.get("url") or "")
        if not candidate_url or not mime.startswith("image/"):
            continue
        lowered = title.lower()
        score = 0
        if "logo" in lowered or "emblem" in lowered or "seal" in lowered:
            score += 3
        if "coat" in lowered or "badge" in lowered:
            score += 1
        if mime == "image/png":
            score += 2
        results.append(
            {
                "title": title,
                "url": candidate_url,
                "description_url": str(info.get("descriptionurl", "")),
                "mime": mime,
                "score": score,
            }
        )
    results.sort(key=lambda item: (-item["score"], item["title"]))
    return results


def white_to_alpha(image: Image.Image, threshold: int = 244, softness: int = 12) -> Image.Image:
    array = np.array(image.convert("RGBA")).astype(np.int16)
    rgb = array[:, :, :3]
    minimum = rgb.min(axis=2)
    if softness <= 0:
        alpha = np.where(minimum >= threshold, 0, array[:, :, 3])
    else:
        ramp = np.clip((threshold - minimum) / softness, 0.0, 1.0)
        alpha = (array[:, :, 3] * ramp).astype(np.int16)
    array[:, :, 3] = np.clip(alpha, 0, 255)
    return Image.fromarray(array.astype(np.uint8), "RGBA")


def crop_transparent(image: Image.Image, padding: int = 2) -> Image.Image:
    bbox = image.getchannel("A").getbbox()
    if bbox is None:
        return image
    left, top, right, bottom = bbox
    left = max(0, left - padding)
    top = max(0, top - padding)
    right = min(image.width, right + padding)
    bottom = min(image.height, bottom + padding)
    return image.crop((left, top, right, bottom))


def pick_font(theme: dict, explicit: str | None) -> Path:
    candidates = [explicit] if explicit else []
    fonts = theme.get("fonts") or {}
    for key in ("display", "title", "body"):
        candidates.extend(fonts.get(key) or [])
    for candidate in candidates:
        if not candidate:
            continue
        path = Path(candidate)
        if path.exists():
            return path
    raise FileNotFoundError("No usable display font found. Pass --font with a local font path.")


def pick_latin_font(theme: dict, fallback: Path) -> Path:
    for candidate in (theme.get("fonts") or {}).get("latin") or []:
        path = Path(candidate)
        if path.exists():
            return path
    return fallback


def compose_lockup(
    emblem: Image.Image,
    name: str,
    latin: str,
    font_path: Path,
    primary: tuple[int, int, int],
    accent: tuple[int, int, int],
    emblem_height: int = 132,
    latin_font_path: Path | None = None,
) -> Image.Image:
    scale = emblem_height / emblem.height
    emblem = emblem.resize(
        (max(1, round(emblem.width * scale)), emblem_height),
        Image.Resampling.LANCZOS,
    )
    title_size = max(28, round(emblem_height * 0.59))
    latin_size = max(11, round(emblem_height * 0.12))
    title_font = ImageFont.truetype(str(font_path), title_size)
    latin_font = ImageFont.truetype(str(latin_font_path or font_path), latin_size)
    probe = ImageDraw.Draw(Image.new("RGBA", (10, 10)))
    title_width = probe.textbbox((0, 0), name, font=title_font)[2]
    latin_width = probe.textbbox((0, 0), latin, font=latin_font)[2] if latin else 0
    gap = max(16, round(emblem_height * 0.16))
    padding = max(8, round(emblem_height * 0.11))
    text_width = max(title_width, latin_width)
    width = padding * 2 + emblem.width + gap + text_width
    height = max(emblem_height, title_size + 34) + padding * 2
    image = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    image.paste(emblem, (padding, (height - emblem.height) // 2), emblem)
    draw = ImageDraw.Draw(image)
    text_x = padding + emblem.width + gap
    title_top = padding
    title_box = probe.textbbox((text_x, title_top), name, font=title_font)
    draw.text((text_x, title_top), name, font=title_font, fill=(*primary, 255))
    rule_y = title_box[3] + 5
    draw.rectangle((text_x, rule_y, text_x + title_width, rule_y + 4), fill=(*accent, 215))
    if latin:
        draw.text((text_x, rule_y + 7), latin, font=latin_font, fill=(*accent, 235))
    return image


def build(
    theme_path: Path,
    name: str,
    emblem_source: str,
    latin: str = "",
    out_dir: Path | None = None,
    font: str | None = None,
    emblem_height: int = 132,
    remove_white: bool = True,
    update_theme: bool = True,
) -> dict:
    theme_path = theme_path.resolve()
    theme = load_json(theme_path)
    target_dir = (out_dir or theme_path.parent).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    emblem, source_meta = read_image(emblem_source)
    if remove_white:
        emblem = white_to_alpha(emblem)
    emblem = crop_transparent(emblem)
    emblem_path = target_dir / "emblem_mark.png"
    emblem.save(emblem_path)
    font_path = pick_font(theme, font)
    latin_font_path = pick_latin_font(theme, font_path)
    primary = hex_rgb(theme["colors"]["blue"])
    accent = hex_rgb(theme["colors"]["blue_mid"])
    lockup = compose_lockup(
        emblem,
        name,
        latin,
        font_path,
        primary,
        accent,
        emblem_height,
        latin_font_path=latin_font_path,
    )
    lockup_path = target_dir / "brand_lockup.png"
    lockup.save(lockup_path)
    source_meta.update(
        {
            "name": name,
            "latin": latin,
            "font": str(font_path),
            "latin_font": str(latin_font_path),
            "emblem_size": list(emblem.size),
            "lockup_size": list(lockup.size),
        }
    )
    write_json(target_dir / "BRAND_SOURCE.json", source_meta)
    if update_theme:
        theme.setdefault("brand", {})
        theme["brand"]["lockup"] = "brand_lockup.png"
        theme["brand"]["emblem"] = "emblem_mark.png"
        theme["brand"]["fallback_text"] = name
        theme_path.write_text(json.dumps(theme, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return {
        "theme": str(theme_path),
        "emblem": str(emblem_path),
        "lockup": str(lockup_path),
        "source": source_meta,
        "theme_updated": update_theme,
    }


def main() -> None:
    args = parse_args()
    emblem_source = args.emblem
    if not emblem_source:
        if not args.search_query:
            raise SystemExit("Provide --emblem or --search-query.")
        results = search_commons(args.search_query)
        if not results:
            raise SystemExit("No emblem found. Provide --emblem with a direct image URL.")
        emblem_source = results[0]["url"]
        print(json.dumps({"candidates": results, "selected": emblem_source}, ensure_ascii=False, indent=2))
    report = build(
        Path(args.theme),
        args.name,
        emblem_source,
        latin=args.latin,
        out_dir=Path(args.out_dir) if args.out_dir else None,
        font=args.font,
        emblem_height=args.emblem_height,
        remove_white=not args.keep_white,
        update_theme=not args.no_update_theme,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
