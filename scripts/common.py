from __future__ import annotations

import json
import math
import re
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw, ImageFont


SKILL_ROOT = Path(__file__).resolve().parent.parent
PROJECTION_TYPE_SCALE = {
    "cover_kicker": 30,
    "cover_title": 96,
    "cover_subtitle": 38,
    "cover_info": 34,
    "header_title": 54,
    "eyebrow": 54,
    "lead": 38,
    "lead_label": 24,
    "takeaway": 34,
    "sidebar": 32,
    "block_title": 32,
    "body": 28,
    "bullet": 24,
    "small": 21,
    "micro": 16,
    "metric_value": 78,
    "metric_label": 28,
}
PUNCTUATION = set("，。；：？！、）》」』】”’%）")
CLOSING_PUNCTUATION = set(".,;:!?)]}%\"'")
LATIN_CHUNK_RE = re.compile(
    r"[A-Za-zÀ-ÖØ-öø-ÿ0-9]+(?:[.'’\-–—/+×%:][A-Za-zÀ-ÖØ-öø-ÿ0-9]+)*|\s+|.",
    re.S,
)


def load_json(path: str | Path) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write_json(path: str | Path, data: dict) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def resolve_path(value: str | None, base: Path) -> Path | None:
    if not value:
        return None
    path = Path(value)
    if not path.is_absolute():
        path = base / path
    return path


def hex_rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    if len(value) != 6:
        raise ValueError(f"Expected six-digit color, got {value!r}")
    return tuple(int(value[index:index + 2], 16) for index in (0, 2, 4))


def mix_color(first: tuple[int, int, int], second: tuple[int, int, int], amount: float) -> tuple[int, int, int]:
    amount = max(0.0, min(1.0, amount))
    return tuple(round(a * (1 - amount) + b * amount) for a, b in zip(first, second))


def load_theme(path: str | Path) -> dict:
    theme_path = Path(path).resolve()
    theme = load_json(theme_path)
    theme["_path"] = str(theme_path)
    theme["_dir"] = str(theme_path.parent)
    theme["_colors"] = {key: hex_rgb(value) for key, value in theme["colors"].items()}
    lockup = theme.get("brand", {}).get("lockup")
    resolved_lockup = resolve_path(lockup, theme_path.parent)
    theme["brand"]["lockup_path"] = str(resolved_lockup) if resolved_lockup else None
    return theme


def apply_type_profile(theme: dict, profile: str | None = None) -> dict:
    """Apply a named type profile without changing layout coordinates.

    `document` keeps the existing theme values. `projection` raises the
    physical PowerPoint type scale so the same canvas remains legible in a
    lecture room. Theme authors can override either profile via
    `theme["type_profiles"][profile]["type_scale"]`.
    """
    name = str(profile or "document").strip().lower() or "document"
    profiles = theme.get("type_profiles") or {}
    configured = profiles.get(name) or {}
    if isinstance(configured, dict):
        scale = configured.get("type_scale") or configured
    else:
        scale = {}
    if not scale and name == "projection":
        scale = PROJECTION_TYPE_SCALE
    if isinstance(scale, dict) and scale:
        theme.setdefault("type_scale", {}).update(scale)
    theme["active_type_profile"] = name
    return theme


def resolve_font(candidates: Iterable[str], size: int) -> ImageFont.FreeTypeFont:
    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    raise FileNotFoundError("No usable font found. Add a CJK font path to the theme.")


def wrap_cjk_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    font: ImageFont.FreeTypeFont,
    max_width: int,
) -> list[str]:
    if not text:
        return []

    def keep_with_previous(token: str) -> bool:
        return token in PUNCTUATION or token in CLOSING_PUNCTUATION

    def split_long_token(token: str) -> list[str]:
        chunks: list[str] = []
        current = ""
        for character in token:
            candidate = current + character
            if current and draw.textbbox((0, 0), candidate, font=font)[2] > max_width:
                chunks.append(current)
                current = character
            else:
                current = candidate
        if current:
            chunks.append(current)
        return chunks

    lines: list[str] = []
    for paragraph in text.split("\n"):
        if not paragraph:
            lines.append("")
            continue
        current = ""
        pending_space = False
        for token in LATIN_CHUNK_RE.findall(paragraph):
            if token.isspace():
                pending_space = bool(current)
                continue
            joiner = " " if pending_space and current else ""
            candidate = current + joiner + token
            width = draw.textbbox((0, 0), candidate, font=font)[2]
            if width <= max_width:
                current = candidate
                pending_space = False
                continue
            if not current:
                pieces = split_long_token(token)
                lines.extend(pieces[:-1])
                current = pieces[-1] if pieces else ""
                pending_space = False
                continue
            if keep_with_previous(token) and len(token) <= 2:
                current += token
                lines.append(current)
                current = ""
            else:
                lines.append(current.rstrip())
                current = token
            pending_space = False
        if current:
            lines.append(current)
    return lines


def text_width(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont) -> int:
    return draw.textbbox((0, 0), text, font=font)[2]


def line_height(font: ImageFont.FreeTypeFont, ratio: float = 1.3) -> int:
    bbox = font.getbbox("国Ag")
    return max(1, round((bbox[3] - bbox[1]) * ratio))


def fit_text_block(
    draw: ImageDraw.ImageDraw,
    text: str,
    font_candidates: Iterable[str],
    max_width: int,
    max_height: int,
    max_size: int,
    min_size: int,
    max_lines: int | None = None,
    line_ratio: float = 1.32,
) -> tuple[ImageFont.FreeTypeFont, list[str], int, bool]:
    size = max_size
    while size >= min_size:
        font = resolve_font(font_candidates, size)
        lines = wrap_cjk_text(draw, text, font, max_width)
        height = line_height(font, line_ratio) * max(1, len(lines))
        if height <= max_height and (max_lines is None or len(lines) <= max_lines):
            return font, lines, height, True
        size -= 2
    font = resolve_font(font_candidates, min_size)
    lines = wrap_cjk_text(draw, text, font, max_width)
    return font, lines, line_height(font, line_ratio) * max(1, len(lines)), False


def draw_lines(
    draw: ImageDraw.ImageDraw,
    lines: list[str],
    xy: tuple[int, int],
    font: ImageFont.FreeTypeFont,
    fill: tuple[int, int, int],
    line_ratio: float = 1.32,
    align: str = "left",
    max_width: int | None = None,
) -> int:
    x, y = xy
    step = line_height(font, line_ratio)
    for line in lines:
        draw_x = x
        if align == "center" and max_width is not None:
            draw_x = x + (max_width - text_width(draw, line, font)) // 2
        draw.text((draw_x, y), line, font=font, fill=fill)
        y += step
    return y


def draw_bullet_list(
    draw: ImageDraw.ImageDraw,
    bullets: list[str],
    xy: tuple[int, int],
    width: int,
    height: int,
    font_candidates: Iterable[str],
    fill: tuple[int, int, int],
    bullet_color: tuple[int, int, int],
    max_size: int = 26,
    min_size: int = 20,
) -> tuple[bool, int]:
    x, y = xy
    if not bullets:
        return True, 0
    gap = 14
    usable_height = height - gap * (len(bullets) - 1)
    item_height = min(112, max(54, usable_height // max(1, len(bullets))))
    all_fit = True
    for bullet in bullets:
        text_x = x + 26
        text_width_available = width - 26
        font, lines, block_height, fits = fit_text_block(
            draw,
            bullet,
            font_candidates,
            text_width_available,
            item_height,
            max_size,
            min_size,
            max_lines=3,
            line_ratio=1.24,
        )
        all_fit = all_fit and fits
        radius = 5
        draw.ellipse((x + 4, y + 8, x + 4 + radius * 2, y + 8 + radius * 2), fill=bullet_color)
        draw_lines(draw, lines, (text_x, y), font, fill, 1.24)
        y += max(item_height, block_height) + gap
    return all_fit, y


def draw_wrapped_block(
    draw: ImageDraw.ImageDraw,
    text: str,
    xy: tuple[int, int],
    width: int,
    height: int,
    font_candidates: Iterable[str],
    fill: tuple[int, int, int],
    max_size: int,
    min_size: int,
    max_lines: int | None = None,
    align: str = "left",
    line_ratio: float = 1.3,
) -> tuple[bool, int]:
    font, lines, used_height, fits = fit_text_block(
        draw,
        text,
        font_candidates,
        width,
        height,
        max_size,
        min_size,
        max_lines=max_lines,
        line_ratio=line_ratio,
    )
    end = draw_lines(draw, lines, xy, font, fill, line_ratio, align, width)
    return fits, end - xy[1]


def draw_panel(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    fill: tuple[int, int, int],
    outline: tuple[int, int, int] | None = None,
    radius: int = 8,
    width: int = 2,
) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def paste_contain(
    canvas: Image.Image,
    image: Image.Image,
    box: tuple[int, int, int, int],
    background: tuple[int, int, int] = (255, 255, 255),
) -> tuple[int, int, int, int]:
    left, top, right, bottom = box
    width, height = right - left, bottom - top
    frame = Image.new("RGB", (width, height), background)
    image = image.convert("RGB")
    scale = min(width / image.width, height / image.height)
    size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
    resized = image.resize(size, Image.Resampling.LANCZOS)
    x = (width - size[0]) // 2
    y = (height - size[1]) // 2
    frame.paste(resized, (x, y))
    canvas.paste(frame, (left, top))
    return left + x, top + y, left + x + size[0], top + y + size[1]


def count_chinese_chars(text: str) -> int:
    return sum(1 for char in text if "\u4e00" <= char <= "\u9fff")


def estimate_seconds(text: str) -> int:
    characters = len(text.strip())
    chinese = count_chinese_chars(text)
    english_words = max(0, characters - chinese) / 5
    units = chinese + english_words
    return max(15, math.ceil(units / 3.8))


def format_duration(seconds: int) -> str:
    minutes, remainder = divmod(max(0, seconds), 60)
    return f"{minutes}m {remainder:02d}s"
