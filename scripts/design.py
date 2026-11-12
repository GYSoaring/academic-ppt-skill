from __future__ import annotations

import io
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

from common import (
    PUNCTUATION,
    draw_lines,
    fit_text_block,
    hex_rgb,
    line_height,
    mix_color,
    resolve_font,
    text_width,
)


@dataclass(frozen=True)
class TextStyle:
    bold: bool = False
    accent: bool = False
    underline: bool = False


@dataclass(frozen=True)
class Run:
    text: str
    style: TextStyle


@dataclass(frozen=True)
class Token:
    text: str
    style: TextStyle


MARKUP_RE = re.compile(r"(\*\*.*?\*\*|__.*?__|\[\[.*?\]\])", re.S)
TOKEN_RE = re.compile(r"[A-Za-z0-9]+(?:[\.\-+/×%][A-Za-z0-9]+)*|[\u4e00-\u9fff]|\s+|.", re.S)


class FontBook:
    def __init__(self, title_paths: Sequence[str], body_paths: Sequence[str]) -> None:
        self.title_paths = list(title_paths)
        self.body_paths = list(body_paths)
        self._cache: dict[tuple[bool, int], ImageFont.FreeTypeFont] = {}

    def get(self, style: TextStyle, size: int) -> ImageFont.FreeTypeFont:
        key = (bool(style.bold), int(size))
        if key not in self._cache:
            paths = self.title_paths if style.bold else self.body_paths
            self._cache[key] = resolve_font(paths, int(size))
        return self._cache[key]


def render_math_image(
    formula: str,
    color: tuple[int, int, int],
    font_size: int = 32,
    dpi: int = 300,
) -> Image.Image | None:
    """Render LaTeX-style math through matplotlib mathtext with a graceful fallback."""
    source = (formula or "").strip()
    if not source:
        return None
    if source.startswith("$") and source.endswith("$"):
        source = source[1:-1]
    try:
        import matplotlib

        from matplotlib.figure import Figure

        matplotlib.rcParams["mathtext.fontset"] = "cm"
        matplotlib.rcParams["font.family"] = "serif"
        figure = Figure(figsize=(6, 1.4), dpi=dpi)
        figure.patch.set_alpha(0)
        figure.text(
            0.02,
            0.12,
            f"${source}$",
            fontsize=font_size,
            color=tuple(channel / 255 for channel in color),
            ha="left",
            va="bottom",
        )
        buffer = io.BytesIO()
        figure.savefig(
            buffer,
            format="png",
            transparent=True,
            bbox_inches="tight",
            pad_inches=0.02,
        )
        buffer.seek(0)
        return Image.open(buffer).convert("RGBA")
    except Exception:
        return None


def parse_markup(text: str) -> list[Run]:
    runs: list[Run] = []
    for part in MARKUP_RE.split(text or ""):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            runs.append(Run(part[2:-2], TextStyle(bold=True, accent=True, underline=False)))
        elif part.startswith("__") and part.endswith("__") and len(part) >= 4:
            runs.append(Run(part[2:-2], TextStyle(bold=True, underline=True)))
        elif part.startswith("[[") and part.endswith("]]") and len(part) >= 4:
            runs.append(Run(part[2:-2], TextStyle(bold=True)))
        else:
            runs.append(Run(part, TextStyle()))
    return runs


def _iter_tokens(runs: Sequence[Run]) -> Iterable[Token | None]:
    for run in runs:
        if "\n" in run.text:
            chunks = run.text.split("\n")
            for index, chunk in enumerate(chunks):
                if index:
                    yield None
                for raw in TOKEN_RE.findall(chunk):
                    yield Token(raw, run.style)
        else:
            for raw in TOKEN_RE.findall(run.text):
                yield Token(raw, run.style)


def wrap_styled(
    draw: ImageDraw.ImageDraw,
    runs: Sequence[Run],
    fonts: FontBook,
    size: int,
    max_width: int,
) -> list[list[Token]]:
    lines: list[list[Token]] = []
    line: list[Token] = []
    width = 0.0
    for token in _iter_tokens(runs):
        if token is None:
            lines.append(line)
            line = []
            width = 0.0
            continue
        if token.text.isspace() and not line:
            continue
        font = fonts.get(token.style, size)
        token_width = float(text_width(draw, token.text, font))
        if line and width + token_width > max_width:
            if token.text in PUNCTUATION:
                line.append(token)
                lines.append(line)
                line = []
                width = 0.0
                continue
            lines.append(line)
            line = []
            width = 0.0
            if token.text.isspace():
                continue
        line.append(token)
        width += token_width
    if line:
        lines.append(line)
    return lines


def fit_styled(
    draw: ImageDraw.ImageDraw,
    text: str,
    fonts: FontBook,
    max_width: int,
    max_height: int,
    max_size: int,
    min_size: int,
    line_ratio: float = 1.32,
    max_lines: int | None = None,
) -> tuple[int, list[list[Token]], int, bool]:
    runs = parse_markup(text)
    if not any(run.text.strip() for run in runs):
        font = fonts.get(TextStyle(), min_size)
        return min_size, [], line_height(font, line_ratio), True
    size = max_size
    while size >= min_size:
        lines = wrap_styled(draw, runs, fonts, size, max_width)
        height = line_height(fonts.get(TextStyle(), size), line_ratio) * max(1, len(lines))
        if height <= max_height and (max_lines is None or len(lines) <= max_lines):
            return size, lines, height, True
        size -= 2
    size = min_size
    lines = wrap_styled(draw, runs, fonts, size, max_width)
    height = line_height(fonts.get(TextStyle(), size), line_ratio) * max(1, len(lines))
    return size, lines, height, False


def draw_styled(
    draw: ImageDraw.ImageDraw,
    lines: Sequence[Sequence[Token]],
    xy: tuple[int, int],
    size: int,
    fonts: FontBook,
    colors: dict[str, tuple[int, int, int]],
    align: str = "left",
    max_width: int | None = None,
    line_ratio: float = 1.32,
) -> int:
    x0, y = xy
    step = line_height(fonts.get(TextStyle(), size), line_ratio)
    for line in lines:
        widths = [float(text_width(draw, token.text, fonts.get(token.style, size))) for token in line]
        total = sum(widths)
        x = float(x0)
        if align == "center" and max_width is not None:
            x = x0 + (max_width - total) / 2
        elif align == "right" and max_width is not None:
            x = x0 + max_width - total
        for token, token_width in zip(line, widths):
            font = fonts.get(token.style, size)
            if token.style.accent:
                fill = colors["red"]
            elif token.style.bold:
                fill = colors["blue_dark"]
            else:
                fill = colors["ink"]
            draw.text((int(round(x)), int(y)), token.text, font=font, fill=fill)
            if token.style.underline and token.text.strip():
                underline_color = colors["red"] if token.style.accent else colors["blue"]
                text_bbox = font.getbbox(token.text)
                underline_offset = max(6, int(text_bbox[3]) + max(3, size // 12))
                draw.line(
                    (
                        int(round(x)),
                        int(y + underline_offset),
                        int(round(x + token_width)),
                        int(y + underline_offset),
                    ),
                    fill=underline_color,
                    width=max(2, size // 16),
                )
            x += token_width
        y += step
    return y - xy[1]


def draw_rich_block(
    draw: ImageDraw.ImageDraw,
    text: str,
    xy: tuple[int, int],
    width: int,
    height: int,
    fonts: FontBook,
    colors: dict[str, tuple[int, int, int]],
    max_size: int,
    min_size: int,
    max_lines: int | None = None,
    align: str = "left",
    line_ratio: float = 1.3,
) -> tuple[bool, int]:
    size, lines, used, fits = fit_styled(
        draw,
        text,
        fonts,
        width,
        height,
        max_size,
        min_size,
        line_ratio=line_ratio,
        max_lines=max_lines,
    )
    draw_styled(draw, lines, xy, size, fonts, colors, align, width, line_ratio)
    return fits, used


def draw_shadowed_panel(
    image: Image.Image,
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    fill: tuple[int, int, int],
    outline: tuple[int, int, int] | None = None,
    radius: int = 10,
    width: int = 2,
    shadow_offset: tuple[int, int] = (0, 7),
    shadow_blur: int = 12,
    shadow_alpha: int = 38,
) -> None:
    left, top, right, bottom = box
    w, h = right - left, bottom - top
    pad = shadow_blur * 2 + 8
    layer = Image.new("RGBA", (w + pad * 2, h + pad * 2), (0, 0, 0, 0))
    layer_draw = ImageDraw.Draw(layer)
    layer_draw.rounded_rectangle(
        (
            pad + shadow_offset[0],
            pad + shadow_offset[1],
            pad + w + shadow_offset[0],
            pad + h + shadow_offset[1],
        ),
        radius=radius,
        fill=(13, 45, 92, shadow_alpha),
    )
    layer = layer.filter(ImageFilter.GaussianBlur(shadow_blur))
    image.paste(layer, (left - pad, top - pad), layer)
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def draw_gradient_panel(
    image: Image.Image,
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    top_color: tuple[int, int, int],
    bottom_color: tuple[int, int, int],
    outline: tuple[int, int, int] | None = None,
    radius: int = 8,
    width: int = 2,
) -> None:
    left, top, right, bottom = box
    w, h = right - left, bottom - top
    if w <= 0 or h <= 0:
        return
    gradient = Image.new("RGB", (w, h))
    gradient_draw = ImageDraw.Draw(gradient)
    for row in range(h):
        amount = row / max(1, h - 1)
        gradient_draw.line((0, row, w, row), fill=mix_color(top_color, bottom_color, amount))
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, w - 1, h - 1), radius=radius, fill=255)
    image.paste(gradient, (left, top), mask)
    if outline is not None:
        draw.rounded_rectangle(box, radius=radius, outline=outline, width=width)


def draw_tag(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    font: ImageFont.FreeTypeFont,
    fill: tuple[int, int, int],
    text_fill: tuple[int, int, int],
    radius: int = 5,
    align: str = "center",
) -> None:
    left, top, right, bottom = box
    draw.rounded_rectangle(box, radius=radius, fill=fill)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    if align == "center":
        x = left + (right - left - tw) // 2
    elif align == "right":
        x = right - 10 - tw
    else:
        x = left + 10
    y = top + (bottom - top - th) // 2 - bbox[1]
    draw.text((x, y), text, font=font, fill=text_fill)


def draw_chevrons(
    draw: ImageDraw.ImageDraw,
    left: int,
    top: int,
    height: int,
    color: tuple[int, int, int],
    count: int = 2,
    width: int = 4,
    gap: int = 12,
) -> None:
    for index in range(count):
        x = left + index * gap
        draw.line((x, top, x + height * 0.45, top + height / 2), fill=color, width=width)
        draw.line((x + height * 0.45, top + height / 2, x, top + height), fill=color, width=width)


def draw_arrow(
    draw: ImageDraw.ImageDraw,
    start: tuple[int, int],
    end: tuple[int, int],
    color: tuple[int, int, int],
    width: int = 3,
    head: int = 12,
) -> None:
    draw.line((*start, *end), fill=color, width=width)
    angle = math.atan2(end[1] - start[1], end[0] - start[0])
    for delta in (math.pi * 0.78, -math.pi * 0.78):
        x = end[0] + head * math.cos(angle + delta)
        y = end[1] + head * math.sin(angle + delta)
        draw.line((*end, int(x), int(y)), fill=color, width=width)


def draw_dashed_line(
    draw: ImageDraw.ImageDraw,
    start: tuple[int, int],
    end: tuple[int, int],
    color: tuple[int, int, int],
    width: int = 2,
    dash: int = 10,
    gap: int = 7,
) -> None:
    x1, y1 = start
    x2, y2 = end
    length = math.hypot(x2 - x1, y2 - y1)
    if length <= 0:
        return
    ux, uy = (x2 - x1) / length, (y2 - y1) / length
    position = 0.0
    while position < length:
        end_position = min(length, position + dash)
        draw.line(
            (
                int(x1 + ux * position),
                int(y1 + uy * position),
                int(x1 + ux * end_position),
                int(y1 + uy * end_position),
            ),
            fill=color,
            width=width,
        )
        position += dash + gap


def draw_corner_marks(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    color: tuple[int, int, int],
    length: int = 18,
    width: int = 2,
) -> None:
    left, top, right, bottom = box
    for x, y, sx, sy in (
        (left, top, 1, 1),
        (right, top, -1, 1),
        (left, bottom, 1, -1),
        (right, bottom, -1, -1),
    ):
        draw.line((x, y, x + sx * length, y), fill=color, width=width)
        draw.line((x, y, x, y + sy * length), fill=color, width=width)


def _circle_icon(draw: ImageDraw.ImageDraw, box, color, width):
    left, top, right, bottom = box
    draw.ellipse(box, outline=color, width=width)


def draw_icon(
    draw: ImageDraw.ImageDraw,
    name: str,
    box: tuple[int, int, int, int],
    color: tuple[int, int, int],
    paper: tuple[int, int, int],
    stroke: int = 3,
) -> None:
    left, top, right, bottom = box
    w, h = right - left, bottom - top
    cx, cy = (left + right) / 2, (top + bottom) / 2
    s = min(w, h)
    if name in {"laser", "transmit"}:
        draw.rounded_rectangle(
            (left + s * 0.05, top + s * 0.28, left + s * 0.46, top + s * 0.67),
            radius=max(2, int(s * 0.05)),
            outline=color,
            width=stroke,
        )
        draw.line((left + s * 0.13, top + s * 0.36, left + s * 0.13, top + s * 0.58), fill=color, width=stroke)
        draw.line((left + s * 0.46, cy, right - s * 0.02, cy - s * 0.14), fill=color, width=stroke)
        draw.line((left + s * 0.46, cy, right - s * 0.02, cy + s * 0.14), fill=color, width=stroke)
        draw.line((left + s * 0.58, cy - s * 0.25, right - s * 0.02, cy - s * 0.36), fill=color, width=1)
        draw.line((left + s * 0.58, cy + s * 0.25, right - s * 0.02, cy + s * 0.36), fill=color, width=1)
    elif name in {"photon", "spad"}:
        draw.ellipse((cx - s * 0.08, cy - s * 0.08, cx + s * 0.08, cy + s * 0.08), fill=color)
        for radius in (0.22, 0.36):
            draw.arc(
                (cx - s * radius, cy - s * radius, cx + s * radius, cy + s * radius),
                start=205,
                end=335,
                fill=color,
                width=stroke,
            )
        draw.line((left + s * 0.04, top + s * 0.12, left + s * 0.24, top + s * 0.28), fill=color, width=stroke)
    elif name in {"clock", "time"}:
        draw.ellipse((left + s * 0.06, top + s * 0.06, right - s * 0.06, bottom - s * 0.06), outline=color, width=stroke)
        draw.line((cx, cy, cx, top + s * 0.22), fill=color, width=stroke)
        draw.line((cx, cy, left + s * 0.7, cy + s * 0.16), fill=color, width=stroke)
        draw.ellipse((cx - 3, cy - 3, cx + 3, cy + 3), fill=color)
    elif name in {"histogram", "chart"}:
        draw.line((left + s * 0.06, top + s * 0.05, left + s * 0.06, bottom - s * 0.06), fill=color, width=stroke)
        draw.line((left + s * 0.06, bottom - s * 0.06, right - s * 0.02, bottom - s * 0.06), fill=color, width=stroke)
        for index, height in enumerate((0.22, 0.42, 0.64, 0.48, 0.3)):
            x = left + s * (0.18 + index * 0.14)
            draw.rectangle((x, bottom - s * 0.06 - s * height, x + s * 0.08, bottom - s * 0.08), fill=color)
    elif name in {"check", "result"}:
        draw.ellipse((left + s * 0.04, top + s * 0.04, right - s * 0.04, bottom - s * 0.04), outline=color, width=stroke)
        draw.line((left + s * 0.28, cy, left + s * 0.44, bottom - s * 0.3, right - s * 0.2, top + s * 0.3), fill=color, width=stroke)
    elif name in {"target", "focus"}:
        draw.ellipse((left + s * 0.06, top + s * 0.06, right - s * 0.06, bottom - s * 0.06), outline=color, width=stroke)
        draw.ellipse((left + s * 0.28, top + s * 0.28, right - s * 0.28, bottom - s * 0.28), outline=color, width=stroke)
        draw.ellipse((cx - s * 0.03, cy - s * 0.03, cx + s * 0.03, cy + s * 0.03), fill=color)
        draw.line((cx, top + s * 0.02, cx, top + s * 0.2), fill=color, width=stroke)
        draw.line((cx, bottom - s * 0.2, cx, bottom - s * 0.02), fill=color, width=stroke)
        draw.line((left + s * 0.02, cy, left + s * 0.2, cy), fill=color, width=stroke)
        draw.line((right - s * 0.2, cy, right - s * 0.02, cy), fill=color, width=stroke)
    elif name in {"layers", "fold"}:
        draw.polygon(
            [(cx, top + s * 0.05), (right - s * 0.05, top + s * 0.27), (cx, top + s * 0.49), (left + s * 0.05, top + s * 0.27)],
            outline=color,
            width=stroke,
        )
        draw.line((left + s * 0.05, top + s * 0.52, cx, top + s * 0.74, right - s * 0.05, top + s * 0.52), fill=color, width=stroke)
        draw.line((left + s * 0.05, top + s * 0.72, cx, top + s * 0.94, right - s * 0.05, top + s * 0.72), fill=color, width=stroke)
    elif name in {"signal", "wave"}:
        points = [
            (left, top + s * 0.58),
            (left + s * 0.18, top + s * 0.38),
            (left + s * 0.36, top + s * 0.68),
            (left + s * 0.54, top + s * 0.2),
            (left + s * 0.72, top + s * 0.6),
            (right, top + s * 0.32),
        ]
        draw.line(points, fill=color, width=stroke, joint="curve")
    elif name in {"atom", "sensitivity"}:
        draw.ellipse((cx - s * 0.07, cy - s * 0.07, cx + s * 0.07, cy + s * 0.07), fill=color)
        draw.ellipse((left + s * 0.06, top + s * 0.28, right - s * 0.06, bottom - s * 0.28), outline=color, width=stroke)
        draw.ellipse((left + s * 0.28, top + s * 0.06, right - s * 0.28, bottom - s * 0.06), outline=color, width=stroke)
        draw.ellipse((left + s * 0.16, top + s * 0.16, right - s * 0.16, bottom - s * 0.16), outline=color, width=1)
    elif name in {"mountain", "range"}:
        draw.line((left + s * 0.04, bottom - s * 0.08, left + s * 0.36, top + s * 0.18, left + s * 0.58, bottom - s * 0.08), fill=color, width=stroke)
        draw.line((left + s * 0.36, top + s * 0.18, left + s * 0.52, top + s * 0.42, left + s * 0.66, bottom - s * 0.08), fill=color, width=stroke)
        draw.line((left + s * 0.58, bottom - s * 0.08, right - s * 0.02, top + s * 0.34, right - s * 0.02, bottom - s * 0.08), fill=color, width=stroke)
    elif name in {"sun", "background"}:
        draw.ellipse((cx - s * 0.15, cy - s * 0.15, cx + s * 0.15, cy + s * 0.15), outline=color, width=stroke)
        for angle in range(0, 360, 45):
            rad = math.radians(angle)
            draw.line(
                (
                    int(cx + math.cos(rad) * s * 0.25),
                    int(cy + math.sin(rad) * s * 0.25),
                    int(cx + math.cos(rad) * s * 0.38),
                    int(cy + math.sin(rad) * s * 0.38),
                ),
                fill=color,
                width=stroke,
            )
    elif name in {"moon", "dark"}:
        draw.ellipse((cx - s * 0.24, cy - s * 0.24, cx + s * 0.24, cy + s * 0.24), fill=color)
        draw.ellipse((cx - s * 0.12, cy - s * 0.3, cx + s * 0.34, cy + s * 0.16), fill=paper)
    elif name in {"shield", "robust"}:
        points = [
            (cx, top + s * 0.04),
            (right - s * 0.08, top + s * 0.2),
            (right - s * 0.12, bottom - s * 0.28),
            (cx, bottom - s * 0.04),
            (left + s * 0.12, bottom - s * 0.28),
            (left + s * 0.08, top + s * 0.2),
        ]
        draw.polygon(points, outline=color, width=stroke)
        draw.line(
            (left + s * 0.2, cy, left + s * 0.34, cy + s * 0.08, left + s * 0.48, cy - s * 0.14, left + s * 0.62, cy + s * 0.08, right - s * 0.2, cy),
            fill=color,
            width=stroke,
            joint="curve",
        )
    elif name in {"database", "memory"}:
        draw.ellipse((left + s * 0.12, top + s * 0.06, right - s * 0.12, top + s * 0.3), outline=color, width=stroke)
        draw.line((left + s * 0.12, top + s * 0.18, left + s * 0.12, bottom - s * 0.2), fill=color, width=stroke)
        draw.line((right - s * 0.12, top + s * 0.18, right - s * 0.12, bottom - s * 0.2), fill=color, width=stroke)
        draw.arc((left + s * 0.12, bottom - s * 0.3, right - s * 0.12, bottom - s * 0.06), 0, 180, fill=color, width=stroke)
        for y in (0.45, 0.64):
            draw.arc((left + s * 0.12, top + s * y, right - s * 0.12, top + s * (y + 0.16)), 0, 180, fill=color, width=1)
    elif name in {"chip", "fpga"}:
        draw.rounded_rectangle((left + s * 0.18, top + s * 0.18, right - s * 0.18, bottom - s * 0.18), radius=max(2, int(s * 0.04)), outline=color, width=stroke)
        for index in range(3):
            offset = left + s * (0.3 + index * 0.18)
            draw.line((offset, top + s * 0.08, offset, top + s * 0.18), fill=color, width=stroke)
            draw.line((offset, bottom - s * 0.18, offset, bottom - s * 0.08), fill=color, width=stroke)
            draw.line((left + s * 0.08, top + s * (0.3 + index * 0.18), left + s * 0.18, top + s * (0.3 + index * 0.18)), fill=color, width=stroke)
            draw.line((right - s * 0.18, top + s * (0.3 + index * 0.18), right - s * 0.08, top + s * (0.3 + index * 0.18)), fill=color, width=stroke)
    elif name in {"book", "review"}:
        draw.line((cx, top + s * 0.16, cx, bottom - s * 0.1), fill=color, width=stroke)
        draw.line((cx, top + s * 0.16, left + s * 0.1, top + s * 0.22, left + s * 0.1, bottom - s * 0.16, cx, bottom - s * 0.1), fill=color, width=stroke)
        draw.line((cx, top + s * 0.16, right - s * 0.1, top + s * 0.22, right - s * 0.1, bottom - s * 0.16, cx, bottom - s * 0.1), fill=color, width=stroke)
    elif name in {"alert", "boundary"}:
        draw.polygon([(cx, top + s * 0.04), (right - s * 0.04, bottom - s * 0.08), (left + s * 0.04, bottom - s * 0.08)], outline=color, width=stroke)
        draw.line((cx, top + s * 0.3, cx, bottom - s * 0.32), fill=color, width=stroke)
        draw.ellipse((cx - 3, bottom - s * 0.22, cx + 3, bottom - s * 0.14), fill=color)
    elif name in {"matrix", "compare"}:
        cell = s * 0.23
        for row in range(3):
            for col in range(3):
                x = left + s * 0.1 + col * cell
                y = top + s * 0.1 + row * cell
                draw.rounded_rectangle((x, y, x + cell * 0.7, y + cell * 0.7), radius=2, outline=color, width=2)
    elif name in {"route", "pipeline"}:
        nodes = [(left + s * 0.12, cy), (cx, cy), (right - s * 0.12, cy)]
        for x, y in nodes:
            draw.ellipse((x - s * 0.07, y - s * 0.07, x + s * 0.07, y + s * 0.07), fill=color)
        draw.line((nodes[0][0] + s * 0.08, cy, nodes[1][0] - s * 0.08, cy), fill=color, width=stroke)
        draw.line((nodes[1][0] + s * 0.08, cy, nodes[2][0] - s * 0.08, cy), fill=color, width=stroke)
    elif name in {"formula", "math"}:
        font = resolve_font([str(Path("C:/Windows/Fonts/cambria.ttc"))], max(16, int(s * 0.42)))
        bbox = draw.textbbox((0, 0), "f(x)", font=font)
        draw.text((cx - (bbox[2] - bbox[0]) / 2, cy - (bbox[3] - bbox[1]) / 2 - bbox[1]), "f(x)", font=font, fill=color)
    elif name in {"arrow", "next"}:
        draw.line((left + s * 0.08, cy, right - s * 0.1, cy), fill=color, width=stroke)
        draw.line((right - s * 0.28, cy - s * 0.16, right - s * 0.08, cy), fill=color, width=stroke)
        draw.line((right - s * 0.28, cy + s * 0.16, right - s * 0.08, cy), fill=color, width=stroke)
    else:
        draw.ellipse((cx - s * 0.14, cy - s * 0.14, cx + s * 0.14, cy + s * 0.14), fill=color)


def draw_icon_badge(
    image: Image.Image,
    draw: ImageDraw.ImageDraw,
    center: tuple[int, int],
    size: int,
    name: str,
    colors: dict[str, tuple[int, int, int]],
    icon_color: tuple[int, int, int] | None = None,
    fill: tuple[int, int, int] | None = None,
    outline: tuple[int, int, int] | None = None,
) -> None:
    cx, cy = center
    radius = size // 2
    box = (cx - radius, cy - radius, cx + radius, cy + radius)
    if fill is None:
        fill = colors["blue_wash"]
    if outline is None:
        outline = colors["line"]
    draw.ellipse(box, fill=fill, outline=outline, width=2)
    inner = int(size * 0.58)
    draw_icon(
        draw,
        name,
        (cx - inner // 2, cy - inner // 2, cx + inner // 2, cy + inner // 2),
        icon_color or colors["blue"],
        colors["paper"],
        stroke=max(2, size // 22),
    )


def draw_chart(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    chart: dict,
    colors: dict[str, tuple[int, int, int]],
    fonts: FontBook,
    small_size: int = 16,
) -> None:
    left, top, right, bottom = box
    draw.rounded_rectangle(box, radius=7, fill=colors["white"], outline=colors["line"], width=1)
    plot_left = left + 58
    plot_right = right - 26
    plot_top = top + 32
    plot_bottom = bottom - 52
    if plot_right <= plot_left or plot_bottom <= plot_top:
        return
    grid = mix_color(colors["line"], colors["paper"], 0.55)
    for index in range(5):
        y = plot_top + (plot_bottom - plot_top) * index / 4
        draw.line((plot_left, int(y), plot_right, int(y)), fill=grid, width=1)
    for index in range(4):
        x = plot_left + (plot_right - plot_left) * index / 3
        draw.line((int(x), plot_top, int(x), plot_bottom), fill=grid, width=1)
    axis_color = mix_color(colors["blue_dark"], colors["ink"], 0.2)
    draw.line((plot_left, plot_bottom, plot_right, plot_bottom), fill=axis_color, width=2)
    draw.line((plot_left, plot_top, plot_left, plot_bottom), fill=axis_color, width=2)
    label_font = fonts.get(TextStyle(), small_size)
    muted = colors["muted"]
    kind = str(chart.get("type", "histogram"))
    values = [float(value) for value in chart.get("values", [])]
    if not values:
        values = [0.1, 0.2, 0.4, 0.7, 0.45, 0.25, 0.12]
    maximum = max(values) or 1.0
    highlight = int(chart.get("highlight", max(range(len(values)), key=lambda index: values[index])))
    highlight = max(0, min(len(values) - 1, highlight))
    if kind == "curve":
        minimum = min(values)
        span = (maximum - minimum) or 1.0
        points = []
        for index, value in enumerate(values):
            x = plot_left + (plot_right - plot_left) * (index / max(1, len(values) - 1))
            y = plot_bottom - (plot_bottom - plot_top) * (value - minimum) / span
            points.append((int(x), int(y)))
        draw.line(points, fill=colors["blue_mid"], width=4, joint="curve")
        for index, (x, y) in enumerate(points):
            draw.ellipse((x - 5, y - 5, x + 5, y + 5), fill=colors["red"] if index == highlight else colors["blue"], outline=colors["white"], width=2)
    else:
        slot = (plot_right - plot_left) / max(1, len(values))
        bar_width = max(6, slot * 0.62)
        for index, value in enumerate(values):
            x = plot_left + index * slot + (slot - bar_width) / 2
            y = plot_bottom - (plot_bottom - plot_top) * value / (maximum * 1.12)
            fill = colors["red"] if index == highlight else colors["blue_mid"]
            draw.rounded_rectangle((int(x), int(y), int(x + bar_width), plot_bottom), radius=2, fill=fill)
            if index == highlight:
                draw.rounded_rectangle((int(x), int(y), int(x + bar_width), int(y + 8)), radius=2, fill=mix_color(fill, colors["white"], 0.28))
        highlight_x = plot_left + highlight * slot + slot / 2
        draw_dashed_line(
            draw,
            (int(highlight_x), plot_top),
            (int(highlight_x), plot_bottom),
            colors["red"],
            width=2,
            dash=8,
            gap=6,
        )
    x_label = str(chart.get("x_label", ""))
    y_label = str(chart.get("y_label", ""))
    if x_label:
        width = text_width(draw, x_label, label_font)
        draw.text((plot_right - width, bottom - 32), x_label, font=label_font, fill=muted)
    if y_label:
        draw.text((left + 14, top + 10), y_label, font=label_font, fill=muted)
    if chart.get("annotation"):
        annotation = str(chart["annotation"])
        width = text_width(draw, annotation, label_font)
        x = plot_left + highlight * ((plot_right - plot_left) / max(1, len(values))) + 14
        draw.text((int(min(x, plot_right - width - 8)), plot_top + 8), annotation, font=label_font, fill=colors["red"])


def draw_mini_visual(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    kind: str,
    colors: dict[str, tuple[int, int, int]],
    fonts: FontBook,
) -> None:
    left, top, right, bottom = box
    if right - left < 40 or bottom - top < 30:
        return
    width, height = right - left, bottom - top
    draw.rounded_rectangle(
        box,
        radius=6,
        fill=mix_color(colors["blue_wash"], colors["white"], 0.28),
        outline=mix_color(colors["line"], colors["white"], 0.22),
        width=1,
    )
    pad_x = max(12, int(width * 0.06))
    pad_y = max(8, int(height * 0.12))
    plot_left, plot_right = left + pad_x, right - pad_x
    plot_top, plot_bottom = top + pad_y, bottom - pad_y
    baseline = plot_bottom - max(8, int(height * 0.12))
    ink = colors["blue"]
    mid = colors["blue_mid"]
    pale = mix_color(colors["blue_mid"], colors["white"], 0.55)
    if kind in {"pulse", "timeline"}:
        draw.line((plot_left, baseline, plot_right, baseline), fill=colors["line"], width=2)
        count = 9 if kind == "pulse" else 6
        for index in range(count):
            x = plot_left + (plot_right - plot_left) * (index + 0.5) / count
            if kind == "timeline":
                color = colors["red"] if index == 3 else mid
                draw.ellipse((x - 5, baseline - 5, x + 5, baseline + 5), fill=color, outline=colors["white"], width=2)
            else:
                pulse_height = (0.45 + 0.3 * ((index * 7) % 5) / 4) * (baseline - plot_top)
                color = colors["red"] if index == 6 else mid
                draw.line((x, baseline, x, baseline - pulse_height), fill=color, width=3)
                draw.line((x - 4, baseline - pulse_height, x + 4, baseline - pulse_height), fill=color, width=2)
        if kind == "pulse":
            x = plot_left + (plot_right - plot_left) * 6.5 / count
            draw_dashed_line(draw, (int(x), plot_top), (int(x), baseline), colors["red"], width=1, dash=5, gap=4)
    elif kind in {"events", "noise"}:
        draw.line((plot_left, baseline, plot_right, baseline), fill=colors["line"], width=2)
        for index in range(14):
            x = plot_left + (plot_right - plot_left) * (index + 0.5) / 14
            if kind == "noise":
                bar = (0.18 + 0.62 * ((index * 3) % 7) / 6) * (baseline - plot_top)
                draw.rectangle((x - 4, baseline - bar, x + 4, baseline), fill=pale if index % 4 else mid)
            else:
                y = baseline - ((index * 5) % 9) / 9 * (baseline - plot_top) * 0.75
                draw.ellipse((x - 3, y - 3, x + 3, y + 3), fill=colors["red"] if index in {4, 9} else mid)
        if kind == "events" and plot_right - plot_left > 60:
            draw_arrow(draw, (int(plot_left + 5), int(baseline - (baseline - plot_top) * 0.75)), (int(plot_right - 8), int(plot_top + 6)), colors["red"], width=1, head=5)
    elif kind in {"histogram", "peak"}:
        count = 10
        if kind == "peak":
            values = [0.15, 0.25, 0.42, 0.7, 1.0, 0.68, 0.39, 0.24, 0.15, 0.1]
        else:
            values = [0.14, 0.24, 0.38, 0.56, 0.82, 0.55, 0.35, 0.22, 0.16, 0.11]
        slot = (plot_right - plot_left) / count
        bar_width = max(4, slot * 0.62)
        for index, value in enumerate(values):
            x = plot_left + index * slot + (slot - bar_width) / 2
            y = baseline - value * (baseline - plot_top)
            draw.rounded_rectangle((int(x), int(y), int(x + bar_width), baseline), radius=1, fill=colors["red"] if index == 4 else mid)
        peak_x = plot_left + 4.5 * slot
        draw_dashed_line(draw, (int(peak_x), plot_top), (int(peak_x), baseline), colors["red"], width=1, dash=5, gap=4)
    elif kind in {"target", "range"}:
        cx, cy = (plot_left + plot_right) / 2, (plot_top + baseline) / 2
        radius = min((plot_right - plot_left) * 0.18, (baseline - plot_top) * 0.4)
        draw.line((plot_left, cy, plot_right, cy), fill=mix_color(mid, colors["white"], 0.48), width=1)
        draw.line((cx, plot_top, cx, baseline), fill=mix_color(mid, colors["white"], 0.48), width=1)
        for factor in (1.0, 0.68, 0.38):
            draw.ellipse((cx - radius * factor, cy - radius * factor, cx + radius * factor, cy + radius * factor), outline=mid if factor < 1.0 else colors["red"], width=2 if factor == 1.0 else 1)
        for dx, dy in ((-1, -1), (1, -1), (-1, 1), (1, 1)):
            x = cx + dx * radius * 1.24
            y = cy + dy * radius * 1.24
            draw.line((x, y, x - dx * 12, y), fill=mid, width=2)
            draw.line((x, y, x, y - dy * 12), fill=mid, width=2)
        if kind == "range":
            rule_y = baseline - 2
            draw.line((plot_left, rule_y, plot_right, rule_y), fill=colors["blue"], width=2)
            for index in range(9):
                x = plot_left + (plot_right - plot_left) * index / 8
                tick = 8 if index % 2 == 0 else 4
                draw.line((x, rule_y, x, rule_y - tick), fill=colors["blue"], width=1)
            draw.line((plot_left + 6, baseline - 2, cx - radius, cy + radius * 0.28), fill=colors["red"], width=2)
            draw.text((plot_left + 4, top + 2), "R0", font=fonts.get(TextStyle(bold=True), max(10, min(15, int(height * 0.16)))), fill=colors["red"])
    elif kind in {"nodes", "pipeline"}:
        labels = ["01", "02", "03"]
        node_w = (plot_right - plot_left) * 0.2
        gap = ((plot_right - plot_left) - node_w * 3) / 2
        cy = (plot_top + baseline) / 2
        for index, label in enumerate(labels):
            x = plot_left + index * (node_w + gap)
            draw.rounded_rectangle((x, cy - 15, x + node_w, cy + 15), radius=5, fill=colors["blue_wash"], outline=mid, width=1)
            draw.text((x + node_w * 0.3, cy - 9), label, font=fonts.get(TextStyle(bold=True), max(10, min(14, int(height * 0.15)))), fill=colors["blue_dark"])
            if index < 2:
                draw_arrow(draw, (int(x + node_w + 4), int(cy)), (int(x + node_w + gap - 4), int(cy)), mid, width=1, head=5)
    elif kind == "formula":
        draw.text((left + 14, top + 8), "f(x)", font=fonts.get(TextStyle(bold=True), max(18, int(height * 0.3))), fill=mid)
        draw.line((left + 14, bottom - 14, right - 14, bottom - 14), fill=colors["line"], width=1)
    else:
        draw.line((plot_left, baseline, plot_right, baseline), fill=colors["line"], width=2)
        draw_chevrons(draw, plot_left + 8, baseline - 26, 18, mid, count=2, width=2, gap=7)


def _diagram_fit_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    box: tuple[int, int, int, int],
    fonts: FontBook,
    fill: tuple[int, int, int],
    max_size: int,
    min_size: int,
    bold: bool = False,
    align: str = "left",
    max_lines: int = 2,
    line_ratio: float = 1.16,
    valign: str = "top",
) -> tuple[bool, int]:
    left, top, right, bottom = box
    candidates = fonts.title_paths if bold else fonts.body_paths
    font, lines, block_height, fits = fit_text_block(
        draw,
        text,
        candidates,
        max(20, right - left),
        max(20, bottom - top),
        max_size,
        min_size,
        max_lines=max_lines,
        line_ratio=line_ratio,
    )
    if valign == "middle":
        y = top + max(0, (bottom - top - block_height) // 2)
    elif valign == "bottom":
        y = max(top, bottom - block_height)
    else:
        y = top
    draw_lines(
        draw,
        lines,
        (left, y),
        font,
        fill,
        line_ratio,
        align=align,
        max_width=right - left,
    )
    return fits, block_height


def _diagram_index(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int],
    index: int,
    colors: dict[str, tuple[int, int, int]],
    fonts: FontBook,
    dark: bool = False,
) -> None:
    font = fonts.get(TextStyle(bold=True), 13)
    text = f"{index:02d}"
    width = text_width(draw, text, font) + 16
    left, top = xy
    fill = colors["blue"] if dark else colors["blue_pale"]
    text_fill = colors["white"] if dark else colors["blue_dark"]
    draw.rounded_rectangle((left, top, left + width, top + 24), radius=3, fill=fill)
    draw.text((left + 8, top + 4), text, font=font, fill=text_fill)


def draw_diagram(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    diagram: dict,
    colors: dict[str, tuple[int, int, int]],
    fonts: FontBook,
    label: str = "",
    frame: bool = True,
) -> None:
    """Draw a compact, data-driven academic concept diagram.

    The grammar is intentionally semantic rather than decorative. It supports
    sequence, branch, stack, ruler, matrix, and cycle structures. Labels and
    values come from deck.json; the renderer only owns geometry and typography.
    """
    if not diagram:
        return
    left, top, right, bottom = box
    if right - left < 100 or bottom - top < 80:
        return
    if frame:
        draw.rounded_rectangle(
            box,
            radius=6,
            fill=mix_color(colors["blue_wash"], colors["white"], 0.2),
            outline=colors["line"],
            width=1,
        )
    pad = max(10, min(18, int((right - left) * 0.018)))
    inner_left, inner_right = left + pad, right - pad
    inner_top, inner_bottom = top + pad, bottom - pad
    if label:
        label_font = fonts.get(TextStyle(bold=True), 13)
        label_width = text_width(draw, label, label_font) + 22
        label_box = (inner_left, inner_top, inner_left + label_width, inner_top + 25)
        draw.rounded_rectangle(label_box, radius=3, fill=colors["blue"])
        draw.text((label_box[0] + 10, label_box[1] + 4), label, font=label_font, fill=colors["white"])
        inner_top += 34
    note = str(diagram.get("note", "")).strip()
    note_height = 25 if note else 0
    if note:
        inner_bottom -= 30
    kind = str(diagram.get("kind", "sequence")).strip().lower()
    items = [item for item in (diagram.get("items") or []) if isinstance(item, dict)]
    highlight = int(diagram.get("highlight", -1))
    blue = colors["blue"]
    blue_dark = colors["blue_dark"]
    blue_mid = colors["blue_mid"]
    blue_pale = colors["blue_pale"]
    blue_wash = colors["blue_wash"]
    red = colors["red"]
    ink = colors["ink"]
    muted = colors["muted"]
    white = colors["white"]

    if kind == "sequence":
        items = items[:5]
        count = max(1, len(items))
        gap = max(20, min(34, int((inner_right - inner_left) * 0.035)))
        node_width = max(64, (inner_right - inner_left - gap * (count - 1)) // count)
        node_height = min(112, max(72, int((inner_bottom - inner_top) * 0.56)))
        node_top = inner_top + max(0, (inner_bottom - inner_top - node_height) // 2)
        centers: list[tuple[int, int]] = []
        for index, item in enumerate(items):
            x = inner_left + index * (node_width + gap)
            is_focus = index == highlight
            node_box = (x, node_top, x + node_width, node_top + node_height)
            draw.rounded_rectangle(
                node_box,
                radius=5,
                fill=blue if is_focus else white,
                outline=blue if is_focus else colors["line"],
                width=2 if is_focus else 1,
            )
            if is_focus:
                draw.rectangle((x + 1, node_top + 1, x + node_width - 1, node_top + 5), fill=blue_mid)
            _diagram_index(
                draw,
                (x + 10, node_top + 10),
                index + 1,
                colors,
                fonts,
                dark=is_focus,
            )
            text_fill = white if is_focus else blue_dark
            _diagram_fit_text(
                draw,
                str(item.get("title") or item.get("label") or ""),
                (x + 10, node_top + 40, x + node_width - 10, node_top + 68),
                fonts,
                text_fill,
                18,
                12,
                bold=True,
                align="center",
                max_lines=2,
                line_ratio=1.05,
            )
            body = str(item.get("body", "")).strip()
            if body:
                _diagram_fit_text(
                    draw,
                    body,
                    (x + 10, node_top + 70, x + node_width - 10, node_top + node_height - 8),
                    fonts,
                    mix_color(white, blue_pale, 0.42) if is_focus else muted,
                    14,
                    10,
                    align="center",
                    max_lines=2,
                    line_ratio=1.08,
                    valign="top",
                )
            centers.append((x + node_width // 2, node_top + node_height // 2))
        for start, end in zip(centers[:-1], centers[1:]):
            draw_arrow(
                draw,
                (start[0] + node_width // 2 + 4, start[1]),
                (end[0] - node_width // 2 - 4, end[1]),
                blue_mid,
                width=2,
                head=7,
            )
    elif kind == "branch":
        items = items[:4]
        root = diagram.get("root") if isinstance(diagram.get("root"), dict) else {}
        root_width = int((inner_right - inner_left) * 0.31)
        root_box = (inner_left, inner_top + 18, inner_left + root_width, inner_bottom - 18)
        draw.rounded_rectangle(
            root_box,
            radius=5,
            fill=blue,
            outline=blue,
            width=1,
        )
        _diagram_fit_text(
            draw,
            str(root.get("title") or diagram.get("center") or "共同问题"),
            (root_box[0] + 12, root_box[1] + 16, root_box[2] - 12, root_box[1] + 58),
            fonts,
            white,
            24,
            14,
            bold=True,
            align="center",
            max_lines=2,
        )
        root_body = str(root.get("body", "")).strip()
        if root_body:
            _diagram_fit_text(
                draw,
                root_body,
                (root_box[0] + 12, root_box[1] + 66, root_box[2] - 12, root_box[3] - 12),
                fonts,
                mix_color(white, blue_pale, 0.48),
                16,
                11,
                align="center",
                max_lines=3,
            )
        gap = 12
        item_left = inner_left + root_width + max(34, int((inner_right - inner_left) * 0.08))
        item_width = inner_right - item_left
        item_height = max(48, (inner_bottom - inner_top - gap * (len(items) - 1)) // max(1, len(items)))
        root_center = ((root_box[0] + root_box[2]) // 2, (root_box[1] + root_box[3]) // 2)
        for index, item in enumerate(items):
            y = inner_top + index * (item_height + gap)
            is_focus = index == highlight
            item_box = (item_left, y, item_left + item_width, y + item_height)
            draw.rounded_rectangle(
                item_box,
                radius=5,
                fill=blue_wash if not is_focus else blue_pale,
                outline=red if is_focus else colors["line"],
                width=2 if is_focus else 1,
            )
            _diagram_index(draw, (item_left + 10, y + 9), index + 1, colors, fonts, dark=is_focus)
            _diagram_fit_text(
                draw,
                str(item.get("title") or item.get("label") or ""),
                (item_left + 52, y + 8, item_box[2] - 10, y + 34),
                fonts,
                blue_dark,
                17,
                12,
                bold=True,
                max_lines=1,
            )
            body = str(item.get("body", "")).strip()
            if body:
                _diagram_fit_text(
                    draw,
                    body,
                    (item_left + 52, y + 35, item_box[2] - 10, item_box[3] - 7),
                    fonts,
                    ink,
                    14,
                    10,
                    max_lines=2,
                    line_ratio=1.08,
                )
            draw_arrow(
                draw,
                (root_center[0] + root_width // 2 - 8, root_center[1]),
                (item_left - 6, y + item_height // 2),
                blue_mid,
                width=2,
                head=7,
            )
    elif kind == "stack":
        items = items[:5]
        gap = 8
        row_height = max(38, (inner_bottom - inner_top - gap * (len(items) - 1)) // max(1, len(items)))
        for index, item in enumerate(items):
            y = inner_top + index * (row_height + gap)
            is_focus = index == highlight
            indent = min(34, index * 10)
            row_box = (inner_left + indent, y, inner_right, y + row_height)
            draw.rounded_rectangle(
                row_box,
                radius=4,
                fill=blue if is_focus else (white if index % 2 == 0 else blue_wash),
                outline=blue_mid if is_focus else colors["line"],
                width=1,
            )
            draw.rectangle((row_box[0], row_box[1], row_box[0] + 5, row_box[3]), fill=blue if not is_focus else red)
            _diagram_index(
                draw,
                (row_box[0] + 12, row_box[1] + max(7, (row_height - 24) // 2)),
                index + 1,
                colors,
                fonts,
                dark=is_focus,
            )
            text_left = row_box[0] + 64
            text_color = white if is_focus else blue_dark
            _diagram_fit_text(
                draw,
                str(item.get("title") or item.get("label") or ""),
                (text_left, row_box[1] + 5, text_left + int((row_box[2] - text_left) * 0.34), row_box[3] - 5),
                fonts,
                text_color,
                17,
                12,
                bold=True,
                max_lines=1,
                valign="middle",
            )
            body = str(item.get("body", "")).strip()
            if body:
                _diagram_fit_text(
                    draw,
                    body,
                    (text_left + int((row_box[2] - text_left) * 0.36), row_box[1] + 5, row_box[2] - 10, row_box[3] - 5),
                    fonts,
                    mix_color(white, blue_pale, 0.45) if is_focus else muted,
                    14,
                    10,
                    max_lines=1,
                    valign="middle",
                )
            if index < len(items) - 1:
                draw_arrow(
                    draw,
                    ((row_box[0] + row_box[2]) // 2, row_box[3] + 1),
                    ((row_box[0] + row_box[2]) // 2, row_box[3] + gap - 1),
                    blue_mid,
                    width=2,
                    head=5,
                )
    elif kind == "ruler":
        items = items[:8]
        values: list[float] = []
        for item in items:
            try:
                values.append(float(item.get("value", 0)))
            except (TypeError, ValueError):
                values.append(0.0)
        if values:
            minimum, maximum = min(values), max(values)
            span = (maximum - minimum) or 1.0
            axis_y = inner_top + int((inner_bottom - inner_top) * 0.55)
            draw.line((inner_left, axis_y, inner_right, axis_y), fill=blue_dark, width=3)
            draw_arrow(draw, (inner_left, axis_y), (inner_right, axis_y), blue_dark, width=3, head=8)
            positions: dict[float, int] = {}
            for index, (item, value) in enumerate(zip(items, values)):
                x = inner_left + int((inner_right - inner_left - 12) * (value - minimum) / span)
                positions[value] = x
                is_focus = index == highlight
                draw.line((x, axis_y - 16, x, axis_y + 5), fill=red if is_focus else blue_mid, width=3)
                draw.ellipse((x - 5, axis_y - 5, x + 5, axis_y + 5), fill=red if is_focus else blue, outline=white, width=2)
                _diagram_fit_text(
                    draw,
                    str(item.get("label") or item.get("title") or value),
                    (x - 36, axis_y + 12, x + 36, axis_y + 40),
                    fonts,
                    red if is_focus else blue_dark,
                    18,
                    11,
                    bold=True,
                    align="center",
                    max_lines=1,
                )
                body = str(item.get("body", "")).strip()
                if body:
                    _diagram_fit_text(
                        draw,
                        body,
                        (x - 48, axis_y + 42, x + 48, inner_bottom),
                        fonts,
                        muted,
                        12,
                        9,
                        align="center",
                        max_lines=2,
                    )
            pairs = [pair for pair in (diagram.get("pairs") or []) if isinstance(pair, dict)]
            for pair_index, pair in enumerate(pairs[:4]):
                try:
                    start_value = float(pair.get("from", minimum))
                    end_value = float(pair.get("to", maximum))
                except (TypeError, ValueError):
                    continue
                if start_value not in positions or end_value not in positions:
                    continue
                x1, x2 = positions[start_value], positions[end_value]
                if x1 > x2:
                    x1, x2 = x2, x1
                bracket_y = inner_top + 18 + pair_index * 18
                draw.line((x1, bracket_y, x2, bracket_y), fill=blue_mid, width=2)
                draw.line((x1, bracket_y - 5, x1, bracket_y + 5), fill=blue_mid, width=2)
                draw.line((x2, bracket_y - 5, x2, bracket_y + 5), fill=blue_mid, width=2)
                label_text = str(pair.get("label", "")).strip()
                if label_text:
                    pair_font = fonts.get(TextStyle(bold=True), 13)
                    label_width = text_width(draw, label_text, pair_font)
                    draw.rectangle(
                        ((x1 + x2 - label_width) // 2 - 6, bracket_y - 10, (x1 + x2 + label_width) // 2 + 6, bracket_y + 3),
                        fill=mix_color(colors["blue_wash"], colors["white"], 0.1),
                    )
                    draw.text(
                        ((x1 + x2 - label_width) // 2, bracket_y - 11),
                        label_text,
                        font=pair_font,
                        fill=blue_mid,
                    )
    elif kind == "matrix":
        items = items[:4]
        gap = 8
        grid_left = inner_left + 42
        grid_top = inner_top + 6
        grid_right = inner_right
        grid_bottom = inner_bottom - 34
        cell_width = (grid_right - grid_left - gap) // 2
        cell_height = (grid_bottom - grid_top - gap) // 2
        for index, item in enumerate(items):
            row = index // 2
            column = index % 2
            x = grid_left + column * (cell_width + gap)
            y = grid_top + row * (cell_height + gap)
            is_focus = index == highlight
            cell_box = (x, y, x + cell_width, y + cell_height)
            draw.rounded_rectangle(
                cell_box,
                radius=4,
                fill=blue if is_focus else (white if row == column else blue_wash),
                outline=red if is_focus else colors["line"],
                width=2 if is_focus else 1,
            )
            _diagram_index(draw, (x + 9, y + 8), index + 1, colors, fonts, dark=is_focus)
            _diagram_fit_text(
                draw,
                str(item.get("title") or item.get("label") or ""),
                (x + 34, y + 8, x + cell_width - 8, y + 38),
                fonts,
                white if is_focus else blue_dark,
                16,
                11,
                bold=True,
                max_lines=1,
            )
            body = str(item.get("body", "")).strip()
            if body:
                _diagram_fit_text(
                    draw,
                    body,
                    (x + 10, y + 40, x + cell_width - 8, y + cell_height - 7),
                    fonts,
                    mix_color(white, blue_pale, 0.5) if is_focus else ink,
                    13,
                    9,
                    max_lines=2,
                    line_ratio=1.08,
                )
        draw.line((grid_left, grid_top, grid_left, grid_bottom), fill=blue_dark, width=2)
        draw.line((grid_left, grid_bottom, grid_right, grid_bottom), fill=blue_dark, width=2)
        x_axis = diagram.get("x_axis") if isinstance(diagram.get("x_axis"), dict) else {}
        y_axis = diagram.get("y_axis") if isinstance(diagram.get("y_axis"), dict) else {}
        if x_axis.get("low"):
            _diagram_fit_text(draw, str(x_axis["low"]), (grid_left, grid_bottom + 6, grid_left + cell_width, inner_bottom), fonts, muted, 12, 9, align="center", max_lines=1)
        if x_axis.get("high"):
            _diagram_fit_text(draw, str(x_axis["high"]), (grid_right - cell_width, grid_bottom + 6, grid_right, inner_bottom), fonts, muted, 12, 9, align="center", max_lines=1)
        if y_axis.get("high"):
            _diagram_fit_text(draw, str(y_axis["high"]), (inner_left, grid_top, grid_left - 6, grid_top + cell_height), fonts, muted, 12, 9, align="center", max_lines=2, valign="middle")
        if y_axis.get("low"):
            _diagram_fit_text(draw, str(y_axis["low"]), (inner_left, grid_bottom - cell_height, grid_left - 6, grid_bottom), fonts, muted, 12, 9, align="center", max_lines=2, valign="middle")
    elif kind == "cycle":
        items = items[:5]
        count = max(3, len(items))
        center_x = (inner_left + inner_right) // 2
        center_y = (inner_top + inner_bottom) // 2
        radius_x = max(70, (inner_right - inner_left) // 3)
        radius_y = max(52, (inner_bottom - inner_top) // 3)
        points: list[tuple[int, int]] = []
        for index in range(count):
            angle = -math.pi / 2 + index * 2 * math.pi / count
            points.append((int(center_x + math.cos(angle) * radius_x), int(center_y + math.sin(angle) * radius_y)))
        for index, point in enumerate(points):
            next_point = points[(index + 1) % len(points)]
            draw_arrow(draw, (point[0] + 20, point[1]), (next_point[0] - 20, next_point[1]), blue_mid, width=2, head=7)
        node_width = min(150, max(90, (inner_right - inner_left) // 4))
        for index, (item, point) in enumerate(zip(items, points)):
            is_focus = index == highlight
            node_box = (point[0] - node_width // 2, point[1] - 28, point[0] + node_width // 2, point[1] + 28)
            draw.rounded_rectangle(
                node_box,
                radius=5,
                fill=blue if is_focus else white,
                outline=red if is_focus else colors["line"],
                width=2 if is_focus else 1,
            )
            _diagram_fit_text(
                draw,
                str(item.get("title") or item.get("label") or ""),
                (node_box[0] + 8, node_box[1] + 5, node_box[2] - 8, node_box[3] - 5),
                fonts,
                white if is_focus else blue_dark,
                16,
                11,
                bold=True,
                align="center",
                max_lines=2,
                valign="middle",
            )
    if note:
        _diagram_fit_text(
            draw,
            note,
            (inner_left, inner_bottom + 6, inner_right, inner_bottom + 28),
            fonts,
            muted,
            13,
            9,
            max_lines=1,
        )


def draw_figure_card(
    image: Image.Image,
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    image_path: Path | None,
    label: str,
    caption: str,
    colors: dict[str, tuple[int, int, int]],
    fonts: FontBook,
    chart: dict | None = None,
    badge: str = "",
    badge_color: tuple[int, int, int] | None = None,
    radius: int = 12,
) -> None:
    left, top, right, bottom = box
    draw_shadowed_panel(image, draw, box, colors["white"], colors["line"], radius=radius, width=1)
    label_font = fonts.get(TextStyle(bold=True), 20)
    if label:
        label_width = text_width(draw, label, label_font) + 34
        label_box = (left + 18, top - 15, left + 18 + label_width, top + 19)
        draw_shadowed_panel(image, draw, label_box, colors["blue"], None, radius=5, width=0, shadow_blur=8, shadow_alpha=30)
        draw.text((label_box[0] + 17, label_box[1] + 7), label, font=label_font, fill=colors["white"])
    if badge:
        badge_font = fonts.get(TextStyle(bold=True), 16)
        badge_width = text_width(draw, badge, badge_font) + 22
        badge_box = (right - 18 - badge_width, top + 13, right - 18, top + 39)
        draw_tag(draw, badge_box, badge, badge_font, badge_color or colors["blue_pale"], colors["blue_dark"], radius=4)
    caption_height = 34 if caption else 0
    image_box = (left + 14, top + 26, right - 14, bottom - 14 - caption_height)
    if chart:
        draw_chart(draw, image_box, chart, colors, fonts)
    elif image_path is not None and image_path.exists():
        figure = Image.open(image_path).convert("RGB")
        # White photo mount, then contain the source without stretching.
        draw.rectangle(image_box, fill=colors["white"])
        max_w, max_h = image_box[2] - image_box[0] - 16, image_box[3] - image_box[1] - 16
        scale = min(max_w / figure.width, max_h / figure.height)
        size = (max(1, round(figure.width * scale)), max(1, round(figure.height * scale)))
        resized = figure.resize(size, Image.Resampling.LANCZOS)
        px = image_box[0] + (image_box[2] - image_box[0] - size[0]) // 2
        py = image_box[1] + (image_box[3] - image_box[1] - size[1]) // 2
        image.paste(resized, (px, py))
        draw.rectangle((px, py, px + size[0], py + size[1]), outline=mix_color(colors["line"], colors["paper"], 0.25), width=1)
    else:
        draw.rectangle(image_box, fill=colors["blue_wash"])
        missing_font = fonts.get(TextStyle(), 20)
        text = "图片缺失"
        width = text_width(draw, text, missing_font)
        draw.text((image_box[0] + (image_box[2] - image_box[0] - width) // 2, image_box[1] + 30), text, font=missing_font, fill=colors["muted"])
    if caption:
        caption_box = (left + 2, bottom - 36, right - 2, bottom - 2)
        draw.rectangle(caption_box, fill=colors["blue_wash"])
        caption_font = fonts.get(TextStyle(), 17)
        draw.text((left + 16, bottom - 29), caption, font=caption_font, fill=colors["muted"])
    draw_corner_marks(draw, (left + 8, top + 8, right - 8, bottom - 8), mix_color(colors["blue"], colors["white"], 0.55), length=16, width=2)


def make_background(
    width: int,
    height: int,
    colors: dict[str, tuple[int, int, int]],
    cover: bool = False,
    seed: int = 0,
    settings: dict | None = None,
) -> Image.Image:
    settings = settings or {}
    motif = str(settings.get("motif", "arcs"))
    motif_opacity = max(0.0, min(1.0, float(settings.get("motif_opacity", 0.22))))
    noise_sigma = max(0.0, float(settings.get("noise", 1.1)))
    paper = np.array(colors["paper"], dtype=np.float32)
    paper_alt = np.array(colors["paper_alt"], dtype=np.float32)
    gradient = np.zeros((height, width, 3), dtype=np.float32)
    for row in range(height):
        amount = row / max(1, height - 1)
        gradient[row, :, :] = paper * (1 - amount * 0.52) + paper_alt * (amount * 0.52)
    # A soft highlight keeps the paper from looking like a flat gray wash.
    yy, xx = np.mgrid[0:height, 0:width]
    glow = np.exp(-(((xx - width * 0.5) / (width * 0.52)) ** 2 + ((yy - height * 0.1) / (height * 0.5)) ** 2))
    gradient = gradient + glow[:, :, None] * np.array([4, 7, 11], dtype=np.float32)
    rng = np.random.default_rng(4731 + seed)
    noise = rng.normal(0, noise_sigma, size=(height, width, 1))
    gradient = np.clip(gradient + noise, 0, 255).astype(np.uint8)
    image = Image.fromarray(gradient, mode="RGB").convert("RGBA")
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    line_color = mix_color(colors["line"], colors["paper"], 0.28) + (150,)
    accent_color = mix_color(colors["blue_mid"], colors["paper"], 0.68) + (88,)
    if settings.get("arc", True):
        draw.arc((-540, 900, 2600, 1790), 200, 340, fill=line_color, width=5)
        draw.arc((-280, 1010, 2330, 1600), 198, 342, fill=accent_color, width=3)
    if motif == "campus":
        # Optional institutional line-art; only enabled by a theme that wants it.
        ghost = mix_color(colors["blue"], colors["paper"], 0.88) + (max(0, round(18 * motif_opacity / 0.22)),)
        base_y = height - 34
        blocks = [
            (int(width * 0.705), int(height * 0.855), int(width * 0.795), base_y),
            (int(width * 0.795), int(height * 0.805), int(width * 0.875), base_y),
            (int(width * 0.875), int(height * 0.875), int(width * 0.965), base_y),
        ]
        for left, top, right, bottom in blocks:
            draw.rectangle((left, top, right, bottom), outline=ghost, width=2)
            for window in range(3):
                y = top + (bottom - top) * (0.28 + window * 0.2)
                draw.line((left + 8, y, right - 8, y), fill=ghost, width=1)
        tower_left = int(width * 0.815)
        tower_right = int(width * 0.865)
        tower_top = int(height * 0.735)
        draw.rectangle((tower_left, tower_top, tower_right, base_y), outline=ghost, width=2)
        draw.polygon(
            [(tower_left - 10, tower_top), ((tower_left + tower_right) // 2, tower_top - 28), (tower_right + 10, tower_top)],
            outline=ghost,
            width=2,
        )
        draw.ellipse(
            (
                (tower_left + tower_right) // 2 - 14,
                tower_top + 18,
                (tower_left + tower_right) // 2 + 14,
                tower_top + 46,
            ),
            outline=ghost,
            width=2,
        )
    elif motif == "contour":
        contour = mix_color(colors["blue_mid"], colors["paper"], 0.76) + (max(0, round(120 * motif_opacity / 0.22)),)
        for band in range(7):
            points = []
            phase = band * 0.52
            base = height * (0.74 + band * 0.025)
            for x in range(-40, width + 90, 18):
                t = x / max(1, width)
                wave = math.sin(t * math.pi * 1.55 + phase) * 34
                wave += math.sin(t * math.pi * 3.2 - phase * 0.7) * 12
                points.append((x, int(base + wave + band * 12)))
            draw.line(points, fill=contour, width=2)
        for ring in range(4):
            inset = ring * 38
            draw.arc(
                (
                    int(width * 0.64) + inset,
                    int(height * 0.32) + inset,
                    int(width * 0.98) - inset,
                    int(height * 0.83) - inset,
                ),
                212,
                338,
                fill=accent_color,
                width=2,
            )
    if cover:
        draw.arc((-860, 760, 2900, 1840), 200, 340, fill=mix_color(colors["line"], colors["paper"], 0.2) + (200,), width=9)
        draw.arc((-440, 950, 2480, 1680), 198, 342, fill=accent_color, width=4)
    return Image.alpha_composite(image, overlay).convert("RGB")


def make_section_number(
    draw: ImageDraw.ImageDraw,
    number: str,
    xy: tuple[int, int],
    fonts: FontBook,
    color: tuple[int, int, int],
    size: int = 180,
) -> tuple[int, int]:
    font = fonts.get(TextStyle(bold=True), size)
    draw.text(xy, number, font=font, fill=color)
    bbox = draw.textbbox(xy, number, font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]
