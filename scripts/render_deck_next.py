from __future__ import annotations

import argparse
import json
import math
import shutil
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches

from common import (
    apply_navigation_mode,
    apply_type_profile,
    clamp_defense_size_px,
    content_section_openers,
    DEFENSE_BODY_PX,
    DEFENSE_MIN_ANY_PX,
    figure_width_ratio,
    draw_lines,
    fit_text_block,
    line_height,
    load_json,
    load_theme,
    mix_color,
    navigation_mode_for,
    resolve_font,
    resolve_path,
    source_line_for_slide,
    text_width,
    write_json,
)
from design import (
    FontBook,
    TextStyle,
    draw_arrow,
    draw_chart,
    draw_chevrons,
    draw_diagram,
    draw_figure_card,
    draw_gradient_panel,
    draw_icon,
    draw_icon_badge,
    draw_mini_visual,
    draw_rich_block,
    draw_shadowed_panel,
    draw_tag,
    make_background,
    render_math_image,
)

HEADER_EYEBROW_GAP = 56


def copy_if_distinct(source: Path, target: Path) -> bool:
    """Copy a deck only when source and target are different files."""
    if source.resolve() == target.resolve():
        return False
    shutil.copy2(source, target)
    return True


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render an academic defense deck with the editorial visual system.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--theme", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--columns", type=int, default=3)
    parser.add_argument("--profile", choices=["document", "projection", "lecture", "defense"], default=None)
    parser.add_argument("--sample", type=int, default=None, help="Render only the first N slides as a visual sample.")
    parser.add_argument("--pages", default=None, help="Comma-separated 1-based slide numbers, e.g. 1,4,7.")
    return parser.parse_args()


class DeckRenderer:
    def __init__(
        self,
        deck_path: Path,
        theme_path: Path,
        output_dir: Path,
        profile: str | None = None,
        pages: list[int] | None = None,
        sample: int | None = None,
    ) -> None:
        self.deck_path = deck_path.resolve()
        self.deck = load_json(self.deck_path)
        self.theme = load_theme(theme_path)
        self.profile = profile or str(self.deck.get("meta", {}).get("type_profile", "document"))
        apply_type_profile(self.theme, self.profile)
        self.navigation_mode = apply_navigation_mode(
            self.theme,
            navigation_mode_for(self.deck, self.theme),
        )
        self.output_dir = output_dir.resolve()
        self.render_dir = self.output_dir / "rendered"
        self.render_dir.mkdir(parents=True, exist_ok=True)
        if pages:
            self.page_selection = [page for page in pages if page >= 1]
        elif sample:
            self.page_selection = list(range(1, max(0, sample) + 1))
        else:
            self.page_selection = None
        self.canvas = self.theme["canvas"]
        self.width = int(self.canvas["width"])
        self.height = int(self.canvas["height"])
        self.colors = self.theme["_colors"]
        self.layout = self.theme["layout"]
        self.type_scale = self.theme.get("type_scale", {})
        body_bold_min_px = float(self.theme.get("text", {}).get("body_bold_min_px", 40))
        if self.profile == "defense":
            body_bold_min_px = DEFENSE_BODY_PX
        self.fonts = FontBook(
            self.theme["fonts"]["title"],
            self.theme["fonts"]["body"],
            body_bold=bool(self.theme.get("text", {}).get("body_bold", False)),
            body_bold_min_px=body_bold_min_px,
            body_bold_max_px=float(self.theme.get("text", {}).get("body_bold_max_px", 56)),
            min_size_px=0,
        )
        self._header_title_left: int | None = None
        self.serif_fonts = self.theme["fonts"]["serif"]
        self.display_fonts = self.theme["fonts"].get("display", self.serif_fonts)
        self.formula_fonts = self.theme["fonts"].get("formula", self.serif_fonts)
        self.latin_fonts = self.theme["fonts"].get("latin", self.formula_fonts)
        self.panel_radius = int(self.layout.get("panel_radius", 8))
        self.panel_shadow_alpha = int(self.layout.get("panel_shadow_alpha", 20))
        self.warnings: list[dict] = []
        self.section_map = {section["id"]: index for index, section in enumerate(self.deck.get("sections", []))}
        # Sidebar (mini directory) policy: "always" keeps the historical behaviour,
        # "section" shows it only where a new section starts, "off" removes it and
        # lets the content box widen to `layout.content_no_sidebar`.
        self.sidebar_mode = str(self.layout.get("sidebar_mode", "always")).strip().lower() or "always"
        self.section_openers = content_section_openers(self.deck)
        import design as _design_module

        self.block_icons = str(self.layout.get("block_icons", "on")).strip().lower() != "off"
        _design_module.DRAW_ICONS = self.block_icons
        _design_module.DRAW_CORNER_MARKS = bool((self.theme.get("panel", {}) or {}).get("corner_marks", True))
        _panel_cfg = self.theme.get("panel", {}) or {}
        self.card_index = bool(_panel_cfg.get("card_index", True))
        self.card_highlight = bool(_panel_cfg.get("card_highlight", True))
        if self.sidebar_mode == "off":
            wide = self.layout.get("content_no_sidebar")
            if isinstance(wide, (list, tuple)) and len(wide) == 4:
                self.layout["content"] = [int(value) for value in wide]
        self.lockup = self.load_lockup()
        self.emblem = self.load_emblem()

    def color(self, name: str) -> tuple[int, int, int]:
        return self.colors[name]

    @staticmethod
    def has_cjk(text: str) -> bool:
        return any("\u4e00" <= char <= "\u9fff" for char in text)

    def type_size(self, name: str, fallback: int) -> int:
        value = int(self.type_scale.get(name, fallback))
        return int(round(clamp_defense_size_px(value, self.profile)))

    def fit_text(
        self,
        draw: ImageDraw.ImageDraw,
        text: str,
        font_candidates: Iterable[str],
        max_width: int,
        max_height: int,
        max_size: int,
        min_size: int,
        *args,
        body_floor: bool = False,
        **kwargs,
    ):
        max_size = int(round(clamp_defense_size_px(max_size, self.profile, body_floor=body_floor)))
        min_size = int(round(clamp_defense_size_px(min_size, self.profile)))
        if self.profile == "defense" and not body_floor and max_size < DEFENSE_BODY_PX:
            min_size = min(min_size, 12)
        if min_size > max_size:
            min_size = max_size
        return fit_text_block(
            draw,
            text,
            font_candidates,
            max_width,
            max_height,
            max_size,
            min_size,
            *args,
            **kwargs,
        )

    def load_lockup(self) -> Image.Image | None:
        lockup_path = self.theme.get("brand", {}).get("lockup_path")
        if not lockup_path:
            return None
        path = Path(lockup_path)
        if not path.exists():
            self.warn("brand", f"Brand lockup missing: {path}")
            return None
        return Image.open(path).convert("RGBA")

    def load_emblem(self) -> Image.Image | None:
        emblem_path = self.theme.get("brand", {}).get("emblem")
        if not emblem_path:
            return None
        path = resolve_path(emblem_path, Path(self.theme["_dir"]))
        if path is None or not path.exists():
            self.warn("brand", f"Brand emblem missing: {path}")
            return None
        return Image.open(path).convert("RGBA")

    def warn(self, page: str, message: str) -> None:
        self.warnings.append({"page": page, "message": message})

    def background(self, page_number: int, cover: bool = False) -> Image.Image:
        return make_background(
            self.width,
            self.height,
            self.colors,
            cover=cover,
            seed=page_number,
            settings=self.theme.get("background", {}),
        )

    def draw_brand(
        self,
        image: Image.Image,
        box: tuple[int, int, int, int],
        fallback_size: int = 30,
        align: str = "center",
    ) -> None:
        if self.lockup is not None:
            fitted, position = self.fit_transparent(self.lockup, box)
            image.paste(fitted, position, fitted)
            return
        draw = ImageDraw.Draw(image)
        text = self.theme.get("brand", {}).get("fallback_text") or self.deck.get("meta", {}).get("institution", "")
        if not text:
            return
        left, top, right, bottom = box
        inside_width = right - left
        font, lines, _, fits = self.fit_text(
            draw,
            text,
            self.theme["fonts"].get("brand") or self.theme["fonts"]["title"],
            inside_width,
            bottom - top,
            fallback_size,
            max(15, fallback_size - 10),
            max_lines=2,
        )
        if not fits:
            self.warn("brand", "Fallback brand text did not fit its box.")
        block_height = line_height(font) * max(1, len(lines))
        start_y = top + (bottom - top - block_height) // 2
        draw_lines(
            draw,
            lines,
            (left, start_y),
            font,
            self.color("blue"),
            align=align,
            max_width=inside_width,
        )

    @staticmethod
    def fit_transparent(image: Image.Image, box: tuple[int, int, int, int]) -> tuple[Image.Image, tuple[int, int]]:
        left, top, right, bottom = box
        width, height = right - left, bottom - top
        scale = min(width / image.width, height / image.height)
        size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
        resized = image.resize(size, Image.Resampling.LANCZOS)
        return resized, (left + (width - size[0]) // 2, top + (height - size[1]) // 2)

    def draw_header(self, image: Image.Image, slide: dict, with_title: bool = True) -> None:
        draw = ImageDraw.Draw(image)
        blue = self.color("blue")
        blue_mid = self.color("blue_mid")
        slash_left = int(self.layout["slash_x"])
        slash_top = int(self.layout["slash_y"])
        slash_width = int(self.layout["slash_width"])
        slash_height = int(self.layout["slash_height"])
        first = [
            (slash_left, slash_top + slash_height),
            (slash_left + slash_width * 0.55, slash_top),
            (slash_left + slash_width * 0.98, slash_top),
            (slash_left + slash_width * 0.43, slash_top + slash_height),
        ]
        second = [(point[0] + slash_width + 4, point[1]) for point in first]
        draw.polygon(first, fill=blue)
        draw.polygon(second, fill=blue_mid)

        eyebrow = str(slide.get("eyebrow", "")).strip()
        chapter_x = int(self.layout["chapter_x"])
        title_left = self.header_title_left()
        title = str(slide.get("title", "")).strip()
        title_width = int(self.theme["brand"]["header_box"][0] - title_left - 54)
        header_height = max(56, int(self.type_size("header_title", 48) * 1.25))
        font, lines, _, fits = self.fit_text(
            draw,
            title,
            self.theme["fonts"]["title"],
            title_width,
            header_height,
            self.type_size("header_title", 48),
            max(34, int(self.type_size("header_title", 48) * 0.72)),
            max_lines=1,
        )
        if not fits:
            self.warn(str(slide.get("id", "")), f"Header title shrank or overflowed: {title}")
        if eyebrow:
            # Keep a real text gutter inside the label column. The previous
            # width ended exactly at the measured glyph edge, so PowerPoint
            # could repack "1.2" as two lines even though the PNG looked fine.
            eyebrow_width_available = max(24, title_left - chapter_x - 24)
            configured_title_size = self.type_size("header_title", 48)
            configured_eyebrow_size = self.type_size("eyebrow", configured_title_size)
            if configured_eyebrow_size != configured_title_size:
                self.warn(
                    str(slide.get("id", "")),
                    "Header eyebrow type size differs from header_title; set both to the same value in type_scale.",
                )
            eyebrow_size = font.size
            eyebrow_font = self.fonts.get(TextStyle(bold=True), eyebrow_size)
            eyebrow_width = text_width(draw, eyebrow, eyebrow_font)
            if eyebrow_width > eyebrow_width_available:
                self.warn(
                    str(slide.get("id", "")),
                    "Header eyebrow exceeds its column; shorten the label instead of shrinking only the number.",
                )
            title_y = int(self.layout["title_y"])
            draw.text((chapter_x, title_y), eyebrow, font=eyebrow_font, fill=blue)
            divider_x = chapter_x + eyebrow_width + HEADER_EYEBROW_GAP // 2
            if divider_x < title_left - 8:
                draw.line((divider_x, title_y + 4, divider_x, title_y + eyebrow_font.size * 0.92), fill=self.color("line"), width=2)
        draw.text((title_left, int(self.layout["title_y"])), lines[0] if lines else title, font=font, fill=self.color("blue_dark"))

        divider_y = int(self.layout["divider_y"])
        draw.rectangle((24, divider_y, self.width - 24, divider_y + 4), fill=blue)
        draw.rectangle((24, divider_y + 6, self.width - 24, divider_y + 7), fill=mix_color(blue, self.color("paper"), 0.55))
        self.draw_brand(image, tuple(self.theme["brand"]["header_box"]))
        draw.rectangle((24, int(self.layout["footer_y"]), self.width - 24, int(self.layout["footer_y"]) + 3), fill=blue)

    def header_title_left(self) -> int:
        """Return one title x for every page, wide enough for the widest eyebrow.

        The chapter label used to live in the fixed gap between ``chapter_x``
        and ``title_x``. Once the label and the title share one type size, a
        label such as ``1.1`` at 72 px is wider than that gap and the two runs
        collide. Measure the whole deck once and reserve a real column; the
        title then stays on one vertical alignment across pages.
        """
        if self._header_title_left is not None:
            return self._header_title_left
        chapter_x = int(self.layout["chapter_x"])
        title_left = int(self.layout["title_x"])
        size = self.type_size("header_title", 48)
        font = self.fonts.get(TextStyle(bold=True), size)
        draw = ImageDraw.Draw(Image.new("RGB", (4, 4)))
        eyebrow_width = 0
        for item in self.deck.get("slides", []):
            label = str(item.get("eyebrow", "")).strip()
            if label:
                eyebrow_width = max(eyebrow_width, text_width(draw, label, font))
        if eyebrow_width:
            title_left = max(title_left, chapter_x + eyebrow_width + HEADER_EYEBROW_GAP)
        self._header_title_left = title_left
        return title_left

    def draw_sidebar(self, image: Image.Image, active_section: str | None, slide_id: str | None = None) -> None:
        sections = self.deck.get("sections", [])
        if not sections:
            return
        if self.sidebar_mode == "off":
            return
        if self.sidebar_mode == "section":
            if str(slide_id or "") not in self.section_openers:
                return
            # The directory needs 300 px of left margin. If the theme widened the
            # text box into that margin, drawing it would overlap the body text.
            if float(self.layout["content"][0]) < 300:
                return
        draw = ImageDraw.Draw(image)
        left, top, right, bottom = map(int, self.layout["sidebar"])
        gap = 10
        row_height = min(96, (bottom - top - gap * (len(sections) - 1)) // len(sections))
        for index, section in enumerate(sections):
            row_top = top + index * (row_height + gap)
            active = section["id"] == active_section
            if active:
                points = [
                    (left, row_top),
                    (right - 18, row_top),
                    (right, row_top + row_height // 2),
                    (right - 18, row_top + row_height),
                    (left, row_top + row_height),
                ]
                # A narrow cast shadow gives the active tab the same physical feel as the reference deck.
                shadow = Image.new("RGBA", (right - left + 40, row_height + 24), (0, 0, 0, 0))
                shadow_draw = ImageDraw.Draw(shadow)
                shadow_draw.polygon([(x + 2, y + 7) for x, y in points], fill=(13, 45, 92, 30))
                from PIL import ImageFilter
                shadow = shadow.filter(ImageFilter.GaussianBlur(7))
                image.paste(shadow, (left, row_top - 6), shadow)
                draw = ImageDraw.Draw(image)
                draw.polygon(points, fill=self.color("blue"))
                draw.line((left, row_top + 1, right - 18, row_top + 1), fill=mix_color(self.color("blue"), self.color("white"), 0.3), width=2)
            else:
                draw_shadowed_panel(
                    image,
                    draw,
                    (left, row_top, right - 8, row_top + row_height),
                    self.color("white"),
                    self.color("line"),
                    radius=6,
                    width=2,
                    shadow_offset=(0, 4),
                    shadow_blur=7,
                    shadow_alpha=20,
                )
                draw.rectangle((left, row_top + 14, left + 4, row_top + row_height - 14), fill=self.color("blue_pale"))
            text_color = self.color("white") if active else self.color("blue_dark")
            font, lines, _, fits = self.fit_text(
                draw,
                section["title"],
                self.theme["fonts"]["title"],
                (right - left) - 30,
                row_height - 10,
                self.type_size("sidebar", 29),
                17,
                max_lines=2,
                line_ratio=1.15,
            )
            if not fits:
                self.warn(str(section.get("id", "")), "Sidebar label did not fit.")
            block_height = line_height(font, 1.15) * len(lines)
            text_top = row_top + (row_height - block_height) // 2
            text_left = left + 14
            text_max_width = (right - left) - 24 if active else (right - left) - 24
            draw_lines(draw, lines, (text_left, text_top), font, text_color, 1.15, "center", text_max_width)

    def draw_lead(self, image: Image.Image, text: str, top: int, page_id: str) -> int:
        if not text:
            return top
        draw = ImageDraw.Draw(image)
        left, _, right, _ = map(int, self.layout["content"])
        base_height = int(self.layout.get("lead_height", 88))
        label = str(self.layout.get("lead_label", "")).strip()
        if label:
            label_width = int(self.layout.get("lead_label_width", 168))
            label_width = min(label_width, max(120, (right - left) // 3))
            text_left = left + label_width + 24
            text_width_available = right - text_left - 24
        else:
            label_width = 0
            text_left = left + 96
            text_width_available = right - left - 118

        # Measure the real line count instead of assuming an 88 px strip. A
        # two-line lead used to be clipped by the first content card because
        # the height stayed fixed while PowerPoint correctly wrapped the text.
        lead_size = self.type_size("lead", 31)
        minimum_size = max(22, self.type_size("body", 22))
        font, lines, block_height, fits = self.fit_text(
            draw,
            text,
            self.theme["fonts"]["title"],
            text_width_available,
            max(base_height, lead_size * 3),
            lead_size,
            minimum_size,
            max_lines=2,
            line_ratio=1.2,
        )
        if not fits:
            self.warn(page_id, "Lead text overflowed even after reserving two lines.")
        height = max(base_height, block_height + 24)
        if self.layout.get("lead_panel", True):
            draw_gradient_panel(
                image,
                draw,
                (left, top, right, top + height),
                mix_color(self.color("blue_wash"), self.color("white"), 0.12),
                self.color("white"),
                self.color("line"),
                radius=self.panel_radius,
                width=1,
            )
        if label:
            points = [
                (left, top + 1),
                (left + label_width - 18, top + 1),
                (left + label_width, top + height // 2),
                (left + label_width - 18, top + height - 1),
                (left, top + height - 1),
            ]
            draw.polygon(points, fill=self.color("blue"))
            draw.line(
                (left + 1, top + 1, left + 1, top + height - 1),
                fill=mix_color(self.color("blue_mid"), self.color("white"), 0.28),
                width=2,
            )
            label_font = self.fonts.get(TextStyle(bold=True), self.type_size("lead_label", 22))
            label_width_actual = text_width(draw, label, label_font)
            draw.text(
                (left + (label_width - 18 - label_width_actual) // 2, top + (height - line_height(label_font)) // 2 - 2),
                label,
                font=label_font,
                fill=self.color("white"),
            )
        else:
            draw.rectangle((left, top + 1, left + 9, top + height - 1), fill=self.color("blue"))
            draw_chevrons(draw, left + 28, top + 26, 34, self.color("blue_mid"), count=2, width=4, gap=13)
        fits, _ = draw_rich_block(
            draw,
            text,
            (text_left, top + 12),
            text_width_available,
            height - 24,
            self.fonts,
            self.colors,
            font.size,
            font.size,
            max_lines=2,
            line_ratio=1.2,
        )
        if not fits:
            self.warn(page_id, "Lead text overflowed.")
        return top + height + (20 if self.layout.get("lead_panel", True) else 12)

    def draw_takeaway(self, image: Image.Image, slide: dict, page_id: str) -> None:
        text = str(slide.get("takeaway", "")).strip()
        if not text:
            return
        draw = ImageDraw.Draw(image)
        left = int(self.layout["content"][0])
        right = int(self.layout["content"][2])
        top = int(self.layout["takeaway_top"])
        height = int(self.layout["takeaway_height"])
        draw_gradient_panel(
            image,
            draw,
            (left, top, right, top + height),
            self.color("blue"),
            self.color("blue_dark"),
            None,
            radius=self.panel_radius,
            width=0,
        )
        draw.line((left + 2, top + 1, right - 2, top + 1), fill=mix_color(self.color("blue_mid"), self.color("white"), 0.3), width=2)
        draw_chevrons(draw, left + 26, top + 21, 28, self.color("white"), count=2, width=4, gap=11)
        light_colors = dict(self.colors)
        light_colors.update(
            {
                "ink": self.color("white"),
                "blue_dark": self.color("white"),
                "muted": mix_color(self.color("white"), self.color("blue_pale"), 0.36),
                "red": mix_color(self.color("white"), self.color("red"), 0.58),
            }
        )
        fits, _ = draw_rich_block(
            draw,
            text,
            (left + 88, top + 12),
            right - left - 112,
            height - 18,
            self.fonts,
            light_colors,
            self.type_size("takeaway", 29),
            19,
            max_lines=2,
            line_ratio=1.16,
        )
        if not fits:
            self.warn(page_id, "Takeaway overflowed.")

    def content_bounds(self, slide: dict) -> tuple[int, int, int, int]:
        left, top, right, bottom = map(int, self.layout["content"])
        if slide.get("takeaway"):
            bottom = int(self.layout["takeaway_top"]) - 20
        return left, top, right, bottom

    def draw_cover(self, slide: dict) -> Image.Image:
        image = self.background(1, cover=True)
        cover = self.theme["cover"]
        if str(cover.get("style", "centered")) == "editorial":
            return self.draw_cover_editorial(slide, image)
        return self.draw_cover_centered(slide, image)

    def draw_cover_centered(self, slide: dict, image: Image.Image) -> Image.Image:
        draw = ImageDraw.Draw(image)
        cover = self.theme["cover"]
        meta = self.deck.get("meta", {})
        self.draw_brand(image, tuple(self.theme["brand"]["cover_box"]), fallback_size=46)

        title_width = 1810
        title_left = (self.width - title_width) // 2
        title_text = str(slide.get("title", meta.get("title", "")))
        title_lines = slide.get("title_lines") or meta.get("title_lines")
        if isinstance(title_lines, list) and title_lines:
            title_text = "\n".join(str(line) for line in title_lines)
        font, lines, _, fits = self.fit_text(
            draw,
            title_text,
            self.display_fonts,
            title_width,
            128,
            self.type_size("cover_title", 80),
            48,
            max_lines=2,
            line_ratio=1.14,
        )
        if not fits:
            self.warn(str(slide.get("id", "")), "Cover title overflowed.")
        draw_lines(draw, lines, (title_left, int(cover["title_y"])), font, self.color("blue"), 1.14, "center", title_width)

        subtitle = slide.get("subtitle") or meta.get("subtitle")
        if subtitle:
            subtitle_font, subtitle_lines, _, subtitle_fits = self.fit_text(
                draw,
                subtitle,
                self.theme["fonts"]["body"],
                title_width,
                54,
                self.type_size("cover_subtitle", 35),
                22,
                max_lines=1,
            )
            if not subtitle_fits:
                self.warn(str(slide.get("id", "")), "Cover subtitle overflowed.")
            draw_lines(
                draw,
                subtitle_lines,
                (title_left, int(cover["subtitle_y"])),
                subtitle_font,
                self.color("muted"),
                align="center",
                max_width=title_width,
            )

        rule_y = int(cover["rule_y"])
        draw.rectangle((300, rule_y, self.width - 300, rule_y + 4), fill=self.color("blue"))
        draw.rectangle((300, rule_y + 10, self.width - 300, rule_y + 11), fill=mix_color(self.color("blue_mid"), self.color("paper"), 0.2))

        presenter = meta.get("presenter", "")
        advisor = meta.get("advisor", "")
        date = meta.get("date", "")
        info_font = self.fonts.get(TextStyle(), self.type_size("cover_info", 31))
        first_line = f"汇报人：{presenter}"
        second_line = f"指导教师：{advisor}"
        date_line = f"汇报时间：{date}"
        group_width = text_width(draw, first_line, info_font) + text_width(draw, second_line, info_font) + 102
        group_left = (self.width - group_width) // 2
        info_y = int(cover["info_y"])
        icon_color = self.color("blue")
        self.draw_person_icon(draw, group_left, info_y + 4, 28, icon_color)
        draw.text((group_left + 42, info_y), first_line, font=info_font, fill=self.color("ink"))
        separator_x = group_left + 42 + text_width(draw, first_line, info_font) + 34
        draw.line((separator_x, info_y + 1, separator_x, info_y + 38), fill=self.color("line"), width=2)
        self.draw_cap_icon(draw, separator_x + 26, info_y + 3, 30, icon_color)
        draw.text((separator_x + 66, info_y), second_line, font=info_font, fill=self.color("ink"))
        date_width = text_width(draw, date_line, info_font)
        date_left = (self.width - date_width) // 2
        self.draw_calendar_icon(draw, date_left - 44, info_y + 68, 28, icon_color)
        draw.text((date_left, info_y + 62), date_line, font=info_font, fill=self.color("ink"))
        draw.rectangle((24, int(self.layout["footer_y"]), self.width - 24, int(self.layout["footer_y"]) + 3), fill=self.color("blue"))
        return image

    def draw_cover_editorial(self, slide: dict, image: Image.Image) -> Image.Image:
        draw = ImageDraw.Draw(image)
        cover = self.theme["cover"]
        meta = self.deck.get("meta", {})
        accent_x = int(cover.get("accent_x", 124))
        title_width = int(cover.get("title_width", 1092))

        draw.rectangle((0, 0, 12, self.height), fill=self.color("blue"))
        brand_box = tuple(cover.get("brand_box", self.theme["brand"]["cover_box"]))
        self.draw_brand(image, brand_box, fallback_size=46, align="left")

        panel_left = int(cover.get("panel_left", int(self.width * 0.615)))
        draw.line(
            (panel_left, 210, panel_left, 930),
            fill=mix_color(self.color("line"), self.color("paper"), 0.35),
            width=2,
        )

        if self.emblem is not None:
            watermark_box = tuple(cover.get("watermark_box", (1288, 228, 1964, 862)))
            fitted, position = self.fit_transparent(self.emblem, watermark_box)
            watermark = fitted.copy()
            opacity = float(cover.get("watermark_alpha", 0.09))
            alpha = watermark.getchannel("A").point(lambda value: int(value * opacity))
            watermark.putalpha(alpha)
            image.paste(watermark, position, watermark)
            draw = ImageDraw.Draw(image)
            ring_color = mix_color(self.color("blue_pale"), self.color("paper"), 0.48)
            for ring in range(3):
                inset = ring * 46
                draw.arc(
                    (
                        watermark_box[0] + 16 + inset,
                        watermark_box[1] + 38 + inset,
                        watermark_box[2] - 16 - inset,
                        watermark_box[3] - 38 - inset,
                    ),
                    208,
                    338,
                    fill=ring_color,
                    width=2,
                )

        kicker = str(meta.get("report_type") or meta.get("subtitle") or "学术报告").strip()
        kicker_y = int(cover.get("kicker_y", 252))
        draw.rectangle((accent_x, kicker_y + 1, accent_x + 7, kicker_y + 38), fill=self.color("blue"))
        kicker_font, kicker_lines, _, kicker_fits = self.fit_text(
            draw,
            kicker,
            self.theme["fonts"]["title"],
            title_width - 40,
            44,
            self.type_size("cover_kicker", 27),
            17,
            max_lines=1,
            line_ratio=1.12,
        )
        if not kicker_fits:
            self.warn(str(slide.get("id", "")), "Cover kicker overflowed.")
        draw.text(
            (accent_x + 24, kicker_y),
            kicker_lines[0] if kicker_lines else kicker,
            font=kicker_font,
            fill=self.color("blue"),
        )

        title = str(slide.get("title") or meta.get("title") or "").strip()
        explicit_title_lines = slide.get("title_lines") or meta.get("title_lines")
        if isinstance(explicit_title_lines, list) and explicit_title_lines:
            explicit_title_lines = [str(line) for line in explicit_title_lines if str(line).strip()]
            title = "\n".join(explicit_title_lines)
        title_y = int(cover.get("title_y", 304))
        title_fits = True
        if isinstance(explicit_title_lines, list) and explicit_title_lines:
            title_size = self.type_size("cover_title", 86)
            while title_size >= 54:
                title_font = resolve_font(self.display_fonts, title_size)
                title_lines = explicit_title_lines
                title_fits = all(
                    text_width(draw, line, title_font) <= title_width
                    for line in title_lines
                ) and line_height(title_font, 1.12) * len(title_lines) <= 270
                if title_fits:
                    break
                title_size -= 2
        else:
            title_font, title_lines, _, title_fits = self.fit_text(
                draw,
                title,
                self.display_fonts,
                title_width,
                270,
                self.type_size("cover_title", 86),
                54,
                max_lines=3,
                line_ratio=1.12,
            )
        if not title_fits:
            self.warn(str(slide.get("id", "")), "Cover title overflowed.")
        draw_lines(
            draw,
            title_lines,
            (accent_x, title_y),
            title_font,
            self.color("blue"),
            1.12,
            "left",
            title_width,
        )
        title_bottom = title_y + line_height(title_font, 1.12) * max(1, len(title_lines))

        subtitle = str(slide.get("subtitle") or meta.get("subtitle") or "").strip()
        subtitle_bottom = title_bottom
        if subtitle and subtitle != kicker:
            subtitle_y = title_bottom + int(cover.get("subtitle_gap", 22))
            subtitle_font, subtitle_lines, _, subtitle_fits = self.fit_text(
                draw,
                subtitle,
                self.theme["fonts"]["body"],
                title_width,
                96,
                self.type_size("cover_subtitle", 34),
                22,
                max_lines=2,
                line_ratio=1.25,
            )
            if not subtitle_fits:
                self.warn(str(slide.get("id", "")), "Cover subtitle overflowed.")
            draw_lines(
                draw,
                subtitle_lines,
                (accent_x, subtitle_y),
                subtitle_font,
                self.color("muted"),
                1.25,
                "left",
                title_width,
            )
            subtitle_bottom = subtitle_y + line_height(subtitle_font, 1.25) * max(1, len(subtitle_lines))

        rule_y = subtitle_bottom + int(cover.get("rule_gap", 34))
        draw.rectangle((accent_x, rule_y, accent_x + title_width, rule_y + 4), fill=self.color("blue"))
        draw.rectangle(
            (accent_x, rule_y + 10, accent_x + title_width, rule_y + 11),
            fill=mix_color(self.color("blue_mid"), self.color("paper"), 0.25),
        )

        info_y = max(int(cover.get("info_y", 708)), rule_y + int(cover.get("info_gap", 34)))
        label_font = self.fonts.get(TextStyle(), max(22, int(self.type_size("cover_info", 28) * 0.62)))
        value_size = self.type_size("cover_info", 28)
        value_font = self.fonts.get(TextStyle(), value_size)
        info_items = [
            ("汇报人", str(meta.get("presenter") or "—")),
            ("指导教师", str(meta.get("advisor") or "—")),
            ("日期", str(meta.get("date") or "—")),
        ]
        info_right = accent_x + title_width
        date_text = info_items[2][1]
        date_width = text_width(draw, date_text, value_font)
        date_x = info_right - date_width
        presenter_x = accent_x
        advisor_x = accent_x + max(280, int(title_width * 0.30))
        advisor_width = max(120, date_x - 56 - advisor_x)
        advisor_font, advisor_lines, _, advisor_fits = self.fit_text(
            draw,
            info_items[1][1],
            self.theme["fonts"]["body"],
            advisor_width,
            int(value_size * 1.35),
            value_size,
            max(24, int(value_size * 0.72)),
            max_lines=1,
            line_ratio=1.0,
        )
        if not advisor_fits:
            self.warn(str(slide.get("id", "")), "Cover advisor line overflowed.")
        info_x = [presenter_x, advisor_x, date_x]
        for index, (x, (label, value)) in enumerate(zip(info_x, info_items)):
            draw.text((x, info_y), label, font=label_font, fill=self.color("muted"))
            font = advisor_font if index == 1 else value_font
            draw.text((x, info_y + 30), value, font=font, fill=self.color("ink"))

        department = str(meta.get("department") or meta.get("institution") or "").strip()
        if department:
            dept_size = max(22, int(value_size * 0.62))
            dept_font, dept_lines, _, dept_fits = self.fit_text(
                draw,
                department,
                self.theme["fonts"]["body"],
                title_width,
                int(dept_size * 1.4),
                dept_size,
                max(18, int(dept_size * 0.8)),
                max_lines=1,
                line_ratio=1.0,
            )
            if not dept_fits:
                self.warn(str(slide.get("id", "")), "Cover department line overflowed.")
            draw.text(
                (accent_x, info_y + 88),
                dept_lines[0] if dept_lines else department,
                font=dept_font,
                fill=self.color("muted"),
            )

        cover_line = str(meta.get("cover_line") or "").strip()
        if cover_line:
            band_y = int(cover.get("bottom_band_y", 1002))
            band_height = int(cover.get("bottom_band_height", 92))
            draw.rectangle((0, band_y, self.width, band_y + band_height), fill=self.color("blue_dark"))
            draw.rectangle((0, band_y, 10, band_y + band_height), fill=self.color("red"))
            line_font, line_text, _, line_fits = self.fit_text(
                draw,
                cover_line,
                self.theme["fonts"]["body"],
                self.width - accent_x * 2,
                band_height - 22,
                26,
                18,
                max_lines=1,
                line_ratio=1.15,
            )
            if not line_fits:
                self.warn(str(slide.get("id", "")), "Cover bottom line overflowed.")
            line_text = line_text[0] if line_text else cover_line
            draw.text(
                (accent_x, band_y + (band_height - line_height(line_font)) // 2 - 2),
                line_text,
                font=line_font,
                fill=self.color("white"),
            )

        draw.rectangle(
            (24, int(self.layout["footer_y"]), self.width - 24, int(self.layout["footer_y"]) + 3),
            fill=self.color("blue"),
        )
        return image

    def draw_person_icon(self, draw: ImageDraw.ImageDraw, left: int, top: int, size: int, color: tuple[int, int, int]) -> None:
        draw.ellipse((left + size * 0.32, top, left + size * 0.68, top + size * 0.36), outline=color, width=2)
        draw.arc((left + size * 0.12, top + size * 0.42, left + size * 0.88, top + size), 200, 340, fill=color, width=2)

    def draw_cap_icon(self, draw: ImageDraw.ImageDraw, left: int, top: int, size: int, color: tuple[int, int, int]) -> None:
        draw.line((left, top + size * 0.34, left + size * 0.5, top), fill=color, width=2)
        draw.line((left + size * 0.5, top, left + size, top + size * 0.34), fill=color, width=2)
        draw.line((left, top + size * 0.34, left + size * 0.5, top + size * 0.68), fill=color, width=2)
        draw.line((left + size, top + size * 0.34, left + size * 0.5, top + size * 0.68), fill=color, width=2)
        draw.line((left + size * 0.5, top + size * 0.68, left + size * 0.5, top + size), fill=color, width=2)

    def draw_calendar_icon(self, draw: ImageDraw.ImageDraw, left: int, top: int, size: int, color: tuple[int, int, int]) -> None:
        draw.rectangle((left, top + size * 0.16, left + size, top + size), outline=color, width=2)
        draw.line((left, top + size * 0.42, left + size, top + size * 0.42), fill=color, width=2)
        draw.line((left + size * 0.28, top, left + size * 0.28, top + size * 0.22), fill=color, width=2)
        draw.line((left + size * 0.72, top, left + size * 0.72, top + size * 0.22), fill=color, width=2)

    def draw_outline(self, slide: dict) -> Image.Image:
        image = self.background(2)
        self.draw_header(image, slide)
        draw = ImageDraw.Draw(image)
        outline = self.theme["outline"]
        left = int(outline["x"])
        top = int(outline["y"])
        width = int(outline["width"])
        row_height = int(outline.get("row_height", 112))
        gap = int(outline.get("gap", 24))
        sections = self.deck.get("sections", [])
        spine_x = left - 34
        last_row_bottom = top + len(sections) * (row_height + gap) - gap
        index_style = str(self.layout.get("outline_style", "cards")).strip().lower() == "index"
        if index_style:
            # 参考图实测：单列五行铺满版心，序号在左、标题与说明在右、行间细线，
            # 右上角一个小号 CONTENTS。两列会把 5 条挤在上半页，整页留白失衡。
            items = list(sections)
            rows = max(1, len(items))
            top_y = int(self.layout.get("header_height", 132)) + 48
            bottom_y = int(self.layout.get("footer_y", self.height - 96)) - 48
            row_h = (bottom_y - top_y) // rows
            kicker_font = self.fonts.get(TextStyle(bold=False), 34)
            draw.text((left + width, top_y - 64), "C O N T E N T S", font=kicker_font,
                      fill=self.color("blue_mid"), anchor="ra")
            draw.line((left - 34, top_y + 6, left - 34, bottom_y - 6), fill=self.color("blue"), width=3)
            for index, section in enumerate(items):
                y = top_y + index * row_h
                preferred_size = max(62, self.type_size("header_title", 62))
                title_font, title_lines, _, _ = self.fit_text(
                    draw,
                    str(section["title"]),
                    self.theme["fonts"]["title"],
                    width - 170,
                    max(72, row_h - 48),
                    preferred_size,
                    max(40, preferred_size - 24),
                    max_lines=1,
                )
                desc = str(section.get("desc", "")).strip()
                desc_font = self.fonts.get(TextStyle(bold=False), 30)
                ascent, descent = title_font.getmetrics()
                desc_gap = 8
                desc_height = line_height(desc_font, 1.2) if desc else 0
                content_height = ascent + descent + (desc_gap + desc_height if desc else 0)
                baseline = y + (row_h - content_height) // 2 + ascent
                draw.text(
                    (left - 6, baseline),
                    f"{index + 1:02d}",
                    font=title_font,
                    fill=self.color("blue"),
                    anchor="ls",
                )
                title_text = title_lines[0] if title_lines else str(section["title"])
                draw.text(
                    (left + 150, baseline),
                    title_text,
                    font=title_font,
                    fill=self.color("ink"),
                    anchor="ls",
                )
                if desc:
                    draw.text(
                        (left + 150, baseline + descent + desc_gap),
                        desc,
                        font=desc_font,
                        fill=self.color("muted"),
                    )
                if index < len(items) - 1:
                    draw.line((left + 150, y + row_h - 12, left + width, y + row_h - 12),
                              fill=self.color("line"), width=2)
            return image

        lines_style = str(self.layout.get("outline_style", "cards")).strip().lower() == "lines"
        if not lines_style:
            draw.line((spine_x, top + 15, spine_x, last_row_bottom - 15), fill=self.color("line"), width=2)
        for index, section in enumerate(sections):
            row_top = top + index * (row_height + gap)
            if lines_style:
                number_font, _, _, _ = self.fit_text(
                    draw, f"{index + 1:02d}", self.theme["fonts"]["title"], 130, row_height - 8,
                    self.type_size("header_title", 46), 20, max_lines=1,
                )
                draw.text((left, row_top + (row_height - line_height(number_font)) // 2), f"{index + 1:02d}",
                          font=number_font, fill=self.color("blue"))
                title_font, title_lines, _, _ = self.fit_text(
                    draw, str(section["title"]), self.theme["fonts"]["title"], width - 210, row_height - 8,
                    self.type_size("block_title", 44), 20, max_lines=1,
                )
                draw.text((left + 140, row_top + (row_height - line_height(title_font)) // 2),
                          title_lines[0] if title_lines else str(section["title"]), font=title_font, fill=self.color("ink"))
                draw.line((left, row_top + row_height - 6, left + width, row_top + row_height - 6), fill=self.color("line"), width=2)
                continue
            draw.ellipse((spine_x - 6, row_top + row_height // 2 - 6, spine_x + 6, row_top + row_height // 2 + 6), fill=self.color("blue_mid"))
            draw_shadowed_panel(
                image,
                draw,
                (left, row_top, left + width, row_top + row_height),
                self.color("white"),
                self.color("line"),
                radius=9,
                width=2,
                shadow_offset=(0, 5),
                shadow_blur=9,
                shadow_alpha=24,
            )
            number_width = 168
            number_box = [
                (left + 2, row_top + 2),
                (left + number_width - 24, row_top + 2),
                (left + number_width, row_top + row_height // 2),
                (left + number_width - 24, row_top + row_height - 2),
                (left + 2, row_top + row_height - 2),
            ]
            draw.polygon(number_box, fill=self.color("blue"))
            title_font, title_lines, _, title_fits = self.fit_text(
                draw,
                section["title"],
                self.theme["fonts"]["title"],
                width - number_width - 110,
                row_height - 20,
                44,
                26,
                max_lines=1,
            )
            if not title_fits:
                self.warn(str(slide.get("id", "")), f"Outline item did not fit: {section['title']}")
            number_font = self.fonts.get(TextStyle(bold=True), title_font.size)
            number_text = f"{index + 1:02d}"
            number_bbox = draw.textbbox((0, 0), number_text, font=number_font)
            number_y = row_top + (row_height - (number_bbox[3] - number_bbox[1])) // 2 - number_bbox[1]
            draw.text((left + 46, number_y), number_text, font=number_font, fill=self.color("white"))
            title_y = row_top + (row_height - line_height(title_font)) // 2 - 2
            draw.text((left + number_width + 32, title_y), title_lines[0] if title_lines else section["title"], font=title_font, fill=self.color("blue_dark"))
            draw_chevrons(draw, left + width - 52, row_top + row_height // 2 - 9, 18, self.color("blue_pale"), count=1, width=3, gap=10)
        return image

    def draw_section(self, slide: dict, page_number: int) -> Image.Image:
        """章节分隔页：复刻参考模板的「蓝带 + 跨带大数字」版式。

        实测参考图（991x620）折算到 2048x1152：
          页眉细线 y≈0.19；蓝带 y 0.46–0.855，右端留白 3%；
          中文标题 100 px、英文 38 px、细白线 y≈0.79；
          大数字 x 0.605、宽 0.313、纵向跨带（0.242–0.63），上深下浅渐变。
        """
        image = self.background(page_number)
        draw = ImageDraw.Draw(image)
        width, height = self.width, self.height
        sections = self.deck.get("sections", []) or []
        section_id = str(slide.get("section", ""))
        index = next((i for i, item in enumerate(sections) if item.get("id") == section_id), 0)

        # 页眉：左校徽锁版（没有锁版就退回文字校名），右校训，细线
        self.draw_brand(image, (56, 30, 520, 196), align="left")
        draw = ImageDraw.Draw(image)
        motto = str(self.theme.get("brand", {}).get("motto") or "")
        if motto:
            motto_font = self.fonts.get(TextStyle(bold=False), 34)
            draw.text((width - 64, 96), motto, font=motto_font, fill=self.color("muted"), anchor="rt")
        rule_y = int(height * 0.19)
        draw.rectangle((40, rule_y, width - 40, rule_y + 3), fill=self.color("blue_dark"))

        # 蓝带
        band_top = int(height * 0.46)
        band_bottom = int(height * 0.855)
        band_right = int(width * 0.968)
        draw.rectangle((0, band_top, band_right, band_bottom), fill=self.color("blue"))

        # 跨带大数字：先做蒙版，再用纵向渐变填充
        number = f"{index + 1:02d}"
        num_left = int(width * 0.605)
        num_top = int(height * 0.222)
        num_w = int(width * 0.335)
        num_h = int(height * 0.44)
        probe = 400
        base_font = self.fonts.get(TextStyle(bold=True), probe)
        temp = Image.new("L", (probe * 3, probe * 2), 0)
        ImageDraw.Draw(temp).text((probe // 2, probe // 4), number, font=base_font, fill=255)
        bbox = temp.getbbox()
        glyph = temp.crop(bbox) if bbox else temp
        scale = min(num_w / glyph.width, num_h / glyph.height)
        size = (max(1, int(glyph.width * scale)), max(1, int(glyph.height * scale)))
        mask = glyph.resize(size, Image.Resampling.LANCZOS)
        gradient = Image.new("RGB", size)
        gdraw = ImageDraw.Draw(gradient)
        top_rgb = self.color("blue_dark")
        bottom_rgb = mix_color(self.color("blue_mid"), self.color("white"), 0.55)
        for row in range(size[1]):
            p = row / max(1, size[1] - 1)
            gdraw.line((0, row, size[0], row),
                       fill=tuple(int(top_rgb[c] * (1 - p) + bottom_rgb[c] * p) for c in range(3)))
        image.paste(gradient, (num_left + (num_w - size[0]) // 2, num_top + (num_h - size[1]) // 2), mask)
        draw = ImageDraw.Draw(image)

        # 中文标题 + 英文 + 细线
        title = str(slide.get("section_title") or slide.get("title") or "")
        title_font, lines, _, fits = self.fit_text(draw, title, self.theme["fonts"]["title"],
                                                  int(width * 0.42), 150, 100, 62, max_lines=1)
        if not fits:
            self.warn(str(slide.get("id", "")), "Section title overflowed.")
        title_y = band_top + int((band_bottom - band_top) * 0.24)
        draw.text((104, title_y), lines[0] if lines else title, font=title_font, fill=self.color("white"))

        label = str(slide.get("label_en") or "").upper()
        if label:
            label_font = self.fonts.get(TextStyle(bold=False), 38)
            draw.text((108, title_y + line_height(title_font) + 18), label, font=label_font,
                      fill=mix_color(self.color("white"), self.color("blue"), 0.25))

        line_y = band_bottom - int((band_bottom - band_top) * 0.19)
        line_left, line_right = 104, int(width * 0.90)
        draw.rectangle((line_left, line_y, line_right, line_y + 2),
                       fill=mix_color(self.color("white"), self.color("blue"), 0.35))
        draw.ellipse((line_right - 7, line_y - 5, line_right + 7, line_y + 9), fill=self.color("white"))
        return image

    def draw_frame(self, slide: dict, page_number: int) -> tuple[Image.Image, ImageDraw.ImageDraw]:
        self.current_slide = slide
        image = self.background(page_number)
        self.draw_header(image, slide)
        self.draw_sidebar(image, slide.get("section"), str(slide.get("id", "")))
        draw = ImageDraw.Draw(image)
        return image, draw

    def draw_module_header(
        self,
        image: Image.Image,
        draw: ImageDraw.ImageDraw,
        box: tuple[int, int, int, int],
        heading: str,
        index: str = "",
        icon: str = "",
        height: int = 56,
    ) -> int:
        """Draw one measured module header and return its real height.

        The index and the heading share one font size and one baseline. The
        index column is measured from the glyphs instead of a fixed 20 px
        stamp, so ``01`` can never be cropped to ``0``.
        """
        left, top, right, bottom = box
        configured_size = self.type_size("block_title", 26)
        minimum_height = int(round(configured_size * 1.1))
        if self.profile == "defense":
            minimum_height = max(minimum_height, 76)
        height = max(int(height), minimum_height)
        header_box = (left + 2, top + 2, right - 2, top + height)
        header_style = str(self.theme.get("panel", {}).get("header_style", "gradient"))
        if header_style == "flat":
            draw.rounded_rectangle(
                header_box,
                radius=self.panel_radius,
                fill=self.color("blue"),
            )
            draw.rectangle(
                (header_box[0] + 1, header_box[1] + 1, header_box[2] - 1, header_box[1] + 4),
                fill=mix_color(self.color("blue_mid"), self.color("white"), 0.18),
            )
        else:
            draw_gradient_panel(
                image,
                draw,
                header_box,
                self.color("blue_mid"),
                self.color("blue"),
                None,
                radius=self.panel_radius,
                width=0,
            )
            draw.rectangle((header_box[0], header_box[3] - 1, header_box[2], header_box[3]), fill=self.color("blue"))
        probe_font = self.fonts.get(TextStyle(bold=True), configured_size)
        index_width = 0
        if index and getattr(self, "card_index", True):
            index_width = max(60, int(round(text_width(draw, index, probe_font) + 20)))
        text_left = left + 16 + (index_width + 12 if index_width else 0)
        text_height = height - 12
        heading_size = configured_size
        fits = True
        rendered = probe_font
        lines = [heading] if heading else []
        if heading:
            rendered, lines, _, fits = self.fit_text(
                draw,
                heading,
                self.theme["fonts"]["title"],
                right - text_left - (56 if icon else 18),
                text_height,
                configured_size,
                17,
                max_lines=1,
            )
            heading_size = rendered.size
        if index_width:
            index_font = self.fonts.get(TextStyle(bold=True), heading_size)
            ascent, descent = index_font.getmetrics()
            index_y = top + 6 + (text_height - (ascent + descent)) // 2
            draw.text((left + 16, index_y), index, font=index_font, fill=self.color("blue_pale"))
        if heading:
            ascent, descent = rendered.getmetrics()
            heading_y = top + 6 + (text_height - (ascent + descent)) // 2
            draw.text((text_left, heading_y), lines[0] if lines else heading, font=rendered, fill=self.color("white"))
        if not fits:
            self.warn("panel", f"Panel heading overflowed: {heading}")
        if icon:
            draw_icon(
                draw,
                icon,
                (right - 48, top + (height - 32) // 2, right - 16, top + (height + 32) // 2),
                self.color("white"),
                self.color("blue"),
                stroke=2,
            )
        return height

    def draw_compact_bullets(
        self,
        draw: ImageDraw.ImageDraw,
        bullets: list[str],
        left: int,
        top: int,
        width: int,
        height: int,
        max_size: int = 21,
        min_size: int = 15,
        gap: int = 10,
    ) -> tuple[bool, int]:
        max_size = max(max_size, min_size)
        current_top = top
        all_fit = True
        for bullet in bullets:
            font, lines, block_height, fits = self.fit_text(
                draw,
                bullet,
                self.theme["fonts"]["body"],
                width - 26,
                min(64, height),
                max_size,
                min_size,
                max_lines=3,
                line_ratio=1.24,
            )
            if current_top + block_height > top + height:
                all_fit = False
                break
            size = 5
            y = current_top + 9
            draw.polygon(
                [(left + 3, y), (left + 3 + size, y + size), (left + 3, y + size * 2), (left + 3 - size, y + size)],
                fill=self.color("blue_mid"),
            )
            draw_lines(draw, lines, (left + 24, current_top), font, self.color("ink"), 1.24)
            current_top += max(block_height, line_height(font, 1.24)) + gap
            all_fit = all_fit and fits
        return all_fit, current_top

    def draw_bullets(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        columns = slide.get("columns") or []
        if not columns and slide.get("bullets"):
            columns = [{"heading": "要点", "bullets": slide.get("bullets", [])}]
        if not columns:
            self.warn(str(slide.get("id", "")), "Bullets slide has no content.")
            columns = [{"heading": "内容待补充", "bullets": []}]
        weights = slide.get("column_weights")
        if not weights or len(weights) != len(columns):
            weights = [1.35, 1.0, 1.0][: len(columns)] if len(columns) == 3 else [1.0] * len(columns)
        gap = 18
        available = right - left - gap * (len(columns) - 1)
        widths = [round(available * weight / sum(weights)) for weight in weights]
        widths[-1] = available - sum(widths[:-1])
        x = left
        for index, (column, width) in enumerate(zip(columns, widths)):
            box = (x, body_top, x + width, bottom)
            draw_shadowed_panel(
                image,
                draw,
                box,
                self.color("blue_wash") if index == 0 else self.color("white"),
                self.color("line"),
                radius=10,
                width=2,
                shadow_offset=(0, 6),
                shadow_blur=10,
                shadow_alpha=26,
            )
            header_height = self.draw_module_header(
                image,
                draw,
                box,
                str(column.get("heading", "")),
                index=f"{index + 1:02d}",
                icon=str(column.get("icon", "")),
            )
            content_top = body_top + header_height + 12
            body_text = str(column.get("body", "")).strip()
            if body_text:
                fits, used = draw_rich_block(
                    draw,
                    body_text,
                    (x + 18, content_top),
                    width - 36,
                    100,
                    self.fonts,
                    self.colors,
                    self.type_size("body", 22),
                    self.type_size("small", 16),
                    max_lines=3,
                    line_ratio=1.28,
                )
                if not fits:
                    self.warn(str(slide.get("id", "")), f"Column body overflowed: {column.get('heading', '')}")
                content_top += used + 12
            bullets = column.get("bullets") or []
            if bullets:
                fits, _ = self.draw_compact_bullets(
                    draw,
                    bullets,
                    x + 18,
                    content_top,
                    width - 36,
                    bottom - content_top - 18,
                    max_size=self.type_size("bullet", 21),
                    min_size=self.type_size("micro", 14),
                )
                if not fits:
                    self.warn(str(slide.get("id", "")), f"Column bullets overflowed: {column.get('heading', '')}")
            x += width + gap
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_grid_block(
        self,
        image: Image.Image,
        draw: ImageDraw.ImageDraw,
        block: dict,
        box: tuple[int, int, int, int],
        index: int,
        page_id: str,
    ) -> None:
        left, top, right, bottom = box
        fill = self.color("blue_wash") if block.get("tone") == "pale" else self.color("white")
        draw_shadowed_panel(
            image,
            draw,
            box,
            fill,
            self.color("line"),
            radius=self.panel_radius,
            width=1,
            shadow_offset=(0, 5),
            shadow_blur=8,
            shadow_alpha=self.panel_shadow_alpha,
        )
        header_height = self.draw_module_header(
            image,
            draw,
            box,
            str(block.get("heading", "")),
            index=f"{index:02d}",
            icon=str(block.get("icon", "")),
        )
        content_top = top + header_height + 12
        footer_text = str(block.get("footer") or block.get("output") or "").strip()
        footer_label = str(block.get("footer_label") or "产出").strip()
        footer_height = 52 if footer_text else 0
        footer_top = bottom - footer_height - 12 if footer_text else bottom
        content_bottom = footer_top - 8 if footer_text else bottom - 14
        image_path = resolve_path(block.get("image"), self.deck_path.parent)
        image_side = str(block.get("image_side", "right"))
        if image_path is not None and image_path.exists():
            if (right - left) > 520:
                if image_side == "left":
                    image_box = (left + 14, content_top, left + int((right - left) * 0.46), content_bottom)
                    text_left = image_box[2] + 14
                    text_width_available = right - text_left - 14
                else:
                    image_box = (right - int((right - left) * 0.46), content_top, right - 14, content_bottom)
                    text_left = left + 16
                    text_width_available = image_box[0] - text_left - 14
                draw_figure_card(
                    image,
                    draw,
                    image_box,
                    image_path,
                    "",
                    "",
                    self.colors,
                    self.fonts,
                    badge=str(block.get("badge", "")),
                    scale=1.8 if self.profile == "defense" else 1.0,
                )
                draw = ImageDraw.Draw(image)
            else:
                image_box = (left + 14, content_top, right - 14, content_top + int((content_bottom - content_top) * 0.38))
                draw_figure_card(
                    image,
                    draw,
                    image_box,
                    image_path,
                    "",
                    "",
                    self.colors,
                    self.fonts,
                    badge=str(block.get("badge", "")),
                    scale=1.8 if self.profile == "defense" else 1.0,
                )
                draw = ImageDraw.Draw(image)
                content_top = image_box[3] + 10
                text_left = left + 16
                text_width_available = right - left - 32
        else:
            text_left = left + 16
            text_width_available = right - left - 32
        stat = str(block.get("stat", "")).strip()
        if stat:
            stat_font = self.fonts.get(TextStyle(bold=True), 48)
            draw.text((text_left, content_top + 2), stat, font=stat_font, fill=self.color("blue"))
            stat_width = text_width(draw, stat, stat_font)
            stat_label = str(block.get("stat_label", "")).strip()
            if stat_label:
                label_font = self.fonts.get(TextStyle(), 18)
                draw.text((text_left + stat_width + 12, content_top + 20), stat_label, font=label_font, fill=self.color("muted"))
            content_top += 62
        formula = str(block.get("formula", "")).strip()
        formula_latex = str(block.get("formula_latex", "")).strip()
        if formula or formula_latex:
            formula_box = (text_left, content_top, text_left + text_width_available, content_top + 58)
            draw_gradient_panel(image, draw, formula_box, self.color("blue_wash"), self.color("white"), self.color("line"), radius=5, width=1)
            math_image = (
                render_math_image(formula_latex, self.color("blue_dark"), font_size=34)
                if formula_latex
                else None
            )
            if math_image is not None:
                inner_width = text_width_available - 28
                inner_height = 42
                scale = min(inner_width / math_image.width, inner_height / math_image.height)
                size = (
                    max(1, round(math_image.width * scale)),
                    max(1, round(math_image.height * scale)),
                )
                resized = math_image.resize(size, Image.Resampling.LANCZOS)
                paste_x = formula_box[0] + (formula_box[2] - formula_box[0] - size[0]) // 2
                paste_y = formula_box[1] + (formula_box[3] - formula_box[1] - size[1]) // 2
                image.paste(resized, (paste_x, paste_y), resized)
                draw = ImageDraw.Draw(image)
            else:
                fallback_formula = formula or formula_latex
                formula_font, formula_lines, _, formula_fits = self.fit_text(
                    draw,
                    fallback_formula,
                    self.formula_fonts,
                    text_width_available - 18,
                    42,
                    27,
                    15,
                    max_lines=1,
                )
                if not formula_fits:
                    self.warn(page_id, f"Formula overflowed: {fallback_formula}")
                lines_height = line_height(formula_font, 1.1) * max(1, len(formula_lines))
                draw_lines(
                    draw,
                    formula_lines,
                    (formula_box[0] + 9, formula_box[1] + (formula_box[3] - formula_box[1] - lines_height) // 2),
                    formula_font,
                    self.color("blue_dark"),
                    1.1,
                    align="center",
                    max_width=text_width_available - 18,
                )
            content_top = formula_box[3] + 10
        body = str(block.get("body", "")).strip()
        if body:
            fits, used = draw_rich_block(
                draw,
                body,
                (text_left, content_top),
                text_width_available,
                max(34, content_bottom - content_top - 10),
                self.fonts,
                self.colors,
                int(block.get("body_size", self.type_size("body", 22))),
                self.type_size("small", 15),
                max_lines=4,
                line_ratio=1.26,
            )
            if not fits:
                self.warn(page_id, f"Grid body overflowed: {block.get('heading', '')}")
            content_top += used + 8
        diagram = block.get("diagram")
        has_block_image = image_path is not None and image_path.exists()
        if isinstance(diagram, dict) and diagram and not has_block_image and content_bottom - content_top > 96:
            diagram_height = int(
                block.get(
                    "diagram_height",
                    min(170, max(88, int((content_bottom - content_top) * 0.46))),
                )
            )
            draw_diagram(
                draw,
                (text_left, content_top, right - 16, content_top + diagram_height),
                diagram,
                self.colors,
                self.fonts,
                label=str(block.get("diagram_label", "")),
                frame=True,
                min_size=DEFENSE_MIN_ANY_PX if self.profile == "defense" else 0,
            )
            content_top += diagram_height + 8
        visual = str(block.get("visual", "")).strip() if block.get("visual_semantic") else ""
        bullets_height = content_bottom - content_top - 10
        if visual and content_bottom - content_top > 70:
            visual_height = max(52, min(104, int((content_bottom - content_top) * 0.34)))
            visual_top = content_bottom - visual_height
            draw_mini_visual(
                draw,
                (text_left, visual_top, right - 16, content_bottom),
                visual,
                self.colors,
                self.fonts,
            )
            bullets_height = visual_top - content_top - 6
        bullets = block.get("bullets") or []
        if bullets:
            fits, _ = self.draw_compact_bullets(
                draw,
                bullets,
                text_left,
                content_top,
                text_width_available,
                bullets_height,
                max_size=self.type_size("bullet", 19),
                min_size=self.type_size("micro", 14),
                gap=7,
            )
            if not fits:
                self.warn(page_id, f"Grid bullets overflowed: {block.get('heading', '')}")
        if footer_text:
            self.draw_output_strip(
                draw,
                (left + 14, footer_top, right - 14, bottom - 12),
                footer_label,
                footer_text,
                page_id,
            )

    def draw_output_strip(
        self,
        draw: ImageDraw.ImageDraw,
        box: tuple[int, int, int, int],
        label: str,
        text: str,
        page_id: str,
    ) -> None:
        left, top, right, bottom = box
        if not text or right - left < 80 or bottom - top < 28:
            return
        draw.rounded_rectangle(
            box,
            radius=4,
            fill=mix_color(self.color("blue_wash"), self.color("white"), 0.08),
            outline=mix_color(self.color("line"), self.color("white"), 0.18),
            width=1,
        )
        label_font = self.fonts.get(TextStyle(bold=True), 14)
        label_width = min(
            max(58, text_width(draw, label, label_font) + 24),
            max(58, (right - left) // 3),
        )
        draw.rounded_rectangle(
            (left + 1, top + 1, left + label_width, bottom - 1),
            radius=3,
            fill=self.color("blue"),
        )
        draw.text(
            (left + (label_width - text_width(draw, label, label_font)) // 2, top + (bottom - top - 17) // 2 - 1),
            label,
            font=label_font,
            fill=self.color("white"),
        )
        text_left = left + label_width + 12
        text_width_available = right - text_left - 10
        font, lines, _, fits = self.fit_text(
            draw,
            text,
            self.theme["fonts"]["body"],
            text_width_available,
            bottom - top - 10,
            17,
            12,
            max_lines=2,
            line_ratio=1.15,
        )
        if not fits:
            self.warn(page_id, f"Output strip overflowed: {text}")
        block_height = line_height(font, 1.15) * max(1, len(lines))
        draw_lines(
            draw,
            lines,
            (text_left, top + (bottom - top - block_height) // 2 - 1),
            font,
            self.color("blue_dark"),
            1.15,
        )

    def draw_grid(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        blocks = slide.get("blocks") or []
        if not blocks:
            self.warn(str(slide.get("id", "")), "Grid layout has no blocks.")
            blocks = [{"heading": "内容待补充", "body": ""}]
        columns = max(2, min(3, int(slide.get("columns", 2))))
        rows: list[list[tuple[dict, int]]] = []
        current: list[tuple[dict, int]] = []
        current_span = 0
        for block in blocks:
            span = max(1, min(columns, int(block.get("span", 1))))
            if current and current_span + span > columns:
                rows.append(current)
                current = []
                current_span = 0
            current.append((block, span))
            current_span += span
            if current_span == columns:
                rows.append(current)
                current = []
                current_span = 0
        if current:
            rows.append(current)
        gap = 18
        cell_height = (bottom - body_top - gap * (len(rows) - 1)) // max(1, len(rows))
        block_index = 1
        for row_index, row in enumerate(rows):
            row_top = body_top + row_index * (cell_height + gap)
            row_span = sum(span for _, span in row)
            available = right - left - gap * (len(row) - 1)
            widths = [round(available * span / row_span) for _, span in row]
            widths[-1] = available - sum(widths[:-1])
            x = left
            for (block, span), width in zip(row, widths):
                self.draw_grid_block(
                    image,
                    draw,
                    block,
                    (x, row_top, x + width, row_top + cell_height),
                    block_index,
                    str(slide.get("id", "")),
                )
                draw = ImageDraw.Draw(image)
                x += width + gap
                block_index += 1
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_balance_rail(
        self,
        image: Image.Image,
        draw: ImageDraw.ImageDraw,
        rail: dict,
        box: tuple[int, int, int, int],
        index: int,
        page_id: str,
    ) -> None:
        left, top, right, bottom = box
        draw_shadowed_panel(
            image,
            draw,
            box,
            self.color("blue_wash") if index == 0 else self.color("white"),
            self.color("line"),
            radius=8,
            width=2,
            shadow_offset=(0, 4),
            shadow_blur=8,
            shadow_alpha=18,
        )
        header_height = self.draw_module_header(
            image,
            draw,
            box,
            str(rail.get("heading", "")),
            icon=str(rail.get("icon", "check" if index == 0 else "alert")),
        )
        items = rail.get("items") or []
        if not items:
            self.warn(page_id, "Balance rail has no items.")
            return
        header_bottom = top + header_height
        item_height = (bottom - header_bottom - 8) // len(items)
        for item_index, item in enumerate(items):
            item_top = header_bottom + item_index * item_height
            if item_index:
                draw.line((left + 18, item_top, right - 18, item_top), fill=self.color("line"), width=1)
            icon_size = 80 if item_height >= 180 else 66
            draw_icon_badge(
                image,
                draw,
                (left + 54, item_top + item_height // 2 - 4),
                icon_size,
                str(item.get("icon", "check" if index == 0 else "alert")),
                self.colors,
            )
            text_left = left + 106
            text_width_available = right - text_left - 20
            heading_font, heading_lines, _, heading_fits = self.fit_text(
                draw,
                str(item.get("heading", "")),
                self.theme["fonts"]["title"],
                text_width_available,
                36,
                30,
                20,
                max_lines=1,
            )
            if not heading_fits:
                self.warn(page_id, f"Balance item heading overflowed: {item.get('heading', '')}")
            heading_y = item_top + 24
            draw.text(
                (text_left, heading_y),
                heading_lines[0] if heading_lines else str(item.get("heading", "")),
                font=heading_font,
                fill=self.color("blue_dark"),
            )
            bullets = item.get("bullets") or []
            body = str(item.get("body", "")).strip()
            body_top = heading_y + 46
            body_height = max(28, item_top + item_height - body_top - 14)
            if body:
                fits, used = draw_rich_block(
                    draw,
                    body,
                    (text_left, body_top),
                    text_width_available,
                    body_height,
                    self.fonts,
                    self.colors,
                    22,
                    17,
                    max_lines=3,
                    line_ratio=1.22,
                )
                if not fits:
                    self.warn(page_id, f"Balance item body overflowed: {item.get('heading', '')}")
                body_top += min(used, body_height) + 4
            if bullets and not body:
                fits, _ = self.draw_compact_bullets(
                    draw,
                    bullets[:2],
                    text_left,
                    body_top,
                    text_width_available,
                    max(18, item_top + item_height - body_top - 6),
                    max_size=17,
                    min_size=12,
                    gap=3,
                )
                if not fits:
                    self.warn(page_id, f"Balance item bullets overflowed: {item.get('heading', '')}")

    def draw_balance(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        gap = 24
        rail_width = int((right - left) * 0.255)
        center_left = left + rail_width + gap
        center_right = right - rail_width - gap
        if center_right <= center_left:
            rail_width = int((right - left) * 0.22)
            center_left = left + rail_width + gap
            center_right = right - rail_width - gap
        left_rail = slide.get("left") or {}
        right_rail = slide.get("right") or {}
        self.draw_balance_rail(
            image,
            draw,
            left_rail,
            (left, body_top, left + rail_width, bottom),
            0,
            str(slide.get("id", "")),
        )
        self.draw_balance_rail(
            image,
            draw,
            right_rail,
            (right - rail_width, body_top, right, bottom),
            1,
            str(slide.get("id", "")),
        )
        draw = ImageDraw.Draw(image)
        draw_figure_card(
            image,
            draw,
            (center_left, body_top, center_right, bottom),
            resolve_path(slide.get("image"), self.deck_path.parent),
            str(slide.get("image_label", "系统结构")),
            str(slide.get("image_caption", "")),
            self.colors,
            self.fonts,
            chart=slide.get("chart"),
            badge=str(slide.get("image_badge", "")),
            scale=1.8 if self.profile == "defense" else 1.0,
        )
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_hub_card(
        self,
        image: Image.Image,
        draw: ImageDraw.ImageDraw,
        item: dict,
        box: tuple[int, int, int, int],
        index: int,
        page_id: str,
    ) -> None:
        left, top, right, bottom = box
        fill = self.color("blue_wash") if item.get("tone") == "pale" else self.color("white")
        draw_shadowed_panel(
            image,
            draw,
            box,
            fill,
            self.color("line"),
            radius=self.panel_radius,
            width=1,
            shadow_offset=(0, 4),
            shadow_blur=7,
            shadow_alpha=self.panel_shadow_alpha,
        )
        draw.rectangle((left, top + 15, left + 5, bottom - 15), fill=self.color("blue") if index % 2 else self.color("blue_mid"))
        draw_icon_badge(
            image,
            draw,
            (left + 42, top + 40),
            42,
            str(item.get("icon", "target")),
            self.colors,
        )
        text_left = left + 76
        text_width_available = right - text_left - 18
        heading_font, heading_lines, _, heading_fits = self.fit_text(
            draw,
            str(item.get("heading", "")),
            self.theme["fonts"]["title"],
            text_width_available,
            34,
            26,
            16,
            max_lines=1,
        )
        if not heading_fits:
            self.warn(page_id, f"Hub heading overflowed: {item.get('heading', '')}")
        draw.text(
            (text_left, top + 17),
            heading_lines[0] if heading_lines else str(item.get("heading", "")),
            font=heading_font,
            fill=self.color("blue_dark"),
        )
        body_top = top + 58
        visual = str(item.get("visual", "")).strip() if item.get("visual_semantic") else ""
        bullets = item.get("bullets") or []
        visual_height = 58 if visual and bottom - body_top > 128 else 0
        bullets_height = 48 if bullets and bottom - body_top > 108 else 0
        body_height = max(30, bottom - body_top - 12 - visual_height - bullets_height)
        body = str(item.get("body", "")).strip()
        if body:
            fits, used = draw_rich_block(
                draw,
                body,
                (left + 16, body_top),
                right - left - 32,
                body_height,
                self.fonts,
                self.colors,
                19,
                13,
                max_lines=3,
                line_ratio=1.22,
            )
            if not fits:
                self.warn(page_id, f"Hub body overflowed: {item.get('heading', '')}")
            body_top += min(body_height, used) + 4
        if bullets:
            fits, _ = self.draw_compact_bullets(
                draw,
                bullets[:2],
                left + 16,
                body_top,
                right - left - 32,
                max(20, bottom - body_top - visual_height - 8),
                max_size=15,
                min_size=11,
                gap=3,
            )
            if not fits:
                self.warn(page_id, f"Hub bullets overflowed: {item.get('heading', '')}")
        if visual and visual_height:
            draw_mini_visual(
                draw,
                (left + 16, bottom - visual_height - 10, right - 16, bottom - 10),
                visual,
                self.colors,
                self.fonts,
            )

    def draw_hub(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        items = slide.get("items") or []
        if not 4 <= len(items) <= 6:
            self.warn(str(slide.get("id", "")), "Hub layout expects four to six surrounding items.")
        items = items[:6]
        if not items:
            items = [{"heading": "内容待补充", "body": ""}]
        left_count = math.ceil(len(items) / 2)
        left_items = items[:left_count]
        right_items = items[left_count:]
        rows = max(len(left_items), len(right_items))
        gap = 18
        side_width = int((right - left - 380 - 80) / 2)
        side_width = max(420, min(side_width, 650))
        card_height = min(244, (bottom - body_top - gap * (rows - 1)) // max(1, rows))
        block_height = rows * card_height + gap * (rows - 1)
        start_y = body_top + max(0, (bottom - body_top - block_height) // 2)
        center_cx = (left + right) // 2
        center_cy = body_top + (bottom - body_top) // 2
        center_radius = min(162, (bottom - body_top) // 2 - 24, side_width // 2)
        side_left_x = left
        side_right_x = right - side_width
        for column_items, is_left in ((left_items, True), (right_items, False)):
            for index, item in enumerate(column_items):
                card_y = start_y + index * (card_height + gap)
                anchor_y = card_y + card_height // 2
                if is_left:
                    start = (center_cx - center_radius + 6, center_cy)
                    end = (side_left_x + side_width + 9, anchor_y)
                else:
                    start = (center_cx + center_radius - 6, center_cy)
                    end = (side_right_x - 9, anchor_y)
                draw_arrow(draw, start, end, mix_color(self.color("blue_mid"), self.color("line"), 0.32), width=2, head=9)
        for column_items, x in ((left_items, side_left_x), (right_items, side_right_x)):
            for index, item in enumerate(column_items):
                card_y = start_y + index * (card_height + gap)
                self.draw_hub_card(
                    image,
                    draw,
                    item,
                    (x, card_y, x + side_width, card_y + card_height),
                    index + (0 if x == side_left_x else left_count),
                    str(slide.get("id", "")),
                )
                draw = ImageDraw.Draw(image)
        center_box = (
            center_cx - center_radius,
            center_cy - center_radius,
            center_cx + center_radius,
            center_cy + center_radius,
        )
        draw.ellipse(center_box, fill=self.color("blue_wash"), outline=self.color("blue_mid"), width=3)
        draw.ellipse(
            (
                center_box[0] + 12,
                center_box[1] + 12,
                center_box[2] - 12,
                center_box[3] - 12,
            ),
            outline=mix_color(self.color("blue_pale"), self.color("white"), 0.18),
            width=2,
        )
        center = slide.get("center") or {}
        label = str(center.get("label", ""))
        value = str(center.get("value", ""))
        note = str(center.get("note", ""))
        label_font = self.fonts.get(TextStyle(bold=True), 34)
        if label:
            label_width = text_width(draw, label, label_font)
            draw.text((center_cx - label_width // 2, center_cy - 92), label, font=label_font, fill=self.color("blue_dark"))
        if value:
            value_font, value_lines, _, value_fits = self.fit_text(
                draw,
                value,
                self.display_fonts,
                center_radius * 2 - 44,
                92,
                100,
                46,
                max_lines=1,
            )
            if not value_fits:
                self.warn(str(slide.get("id", "")), "Hub center value overflowed.")
            value_width = text_width(draw, value_lines[0] if value_lines else value, value_font)
            draw.text(
                (center_cx - value_width // 2, center_cy - 44),
                value_lines[0] if value_lines else value,
                font=value_font,
                fill=self.color("blue"),
            )
        if note:
            note_font, note_lines, _, note_fits = self.fit_text(
                draw,
                note,
                self.theme["fonts"]["body"],
                center_radius * 2 - 34,
                46,
                18,
                12,
                max_lines=2,
            )
            if not note_fits:
                self.warn(str(slide.get("id", "")), "Hub center note overflowed.")
            draw_lines(
                draw,
                note_lines,
                (center_cx - (center_radius - 18), center_cy + 66),
                note_font,
                self.color("muted"),
                align="center",
                max_width=center_radius * 2 - 36,
            )
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_figure_panels(
        self,
        image: Image.Image,
        draw: ImageDraw.ImageDraw,
        slide: dict,
        box: tuple[int, int, int, int],
    ) -> None:
        left, top, right, bottom = box
        panels = slide.get("panels") or []
        if not panels:
            panels = [{"heading": "图示", "body": slide.get("image_role", "请结合原图讲解。")}]
        gap = 10
        panel_height = (bottom - top - gap * (len(panels) - 1)) // len(panels)
        for index, panel in enumerate(panels):
            panel_top = top + index * (panel_height + gap)
            panel_box = (left, panel_top, right, panel_top + panel_height)
            draw_shadowed_panel(
                image,
                draw,
                panel_box,
                self.color("white"),
                self.color("line"),
                radius=8,
                width=2,
                shadow_offset=(0, 4),
                shadow_blur=8,
                shadow_alpha=20,
            )
            header_height = self.draw_module_header(
                image,
                draw,
                panel_box,
                str(panel.get("heading", "")),
                index=f"{index + 1:02d}",
                icon=str(panel.get("icon", "")),
            )
            content_top = panel_top + header_height + 12
            body = str(panel.get("body", "")).strip()
            bullets = panel.get("bullets") or []
            if body:
                fits, used = draw_rich_block(
                    draw,
                    body,
                    (left + 16, content_top),
                    right - left - 32,
                    max(30, panel_top + panel_height - content_top - 10),
                    self.fonts,
                    self.colors,
                    self.type_size("body", 18),
                    self.type_size("small", 14),
                    max_lines=2 if bullets else 3,
                    line_ratio=1.25,
                )
                if not fits:
                    self.warn(str(slide.get("id", "")), f"Figure panel overflowed: {panel.get('heading', '')}")
                content_top += used + 6
            if bullets:
                fits, _ = self.draw_compact_bullets(
                    draw,
                    bullets,
                    left + 16,
                    content_top,
                    right - left - 32,
                    panel_top + panel_height - content_top - 10,
                    max_size=16,
                    min_size=12,
                    gap=4,
                )
                if not fits:
                    self.warn(str(slide.get("id", "")), f"Figure panel bullets overflowed: {panel.get('heading', '')}")

    def draw_figure(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        image_path = resolve_path(slide.get("image"), self.deck_path.parent)
        side = str(slide.get("image_side", "left"))
        variant = str(slide.get("variant", "")).strip().lower()
        if variant == "image_right":
            side = "right"
        elif variant == "image_left":
            side = "left"
        total_width = right - left
        gap = 20
        image_width = round(total_width * figure_width_ratio(slide))
        text_width_available = total_width - image_width - gap
        if side == "right":
            image_box = (right - image_width, body_top, right, bottom)
            text_box = (left, body_top, left + text_width_available, bottom)
        else:
            image_box = (left, body_top, left + image_width, bottom)
            text_box = (left + image_width + gap, body_top, right, bottom)
        draw_figure_card(
            image,
            draw,
            image_box,
            image_path,
            str(slide.get("image_label", "图版")),
            str(slide.get("image_caption", "")),
            self.colors,
            self.fonts,
            chart=slide.get("chart"),
            badge=str(slide.get("image_badge", "")),
            scale=1.8 if self.profile == "defense" else 1.0,
        )
        draw = ImageDraw.Draw(image)
        self.draw_figure_panels(image, draw, slide, text_box)
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_figure_flow(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        flow = slide.get("flow") or []
        if not flow:
            self.warn(str(slide.get("id", "")), "Figure-flow slide has no flow steps.")
            flow = [{"title": "步骤", "body": ""}]
        flow_gap = 18
        flow_width = (right - left - flow_gap * (len(flow) - 1)) // len(flow)
        flow_height = 112
        for index, step in enumerate(flow):
            x = left + index * (flow_width + flow_gap)
            box = (x, body_top, x + flow_width, body_top + flow_height)
            draw_shadowed_panel(
                image,
                draw,
                box,
                self.color("blue_wash") if index % 2 == 0 else self.color("white"),
                self.color("line"),
                radius=8,
                width=2,
                shadow_offset=(0, 4),
                shadow_blur=7,
                shadow_alpha=18,
            )
            draw_icon_badge(
                image,
                draw,
                (x + 38, body_top + flow_height // 2),
                52,
                str(step.get("icon", "signal")),
                self.colors,
            )
            title_font, title_lines, _, title_fits = self.fit_text(
                draw,
                str(step.get("title", "")),
                self.theme["fonts"]["title"],
                flow_width - 86,
                30,
                21,
                14,
                max_lines=1,
            )
            if not title_fits:
                self.warn(str(slide.get("id", "")), f"Flow title overflowed: {step.get('title', '')}")
            draw.text((x + 74, body_top + 18), title_lines[0] if title_lines else str(step.get("title", "")), font=title_font, fill=self.color("blue_dark"))
            body_font, body_lines, _, body_fits = self.fit_text(
                draw,
                str(step.get("body", "")),
                self.theme["fonts"]["body"],
                flow_width - 86,
                42,
                15,
                10,
                max_lines=2,
            )
            if not body_fits:
                self.warn(str(slide.get("id", "")), f"Flow body overflowed: {step.get('title', '')}")
            draw_lines(draw, body_lines, (x + 74, body_top + 52), body_font, self.color("muted"), 1.2)
            if index < len(flow) - 1:
                draw_chevrons(draw, x + flow_width + 3, body_top + 46, 20, self.color("blue_mid"), count=1, width=3, gap=8)
        lower_top = body_top + flow_height + 16
        lower_bottom = bottom
        image_width = round((right - left) * figure_width_ratio(slide, 0.60))
        text_left = left + image_width + 20
        draw_figure_card(
            image,
            draw,
            (left, lower_top, left + image_width, lower_bottom),
            resolve_path(slide.get("image"), self.deck_path.parent),
            str(slide.get("image_label", "图版")),
            str(slide.get("image_caption", "")),
            self.colors,
            self.fonts,
            chart=slide.get("chart"),
            badge=str(slide.get("image_badge", "")),
            scale=1.8 if self.profile == "defense" else 1.0,
        )
        draw = ImageDraw.Draw(image)
        self.draw_figure_panels(image, draw, slide, (text_left, lower_top, right, lower_bottom))
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_concept(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        diagram = slide.get("diagram") or {}
        panels = slide.get("panels") or []
        gap = 18
        panel_width = 0
        if panels:
            diagram_ratio = float(slide.get("diagram_width_ratio", 0.58))
            diagram_ratio = max(0.48, min(0.68, diagram_ratio))
            diagram_right = left + round((right - left) * diagram_ratio)
            panel_width = right - diagram_right - gap
        else:
            diagram_right = right
        draw_diagram(
            draw,
            (left, body_top, diagram_right, bottom),
            diagram,
            self.colors,
            self.fonts,
            label=str(slide.get("diagram_label", "机制示意")),
            frame=True,
            min_size=DEFENSE_MIN_ANY_PX if self.profile == "defense" else 0,
        )
        if panels:
            panel_left = diagram_right + gap
            panel_gap = 10
            panel_height = (bottom - body_top - panel_gap * (len(panels) - 1)) // max(1, len(panels))
            for index, panel in enumerate(panels):
                panel_top = body_top + index * (panel_height + panel_gap)
                panel_box = (panel_left, panel_top, panel_left + panel_width, panel_top + panel_height)
                draw_shadowed_panel(
                    image,
                    draw,
                    panel_box,
                    self.color("white") if index % 2 == 0 else self.color("blue_wash"),
                    self.color("line"),
                    radius=self.panel_radius,
                    width=1,
                    shadow_offset=(0, 3),
                    shadow_blur=6,
                    shadow_alpha=max(10, self.panel_shadow_alpha - 6),
                )
                header_height = self.draw_module_header(
                    image,
                    draw,
                    panel_box,
                    str(panel.get("heading", "")),
                    index=f"{index + 1:02d}",
                    icon=str(panel.get("icon", "")),
                )
                content_top = panel_top + header_height + 12
                body = str(panel.get("body", "")).strip()
                bullets = panel.get("bullets") or []
                if body:
                    fits, used = draw_rich_block(
                        draw,
                        body,
                        (panel_left + 14, content_top),
                        panel_width - 28,
                        max(34, panel_top + panel_height - content_top - 10),
                        self.fonts,
                        self.colors,
                        self.type_size("body", 20),
                        self.type_size("small", 14),
                        max_lines=4 if not bullets else 3,
                        line_ratio=1.22,
                    )
                    if not fits:
                        self.warn(str(slide.get("id", "")), f"Concept panel body overflowed: {panel.get('heading', '')}")
                    content_top += used + 8
                if bullets:
                    fits, _ = self.draw_compact_bullets(
                        draw,
                        bullets,
                        panel_left + 14,
                        content_top,
                        panel_width - 28,
                        panel_top + panel_height - content_top - 10,
                        max_size=self.type_size("bullet", 17),
                        min_size=self.type_size("micro", 12),
                        gap=5,
                    )
                    if not fits:
                        self.warn(str(slide.get("id", "")), f"Concept panel bullets overflowed: {panel.get('heading', '')}")
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_formula(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        gap = 18
        left_ratio = float(slide.get("formula_width_ratio", 0.62))
        left_ratio = max(0.52, min(0.70, left_ratio))
        left_right = left + round((right - left) * left_ratio)
        right_left = left_right + gap
        formula_panel_height = int(slide.get("formula_panel_height", 236))
        formula_box = (left, body_top, left_right, body_top + formula_panel_height)
        draw_gradient_panel(
            image,
            draw,
            formula_box,
            self.color("blue_wash"),
            self.color("white"),
            self.color("line"),
            radius=self.panel_radius,
            width=1,
        )
        draw.rectangle((formula_box[0], formula_box[1], formula_box[0] + 6, formula_box[3]), fill=self.color("blue"))
        label = str(slide.get("formula_label", "核心关系"))
        label_size = max(16, self.type_size("small", 16))
        label_font = self.fonts.get(TextStyle(bold=True), label_size)
        label_width = min(
            formula_box[2] - formula_box[0] - 36,
            text_width(draw, label, label_font) + 28,
        )
        label_height = max(30, int(round(label_size * 1.12)))
        label_top = formula_box[1] + 10
        label_box = (
            formula_box[0] + 18,
            label_top,
            formula_box[0] + 18 + label_width,
            label_top + label_height,
        )
        draw.rounded_rectangle(label_box, radius=3, fill=self.color("blue"))
        draw.text(
            (label_box[0] + 14, (label_box[1] + label_box[3]) // 2),
            label,
            font=label_font,
            fill=self.color("white"),
            anchor="lm",
        )

        formula_text = str(slide.get("formula", "")).strip()
        formula_latex = str(slide.get("formula_latex", "")).strip()
        condition = str(slide.get("formula_condition", "")).strip()
        condition_size = self.type_size("micro", 14)
        condition_height = 0
        if condition:
            condition_width = formula_box[2] - formula_box[0] - 56
            condition_font, condition_lines, _, _ = self.fit_text(
                draw,
                condition,
                self.theme["fonts"]["body"],
                condition_width,
                int(round(condition_size * 2.8)),
                condition_size,
                max(18, condition_size - 4),
                max_lines=2,
                line_ratio=1.05,
            )
            condition_height = int(
                round(line_height(condition_font, 1.05) * max(1, len(condition_lines)) + 12)
            )
        math_bottom = formula_box[3] - (condition_height + 12 if condition else 22)
        math_top = label_box[3] + 8
        math_box = (formula_box[0] + 28, math_top, formula_box[2] - 28, math_bottom)
        math_image = (
            render_math_image(formula_latex, self.color("blue_dark"), font_size=48)
            if formula_latex
            else None
        )
        if math_image is not None:
            inner_width = math_box[2] - math_box[0]
            inner_height = math_box[3] - math_box[1]
            scale = min(inner_width / math_image.width, inner_height / math_image.height)
            size = (max(1, round(math_image.width * scale)), max(1, round(math_image.height * scale)))
            resized = math_image.resize(size, Image.Resampling.LANCZOS)
            paste_x = math_box[0] + (inner_width - size[0]) // 2
            paste_y = math_box[1] + (inner_height - size[1]) // 2
            image.paste(resized, (paste_x, paste_y), resized)
            draw = ImageDraw.Draw(image)
        else:
            fallback = formula_text or formula_latex or "公式待补充"
            fallback_font, fallback_lines, _, fallback_fits = self.fit_text(
                draw,
                fallback,
                self.formula_fonts,
                math_box[2] - math_box[0],
                math_box[3] - math_box[1],
                46,
                18,
                max_lines=3,
                line_ratio=1.16,
            )
            if not fallback_fits:
                self.warn(str(slide.get("id", "")), "Formula text overflowed.")
            block_height = line_height(fallback_font, 1.16) * max(1, len(fallback_lines))
            draw_lines(
                draw,
                fallback_lines,
                (math_box[0], math_box[1] + (math_box[3] - math_box[1] - block_height) // 2),
                fallback_font,
                self.color("blue_dark"),
                1.16,
                align="center",
                max_width=math_box[2] - math_box[0],
            )
        if condition:
            condition_box = (
                formula_box[0] + 18,
                formula_box[3] - condition_height - 6,
                formula_box[2] - 18,
                formula_box[3] - 6,
            )
            draw.rounded_rectangle(
                condition_box,
                radius=3,
                fill=mix_color(self.color("blue_pale"), self.color("white"), 0.3),
                outline=mix_color(self.color("line"), self.color("white"), 0.15),
                width=1,
            )
            condition_fits, _ = draw_rich_block(
                draw,
                condition,
                (condition_box[0] + 10, condition_box[1] + 4),
                condition_box[2] - condition_box[0] - 20,
                condition_box[3] - condition_box[1] - 6,
                self.fonts,
                self.colors,
                condition_size,
                max(18, condition_size - 4),
                max_lines=2,
                line_ratio=1.05,
            )
            if not condition_fits:
                self.warn(str(slide.get("id", "")), "Formula condition overflowed.")

        variables_top = formula_box[3] + 18
        variables = [item for item in (slide.get("variables") or []) if isinstance(item, dict)]
        variable_title = str(slide.get("variables_title", "符号与条件"))
        variable_title_size = self.type_size("small", 18)
        title_font = self.fonts.get(TextStyle(bold=True), variable_title_size)
        title_height = max(30, int(variable_title_size * 1.6))
        draw.text((left, variables_top), variable_title, font=title_font, fill=self.color("blue_dark"))
        draw.line((left, variables_top + title_height, left_right, variables_top + title_height), fill=self.color("line"), width=2)
        if variables:
            # A symbol followed by a full sentence needs a real definition-list
            # row, not two narrow newspaper columns. Up to six entries stay in
            # one column; longer lists fall back to two balanced columns.
            columns = 2 if len(variables) > 6 else 1
            column_gap = 24
            column_width = (left_right - left - column_gap * (columns - 1)) // columns
            rows = math.ceil(len(variables) / columns)
            body_size = self.type_size("small", 16)
            body_font = self.fonts.get(TextStyle(bold=True), body_size)
            symbol_width = max(
                [text_width(draw, str(item.get("symbol", "")).strip(), body_font) for item in variables]
                or [body_size]
            )
            symbol_width = max(64, min(180, int(symbol_width) + 24))
            list_top = variables_top + title_height + 14
            available_height = max(60, bottom - list_top - 8)
            row_height = available_height // max(1, rows)
            for column in range(columns):
                chunk = variables[column * rows : (column + 1) * rows]
                if not chunk:
                    continue
                x = left + column * (column_width + column_gap)
                for row_index, item in enumerate(chunk):
                    symbol = str(item.get("symbol", "")).strip()
                    meaning = str(item.get("meaning", "")).strip()
                    y = list_top + row_index * row_height
                    symbol_top = y + max(0, (row_height - int(body_size * 1.3)) // 2)
                    draw.text((x, symbol_top), symbol, font=body_font, fill=self.color("blue_dark"))
                    font, lines, _, fits = self.fit_text(
                        draw,
                        meaning,
                        self.theme["fonts"]["body"],
                        column_width - symbol_width - 10,
                        row_height - 4,
                        body_size,
                        max(18, body_size - 8),
                        max_lines=1,
                    )
                    if not fits:
                        self.warn(str(slide.get("id", "")), f"Formula variable overflowed: {symbol}")
                    draw.text(
                        (x + symbol_width, symbol_top),
                        lines[0] if lines else meaning,
                        font=font,
                        fill=self.color("ink"),
                    )

        panels = slide.get("panels") or []
        if panels:
            panel_gap = 10
            panel_height = (bottom - body_top - panel_gap * (len(panels) - 1)) // max(1, len(panels))
            panel_width = right - right_left
            for index, panel in enumerate(panels):
                panel_top = body_top + index * (panel_height + panel_gap)
                panel_box = (right_left, panel_top, right, panel_top + panel_height)
                draw_shadowed_panel(
                    image,
                    draw,
                    panel_box,
                    self.color("white") if index % 2 == 0 else self.color("blue_wash"),
                    self.color("line"),
                    radius=self.panel_radius,
                    width=1,
                    shadow_offset=(0, 3),
                    shadow_blur=6,
                    shadow_alpha=max(10, self.panel_shadow_alpha - 6),
                )
                header_height = self.draw_module_header(
                    image,
                    draw,
                    panel_box,
                    str(panel.get("heading", "")),
                    index=f"{index + 1:02d}",
                    icon=str(panel.get("icon", "")),
                )
                content_top = panel_top + header_height + 12
                body = str(panel.get("body", "")).strip()
                bullets = panel.get("bullets") or []
                if body:
                    fits, used = draw_rich_block(
                        draw,
                        body,
                        (right_left + 14, content_top),
                        panel_width - 28,
                        max(30, panel_top + panel_height - content_top - 10),
                        self.fonts,
                        self.colors,
                        self.type_size("body", 18),
                        self.type_size("small", 13),
                        max_lines=4 if not bullets else 3,
                        line_ratio=1.2,
                    )
                    if not fits:
                        self.warn(str(slide.get("id", "")), f"Formula panel body overflowed: {panel.get('heading', '')}")
                    content_top += used + 8
                if bullets:
                    fits, _ = self.draw_compact_bullets(
                        draw,
                        bullets,
                        right_left + 14,
                        content_top,
                        panel_width - 28,
                        panel_top + panel_height - content_top - 10,
                        max_size=self.type_size("bullet", 16),
                        min_size=self.type_size("micro", 11),
                        gap=4,
                    )
                    if not fits:
                        self.warn(str(slide.get("id", "")), f"Formula panel bullets overflowed: {panel.get('heading', '')}")
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_pipeline(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        steps = slide.get("steps") or []
        if not steps:
            self.warn(str(slide.get("id", "")), "Pipeline has no steps.")
            steps = [{"title": "步骤", "body": "内容待补充"}]
        gap = 18
        step_width = (right - left - gap * (len(steps) - 1)) // len(steps)
        line_top = body_top + 34
        draw.line((left + 30, line_top, right - 30, line_top), fill=self.color("line"), width=3)
        card_top = body_top + 74
        has_footer = bool(slide.get("footer_row"))
        has_visual = any(str(step.get("visual", "")).strip() for step in steps)
        default_card_height = 500 if has_visual else 460
        available_height = bottom - card_top - (106 if has_footer else 24)
        card_height = min(int(slide.get("card_height", default_card_height)), available_height)
        for index, step in enumerate(steps):
            x = left + index * (step_width + gap)
            center = x + step_width // 2
            draw.ellipse((center - 29, line_top - 29, center + 29, line_top + 29), fill=self.color("blue"), outline=self.color("white"), width=4)
            number_font = self.fonts.get(TextStyle(bold=True), 24)
            number_text = f"{index + 1}"
            num_bbox = draw.textbbox((0, 0), number_text, font=number_font)
            draw.text((center - (num_bbox[2] - num_bbox[0]) // 2, line_top - 15), number_text, font=number_font, fill=self.color("white"))
            box = (x, card_top, x + step_width, card_top + card_height)
            draw_shadowed_panel(
                image,
                draw,
                box,
                self.color("white"),
                self.color("line"),
                radius=self.panel_radius,
                width=1,
                shadow_offset=(0, 6),
                shadow_blur=8,
                shadow_alpha=self.panel_shadow_alpha,
            )
            draw_icon_badge(image, draw, (x + 42, card_top + 42), 54, str(step.get("icon", "route")), self.colors)
            step_font = self.fonts.get(TextStyle(bold=True), 15)
            draw.text((x + 74, card_top + 22), f"STEP {index + 1:02d}", font=step_font, fill=self.color("muted"))
            title_font, title_lines, _, title_fits = self.fit_text(
                draw,
                str(step.get("title", "")),
                self.theme["fonts"]["title"],
                step_width - 32,
                50,
                27,
                18,
                max_lines=2,
            )
            if not title_fits:
                self.warn(str(slide.get("id", "")), f"Pipeline title overflowed: {step.get('title', '')}")
            draw_lines(draw, title_lines, (x + 18, card_top + 78), title_font, self.color("blue_dark"), 1.16)
            body_font, body_lines, _, body_fits = self.fit_text(
                draw,
                str(step.get("body", "")),
                self.theme["fonts"]["body"],
                step_width - 36,
                102,
                22,
                14,
                max_lines=4,
                line_ratio=1.3,
            )
            if not body_fits:
                self.warn(str(slide.get("id", "")), f"Pipeline body overflowed: {step.get('title', '')}")
            draw_lines(draw, body_lines, (x + 18, card_top + 146), body_font, self.color("ink"), 1.3)
            bullets = step.get("bullets") or []
            visual = str(step.get("visual", "")).strip() if step.get("visual_semantic") else ""
            output = str(step.get("output", "")).strip()
            output_height = 52 if output else 0
            output_top = card_top + card_height - output_height - 14 if output else card_top + card_height - 14
            bullets_bottom = output_top - 10
            if visual:
                visual_height = 82
                visual_top = output_top - visual_height - 8
                draw_mini_visual(
                    draw,
                    (x + 18, visual_top, x + step_width - 18, output_top - 8),
                    visual,
                    self.colors,
                    self.fonts,
                )
                bullets_bottom = visual_top - 6
            if bullets:
                body_height = line_height(body_font, 1.3) * max(1, len(body_lines))
                bullets_top = min(card_top + 258, card_top + 146 + body_height + 18)
                fits, _ = self.draw_compact_bullets(
                    draw,
                    bullets,
                    x + 18,
                    bullets_top,
                    step_width - 36,
                    bullets_bottom - bullets_top,
                    max_size=19,
                    min_size=13,
                    gap=7,
                )
                if not fits:
                    self.warn(str(slide.get("id", "")), f"Pipeline bullets overflowed: {step.get('title', '')}")
            if output:
                self.draw_output_strip(
                    draw,
                    (x + 16, output_top, x + step_width - 16, card_top + card_height - 14),
                    "输出",
                    output,
                    str(slide.get("id", "")),
                )
            if index < len(steps) - 1:
                draw_chevrons(draw, x + step_width + 4, line_top - 9, 18, self.color("blue_mid"), count=1, width=3, gap=8)
        footer_row = slide.get("footer_row") or []
        if footer_row:
            footer_top = card_top + card_height + 16
            footer_height = min(84, bottom - footer_top)
            if footer_height > 20:
                gap2 = 12
                item_width = (right - left - gap2 * (len(footer_row) - 1)) // len(footer_row)
                for index, item in enumerate(footer_row):
                    x = left + index * (item_width + gap2)
                    draw_shadowed_panel(
                        image,
                        draw,
                        (x, footer_top, x + item_width, footer_top + footer_height),
                        self.color("blue_wash"),
                        self.color("line"),
                        radius=6,
                        width=1,
                        shadow_offset=(0, 3),
                        shadow_blur=6,
                        shadow_alpha=16,
                    )
                    draw_icon(
                        draw,
                        str(item.get("icon", "alert")),
                        (x + 12, footer_top + 12, x + 42, footer_top + 42),
                        self.color("blue"),
                        self.color("blue_wash"),
                        stroke=2,
                    )
                    label_font = self.fonts.get(TextStyle(bold=True), 17)
                    draw.text((x + 50, footer_top + 10), str(item.get("title", "")), font=label_font, fill=self.color("blue_dark"))
                    value_font = self.fonts.get(TextStyle(), 14)
                    draw.text((x + 50, footer_top + 34), str(item.get("body", "")), font=value_font, fill=self.color("muted"))
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_compare(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        columns = slide.get("columns") or []
        rows = slide.get("rows") or []
        if not columns:
            self.warn(str(slide.get("id", "")), "Compare table has no columns.")
            return image
        table = self.theme["table"]
        table_header_size = self.type_size("table_header", 25)
        table_cell_size = self.type_size("table_cell", 24)
        header_height = max(
            int(table.get("header_height", 66)),
            int(table_header_size * 1.7),
        )
        row_height = min(
            max(112, int(table_cell_size * 2.1)),
            max(80, (bottom - body_top - header_height - 4) // max(1, len(rows))),
        )
        weights = slide.get("column_weights") or ([1.4] + [1.0] * (len(columns) - 1))
        if len(weights) != len(columns):
            weights = [1.4] + [1.0] * (len(columns) - 1)
        total_weight = sum(weights)
        widths = [round((right - left) * weight / total_weight) for weight in weights]
        widths[-1] = right - left - sum(widths[:-1])
        table_height = header_height + row_height * len(rows)
        header_top = body_top + max(0, (bottom - body_top - table_height - 18) // 2)
        draw_shadowed_panel(
            image,
            draw,
            (left, header_top, right, header_top + table_height),
            self.color("white"),
            self.color("line"),
            radius=self.panel_radius,
            width=1,
            shadow_offset=(0, 6),
            shadow_blur=8,
            shadow_alpha=self.panel_shadow_alpha,
        )
        current_left = left
        for index, column in enumerate(columns):
            cell_box = (current_left, header_top, current_left + widths[index], header_top + header_height)
            draw_gradient_panel(image, draw, cell_box, self.color("blue_mid"), self.color("blue_dark"), None, radius=0, width=0)
            heading_font, heading_lines, _, heading_fits = self.fit_text(
                draw,
                column,
                self.theme["fonts"]["title"],
                widths[index] - 20,
                header_height - 12,
                table_header_size,
                16,
                max_lines=1,
            )
            if not heading_fits:
                self.warn(str(slide.get("id", "")), f"Table heading overflowed: {column}")
            heading_width = text_width(draw, heading_lines[0] if heading_lines else column, heading_font)
            draw.text((current_left + (widths[index] - heading_width) // 2, header_top + 17), heading_lines[0] if heading_lines else column, font=heading_font, fill=self.color("white"))
            if index:
                draw.line((current_left, header_top, current_left, header_top + table_height), fill=self.color("white"), width=2)
            current_left += widths[index]
        highlight_rows = set(int(value) for value in slide.get("highlight_rows", []))
        for row_index, row in enumerate(rows):
            row_top = header_top + header_height + row_index * row_height
            fill = self.color("blue_wash") if row_index % 2 == 1 or row_index in highlight_rows else self.color("white")
            current_left = left
            for column_index, value in enumerate(row[: len(columns)]):
                cell_box = (current_left, row_top, current_left + widths[column_index], row_top + row_height)
                draw.rectangle(cell_box, fill=fill)
                if column_index == 0:
                    draw.rectangle((current_left, row_top, current_left + 5, row_top + row_height), fill=self.color("blue") if row_index not in highlight_rows else self.color("red"))
                if column_index:
                    draw.line((current_left, row_top, current_left, row_top + row_height), fill=self.color("line"), width=1)
                font_paths = self.theme["fonts"]["title"] if column_index == 0 else self.theme["fonts"]["body"]
                value_font, value_lines, _, value_fits = self.fit_text(
                    draw,
                    str(value),
                    font_paths,
                    widths[column_index] - 26,
                    row_height - 16,
                table_cell_size,
                    14,
                    max_lines=3,
                    line_ratio=1.2,
                )
                if not value_fits:
                    self.warn(str(slide.get("id", "")), f"Table cell overflowed: {value}")
                block_height = line_height(value_font, 1.2) * len(value_lines)
                draw_lines(
                    draw,
                    value_lines,
                    (current_left + 14, row_top + (row_height - block_height) // 2),
                    value_font,
                    self.color("blue_dark") if column_index == 0 else self.color("ink"),
                    1.2,
                )
                current_left += widths[column_index]
            draw.line((left, row_top, right, row_top), fill=self.color("line"), width=1)
        notes = slide.get("notes") or []
        table_bottom = header_top + table_height
        if notes and table_bottom + 16 < int(self.layout["takeaway_top"]) - 26:
            note_top = table_bottom + 16
            note_height = min(64, int(self.layout["takeaway_top"]) - note_top - 20)
            columns_notes = 2
            note_width = (right - left - 14) // 2
            half = math.ceil(len(notes) / columns_notes)
            note_size = self.type_size("micro", 17)
            self.draw_compact_bullets(draw, notes[:half], left, note_top, note_width, note_height, max_size=note_size, min_size=12, gap=4)
            self.draw_compact_bullets(draw, notes[half:], left + note_width + 14, note_top, note_width, note_height, max_size=note_size, min_size=12, gap=4)
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_metric_card(
        self,
        image: Image.Image,
        draw: ImageDraw.ImageDraw,
        box: tuple[int, int, int, int],
        metric: dict,
        index: int,
    ) -> None:
        left, top, right, bottom = box
        draw_shadowed_panel(
            image,
            draw,
            box,
            self.color("blue_wash") if index == 0 else self.color("white"),
            self.color("line"),
            radius=10,
            width=2,
            shadow_offset=(0, 6),
            shadow_blur=10,
            shadow_alpha=22,
        )
        draw.rectangle((left + 2, top + 2, right - 2, top + 7), fill=self.color("blue") if index == 0 else self.color("blue_mid"))
        icon = str(metric.get("icon", "chart"))
        draw_icon_badge(image, draw, (right - 42, top + 44), 48, icon, self.colors)
        value_font, value_lines, _, value_fits = self.fit_text(
            draw,
            str(metric.get("value", "")),
            self.theme["fonts"]["title"],
            right - left - 98,
            80,
            64,
            28,
            max_lines=1,
        )
        if not value_fits:
            self.warn("metric", f"Metric value overflowed: {metric.get('value', '')}")
        value_y = top + 18
        draw_lines(draw, value_lines, (left + 20, value_y), value_font, self.color("blue"), 1.05)
        label_font, label_lines, _, label_fits = self.fit_text(
            draw,
            str(metric.get("label", "")),
            self.theme["fonts"]["title"],
            right - left - 96,
            58,
            max(24, self.type_size("metric_label", 24)),
            max(18, int(self.type_size("metric_label", 24) * 0.8)),
            max_lines=1,
        )
        if not label_fits:
            self.warn("metric", f"Metric label overflowed: {metric.get('label', '')}")
        label_y = value_y + line_height(value_font, 1.05) + 4
        draw_lines(draw, label_lines, (left + 20, label_y), label_font, self.color("ink"), 1.15)
        note_size = max(16, self.type_size("micro", 14))
        note_font, note_lines, _, note_fits = self.fit_text(
            draw,
            str(metric.get("note", "")),
            self.theme["fonts"]["body"],
            right - left - 36,
            int(note_size * 1.45),
            note_size,
            max(16, int(note_size * 0.8)),
            max_lines=1,
        )
        if not note_fits:
            self.warn("metric", f"Metric note overflowed: {metric.get('note', '')}")
        note_y = bottom - line_height(note_font, 1.15) - 8
        draw_lines(draw, note_lines, (left + 20, note_y), note_font, self.color("muted"), 1.15)

    def draw_condition_grid(
        self,
        draw: ImageDraw.ImageDraw,
        conditions: list,
        box: tuple[int, int, int, int],
        columns: int = 2,
    ) -> None:
        left, top, right, bottom = box
        if not conditions:
            return
        defense = str(getattr(self, "profile", "")).strip().lower() == "defense"
        rows = math.ceil(len(conditions) / columns)
        gap = 16 if defense else 12
        cell_width = (right - left - gap * (columns - 1)) // columns
        available = bottom - top
        if defense:
            # Defense type cannot go below ~18 pt. Give each label/value pair
            # its own band instead of stacking them into a 58 px row.
            row_height = min(104, max(72, available // max(1, rows)))
        else:
            row_height = min(58, available // max(1, rows))
        label_font = self.fonts.get(TextStyle(), 38 if defense else 15)
        value_font = self.fonts.get(TextStyle(bold=True), 40 if defense else 18)
        for index, item in enumerate(conditions):
            row = index // columns
            column = index % columns
            x = left + column * (cell_width + gap)
            y = top + row * row_height
            if isinstance(item, dict):
                label = str(item.get("label", ""))
                value = str(item.get("value", ""))
            else:
                text = str(item)
                if "：" in text:
                    label, value = text.split("：", 1)
                elif ":" in text:
                    label, value = text.split(":", 1)
                else:
                    label, value = "", text
            if label:
                draw.text((x, y + 2), label, font=label_font, fill=self.color("muted"))
            value_width = cell_width - (0 if not label else 0)
            value_rendered, value_lines, _, value_fits = self.fit_text(
                draw,
                value,
                self.theme["fonts"]["title"],
                value_width,
                row_height - (2 if not label else line_height(label_font, 1.0) + 2),
                40 if defense else 18,
                14 if defense else 12,
                max_lines=1 if defense else 2,
            )
            if not value_fits:
                self.warn("conditions", f"Condition value overflowed: {value}")
            value_top = y + (line_height(label_font, 1.0) + 3 if label else 3)
            draw_lines(draw, value_lines, (x, value_top), value_rendered, self.color("blue_dark"), 1.05)
            if not defense:
                draw.line((x, y + row_height - 6, x + cell_width, y + row_height - 6), fill=self.color("line"), width=1)

    def draw_metrics(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        metrics = slide.get("metrics") or []
        if not metrics:
            self.warn(str(slide.get("id", "")), "Metrics slide has no metrics.")
            metrics = [{"value": "—", "label": "指标", "note": ""}]
        gap = 16
        card_width = (right - left - gap * (len(metrics) - 1)) // len(metrics)
        value_size = min(64, self.type_size("metric_value", 64))
        label_size = max(24, self.type_size("metric_label", 24))
        note_size = max(16, self.type_size("micro", 14))
        card_height = max(
            196,
            int(value_size * 1.05 + label_size * 1.3 + note_size * 1.35 + 42),
        )
        for index, metric in enumerate(metrics):
            x = left + index * (card_width + gap)
            self.draw_metric_card(image, draw, (x, body_top, x + card_width, body_top + card_height), metric, index)
        lower_top = body_top + card_height + 40
        lower_bottom = bottom
        image_path = resolve_path(slide.get("image"), self.deck_path.parent)
        if image_path is not None or slide.get("chart"):
            image_width = round((right - left) * 0.58)
            draw_figure_card(
                image,
                draw,
                (left, lower_top, left + image_width, lower_bottom),
                image_path,
                str(slide.get("image_label", "测量结果")),
                str(slide.get("image_caption", "")),
                self.colors,
                self.fonts,
                chart=slide.get("chart"),
                badge=str(slide.get("image_badge", "")),
                scale=1.8 if self.profile == "defense" else 1.0,
            )
            draw = ImageDraw.Draw(image)
            side_box = (left + image_width + 18, lower_top, right, lower_bottom)
            draw_shadowed_panel(image, draw, side_box, self.color("white"), self.color("line"), radius=9, width=2, shadow_offset=(0, 5), shadow_blur=9, shadow_alpha=20)
            header_height = self.draw_module_header(
                image,
                draw,
                side_box,
                str(slide.get("conditions_title", "实验条件")),
                icon="alert",
            )
            content_top = lower_top + header_height + 12
            conditions = slide.get("conditions") or slide.get("bullets") or []
            note = str(slide.get("conditions_note", "")).strip()
            note_height = 70 if note else 0
            if conditions and isinstance(conditions[0], dict):
                self.draw_condition_grid(
                    draw,
                    conditions,
                    (left + image_width + 34, content_top, right - 18, lower_bottom - 16 - note_height),
                    columns=2,
                )
            else:
                fits, _ = self.draw_compact_bullets(
                    draw,
                    conditions,
                    left + image_width + 18 + 16,
                    content_top,
                    (right - left - image_width - 18) - 32,
                    lower_bottom - content_top - 16 - note_height,
                    max_size=18,
                    min_size=13,
                    gap=6,
                )
                if not fits:
                    self.warn(str(slide.get("id", "")), "Conditions list overflowed.")
            if note:
                note_top = lower_bottom - note_height
                note_box = (left + image_width + 30, note_top, right - 18, lower_bottom - 16)
                draw.rounded_rectangle(note_box, radius=6, fill=self.color("blue_wash"), outline=self.color("line"), width=1)
                draw.rectangle((note_box[0], note_box[1], note_box[0] + 5, note_box[3]), fill=self.color("red"))
                draw_icon(
                    draw,
                    "alert",
                    (note_box[0] + 14, note_box[1] + 13, note_box[0] + 36, note_box[1] + 35),
                    self.color("red"),
                    self.color("blue_wash"),
                    stroke=2,
                )
                draw_rich_block(
                    draw,
                    note,
                    (note_box[0] + 44, note_box[1] + 10),
                    note_box[2] - note_box[0] - 56,
                    note_box[3] - note_box[1] - 16,
                    self.fonts,
                    self.colors,
                    16,
                    11,
                    max_lines=3,
                    line_ratio=1.22,
                )
        else:
            bullets = slide.get("bullets") or []
            if bullets:
                columns = 2
                available = right - left - 18
                width = available // 2
                half = math.ceil(len(bullets) / columns)
                self.draw_compact_bullets(draw, bullets[:half], left, lower_top, width, lower_bottom - lower_top, max_size=21, min_size=15, gap=8)
                self.draw_compact_bullets(draw, bullets[half:], left + width + 18, lower_top, width, lower_bottom - lower_top, max_size=21, min_size=15, gap=8)
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_footnote(self, image: Image.Image, text: str) -> None:
        text = str(text or "").strip()
        source = source_line_for_slide(self.deck, getattr(self, "current_slide", {}))
        if source:
            text = f"{text}　｜　{source}" if text else source
        if not text:
            return
        draw = ImageDraw.Draw(image)
        left = int(self.layout["content"][0])
        right = int(self.layout["content"][2])
        size = self.type_size("footnote", 20)
        text_y = int(self.layout["takeaway_top"]) + int(self.layout["takeaway_height"]) + 6
        available_height = max(24, int(self.layout["footer_y"]) - text_y - 6)
        height = min(max(24, int(round(size * 1.18))), available_height)
        font, lines, _, fits = self.fit_text(
            draw,
            text,
            self.theme["fonts"]["body"],
            right - left,
            height,
            size,
            max(14, size - 4),
            max_lines=1,
            line_ratio=1.0,
        )
        if not fits:
            self.warn("footnote", f"Footnote overflowed: {text}")
        ascent, descent = font.getmetrics()
        text_y += max(0, (height - (ascent + descent)) // 2)
        draw.text(
            (left, text_y),
            lines[0] if lines else text,
            font=font,
            fill=self.color("muted"),
        )

    def draw_quote(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        quote = str(slide.get("quote", "")).strip()
        attribution = str(slide.get("attribution", "")).strip()
        commentary = str(slide.get("commentary", "")).strip()
        quote_note = str(slide.get("quote_note", "")).strip()
        variant = str(slide.get("variant", "centered")).strip().lower()
        if not quote:
            self.warn(str(slide.get("id", "")), "Quote layout has no quotation.")
            quote = "引文待补充"
        if variant == "split":
            quote_box = (left, body_top, left + round((right - left) * 0.64), bottom)
            rail_box = (quote_box[2] + 20, body_top, right, bottom)
        else:
            quote_height = round((bottom - body_top) * 0.66)
            quote_box = (left, body_top, right, body_top + quote_height)
            rail_box = (left, quote_box[3] + 16, right, bottom)
        draw_shadowed_panel(
            image,
            draw,
            quote_box,
            self.color("white"),
            self.color("line"),
            radius=8,
            width=2,
            shadow_offset=(0, 4),
            shadow_blur=8,
            shadow_alpha=18,
        )
        draw.rectangle(
            (quote_box[0], quote_box[1], quote_box[0] + 8, quote_box[3]),
            fill=self.color("blue"),
        )
        mark_font = self.fonts.get(TextStyle(), 92)
        draw.text((quote_box[0] + 26, quote_box[1] + 2), "“", font=mark_font, fill=self.color("blue_pale"))
        text_left = quote_box[0] + 82
        quote_width = quote_box[2] - text_left - 42
        text_top = quote_box[1] + 46
        text_bottom = quote_box[3] - 76
        quote_font, quote_lines, _, quote_fits = self.fit_text(
            draw,
            quote,
            self.theme["fonts"]["serif"],
            quote_width,
            max(40, text_bottom - text_top),
            self.type_size("quote_text", 42),
            24,
            max_lines=7,
            line_ratio=1.45,
        )
        if not quote_fits:
            self.warn(str(slide.get("id", "")), "Quotation overflowed.")
        draw_lines(draw, quote_lines, (text_left, text_top), quote_font, self.color("ink"), 1.45)
        if attribution:
            attr_text = f"— {attribution}"
            attr_font, attr_lines, _, attr_fits = self.fit_text(
                draw,
                attr_text,
                self.theme["fonts"]["body"],
                quote_width,
                32,
                self.type_size("small", 20),
                14,
                max_lines=1,
            )
            if not attr_fits:
                self.warn(str(slide.get("id", "")), "Quote attribution overflowed.")
            attr_width = text_width(draw, attr_lines[0] if attr_lines else attr_text, attr_font)
            draw.text(
                (quote_box[2] - 40 - attr_width, quote_box[3] - 46),
                attr_lines[0] if attr_lines else attr_text,
                font=attr_font,
                fill=self.color("muted"),
            )
        draw_shadowed_panel(
            image,
            draw,
            rail_box,
            self.color("blue_wash"),
            self.color("line"),
            radius=8,
            width=2,
            shadow_offset=(0, 4),
            shadow_blur=8,
            shadow_alpha=16,
        )
        header_height = self.draw_module_header(
            image,
            draw,
            rail_box,
            "解读",
            icon="target",
        )
        rail_left = rail_box[0] + 18
        rail_top = rail_box[1] + header_height + 16
        rail_width = rail_box[2] - rail_left - 18
        if commentary:
            fits, used = draw_rich_block(
                draw,
                commentary,
                (rail_left, rail_top),
                rail_width,
                max(40, rail_box[3] - rail_top - (54 if quote_note else 18)),
                self.fonts,
                self.colors,
                self.type_size("body", 22),
                self.type_size("small", 15),
                max_lines=6,
                line_ratio=1.28,
            )
            if not fits:
                self.warn(str(slide.get("id", "")), "Quote commentary overflowed.")
            rail_top += used + 10
        if quote_note:
            note_font, note_lines, _, note_fits = self.fit_text(
                draw,
                quote_note,
                self.theme["fonts"]["body"],
                rail_width,
                max(24, rail_box[3] - rail_top - 12),
                self.type_size("micro", 14),
                11,
                max_lines=2,
            )
            if not note_fits:
                self.warn(str(slide.get("id", "")), "Quote note overflowed.")
            draw_lines(draw, note_lines, (rail_left, rail_top), note_font, self.color("muted"), 1.2)
        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_closing(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        points = list(slide.get("points") or [])
        has_bottom_cards = bool(
            slide.get("boundary")
            or slide.get("boundary_points")
            or slide.get("next")
            or slide.get("next_points")
        )
        if not points and not has_bottom_cards:
            points = ["结论待补充"]
        row_height = 88 if has_bottom_cards else 104
        gap = 14
        points_bottom = body_top + len(points) * (row_height + gap) - gap if points else body_top
        for index, point in enumerate(points):
            row_top = body_top + index * (row_height + gap)
            draw_shadowed_panel(
                image,
                draw,
                (left, row_top, right, row_top + row_height),
                self.color("white") if index else self.color("blue_wash"),
                self.color("line"),
                radius=8,
                width=2,
                shadow_offset=(0, 4),
                shadow_blur=8,
                shadow_alpha=18,
            )
            draw.rectangle((left, row_top, left + 7, row_top + row_height), fill=self.color("blue") if index else self.color("red"))
            number_font = self.fonts.get(TextStyle(bold=True), 44)
            number_color = mix_color(self.color("blue_mid"), self.color("white"), 0.35)
            draw.text((left + 24, row_top + 24), f"{index + 1:02d}", font=number_font, fill=number_color)
            fits, _ = draw_rich_block(
                draw,
                str(point),
                (left + 108, row_top + 22),
                right - left - 136,
                row_height - 34,
                self.fonts,
                self.colors,
                self.type_size("lead", 27),
                17,
                max_lines=2,
                line_ratio=1.2,
            )
            if not fits:
                self.warn(str(slide.get("id", "")), f"Closing point overflowed: {point}")
        cards = []
        if slide.get("boundary") or slide.get("boundary_points"):
            cards.append(
                (
                    "边界",
                    "alert",
                    str(slide.get("boundary", "")),
                    slide.get("boundary_points") or [],
                    self.color("blue_wash"),
                )
            )
        if slide.get("next") or slide.get("next_points"):
            cards.append(
                (
                    "下一步",
                    "arrow",
                    str(slide.get("next", "")),
                    slide.get("next_points") or [],
                    self.color("white"),
                )
            )
        cards_top = points_bottom + (16 if points else 0)
        # Give the closing cards the whole band between the lead and the
        # closing line. A fixed 100 px reserve for the closing line left the
        # second boundary bullet outside the card while the area below the
        # closing line stayed empty. Stop the cards above the footer and keep
        # the closing line in the paper margin.
        footer_y = int(self.layout.get("footer_y", self.height - 80))
        cards_bottom = min(bottom - 26, footer_y - 132)
        if cards and cards_bottom - cards_top > 80:
            gap_cards = 16
            card_width = (right - left - gap_cards * (len(cards) - 1)) // len(cards)
            for index, (title, icon, body, points, fill) in enumerate(cards):
                x = left + index * (card_width + gap_cards)
                box = (x, cards_top, x + card_width, cards_bottom)
                draw_shadowed_panel(image, draw, box, fill, self.color("line"), radius=8, width=2, shadow_offset=(0, 4), shadow_blur=8, shadow_alpha=18)
                draw_icon(
                    draw,
                    icon,
                    (x + card_width - 50, cards_top + 18, x + card_width - 18, cards_top + 50),
                    self.color("red") if index == 0 else self.color("blue"),
                    fill,
                    stroke=2,
                )
                title_font = self.fonts.get(
                    TextStyle(bold=True),
                    max(52, int(self.type_size("block_title", 48) * 1.08)),
                )
                draw.text((x + 18, cards_top + 20), title, font=title_font, fill=self.color("blue_dark"))
                body_top = cards_top + max(58, int(title_font.size * 1.55))
                if body:
                    _, used = draw_rich_block(
                        draw,
                        body,
                        (x + 18, body_top),
                        card_width - 36,
                        max(34, cards_bottom - body_top - 16),
                        self.fonts,
                        self.colors,
                        self.type_size("small", 16),
                        max(18, int(self.type_size("small", 16) * 0.8)),
                        max_lines=2,
                        line_ratio=1.22,
                    )
                    body_top += used + 4
                if points:
                    fits, _ = self.draw_compact_bullets(
                        draw,
                        points,
                        x + 18,
                        body_top,
                        card_width - 36,
                        max(24, cards_bottom - body_top - 12),
                        max_size=max(32, int(self.type_size("small", 16) * 0.95)),
                        min_size=26,
                        gap=4,
                    )
                    if not fits:
                        self.warn(str(slide.get("id", "")), f"Closing card bullets overflowed: {title}")
        closing = str(slide.get("closing", "谢谢各位老师")).strip()
        closing_font, closing_lines, _, closing_fits = self.fit_text(
            draw,
            closing,
            self.theme["fonts"]["title"],
            right - left - 80,
            64,
            48,
            28,
            max_lines=1,
        )
        if not closing_fits:
            self.warn(str(slide.get("id", "")), "Closing line overflowed.")
        closing_y = cards_bottom + 30 if cards and cards_bottom - cards_top > 80 else bottom - 62
        closing_y = min(closing_y, footer_y - 96)
        draw.line(
            (left, closing_y - 24, right, closing_y - 24),
            fill=mix_color(self.color("line"), self.color("white"), 0.2),
            width=1,
        )
        draw_lines(draw, closing_lines, (left + 40, closing_y), closing_font, self.color("blue"), align="center", max_width=right - left - 80)
        return image

    def draw_timeline(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        phases = slide.get("phases") or []
        if not 3 <= len(phases) <= 5:
            self.warn(str(slide.get("id", "")), "Timeline layout expects three to five phases.")
        if not phases:
            phases = [{"period": "阶段", "title": "内容待补充"}]

        axis_y = body_top + 34
        axis_left = left + 30
        axis_right = right - 30
        draw.line((axis_left, axis_y, axis_right, axis_y), fill=self.color("line"), width=3)
        draw.polygon(
            [
                (axis_right - 2, axis_y),
                (axis_right - 20, axis_y - 9),
                (axis_right - 20, axis_y + 9),
            ],
            fill=self.color("blue_mid"),
        )
        period_note = str(slide.get("period_note", "")).strip()
        if period_note:
            note_font = self.fonts.get(TextStyle(), 16)
            note_width = text_width(draw, period_note, note_font)
            draw.text(
                (axis_right - note_width, axis_y + 14),
                period_note,
                font=note_font,
                fill=self.color("muted"),
            )

        card_top = body_top + 86
        timeline_note = str(slide.get("timeline_note", "")).strip()
        note_height = 56 if timeline_note else 0
        max_card_bottom = bottom - (note_height + 18 if timeline_note else 0)
        card_bottom = min(max_card_bottom, card_top + int(slide.get("card_height", 470)))
        gap = 18
        card_width = (right - left - gap * (len(phases) - 1)) // len(phases)
        for index, phase in enumerate(phases):
            x = left + index * (card_width + gap)
            card = (x, card_top, x + card_width, card_bottom)
            center_x = x + card_width // 2
            draw.line(
                (center_x, axis_y + 14, center_x, card_top),
                fill=mix_color(self.color("blue_mid"), self.color("line"), 0.34),
                width=2,
            )
            draw.ellipse(
                (center_x - 15, axis_y - 15, center_x + 15, axis_y + 15),
                fill=self.color("white"),
                outline=self.color("blue"),
                width=3,
            )
            draw.ellipse(
                (center_x - 5, axis_y - 5, center_x + 5, axis_y + 5),
                fill=self.color("blue"),
            )

            draw_shadowed_panel(
                image,
                draw,
                card,
                self.color("white"),
                self.color("line"),
                radius=self.panel_radius,
                width=1,
                shadow_offset=(0, 5),
                shadow_blur=8,
                shadow_alpha=self.panel_shadow_alpha,
            )
            ghost_font = self.fonts.get(TextStyle(bold=True), 66)
            ghost_text = f"{index + 1:02d}"
            ghost_width = text_width(draw, ghost_text, ghost_font)
            draw.text(
                (card[2] - ghost_width - 18, card_top + 16),
                ghost_text,
                font=ghost_font,
                fill=mix_color(self.color("blue_pale"), self.color("white"), 0.42),
            )

            period = str(phase.get("period", "")).strip()
            period_font, period_lines, _, period_fits = self.fit_text(
                draw,
                period,
                self.theme["fonts"]["title"],
                card_width - 54,
                34,
                19,
                13,
                max_lines=1,
                line_ratio=1.1,
            )
            if not period_fits:
                self.warn(str(slide.get("id", "")), f"Timeline period overflowed: {period}")
            period_text = period_lines[0] if period_lines else period
            pill_width = min(card_width - 36, max(92, text_width(draw, period_text, period_font) + 28))
            pill_box = (x + 18, card_top + 18, x + 18 + pill_width, card_top + 52)
            draw.rounded_rectangle(pill_box, radius=4, fill=self.color("blue"))
            draw.text(
                (pill_box[0] + 14, pill_box[1] + 5),
                period_text,
                font=period_font,
                fill=self.color("white"),
            )

            title_y = card_top + 72
            title_font, title_lines, _, title_fits = self.fit_text(
                draw,
                str(phase.get("title", "")),
                self.theme["fonts"]["title"],
                card_width - 36,
                80,
                27,
                18,
                max_lines=2,
                line_ratio=1.16,
            )
            if not title_fits:
                self.warn(str(slide.get("id", "")), f"Timeline title overflowed: {phase.get('title', '')}")
            draw_lines(
                draw,
                title_lines,
                (x + 18, title_y),
                title_font,
                self.color("blue_dark"),
                1.16,
            )
            cursor = title_y + line_height(title_font, 1.16) * max(1, len(title_lines)) + 8

            body = str(phase.get("body", "")).strip()
            if body:
                fits, used = draw_rich_block(
                    draw,
                    body,
                    (x + 18, cursor),
                    card_width - 36,
                    112,
                    self.fonts,
                    self.colors,
                    18,
                    13,
                    max_lines=3,
                    line_ratio=1.25,
                )
                if not fits:
                    self.warn(str(slide.get("id", "")), f"Timeline body overflowed: {phase.get('title', '')}")
                cursor += min(used, 112) + 8

            milestone = str(phase.get("milestone", "")).strip()
            milestone_height = 58 if milestone else 0
            milestone_top = card_bottom - milestone_height - 12
            bullets_height = max(24, milestone_top - cursor - 8) if milestone else max(24, card_bottom - cursor - 18)
            bullets = phase.get("bullets") or []
            if bullets:
                fits, _ = self.draw_compact_bullets(
                    draw,
                    bullets,
                    x + 18,
                    cursor,
                    card_width - 36,
                    bullets_height,
                    max_size=18,
                    min_size=12,
                    gap=6,
                )
                if not fits:
                    self.warn(str(slide.get("id", "")), f"Timeline bullets overflowed: {phase.get('title', '')}")

            if milestone:
                milestone_box = (x + 14, milestone_top, card[2] - 14, card_bottom - 12)
                draw.rounded_rectangle(
                    milestone_box,
                    radius=4,
                    fill=self.color("blue_wash"),
                    outline=mix_color(self.color("line"), self.color("white"), 0.18),
                    width=1,
                )
                badge_width = 74
                draw.rounded_rectangle(
                    (milestone_box[0], milestone_box[1], milestone_box[0] + badge_width, milestone_box[3]),
                    radius=4,
                    fill=self.color("blue"),
                )
                badge_font = self.fonts.get(TextStyle(bold=True), 14)
                draw.text(
                    (milestone_box[0] + 12, milestone_box[1] + 18),
                    "里程碑",
                    font=badge_font,
                    fill=self.color("white"),
                )
                milestone_font, milestone_lines, _, milestone_fits = self.fit_text(
                    draw,
                    milestone,
                    self.theme["fonts"]["body"],
                    milestone_box[2] - milestone_box[0] - badge_width - 24,
                    milestone_box[3] - milestone_box[1] - 12,
                    16,
                    11,
                    max_lines=2,
                    line_ratio=1.15,
                )
                if not milestone_fits:
                    self.warn(str(slide.get("id", "")), f"Timeline milestone overflowed: {milestone}")
                draw_lines(
                    draw,
                    milestone_lines,
                    (milestone_box[0] + badge_width + 10, milestone_box[1] + 10),
                    milestone_font,
                    self.color("blue_dark"),
                    1.15,
                )

        if timeline_note and card_bottom + 18 < bottom:
            note_box = (left + 30, card_bottom + 18, right - 30, min(bottom, card_bottom + 18 + note_height))
            draw.rounded_rectangle(
                note_box,
                radius=self.panel_radius,
                fill=self.color("blue_wash"),
                outline=mix_color(self.color("line"), self.color("white"), 0.18),
                width=1,
            )
            draw_icon_badge(
                image,
                draw,
                (note_box[0] + 30, (note_box[1] + note_box[3]) // 2),
                38,
                "alert",
                self.colors,
            )
            note_font, note_lines, _, note_fits = self.fit_text(
                draw,
                timeline_note,
                self.theme["fonts"]["body"],
                note_box[2] - note_box[0] - 82,
                note_box[3] - note_box[1] - 14,
                18,
                13,
                max_lines=1,
                line_ratio=1.15,
            )
            if not note_fits:
                self.warn(str(slide.get("id", "")), "Timeline note overflowed.")
            draw.text(
                (note_box[0] + 60, note_box[1] + (note_box[3] - note_box[1] - line_height(note_font)) // 2 - 2),
                note_lines[0] if note_lines else timeline_note,
                font=note_font,
                fill=self.color("blue_dark"),
            )

        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_references(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        references = slide.get("references") or []
        if not 4 <= len(references) <= 10:
            self.warn(str(slide.get("id", "")), "References layout expects four to ten entries.")
        if not references:
            references = [{"authors": "", "title": "参考文献待补充", "venue": "", "doi": ""}]

        column_gap = 44
        column_width = (right - left - column_gap) // 2
        split = math.ceil(len(references) / 2)
        columns = [(references[:split], left), (references[split:], left + column_width + column_gap)]
        divider_x = left + column_width + column_gap // 2
        draw.line((divider_x, body_top + 6, divider_x, bottom - 6), fill=mix_color(self.color("line"), self.color("white"), 0.25), width=1)
        reference_meta_size = self.type_size("reference_meta", 16)
        reference_title_size = self.type_size("reference_title", 22)

        for column_index, (items, x) in enumerate(columns):
            if not items:
                continue
            start_number = 1 + (0 if column_index == 0 else split)
            end_number = start_number + len(items) - 1
            label_font = self.fonts.get(TextStyle(bold=True), reference_meta_size)
            label = str(slide.get("reference_label", "参考文献"))
            draw.text(
                (x, body_top),
                f"{label}  {start_number:02d}–{end_number:02d}",
                font=label_font,
                fill=self.color("muted"),
            )
            label_line_y = body_top + int(reference_meta_size * 1.7)
            draw.line((x, label_line_y, x + column_width, label_line_y), fill=self.color("line"), width=1)
            list_top = body_top + int(reference_meta_size * 2.5)
            item_gap = 16
            item_height = (bottom - list_top - item_gap * (len(items) - 1)) // len(items)
            for item_index, reference in enumerate(items):
                item_top = list_top + item_index * (item_height + item_gap)
                number = start_number + item_index
                number_size = max(28, int(reference_title_size * 0.8))
                number_font = self.fonts.get(TextStyle(bold=True), number_size)
                draw.text((x, item_top + 2), f"{number:02d}", font=number_font, fill=self.color("blue_pale"))
                number_gap = max(52, int(number_size * 1.6))
                text_left = x + number_gap
                text_width_available = column_width - number_gap
                cursor = item_top
                authors = str(reference.get("authors", "")).strip()
                if authors:
                    authors_font, authors_lines, _, authors_fits = self.fit_text(
                        draw,
                        authors,
                        self.theme["fonts"]["body"] if self.has_cjk(authors) else self.latin_fonts,
                        text_width_available,
                        reference_meta_size,
                        12,
                        12,
                        max_lines=1,
                        line_ratio=1.15,
                    )
                    if not authors_fits:
                        self.warn(str(slide.get("id", "")), f"Reference authors overflowed: {authors}")
                    draw.text(
                        (text_left, cursor),
                        authors_lines[0] if authors_lines else authors,
                        font=authors_font,
                        fill=self.color("muted"),
                    )
                    cursor += int(line_height(authors_font, 1.15) * 1.05)

                title_text = str(reference.get("title", "")).strip()
                title_font, title_lines, _, title_fits = self.fit_text(
                    draw,
                    title_text,
                    self.theme["fonts"]["title"] if self.has_cjk(title_text) else self.latin_fonts,
                    text_width_available,
                    reference_title_size * 2.4,
                    reference_title_size,
                    16,
                    max_lines=2,
                    line_ratio=1.2,
                )
                if not title_fits:
                    self.warn(str(slide.get("id", "")), f"Reference title overflowed: {title_text}")
                draw_lines(
                    draw,
                    title_lines,
                    (text_left, cursor),
                    title_font,
                    self.color("blue_dark"),
                    1.2,
                )
                cursor += line_height(title_font, 1.2) * max(1, len(title_lines)) + 4

                venue = str(reference.get("venue", "")).strip()
                if venue:
                    venue_font, venue_lines, _, venue_fits = self.fit_text(
                        draw,
                        venue,
                        self.theme["fonts"]["body"] if self.has_cjk(venue) else self.latin_fonts,
                        text_width_available,
                        reference_meta_size * 1.6,
                        reference_meta_size,
                        13,
                        max_lines=1,
                        line_ratio=1.15,
                    )
                    if not venue_fits:
                        self.warn(str(slide.get("id", "")), f"Reference venue overflowed: {venue}")
                    draw.text(
                        (text_left, cursor),
                        venue_lines[0] if venue_lines else venue,
                        font=venue_font,
                        fill=self.color("ink"),
                    )
                    cursor += int(line_height(venue_font, 1.15) * 1.05)

                doi = str(reference.get("doi", "")).strip()
                if doi:
                    doi_font, doi_lines, _, doi_fits = self.fit_text(
                        draw,
                        doi,
                        self.theme["fonts"]["body"] if self.has_cjk(doi) else self.latin_fonts,
                        text_width_available,
                        reference_meta_size * 1.4,
                        max(16, int(reference_meta_size * 0.9)),
                        11,
                        max_lines=1,
                        line_ratio=1.1,
                    )
                    if not doi_fits:
                        self.warn(str(slide.get("id", "")), f"Reference DOI overflowed: {doi}")
                    draw.text(
                        (text_left, cursor),
                        doi_lines[0] if doi_lines else doi,
                        font=doi_font,
                        fill=self.color("muted"),
                    )
                if item_index < len(items) - 1:
                    draw.line(
                        (x, item_top + item_height - 4, x + column_width, item_top + item_height - 4),
                        fill=mix_color(self.color("line"), self.color("white"), 0.2),
                        width=1,
                    )

        self.draw_takeaway(image, slide, str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def render(self) -> dict:
        slides = self.deck.get("slides", [])
        if self.page_selection:
            slides = [slides[page - 1] for page in self.page_selection if page <= len(slides)]
        rendered_paths: list[Path] = []
        for index, slide in enumerate(slides, start=1):
            layout = slide.get("layout")
            if layout == "cover":
                image = self.draw_cover(slide)
            elif layout == "outline":
                image = self.draw_outline(slide)
            elif layout == "section":
                image = self.draw_section(slide, index)
            elif layout == "bullets":
                image = self.draw_bullets(slide, index)
            elif layout == "figure":
                image = self.draw_figure(slide, index)
            elif layout == "figure_flow":
                image = self.draw_figure_flow(slide, index)
            elif layout == "concept":
                image = self.draw_concept(slide, index)
            elif layout == "formula":
                image = self.draw_formula(slide, index)
            elif layout == "pipeline":
                image = self.draw_pipeline(slide, index)
            elif layout == "compare":
                image = self.draw_compare(slide, index)
            elif layout == "metrics":
                image = self.draw_metrics(slide, index)
            elif layout == "closing":
                image = self.draw_closing(slide, index)
            elif layout == "grid":
                image = self.draw_grid(slide, index)
            elif layout == "balance":
                image = self.draw_balance(slide, index)
            elif layout == "hub":
                image = self.draw_hub(slide, index)
            elif layout == "timeline":
                image = self.draw_timeline(slide, index)
            elif layout == "references":
                image = self.draw_references(slide, index)
            elif layout == "quote":
                image = self.draw_quote(slide, index)
            else:
                raise ValueError(f"Unsupported layout {layout!r} on slide {index}")
            output_path = self.render_dir / f"slide_{index:02d}.png"
            if self.profile == "defense":
                draw = ImageDraw.Draw(image)
                draw.rectangle(
                    (24, int(self.layout["footer_y"]), self.width - 24, int(self.layout["footer_y"]) + 3),
                    fill=self.color("blue"),
                )
            image.save(output_path, optimize=True)
            rendered_paths.append(output_path)
        pptx_path = self.output_dir / "defense_deck_raster.pptx"
        self.build_pptx(rendered_paths, pptx_path)
        contact_path = self.output_dir / "contact_sheet.jpg"
        self.build_contact_sheet(rendered_paths, contact_path)
        copied_deck = self.output_dir / "deck.json"
        copy_if_distinct(self.deck_path, copied_deck)
        report = {
            "theme": self.theme["name"],
            "canvas": [self.width, self.height],
            "slide_count": len(slides),
            "rendered": [str(path) for path in rendered_paths],
            "raster_pptx": str(pptx_path),
            "contact_sheet": str(contact_path),
            "warnings": self.warnings,
        }
        write_json(self.output_dir / "render_report.json", report)
        return report

    def build_pptx(self, rendered_paths: list[Path], target: Path) -> None:
        presentation = Presentation()
        presentation.slide_width = Inches(13.333333)
        presentation.slide_height = Inches(7.5)
        blank_layout = presentation.slide_layouts[6]
        slides = self.deck.get("slides", [])
        for index, image_path in enumerate(rendered_paths):
            slide = presentation.slides.add_slide(blank_layout)
            slide.shapes.add_picture(str(image_path), 0, 0, width=presentation.slide_width, height=presentation.slide_height)
            speech = str(slides[index].get("speech", "")).strip()
            if speech:
                slide.notes_slide.notes_text_frame.text = speech
        presentation.save(str(target))

    @staticmethod
    def contact_columns(count: int) -> int:
        if count <= 1:
            return 1
        best_columns = 3
        best_score: tuple[int, float] | None = None
        first_columns = 3 if count >= 5 else 2
        for columns in range(first_columns, 6):
            rows = math.ceil(count / columns)
            blanks = rows * columns - count
            score = (blanks, abs(columns - 3.6))
            if best_score is None or score < best_score:
                best_score = score
                best_columns = columns
        return best_columns

    def build_contact_sheet(self, rendered_paths: list[Path], target: Path) -> None:
        columns = self.contact_columns(len(rendered_paths))
        thumb_width = 620 if columns <= 4 else 500
        thumb_height, label_height = round(thumb_width * 9 / 16), 42
        rows = math.ceil(len(rendered_paths) / columns)
        sheet = Image.new("RGB", (thumb_width * columns, (thumb_height + label_height) * rows), self.color("paper"))
        draw = ImageDraw.Draw(sheet)
        label_font = self.fonts.get(TextStyle(bold=True), 18)
        slides = self.deck.get("slides", [])
        for index, path in enumerate(rendered_paths):
            image = Image.open(path).convert("RGB").resize((thumb_width, thumb_height), Image.Resampling.LANCZOS)
            column = index % columns
            row = index // columns
            left = column * thumb_width
            top = row * (thumb_height + label_height)
            sheet.paste(image, (left, top))
            draw.rectangle((left, top + thumb_height, left + thumb_width, top + thumb_height + label_height), fill=self.color("blue_wash"))
            title = str(slides[index].get("title", ""))
            draw.text((left + 12, top + thumb_height + 9), f"{index + 1:02d}｜{title}", font=label_font, fill=self.color("blue"))
        sheet.save(target, quality=94)


def main() -> None:
    args = parse_args()
    pages = None
    if args.pages:
        pages = [int(value) for value in str(args.pages).replace("，", ",").split(",") if value.strip()]
    renderer = DeckRenderer(
        Path(args.deck),
        Path(args.theme),
        Path(args.out),
        profile=args.profile,
        pages=pages,
        sample=args.sample,
    )
    report = renderer.render()
    print(json.dumps(report, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
