from __future__ import annotations

import argparse
import json
import math
import shutil
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches

from common import (
    draw_bullet_list,
    draw_lines,
    draw_panel,
    draw_wrapped_block,
    fit_text_block,
    line_height,
    load_json,
    load_theme,
    mix_color,
    paste_contain,
    resolve_font,
    resolve_path,
    text_width,
    wrap_cjk_text,
    write_json,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render an academic defense deck from deck.json.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--theme", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--columns", type=int, default=4)
    return parser.parse_args()


class DeckRenderer:
    def __init__(self, deck_path: Path, theme_path: Path, output_dir: Path) -> None:
        self.deck_path = deck_path.resolve()
        self.deck = load_json(self.deck_path)
        self.theme = load_theme(theme_path)
        self.output_dir = output_dir.resolve()
        self.render_dir = self.output_dir / "rendered"
        self.render_dir.mkdir(parents=True, exist_ok=True)
        self.canvas = self.theme["canvas"]
        self.width = int(self.canvas["width"])
        self.height = int(self.canvas["height"])
        self.colors = self.theme["_colors"]
        self.layout = self.theme["layout"]
        self.title_fonts = self.theme["fonts"]["title"]
        self.body_fonts = self.theme["fonts"]["body"]
        self.serif_fonts = self.theme["fonts"]["serif"]
        self.display_fonts = self.theme["fonts"].get("display", self.serif_fonts)
        self.formula_fonts = self.theme["fonts"].get("formula", self.serif_fonts)
        self.type_scale = self.theme.get("type_scale", {})
        self.warnings: list[dict] = []
        self.section_map = {section["id"]: index for index, section in enumerate(self.deck.get("sections", []))}
        self.lockup = self.load_lockup()

    def color(self, name: str) -> tuple[int, int, int]:
        return self.colors[name]

    def type_size(self, name: str, fallback: int) -> int:
        return int(self.type_scale.get(name, fallback))

    def load_lockup(self) -> Image.Image | None:
        lockup_path = self.theme.get("brand", {}).get("lockup_path")
        if not lockup_path:
            return None
        path = Path(lockup_path)
        if not path.exists():
            self.warn("brand", f"Brand lockup missing: {path}")
            return None
        return Image.open(path).convert("RGBA")

    def warn(self, page: str, message: str) -> None:
        self.warnings.append({"page": page, "message": message})

    def background(self, page_number: int, cover: bool = False) -> Image.Image:
        paper = np.array(self.color("paper"), dtype=np.float32)
        paper_alt = np.array(self.color("paper_alt"), dtype=np.float32)
        gradient = np.zeros((self.height, self.width, 3), dtype=np.float32)
        for row in range(self.height):
            amount = row / max(1, self.height - 1)
            gradient[row, :, :] = paper * (1 - amount * 0.55) + paper_alt * (amount * 0.55)
        random = np.random.default_rng(4731 + page_number)
        noise = random.normal(0, 1.4, size=(self.height, self.width, 1))
        gradient = np.clip(gradient + noise, 0, 255).astype(np.uint8)
        image = Image.fromarray(gradient, mode="RGB").convert("RGBA")
        overlay = Image.new("RGBA", (self.width, self.height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)
        line_color = mix_color(self.color("line"), self.color("paper"), 0.35) + (170,)
        accent_color = mix_color(self.color("blue_mid"), self.color("paper"), 0.72) + (90,)
        draw.arc((-520, 880, 2580, 1770), 200, 340, fill=line_color, width=5)
        draw.arc((-260, 1010, 2300, 1590), 198, 342, fill=accent_color, width=3)
        if cover:
            draw.arc((-820, 760, 2860, 1840), 200, 340, fill=line_color, width=8)
            draw.arc((-420, 950, 2460, 1680), 198, 342, fill=accent_color, width=4)
        return Image.alpha_composite(image, overlay).convert("RGB")

    def draw_brand(self, image: Image.Image, box: tuple[int, int, int, int], fallback_size: int = 30) -> None:
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
        font, lines, _, fits = fit_text_block(
            draw,
            text,
            self.title_fonts,
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
        draw_lines(draw, lines, (left, start_y), font, self.color("blue"), align="center", max_width=inside_width)

    @staticmethod
    def fit_transparent(image: Image.Image, box: tuple[int, int, int, int]) -> tuple[Image.Image, tuple[int, int]]:
        left, top, right, bottom = box
        width, height = right - left, bottom - top
        scale = min(width / image.width, height / image.height)
        size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
        resized = image.resize(size, Image.Resampling.LANCZOS)
        return resized, (left + (width - size[0]) // 2, top + (height - size[1]) // 2)

    def draw_header(self, image: Image.Image, slide: dict) -> None:
        draw = ImageDraw.Draw(image)
        blue = self.color("blue")
        blue_mid = self.color("blue_mid")
        slash_left = self.layout["slash_x"]
        slash_top = self.layout["slash_y"]
        slash_width = self.layout["slash_width"]
        slash_height = self.layout["slash_height"]
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
        if eyebrow:
            font = resolve_font(self.title_fonts, self.type_size("eyebrow", int(self.layout["chapter_size"])))
            draw.text((self.layout["chapter_x"], self.layout["chapter_y"]), eyebrow, font=font, fill=blue)

        title = str(slide.get("title", "")).strip()
        title_left = int(self.layout["title_x"])
        title_top = int(self.layout["title_y"])
        title_width = int(self.theme["brand"]["header_box"][0] - title_left - 46)
        font, lines, _, fits = fit_text_block(
            draw,
            title,
            self.title_fonts,
            title_width,
            54,
            self.type_size("header_title", int(self.layout["title_size"])),
            34,
            max_lines=1,
        )
        if not fits:
            self.warn(str(slide.get("id", "")), f"Header title shrank or overflowed: {title}")
        draw.text((title_left, title_top), lines[0] if lines else title, font=font, fill=blue)

        divider_y = int(self.layout["divider_y"])
        draw.rectangle((24, divider_y, self.width - 24, divider_y + 3), fill=blue)
        draw.rectangle((24, divider_y + 5, self.width - 24, divider_y + 6), fill=mix_color(blue, self.color("paper"), 0.55))
        self.draw_brand(image, tuple(self.layout["brand"]["header_box"]) if "brand" in self.layout else tuple(self.theme["brand"]["header_box"]))
        draw.rectangle((24, int(self.layout["footer_y"]), self.width - 24, int(self.layout["footer_y"]) + 3), fill=blue)

    def draw_sidebar(self, image: Image.Image, active_section: str | None) -> None:
        sections = self.deck.get("sections", [])
        if not sections:
            return
        draw = ImageDraw.Draw(image)
        left, top, right, bottom = map(int, self.layout["sidebar"])
        width = right - left
        gap = 10
        row_height = min(92, (bottom - top - gap * (len(sections) - 1)) // len(sections))
        for index, section in enumerate(sections):
            row_top = top + index * (row_height + gap)
            active = section["id"] == active_section
            fill = self.color("blue") if active else self.color("white")
            outline = self.color("blue") if active else self.color("line")
            draw_panel(draw, (left, row_top, right, row_top + row_height), fill, outline, radius=7, width=2)
            text_color = self.color("white") if active else self.color("blue")
            font, lines, _, fits = fit_text_block(
                draw,
                section["title"],
                self.title_fonts,
                width - 32,
                row_height - 12,
                self.type_size("sidebar", 28),
                17,
                max_lines=2,
                line_ratio=1.16,
            )
            if not fits:
                self.warn(str(section.get("id", "")), "Sidebar label did not fit.")
            block_height = line_height(font, 1.16) * len(lines)
            text_top = row_top + (row_height - block_height) // 2
            draw_lines(draw, lines, (left + 16, text_top), font, text_color, 1.16, "center", width - 32)
            if active:
                draw.polygon(
                    [
                        (right - 17, row_top + row_height // 2 - 8),
                        (right - 8, row_top + row_height // 2),
                        (right - 17, row_top + row_height // 2 + 8),
                    ],
                    fill=self.color("white"),
                )

    def draw_lead(self, image: Image.Image, text: str, top: int, page_id: str) -> int:
        if not text:
            return top
        draw = ImageDraw.Draw(image)
        left, _, right, _ = map(int, self.layout["content"])
        height = int(self.layout["lead_height"])
        draw_panel(draw, (left, top, right, top + height), self.color("blue_wash"), self.color("line"), radius=6, width=2)
        draw.rectangle((left, top, left + 11, top + height), fill=self.color("blue"))
        marker_top = top + 20
        for marker_index in range(3):
            marker_left = left + 28 + marker_index * 17
            draw.line(
                (
                    marker_left,
                    marker_top,
                    marker_left + 12,
                    marker_top + 17,
                    marker_left,
                    marker_top + 34,
                ),
                fill=self.color("blue_mid") if marker_index == 0 else self.color("line"),
                width=4,
            )
        fits, _ = draw_wrapped_block(
            draw,
            text,
            (left + 88, top + 13),
            right - left - 108,
            height - 22,
            self.title_fonts,
            self.color("blue_dark"),
            self.type_size("lead", 29),
            21,
            max_lines=2,
            line_ratio=1.18,
        )
        if not fits:
            self.warn(page_id, "Lead text overflowed.")
        return top + height + 22

    def draw_takeaway(self, image: Image.Image, text: str, page_id: str) -> None:
        if not text:
            return
        draw = ImageDraw.Draw(image)
        left = int(self.layout["content"][0])
        right = int(self.layout["content"][2])
        top = int(self.layout["takeaway_top"])
        height = int(self.layout["takeaway_height"])
        draw_panel(draw, (left, top, right, top + height), self.color("blue_dark"), None, radius=7, width=0)
        fits, _ = draw_wrapped_block(
            draw,
            text,
            (left + 24, top + 12),
            right - left - 48,
            height - 20,
            self.title_fonts,
            self.color("white"),
            self.type_size("takeaway", 28),
            19,
            max_lines=2,
            align="center",
            line_ratio=1.16,
        )
        if not fits:
            self.warn(page_id, "Takeaway overflowed.")

    def content_bounds(self, slide: dict) -> tuple[int, int, int, int]:
        left, top, right, bottom = map(int, self.layout["content"])
        if slide.get("takeaway"):
            bottom = int(self.layout["takeaway_top"]) - 18
        return left, top, right, bottom

    def draw_cover(self, slide: dict) -> Image.Image:
        image = self.background(1, cover=True)
        draw = ImageDraw.Draw(image)
        cover = self.theme["cover"]
        meta = self.deck.get("meta", {})
        self.draw_brand(image, tuple(self.theme["brand"]["cover_box"]), fallback_size=46)

        title_width = 1780
        title_left = (self.width - title_width) // 2
        font, lines, _, fits = fit_text_block(
            draw,
            slide.get("title", meta.get("title", "")),
            self.display_fonts,
            title_width,
            118,
            self.type_size("cover_title", int(cover["title_size"])),
            50,
            max_lines=2,
            line_ratio=1.16,
        )
        if not fits:
            self.warn(str(slide.get("id", "")), "Cover title overflowed.")
        title_y = int(cover["title_y"])
        draw_lines(draw, lines, (title_left, title_y), font, self.color("blue"), 1.16, "center", title_width)

        subtitle = slide.get("subtitle") or meta.get("subtitle")
        if subtitle:
            subtitle_font, subtitle_lines, _, subtitle_fits = fit_text_block(
                draw,
                subtitle,
                self.body_fonts,
                title_width,
                54,
                self.type_size("cover_subtitle", int(cover["subtitle_size"])),
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
        draw.rectangle((300, rule_y, self.width - 300, rule_y + 3), fill=self.color("blue"))
        draw.rectangle((300, rule_y + 9, self.width - 300, rule_y + 10), fill=mix_color(self.color("blue_mid"), self.color("paper"), 0.25))

        presenter = meta.get("presenter", "")
        advisor = meta.get("advisor", "")
        date = meta.get("date", "")
        info_font = resolve_font(self.body_fonts, int(cover["info_size"]))
        first_line = f"汇报人：{presenter}"
        second_line = f"指导教师：{advisor}"
        date_line = f"汇报时间：{date}"
        group_width = text_width(draw, first_line, info_font) + text_width(draw, second_line, info_font) + 92
        group_left = (self.width - group_width) // 2
        info_y = int(cover["info_y"])
        icon_color = self.color("blue")
        self.draw_person_icon(draw, group_left, info_y + 4, 28, icon_color)
        draw.text((group_left + 40, info_y), first_line, font=info_font, fill=self.color("ink"))
        separator_x = group_left + 40 + text_width(draw, first_line, info_font) + 32
        draw.line((separator_x, info_y + 2, separator_x, info_y + 36), fill=self.color("line"), width=3)
        self.draw_cap_icon(draw, separator_x + 24, info_y + 3, 30, icon_color)
        draw.text((separator_x + 64, info_y), second_line, font=info_font, fill=self.color("ink"))
        date_width = text_width(draw, date_line, info_font)
        date_left = (self.width - date_width) // 2
        self.draw_calendar_icon(draw, date_left - 42, info_y + 68, 28, icon_color)
        draw.text((date_left, info_y + 62), date_line, font=info_font, fill=self.color("ink"))
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
        row_height = int(outline["row_height"])
        gap = int(outline["gap"])
        sections = self.deck.get("sections", [])
        draw.line((left - 20, top - 10, left - 20, top + len(sections) * (row_height + gap) - gap + 10), fill=self.color("blue"), width=3)
        for index, section in enumerate(sections):
            row_top = top + index * (row_height + gap)
            draw_panel(draw, (left, row_top, left + width, row_top + row_height), self.color("blue_wash"), self.color("line"), radius=8, width=2)
            number_width = 190
            number_box = [
                (left, row_top),
                (left + number_width - 24, row_top),
                (left + number_width, row_top + row_height // 2),
                (left + number_width - 24, row_top + row_height),
                (left, row_top + row_height),
            ]
            draw.polygon(number_box, fill=self.color("blue"))
            number_font = resolve_font(self.title_fonts, 54)
            number_text = f"{index + 1:02d}"
            number_bbox = draw.textbbox((0, 0), number_text, font=number_font)
            number_y = row_top + (row_height - (number_bbox[3] - number_bbox[1])) // 2 - 4
            draw.text((left + 48, number_y), number_text, font=number_font, fill=self.color("white"))
            title_font, title_lines, _, title_fits = fit_text_block(
                draw,
                section["title"],
                self.title_fonts,
                width - number_width - 48,
                row_height - 22,
                42,
                26,
                max_lines=1,
            )
            if not title_fits:
                self.warn(str(slide.get("id", "")), f"Outline item did not fit: {section['title']}")
            title_y = row_top + (row_height - line_height(title_font)) // 2 - 2
            draw.text((left + number_width + 30, title_y), title_lines[0] if title_lines else section["title"], font=title_font, fill=self.color("blue_dark"))
        return image

    def draw_frame(self, slide: dict, page_number: int) -> tuple[Image.Image, ImageDraw.ImageDraw]:
        image = self.background(page_number)
        self.draw_header(image, slide)
        self.draw_sidebar(image, slide.get("section"))
        draw = ImageDraw.Draw(image)
        return image, draw

    def draw_bullets(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        columns = slide.get("columns") or []
        if not columns and slide.get("bullets"):
            columns = [{"heading": "", "bullets": slide.get("bullets", [])}]
        if not columns:
            self.warn(str(slide.get("id", "")), "Bullets slide has no content.")
            columns = [{"heading": "", "bullets": ["内容待补充"]}]
        column_count = len(columns)
        gap = 22 if column_count > 2 else 28
        column_width = (right - left - gap * (column_count - 1)) // column_count
        panel_top = body_top
        panel_bottom = bottom
        for index, column in enumerate(columns):
            column_left = left + index * (column_width + gap)
            draw_panel(draw, (column_left, panel_top, column_left + column_width, panel_bottom), self.color("white"), self.color("line"), radius=8, width=2)
            heading = str(column.get("heading", "")).strip()
            heading_height = 0
            if heading:
                heading_font, heading_lines, _, heading_fits = fit_text_block(
                    draw,
                    heading,
                    self.title_fonts,
                    column_width - 36,
                    48,
                    self.type_size("block_title", 30),
                    21,
                    max_lines=1,
                )
                if not heading_fits:
                    self.warn(str(slide.get("id", "")), f"Column heading overflowed: {heading}")
                draw.text((column_left + 18, panel_top + 18), heading_lines[0] if heading_lines else heading, font=heading_font, fill=self.color("blue"))
                draw.line((column_left + 18, panel_top + 62, column_left + column_width - 18, panel_top + 62), fill=self.color("line"), width=2)
                heading_height = 54
            bullet_top = panel_top + 72 + heading_height
            fits, _ = draw_bullet_list(
                draw,
                column.get("bullets", []),
                (column_left + 18, bullet_top),
                column_width - 36,
                panel_bottom - bullet_top - 16,
                self.body_fonts,
                self.color("ink"),
                self.color("blue_mid"),
                max_size=self.type_size("bullet", 26),
                min_size=self.type_size("small", 18),
            )
            if not fits:
                self.warn(str(slide.get("id", "")), f"Column text overflowed: {heading or index + 1}")
        self.draw_takeaway(image, slide.get("takeaway", ""), str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_figure(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        image_path = resolve_path(slide.get("image"), self.deck_path.parent)
        if image_path is None or not image_path.exists():
            self.warn(str(slide.get("id", "")), f"Missing figure: {slide.get('image')}")
            draw_panel(draw, (left, body_top, left + (right - left) // 2, bottom), self.color("blue_wash"), self.color("line"), radius=8, width=2)
            draw_wrapped_block(
                draw,
                "图片缺失\n请检查 deck.json 中的 image 路径",
                (left + 24, body_top + 24),
                (right - left) // 2 - 48,
                bottom - body_top - 48,
                self.body_fonts,
                self.color("muted"),
                self.type_size("body", 28),
                20,
                max_lines=4,
            )
            return image
        side = slide.get("image_side", "left")
        total_width = right - left
        gap = 24
        image_width = round(total_width * 0.55)
        text_width_available = total_width - image_width - gap
        if side == "right":
            image_box = (right - image_width, body_top, right, bottom)
            text_left = left
        else:
            image_box = (left, body_top, left + image_width, bottom)
            text_left = left + image_width + gap
        draw_panel(draw, image_box, self.color("white"), self.color("line"), radius=8, width=2)
        inset = 14
        figure = Image.open(image_path).convert("RGB")
        paste_contain(image, figure, (image_box[0] + inset, image_box[1] + inset, image_box[2] - inset, image_box[3] - inset), self.color("white"))
        draw = ImageDraw.Draw(image)
        panels = slide.get("panels") or []
        if not panels:
            panels = [{"heading": "图示", "body": slide.get("image_role", "请结合原图讲解。")}]
        panel_gap = 16
        panel_height = (bottom - body_top - panel_gap * (len(panels) - 1)) // len(panels)
        for index, panel in enumerate(panels):
            panel_top = body_top + index * (panel_height + panel_gap)
            draw_panel(
                draw,
                (text_left, panel_top, text_left + text_width_available, panel_top + panel_height),
                self.color("white"),
                self.color("line"),
                radius=7,
                width=2,
            )
            heading = str(panel.get("heading", "")).strip()
            body = str(panel.get("body", "")).strip()
            heading_space = 0
            if heading:
                heading_bar_height = 36
                draw.rectangle(
                    (text_left + 2, panel_top + 2, text_left + text_width_available - 2, panel_top + heading_bar_height),
                    fill=self.color("blue"),
                )
                heading_font, heading_lines, _, heading_fits = fit_text_block(
                    draw,
                    heading,
                    self.title_fonts,
                    text_width_available - 36,
                    30,
                    self.type_size("block_title", 25),
                    19,
                    max_lines=1,
                )
                if not heading_fits:
                    self.warn(str(slide.get("id", "")), f"Figure panel heading overflowed: {heading}")
                draw.text((text_left + 16, panel_top + 6), heading_lines[0] if heading_lines else heading, font=heading_font, fill=self.color("white"))
                heading_space = 46
            fits, _ = draw_wrapped_block(
                draw,
                body,
                (text_left + 16, panel_top + 8 + heading_space),
                text_width_available - 32,
                panel_height - 20 - heading_space,
                self.body_fonts,
                self.color("ink"),
                self.type_size("body", 22),
                self.type_size("small", 16),
                max_lines=4,
                line_ratio=1.24,
            )
            if not fits:
                self.warn(str(slide.get("id", "")), f"Figure panel overflowed: {heading or index + 1}")
        self.draw_takeaway(image, slide.get("takeaway", ""), str(slide.get("id", "")))
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
        row_top = body_top + 36
        step_height = min(420, bottom - row_top - 28)
        connector_y = row_top + 54
        draw.line((left + 20, connector_y, right - 20, connector_y), fill=self.color("line"), width=4)
        for index, step in enumerate(steps):
            step_left = left + index * (step_width + gap)
            center_x = step_left + step_width // 2
            draw.ellipse((center_x - 25, connector_y - 25, center_x + 25, connector_y + 25), fill=self.color("blue"), outline=self.color("white"), width=4)
            number_font = resolve_font(self.title_fonts, 23)
            number_text = str(index + 1)
            number_box = draw.textbbox((0, 0), number_text, font=number_font)
            draw.text((center_x - (number_box[2] - number_box[0]) // 2, connector_y - 14), number_text, font=number_font, fill=self.color("white"))
            draw_panel(draw, (step_left, row_top + 82, step_left + step_width, row_top + step_height), self.color("white"), self.color("line"), radius=8, width=2)
            heading_font, heading_lines, _, heading_fits = fit_text_block(
                draw,
                step.get("title", ""),
                self.title_fonts,
                step_width - 34,
                42,
                self.type_size("block_title", 28),
                19,
                max_lines=2,
            )
            if not heading_fits:
                self.warn(str(slide.get("id", "")), f"Pipeline title overflowed: {step.get('title')}")
            draw_lines(draw, heading_lines, (step_left + 17, row_top + 102), heading_font, self.color("blue"), 1.2, "center", step_width - 34)
            fits, _ = draw_wrapped_block(
                draw,
                step.get("body", ""),
                (step_left + 17, row_top + 164),
                step_width - 34,
                96,
                self.body_fonts,
                self.color("ink"),
                self.type_size("body", 23),
                self.type_size("small", 16),
                max_lines=3,
                line_ratio=1.24,
            )
            if not fits:
                self.warn(str(slide.get("id", "")), f"Pipeline body overflowed: {step.get('title')}")
            bullets = step.get("bullets") or []
            if bullets:
                fits, _ = self.draw_compact_bullets(
                    draw,
                    bullets,
                    step_left + 17,
                    row_top + 268,
                    step_width - 34,
                    step_height - 286,
                    max_size=self.type_size("bullet", 21),
                    min_size=self.type_size("micro", 15),
                )
                if not fits:
                    self.warn(str(slide.get("id", "")), f"Pipeline bullets overflowed: {step.get('title')}")
        self.draw_takeaway(image, slide.get("takeaway", ""), str(slide.get("id", "")))
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
        header_height = int(table["header_height"])
        row_height = min(132, max(72, (bottom - body_top - header_height) // max(1, len(rows))))
        weights = [1.45] + [1.0] * (len(columns) - 1)
        total_weight = sum(weights)
        widths = [round((right - left) * weight / total_weight) for weight in weights]
        widths[-1] = right - left - sum(widths[:-1])
        table_height = header_height + row_height * len(rows)
        header_top = body_top + max(0, (bottom - body_top - table_height) // 2)
        current_left = left
        for index, column in enumerate(columns):
            cell_box = (current_left, header_top, current_left + widths[index], header_top + header_height)
            draw.rectangle(cell_box, fill=self.color("blue_dark"))
            if index:
                draw.line((current_left, header_top, current_left, bottom), fill=self.color("white"), width=2)
            heading_font, heading_lines, _, heading_fits = fit_text_block(
                draw,
                column,
                self.title_fonts,
                widths[index] - 18,
                header_height - 12,
                self.type_size("block_title", 26),
                17,
                max_lines=1,
            )
            if not heading_fits:
                self.warn(str(slide.get("id", "")), f"Table heading overflowed: {column}")
            draw.text((current_left + 9, header_top + 15), heading_lines[0] if heading_lines else column, font=heading_font, fill=self.color("white"))
            current_left += widths[index]
        for row_index, row in enumerate(rows):
            row_top = header_top + header_height + row_index * row_height
            fill = self.color("white") if row_index % 2 == 0 else self.color("blue_wash")
            current_left = left
            for column_index, value in enumerate(row[:len(columns)]):
                cell_box = (current_left, row_top, current_left + widths[column_index], row_top + row_height)
                draw.rectangle(cell_box, fill=fill, outline=self.color("line"), width=1)
                text_font, text_lines, _, text_fits = fit_text_block(
                    draw,
                    str(value),
                    self.body_fonts,
                    widths[column_index] - 20,
                    row_height - 18,
                    self.type_size("body", 24),
                    self.type_size("micro", 15),
                    max_lines=3,
                    line_ratio=1.2,
                )
                if not text_fits:
                    self.warn(str(slide.get("id", "")), f"Table cell overflowed: {value}")
                block_height = line_height(text_font, 1.2) * len(text_lines)
                draw_lines(
                    draw,
                    text_lines,
                    (current_left + 10, row_top + (row_height - block_height) // 2),
                    text_font,
                    self.color("ink"),
                    1.2,
                )
                current_left += widths[column_index]
        notes = slide.get("notes") or []
        table_bottom = header_top + table_height
        if notes and table_bottom + 20 < int(self.layout["takeaway_top"]) - 10:
            note_width = (right - left - 18) // 2
            note_height = int(self.layout["takeaway_top"]) - table_bottom - 24
            half = math.ceil(len(notes) / 2)
            self.draw_compact_bullets(draw, notes[:half], left, table_bottom + 18, note_width, note_height, self.type_size("bullet", 21), self.type_size("micro", 15))
            self.draw_compact_bullets(draw, notes[half:], left + note_width + 18, table_bottom + 18, note_width, note_height, self.type_size("bullet", 21), self.type_size("micro", 15))
        self.draw_takeaway(image, slide.get("takeaway", ""), str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_metrics(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        metrics = slide.get("metrics") or []
        if not metrics:
            self.warn(str(slide.get("id", "")), "Metrics slide has no metrics.")
            metrics = [{"value": "—", "label": "指标", "note": ""}]
        gap = 20
        card_width = (right - left - gap * (len(metrics) - 1)) // len(metrics)
        card_height = min(252, bottom - body_top - 140)
        for index, metric in enumerate(metrics):
            card_left = left + index * (card_width + gap)
            draw_panel(draw, (card_left, body_top, card_left + card_width, body_top + card_height), self.color("blue_wash") if index == 0 else self.color("white"), self.color("line"), radius=8, width=2)
            value_font, value_lines, _, value_fits = fit_text_block(
                draw,
                str(metric.get("value", "")),
                self.title_fonts,
                card_width - 32,
                84,
                self.type_size("metric_value", 58),
                27,
                max_lines=1,
            )
            if not value_fits:
                self.warn(str(slide.get("id", "")), f"Metric value overflowed: {metric.get('value')}")
            draw_lines(draw, value_lines, (card_left + 16, body_top + 28), value_font, self.color("blue"), align="center", max_width=card_width - 32)
            label_font, label_lines, _, label_fits = fit_text_block(
                draw,
                str(metric.get("label", "")),
                self.title_fonts,
                card_width - 32,
                50,
                self.type_size("metric_label", 26),
                self.type_size("small", 18),
                max_lines=2,
            )
            if not label_fits:
                self.warn(str(slide.get("id", "")), f"Metric label overflowed: {metric.get('label')}")
            draw_lines(draw, label_lines, (card_left + 16, body_top + 126), label_font, self.color("ink"), 1.18, "center", card_width - 32)
            note_font, note_lines, _, note_fits = fit_text_block(
                draw,
                str(metric.get("note", "")),
                self.body_fonts,
                card_width - 32,
                60,
                self.type_size("small", 21),
                self.type_size("micro", 15),
                max_lines=2,
            )
            if not note_fits:
                self.warn(str(slide.get("id", "")), f"Metric note overflowed: {metric.get('note')}")
            draw_lines(draw, note_lines, (card_left + 16, body_top + 186), note_font, self.color("muted"), 1.18, "center", card_width - 32)
        bullet_top = body_top + card_height + 30
        bullets = slide.get("bullets") or []
        if bullets:
            bullet_width = (right - left - gap) // 2
            half = math.ceil(len(bullets) / 2)
            self.draw_compact_bullets(
                draw,
                bullets[:half],
                left,
                bullet_top,
                bullet_width,
                bottom - bullet_top,
                max_size=self.type_size("bullet", 23),
                min_size=self.type_size("small", 17),
            )
            self.draw_compact_bullets(
                draw,
                bullets[half:],
                left + bullet_width + gap,
                bullet_top,
                bullet_width,
                bottom - bullet_top,
                max_size=self.type_size("bullet", 23),
                min_size=self.type_size("small", 17),
            )
        self.draw_takeaway(image, slide.get("takeaway", ""), str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_footnote(self, image: Image.Image, text: str) -> None:
        if not text:
            return
        draw = ImageDraw.Draw(image)
        left, _, right, _ = map(int, self.layout["content"])
        font, lines, _, fits = fit_text_block(
            draw,
            text,
            self.body_fonts,
            right - left,
            self.type_size("small", 26),
            self.type_size("micro", 17),
            14,
            max_lines=1,
        )
        if not fits:
            self.warn("footnote", f"Footnote overflowed: {text}")
        draw.text((left, int(self.layout["takeaway_top"]) + int(self.layout["takeaway_height"]) + 10), lines[0] if lines else text, font=font, fill=self.color("muted"))

    def draw_compact_bullets(
        self,
        draw: ImageDraw.ImageDraw,
        bullets: list[str],
        left: int,
        top: int,
        width: int,
        height: int,
        max_size: int = 23,
        min_size: int = 16,
    ) -> tuple[bool, int]:
        current_top = top
        all_fit = True
        for bullet in bullets:
            font, lines, block_height, fits = fit_text_block(
                draw,
                bullet,
                self.body_fonts,
                width - 24,
                58,
                max_size,
                min_size,
                max_lines=2,
                line_ratio=1.2,
            )
            if current_top + block_height > top + height:
                all_fit = False
                break
            draw.ellipse((left + 3, current_top + 9, left + 11, current_top + 17), fill=self.color("blue_mid"))
            draw_lines(draw, lines, (left + 22, current_top), font, self.color("ink"), 1.2)
            current_top += block_height + 10
            all_fit = all_fit and fits
        return all_fit, current_top

    def draw_academic_icon(
        self,
        draw: ImageDraw.ImageDraw,
        name: str,
        left: int,
        top: int,
        size: int,
        color: tuple[int, int, int],
        width: int = 2,
    ) -> None:
        right = left + size
        bottom = top + size
        center_x = left + size // 2
        center_y = top + size // 2
        if name in {"signal", "wave"}:
            points = [
                (left, top + size * 0.58),
                (left + size * 0.18, top + size * 0.40),
                (left + size * 0.36, top + size * 0.68),
                (left + size * 0.54, top + size * 0.22),
                (left + size * 0.72, top + size * 0.60),
                (right, top + size * 0.34),
            ]
            draw.line(points, fill=color, width=width, joint="curve")
        elif name == "target":
            draw.ellipse((left + 2, top + 2, right - 2, bottom - 2), outline=color, width=width)
            draw.ellipse((left + size * 0.30, top + size * 0.30, left + size * 0.70, top + size * 0.70), outline=color, width=width)
            draw.line((center_x, top, center_x, bottom), fill=color, width=1)
            draw.line((left, center_y, right, center_y), fill=color, width=1)
        elif name == "clock":
            draw.ellipse((left + 2, top + 2, right - 2, bottom - 2), outline=color, width=width)
            draw.line((center_x, center_y, center_x, top + size * 0.30), fill=color, width=width)
            draw.line((center_x, center_y, left + size * 0.68, center_y + size * 0.18), fill=color, width=width)
        elif name == "layers":
            draw.polygon([(center_x, top), (right, top + size * 0.25), (center_x, top + size * 0.50), (left, top + size * 0.25)], outline=color, width=width)
            draw.line((left, top + size * 0.52, center_x, top + size * 0.76, right, top + size * 0.52), fill=color, width=width, joint="curve")
            draw.line((left, top + size * 0.72, center_x, top + size * 0.96, right, top + size * 0.72), fill=color, width=width, joint="curve")
        elif name == "chart":
            draw.line((left + 2, top, left + 2, bottom - 2, right, bottom - 2), fill=color, width=width, joint="curve")
            draw.rectangle((left + size * 0.22, top + size * 0.54, left + size * 0.40, bottom - 4), outline=color, width=width)
            draw.rectangle((left + size * 0.48, top + size * 0.34, left + size * 0.66, bottom - 4), outline=color, width=width)
            draw.rectangle((left + size * 0.74, top + size * 0.16, left + size * 0.92, bottom - 4), outline=color, width=width)
        elif name == "check":
            draw.ellipse((left + 2, top + 2, right - 2, bottom - 2), outline=color, width=width)
            draw.line((left + size * 0.28, center_y, left + size * 0.44, top + size * 0.70, left + size * 0.76, top + size * 0.30), fill=color, width=width, joint="curve")
        elif name == "alert":
            draw.polygon([(center_x, top + 1), (right - 2, bottom - 3), (left + 2, bottom - 3)], outline=color, width=width)
            draw.line((center_x, top + size * 0.34, center_x, top + size * 0.64), fill=color, width=width)
            draw.ellipse((center_x - 2, bottom - 12, center_x + 2, bottom - 8), fill=color)
        elif name == "matrix":
            cell = size / 3
            for row in range(3):
                for column in range(3):
                    draw.rectangle((left + column * cell + 1, top + row * cell + 1, left + (column + 1) * cell - 2, top + (row + 1) * cell - 2), outline=color, width=1)
        else:
            draw.ellipse((left + size * 0.24, top + size * 0.24, right - size * 0.24, bottom - size * 0.24), fill=color)

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
        tone = block.get("tone", "white")
        fill = self.color("blue_wash") if tone == "pale" else self.color("white")
        draw_panel(draw, box, fill, self.color("line"), radius=7, width=2)
        heading = str(block.get("heading", "")).strip()
        heading_height = 40
        if heading:
            draw.rectangle((left + 2, top + 2, right - 2, top + heading_height), fill=self.color("blue"))
            number_font = resolve_font(self.title_fonts, 23)
            draw.text((left + 14, top + 7), f"{index:02d}", font=number_font, fill=self.color("blue_pale"))
            heading_font, heading_lines, _, heading_fits = fit_text_block(
                draw,
                heading,
                self.title_fonts,
                right - left - 78,
                30,
                self.type_size("block_title", 26),
                self.type_size("small", 18),
                max_lines=1,
            )
            if not heading_fits:
                self.warn(page_id, f"Grid heading overflowed: {heading}")
            draw.text((left + 58, top + 7), heading_lines[0] if heading_lines else heading, font=heading_font, fill=self.color("white"))
            icon_name = str(block.get("icon", "")).strip()
            if icon_name:
                self.draw_academic_icon(draw, icon_name, right - 36, top + 10, 20, self.color("white"), 2)
        content_top = top + heading_height + 12
        formula = str(block.get("formula", "")).strip()
        if formula:
            formula_box = (left + 14, content_top, right - 14, content_top + 54)
            draw_panel(draw, formula_box, self.color("blue_wash"), self.color("line"), radius=4, width=1)
            formula_font, formula_lines, _, formula_fits = fit_text_block(
                draw,
                formula,
                self.formula_fonts,
                formula_box[2] - formula_box[0] - 20,
                42,
                26,
                16,
                max_lines=1,
            )
            if not formula_fits:
                self.warn(page_id, f"Formula overflowed: {formula}")
            draw_lines(
                draw,
                formula_lines,
                (formula_box[0] + 10, formula_box[1] + 8),
                formula_font,
                self.color("blue_dark"),
                align="center",
                max_width=formula_box[2] - formula_box[0] - 20,
            )
            content_top = formula_box[3] + 10
        image_path = resolve_path(block.get("image"), self.deck_path.parent)
        if image_path and image_path.exists():
            image_height = round((bottom - content_top) * 0.44)
            image_box = (left + 14, content_top, right - 14, content_top + image_height)
            draw_panel(draw, image_box, self.color("white"), self.color("line"), radius=4, width=1)
            figure = Image.open(image_path).convert("RGB")
            paste_contain(image, figure, (image_box[0] + 8, image_box[1] + 8, image_box[2] - 8, image_box[3] - 8), self.color("white"))
            draw = ImageDraw.Draw(image)
            content_top = image_box[3] + 12
        body = str(block.get("body", "")).strip()
        if body:
            fits, body_height = draw_wrapped_block(
                draw,
                body,
                (left + 14, content_top),
                right - left - 28,
                max(34, bottom - content_top - 10),
                self.body_fonts,
                self.color("ink"),
                self.type_size("body", 22),
                self.type_size("small", 16),
                max_lines=3,
                line_ratio=1.22,
            )
            if not fits:
                self.warn(page_id, f"Grid body overflowed: {heading}")
            content_top += body_height + 8
        bullets = block.get("bullets") or []
        if bullets:
            fits, _ = self.draw_compact_bullets(
                draw,
                bullets,
                left + 14,
                content_top,
                right - left - 28,
                bottom - content_top - 10,
            )
            if not fits:
                self.warn(page_id, f"Grid bullets overflowed: {heading}")

    def draw_grid(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        blocks = slide.get("blocks") or []
        if not blocks:
            self.warn(str(slide.get("id", "")), "Grid layout has no blocks.")
            blocks = [{"heading": "内容待补充", "body": ""}]
        column_count = int(slide.get("columns", 2))
        column_count = max(2, min(3, column_count))
        row_count = math.ceil(len(blocks) / column_count)
        gap = 18
        cell_width = (right - left - gap * (column_count - 1)) // column_count
        cell_height = (bottom - body_top - gap * (row_count - 1)) // row_count
        for index, block in enumerate(blocks):
            row = index // column_count
            column = index % column_count
            block_left = left + column * (cell_width + gap)
            block_top = body_top + row * (cell_height + gap)
            remaining_in_row = min(column_count, len(blocks) - row * column_count)
            block_right = block_left + cell_width
            if remaining_in_row == 1 and column_count > 1:
                block_right = right
            self.draw_grid_block(
                image,
                draw,
                block,
                (block_left, block_top, block_right, block_top + cell_height),
                index + 1,
                str(slide.get("id", "")),
            )
        self.draw_takeaway(image, slide.get("takeaway", ""), str(slide.get("id", "")))
        self.draw_footnote(image, slide.get("footnote", ""))
        return image

    def draw_closing(self, slide: dict, page_number: int) -> Image.Image:
        image, draw = self.draw_frame(slide, page_number)
        left, top, right, bottom = self.content_bounds(slide)
        body_top = self.draw_lead(image, slide.get("lead", ""), top, str(slide.get("id", "")))
        points = slide.get("points") or []
        if not points:
            points = ["结论待补充"]
        gap = 18
        point_height = min(112, (bottom - body_top - 110 - gap * (len(points) - 1)) // len(points))
        for index, point in enumerate(points):
            row_top = body_top + index * (point_height + gap)
            draw_panel(draw, (left, row_top, right, row_top + point_height), self.color("white"), self.color("line"), radius=7, width=2)
            draw.rectangle((left, row_top, left + 8, row_top + point_height), fill=self.color("blue"))
            number_font = resolve_font(self.title_fonts, 25)
            draw.text((left + 26, row_top + 24), f"{index + 1:02d}", font=number_font, fill=self.color("blue_mid"))
            fits, _ = draw_wrapped_block(
                draw,
                str(point),
                (left + 94, row_top + 18),
                right - left - 124,
                point_height - 32,
                self.title_fonts,
                self.color("ink"),
                28,
                18,
                max_lines=2,
                line_ratio=1.2,
            )
            if not fits:
                self.warn(str(slide.get("id", "")), f"Closing point overflowed: {point}")
        closing = str(slide.get("closing", "谢谢各位老师")).strip()
        closing_font, closing_lines, _, closing_fits = fit_text_block(
            draw,
            closing,
            self.title_fonts,
            right - left - 80,
            70,
            46,
            28,
            max_lines=1,
        )
        if not closing_fits:
            self.warn(str(slide.get("id", "")), "Closing line overflowed.")
        draw_lines(draw, closing_lines, (left + 40, bottom - 74), closing_font, self.color("blue"), align="center", max_width=right - left - 80)
        return image

    def render(self) -> dict:
        slides = self.deck.get("slides", [])
        rendered_paths: list[Path] = []
        for index, slide in enumerate(slides, start=1):
            layout = slide.get("layout")
            if layout == "cover":
                image = self.draw_cover(slide)
            elif layout == "outline":
                image = self.draw_outline(slide)
            elif layout == "bullets":
                image = self.draw_bullets(slide, index)
            elif layout == "figure":
                image = self.draw_figure(slide, index)
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
            else:
                raise ValueError(f"Unsupported layout {layout!r} on slide {index}")
            output_path = self.render_dir / f"slide_{index:02d}.png"
            image.save(output_path, optimize=True)
            rendered_paths.append(output_path)
        pptx_path = self.output_dir / "defense_deck.pptx"
        self.build_pptx(rendered_paths)
        contact_path = self.output_dir / "contact_sheet.jpg"
        self.build_contact_sheet(rendered_paths, contact_path)
        copied_deck = self.output_dir / "deck.json"
        shutil.copy2(self.deck_path, copied_deck)
        report = {
            "theme": self.theme["name"],
            "canvas": [self.width, self.height],
            "slide_count": len(slides),
            "rendered": [str(path) for path in rendered_paths],
            "pptx": str(pptx_path),
            "contact_sheet": str(contact_path),
            "warnings": self.warnings,
        }
        write_json(self.output_dir / "render_report.json", report)
        return report

    def build_pptx(self, rendered_paths: list[Path]) -> None:
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
        presentation.save(str(self.output_dir / "defense_deck.pptx"))

    def build_contact_sheet(self, rendered_paths: list[Path], target: Path) -> None:
        columns = 4
        thumb_width, thumb_height, label_height = 512, 288, 32
        rows = math.ceil(len(rendered_paths) / columns)
        sheet = Image.new("RGB", (thumb_width * columns, (thumb_height + label_height) * rows), self.color("paper"))
        draw = ImageDraw.Draw(sheet)
        label_font = resolve_font(self.title_fonts, 18)
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
            draw.text((left + 10, top + thumb_height + 6), f"{index + 1:02d}｜{title}", font=label_font, fill=self.color("blue"))
        sheet.save(target, quality=94)


def main() -> None:
    args = parse_args()
    renderer = DeckRenderer(Path(args.deck), Path(args.theme), Path(args.out))
    report = renderer.render()
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
