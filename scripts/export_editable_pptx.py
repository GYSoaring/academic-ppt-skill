from __future__ import annotations

import argparse
import json
import math
import re
from pathlib import Path
from typing import Iterable, Sequence

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (
    XL_CHART_TYPE,
    XL_MARKER_STYLE,
    XL_TICK_LABEL_POSITION,
    XL_TICK_MARK,
)
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from common import (
    apply_navigation_mode,
    apply_type_profile,
    clamp_defense_size_px,
    clamp_defense_body_size_px,
    content_section_openers,
    defense_body_bold,
    figure_width_ratio,
    fit_text_block,
    load_json,
    load_theme,
    mix_color,
    navigation_mode_for,
    resolve_path,
    resolve_font,
    slide_display_title,
    source_line_for_slide,
    text_width,
    wrap_cjk_text,
    write_json,
)
from accessibility import finalize_slide_accessibility, mark_decorative, set_alt_text
from design import (
    DIAGRAM_BRANCH_TEXT,
    Run,
    TextStyle,
    parse_markup,
    relation_arrow_colors,
    relation_arrow_polygon,
    relation_connector_geometry,
    relation_label_box,
    relation_node_boxes,
    relation_route_points,
)
from render_deck_next import DeckRenderer, HEADER_EYEBROW_GAP


PX_PER_INCH = 153.6
EMU_PER_PX = 914400.0 / PX_PER_INCH
PT_PER_PX = 72.0 / PX_PER_INCH
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
MARKUP_RE = re.compile(r"(\*\*.*?\*\*|__.*?__|\[\[.*?\]\])", re.S)


def px(value: float) -> Emu:
    return Emu(int(round(float(value) * EMU_PER_PX)))


def pt(value: float) -> Pt:
    return Pt(float(value) * PT_PER_PX)


def align_value(name: str) -> PP_ALIGN:
    return {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
        "justify": PP_ALIGN.JUSTIFY,
    }.get(name, PP_ALIGN.LEFT)


def anchor_value(name: str) -> MSO_ANCHOR:
    return {
        "top": MSO_ANCHOR.TOP,
        "middle": MSO_ANCHOR.MIDDLE,
        "bottom": MSO_ANCHOR.BOTTOM,
    }.get(name, MSO_ANCHOR.TOP)


def strip_markup(text: str) -> str:
    return MARKUP_RE.sub(lambda match: match.group(0)[2:-2] if match.group(0)[:2] in {"**", "__", "[["} else match.group(0), text or "")


def defense_line_text(text: str, width_px: float, size_px: float, font_paths: Iterable[str] | None = None) -> str:
    """Insert explicit CJK line breaks so PowerPoint does not repack a measured line."""
    if not text:
        return ""
    draw = ImageDraw.Draw(Image.new("RGB", (8, 8)))
    candidates = list(font_paths or []) or [
        r"C:\Windows\Fonts\Noto Sans SC Bold (TrueType).otf",
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        "/System/Library/Fonts/PingFang.ttc",
    ]
    font = resolve_font(candidates, max(1, int(round(size_px))))
    lines = wrap_cjk_text(draw, strip_markup(text), font, max(20, int(width_px)))
    return "\n".join(lines) if lines else ""


def family_from_path(path: str | Path | None, fallback: str) -> str:
    if not path:
        return fallback
    name = Path(path).name.lower()
    if "notosanssc" in name or "noto sans sc" in name:
        return "Noto Sans SC"
    if "sourcehanserifcn" in name:
        return "Source Han Serif CN SemiBold" if "semibold" in name else "Source Han Serif CN"
    if "msyh" in name:
        return "Microsoft YaHei"
    if "simhei" in name:
        return "SimHei"
    if "simsun" in name:
        return "SimSun"
    if "stsong" in name or "stzhongs" in name:
        return "STZhongsong"
    if "pingfang" in name:
        return "PingFang SC"
    if "arial" in name:
        return "Arial"
    if "times" in name:
        return "Times New Roman"
    if "cambria" in name:
        return "Cambria Math"
    return fallback


def first_existing(paths: Iterable[str] | None) -> str | None:
    for item in paths or []:
        path = Path(item)
        if path.exists():
            return str(path)
    return None


def set_cell_border(cell, color: RGBColor, width_pt: float = 0.75) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    width = int(round(width_pt * 12700))
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        line = tc_pr.find(qn(tag))
        if line is None:
            line = parse_xml(f'<{tag.replace(":", ":")} xmlns:a="{NS_A}"/>')
            tc_pr.append(line)
        line.set("w", str(width))
        line.set("cap", "flat")
        line.set("cmpd", "sng")
        line.set("algn", "ctr")
        for child in list(line):
            line.remove(child)
        solid = parse_xml(f'<a:solidFill xmlns:a="{NS_A}"/>')
        srgb = parse_xml(f'<a:srgbClr xmlns:a="{NS_A}" val="{str(color)}"/>')
        solid.append(srgb)
        line.append(solid)
        dash = parse_xml(f'<a:prstDash xmlns:a="{NS_A}" val="solid"/>')
        line.append(dash)


def set_shape_effect(shape, color: tuple[int, int, int], blur_px: float = 8, distance_px: float = 5) -> None:
    sp_pr = shape._element.spPr
    old = sp_pr.find(qn("a:effectLst"))
    if old is not None:
        sp_pr.remove(old)
    effect = parse_xml(
        f'<a:effectLst xmlns:a="{NS_A}">'
        f'<a:outerShdw blurRad="{int(px(blur_px))}" dist="{int(px(distance_px))}" dir="5400000" algn="ctr" rotWithShape="0">'
        f'<a:srgbClr val="{color[0]:02X}{color[1]:02X}{color[2]:02X}"><a:alpha val="18000"/></a:srgbClr>'
        f"</a:outerShdw></a:effectLst>"
    )
    sp_pr.append(effect)


class EditableDeckExporter:
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
        self.theme_path = theme_path.resolve()
        self.output_dir = output_dir.resolve()
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.asset_dir = self.output_dir / "assets"
        self.asset_dir.mkdir(parents=True, exist_ok=True)
        self.deck = load_json(self.deck_path)
        if pages:
            self.page_selection = [page for page in pages if page >= 1]
        elif sample:
            self.page_selection = list(range(1, max(0, sample) + 1))
        else:
            self.page_selection = None
        self.theme = load_theme(self.theme_path)
        self.profile = profile or str(self.deck.get("meta", {}).get("type_profile", "document"))
        apply_type_profile(self.theme, self.profile)
        self.navigation_mode = apply_navigation_mode(
            self.theme,
            navigation_mode_for(self.deck, self.theme),
        )
        self.renderer = DeckRenderer(self.deck_path, self.theme_path, self.output_dir, profile=self.profile)
        self.width = int(self.theme["canvas"]["width"])
        self.height = int(self.theme["canvas"]["height"])
        self.colors = self.theme["_colors"]
        self.layout = self.theme["layout"]
        self.type_scale = self.theme.get("type_scale", {})
        self.body_bold = bool(self.theme.get("text", {}).get("body_bold", False))
        self.body_bold_min_px = float(self.theme.get("text", {}).get("body_bold_min_px", 40))
        self.body_bold_max_px = float(self.theme.get("text", {}).get("body_bold_max_px", 56))
        self.panel_radius = int(self.layout.get("panel_radius", 8))
        self.panel_shadow_alpha = int(self.layout.get("panel_shadow_alpha", 20))
        self.panel_cfg = self.theme.get("panel", {}) or {}
        self.sidebar_mode = str(self.layout.get("sidebar_mode", "always")).strip().lower() or "always"
        self.block_icons = str(self.layout.get("block_icons", "on")).strip().lower() != "off"
        self.corner_marks = bool((self.theme.get("panel", {}) or {}).get("corner_marks", True))
        self.card_index = bool((self.theme.get("panel", {}) or {}).get("card_index", True))
        self.card_highlight = bool((self.theme.get("panel", {}) or {}).get("card_highlight", True))
        self.section_openers = content_section_openers(self.deck)
        if self.sidebar_mode == "off":
            _wide = self.layout.get("content_no_sidebar")
            if isinstance(_wide, (list, tuple)) and len(_wide) == 4:
                self.layout["content"] = [float(value) for value in _wide]
        self.formula_cache: dict[str, etree._Element] = {}
        self.formula_native = 0
        self.formula_fallback = 0
        self.issues: list[dict] = []
        self.text_shape_count = 0
        self.picture_count = 0
        self.native_table_count = 0
        self.native_chart_count = 0
        self.slide_stats: list[dict] = []
        self.background_paths: list[Path] = []
        self.fonts = self.resolve_pptx_fonts()
        self.xsl_path = Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL")
        self._prepare_backgrounds()
        self._prepare_watermark()

    def selected_slides(self) -> list[dict]:
        slides = self.deck.get("slides", [])
        if not self.page_selection:
            return slides
        return [slides[page - 1] for page in self.page_selection if page <= len(slides)]

    def warn(self, page: str, message: str) -> None:
        self.issues.append({"page": page, "message": message})

    def color(self, name: str) -> tuple[int, int, int]:
        return self.colors[name]

    def _defense_font_paths(self, bold: bool) -> list[str]:
        key = "title" if bold else "body"
        return list(self.renderer.theme["fonts"].get(key) or [])

    def rgb(self, name: str) -> RGBColor:
        return RGBColor(*self.color(name))

    def set_shape_gradient(
        self,
        shape,
        start: tuple[int, int, int],
        end: tuple[int, int, int],
        direction: str = "horizontal",
        angle_degrees: float | None = None,
    ) -> None:
        shape.fill.gradient()
        shape.fill.gradient_angle = (
            float(angle_degrees) % 360.0
            if angle_degrees is not None
            else 0
            if str(direction).strip().lower() != "vertical"
            else 90
        )
        stops = shape.fill.gradient_stops
        stops[0].color.rgb = RGBColor(*start)
        stops[0].position = 0.0
        stops[1].color.rgb = RGBColor(*end)
        stops[1].position = 1.0

    def panel_gradient_colors(
        self,
        fill_name: str,
        *,
        enabled: bool = True,
    ) -> tuple[tuple[int, int, int], tuple[int, int, int]] | None:
        style = str(self.panel_cfg.get("fill_style", "solid")).strip().lower()
        if not enabled:
            return None
        if style == "card_fade":
            fade_names = {
                str(item) for item in (self.panel_cfg.get("card_gradient_colors") or ["white", "blue_wash"])
            }
            if fill_name not in fade_names:
                return None
            start = self.color(str(self.panel_cfg.get("card_gradient_from", "blue_bright")))
            end = self.color(str(self.panel_cfg.get("card_gradient_to", "white")))
            return start, end
        if style not in {"soft_gradient", "all_gradient"}:
            return None
        if fill_name == "white":
            if style != "all_gradient":
                return None
            start = self.color("white")
            target = self.color(str(self.panel_cfg.get("white_gradient_to", "blue_wash")))
            amount = max(0.0, min(1.0, float(self.panel_cfg.get("white_gradient_fade", 0.72))))
            return start, mix_color(start, target, amount)
        fade_names = {str(item) for item in (self.panel_cfg.get("fill_gradient_colors") or [])}
        if style != "all_gradient" and fill_name not in fade_names:
            return None
        start = self.color(fill_name)
        target = self.color(str(self.panel_cfg.get("fill_gradient_to", "white")))
        amount = max(0.0, min(1.0, float(self.panel_cfg.get("fill_gradient_fade", 0.86))))
        return start, mix_color(start, target, amount)

    def type_size(self, name: str, fallback: int) -> int:
        return int(self.type_scale.get(name, fallback))

    def header_title_left(self) -> float:
        """Use the same deck-wide title column as the PNG renderer."""
        cached = getattr(self, "_header_title_left", None)
        if cached is not None:
            return cached
        chapter_x = float(self.layout["chapter_x"])
        title_left = float(self.layout["title_x"])
        size = self.type_size("header_title", 48)
        font = self.renderer.fonts.get(TextStyle(bold=True), size)
        draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
        eyebrow_width = 0.0
        for item in self.deck.get("slides", []):
            label = str(item.get("eyebrow", "")).strip()
            if label:
                eyebrow_width = max(eyebrow_width, text_width(draw, label, font))
        if eyebrow_width:
            title_left = max(title_left, chapter_x + eyebrow_width + HEADER_EYEBROW_GAP)
        self._header_title_left = title_left
        return title_left

    @staticmethod
    def image_alt(slide_data: dict) -> str:
        return str(
            slide_data.get("image_alt")
            or slide_data.get("image_role")
            or slide_data.get("image_caption")
            or slide_data.get("image_label")
            or "内容配图"
        ).strip()

    def resolve_pptx_fonts(self) -> dict[str, str]:
        configured = self.theme.get("pptx_fonts", {})
        theme_fonts = self.theme.get("fonts", {})
        title_path = first_existing(theme_fonts.get("title"))
        body_path = first_existing(theme_fonts.get("body"))
        display_path = first_existing(theme_fonts.get("display"))
        formula_path = first_existing(theme_fonts.get("formula"))
        latin_path = first_existing(theme_fonts.get("latin"))
        return {
            "title": str(configured.get("title") or family_from_path(title_path, "Microsoft YaHei")),
            "brand": str(configured.get("brand") or family_from_path(first_existing(theme_fonts.get("brand")), configured.get("title") or "Microsoft YaHei")),
            "body": str(configured.get("body") or family_from_path(body_path, "Microsoft YaHei")),
            "display": str(configured.get("display") or family_from_path(display_path, "STZhongsong")),
            "formula": str(configured.get("formula") or "Cambria Math"),
            "latin": str(configured.get("latin") or family_from_path(latin_path, "Arial")),
        }

    def _prepare_backgrounds(self) -> None:
        for index, slide in enumerate(self.selected_slides(), start=1):
            image = self.renderer.background(index, cover=slide.get("layout") == "cover")
            path = self.asset_dir / f"background_{index:02d}.png"
            image.save(path, optimize=True)
            self.background_paths.append(path)
        cover_bg = self.renderer.background(1, cover=True)
        cover_path = self.asset_dir / "background_cover.png"
        cover_bg.save(cover_path, optimize=True)
        self.cover_background = cover_path

    def _prepare_watermark(self) -> None:
        self.watermark_path: Path | None = None
        emblem = self.theme.get("brand", {}).get("emblem")
        if not emblem:
            return
        path = resolve_path(emblem, Path(self.theme["_dir"]))
        if path is None or not path.exists():
            return
        image = Image.open(path).convert("RGBA")
        opacity = float(self.theme.get("cover", {}).get("watermark_alpha", 0.09))
        alpha = image.getchannel("A").point(lambda value: int(value * opacity))
        image.putalpha(alpha)
        target = self.asset_dir / "watermark_emblem.png"
        image.save(target, optimize=True)
        self.watermark_path = target

    def apply_font(self, run, family: str, size_px: float, color: RGBColor, bold: bool = False, underline: bool = False) -> None:
        size_px = clamp_defense_size_px(size_px, self.profile)
        bold = bold or defense_body_bold(self.theme, size_px, self.profile)
        run.font.name = family
        point_size = float(size_px) * PT_PER_PX
        if self.profile == "defense" and 23.0 < point_size < 24.0:
            point_size = 24.0
        run.font.size = Pt(point_size)
        run.font.bold = bold
        run.font.underline = underline
        run.font.color.rgb = color
        r_pr = run._r.get_or_add_rPr()
        r_pr.set("lang", "zh-CN")
        for tag in ("a:ea", "a:cs"):
            node = r_pr.find(qn(tag))
            if node is None:
                node = parse_xml(f'<{tag} xmlns:a="{NS_A}"/>')
                r_pr.append(node)
            node.set("typeface", family)

    def add_text(
        self,
        slide,
        box: tuple[float, float, float, float],
        text: str,
        size_px: float,
        color_name: str = "ink",
        bold: bool = False,
        family: str | None = None,
        align: str = "left",
        valign: str = "top",
        line_ratio: float = 1.22,
        margin_px: float = 0,
        accent_color: tuple[int, int, int] | None = None,
        name: str = "",
        transparent: bool = True,
        body_floor: bool = False,
        wrap_explicit: bool = True,
        type_floor_px: float | None = None,
    ):
        # Nothing in a projected deck may render below the smallest step of the
        # type ladder. Diagram labels, formula notes and cover metadata used to
        # carry hard-coded 13-18 px sizes; clamp here instead of hunting down
        # every call site.
        floor = (
            int(type_floor_px)
            if type_floor_px is not None
            else int(self.type_size("micro", 16))
        )
        if isinstance(size_px, (int, float)) and size_px < floor:
            size_px = floor
        if body_floor:
            size_px = clamp_defense_body_size_px(size_px, self.profile)
        left, top, width, height = box
        if (
            self.body_bold
            and not bold
            and self.body_bold_min_px <= float(size_px) <= self.body_bold_max_px
        ):
            bold = True
        shape = slide.shapes.add_textbox(px(left), px(top), px(width), px(height))
        if name:
            shape.name = name
        self.text_shape_count += 1
        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = bool(wrap_explicit)
        tf.margin_left = px(margin_px)
        tf.margin_right = px(margin_px)
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = anchor_value(valign)
        tf.auto_size = None
        rendered_text = str(text)
        if self.profile == "defense" and rendered_text and wrap_explicit:
            rendered_text = defense_line_text(
                rendered_text,
                max(20, width - 2 * margin_px),
                clamp_defense_size_px(size_px, self.profile, body_floor=body_floor),
                self._defense_font_paths(bold or (self.body_bold and self.body_bold_min_px <= float(size_px) <= self.body_bold_max_px)),
            )
        paragraphs = rendered_text.split("\n") if rendered_text else [""]
        font_family = family or self.fonts["body"]
        for paragraph_index, paragraph_text in enumerate(paragraphs):
            paragraph = tf.paragraphs[0] if paragraph_index == 0 else tf.add_paragraph()
            paragraph.alignment = align_value(align)
            paragraph.line_spacing = 1.0 if self.profile == "defense" else line_ratio
            paragraph.space_before = Pt(0)
            paragraph.space_after = Pt(0)
            runs = parse_markup(paragraph_text)
            if not runs:
                runs = [Run("", TextStyle())]
            for item in runs:
                if not item.text:
                    continue
                run = paragraph.add_run()
                run.text = item.text
                run_color = self.rgb(color_name)
                if item.style.accent:
                    run_color = RGBColor(*accent_color) if accent_color else self.rgb("red")
                elif item.style.bold:
                    run_color = self.rgb("blue_dark") if color_name == "ink" else self.rgb(color_name)
                self.apply_font(
                    run,
                    font_family,
                    size_px,
                    run_color,
                    bold=bold or item.style.bold,
                    underline=item.style.underline,
                )
        return shape

    def measure(
        self,
        text: str,
        kind: str,
        width: int,
        height: int,
        max_size: int,
        min_size: int,
        max_lines: int | None = None,
        line_ratio: float = 1.22,
        bold: bool = False,
    ) -> tuple[int, list[str]]:
        max_size = int(round(clamp_defense_size_px(max_size, self.profile)))
        min_size = int(round(clamp_defense_size_px(min_size, self.profile)))
        if min_size > max_size:
            min_size = max_size
        draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
        font_key = "title" if bold else kind
        fonts = self.renderer.theme["fonts"].get(font_key) or self.renderer.theme["fonts"]["body"]
        font, lines, _, fits = fit_text_block(
            draw,
            strip_markup(text),
            fonts,
            width,
            height,
            max_size,
            min_size,
            max_lines=max_lines,
            line_ratio=line_ratio,
        )
        return int(font.size), lines

    def add_fit_text(
        self,
        slide,
        box: tuple[float, float, float, float],
        text: str,
        kind: str,
        max_size: int,
        min_size: int,
        color_name: str = "ink",
        bold: bool = False,
        align: str = "left",
        valign: str = "top",
        max_lines: int | None = None,
        line_ratio: float = 1.22,
        family: str | None = None,
        accent_color: tuple[int, int, int] | None = None,
        name: str = "",
        body_floor: bool = False,
        wrap_explicit: bool = True,
        type_floor_px: float | None = None,
    ):
        size, _ = self.measure(
            text,
            kind,
            max(20, int(box[2])),
            max(20, int(box[3])),
            max_size,
            min_size,
            max_lines=max_lines,
            line_ratio=line_ratio,
            bold=bold,
        )
        return self.add_text(
            slide,
            box,
            text,
            size,
            color_name=color_name,
            bold=bold,
            family=family or (self.fonts["title"] if kind == "title" else self.fonts["body"]),
            align=align,
            valign=valign,
            line_ratio=line_ratio,
            accent_color=accent_color,
            name=name,
            body_floor=body_floor,
            wrap_explicit=wrap_explicit,
            type_floor_px=type_floor_px,
        )

    def add_panel(
        self,
        slide,
        box: tuple[float, float, float, float],
        fill_name: str = "white",
        line_name: str | None = "line",
        radius_px: int | None = None,
        line_width_px: float = 1,
        shadow: bool = False,
        name: str = "",
        gradient: bool = True,
        fill_to: str | None = None,
        gradient_direction: str | None = None,
    ):
        left, top, width, height = box
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(left), px(top), px(width), px(height))
        if name:
            shape.name = name
        radius = self.panel_radius if radius_px is None else radius_px
        if radius <= 0:
            shape.adjustments[0] = 0
        else:
            shape.adjustments[0] = min(0.5, float(radius) / max(1.0, min(width, height)))
        large_enough = (
            width >= float(self.panel_cfg.get("card_gradient_min_width", 300))
            and height >= float(self.panel_cfg.get("card_gradient_min_height", 80))
        )
        gradient_colors = (
            (self.color(fill_name), self.color(fill_to))
            if fill_to
            else self.panel_gradient_colors(
                fill_name,
                enabled=gradient and large_enough,
            )
        )
        if gradient_colors:
            start, end = gradient_colors
            self.set_shape_gradient(
                shape,
                start,
                end,
                gradient_direction
                or str(self.panel_cfg.get("gradient_direction", "horizontal")),
            )
            if not name:
                shape.name = "soft_gradient_panel"
        else:
            shape.fill.solid()
            shape.fill.fore_color.rgb = self.rgb(fill_name)
        if line_name:
            shape.line.color.rgb = self.rgb(line_name)
            shape.line.width = pt(line_width_px)
        else:
            shape.line.fill.background()
        shape.shadow.inherit = False
        if shadow:
            set_shape_effect(shape, self.color("blue_dark"), blur_px=8, distance_px=5)
        return shape

    def add_rect(
        self,
        slide,
        box: tuple[float, float, float, float],
        fill_name: str,
        line_name: str | None = None,
        name: str = "",
        fill_to: str | None = None,
        gradient_direction: str | None = None,
    ):
        left, top, width, height = box
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(left), px(top), px(width), px(height))
        if name:
            shape.name = name
        if fill_to:
            self.set_shape_gradient(
                shape,
                self.color(fill_name),
                self.color(fill_to),
                gradient_direction or str(self.panel_cfg.get("gradient_direction", "horizontal")),
            )
        else:
            shape.fill.solid()
            shape.fill.fore_color.rgb = self.rgb(fill_name)
        if line_name:
            shape.line.color.rgb = self.rgb(line_name)
            shape.line.width = pt(1)
        else:
            shape.line.fill.background()
        shape.shadow.inherit = False
        return shape

    def add_line(
        self,
        slide,
        start: tuple[float, float],
        end: tuple[float, float],
        color_name: str = "line",
        width_px: float = 2,
        arrow: bool = False,
        dash: str | None = None,
        arrow_size: str = "sm",
        name: str = "",
    ):
        shape = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, px(start[0]), px(start[1]), px(end[0]), px(end[1]))
        if name:
            shape.name = name
        shape.line.color.rgb = self.rgb(color_name)
        shape.line.width = pt(width_px)
        shape.shadow.inherit = False
        ln = shape.line._get_or_add_ln()
        if dash:
            existing = ln.find(qn("a:prstDash"))
            if existing is not None:
                ln.remove(existing)
            ln.append(parse_xml(f'<a:prstDash xmlns:a="{NS_A}" val="{dash}"/>'))
        if arrow:
            size = arrow_size if arrow_size in {"sm", "med", "lg"} else "sm"
            tail = parse_xml(f'<a:tailEnd xmlns:a="{NS_A}" type="triangle" w="{size}" len="{size}"/>')
            ln.append(tail)
        return shape

    def add_ellipse(
        self,
        slide,
        box: tuple[float, float, float, float],
        fill_name: str | None = None,
        line_name: str | None = None,
        line_width_px: float = 2,
        name: str = "",
        fill_to: str | None = None,
    ):
        left, top, width, height = box
        shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, px(left), px(top), px(width), px(height))
        if name:
            shape.name = name
        if fill_name:
            if fill_to:
                self.set_shape_gradient(shape, self.color(fill_name), self.color(fill_to), "vertical")
            else:
                shape.fill.solid()
                shape.fill.fore_color.rgb = self.rgb(fill_name)
        else:
            shape.fill.background()
        if line_name:
            shape.line.color.rgb = self.rgb(line_name)
            shape.line.width = pt(line_width_px)
        else:
            shape.line.fill.background()
        shape.shadow.inherit = False
        return shape

    def add_polygon(
        self,
        slide,
        points: Sequence[tuple[float, float]],
        fill_name: str,
        line_name: str | None = None,
        line_width_px: float = 0,
        name: str = "",
        fill_to: str | None = None,
        gradient_direction: str | None = None,
    ):
        if len(points) < 3:
            return None
        x0, y0 = points[0]
        builder = slide.shapes.build_freeform(x0, y0, scale=EMU_PER_PX)
        builder.add_line_segments(points[1:], close=True)
        shape = builder.convert_to_shape()
        if name:
            shape.name = name
        if fill_to:
            self.set_shape_gradient(
                shape,
                self.color(fill_name),
                self.color(fill_to),
                gradient_direction or str(self.panel_cfg.get("gradient_direction", "horizontal")),
            )
        else:
            shape.fill.solid()
            shape.fill.fore_color.rgb = self.rgb(fill_name)
        if line_name:
            shape.line.color.rgb = self.rgb(line_name)
            shape.line.width = pt(line_width_px)
        else:
            shape.line.fill.background()
        shape.shadow.inherit = False
        return shape

    def add_picture_fit(
        self,
        slide,
        path: str | Path | None,
        box: tuple[float, float, float, float],
        align: str = "center",
        valign: str = "middle",
        name: str = "",
        alt_text: str = "",
        decorative: bool = False,
    ):
        if not path:
            return None
        image_path = Path(path)
        if not image_path.exists():
            return None
        with Image.open(image_path) as image:
            image_width, image_height = image.size
        left, top, width, height = box
        scale = min(width / image_width, height / image_height)
        draw_width = image_width * scale
        draw_height = image_height * scale
        if align == "left":
            target_left = left
        elif align == "right":
            target_left = left + width - draw_width
        else:
            target_left = left + (width - draw_width) / 2
        if valign == "top":
            target_top = top
        elif valign == "bottom":
            target_top = top + height - draw_height
        else:
            target_top = top + (height - draw_height) / 2
        shape = slide.shapes.add_picture(str(image_path), px(target_left), px(target_top), px(draw_width), px(draw_height))
        if name:
            shape.name = name
        if decorative:
            mark_decorative(shape)
        else:
            set_alt_text(shape, alt_text or "内容配图")
        self.picture_count += 1
        return shape

    def add_corner_marks(self, slide, box: tuple[float, float, float, float], color_name: str = "line") -> None:
        if not getattr(self, "corner_marks", True):
            return
        left, top, width, height = box
        inset = 8
        length = 16
        x0, y0 = left + inset, top + inset
        x1, y1 = left + width - inset, top + height - inset
        segments = [
            ((x0, y0), (x0 + length, y0)),
            ((x0, y0), (x0, y0 + length)),
            ((x1, y0), (x1 - length, y0)),
            ((x1, y0), (x1, y0 + length)),
            ((x0, y1), (x0 + length, y1)),
            ((x0, y1), (x0, y1 - length)),
            ((x1, y1), (x1 - length, y1)),
            ((x1, y1), (x1, y1 - length)),
        ]
        for start, end in segments:
            self.add_line(slide, start, end, color_name, 1.5)

    def add_native_chart(self, slide, box: tuple[float, float, float, float], chart_data: dict, name: str = "native_chart"):
        values = [float(value) for value in chart_data.get("values", []) if value is not None]
        if not values:
            return None
        kind = str(chart_data.get("type", "histogram")).strip().lower()
        if kind in {"curve", "line"}:
            chart_type = XL_CHART_TYPE.LINE_MARKERS
        elif kind in {"bar", "column", "histogram"}:
            chart_type = XL_CHART_TYPE.COLUMN_CLUSTERED
        else:
            return None

        data = CategoryChartData()
        data.categories = [str(index + 1) for index in range(len(values))]
        data.add_series(str(chart_data.get("series_name", "数值")), values)
        left, top, width, height = box
        graphic = slide.shapes.add_chart(
            chart_type,
            px(left),
            px(top),
            px(width),
            px(height),
            data,
        )
        graphic.name = name
        set_alt_text(
            graphic,
            str(
                chart_data.get("alt_text")
                or chart_data.get("caption")
                or chart_data.get("annotation")
                or f"图表：{chart_data.get('series_name', '数据')}"
            ),
        )

        chart = graphic.chart
        chart.has_legend = False
        chart.has_title = False
        chart.font.name = self.fonts["body"]
        chart.font.size = pt(clamp_defense_size_px(15, self.profile))
        chart.font.color.rgb = self.rgb("muted")

        plot = chart.plots[0]
        plot.has_data_labels = False
        maximum = max(values) or 1.0
        highlight = int(chart_data.get("highlight", max(range(len(values)), key=lambda index: values[index])))
        highlight = max(0, min(len(values) - 1, highlight))
        axis_color = self.rgb("blue_dark")
        grid_color = mix_color(self.color("line"), self.color("paper"), 0.55)

        for axis in (chart.category_axis, chart.value_axis):
            axis.has_major_gridlines = True
            axis.has_minor_gridlines = False
            axis.major_tick_mark = XL_TICK_MARK.NONE
            axis.minor_tick_mark = XL_TICK_MARK.NONE
            axis.tick_label_position = XL_TICK_LABEL_POSITION.NONE
            axis.format.line.color.rgb = axis_color
            axis.format.line.width = pt(1.5)
            axis.major_gridlines.format.line.color.rgb = RGBColor(*grid_color)
            axis.major_gridlines.format.line.width = pt(1)

        chart.value_axis.minimum_scale = 0
        chart.value_axis.maximum_scale = maximum * 1.12
        series = plot.series[0]
        if chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED:
            plot.gap_width = 38
            plot.vary_by_categories = False
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = self.rgb("blue_mid")
            series.format.line.fill.background()
            point = series.points[highlight]
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = self.rgb("red")
            point.format.line.fill.background()
        else:
            series.smooth = False
            series.format.line.color.rgb = self.rgb("blue_mid")
            series.format.line.width = pt(3)
            series.marker.style = XL_MARKER_STYLE.CIRCLE
            series.marker.size = 6
            series.marker.format.fill.solid()
            series.marker.format.fill.fore_color.rgb = self.rgb("blue")
            series.marker.format.line.color.rgb = self.rgb("white")
            series.marker.format.line.width = pt(1.5)
            point = series.points[highlight]
            point.marker.style = XL_MARKER_STYLE.CIRCLE
            point.marker.size = 8
            point.marker.format.fill.solid()
            point.marker.format.fill.fore_color.rgb = self.rgb("red")
            point.marker.format.line.color.rgb = self.rgb("white")
            point.marker.format.line.width = pt(1.5)

        plot_area = chart._chartSpace.chart.plotArea
        old_layout = plot_area.find(qn("c:layout"))
        if old_layout is not None:
            plot_area.remove(old_layout)
        layout = parse_xml(
            '<c:layout xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            '<c:manualLayout>'
            '<c:layoutTarget val="inner"/>'
            '<c:xMode val="edge"/><c:yMode val="edge"/>'
            '<c:x val="0.035"/><c:y val="0.04"/><c:w val="0.95"/><c:h val="0.93"/>'
            '</c:manualLayout>'
            '</c:layout>'
        )
        plot_area.insert(0, layout)
        return chart

    def add_chart_card(
        self,
        slide,
        box: tuple[float, float, float, float],
        chart_data: dict,
        label: str = "",
        caption: str = "",
        badge: str = "",
        name: str = "native_chart",
    ):
        left, top, width, height = box
        meta_bits: list[str] = []
        if chart_data.get("unit"):
            meta_bits.append(f"单位：{chart_data['unit']}")
        if chart_data.get("sample_size") not in (None, ""):
            meta_bits.append(f"n={chart_data['sample_size']}")
        if chart_data.get("condition"):
            meta_bits.append(str(chart_data["condition"]))
        if chart_data.get("error_bar"):
            meta_bits.append("含误差条")
        if meta_bits:
            meta_text = " · ".join(meta_bits)
            caption = f"{caption}　{meta_text}".strip() if caption else meta_text
        self.add_panel(
            slide,
            box,
            "white",
            "line",
            radius_px=12,
            line_width_px=1.2,
            shadow=True,
            gradient=False,
        )
        if label:
            label_width = max(120.0, len(label) * 20.0 + 34.0)
            label_box = (left + 18, top - 15, min(width - 36, label_width), 34)
            self.add_panel(slide, label_box, "blue_pale", "line", radius_px=5, line_width_px=1, shadow=True)
            self.add_text(
                slide,
                (label_box[0] + 12, label_box[1] + 4, label_box[2] - 20, label_box[3] - 7),
                label,
                20,
                "blue_dark",
                bold=True,
                valign="middle",
                line_ratio=1.0,
                name="chart_label",
            )
        if badge:
            badge_width = max(70.0, len(badge) * 18.0 + 22.0)
            self.add_panel(slide, (left + width - 18 - badge_width, top + 13, badge_width, 28), "blue_pale", None, radius_px=4, line_width_px=0)
            self.add_text(
                slide,
                (left + width - 18 - badge_width, top + 14, badge_width, 25),
                badge,
                16,
                "blue_dark",
                bold=True,
                align="center",
                valign="middle",
                line_ratio=1.0,
                name="chart_badge",
            )

        caption_height = 34 if caption else 0
        chart_box = (left + 58, top + 32, width - 84, height - 32 - 52 - max(0, caption_height - 34))
        native_chart = self.add_native_chart(slide, chart_box, chart_data, name=name)
        x_label = str(chart_data.get("x_label", "")).strip()
        if x_label:
            self.add_text(
                slide,
                (left + width - 280, top + height - (76 if caption else 52), 254, 26),
                x_label,
                16,
                "muted",
                align="right",
                line_ratio=1.0,
                name="chart_x_label",
            )
        y_label = str(chart_data.get("y_label", "")).strip()
        if y_label:
            self.add_fit_text(
                slide,
                (left + 66, top + 37, max(120, width - 260), 24),
                y_label,
                "body",
                16,
                12,
                "muted",
                max_lines=1,
                line_ratio=1.0,
                name="chart_y_label",
            )

        values = [float(value) for value in chart_data.get("values", []) if value is not None]
        if values:
            highlight = int(chart_data.get("highlight", max(range(len(values)), key=lambda index: values[index])))
            highlight = max(0, min(len(values) - 1, highlight))
            plot_left = left + 58
            plot_right = left + width - 26
            slot = (plot_right - plot_left) / max(1, len(values))
            highlight_x = plot_left + highlight * slot + slot / 2
            self.add_line(
                slide,
                (highlight_x, top + 32),
                (highlight_x, top + height - 52),
                "red",
                1.5,
                dash="dash",
                name="chart_highlight",
            )
            annotation = str(chart_data.get("annotation", "")).strip()
            if annotation:
                annotation_left = min(highlight_x + 12, plot_right - 48)
                self.add_text(
                    slide,
                    (annotation_left, top + 35, 48, 26),
                    annotation,
                    17,
                    "red",
                    bold=True,
                    line_ratio=1.0,
                    name="chart_annotation",
                )

        if caption:
            self.add_rect(slide, (left + 2, top + height - 36, width - 4, 34), "blue_wash")
            self.add_fit_text(
                slide,
                (left + 16, top + height - 30, width - 32, 22),
                caption,
                "body",
                17,
                12,
                "muted",
                max_lines=1,
                line_ratio=1.0,
                name="chart_caption",
            )
        self.add_corner_marks(slide, box, "line")
        return native_chart

    def add_brand(self, slide, box: tuple[float, float, float, float], align: str = "center", name: str = "brand_lockup") -> None:
        lockup = self.theme.get("brand", {}).get("lockup_path")
        if lockup and Path(lockup).exists():
            fallback = str(
                self.theme.get("brand", {}).get("fallback_text")
                or self.deck.get("meta", {}).get("institution")
                or "院校"
            )
            self.add_picture_fit(
                slide,
                lockup,
                box,
                align=align,
                name=name,
                alt_text=f"{fallback}校名与校徽",
            )
            return
        fallback = str(self.theme.get("brand", {}).get("fallback_text") or self.deck.get("meta", {}).get("institution") or "")
        if not fallback:
            return
        self.add_fit_text(
            slide,
            box,
            fallback,
            "title",
            46,
            18,
            "blue",
            bold=True,
            align=align,
            valign="middle",
            max_lines=2,
            line_ratio=1.0,
            family=self.fonts.get("brand") or None,
            name=name,
        )

    def add_icon(self, *args, **kwargs):
        if not getattr(self, "block_icons", True):
            return None
        return self._add_icon(*args, **kwargs)

    def _add_icon(
        self,
        slide,
        name: str,
        box: tuple[float, float, float, float],
        color_name: str = "white",
        stroke_px: float = 2,
    ) -> None:
        left, top, width, height = box
        cx, cy = left + width / 2, top + height / 2
        s = min(width, height)
        if name in {"clock"}:
            self.add_ellipse(slide, (left, top, width, height), None, color_name, stroke_px)
            self.add_line(slide, (cx, cy), (cx, top + s * 0.22), color_name, stroke_px)
            self.add_line(slide, (cx, cy), (cx + s * 0.24, cy + s * 0.12), color_name, stroke_px)
            return
        if name in {"target"}:
            self.add_ellipse(slide, (left, top, width, height), None, color_name, stroke_px)
            self.add_ellipse(slide, (left + s * 0.24, top + s * 0.24, s * 0.52, s * 0.52), None, color_name, stroke_px)
            self.add_ellipse(slide, (cx - s * 0.06, cy - s * 0.06, s * 0.12, s * 0.12), color_name, None)
            self.add_line(slide, (cx, top), (cx, top + s * 0.16), color_name, stroke_px)
            self.add_line(slide, (cx, top + s * 0.84), (cx, top + s), color_name, stroke_px)
            self.add_line(slide, (left, cy), (left + s * 0.16, cy), color_name, stroke_px)
            self.add_line(slide, (left + s * 0.84, cy), (left + s, cy), color_name, stroke_px)
            return
        if name in {"alert"}:
            self.add_polygon(slide, [(cx, top), (left + width, top + height), (left, top + height)], color_name)
            self.add_line(slide, (cx, top + s * 0.35), (cx, top + s * 0.68), "white", max(2, stroke_px))
            self.add_ellipse(slide, (cx - s * 0.05, top + s * 0.77, s * 0.1, s * 0.1), "white", None)
            return
        if name in {"check"}:
            self.add_ellipse(slide, (left, top, width, height), None, color_name, stroke_px)
            self.add_line(slide, (left + s * 0.24, cy), (left + s * 0.43, top + s * 0.68), color_name, stroke_px)
            self.add_line(slide, (left + s * 0.43, top + s * 0.68), (left + s * 0.78, top + s * 0.3), color_name, stroke_px)
            return
        if name in {"layers"}:
            for index in range(3):
                y = top + s * (0.12 + index * 0.27)
                self.add_polygon(slide, [(cx, y), (left + width * 0.9, y + s * 0.16), (cx, y + s * 0.32), (left + width * 0.1, y + s * 0.16)], "none" if False else color_name)
            return
        if name in {"nodes"}:
            points = [(left + s * 0.2, top + s * 0.68), (left + s * 0.52, top + s * 0.3), (left + s * 0.82, top + s * 0.68)]
            self.add_line(slide, points[0], points[1], color_name, stroke_px)
            self.add_line(slide, points[1], points[2], color_name, stroke_px)
            for x, y in points:
                self.add_ellipse(slide, (x - s * 0.1, y - s * 0.1, s * 0.2, s * 0.2), "white", color_name, stroke_px)
            return
        if name in {"histogram", "chart"}:
            base = top + height * 0.86
            for index, height_ratio in enumerate((0.3, 0.55, 0.8, 0.45)):
                bar_width = width * 0.15
                bar_height = height * height_ratio
                self.add_rect(slide, (left + width * (0.1 + index * 0.21), base - bar_height, bar_width, bar_height), color_name)
            self.add_line(slide, (left + width * 0.05, base), (left + width * 0.95, base), color_name, stroke_px)
            return
        if name in {"route"}:
            points = [(left + s * 0.12, top + s * 0.72), (left + s * 0.42, top + s * 0.72), (left + s * 0.42, top + s * 0.3), (left + s * 0.82, top + s * 0.3)]
            self.add_line(slide, points[0], points[1], color_name, stroke_px)
            self.add_line(slide, points[1], points[2], color_name, stroke_px)
            self.add_line(slide, points[2], points[3], color_name, stroke_px)
            for x, y in (points[0], points[3]):
                self.add_ellipse(slide, (x - s * 0.08, y - s * 0.08, s * 0.16, s * 0.16), color_name, None)
            return
        if name in {"database"}:
            shape = slide.shapes.add_shape(MSO_SHAPE.CAN, px(left), px(top), px(width), px(height))
            shape.fill.background()
            shape.line.color.rgb = self.rgb(color_name)
            shape.line.width = pt(stroke_px)
            shape.shadow.inherit = False
            return
        if name in {"shield"}:
            self.add_polygon(slide, [(cx, top), (left + width * 0.82, top + height * 0.2), (left + width * 0.72, top + height * 0.72), (cx, top + height), (left + width * 0.28, top + height * 0.72), (left + width * 0.18, top + height * 0.2)], color_name)
            return
        if name in {"signal"}:
            for index, height_ratio in enumerate((0.26, 0.45, 0.64, 0.82)):
                self.add_line(slide, (left + s * (0.15 + index * 0.22), top + s * 0.9), (left + s * (0.15 + index * 0.22), top + s * (0.9 - height_ratio)), color_name, stroke_px)
            return
        if name in {"mountain"}:
            self.add_polygon(slide, [(left, top + height * 0.85), (left + width * 0.34, top + height * 0.2), (left + width * 0.55, top + height * 0.58), (left + width * 0.72, top + height * 0.36), (left + width, top + height * 0.85)], color_name)
            return
        if name in {"matrix"}:
            for row in range(2):
                for column in range(2):
                    self.add_rect(slide, (left + s * (0.16 + column * 0.38), top + s * (0.16 + row * 0.38), s * 0.26, s * 0.26), color_name)
            return
        if name in {"chip"}:
            self.add_rect(slide, (left + s * 0.22, top + s * 0.22, s * 0.56, s * 0.56), color_name)
            for index in range(3):
                offset = s * (0.32 + index * 0.16)
                self.add_line(slide, (left + offset, top + s * 0.08), (left + offset, top + s * 0.22), color_name, stroke_px)
                self.add_line(slide, (left + offset, top + s * 0.78), (left + offset, top + s * 0.92), color_name, stroke_px)
                self.add_line(slide, (left + s * 0.08, top + offset), (left + s * 0.22, top + offset), color_name, stroke_px)
                self.add_line(slide, (left + s * 0.78, top + offset), (left + s * 0.92, top + offset), color_name, stroke_px)
            return
        if name in {"atom"}:
            self.add_ellipse(slide, (left, top + s * 0.28, width, s * 0.44), None, color_name, stroke_px)
            self.add_ellipse(slide, (left + s * 0.28, top, s * 0.44, height), None, color_name, stroke_px)
            self.add_ellipse(slide, (cx - s * 0.08, cy - s * 0.08, s * 0.16, s * 0.16), color_name, None)
            return
        if name in {"photon"}:
            self.add_line(slide, (left + s * 0.4, top + s * 0.82), (left + s * 0.58, top + s * 0.18), color_name, stroke_px)
            self.add_ellipse(slide, (left + s * 0.52, top + s * 0.08, s * 0.2, s * 0.2), color_name, None)
            self.add_line(slide, (left + s * 0.18, top + s * 0.68), (left + s * 0.34, top + s * 0.58), color_name, stroke_px)
            return
        if name in {"laser"}:
            self.add_line(slide, (left + s * 0.12, top + s * 0.55), (left + s * 0.68, top + s * 0.55), color_name, stroke_px)
            self.add_polygon(slide, [(left + s * 0.88, top + s * 0.55), (left + s * 0.68, top + s * 0.4), (left + s * 0.68, top + s * 0.7)], color_name)
            return
        if name in {"arrow", "next"}:
            self.add_line(slide, (left + s * 0.12, cy), (left + s * 0.78, cy), color_name, stroke_px)
            self.add_polygon(slide, [(left + s * 0.96, cy), (left + s * 0.7, cy - s * 0.16), (left + s * 0.7, cy + s * 0.16)], color_name)
            return
        if name in {"wave"}:
            for index in range(2):
                points = [
                    (left + s * 0.08, top + s * (0.45 + index * 0.18)),
                    (left + s * 0.3, top + s * (0.28 + index * 0.18)),
                    (left + s * 0.5, top + s * (0.58 + index * 0.18)),
                    (left + s * 0.72, top + s * (0.38 + index * 0.18)),
                    (left + s * 0.92, top + s * (0.58 + index * 0.18)),
                ]
                for p1, p2 in zip(points[:-1], points[1:]):
                    self.add_line(slide, p1, p2, color_name, stroke_px)
            return
        self.add_ellipse(slide, (cx - s * 0.14, cy - s * 0.14, s * 0.28, s * 0.28), None, color_name, stroke_px)

    def add_module_header(
        self,
        slide,
        box: tuple[float, float, float, float],
        heading: str,
        index: str = "",
        icon: str = "",
        height: int = 56,
    ) -> int:
        """Draw one measured module header and return its real height.

        The index and the heading use the same title font, size, top and
        middle anchor. A two-digit index gets a measured column instead of a
        fixed 36 px box, which used to clip ``01`` to ``0`` in PowerPoint.
        """
        left, top, width, _ = box
        configured_size = float(self.type_size("block_title", 26))
        minimum_height = int(round(configured_size * 1.1))
        if self.profile == "defense":
            minimum_height = max(minimum_height, 76)
        height = max(int(height), minimum_height)
        header = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(left + 2), px(top + 2), px(width - 4), px(height))
        header.name = "module_header_fill"
        header.adjustments[0] = min(0.5, self.panel_radius / max(1, height))
        header_style = str(self.panel_cfg.get("header_style", "gradient")).strip().lower()
        header_tone = str(self.panel_cfg.get("header_tone", "deep")).strip().lower()
        if header_style == "flat":
            header.fill.solid()
            header.fill.fore_color.rgb = self.rgb("blue")
        elif header_tone == "mist":
            self.set_shape_gradient(
                header,
                self.color(str(self.panel_cfg.get("header_gradient_from", "blue_pale"))),
                self.color(str(self.panel_cfg.get("header_gradient_to", "white"))),
                str(self.panel_cfg.get("gradient_direction", "horizontal")),
            )
            self.add_rect(
                slide,
                (left + 2, top + 2, 7, height),
                str(self.panel_cfg.get("header_accent", "blue")),
                fill_to=str(self.panel_cfg.get("header_gradient_to", "white")),
                gradient_direction="vertical",
                name="module_header_accent",
            )
        else:
            self.set_shape_gradient(
                header,
                self.color(str(self.panel_cfg.get("header_gradient_from", "blue_mid"))),
                self.color(str(self.panel_cfg.get("header_gradient_to", "blue"))),
                str(self.panel_cfg.get("gradient_direction", "horizontal")),
            )
        header.line.fill.background()
        header.shadow.inherit = False
        if header_style == "flat" and getattr(self, "card_highlight", True):
            self.add_rect(slide, (left + 3, top + 3, width - 6, 3), "blue_mid")

        draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
        index_width = 0.0
        if index and getattr(self, "card_index", True):
            configured_font = self.renderer.fonts.get(TextStyle(bold=True), int(round(configured_size)))
            index_width = max(60.0, float(text_width(draw, index, configured_font)) + 20.0)
        text_left = left + 16 + (index_width + 12 if index_width else 0)
        heading_size = configured_size
        if heading:
            available = max(40.0, width - (text_left - left) - (56 if icon else 18))
            heading_size, heading_lines = self.measure(
                heading,
                "title",
                int(available),
                max(20, height - 12),
                int(round(configured_size)),
                max(18, int(round(self.type_size("small", 14)))),
                max_lines=1,
                line_ratio=1.0,
                bold=True,
            )
            if len(heading_lines) > 1:
                self.warn(str(self.current_deck_slide.get("id", "")), f"Module heading overflowed: {heading}")

        text_top = top + 6
        text_height = height - 12
        if index_width:
            self.add_text(
                slide,
                (left + 16, text_top, index_width, text_height),
                index,
                heading_size,
                str(self.panel_cfg.get("header_index", "blue_pale")) if header_tone == "mist" else "blue_pale",
                bold=True,
                family=self.fonts["title"],
                valign="middle",
                line_ratio=1.0,
                name="module_index",
                wrap_explicit=False,
            )
        if heading:
            available = max(40.0, width - (text_left - left) - (56 if icon else 18))
            self.add_text(
                slide,
                (text_left, text_top, available, text_height),
                heading,
                heading_size,
                str(self.panel_cfg.get("header_text", "blue_dark")) if header_tone == "mist" else "white",
                bold=True,
                family=self.fonts["title"],
                valign="middle",
                line_ratio=1.0,
                name="module_heading",
                wrap_explicit=False,
            )
        if icon:
            self.add_icon(
                slide,
                icon,
                (left + width - 48, top + (height - 32) / 2, 32, 32),
                "blue" if header_tone == "mist" else "white",
            )
        return height

    def diagram_index_box(
        self,
        left: float,
        top: float,
        text: str,
        requested_size: int,
    ) -> tuple[tuple[float, float, float, float], int]:
        """Return a measured box for an atomic diagram number.

        A two-digit number at the defense floor is about 44 px wide; the old
        fixed 34-46 px boxes clipped it to a single digit.
        """
        size = int(round(clamp_defense_size_px(requested_size, self.profile)))
        font = self.renderer.fonts.get(TextStyle(bold=True), size)
        draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
        width = max(60.0, float(text_width(draw, text, font)) + 20.0)
        height = max(44.0, float(round(size * 1.3)))
        return (left, top, width, height), size

    def add_bullets(
        self,
        slide,
        bullets: Sequence[str],
        box: tuple[float, float, float, float],
        max_size: int = 19,
        min_size: int = 13,
        gap_px: float = 7,
    ) -> None:
        if not bullets:
            return
        max_size = int(round(clamp_defense_size_px(max_size, self.profile)))
        min_size = int(round(clamp_defense_size_px(min_size, self.profile)))

        max_size = max(max_size, min_size)
        left, top, width, height = box
        chosen = min_size
        for candidate in range(max_size, min_size - 1, -1):
            total = 0
            fits = True
            for bullet in bullets:
                _, lines = self.measure(
                    bullet,
                    "body",
                    max(20, int(width - 26)),
                    max(20, int(height)),
                    candidate,
                    candidate,
                    max_lines=3,
                    line_ratio=1.2,
                )
                if len(lines) > 3:
                    fits = False
                    break
                total += len(lines) * candidate * 1.28 + gap_px
            if fits and total <= height:
                chosen = candidate
                break
        shape = slide.shapes.add_textbox(px(left), px(top), px(width), px(height))
        self.text_shape_count += 1
        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.TOP
        for index, bullet in enumerate(bullets):
            paragraph = tf.paragraphs[0] if index == 0 else tf.add_paragraph()
            paragraph.space_before = Pt(0)
            paragraph.space_after = pt(gap_px)
            paragraph.line_spacing = 1.0 if self.profile == "defense" else 1.18
            marker = paragraph.add_run()
            marker.text = "◆  "
            self.apply_font(marker, self.fonts["body"], chosen, self.rgb("blue_mid"), bold=True)
            run = paragraph.add_run()
            run_text = strip_markup(str(bullet))
            bullet_bold = (
                self.body_bold
                and self.body_bold_min_px <= float(chosen) <= self.body_bold_max_px
            )
            if self.profile == "defense":
                run_text = defense_line_text(run_text, width - 52, chosen, self._defense_font_paths(bullet_bold))
            run.text = run_text
            self.apply_font(
                run,
                self.fonts["title"] if bullet_bold else self.fonts["body"],
                chosen,
                self.rgb("ink"),
                bold=bullet_bold,
            )

    def add_output_strip(self, slide, box: tuple[float, float, float, float], label: str, text: str) -> None:
        left, top, width, height = box
        if not text or width < 80 or height < 28:
            return
        self.add_panel(
            slide,
            box,
            "blue_wash",
            "line",
            radius_px=4,
            line_width_px=1,
            gradient=False,
        )
        label_width = min(max(96, len(strip_markup(label)) * 38 + 20), max(96, width // 3))
        self.add_panel(slide, (left + 1, top + 1, label_width, height - 2), "blue", None, radius_px=3, line_width_px=0)
        self.add_text(slide, (left + 1, top + 1, label_width, height - 2), label, 14, "white", bold=True, align="center", valign="middle")
        self.add_fit_text(
            slide,
            (left + label_width + 12, top + 4, width - label_width - 22, height - 8),
            text,
            "body",
            17,
            12,
            "blue_dark",
            max_lines=2,
            line_ratio=1.12,
        )

    def add_lead(self, slide, text: str, top: int, page_id: str) -> int:
        if not text:
            return top
        draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
        left, _, right, _ = map(int, self.layout["content"])
        base_height = int(self.layout.get("lead_height", 88))
        label = str(self.layout.get("lead_label", "")).strip()
        if label:
            label_width = min(int(self.layout.get("lead_label_width", 168)), max(120, (right - left) // 3))
            text_left = left + label_width + 24
            text_width_available = right - text_left - 24
        else:
            label_width = 0
            text_left = left + 96
            text_width_available = right - left - 118

        lead_size = self.type_size("lead", 38)
        minimum_size = max(22, self.type_size("body", 22) if self.profile == "defense" else 22)
        measured_size, lead_lines = self.measure(
            text,
            "title",
            text_width_available,
            max(base_height, int(lead_size * 3)),
            lead_size,
            minimum_size,
            max_lines=2,
            line_ratio=1.2,
            bold=True,
        )
        if len(lead_lines) > 2:
            self.warn(page_id, "Lead text overflowed even after reserving two lines.")
        blocks = max(1, len(lead_lines))
        line_advance = measured_size * 1.2
        height = max(base_height, int(round(blocks * line_advance + 24)))

        if self.layout.get("lead_panel", True):
            panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(left), px(top), px(right - left), px(height))
            panel.adjustments[0] = min(0.5, self.panel_radius / max(1, height))
            panel.fill.solid()
            panel.fill.fore_color.rgb = self.rgb(str(self.panel_cfg.get("lead_fill", "white")))
            panel.line.color.rgb = self.rgb("line")
            panel.line.width = pt(1)
            panel.shadow.inherit = False
        if label:
            self.add_polygon(
                slide,
                [
                    (left, top + 1),
                    (left + label_width - 18, top + 1),
                    (left + label_width, top + height // 2),
                    (left + label_width - 18, top + height - 1),
                    (left, top + height - 1),
                ],
                "blue",
            )
            self.add_text(
                slide,
                (left + 8, top + 8, label_width - 30, height - 16),
                label,
                self.type_size("lead_label", 22),
                "white",
                bold=True,
                family=self.fonts["title"],
                align="center",
                valign="middle",
                name="lead_label",
                wrap_explicit=False,
            )
        else:
            self.add_rect(slide, (left, top + 1, 9, height - 2), "blue")
            self.add_chevrons(slide, left + 28, top + 26, 34, "blue_mid", count=2)
        self.add_text(
            slide,
            (text_left, top + 12, text_width_available, height - 24),
            text,
            measured_size,
            "ink",
            bold=True,
            family=self.fonts["title"],
            line_ratio=1.2,
            name="lead_claim",
        )
        return top + height + (20 if self.layout.get("lead_panel", True) else 12)

    def add_chevrons(self, slide, left: float, top: float, size: float, color_name: str, count: int = 2) -> None:
        for index in range(count):
            base = left + index * (size * 0.58)
            self.add_line(slide, (base, top), (base + size * 0.42, top + size / 2), color_name, max(3, size / 8))
            self.add_line(slide, (base + size * 0.42, top + size / 2), (base, top + size), color_name, max(3, size / 8))

    def add_takeaway(self, slide, text: str, page_id: str) -> None:
        if not text:
            return
        left = int(self.layout["content"][0])
        right = int(self.layout["content"][2])
        top = int(self.layout["takeaway_top"])
        height = int(self.layout["takeaway_height"])
        takeaway_style = str(self.panel_cfg.get("takeaway_style", "deep")).strip().lower()
        if takeaway_style == "pale":
            self.add_panel(
                slide,
                (left, top, right - left, height),
                str(self.panel_cfg.get("takeaway_fill", "blue_wash")),
                "line",
                radius_px=self.panel_radius,
                line_width_px=1,
                name="takeaway",
                gradient=False,
            )
            self.add_rect(
                slide,
                (left + 2, top + 2, 8, height - 4),
                str(self.panel_cfg.get("takeaway_accent", "blue")),
                name="takeaway_accent",
            )
            self.add_chevrons(slide, left + 26, top + 21, 28, "blue_mid", count=2)
            text_color = "ink"
            accent_color = self.color("red")
        elif takeaway_style == "mist":
            shape = self.add_panel(
                slide,
                (left, top, right - left, height),
                str(self.panel_cfg.get("takeaway_gradient_from", "blue_wash")),
                "line",
                radius_px=self.panel_radius,
                line_width_px=1,
                name="takeaway",
            )
            self.add_rect(
                slide,
                (left + 2, top + 2, 8, height - 4),
                str(self.panel_cfg.get("takeaway_accent", "blue")),
                fill_to=str(self.panel_cfg.get("header_gradient_to", "white")),
                gradient_direction="vertical",
                name="takeaway_accent",
            )
            self.add_chevrons(slide, left + 26, top + 21, 28, "blue_mid", count=2)
            text_color = "ink"
            accent_color = self.color("red")
        elif takeaway_style == "solid":
            shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                px(left),
                px(top),
                px(right - left),
                px(height),
            )
            shape.name = "takeaway"
            shape.adjustments[0] = min(0.5, self.panel_radius / max(1, height))
            shape.fill.solid()
            shape.fill.fore_color.rgb = self.rgb(str(self.panel_cfg.get("takeaway_fill", "blue")))
            shape.line.fill.background()
            shape.shadow.inherit = False
            self.add_rect(
                slide,
                (left + 2, top + 1, right - left - 4, 2),
                str(self.panel_cfg.get("takeaway_highlight", "blue_mid")),
                name="takeaway_highlight",
            )
            self.add_chevrons(slide, left + 26, top + 21, 28, "white", count=2)
            text_color = "white"
            accent_color = mix_color(self.color("white"), self.color("red"), 0.58)
        else:
            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(left), px(top), px(right - left), px(height))
            shape.name = "takeaway"
            shape.adjustments[0] = min(0.5, self.panel_radius / max(1, height))
            shape.fill.gradient()
            shape.fill.gradient_angle = 0
            stops = shape.fill.gradient_stops
            stops[0].color.rgb = self.rgb("blue")
            stops[0].position = 0.0
            stops[1].color.rgb = self.rgb("blue_dark")
            stops[1].position = 1.0
            shape.line.fill.background()
            shape.shadow.inherit = False
            self.add_chevrons(slide, left + 26, top + 21, 28, "white", count=2)
            text_color = "white"
            accent_color = mix_color(self.color("white"), self.color("red"), 0.58)
        self.add_fit_text(
            slide,
            (left + 88, top + 8, right - left - 112, height - 14),
            text,
            "body",
            self.type_size("takeaway", 29),
            19,
            text_color,
            max_lines=2,
            line_ratio=1.08,
            accent_color=accent_color,
            name="takeaway",
            body_floor=True,
        )

    def add_footnote(self, slide, text: str) -> None:
        text = str(text or "").strip()
        source = source_line_for_slide(self.deck, getattr(self, "current_deck_slide", {}))
        if source:
            text = f"{text} ｜ {source}" if text else source
        if not text:
            return
        left = int(self.layout["content"][0])
        right = int(self.layout["content"][2])
        top = int(self.layout["takeaway_top"]) + int(self.layout["takeaway_height"]) + 6
        size = self.type_size("footnote", 20)
        available_height = max(24, int(self.layout["footer_y"]) - top - 6)
        height = min(max(24, int(round(size * 1.18))), available_height)
        self.add_fit_text(
            slide,
            (left, top, right - left, height),
            text,
            "body",
            size,
            max(11, size - 4),
            "muted",
            max_lines=1,
            line_ratio=1.0,
            name="footnote",
            wrap_explicit=False,
        )

    def add_header(self, slide, s: dict) -> None:
        slash_left = float(self.layout["slash_x"])
        slash_top = float(self.layout["slash_y"])
        slash_width = float(self.layout["slash_width"])
        slash_height = float(self.layout["slash_height"])
        first = [
            (slash_left, slash_top + slash_height),
            (slash_left + slash_width * 0.55, slash_top),
            (slash_left + slash_width * 0.98, slash_top),
            (slash_left + slash_width * 0.43, slash_top + slash_height),
        ]
        second = [(x + slash_width + 4, y) for x, y in first]
        self.add_polygon(slide, first, "blue_dark")
        self.add_polygon(slide, second, "blue")
        eyebrow = str(s.get("eyebrow", "")).strip()
        chapter_x = float(self.layout["chapter_x"])
        title_left = self.header_title_left()
        title = str(s.get("title", "")).strip()
        title_width = float(self.theme["brand"]["header_box"][0] - title_left - 54)
        header_height = max(62, self.type_size("header_title", 48) * 1.25)
        title_size, _ = self.measure(
            title,
            "title",
            int(title_width),
            header_height,
            self.type_size("header_title", 48),
            max(34, int(self.type_size("header_title", 48) * 0.72)),
            max_lines=1,
            line_ratio=1.0,
        )
        if eyebrow:
            eyebrow_width_available = max(24.0, title_left - chapter_x - 24)
            draw = ImageDraw.Draw(Image.new("RGB", (4, 4)))
            configured_title_size = self.type_size("header_title", 48)
            configured_eyebrow_size = self.type_size("eyebrow", configured_title_size)
            if configured_eyebrow_size != configured_title_size:
                self.warn(
                    str(s.get("id", "")),
                    "Header eyebrow type size differs from header_title; set both to the same value in type_scale.",
                )
            eyebrow_size = title_size
            font = self.renderer.fonts.get(TextStyle(bold=True), eyebrow_size)
            ew = text_width(draw, eyebrow, font)
            if ew > eyebrow_width_available:
                self.warn(
                    str(s.get("id", "")),
                    "Header eyebrow exceeds its column; shorten the label instead of shrinking only the number.",
                )
            self.add_text(
                slide,
                (chapter_x, float(self.layout["title_y"]) - 4, eyebrow_width_available, max(62, eyebrow_size * 1.28)),
                eyebrow,
                eyebrow_size,
                "blue",
                bold=True,
                valign="top",
                line_ratio=1.0,
                name="chapter",
                family=self.fonts["title"],
                wrap_explicit=False,
            )
            divider_x = chapter_x + ew + HEADER_EYEBROW_GAP / 2
            if divider_x < title_left - 8:
                divider_top = float(self.layout["title_y"]) + 4
                self.add_line(slide, (divider_x, divider_top), (divider_x, divider_top + eyebrow_size * 0.92), "line", 2)
        self.set_slide_title(
            slide,
            title,
            (
                title_left,
                float(self.layout["title_y"]) - 4,
                title_width,
                max(62, int(round(float(title_size) * 1.28))),
            ),
            title_size,
            color_name="blue_dark",
            bold=True,
            family=self.fonts["title"],
            name="header_title",
            line_ratio=1.0,
            wrap_explicit=False,
        )
        divider_y = float(self.layout["divider_y"])
        self.add_rect(slide, (24, divider_y, self.width - 48, 4), "blue")
        self.add_rect(slide, (24, divider_y + 6, self.width - 48, 1), "line")
        box = tuple(self.theme["brand"]["header_box"])
        self.add_brand(slide, (box[0], box[1], box[2] - box[0], box[3] - box[1]), name="brand_lockup")
        self.add_rect(slide, (24, float(self.layout["footer_y"]), self.width - 48, 3), "blue")

    def add_sidebar(self, slide, source: dict | str | None) -> None:
        sections = self.deck.get("sections", [])
        if not sections:
            return
        if isinstance(source, dict):
            active_section = source.get("section")
            slide_id = str(source.get("id", ""))
        else:
            active_section = source
            slide_id = ""
        if self.sidebar_mode == "off":
            return
        if self.sidebar_mode == "section" and slide_id not in self.section_openers:
            return
        if self.sidebar_mode == "section" and float(self.layout["content"][0]) < 300:
            return
        left, top, right, bottom = map(float, self.layout["sidebar"])
        gap = 10
        row_height = min(96, (bottom - top - gap * (len(sections) - 1)) / len(sections))
        for index, section in enumerate(sections):
            row_top = top + index * (row_height + gap)
            active = section.get("id") == active_section
            if active:
                points = [
                    (left, row_top),
                    (right - 18, row_top),
                    (right, row_top + row_height / 2),
                    (right - 18, row_top + row_height),
                    (left, row_top + row_height),
                ]
                self.add_polygon(
                    slide,
                    points,
                    str(self.panel_cfg.get("sidebar_active_from", "blue")),
                    line_name=str(self.panel_cfg.get("sidebar_active_outline", "blue_dark")),
                    line_width_px=1,
                    name="sidebar_active",
                )
            else:
                self.add_panel(
                    slide,
                    (left, row_top, right - left - 8, row_height),
                    "white",
                    "line",
                    radius_px=6,
                    line_width_px=1.4,
                    gradient=False,
                )
                self.add_rect(slide, (left, row_top + 14, 4, row_height - 28), "blue_pale")
            self.add_fit_text(
                slide,
                (left + 12, row_top + 8, right - left - 26, row_height - 16),
                str(section.get("title", "")),
                "title",
                self.type_size("sidebar", 29),
                17,
                str(self.panel_cfg.get("sidebar_active_text", "white")) if active else "blue_dark",
                bold=True,
                align="center",
                valign="middle",
                max_lines=2,
                line_ratio=1.05,
                name=f"sidebar_{section.get('id', index)}",
            )

    def content_bounds(self, s: dict) -> tuple[int, int, int, int]:
        left, top, right, bottom = map(int, self.layout["content"])
        if s.get("takeaway"):
            bottom = int(self.layout["takeaway_top"]) - 20
        return left, top, right, bottom

    def new_slide(self, index: int) -> object:
        layout = (
            self.presentation.slide_layouts[5]
            if len(self.presentation.slide_layouts) > 5
            else self.presentation.slide_layouts[6]
        )
        slide = self.presentation.slides.add_slide(layout)
        background = self.cover_background if index == 1 and self.deck["slides"][index - 1].get("layout") == "cover" else self.background_paths[index - 1]
        picture = slide.shapes.add_picture(str(background), 0, 0, px(self.width), px(self.height))
        picture.name = "background_texture"
        mark_decorative(picture)
        self.send_to_back(picture)
        self.picture_count += 1
        return slide

    def send_to_back(self, shape) -> None:
        element = shape._element
        parent = element.getparent()
        if parent is None:
            return
        parent.remove(element)
        parent.insert(2, element)

    def set_slide_title(
        self,
        slide,
        text: str,
        box: tuple[float, float, float, float],
        size_px: float,
        color_name: str = "blue_dark",
        family: str | None = None,
        bold: bool = True,
        align: str = "left",
        line_ratio: float = 1.0,
        name: str = "slide_title",
        wrap_explicit: bool = True,
    ) -> None:
        """Write the visible header/cover title into the real title placeholder."""
        title_shape = slide.shapes.title
        if title_shape is None:
            self.add_text(
                slide,
                box,
                text,
                size_px,
                color_name,
                bold=bold,
                family=family,
                align=align,
                line_ratio=line_ratio,
                name=name,
                wrap_explicit=wrap_explicit,
            )
            return
        self.text_shape_count += 1
        left, top, width, height = box
        title_shape.left = px(left)
        title_shape.top = px(top)
        title_shape.width = px(width)
        title_shape.height = px(height)
        title_shape.name = name
        tf = title_shape.text_frame
        tf.clear()
        tf.word_wrap = bool(wrap_explicit)
        tf.auto_size = None
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.TOP
        paragraphs = str(text).split("\n") if str(text) else [""]
        for paragraph_index, paragraph_text in enumerate(paragraphs):
            paragraph = tf.paragraphs[0] if paragraph_index == 0 else tf.add_paragraph()
            paragraph.alignment = align_value(align)
            paragraph.line_spacing = 1.0 if self.profile == "defense" else line_ratio
            paragraph.space_before = Pt(0)
            paragraph.space_after = Pt(0)
            run = paragraph.add_run()
            run.text = strip_markup(paragraph_text)
            self.apply_font(
                run,
                family or self.fonts["title"],
                size_px,
                self.rgb(color_name),
                bold=bold,
            )

    def draw_cover(self, slide, s: dict) -> None:
        if str(self.theme.get("cover", {}).get("style", "centered")) != "editorial":
            self.draw_cover_centered(slide, s)
            return
        cover = self.theme["cover"]
        meta = self.deck.get("meta", {})
        accent_x = float(cover.get("accent_x", 124))
        title_width = float(cover.get("title_width", 1092))
        self.add_rect(slide, (0, 0, 12, self.height), "blue_dark", name="cover_accent")
        brand_box = tuple(cover.get("brand_box", self.theme["brand"]["cover_box"]))
        self.add_brand(slide, (brand_box[0], brand_box[1], brand_box[2] - brand_box[0], brand_box[3] - brand_box[1]), align="left", name="cover_brand")
        panel_left = float(cover.get("panel_left", self.width * 0.615))
        self.add_line(slide, (panel_left, 210), (panel_left, 930), "line", 2)
        if self.watermark_path is not None:
            wb = tuple(cover.get("watermark_box", (1288, 228, 1964, 862)))
            self.add_picture_fit(
                slide,
                self.watermark_path,
                (wb[0], wb[1], wb[2] - wb[0], wb[3] - wb[1]),
                name="cover_watermark",
                decorative=True,
            )
        kicker = str(meta.get("report_type") or meta.get("subtitle") or "学术报告").strip()
        kicker_y = float(cover.get("kicker_y", 252))
        self.add_rect(slide, (accent_x, kicker_y + 1, 7, 38), "blue", name="cover_kicker_accent")
        self.add_text(slide, (accent_x + 24, kicker_y - 2, title_width - 40, 44), kicker, self.type_size("cover_kicker", 27), "blue", bold=True, name="cover_kicker")
        title = str(s.get("title") or meta.get("title") or "").strip()
        title_lines = s.get("title_lines") or meta.get("title_lines")
        if isinstance(title_lines, list) and title_lines:
            title = "\n".join(str(line) for line in title_lines)
        title_max_size = float(cover.get("title_max_size", 66))
        title_size, _ = self.measure(
            title,
            "display",
            int(title_width),
            270,
            min(title_max_size, self.type_size("cover_title", 86)),
            44,
            max_lines=3,
            line_ratio=1.10,
        )
        title_top = float(cover.get("title_y", 304))
        self.set_slide_title(
            slide,
            title,
            (accent_x, title_top, title_width, 276),
            title_size,
            color_name="blue",
            family=self.fonts["display"],
            align="left",
            line_ratio=1.10,
            name="cover_title",
        )
        title_bottom = title_top + title_size * 1.22 * max(1, len(title.split("\n"))) + 12
        subtitle = str(s.get("subtitle") or meta.get("subtitle") or "").strip()
        subtitle_bottom = title_bottom
        if subtitle and subtitle != kicker:
            subtitle_y = title_bottom + float(cover.get("subtitle_gap", 22))
            subtitle_size, _ = self.measure(subtitle, "body", int(title_width), 96, self.type_size("cover_subtitle", 34), 22, max_lines=2, line_ratio=1.2)
            self.add_text(slide, (accent_x, subtitle_y, title_width, 96), subtitle, subtitle_size, "muted", line_ratio=1.2, name="cover_subtitle")
            subtitle_bottom = subtitle_y + subtitle_size * 1.2 * max(1, math.ceil(len(subtitle) / 32))
        rule_y = subtitle_bottom + float(cover.get("rule_gap", 34))
        self.add_rect(slide, (accent_x, rule_y, title_width, 4), "blue", name="cover_rule")
        self.add_rect(slide, (accent_x, rule_y + 10, title_width, 1), "line")
        info_y = max(float(cover.get("info_y", 708)), rule_y + float(cover.get("info_gap", 34)))
        info_items = [
            ("汇报人", str(meta.get("presenter") or "—")),
            ("指导教师", str(meta.get("advisor") or "—")),
            ("日期", str(meta.get("date") or "—")),
        ]
        value_size = int(cover.get("info_value_size", self.type_size("cover_info", 28)))
        value_font = self.renderer.fonts.get(TextStyle(), value_size)
        draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
        info_right = accent_x + title_width
        date_width = text_width(draw, info_items[2][1], value_font)
        date_x = info_right - date_width
        advisor_x = accent_x + max(280, int(title_width * 0.30))
        advisor_width = max(120, date_x - 56 - advisor_x)
        advisor_size, _ = self.measure(
            info_items[1][1],
            "body",
            int(advisor_width),
            int(value_size * 1.35),
            value_size,
            max(24, int(value_size * 0.72)),
            max_lines=1,
            line_ratio=1.0,
        )
        label_size = max(38, int(round(value_size * 0.62)))
        label_height = 46
        value_gap = 46
        department_gap = 112
        info_x = [accent_x, advisor_x, date_x]
        for index, (x, (label, value)) in enumerate(zip(info_x, info_items)):
            self.add_text(slide, (x, info_y, 310, label_height), label, label_size, "muted", line_ratio=1.0)
            size = advisor_size if index == 1 else value_size
            width = advisor_width if index == 1 else max(160, date_width + 60 if index == 2 else 330)
            self.add_text(slide, (x, info_y + value_gap, width, 54), value, size, "ink", line_ratio=1.0)
        department = str(meta.get("department") or meta.get("institution") or "").strip()
        if department:
            dept_size = max(22, int(round(value_size * 0.62)))
            self.add_text(
                slide,
                (accent_x, info_y + department_gap, title_width, 34),
                department,
                dept_size,
                "muted",
                line_ratio=1.0,
            )
        cover_line = str(meta.get("cover_line") or "").strip()
        if cover_line:
            band_y = float(cover.get("bottom_band_y", 1002))
            band_height = float(cover.get("bottom_band_height", 92))
            self.add_panel(
                slide,
                (0, band_y, self.width, band_height),
                "blue_wash",
                "line",
                radius_px=0,
                line_width_px=1,
                name="cover_bottom_band",
                gradient=False,
            )
            self.add_rect(slide, (0, band_y, 10, band_height), "blue", name="cover_band_accent")
            self.add_text(slide, (accent_x, band_y + 17, self.width - accent_x * 2, band_height - 22), cover_line, 26, "blue_dark", valign="middle", line_ratio=1.0)
        self.add_rect(slide, (24, float(self.layout["footer_y"]), self.width - 48, 3), "blue")

    def draw_cover_centered(self, slide, s: dict) -> None:
        cover = self.theme["cover"]
        meta = self.deck.get("meta", {})
        brand_box = tuple(self.theme.get("brand", {}).get("cover_box", (672, 54, 1376, 292)))
        self.add_brand(slide, (brand_box[0], brand_box[1], brand_box[2] - brand_box[0], brand_box[3] - brand_box[1]), name="cover_brand")
        title_width = 1810
        title_left = (self.width - title_width) / 2
        title = str(s.get("title") or meta.get("title") or "").strip()
        title_lines = s.get("title_lines") or meta.get("title_lines")
        if isinstance(title_lines, list) and title_lines:
            title = "\n".join(str(line) for line in title_lines)
        title_size, _ = self.measure(title, "display", title_width, 128, self.type_size("cover_title", 80), 44, max_lines=2, line_ratio=1.10)
        self.set_slide_title(
            slide,
            title,
            (title_left, float(cover.get("title_y", 352)), title_width, 136),
            title_size,
            color_name="blue",
            family=self.fonts["display"],
            align="center",
            line_ratio=1.10,
            name="cover_title",
        )
        subtitle = str(s.get("subtitle") or meta.get("subtitle") or "").strip()
        if subtitle:
            subtitle_size, _ = self.measure(subtitle, "body", title_width, 54, self.type_size("cover_subtitle", 35), 22, max_lines=1, line_ratio=1.1)
            self.add_text(slide, (title_left, float(cover.get("subtitle_y", 458)), title_width, 54), subtitle, subtitle_size, "muted", align="center", line_ratio=1.0)
        rule_y = float(cover.get("rule_y", 526))
        self.add_rect(slide, (300, rule_y, self.width - 600, 4), "blue")
        self.add_rect(slide, (300, rule_y + 10, self.width - 600, 1), "line")
        info_y = float(cover.get("info_y", 595))
        presenter = str(meta.get("presenter", ""))
        advisor = str(meta.get("advisor", ""))
        date = str(meta.get("date", ""))
        first_line = f"汇报人：{presenter}"
        second_line = f"指导教师：{advisor}"
        group = f"{first_line}      {second_line}"
        self.add_text(slide, (title_left, info_y, title_width, 38), group, self.type_size("cover_info", 31), "ink", align="center", line_ratio=1.0)
        self.add_text(slide, (title_left, info_y + 62, title_width, 38), f"汇报时间：{date}", self.type_size("cover_info", 31), "ink", align="center", line_ratio=1.0)
        self.add_rect(slide, (24, float(self.layout["footer_y"]), self.width - 48, 3), "blue")

    def draw_outline(self, slide, s: dict) -> None:
        self.add_header(slide, s)
        outline = self.theme["outline"]
        left = float(outline["x"])
        top = float(outline["y"])
        width = float(outline["width"])
        row_height = float(outline.get("row_height", 112))
        gap = float(outline.get("gap", 24))
        sections = self.deck.get("sections", [])
        spine_x = left - 34
        last_row_bottom = top + len(sections) * (row_height + gap) - gap
        index_style = str(self.layout.get("outline_style", "cards")).strip().lower() == "index"
        if index_style:
            # 与 PNG 同一套几何：单列五行铺满版心
            items = list(sections)
            rows = max(1, len(items))
            top_y = int(self.layout.get("header_height", 132)) + 48
            bottom_y = int(self.layout.get("footer_y", self.height - 96)) - 48
            row_h = (bottom_y - top_y) // rows
            self.add_text(slide, (left, top_y - 100, width, 40), "C O N T E N T S", 34, "blue_mid", align="right")
            self.add_line(slide, (left - 34, top_y + 6), (left - 34, bottom_y - 6), "blue", 3)
            for index, section in enumerate(items):
                y = top_y + index * row_h
                preferred_size = max(62, self.type_size("header_title", 62))
                title_size, _ = self.measure(
                    str(section["title"]),
                    "title",
                    width - 170,
                    max(72, row_h - 48),
                    preferred_size,
                    max(40, preferred_size - 24),
                    max_lines=1,
                    line_ratio=1.0,
                )
                desc = str(section.get("desc", "")).strip()
                text_top = y + max(18, (row_h - title_size * 1.2 - (42 if desc else 0)) / 2)
                self.add_text(
                    slide,
                    (left - 6, text_top, 142, title_size * 1.3),
                    f"{index + 1:02d}",
                    title_size,
                    "blue",
                    bold=True,
                    family=self.fonts["title"],
                    valign="top",
                    line_ratio=1.0,
                    name=f"outline_number_{index + 1}",
                    wrap_explicit=False,
                )
                self.add_text(
                    slide,
                    (left + 150, text_top, width - 170, title_size * 1.3),
                    str(section["title"]),
                    title_size,
                    "ink",
                    bold=True,
                    family=self.fonts["title"],
                    valign="top",
                    line_ratio=1.0,
                    name=f"outline_{index + 1}",
                )
                if desc:
                    self.add_text(
                        slide,
                        (left + 150, text_top + title_size * 1.2, width - 170, 44),
                        desc,
                        30,
                        "muted",
                    )
                if index < len(items) - 1:
                    self.add_line(slide, (left + 150, y + row_h - 12), (left + width, y + row_h - 12), "line", 2)
            return

        lines_style = str(self.layout.get("outline_style", "cards")).strip().lower() == "lines"
        if not lines_style:
            self.add_line(slide, (spine_x, top + 15), (spine_x, last_row_bottom - 15), "line", 2)
        for index, section in enumerate(sections):
            row_top = top + index * (row_height + gap)
            if lines_style:
                self.add_text(slide, (left, row_top + 8, 130, row_height - 16), f"{index + 1:02d}",
                              self.type_size("header_title", 46), "blue", bold=True, line_ratio=1.0,
                              name="outline_number", wrap_explicit=False)
                self.add_text(slide, (left + 140, row_top + 12, width - 210, row_height - 24), str(section["title"]),
                              self.type_size("block_title", 44), "ink", line_ratio=1.05, name="outline_title")
                self.add_line(slide, (left, row_top + row_height - 6), (left + width, row_top + row_height - 6), "line", 2)
                continue
            self.add_ellipse(slide, (spine_x - 6, row_top + row_height / 2 - 6, 12, 12), "blue_mid", None)
            self.add_panel(
                slide,
                (left, row_top, width, row_height),
                "white",
                "line",
                radius_px=9,
                line_width_px=1.6,
                shadow=True,
                gradient=False,
                name=f"outline_row_{index + 1}",
            )
            number_width = 168
            points = [
                (left + 2, row_top + 2),
                (left + number_width - 24, row_top + 2),
                (left + number_width, row_top + row_height / 2),
                (left + number_width - 24, row_top + row_height - 2),
                (left + 2, row_top + row_height - 2),
            ]
            self.add_polygon(slide, points, "blue")
            title_size, _ = self.measure(
                str(section["title"]),
                "title",
                int(width - number_width - 110),
                int(row_height - 20),
                44,
                26,
                max_lines=1,
            )
            self.add_text(
                slide,
                (left + 32, row_top + 20, 108, row_height - 40),
                f"{index + 1:02d}",
                title_size,
                "white",
                bold=True,
                valign="middle",
                line_ratio=1.0,
                name=f"outline_number_{index + 1}",
            )
            self.add_text(
                slide,
                (left + number_width + 32, row_top + 10, width - number_width - 110, row_height - 20),
                str(section["title"]),
                title_size,
                "blue_dark",
                bold=True,
                family=self.fonts["title"],
                valign="middle",
                line_ratio=1.0,
                name=f"outline_{index + 1}",
            )
            self.add_chevrons(slide, left + width - 52, row_top + row_height / 2 - 9, 18, "blue_pale", count=1)

    def draw_section(self, slide, s: dict) -> None:
        """与 PNG 预览同一套几何：页眉 + 蓝带 + 跨带渐变数字。"""
        sections = self.deck.get("sections", []) or []
        section_id = str(s.get("section", ""))
        index = next((i for i, item in enumerate(sections) if item.get("id") == section_id), 0)
        width, height = self.width, self.height

        self.add_brand(slide, (56, 30, 464, 166), align="left")
        motto = str(self.theme.get("brand", {}).get("motto") or "")
        if motto:
            self.add_text(slide, (width - 520, 76, 456, 44), motto, 34, "muted", align="right")
        rule_y = int(height * 0.19)
        self.add_rect(slide, (40, rule_y, width - 80, 3), "blue_dark")

        band_top = int(height * 0.46)
        band_bottom = int(height * 0.855)
        band_right = int(width * 0.968)
        self.add_rect(slide, (0, band_top, band_right, band_bottom - band_top), "blue")

        # 渐变数字：生成带透明度的 PNG 再贴，PPTX 里无法给文字上渐变
        number = f"{index + 1:02d}"
        num_left, num_top = int(width * 0.605), int(height * 0.222)
        num_w, num_h = int(width * 0.335), int(height * 0.44)
        number_path = self.asset_dir / f"section_number_{number}.png"
        if not number_path.exists():
            self._build_number_asset(number_path, number, num_w, num_h)
        self.add_picture_fit(slide, number_path, (num_left, num_top, num_w, num_h), name="section_number")

        title = slide_display_title(s)
        title_width = int(width * 0.42)
        title_size, _ = self.measure(
            title,
            "title",
            title_width,
            120,
            100,
            62,
            max_lines=1,
        )
        self.set_slide_title(
            slide,
            title,
            (104, band_top + 62, title_width, 120),
            title_size,
            color_name="white",
            family=self.fonts["title"],
            bold=True,
            align="left",
            line_ratio=1.0,
            name="section_title",
        )
        label = str(s.get("label_en") or "").upper()
        if label:
            self.add_text(slide, (108, band_top + 196, int(width * 0.45), 52), label, 38, "white")
        line_y = band_bottom - int((band_bottom - band_top) * 0.19)
        line_right = int(width * 0.90)
        self.add_rect(slide, (104, line_y, line_right - 104, 2), "white")
        self.add_rect(slide, (line_right - 7, line_y - 5, 14, 14), "white")

    def _build_number_asset(self, path: Path, number: str, box_w: int, box_h: int) -> None:
        """生成填满画布的渐变数字 PNG（上深下浅），供分隔页与 PPTX 共用。

        做法：先在临时画布上渲染 → 按墨迹裁切 → 等比缩放到目标框。
        直接按字号估算会把字裁掉（数字被切边就是这个原因）。
        """
        from PIL import ImageDraw as _ImageDraw
        from PIL import ImageFont as _ImageFont

        font_path = first_existing(self.theme.get("fonts", {}).get("title"))
        if font_path is None:
            font_path = first_existing(self.theme.get("fonts", {}).get("display"))
        probe = 400
        try:
            font = _ImageFont.truetype(font_path, probe)
        except Exception:
            font = _ImageFont.load_default()

        temp = Image.new("L", (probe * 3, probe * 2), 0)
        _ImageDraw.Draw(temp).text((probe // 2, probe // 4), number, font=font, fill=255)
        bbox = temp.getbbox()
        if bbox:
            glyph = temp.crop(bbox)
        else:
            glyph = temp
        scale = min(box_w / glyph.width, box_h / glyph.height)
        size = (max(1, int(glyph.width * scale)), max(1, int(glyph.height * scale)))
        mask = glyph.resize(size, Image.Resampling.LANCZOS)

        number_style = str(self.panel_cfg.get("section_number_style", "gradient")).strip().lower()
        if number_style == "flat":
            flat_rgb = self.color(str(self.panel_cfg.get("section_number_color", "blue_bright")))
            fill_layer = Image.new("RGB", size, flat_rgb)
        else:
            top_rgb = self.color("blue_dark")
            mid_rgb = self.color("blue_mid")
            bottom_rgb = tuple(int(mid_rgb[i] * 0.45 + 255 * 0.55) for i in range(3))
            fill_layer = Image.new("RGB", size)
            gdraw = _ImageDraw.Draw(fill_layer)
            for row in range(size[1]):
                p = row / max(1, size[1] - 1)
                gdraw.line((0, row, size[0], row),
                           fill=tuple(int(top_rgb[c] * (1 - p) + bottom_rgb[c] * p) for c in range(3)))
        layer = Image.new("RGBA", (box_w, box_h), (0, 0, 0, 0))
        layer.paste(fill_layer, ((box_w - size[0]) // 2, (box_h - size[1]) // 2), mask)
        layer.save(path, optimize=True)

    def draw_grid(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        blocks = s.get("blocks") or []
        if not blocks:
            self.warn(str(s.get("id", "")), "Grid layout has no blocks.")
            blocks = [{"heading": "内容待补充", "body": ""}]
        columns = max(2, min(3, int(s.get("columns", 2))))
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
        cell_height = (bottom - body_top - gap * (len(rows) - 1)) / max(1, len(rows))
        block_index = 1
        for row_index, row in enumerate(rows):
            row_top = body_top + row_index * (cell_height + gap)
            row_span = sum(span for _, span in row)
            available = right - left - gap * (len(row) - 1)
            widths = [round(available * span / row_span) for _, span in row]
            widths[-1] = available - sum(widths[:-1])
            x = left
            for (block, _), width in zip(row, widths):
                self.draw_grid_block(slide, block, (x, row_top, width, cell_height), block_index)
                x += width + gap
                block_index += 1
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_grid_block(self, slide, block: dict, box: tuple[float, float, float, float], index: int) -> None:
        left, top, width, height = box
        fill = "blue_wash" if block.get("tone") == "pale" else "white"
        self.add_panel(slide, box, fill, "line", radius_px=self.panel_radius, line_width_px=1, shadow=True)
        header_height = self.add_module_header(
            slide,
            box,
            str(block.get("heading", "")),
            index=f"{index:02d}",
            icon=str(block.get("icon", "")),
        )
        content_top = top + header_height + 12
        footer_text = str(block.get("footer") or block.get("output") or "").strip()
        footer_label = str(block.get("footer_label") or "产出").strip()
        footer_height = 52 if footer_text else 0
        footer_top = top + height - footer_height - 12 if footer_text else top + height
        content_bottom = footer_top - 8 if footer_text else top + height - 14
        image_path = resolve_path(block.get("image"), self.deck_path.parent)
        text_left = left + 16
        text_width_available = width - 32
        if image_path is not None and image_path.exists():
            if width > 520:
                image_width = width * 0.46
                image_box = (left + width - image_width - 14, content_top, image_width, content_bottom - content_top)
                self.add_picture_fit(
                    slide,
                    image_path,
                    image_box,
                    name="block_image",
                    alt_text=str(block.get("image_alt") or block.get("heading") or "内容配图"),
                )
                text_width_available = (image_box[0] - text_left - 14)
            else:
                image_height = max(50, (content_bottom - content_top) * 0.38)
                self.add_picture_fit(
                    slide,
                    image_path,
                    (left + 14, content_top, width - 28, image_height),
                    name="block_image",
                    alt_text=str(block.get("image_alt") or block.get("heading") or "内容配图"),
                )
                content_top += image_height + 10
        stat = str(block.get("stat", "")).strip()
        if stat:
            self.add_text(slide, (text_left, content_top, text_width_available, 54), stat, 46, "blue", bold=True, line_ratio=1.0)
            content_top += 58
        formula = str(block.get("formula", "")).strip()
        formula_latex = str(block.get("formula_latex", "")).strip()
        if formula or formula_latex:
            formula_box = (text_left, content_top, text_width_available, 58)
            self.add_panel(
                slide,
                formula_box,
                "blue_wash",
                "line",
                radius_px=5,
                line_width_px=1,
                gradient=False,
            )
            self.add_formula(slide, formula_latex or formula, (text_left + 8, content_top + 4, text_width_available - 16, 50), 34)
            content_top += 68
        body = str(block.get("body", "")).strip()
        if body:
            body_size, _ = self.measure(body, "body", int(text_width_available), 150, int(block.get("body_size", self.type_size("body", 22))), self.type_size("small", 15), max_lines=4, line_ratio=1.2)
            self.add_text(slide, (text_left, content_top, text_width_available, 112), body, body_size, "ink", line_ratio=1.2)
            content_top += min(112, max(30, body_size * 1.28 * max(1, math.ceil(len(strip_markup(body)) / max(8, text_width_available / max(1, body_size)))))) + 8
        diagram = block.get("diagram")
        has_block_image = image_path is not None and image_path.exists()
        if isinstance(diagram, dict) and diagram and not has_block_image and content_bottom - content_top > 96:
            diagram_height = int(
                block.get(
                    "diagram_height",
                    min(170, max(88, int((content_bottom - content_top) * 0.46))),
                )
            )
            self.draw_diagram_native(
                slide,
                (text_left, content_top, text_width_available, diagram_height),
                diagram,
                label=str(block.get("diagram_label", "")),
                frame=True,
            )
            content_top += diagram_height + 8
        visual = str(block.get("visual", "")).strip() if block.get("visual_semantic") else ""
        bullets_height = max(28, content_bottom - content_top - 10)
        if visual and bullets_height > 70:
            visual_height = max(52, min(104, bullets_height * 0.34))
            self.draw_mini_visual_native(slide, (text_left, content_bottom - visual_height, text_width_available, visual_height), visual)
            bullets_height = content_bottom - visual_height - content_top - 6
        bullets = block.get("bullets") or []
        if bullets:
            self.add_bullets(slide, bullets, (text_left, content_top, text_width_available, max(26, bullets_height)), max_size=self.type_size("bullet", 19), min_size=self.type_size("micro", 14), gap_px=6)
        if footer_text:
            self.add_output_strip(slide, (left + 14, footer_top, width - 28, height - (footer_top - top) - 12), footer_label, footer_text)

    def draw_mini_visual_native(self, slide, box: tuple[float, float, float, float], visual: str) -> None:
        left, top, width, height = box
        if visual in {"pulse", "wave"}:
            points = [(left + width * 0.05, top + height * 0.55), (left + width * 0.22, top + height * 0.3), (left + width * 0.38, top + height * 0.7), (left + width * 0.56, top + height * 0.35), (left + width * 0.78, top + height * 0.6), (left + width * 0.95, top + height * 0.42)]
            for p1, p2 in zip(points[:-1], points[1:]):
                self.add_line(slide, p1, p2, "blue_mid", 2)
            return
        if visual in {"nodes", "target", "peak", "histogram"}:
            name = "nodes" if visual == "nodes" else "target" if visual == "target" else "histogram"
            self.add_icon(slide, name, (left + width * 0.28, top, width * 0.44, height), "blue_mid", 2)
            return
        if visual == "flow":
            for index in range(3):
                self.add_panel(slide, (left + index * width * 0.32, top + height * 0.18, width * 0.2, height * 0.56), "white", "line", radius_px=4, line_width_px=1)

    def add_relation_arrow_native(
        self,
        slide,
        start: tuple[float, float],
        end: tuple[float, float],
        width_px: float,
        style: str,
        route: str,
        tone: str,
        index: int,
    ) -> None:
        start_color, end_color = relation_arrow_colors(tone, self.renderer.colors)
        style = str(style).strip().lower()
        route = str(route).strip().lower()
        line_width = max(3.0, width_px * 0.12)
        if route == "orthogonal":
            points = relation_route_points(route, start, end)
            for segment_index, (segment_start, segment_end) in enumerate(zip(points[:-1], points[1:]), start=1):
                name = f"relation_arrow_{index}" if segment_index == 1 else f"relation_arrow_{index}_seg_{segment_index}"
                self.add_line(
                    slide,
                    segment_start,
                    segment_end,
                    "blue_mid" if tone not in {"alert", "risk", "red"} else "red",
                    line_width,
                    arrow=segment_index == len(points) - 1,
                    arrow_size="med",
                    name=name,
                )
            return
        if style == "line" or math.hypot(end[0] - start[0], end[1] - start[1]) < max(24.0, width_px * 0.8):
            self.add_line(
                slide,
                start,
                end,
                "blue_mid" if tone not in {"alert", "risk", "red"} else "red",
                line_width,
                arrow=True,
                arrow_size="med",
                name=f"relation_arrow_{index}",
            )
            return
        polygon = relation_arrow_polygon(start, end, width_px, style)
        if len(polygon) < 3:
            return
        shape = self.add_polygon(
            slide,
            polygon,
            "blue",
            None,
            name=f"relation_arrow_{index}",
        )
        if shape is None:
            return
        dx, dy = end[0] - start[0], end[1] - start[1]
        angle = (-math.degrees(math.atan2(dy, dx))) % 360.0
        self.set_shape_gradient(
            shape,
            start_color,
            end_color,
            angle_degrees=angle,
        )
        shape.line.fill.background()
        shape.shadow.inherit = False

    def draw_relation_node_native(
        self,
        slide,
        box: tuple[float, float, float, float],
        node: dict,
        index: int,
    ) -> None:
        left, top, right, bottom = box
        width, height = right - left, bottom - top
        if width < 40 or height < 34:
            return
        tone = str(node.get("tone", "pale")).strip().lower()
        accent = str(node.get("accent", "")).strip().lower()
        if tone in {"solid", "primary", "strong"}:
            fill_name = "blue"
            line_name = "blue_dark"
            title_color = "white"
            body_color = "blue_pale"
        elif tone == "outline":
            fill_name = "white"
            line_name = "blue_mid"
            title_color = "blue_dark"
            body_color = "ink"
        else:
            fill_name = "blue_wash"
            line_name = "red" if accent in {"red", "alert"} else "line"
            title_color = "blue_dark"
            body_color = "ink"
        self.add_panel(
            slide,
            (left, top, width, height),
            fill_name,
            line_name,
            radius_px=max(3, min(7, int(min(width, height) * 0.035))),
            line_width_px=2 if line_name in {"blue_dark", "blue_mid", "red"} else 1,
            shadow=True,
            name=f"relation_node_{index}",
            gradient=False,
        )
        accent_width = max(4.0, min(8.0, width * 0.018))
        if tone not in {"solid", "primary", "strong"}:
            self.add_rect(
                slide,
                (left + 1, top + 7, accent_width, max(12.0, height - 14)),
                "red" if accent in {"red", "alert"} else "blue",
                name="relation_node_accent",
            )
        pad_x = max(10.0, min(24.0, width * 0.07))
        pad_y = max(8.0, min(20.0, height * 0.07))
        text_left = left + pad_x + (accent_width + 5 if tone not in {"solid", "primary", "strong"} else 0)
        text_width = max(24.0, right - pad_x - text_left)
        title_text = str(node.get("title") or node.get("label") or "").strip()
        body_text = str(node.get("body", "")).strip()
        defense = self.profile == "defense"
        title_min = 39 if defense else 22
        body_min = 39 if defense else 18
        title_max = max(title_min, min(52, int(max(42, width * 0.16))))
        body_max = max(body_min, min(40, int(max(34, width * 0.095))))
        if body_text:
            title_bottom = top + max(
                pad_y + title_min * 1.45,
                min(height * 0.46, pad_y + title_max * 1.35 * 2),
            )
            title_height = max(28.0, title_bottom - (top + pad_y))
            self.add_fit_text(
                slide,
                (text_left, top + pad_y, text_width, title_height),
                title_text,
                "title",
                title_max,
                title_min,
                title_color,
                bold=True,
                max_lines=2 if height >= 150 else 1,
                line_ratio=1.08,
                valign="middle",
                name="relation_node_title",
                type_floor_px=title_min,
            )
            self.add_fit_text(
                slide,
                (
                    text_left,
                    title_bottom + 4,
                    text_width,
                    max(24.0, bottom - pad_y - title_bottom - 4),
                ),
                body_text,
                "body",
                body_max,
                body_min,
                body_color,
                max_lines=3 if height >= 240 else 2,
                line_ratio=1.13,
                valign="top",
                name="relation_node_body",
                type_floor_px=body_min,
            )
        else:
            self.add_fit_text(
                slide,
                (text_left, top + pad_y, text_width, max(24.0, bottom - top - pad_y * 2)),
                title_text,
                "title",
                title_max,
                title_min,
                title_color,
                bold=True,
                max_lines=2,
                line_ratio=1.08,
                valign="middle",
                name="relation_node_title",
                type_floor_px=title_min,
            )

    def draw_relations_native(
        self,
        slide,
        box: tuple[float, float, float, float],
        diagram: dict,
    ) -> None:
        left, top, width, height = box
        if width < 160 or height < 100:
            return
        node_boxes = relation_node_boxes(diagram, left, top, width, height)
        if not node_boxes:
            return
        nodes_by_id = {
            str(node.get("id") or f"n{index}"): node
            for index, node in enumerate(diagram.get("nodes") or diagram.get("items") or [], start=1)
            if isinstance(node, dict)
        }
        node_ids = list(node_boxes)
        relations = [
            relation
            for relation in (diagram.get("relations") or diagram.get("edges") or [])
            if isinstance(relation, dict)
        ]
        if not relations and len(node_ids) > 1:
            relations = [
                {"from": source_id, "to": target_id}
                for source_id, target_id in zip(node_ids[:-1], node_ids[1:])
            ]
        labels: list[tuple[str, tuple[float, float], tuple[float, float], float]] = []
        for relation_index, relation in enumerate(relations, start=1):
            source_id = str(relation.get("from", "")).strip()
            target_id = str(relation.get("to", "")).strip()
            if source_id not in node_boxes or target_id not in node_boxes:
                continue
            width_value = float(relation.get("width", 0.052) or 0.052)
            arrow_width = width_value if width_value > 1.0 else width * width_value
            arrow_width = max(24.0, min(96.0, arrow_width))
            style = str(relation.get("style", "ribbon")).strip().lower()
            route = str(relation.get("route", "straight")).strip().lower()
            gap = (
                max(10.0, arrow_width * 0.25)
                if route == "orthogonal"
                else max(6.0, arrow_width * 0.10)
            )
            start, end = relation_connector_geometry(
                node_boxes[source_id],
                node_boxes[target_id],
                start_side=str(relation.get("start_side", "auto")).strip().lower(),
                end_side=str(relation.get("end_side", "auto")).strip().lower(),
                gap=gap,
            )
            self.add_relation_arrow_native(
                slide,
                start,
                end,
                arrow_width,
                style,
                route,
                str(relation.get("tone", "primary")),
                relation_index,
            )
            label = str(relation.get("label", "")).strip()
            if label:
                labels.append((label, start, end, arrow_width))

        for node_index, (node_id, node_box) in enumerate(node_boxes.items(), start=1):
            self.draw_relation_node_native(
                slide,
                node_box,
                nodes_by_id.get(node_id) or {"title": node_id},
                node_index,
            )

        for label_index, (label, start, end, arrow_width) in enumerate(labels, start=1):
            label_min = 18 if self.profile != "defense" else 39
            label_max = max(label_min, min(26 if self.profile != "defense" else 39, 39))
            size, _ = self.measure(
                label,
                "title",
                1400,
                80,
                label_max,
                label_min,
                max_lines=1,
                line_ratio=1.0,
                bold=True,
            )
            draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
            label_width = min(
                max(56.0, width - 8.0),
                max(56.0, float(text_width(draw, label, self.renderer.fonts.get(TextStyle(bold=True), size))) + 24.0),
            )
            label_height = min(
                max(30.0, height - 6.0),
                max(30.0, float(size) * 1.42),
            )
            label_box = relation_label_box(start, end, label_width, label_height, arrow_width)
            lx1, ly1, lx2, ly2 = label_box
            if lx1 < left:
                shift = left - lx1
                lx1, lx2 = lx1 + shift, lx2 + shift
            elif lx2 > left + width:
                shift = left + width - lx2
                lx1, lx2 = lx1 + shift, lx2 + shift
            if ly1 < top:
                shift = top - ly1
                ly1, ly2 = ly1 + shift, ly2 + shift
            elif ly2 > top + height:
                shift = top + height - ly2
                ly1, ly2 = ly1 + shift, ly2 + shift
            self.add_panel(
                slide,
                (lx1, ly1, lx2 - lx1, ly2 - ly1),
                "white",
                "line",
                radius_px=4,
                line_width_px=1,
                name=f"relation_label_{label_index}",
                gradient=False,
            )
            self.add_text(
                slide,
                (lx1 + 7, ly1 + 2, max(24.0, lx2 - lx1 - 14), max(18.0, ly2 - ly1 - 4)),
                label,
                size,
                "blue_dark",
                bold=True,
                family=self.fonts["title"],
                align="center",
                valign="middle",
                line_ratio=1.0,
                name=f"relation_label_text_{label_index}",
                wrap_explicit=False,
            )

    def draw_diagram_native(
        self,
        slide,
        box: tuple[float, float, float, float],
        diagram: dict,
        label: str = "",
        frame: bool = True,
    ) -> None:
        """Native PowerPoint counterpart of design.draw_diagram."""
        if not diagram:
            return
        left, top, width, height = box
        if width < 100 or height < 80:
            return
        if frame:
            self.add_panel(
                slide,
                box,
                "blue_wash",
                "line",
                radius_px=6,
                line_width_px=1,
                gradient=False,
            )
        pad = max(10, min(18, int(width * 0.018)))
        inner_left, inner_top = left + pad, top + pad
        inner_width, inner_height = width - pad * 2, height - pad * 2
        if label:
            label_size = max(22, min(32, int(height * 0.05)))
            label_height = max(30, int(label_size * 1.45))
            label_width = max(120.0, len(label) * 39.0 + 30.0)
            self.add_panel(
                slide,
                (inner_left, inner_top, label_width, label_height),
                "blue_pale",
                "line",
                radius_px=3,
                line_width_px=1,
                gradient=False,
            )
            self.add_text(
                slide,
                (inner_left + 12, inner_top + 3, label_width - 20, label_height - 6),
                label,
                label_size,
                "blue_dark",
                bold=True,
                valign="middle",
                line_ratio=1.0,
                name="diagram_label",
            )
            inner_top += label_height + 10
            inner_height -= label_height + 10
        note = str(diagram.get("note", "")).strip()
        note_size = max(16, min(26, int(height * 0.036)))
        note_height = max(34, int(note_size * 1.55)) if note else 0
        if note:
            inner_height -= note_height
        kind = str(diagram.get("kind", "sequence")).strip().lower()
        items = [item for item in (diagram.get("items") or []) if isinstance(item, dict)]
        highlight = int(diagram.get("highlight", -1))
        inner_right = inner_left + inner_width
        inner_bottom = inner_top + inner_height

        if kind == "relations":
            self.draw_relations_native(
                slide,
                (inner_left, inner_top, inner_width, inner_height),
                diagram,
            )
        elif kind == "sequence":
            items = items[:5]
            count = max(1, len(items))
            focus_style = str(self.panel_cfg.get("diagram_focus_style", "flat")).strip().lower()
            accent_focus_style = focus_style == "accent_fade"
            focus_from = str(self.panel_cfg.get("diagram_focus_from", "blue"))
            focus_to = str(self.panel_cfg.get("diagram_focus_to", "blue_dark"))
            focus_text = str(self.panel_cfg.get("diagram_focus_text", "white"))
            focus_body = str(self.panel_cfg.get("diagram_focus_body", "blue_pale"))
            gap = max(20, min(34, int(inner_width * 0.035)))
            node_width = max(64, (inner_width - gap * (count - 1)) // count)
            node_height = min(340, max(120, int(inner_height * 0.66)))
            title_size = max(24, min(44, int(node_width * 0.145)))
            body_size = max(22, min(36, int(node_width * 0.118)))
            node_top = inner_top + max(0, (inner_height - node_height) // 2)
            for index, item in enumerate(items):
                x = inner_left + index * (node_width + gap)
                is_focus = index == highlight
                accent_focus = is_focus and accent_focus_style
                self.add_panel(
                    slide,
                    (x, node_top, node_width, node_height),
                    focus_from if accent_focus else "blue_pale" if is_focus else "white",
                    "blue_dark" if accent_focus else "red" if is_focus else "line",
                    radius_px=5,
                    line_width_px=2 if is_focus else 1,
                    name=f"diagram_sequence_node_{index + 1}",
                    gradient=False,
                    fill_to=focus_to if accent_focus else None,
                    gradient_direction="horizontal",
                )
                index_box, index_size = self.diagram_index_box(x + 12, node_top + 10, f"{index + 1:02d}", 26)
                self.add_text(
                    slide,
                    index_box,
                    f"{index + 1:02d}",
                    index_size,
                    focus_text if accent_focus else "blue_dark",
                    bold=True,
                    line_ratio=1.0,
                    name="diagram_index",
                    wrap_explicit=False,
                )
                relative_title_top = max(int(round(index_size * 1.5)) + 20, int(node_height * 0.2))
                title_top = node_top + relative_title_top
                title_height = max(34, int(title_size * 1.5))
                self.add_fit_text(
                    slide,
                    (x + 12, title_top, node_width - 24, title_height),
                    str(item.get("title") or item.get("label") or ""),
                    "title",
                    title_size,
                    24,
                    focus_text if accent_focus else "blue_dark",
                    bold=True,
                    align="center",
                    max_lines=1,
                    line_ratio=1.05,
                    valign="middle",
                    name="diagram_node_title",
                )
                body = str(item.get("body", "")).strip()
                if body:
                    body_height = max(
                        20.0,
                        node_height - relative_title_top - title_height - 20,
                    )
                    self.add_fit_text(
                        slide,
                        (x + 14, title_top + title_height + 6, node_width - 28, body_height),
                        body,
                        "body",
                        body_size,
                        20,
                        focus_body if accent_focus else "ink" if is_focus else "muted",
                        align="center",
                        max_lines=3,
                        line_ratio=1.08,
                        valign="top",
                        name="diagram_node_body",
                    )
                if index < count - 1:
                    self.add_line(
                        slide,
                        (x + node_width + 4, node_top + node_height / 2),
                        (x + node_width + gap - 4, node_top + node_height / 2),
                        "blue_mid",
                        2,
                        arrow=True,
                        name="diagram_connector",
                    )
        elif kind == "branch":
            items = items[:4]
            root = diagram.get("root") if isinstance(diagram.get("root"), dict) else {}
            root_title_size = DIAGRAM_BRANCH_TEXT["root_title"]
            root_body_size = DIAGRAM_BRANCH_TEXT["root_body"]
            item_title_size = DIAGRAM_BRANCH_TEXT["item_title"]
            item_body_size = DIAGRAM_BRANCH_TEXT["item_body"]
            root_width = int(inner_width * 0.31)
            root_height = max(100, inner_height - 36)
            root_top = inner_top + (inner_height - root_height) / 2
            self.add_panel(
                slide,
                (inner_left, root_top, root_width, root_height),
                "blue_pale",
                "red",
                radius_px=5,
                line_width_px=2,
                gradient=False,
            )
            root_title_top = root_top + root_height * 0.30
            root_title_bottom = root_top + root_height * 0.54
            self.add_fit_text(
                slide,
                (
                    inner_left + 20,
                    root_title_top,
                    root_width - 40,
                    max(58, root_title_bottom - root_title_top),
                ),
                str(root.get("title") or diagram.get("center") or "共同问题"),
                "title",
                root_title_size,
                40,
                "blue_dark",
                bold=True,
                align="center",
                max_lines=2,
                valign="middle",
                name="diagram_root_title",
            )
            root_body = str(root.get("body", "")).strip()
            if root_body:
                self.add_fit_text(
                    slide,
                    (
                        inner_left + 20,
                        root_title_bottom + 10,
                        root_width - 40,
                        max(64, root_top + root_height - root_title_bottom - 38),
                    ),
                    root_body,
                    "body",
                    root_body_size,
                    DIAGRAM_BRANCH_TEXT["index"],
                    "ink",
                    align="center",
                    max_lines=3,
                    valign="middle",
                    name="diagram_root_body",
                )
            item_left = inner_left + root_width + max(34, int(inner_width * 0.08))
            item_width = inner_right - item_left
            item_gap = 12
            item_height = max(48, (inner_height - item_gap * (len(items) - 1)) / max(1, len(items)))
            for index, item in enumerate(items):
                y = inner_top + index * (item_height + item_gap)
                is_focus = index == highlight
                self.add_panel(
                    slide,
                    (item_left, y, item_width, item_height),
                    "blue_pale" if is_focus else "white",
                    "red" if is_focus else "line",
                    radius_px=5,
                    line_width_px=2 if is_focus else 1,
                    gradient=False,
                )
                index_box, index_size = self.diagram_index_box(item_left + 14, y + 14, f"{index + 1:02d}", 30)
                self.add_text(
                    slide,
                    index_box,
                    f"{index + 1:02d}",
                    index_size,
                    "blue_dark",
                    bold=True,
                    valign="middle",
                    line_ratio=1.0,
                    name="diagram_index",
                    wrap_explicit=False,
                )
                branch_text_left = index_box[0] + index_box[2] + 12
                self.add_fit_text(
                    slide,
                    (branch_text_left, y + 14, item_width - (branch_text_left - item_left) - 14, 48),
                    str(item.get("title") or item.get("label") or ""),
                    "title",
                    item_title_size,
                    30,
                    "blue_dark",
                    bold=True,
                    max_lines=1,
                    name="diagram_branch_title",
                )
                body = str(item.get("body", "")).strip()
                if body:
                    self.add_fit_text(
                        slide,
                        (branch_text_left, y + 64, item_width - (branch_text_left - item_left) - 14, max(48, item_height - 80)),
                        body,
                        "body",
                        item_body_size,
                        28,
                        "ink",
                        max_lines=2,
                        line_ratio=1.08,
                        name="diagram_branch_body",
                    )
                self.add_line(
                    slide,
                    (inner_left + root_width, root_top + root_height / 2),
                    (item_left - 6, y + item_height / 2),
                    "blue_mid",
                    2,
                    arrow=True,
                    name="diagram_branch_connector",
                )
        elif kind == "stack":
            items = items[:5]
            gap = 8
            row_height = max(38, (inner_height - gap * (len(items) - 1)) / max(1, len(items)))
            for index, item in enumerate(items):
                y = inner_top + index * (row_height + gap)
                is_focus = index == highlight
                indent = min(34, index * 10)
                self.add_panel(
                    slide,
                    (inner_left + indent, y, inner_width - indent, row_height),
                    "blue_pale" if is_focus else ("white" if index % 2 == 0 else "blue_wash"),
                    "red" if is_focus else "line",
                    radius_px=4,
                    line_width_px=2 if is_focus else 1,
                    gradient=False,
                )
                self.add_rect(slide, (inner_left + indent, y, 5, row_height), "red" if is_focus else "blue")
                index_box, index_size = self.diagram_index_box(
                    inner_left + indent + 12,
                    y + max(2, (row_height - 44) / 2),
                    f"{index + 1:02d}",
                    13,
                )
                self.add_text(
                    slide,
                    index_box,
                    f"{index + 1:02d}",
                    index_size,
                    "blue_dark",
                    bold=True,
                    line_ratio=1.0,
                    name="diagram_index",
                    wrap_explicit=False,
                )
                text_left = index_box[0] + index_box[2] + 10
                title_width = (inner_right - text_left) * 0.34
                self.add_fit_text(
                    slide,
                    (text_left, y + 4, title_width, row_height - 8),
                    str(item.get("title") or item.get("label") or ""),
                    "title",
                    17,
                    12,
                    "blue_dark",
                    bold=True,
                    max_lines=1,
                    valign="middle",
                    name="diagram_stack_title",
                )
                body = str(item.get("body", "")).strip()
                if body:
                    self.add_fit_text(
                        slide,
                        (text_left + title_width + 8, y + 4, inner_right - text_left - title_width - 18, row_height - 8),
                        body,
                        "body",
                        14,
                        10,
                        "ink" if is_focus else "muted",
                        max_lines=1,
                        valign="middle",
                        name="diagram_stack_body",
                    )
                if index < len(items) - 1:
                    self.add_line(
                        slide,
                        (inner_left + indent + (inner_width - indent) / 2, y + row_height + 1),
                        (inner_left + indent + (inner_width - indent) / 2, y + row_height + gap - 1),
                        "blue_mid",
                        2,
                        arrow=True,
                        name="diagram_stack_connector",
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
                axis_y = inner_top + inner_height * 0.55
                self.add_line(slide, (inner_left, axis_y), (inner_right, axis_y), "blue_dark", 3, arrow=True, name="diagram_ruler_axis")
                positions: dict[float, float] = {}
                for index, (item, value) in enumerate(zip(items, values)):
                    x = inner_left + (inner_width - 12) * (value - minimum) / span
                    positions[value] = x
                    is_focus = index == highlight
                    self.add_line(slide, (x, axis_y - 16), (x, axis_y + 5), "red" if is_focus else "blue_mid", 3, name="diagram_tick")
                    self.add_ellipse(slide, (x - 5, axis_y - 5, 10, 10), "red" if is_focus else "blue", "white", 2, name="diagram_tick_point")
                    self.add_text(
                        slide,
                        (x - 36, axis_y + 12, 72, 28),
                        str(item.get("label") or item.get("title") or value),
                        18,
                        "red" if is_focus else "blue_dark",
                        bold=True,
                        align="center",
                        line_ratio=1.0,
                        name="diagram_tick_label",
                    )
                    body = str(item.get("body", "")).strip()
                    if body:
                        self.add_fit_text(
                            slide,
                            (x - 48, axis_y + 42, 96, max(20, inner_bottom - axis_y - 42)),
                            body,
                            "body",
                            12,
                            9,
                            "muted",
                            align="center",
                            max_lines=2,
                            line_ratio=1.05,
                            name="diagram_tick_body",
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
                    self.add_line(slide, (x1, bracket_y), (x2, bracket_y), "blue_mid", 2, name="diagram_pair_rule")
                    self.add_line(slide, (x1, bracket_y - 5), (x1, bracket_y + 5), "blue_mid", 2)
                    self.add_line(slide, (x2, bracket_y - 5), (x2, bracket_y + 5), "blue_mid", 2)
                    label_text = str(pair.get("label", "")).strip()
                    if label_text:
                        pair_width = max(28.0, len(label_text) * 14.0 + 12.0)
                        self.add_text(
                            slide,
                            ((x1 + x2 - pair_width) / 2, bracket_y - 11, pair_width, 18),
                            label_text,
                            13,
                            "blue_mid",
                            bold=True,
                            align="center",
                            line_ratio=1.0,
                            name="diagram_pair_label",
                        )
        elif kind == "matrix":
            items = items[:4]
            gap = 8
            grid_left = inner_left + 42
            grid_top = inner_top + 6
            grid_right = inner_right
            grid_bottom = inner_bottom - 34
            cell_width = (grid_right - grid_left - gap) / 2
            cell_height = (grid_bottom - grid_top - gap) / 2
            for index, item in enumerate(items):
                row = index // 2
                column = index % 2
                x = grid_left + column * (cell_width + gap)
                y = grid_top + row * (cell_height + gap)
                is_focus = index == highlight
                self.add_panel(
                    slide,
                    (x, y, cell_width, cell_height),
                    "blue_pale" if is_focus else ("white" if row == column else "blue_wash"),
                    "red" if is_focus else "line",
                    radius_px=4,
                    line_width_px=2 if is_focus else 1,
                    gradient=False,
                )
                index_box, index_size = self.diagram_index_box(x + 9, y + 8, f"{index + 1:02d}", 13)
                self.add_text(
                    slide,
                    index_box,
                    f"{index + 1:02d}",
                    index_size,
                    "blue_dark",
                    bold=True,
                    line_ratio=1.0,
                    name="diagram_index",
                    wrap_explicit=False,
                )
                matrix_text_left = index_box[0] + index_box[2] + 8
                self.add_fit_text(
                    slide,
                    (matrix_text_left, y + 7, cell_width - (matrix_text_left - x) - 10, 32),
                    str(item.get("title") or item.get("label") or ""),
                    "title",
                    16,
                    11,
                    "blue_dark",
                    bold=True,
                    max_lines=1,
                    name="diagram_matrix_title",
                )
                body = str(item.get("body", "")).strip()
                if body:
                    self.add_fit_text(
                        slide,
                        (x + 10, y + 40, cell_width - 18, cell_height - 48),
                        body,
                        "body",
                        13,
                        9,
                        "ink",
                        max_lines=2,
                        line_ratio=1.08,
                        name="diagram_matrix_body",
                    )
            self.add_line(slide, (grid_left, grid_top), (grid_left, grid_bottom), "blue_dark", 2, name="diagram_y_axis")
            self.add_line(slide, (grid_left, grid_bottom), (grid_right, grid_bottom), "blue_dark", 2, name="diagram_x_axis")
            x_axis = diagram.get("x_axis") if isinstance(diagram.get("x_axis"), dict) else {}
            y_axis = diagram.get("y_axis") if isinstance(diagram.get("y_axis"), dict) else {}
            if x_axis.get("low"):
                self.add_fit_text(slide, (grid_left, grid_bottom + 5, cell_width, 24), str(x_axis["low"]), "body", 12, 9, "muted", align="center", max_lines=1, name="diagram_x_low")
            if x_axis.get("high"):
                self.add_fit_text(slide, (grid_right - cell_width, grid_bottom + 5, cell_width, 24), str(x_axis["high"]), "body", 12, 9, "muted", align="center", max_lines=1, name="diagram_x_high")
            if y_axis.get("high"):
                self.add_fit_text(slide, (inner_left, grid_top, grid_left - inner_left - 6, cell_height), str(y_axis["high"]), "body", 12, 9, "muted", align="center", max_lines=2, valign="middle", name="diagram_y_high")
            if y_axis.get("low"):
                self.add_fit_text(slide, (inner_left, grid_bottom - cell_height, grid_left - inner_left - 6, cell_height), str(y_axis["low"]), "body", 12, 9, "muted", align="center", max_lines=2, valign="middle", name="diagram_y_low")
        elif kind == "cycle":
            items = items[:5]
            count = max(3, len(items))
            center_x = inner_left + inner_width / 2
            center_y = inner_top + inner_height / 2
            radius_x = max(70, inner_width / 3)
            radius_y = max(52, inner_height / 3)
            points: list[tuple[float, float]] = []
            for index in range(count):
                angle = -math.pi / 2 + index * 2 * math.pi / count
                points.append((center_x + math.cos(angle) * radius_x, center_y + math.sin(angle) * radius_y))
            for index, point in enumerate(points):
                next_point = points[(index + 1) % len(points)]
                self.add_line(slide, (point[0] + 20, point[1]), (next_point[0] - 20, next_point[1]), "blue_mid", 2, arrow=True, name="diagram_cycle_connector")
            node_width = min(150, max(90, inner_width / 4))
            for index, (item, point) in enumerate(zip(items, points)):
                is_focus = index == highlight
                self.add_panel(
                    slide,
                    (point[0] - node_width / 2, point[1] - 28, node_width, 56),
                    "blue_pale" if is_focus else "white",
                    "red" if is_focus else "line",
                    radius_px=5,
                    line_width_px=2 if is_focus else 1,
                    gradient=False,
                )
                self.add_fit_text(
                    slide,
                    (point[0] - node_width / 2 + 8, point[1] - 23, node_width - 16, 46),
                    str(item.get("title") or item.get("label") or ""),
                    "title",
                    16,
                    11,
                    "blue_dark",
                    bold=True,
                    align="center",
                    max_lines=2,
                    valign="middle",
                    name="diagram_cycle_title",
                )
        else:
            self.add_fit_text(
                slide,
                (inner_left, inner_top, inner_width, inner_height),
                f"暂不支持的原生示意图：{kind}",
                "body",
                16,
                11,
                "red",
                align="center",
                valign="middle",
                name="diagram_unsupported",
            )
        if note:
            self.add_fit_text(
                slide,
                (inner_left, inner_top + inner_height + 8, inner_width, note_height - 8),
                note,
                "body",
                note_size,
                9,
                "muted",
                max_lines=1,
                line_ratio=1.0,
                name="diagram_note",
            )

    def add_formula(self, slide, latex: str, box: tuple[float, float, float, float], size_px: float = 34) -> None:
        source = (latex or "").strip()
        if not source:
            return
        if source.startswith("$") and source.endswith("$"):
            source = source[1:-1]
        try:
            omml = self.latex_to_omml(source)
        except Exception:
            self.formula_fallback += 1
            self.add_text(slide, box, source, size_px * 0.72, "blue_dark", family="Cambria Math", align="center", valign="middle", line_ratio=1.0, name="formula_fallback")
            return
        shape = slide.shapes.add_textbox(px(box[0]), px(box[1]), px(box[2]), px(box[3]))
        shape.name = "formula_native"
        self.text_shape_count += 1
        tf = shape.text_frame
        tf.clear()
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        size_hundredths = max(1200, int(round(size_px * PT_PER_PX * 100)))
        p_xml = (
            f'<a:p xmlns:a="{NS_A}" xmlns:a14="{NS_A14}" xmlns:m="{NS_M}">'
            f'<a:pPr algn="ctr"><a:defRPr sz="{size_hundredths}" b="0">'
            f'<a:latin typeface="Cambria Math"/><a:ea typeface="Cambria Math"/></a:defRPr></a:pPr>'
            f"<a14:m>{etree.tostring(omml, encoding='unicode')}</a14:m>"
            "</a:p>"
        )
        new_p = parse_xml(p_xml)
        tx_body = tf._txBody
        for paragraph in list(tx_body.p_lst):
            tx_body.remove(paragraph)
        tx_body.append(new_p)
        self.formula_native += 1

    def latex_to_omml(self, latex: str) -> etree._Element:
        if latex in self.formula_cache:
            return self.formula_cache[latex]
        from latex2mathml.converter import convert

        if not self.xsl_path.exists():
            raise FileNotFoundError(str(self.xsl_path))
        mathml = etree.fromstring(convert(latex).encode("utf-8"))
        self._stabilize_mathml_fences(mathml)
        transform = etree.XSLT(etree.parse(str(self.xsl_path)))
        result = transform(mathml)
        root = result.getroot()
        if root is None:
            raise RuntimeError("LaTeX to OMML conversion returned no root")
        self.formula_cache[latex] = root
        return root

    @staticmethod
    def _stabilize_mathml_fences(root: etree._Element) -> None:
        """Keep an operator inside a fence from becoming a delimiter separator.

        MML2OMML.XSL treats a direct ``mo`` child of a fenced ``mrow`` as the
        separator whenever every interior operator is identical.  For
        ``[1 - exp(...)]`` the minus is the only direct operator, so the XSL
        writes ``m:sepChr="−"`` and PowerPoint paints it as an arrow.  Grouping
        the interior in one extra ``mrow`` preserves the minus as ordinary math.
        """
        ns = "{http://www.w3.org/1998/Math/MathML}"
        for mrow in list(root.iter(ns + "mrow")):
            children = list(mrow)
            if len(children) < 3:
                continue
            first, last = children[0], children[-1]
            if first.tag != ns + "mo" or last.tag != ns + "mo":
                continue
            if str(first.get("fence", "")).lower() != "true":
                continue
            if str(last.get("fence", "")).lower() != "true":
                continue
            if str(first.get("form", "")) not in {"prefix", ""}:
                continue
            if str(last.get("form", "")) not in {"postfix", ""}:
                continue
            middle = children[1:-1]
            if len(middle) == 1 and middle[0].tag == ns + "mrow":
                continue
            wrapper = etree.Element(ns + "mrow")
            for child in middle:
                mrow.remove(child)
                wrapper.append(child)
            mrow.insert(1, wrapper)

    def draw_concept(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        diagram = s.get("diagram") or {}
        kind = str(diagram.get("kind", "sequence")).strip().lower()
        panels = s.get("panels") or []
        gap = 18
        if panels:
            maximum_ratio = 0.82 if kind == "relations" else 0.68
            diagram_ratio = max(0.48, min(maximum_ratio, float(s.get("diagram_width_ratio", 0.58))))
            diagram_width = (right - left) * diagram_ratio
            panel_width = right - left - diagram_width - gap
        else:
            diagram_width = right - left
            panel_width = 0
        frame = bool(s.get("diagram_frame", not (kind == "relations" and not panels)))
        self.draw_diagram_native(
            slide,
            (left, body_top, diagram_width, bottom - body_top),
            diagram,
            label=str(s.get("diagram_label", "机制示意")),
            frame=frame,
        )
        if panels:
            panel_left = left + diagram_width + gap
            panel_gap = 12
            panel_height = (bottom - body_top - panel_gap * (len(panels) - 1)) / max(1, len(panels))
            for index, panel in enumerate(panels):
                panel_top = body_top + index * (panel_height + panel_gap)
                box = (panel_left, panel_top, panel_width, panel_height)
                self.add_panel(slide, box, "white" if index % 2 == 0 else "blue_wash", "line", radius_px=self.panel_radius, line_width_px=1)
                header_height = self.add_module_header(
                    slide,
                    box,
                    str(panel.get("heading", "")),
                    index=f"{index + 1:02d}",
                    icon=str(panel.get("icon", "")),
                    height=48,
                )
                content_top = panel_top + header_height + 12
                body = str(panel.get("body", "")).strip()
                bullets = panel.get("bullets") or []
                if body:
                    body_size, body_lines = self.measure(
                        body,
                        "title",
                        int(panel_width - 28),
                        max(40, int(panel_top + panel_height - content_top - 10)),
                        self.type_size("body", 19),
                        self.type_size("small", 13),
                        max_lines=4 if not bullets else 3,
                        line_ratio=1.2,
                        bold=True,
                    )
                    body_height = max(
                        36,
                        int(panel_top + panel_height - content_top - 10),
                    )
                    self.add_text(
                        slide,
                        (panel_left + 14, content_top, panel_width - 28, body_height),
                        body,
                        body_size,
                        "ink",
                        bold=True,
                        family=self.fonts["title"],
                        line_ratio=1.2,
                        name="concept_panel_body",
                    )
                    content_top += min(
                        body_height,
                        int(round(body_size * 1.2 * max(1, len(body_lines)) + 8)),
                    )
                if bullets:
                    self.add_bullets(slide, bullets, (panel_left + 14, content_top, panel_width - 28, max(24, panel_top + panel_height - content_top - 10)), max_size=self.type_size("bullet", 16), min_size=self.type_size("micro", 11), gap_px=4)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_formula(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        gap = 18
        left_ratio = max(0.52, min(0.70, float(s.get("formula_width_ratio", 0.62))))
        left_width = (right - left) * left_ratio
        right_left = left + left_width + gap
        panel_height = int(s.get("formula_panel_height", 236))
        formula_box = (left, body_top, left_width, panel_height)
        self.add_panel(
            slide,
            formula_box,
            "blue_wash",
            "line",
            radius_px=self.panel_radius,
            line_width_px=1,
            gradient=False,
        )
        self.add_rect(slide, (left, body_top, 6, panel_height), "blue")
        label = str(s.get("formula_label", "核心关系"))
        label_size = clamp_defense_size_px(
            float(self.type_size("small", 18)),
            self.profile,
        )
        label_font = self.renderer.fonts.get(TextStyle(bold=True), int(round(label_size)))
        draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
        label_width = max(140.0, float(text_width(draw, label, label_font)) + 36.0)
        label_width = min(label_width, left_width - 36.0)
        label_height = max(44.0, round(label_size * 1.12))
        label_top = body_top + 10
        self.add_panel(
            slide,
            (left + 18, label_top, label_width, label_height),
            "blue_pale",
            "line",
            radius_px=3,
            line_width_px=1,
        )
        self.add_text(
            slide,
            (left + 30, label_top + 3, label_width - 24, label_height - 6),
            label,
            label_size,
            "blue_dark",
            bold=True,
            family=self.fonts["title"],
            valign="middle",
            line_ratio=1.0,
            name="formula_label",
            wrap_explicit=False,
        )
        formula = str(s.get("formula_latex") or s.get("formula") or "").strip()
        condition = str(s.get("formula_condition", "")).strip()
        condition_size = self.type_size("micro", 20)
        condition_height = 0
        if condition:
            condition_size, condition_lines = self.measure(
                condition,
                "body",
                int(left_width - 56),
                int(round(condition_size * 2.8)),
                condition_size,
                max(18, condition_size - 4),
                max_lines=2,
                line_ratio=1.05,
            )
            condition_height = max(
                46,
                int(round(condition_size * 1.18 * max(1, len(condition_lines)) + 12)),
            )
        math_bottom = body_top + panel_height - (condition_height + 18 if condition else 22)
        math_top = label_top + label_height + 8
        self.add_formula(
            slide,
            formula,
            (left + 28, math_top, left_width - 56, max(50, math_bottom - math_top)),
            46,
        )
        if condition:
            condition_top = body_top + panel_height - condition_height - 6
            self.add_panel(
                slide,
                (left + 18, condition_top, left_width - 36, condition_height),
                "blue_pale",
                "line",
                radius_px=3,
                line_width_px=1,
            )
            self.add_text(
                slide,
                (left + 28, condition_top + 4, left_width - 56, condition_height - 8),
                condition,
                condition_size,
                "ink",
                line_ratio=1.05,
                name="formula_condition",
            )

        variables_top = body_top + panel_height + 18
        variables = [item for item in (s.get("variables") or []) if isinstance(item, dict)]
        variable_title_size = self.type_size("small", 18)
        title_height = max(30, int(variable_title_size * 1.6))
        self.add_text(
            slide,
            (left, variables_top, left_width, title_height),
            str(s.get("variables_title", "符号与条件")),
            variable_title_size,
            "blue_dark",
            bold=True,
            line_ratio=1.0,
            name="formula_variables_title",
        )
        self.add_line(slide, (left, variables_top + title_height), (left + left_width, variables_top + title_height), "line", 2)
        if variables:
            columns = 2 if len(variables) > 6 else 1
            column_gap = 24
            column_width = (left_width - column_gap * (columns - 1)) / columns
            rows = math.ceil(len(variables) / columns)
            body_size = self.type_size("small", 16)
            symbol_font = self.renderer.fonts.get(TextStyle(bold=True), body_size)
            draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
            symbol_width = max(
                [text_width(draw, str(item.get("symbol", "")).strip(), symbol_font) for item in variables]
                or [body_size]
            )
            symbol_width = max(64, min(180, int(symbol_width) + 24))
            list_top = variables_top + title_height + 14
            available_height = max(60, bottom - list_top - 8)
            row_height = available_height / max(1, rows)
            for column in range(columns):
                chunk = variables[column * rows : (column + 1) * rows]
                if not chunk:
                    continue
                x = left + column * (column_width + column_gap)
                for row_index, item in enumerate(chunk):
                    symbol = str(item.get("symbol", "")).strip()
                    meaning = str(item.get("meaning", "")).strip()
                    y = list_top + row_index * row_height
                    symbol_top = y + max(0, (row_height - body_size * 1.3) / 2)
                    self.add_text(
                        slide,
                        (x, symbol_top, symbol_width, row_height - 4),
                        symbol,
                        body_size,
                        "blue_dark",
                        bold=True,
                        line_ratio=1.0,
                        name="formula_variable_symbol",
                    )
                    meaning_size, _ = self.measure(
                        meaning,
                        "body",
                        int(column_width - symbol_width - 10),
                        int(row_height - 4),
                        body_size,
                        max(18, body_size - 8),
                        max_lines=1,
                        line_ratio=1.0,
                    )
                    self.add_text(
                        slide,
                        (x + symbol_width, symbol_top, column_width - symbol_width - 10, row_height - 4),
                        meaning,
                        meaning_size,
                        "ink",
                        line_ratio=1.0,
                        name="formula_variable_line",
                    )

        panels = s.get("panels") or []
        if panels:
            panel_gap = 12
            panel_height_right = (bottom - body_top - panel_gap * (len(panels) - 1)) / max(1, len(panels))
            panel_width = right - right_left
            for index, panel in enumerate(panels):
                panel_top = body_top + index * (panel_height_right + panel_gap)
                box = (right_left, panel_top, panel_width, panel_height_right)
                self.add_panel(slide, box, "white" if index % 2 == 0 else "blue_wash", "line", radius_px=self.panel_radius, line_width_px=1)
                header_height = self.add_module_header(
                    slide,
                    box,
                    str(panel.get("heading", "")),
                    index=f"{index + 1:02d}",
                    icon=str(panel.get("icon", "")),
                    height=48,
                )
                content_top = panel_top + header_height + 12
                body = str(panel.get("body", "")).strip()
                bullets = panel.get("bullets") or []
                if body:
                    body_size, body_lines = self.measure(
                        body,
                        "body",
                        int(panel_width - 28),
                        max(36, int(panel_top + panel_height_right - content_top - 10)),
                        self.type_size("body", 18),
                        self.type_size("small", 12),
                        max_lines=4 if not bullets else 3,
                        line_ratio=1.18,
                    )
                    body_height = max(
                        32,
                        int(panel_top + panel_height_right - content_top - 10),
                    )
                    self.add_text(
                        slide,
                        (right_left + 14, content_top, panel_width - 28, body_height),
                        body,
                        body_size,
                        "ink",
                        line_ratio=1.18,
                    )
                    content_top += min(
                        body_height,
                        int(round(body_size * 1.18 * max(1, len(body_lines)) + 8)),
                    )
                if bullets:
                    self.add_bullets(slide, bullets, (right_left + 14, content_top, panel_width - 28, max(22, panel_top + panel_height_right - content_top - 10)), max_size=self.type_size("bullet", 16), min_size=self.type_size("micro", 11), gap_px=4)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_pipeline(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        steps = s.get("steps") or []
        if not steps:
            steps = [{"title": "步骤", "body": "内容待补充"}]
        gap = 18
        step_width = (right - left - gap * (len(steps) - 1)) / len(steps)
        line_top = body_top + 34
        self.add_line(slide, (left + 30, line_top), (right - 30, line_top), "line", 3)
        card_top = body_top + 74
        card_height = min(int(s.get("card_height", 460)), bottom - card_top - 24)
        for index, step in enumerate(steps):
            x = left + index * (step_width + gap)
            center = x + step_width / 2
            self.add_ellipse(slide, (center - 29, line_top - 29, 58, 58), "blue_pale", "blue_mid", 3)
            self.add_text(slide, (center - 24, line_top - 15, 48, 30), str(index + 1), 24, "blue_dark", bold=True, align="center", valign="middle", line_ratio=1.0)
            card = (x, card_top, step_width, card_height)
            self.add_panel(slide, card, "white", "line", radius_px=self.panel_radius, line_width_px=1, shadow=True)
            self.add_icon(slide, str(step.get("icon", "route")), (x + 18, card_top + 18, 44, 44), "blue")
            self.add_text(slide, (x + 68, card_top + 24, step_width - 80, 24), f"STEP {index + 1:02d}", 15, "muted", line_ratio=1.0)
            self.add_fit_text(slide, (x + 18, card_top + 72, step_width - 36, 58), str(step.get("title", "")), "title", 27, 18, "blue_dark", bold=True, max_lines=2, line_ratio=1.08)
            self.add_fit_text(slide, (x + 18, card_top + 138, step_width - 36, 96), str(step.get("body", "")), "body", 22, 14, "ink", max_lines=4, line_ratio=1.2)
            bullets = step.get("bullets") or []
            output = str(step.get("output", "")).strip()
            output_height = 52 if output else 0
            output_top = card_top + card_height - output_height - 14 if output else card_top + card_height - 14
            bullets_top = card_top + 246
            self.add_bullets(slide, bullets, (x + 18, bullets_top, step_width - 36, max(24, output_top - bullets_top - 12)), max_size=19, min_size=13, gap_px=6)
            if output:
                self.add_output_strip(slide, (x + 16, output_top, step_width - 32, output_height), "输出", output)
            if index < len(steps) - 1:
                self.add_chevrons(slide, x + step_width + 3, line_top - 9, 18, "blue_mid", count=1)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_compare(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        columns = s.get("columns") or []
        rows = s.get("rows") or []
        if not columns:
            self.warn(str(s.get("id", "")), "Compare table has no columns.")
            return
        table_theme = self.theme["table"]
        table_header_size = self.type_size("table_header", 25)
        table_cell_size = self.type_size("table_cell", 24)
        header_height = max(
            int(table_theme.get("header_height", 66)),
            int(table_header_size * 1.7),
        )
        row_height = min(
            max(112, int(table_cell_size * 2.1)),
            max(80, (bottom - body_top - header_height - 4) // max(1, len(rows))),
        )
        weights = s.get("column_weights") or ([1.4] + [1.0] * (len(columns) - 1))
        if len(weights) != len(columns):
            weights = [1.4] + [1.0] * (len(columns) - 1)
        total_weight = sum(weights)
        widths = [round((right - left) * weight / total_weight) for weight in weights]
        widths[-1] = right - left - sum(widths[:-1])
        table_height = header_height + row_height * len(rows)
        header_top = body_top + max(0, (bottom - body_top - table_height - 18) // 2)
        graphic = slide.shapes.add_table(len(rows) + 1, len(columns), px(left), px(header_top), px(right - left), px(table_height))
        table = graphic.table
        table.first_row = True
        table.horz_banding = False
        self.native_table_count += 1
        for index, width in enumerate(widths):
            table.columns[index].width = px(width)
        table.rows[0].height = px(header_height)
        for index in range(len(rows)):
            table.rows[index + 1].height = px(row_height)
        for column_index, column in enumerate(columns):
            cell = table.cell(0, column_index)
            cell.fill.solid()
            cell.fill.fore_color.rgb = self.rgb("blue_pale")
            cell.margin_left = px(10)
            cell.margin_right = px(10)
            cell.margin_top = px(4)
            cell.margin_bottom = px(4)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            self.add_cell_text(cell, str(column), table_header_size, "blue_dark", bold=True, align="center")
            set_cell_border(cell, self.rgb("line"), 0.7)
        highlight_rows = set(int(value) for value in s.get("highlight_rows", []))
        for row_index, row in enumerate(rows):
            fill_name = "blue_wash" if row_index % 2 == 1 or row_index in highlight_rows else "white"
            for column_index, value in enumerate(row[: len(columns)]):
                cell = table.cell(row_index + 1, column_index)
                cell.fill.solid()
                cell.fill.fore_color.rgb = self.rgb(fill_name)
                cell.margin_left = px(14)
                cell.margin_right = px(14)
                cell.margin_top = px(6)
                cell.margin_bottom = px(6)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                self.add_cell_text(
                    cell,
                    str(value),
                    table_cell_size,
                    "blue_dark" if column_index == 0 else "ink",
                    bold=column_index == 0,
                    align="left",
                )
                set_cell_border(cell, self.rgb("line"), 0.7)
            if row_index in highlight_rows:
                row_top = header_top + header_height + row_index * row_height
                self.add_rect(slide, (left, row_top, 5, row_height), "red")
        notes = s.get("notes") or []
        table_bottom = header_top + table_height
        if notes and table_bottom + 16 < int(self.layout["takeaway_top"]) - 26:
            note_top = table_bottom + 16
            note_width = (right - left - 14) / 2
            half = math.ceil(len(notes) / 2)
            note_size = self.type_size("micro", 17)
            self.add_bullets(slide, notes[:half], (left, note_top, note_width, 54), max_size=note_size, min_size=12, gap_px=3)
            self.add_bullets(slide, notes[half:], (left + note_width + 14, note_top, note_width, 54), max_size=note_size, min_size=12, gap_px=3)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def add_cell_text(self, cell, text: str, size_px: int, color_name: str, bold: bool = False, align: str = "left") -> None:
        tf = cell.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        paragraph = tf.paragraphs[0]
        paragraph.alignment = align_value(align)
        paragraph.line_spacing = 1.0 if self.profile == "defense" else 1.08
        run = paragraph.add_run()
        run.text = strip_markup(text)
        self.apply_font(run, self.fonts["title"] if bold else self.fonts["body"], size_px, self.rgb(color_name), bold=bold)

    def draw_timeline(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        phases = s.get("phases") or []
        if not phases:
            phases = [{"period": "阶段", "title": "内容待补充"}]
        axis_y = body_top + 34
        axis_left = left + 30
        axis_right = right - 30
        self.add_line(slide, (axis_left, axis_y), (axis_right, axis_y), "line", 3, arrow=True)
        period_note = str(s.get("period_note", "")).strip()
        if period_note:
            self.add_text(slide, (axis_right - 196, axis_y + 14, 196, 24), period_note, 16, "muted", align="right", line_ratio=1.0)
        card_top = body_top + 86
        timeline_note = str(s.get("timeline_note", "")).strip()
        note_height = 56 if timeline_note else 0
        max_card_bottom = bottom - (note_height + 18 if timeline_note else 0)
        card_bottom = min(max_card_bottom, card_top + int(s.get("card_height", 470)))
        gap = 18
        card_width = (right - left - gap * (len(phases) - 1)) / len(phases)
        for index, phase in enumerate(phases):
            x = left + index * (card_width + gap)
            center_x = x + card_width / 2
            self.add_line(slide, (center_x, axis_y + 14), (center_x, card_top), "line", 2)
            self.add_ellipse(slide, (center_x - 15, axis_y - 15, 30, 30), "white", "blue", 3)
            self.add_ellipse(slide, (center_x - 5, axis_y - 5, 10, 10), "blue", None)
            card = (x, card_top, card_width, card_bottom - card_top)
            self.add_panel(slide, card, "white", "line", radius_px=self.panel_radius, line_width_px=1, shadow=True)
            self.add_text(slide, (x + width_fix(card_width) - 88, card_top + 14, 76, 62), f"{index + 1:02d}", 62, "blue_pale", bold=True, align="right", line_ratio=1.0)
            period = str(phase.get("period", "")).strip()
            self.add_panel(slide, (x + 18, card_top + 18, min(card_width - 36, max(92, len(period) * 16 + 28)), 34), "blue", None, radius_px=4, line_width_px=0)
            self.add_text(slide, (x + 26, card_top + 20, min(card_width - 52, max(92, len(period) * 16 + 28)) - 16, 28), period, 18, "white", bold=True, valign="middle", line_ratio=1.0)
            title_y = card_top + 72
            self.add_fit_text(slide, (x + 18, title_y, card_width - 36, 66), str(phase.get("title", "")), "title", 27, 18, "blue_dark", bold=True, max_lines=2, line_ratio=1.05)
            body = str(phase.get("body", "")).strip()
            cursor = title_y + 66
            if body:
                body_size, _ = self.measure(body, "body", int(card_width - 36), 104, 18, 13, max_lines=3, line_ratio=1.18)
                self.add_text(slide, (x + 18, cursor, card_width - 36, 100), body, body_size, "ink", line_ratio=1.18)
                cursor += 104
            milestone = str(phase.get("milestone", "")).strip()
            milestone_height = 58 if milestone else 0
            milestone_top = card_bottom - milestone_height - 12
            bullets = phase.get("bullets") or []
            if bullets:
                self.add_bullets(slide, bullets, (x + 18, cursor, card_width - 36, max(24, milestone_top - cursor - 8)), max_size=18, min_size=12, gap_px=5)
            if milestone:
                self.add_panel(
                    slide,
                    (x + 14, milestone_top, card_width - 28, milestone_height),
                    "blue_wash",
                    "line",
                    radius_px=4,
                    line_width_px=1,
                    gradient=False,
                )
                self.add_panel(slide, (x + 14, milestone_top, 74, milestone_height), "blue", None, radius_px=4, line_width_px=0)
                self.add_text(slide, (x + 20, milestone_top + 18, 64, 24), "里程碑", 14, "white", bold=True, line_ratio=1.0)
                self.add_text(slide, (x + 96, milestone_top + 9, card_width - 112, milestone_height - 18), milestone, 16, "blue_dark", valign="middle", line_ratio=1.08)
        if timeline_note and card_bottom + 18 < bottom:
            note_top = card_bottom + 18
            self.add_panel(
                slide,
                (left + 30, note_top, right - left - 60, note_height),
                "blue_wash",
                "line",
                radius_px=self.panel_radius,
                line_width_px=1,
                gradient=False,
            )
            self.add_icon(slide, "alert", (left + 46, note_top + 9, 36, 38), "blue")
            self.add_text(slide, (left + 92, note_top + 7, right - left - 160, note_height - 14), timeline_note, 18, "blue_dark", valign="middle", line_ratio=1.05)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_hub(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        items = (s.get("items") or [])[:6]
        if not items:
            items = [{"heading": "内容待补充", "body": ""}]
        left_count = math.ceil(len(items) / 2)
        left_items = items[:left_count]
        right_items = items[left_count:]
        rows = max(len(left_items), len(right_items))
        gap = 18
        side_width = int((right - left - 380 - 80) / 2)
        side_width = max(420, min(side_width, 650))
        card_height = min(244, (bottom - body_top - gap * (rows - 1)) / max(1, rows))
        block_height = rows * card_height + gap * (rows - 1)
        start_y = body_top + max(0, (bottom - body_top - block_height) / 2)
        center_cx = (left + right) / 2
        center_cy = body_top + (bottom - body_top) / 2
        center_radius = min(162, (bottom - body_top) / 2 - 24, side_width / 2)
        side_left_x = left
        side_right_x = right - side_width
        for column_items, is_left in ((left_items, True), (right_items, False)):
            for index, item in enumerate(column_items):
                card_y = start_y + index * (card_height + gap)
                anchor_y = card_y + card_height / 2
                if is_left:
                    start = (center_cx - center_radius + 6, center_cy)
                    end = (side_left_x + side_width + 9, anchor_y)
                else:
                    start = (center_cx + center_radius - 6, center_cy)
                    end = (side_right_x - 9, anchor_y)
                self.add_line(slide, start, end, "line", 2, arrow=True)
        for column_items, x in ((left_items, side_left_x), (right_items, side_right_x)):
            for index, item in enumerate(column_items):
                card_y = start_y + index * (card_height + gap)
                self.draw_hub_card(slide, item, (x, card_y, side_width, card_height), index + (0 if x == side_left_x else left_count))
        self.add_ellipse(slide, (center_cx - center_radius, center_cy - center_radius, center_radius * 2, center_radius * 2), "blue_wash", "blue_mid", 3)
        self.add_ellipse(slide, (center_cx - center_radius + 12, center_cy - center_radius + 12, (center_radius - 12) * 2, (center_radius - 12) * 2), None, "blue_pale", 2)
        center = s.get("center") or {}
        self.add_text(slide, (center_cx - center_radius + 18, center_cy - 94, (center_radius - 18) * 2, 38), str(center.get("label", "")), 34, "blue_dark", bold=True, align="center", line_ratio=1.0)
        value = str(center.get("value", ""))
        if value:
            self.add_text(slide, (center_cx - center_radius + 24, center_cy - 40, (center_radius - 24) * 2, 92), value, 88, "blue", family=self.fonts["display"], align="center", line_ratio=1.0)
        note = str(center.get("note", ""))
        if note:
            self.add_text(slide, (center_cx - center_radius + 24, center_cy + 64, (center_radius - 24) * 2, 50), note, 18, "muted", align="center", line_ratio=1.1)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_hub_card(self, slide, item: dict, box: tuple[float, float, float, float], index: int) -> None:
        left, top, width, height = box
        fill = "blue_wash" if item.get("tone") == "pale" else "white"
        self.add_panel(slide, box, fill, "line", radius_px=self.panel_radius, line_width_px=1, shadow=True)
        self.add_rect(slide, (left, top + 15, 5, height - 30), "blue" if index % 2 else "blue_mid")
        self.add_icon(slide, str(item.get("icon", "target")), (left + 18, top + 18, 40, 40), "blue")
        self.add_fit_text(slide, (left + 70, top + 15, width - 90, 34), str(item.get("heading", "")), "title", 26, 16, "blue_dark", bold=True, max_lines=1, line_ratio=1.0)
        body = str(item.get("body", "")).strip()
        body_top = top + 56
        if body:
            body_size, _ = self.measure(body, "body", int(width - 32), 100, 19, 13, max_lines=3, line_ratio=1.15)
            self.add_text(slide, (left + 16, body_top, width - 32, 100), body, body_size, "ink", line_ratio=1.15)
            body_top += 100
        bullets = item.get("bullets") or []
        if bullets:
            self.add_bullets(slide, bullets[:2], (left + 16, body_top, width - 32, max(20, top + height - body_top - 10)), max_size=15, min_size=11, gap_px=3)

    def draw_references(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        references = s.get("references") or []
        if not references:
            references = [{"authors": "", "title": "参考文献待补充", "venue": "", "doi": ""}]
        column_gap = 44
        column_width = (right - left - column_gap) / 2
        split = math.ceil(len(references) / 2)
        columns = [(references[:split], left), (references[split:], left + column_width + column_gap)]
        divider_x = left + column_width + column_gap / 2
        self.add_line(slide, (divider_x, body_top + 6), (divider_x, bottom - 6), "line", 1)
        meta_size = self.type_size("reference_meta", 16)
        title_max_size = self.type_size("reference_title", 22)
        for column_index, (items, x) in enumerate(columns):
            if not items:
                continue
            start_number = 1 + (0 if column_index == 0 else split)
            self.add_text(slide, (x, body_top, column_width, meta_size * 1.5), f"{s.get('reference_label', '参考文献')}  {start_number:02d}–{start_number + len(items) - 1:02d}", meta_size, "muted", line_ratio=1.0)
            label_line_y = body_top + meta_size * 1.7
            self.add_line(slide, (x, label_line_y), (x + column_width, label_line_y), "line", 1)
            list_top = body_top + meta_size * 2.5
            item_gap = 16
            item_height = (bottom - list_top - item_gap * (len(items) - 1)) / len(items)
            for item_index, reference in enumerate(items):
                item_top = list_top + item_index * (item_height + item_gap)
                number = start_number + item_index
                number_size = max(28, int(title_max_size * 0.8))
                self.add_text(slide, (x, item_top + 2, number_size * 1.6, number_size * 1.15), f"{number:02d}", number_size, "blue_pale", bold=True, line_ratio=1.0)
                number_gap = max(52, int(number_size * 1.6))
                text_left = x + number_gap
                text_width_available = column_width - number_gap
                authors = str(reference.get("authors", "")).strip()
                cursor = item_top
                if authors:
                    self.add_text(slide, (text_left, cursor, text_width_available, meta_size * 1.35), authors, meta_size, "muted", line_ratio=1.0)
                    cursor += meta_size * 1.3
                title = str(reference.get("title", "")).strip()
                title_size, _ = self.measure(title, "title", int(text_width_available), int(title_max_size * 2.6), title_max_size, 16, max_lines=2, line_ratio=1.12)
                self.add_text(slide, (text_left, cursor, text_width_available, title_max_size * 2.6), title, title_size, "blue_dark", bold=False, line_ratio=1.12)
                cursor += min(title_max_size * 2.6, title_size * 1.12 * max(1, math.ceil(len(title) / max(14, text_width_available / max(1, title_size))))) + 4
                venue = str(reference.get("venue", "")).strip()
                if venue:
                    self.add_text(slide, (text_left, cursor, text_width_available, meta_size * 1.25), venue, meta_size, "ink", line_ratio=1.0)
                    cursor += meta_size * 1.2
                doi = str(reference.get("doi", "")).strip()
                if doi:
                    self.add_text(slide, (text_left, cursor, text_width_available, meta_size * 1.15), doi, max(16, int(meta_size * 0.9)), "muted", line_ratio=1.0)
                if item_index < len(items) - 1:
                    self.add_line(slide, (x, item_top + item_height - 4), (x + column_width, item_top + item_height - 4), "line", 1)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_quote(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        quote = str(s.get("quote", "")).strip() or "引文待补充"
        attribution = str(s.get("attribution", "")).strip()
        commentary = str(s.get("commentary", "")).strip()
        quote_note = str(s.get("quote_note", "")).strip()
        variant = str(s.get("variant", "centered")).strip().lower()
        if variant == "split":
            quote_box = (left, body_top, left + round((right - left) * 0.64), bottom)
            rail_box = (quote_box[2] + 20, body_top, right, bottom)
        else:
            quote_height = round((bottom - body_top) * 0.66)
            quote_box = (left, body_top, right, body_top + quote_height)
            rail_box = (left, quote_box[3] + 16, right, bottom)
        self.add_panel(slide, quote_box, "white", "line", radius_px=8, line_width_px=1.5, shadow=True)
        self.add_rect(slide, (quote_box[0], quote_box[1], 8, quote_box[3] - quote_box[1]), "blue")
        self.add_text(
            slide,
            (quote_box[0] + 24, quote_box[1], 80, 100),
            "“",
            92,
            "blue_pale",
            line_ratio=1.0,
            name="quote_mark",
        )
        text_left = quote_box[0] + 82
        text_width = quote_box[2] - text_left - 42
        text_top = quote_box[1] + 46
        text_height = quote_box[3] - text_top - 76
        self.add_fit_text(
            slide,
            (text_left, text_top, text_width, text_height),
            quote,
            "display",
            self.type_size("quote_text", 40),
            24,
            "ink",
            max_lines=7,
            line_ratio=1.35,
            family=self.fonts["display"],
            name="quote_text",
        )
        if attribution:
            self.add_fit_text(
                slide,
                (text_left, quote_box[3] - 50, text_width, 32),
                f"— {attribution}",
                "body",
                self.type_size("small", 20),
                14,
                "muted",
                align="right",
                max_lines=1,
                line_ratio=1.0,
                name="quote_attribution",
            )
        self.add_panel(slide, rail_box, "blue_wash", "line", radius_px=8, line_width_px=1.4, shadow=True)
        header_height = self.add_module_header(slide, rail_box, "解读", icon="target")
        rail_left = rail_box[0] + 18
        rail_top = rail_box[1] + header_height + 16
        rail_width = rail_box[2] - rail_left - 18
        if commentary:
            self.add_fit_text(
                slide,
                (rail_left, rail_top, rail_width, max(30, rail_box[3] - rail_top - (54 if quote_note else 18))),
                commentary,
                "body",
                self.type_size("body", 22),
                self.type_size("small", 15),
                "ink",
                max_lines=6,
                line_ratio=1.22,
                name="quote_commentary",
            )
        if quote_note:
            self.add_fit_text(
                slide,
                (rail_left, rail_box[3] - 52, rail_width, 36),
                quote_note,
                "body",
                self.type_size("micro", 14),
                11,
                "muted",
                max_lines=2,
                line_ratio=1.15,
                name="quote_note",
            )
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_closing(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        points = list(s.get("points") or [])
        has_bottom_cards = bool(
            s.get("boundary")
            or s.get("boundary_points")
            or s.get("next")
            or s.get("next_points")
        )
        if not points and not has_bottom_cards:
            points = ["结论待补充"]
        row_height = 88 if has_bottom_cards else 104
        gap = 14
        points_bottom = body_top + len(points) * (row_height + gap) - gap if points else body_top
        for index, point in enumerate(points):
            row_top = body_top + index * (row_height + gap)
            self.add_panel(slide, (left, row_top, right - left, row_height), "blue_wash" if index == 0 else "white", "line", radius_px=8, line_width_px=1.4, shadow=True)
            self.add_rect(slide, (left, row_top, 7, row_height), "red" if index == 0 else "blue")
            self.add_text(slide, (left + 24, row_top + 20, 80, 44), f"{index + 1:02d}", 44, "blue_faint", bold=True, line_ratio=1.0)
            self.add_fit_text(slide, (left + 108, row_top + 16, right - left - 136, row_height - 28), str(point), "body", self.type_size("lead", 27), 17, "ink", max_lines=2, line_ratio=1.1)
        cards = []
        if s.get("boundary") or s.get("boundary_points"):
            cards.append(("边界", "alert", str(s.get("boundary", "")), s.get("boundary_points") or [], "blue_wash"))
        if s.get("next") or s.get("next_points"):
            cards.append(("下一步", "arrow", str(s.get("next", "")), s.get("next_points") or [], "white"))
        cards_top = points_bottom + (16 if points else 0)
        # Keep the PNG and native geometry identical: cards use all space
        # between the lead and the closing line instead of a fixed height.
        footer_y = int(self.layout["footer_y"])
        cards_bottom = min(bottom - 26, footer_y - 132)
        if cards and cards_bottom - cards_top > 80:
            gap_cards = 16
            card_width = (right - left - gap_cards * (len(cards) - 1)) / len(cards)
            for index, (title, icon, body, points, fill) in enumerate(cards):
                x = left + index * (card_width + gap_cards)
                self.add_panel(
                    slide,
                    (x, cards_top, card_width, cards_bottom - cards_top),
                    fill,
                    "line",
                    radius_px=8,
                    line_width_px=1.2,
                    shadow=True,
                    name=f"closing_card_{index + 1}",
                )
                self.add_icon(
                    slide,
                    icon,
                    (x + card_width - 50, cards_top + 18, 32, 32),
                    "red" if index == 0 else "blue",
                )
                title_size = max(52, int(self.type_size("block_title", 48) * 1.08))
                self.add_text(
                    slide,
                    (x + 18, cards_top + 18, card_width - 74, title_size * 1.35),
                    title,
                    title_size,
                    "blue_dark",
                    bold=True,
                    line_ratio=1.0,
                    name=f"closing_card_title_{index + 1}",
                )
                cursor = cards_top + max(58, int(title_size * 1.55))
                if body:
                    body_size = self.type_size("small", 16)
                    measured_size, body_lines = self.measure(
                        body,
                        "body",
                        max(20, int(card_width - 36)),
                        max(34, int(cards_bottom - cursor - 16)),
                        body_size,
                        max(18, int(body_size * 0.8)),
                        max_lines=2,
                        line_ratio=1.22,
                    )
                    body_height = max(34, int(measured_size * 1.22 * max(1, len(body_lines))))
                    self.add_text(
                        slide,
                        (x + 18, cursor, card_width - 36, body_height),
                        body,
                        measured_size,
                        "ink",
                        bold=True,
                        line_ratio=1.22,
                        name=f"closing_card_body_{index + 1}",
                    )
                    cursor += body_height + 4
                if points:
                    self.add_bullets(
                        slide,
                        points,
                        (x + 18, cursor, card_width - 36, max(24, cards_bottom - cursor - 12)),
                        max_size=max(32, int(self.type_size("small", 16) * 0.95)),
                        min_size=26,
                        gap_px=4,
                    )
        closing = str(s.get("closing", "谢谢各位老师")).strip()
        closing_y = min(
            cards_bottom + 30 if cards and cards_bottom - cards_top > 80 else bottom - 62,
            footer_y - 96,
        )
        self.add_line(slide, (left, closing_y - 24), (right, closing_y - 24), "line", 1)
        self.add_fit_text(slide, (left + 40, closing_y, right - left - 80, 58), closing, "title", 48, 28, "blue", bold=True, align="center", max_lines=1, line_ratio=1.0)

    def draw_bullets(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        columns = s.get("columns") or [{"heading": "要点", "bullets": s.get("bullets", [])}]
        gap = 18
        weights = s.get("column_weights") or [1.0] * len(columns)
        if len(weights) != len(columns):
            weights = [1.0] * len(columns)
        available = right - left - gap * (len(columns) - 1)
        widths = [round(available * weight / sum(weights)) for weight in weights]
        widths[-1] = available - sum(widths[:-1])
        x = left
        for index, (column, width) in enumerate(zip(columns, widths)):
            box = (x, body_top, width, bottom - body_top)
            self.add_panel(slide, box, "blue_wash" if index == 0 else "white", "line", radius_px=8, line_width_px=1.2, shadow=True)
            header_height = self.add_module_header(slide, box, str(column.get("heading", "")), index=f"{index + 1:02d}", icon=str(column.get("icon", "")))
            content_top = body_top + header_height + 12
            body = str(column.get("body", "")).strip()
            if body:
                body_size, body_lines = self.measure(
                    body,
                    "body",
                    max(20, width - 36),
                    180,
                    self.type_size("body", 22),
                    self.type_size("small", 16),
                    max_lines=3,
                    line_ratio=1.12,
                    bold=True,
                )
                body_height = max(body_size * 1.12, body_size * 1.12 * max(1, len(body_lines)))
                self.add_fit_text(slide, (x + 18, content_top, width - 36, body_height), body, "body", self.type_size("body", 22), self.type_size("small", 16), max_lines=3, line_ratio=1.12)
                content_top += body_height + 12
            self.add_bullets(slide, column.get("bullets") or [], (x + 18, content_top, width - 36, bottom - content_top - 18), max_size=self.type_size("bullet", 21), min_size=self.type_size("micro", 14))
            x += width + gap
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_figure(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        image_path = resolve_path(s.get("image"), self.deck_path.parent)
        side = str(s.get("image_side", "left"))
        variant = str(s.get("variant", "")).strip().lower()
        if variant == "image_right":
            side = "right"
        elif variant == "image_left":
            side = "left"
        gap = 20
        total_width = right - left
        image_width = round(total_width * figure_width_ratio(s))
        text_width_available = total_width - image_width - gap
        if side == "right":
            image_box = (right - image_width, body_top, image_width, bottom - body_top)
            panel_box = (left, body_top, text_width_available, bottom - body_top)
        else:
            image_box = (left, body_top, image_width, bottom - body_top)
            panel_box = (left + image_width + gap, body_top, text_width_available, bottom - body_top)
        chart_data = s.get("chart")
        if chart_data:
            self.add_chart_card(
                slide,
                image_box,
                chart_data,
                str(s.get("image_label", "图版")),
                str(s.get("image_caption", "")),
                str(s.get("image_badge", "")),
                name="hero_chart",
            )
        else:
            self.add_panel(
                slide,
                image_box,
                "white",
                "line",
                radius_px=8,
                line_width_px=1.2,
                shadow=True,
                name="figure_card",
                gradient=False,
            )
            label_text = str(s.get("image_label", "")).strip()
            caption_text = str(s.get("image_caption", "")).strip()
            badge_text = str(s.get("image_badge", "")).strip()
            chrome_size = int(round(clamp_defense_size_px(20, self.profile)))
            inside_header = self.profile == "defense"
            header_bottom = body_top + 14
            measure = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
            if label_text:
                label_width = min(
                    image_width - 36,
                    max(200, int(text_width(measure, label_text, self.renderer.fonts.get(TextStyle(bold=True), chrome_size))) + 40),
                )
                label_height = max(34, int(round(chrome_size * 1.2)))
                label_top = body_top + 10 if inside_header else body_top - 15
                self.add_panel(
                    slide,
                    (image_box[0] + 18, label_top, label_width, label_height),
                    "blue_pale",
                    "line",
                    radius_px=5,
                    line_width_px=1,
                    shadow=True,
                    gradient=False,
                    name="figure_label_fill",
                )
                self.add_text(
                    slide,
                    (image_box[0] + 30, label_top + 3, label_width - 24, label_height - 6),
                    label_text,
                    chrome_size,
                    "blue_dark",
                    bold=True,
                    family=self.fonts["title"],
                    valign="middle",
                    line_ratio=1.0,
                    wrap_explicit=False,
                    name="figure_label",
                )
                header_bottom = max(header_bottom, label_top + label_height + 10)
            if badge_text:
                badge_size = max(16, int(round(chrome_size * 0.85)))
                badge_width = max(70, int(text_width(measure, badge_text, self.renderer.fonts.get(TextStyle(bold=True), badge_size))) + 26)
                badge_height = max(28, int(round(badge_size * 1.45)))
                badge_top = body_top + 10 if inside_header else body_top - 15
                self.add_panel(
                    slide,
                    (image_box[0] + image_width - 18 - badge_width, badge_top, badge_width, badge_height),
                    "blue_pale",
                    None,
                    radius_px=4,
                    line_width_px=0,
                    gradient=False,
                )
                self.add_text(
                    slide,
                    (image_box[0] + image_width - 18 - badge_width, badge_top + 2, badge_width, badge_height - 4),
                    badge_text,
                    badge_size,
                    "blue_dark",
                    bold=True,
                    align="center",
                    valign="middle",
                    line_ratio=1.0,
                    wrap_explicit=False,
                    name="figure_badge",
                )
                header_bottom = max(header_bottom, badge_top + badge_height + 10)
            caption_height = 0
            if caption_text:
                caption_size = int(round(self.type_size("micro", 16)))
                caption_size, caption_lines = self.measure(
                    caption_text,
                    "body",
                    max(40, image_width - 32),
                    200,
                    caption_size,
                    12,
                    max_lines=2,
                    line_ratio=1.0,
                )
                caption_height = max(36, int(round(caption_size * 1.05 * max(1, len(caption_lines)) + 14)))
                self.add_panel(
                    slide,
                    (image_box[0], bottom - caption_height, image_width, caption_height),
                    "blue_wash",
                    None,
                    radius_px=0,
                    line_width_px=0,
                    gradient=False,
                )
                self.add_text(
                    slide,
                    (image_box[0] + 16, bottom - caption_height + 6, image_width - 32, caption_height - 10),
                    "\n".join(caption_lines) if caption_lines else caption_text,
                    caption_size,
                    "muted",
                    line_ratio=1.0,
                    wrap_explicit=False,
                    name="figure_caption",
                )
            if image_path and image_path.exists():
                self.add_picture_fit(
                    slide,
                    image_path,
                    (
                        image_box[0] + 14,
                        header_bottom,
                        image_width - 28,
                        max(40, bottom - header_bottom - caption_height - 10),
                    ),
                    name="hero_figure",
                    alt_text=self.image_alt(s),
                )
            else:
                self.add_text(
                    slide,
                    (image_box[0] + 36, body_top + 36, image_width - 72, 92),
                    "此页需要论文原图或真实图表。请在 deck.json 的 image 或 chart 字段补充素材。",
                    24,
                    "muted",
                    line_ratio=1.2,
                    name="missing_figure",
                )
        panels = s.get("panels") or [{"heading": "图示", "body": s.get("image_role", "请结合原图讲解。")}]
        panel_left, panel_top, panel_width, panel_height_total = panel_box
        gap = 10
        panel_height = (panel_height_total - gap * (len(panels) - 1)) / len(panels)
        for index, panel in enumerate(panels):
            current_top = panel_top + index * (panel_height + gap)
            box = (panel_left, current_top, panel_width, panel_height)
            self.add_panel(slide, box, "white", "line", radius_px=8, line_width_px=1.2, shadow=True)
            header_height = self.add_module_header(slide, box, str(panel.get("heading", "")), index=f"{index + 1:02d}", icon=str(panel.get("icon", "")))
            content_top = current_top + header_height + 12
            self.add_fit_text(
                slide,
                (panel_left + 16, content_top, panel_width - 32, panel_height - header_height - 24),
                str(panel.get("body", "")),
                "body",
                18,
                12,
                max_lines=3 if panel.get("bullets") else 4,
                line_ratio=1.1,
            )
            bullets = panel.get("bullets") or []
            if bullets:
                self.add_bullets(
                    slide,
                    bullets,
                    (panel_left + 16, current_top + panel_height - 88, panel_width - 32, 76),
                    max_size=16,
                    min_size=11,
                    gap_px=3,
                )
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_figure_flow(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        flow = s.get("flow") or []
        if not flow:
            self.warn(str(s.get("id", "")), "Figure-flow layout has no flow steps.")
            flow = [{"title": "步骤", "body": ""}]
        flow_gap = 18
        flow_width = int((right - left - flow_gap * (len(flow) - 1)) / len(flow))
        flow_height = 112
        for index, step in enumerate(flow):
            x = left + index * (flow_width + flow_gap)
            box = (x, body_top, flow_width, flow_height)
            self.add_panel(slide, box, "blue_wash" if index % 2 == 0 else "white", "line", radius_px=8, line_width_px=1.2, shadow=True)
            center_x = x + 38
            center_y = body_top + flow_height / 2
            badge_size = 52
            self.add_ellipse(
                slide,
                (center_x - badge_size / 2, center_y - badge_size / 2, badge_size, badge_size),
                "blue_wash",
                "line",
                2,
            )
            icon_size = badge_size * 0.58
            self.add_icon(
                slide,
                str(step.get("icon", "signal")),
                (center_x - icon_size / 2, center_y - icon_size / 2, icon_size, icon_size),
                "blue",
                2,
            )
            self.add_fit_text(
                slide,
                (x + 74, body_top + 8, flow_width - 86, 44),
                str(step.get("title", "")),
                "title",
                21,
                14,
                "blue_dark",
                bold=True,
                max_lines=1,
                line_ratio=1.0,
                name=f"flow_title_{index + 1}",
            )
            self.add_fit_text(
                slide,
                (x + 74, body_top + 58, flow_width - 86, 50),
                str(step.get("body", "")),
                "body",
                15,
                10,
                "muted",
                max_lines=2,
                line_ratio=1.2,
                name=f"flow_body_{index + 1}",
            )
            if index < len(flow) - 1:
                self.add_chevrons(slide, x + flow_width + 1, body_top + 46, 22, "blue_mid", count=1)

        lower_top = body_top + flow_height + 16
        image_width = round((right - left) * figure_width_ratio(s, 0.60))
        figure_box = (left, lower_top, image_width, bottom - lower_top)
        chart_data = s.get("chart")
        image_path = resolve_path(s.get("image"), self.deck_path.parent)
        if chart_data:
            self.add_chart_card(
                slide,
                figure_box,
                chart_data,
                str(s.get("image_label", "图版")),
                str(s.get("image_caption", "")),
                str(s.get("image_badge", "")),
                name="flow_chart",
            )
        else:
            self.add_panel(
                slide,
                figure_box,
                "white",
                "line",
                radius_px=8,
                line_width_px=1.2,
                shadow=True,
                gradient=False,
            )
            if image_path and image_path.exists():
                self.add_picture_fit(
                    slide,
                    image_path,
                    (left + 14, lower_top + 14, image_width - 28, bottom - lower_top - 28),
                    name="flow_figure",
                    alt_text=self.image_alt(s),
                )
            else:
                self.add_text(
                    slide,
                    (left + 36, lower_top + 36, image_width - 72, 92),
                    "此页需要论文原图、真实图表或明确的机制示意图。请在 deck.json 的 image 或 chart 字段补充素材。",
                    24,
                    "muted",
                    line_ratio=1.2,
                    name="missing_figure",
                )

        panels = s.get("panels") or [{"heading": "图示", "body": s.get("image_role", "请结合原图讲解。")}]
        panel_left = left + image_width + 20
        panel_width = right - panel_left
        gap = 10
        panel_height = (bottom - lower_top - gap * (len(panels) - 1)) / len(panels)
        for index, panel in enumerate(panels):
            current_top = lower_top + index * (panel_height + gap)
            box = (panel_left, current_top, panel_width, panel_height)
            self.add_panel(slide, box, "white", "line", radius_px=8, line_width_px=1.2, shadow=True)
            header_height = self.add_module_header(slide, box, str(panel.get("heading", "")), index=f"{index + 1:02d}", icon=str(panel.get("icon", "")))
            content_top = current_top + header_height + 12
            body = str(panel.get("body", ""))
            bullets = panel.get("bullets") or []
            self.add_fit_text(
                slide,
                (panel_left + 16, content_top, panel_width - 32, 60),
                body,
                "body",
                18,
                12,
                max_lines=2 if bullets else 3,
                line_ratio=1.1,
            )
            if bullets:
                self.add_bullets(
                    slide,
                    bullets,
                    (panel_left + 16, content_top + 66, panel_width - 32, max(24, panel_height - header_height - 90)),
                    max_size=16,
                    min_size=11,
                    gap_px=3,
                )
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_metrics(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        metrics = s.get("metrics") or []
        if not metrics:
            self.warn(str(s.get("id", "")), "Metrics slide has no metrics.")
            metrics = [{"value": "—", "label": "指标", "note": ""}]
        gap = 16
        card_width = (right - left - gap * (len(metrics) - 1)) / len(metrics)
        value_size = min(64, self.type_size("metric_value", 64))
        label_size = max(24, self.type_size("metric_label", 24))
        note_size = max(16, self.type_size("micro", 14))
        card_height = max(
            196,
            int(value_size * 1.05 + label_size * 1.3 + note_size * 1.35 + 42),
        )
        for index, metric in enumerate(metrics):
            x = left + index * (card_width + gap)
            box = (x, body_top, card_width, card_height)
            self.add_panel(
                slide,
                box,
                "blue_wash" if index == 0 else "white",
                "line",
                radius_px=10,
                line_width_px=1.4,
                shadow=True,
                gradient=False,
                name=f"metric_card_{index + 1}",
            )
            self.add_rect(slide, (x + 2, body_top + 2, card_width - 4, 5), "blue" if index == 0 else "blue_mid")
            self.add_fit_text(
                slide,
                (x + 20, body_top + 18, card_width - 98, 80),
                str(metric.get("value", "")),
                "title",
                value_size,
                28,
                "blue",
                bold=True,
                max_lines=1,
                line_ratio=1.05,
                name=f"metric_value_{index + 1}",
            )
            label_y = body_top + 18 + value_size * 1.05 + 4
            self.add_fit_text(
                slide,
                (x + 20, label_y, card_width - 96, 58),
                str(metric.get("label", "")),
                "title",
                label_size,
                max(18, int(label_size * 0.8)),
                "ink",
                bold=True,
                max_lines=1,
                line_ratio=1.15,
                name=f"metric_label_{index + 1}",
            )
            self.add_fit_text(
                slide,
                (x + 20, body_top + card_height - note_size * 1.45 - 8, card_width - 36, note_size * 1.45),
                str(metric.get("note", "")),
                "body",
                note_size,
                max(16, int(note_size * 0.8)),
                "muted",
                max_lines=1,
                line_ratio=1.15,
                name=f"metric_note_{index + 1}",
            )

        lower_top = body_top + card_height + 18
        lower_bottom = bottom
        image_path = resolve_path(s.get("image"), self.deck_path.parent)
        image_width = round((right - left) * 0.58)
        evidence_box = (left, lower_top, image_width, lower_bottom - lower_top)
        chart_data = s.get("chart")
        if chart_data:
            self.add_chart_card(
                slide,
                evidence_box,
                chart_data,
                str(s.get("image_label", "测量结果")),
                str(s.get("image_caption", "")),
                str(s.get("image_badge", "")),
                name="metrics_chart",
            )
        elif image_path and image_path.exists():
            self.add_panel(
                slide,
                evidence_box,
                "white",
                "line",
                radius_px=8,
                line_width_px=1.2,
                shadow=True,
                gradient=False,
            )
            label = str(s.get("image_label", "测量结果"))
            caption = str(s.get("image_caption", ""))
            draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
            label_size = int(round(clamp_defense_size_px(20, self.profile)))
            label_font = self.renderer.fonts.get(TextStyle(bold=True), label_size)
            label_width = min(evidence_box[2] - 36, max(260, int(text_width(draw, label, label_font)) + 40))
            label_height = max(38, int(round(label_size * 1.2)))
            label_top = lower_top + 8
            self.add_panel(
                slide,
                (left + 16, label_top, label_width, label_height),
                "blue_pale",
                "line",
                radius_px=5,
                line_width_px=1,
                gradient=False,
            )
            self.add_text(
                slide,
                (left + 28, label_top + 4, label_width - 24, label_height - 8),
                label,
                label_size,
                color_name="blue_dark",
                bold=True,
                family=self.fonts["title"],
                valign="middle",
                line_ratio=1.0,
                name="metrics_image_label",
                wrap_explicit=False,
            )
            badge_text = str(s.get("image_badge", "")).strip()
            if badge_text:
                badge_size = max(16, int(round(label_size * 0.85)))
                badge_width = max(
                    70,
                    int(text_width(draw, badge_text, self.renderer.fonts.get(TextStyle(bold=True), badge_size))) + 26,
                )
                badge_height = max(28, int(round(badge_size * 1.45)))
                self.add_panel(
                    slide,
                    (left + image_width - 34 - badge_width, label_top, badge_width, badge_height),
                    "blue_pale",
                    None,
                    radius_px=4,
                    line_width_px=0,
                )
                self.add_text(
                    slide,
                    (left + image_width - 34 - badge_width, label_top + 2, badge_width, badge_height - 4),
                    badge_text,
                    badge_size,
                    "blue_dark",
                    bold=True,
                    align="center",
                    valign="middle",
                    line_ratio=1.0,
                    wrap_explicit=False,
                    name="metrics_image_badge",
                )
            caption_size = int(round(self.type_size("micro", 16)))
            caption_lines: list[str] = []
            caption_height = 0
            if caption:
                caption_size, caption_lines = self.measure(
                    caption,
                    "body",
                    max(40, int(image_width - 36)),
                    220,
                    caption_size,
                    12,
                    max_lines=2,
                    line_ratio=1.0,
                )
                caption_height = max(
                    42,
                    int(round(caption_size * 1.05 * max(1, len(caption_lines)) + 14)),
                )
            image_top = label_top + label_height + 10
            image_box = (
                left + 18,
                image_top,
                image_width - 36,
                max(40, lower_bottom - image_top - caption_height - 10),
            )
            self.add_picture_fit(
                slide,
                image_path,
                image_box,
                name="metrics_image",
                alt_text=self.image_alt(s),
            )
            if caption:
                self.add_panel(
                    slide,
                    (left, lower_bottom - caption_height, image_width, caption_height),
                    "blue_wash",
                    None,
                    radius_px=0,
                    line_width_px=0,
                    gradient=False,
                )
                self.add_text(
                    slide,
                    (left + 18, lower_bottom - caption_height + 6, image_width - 36, caption_height - 10),
                    "\n".join(caption_lines) if caption_lines else caption,
                    caption_size,
                    "muted",
                    line_ratio=1.0,
                    name="metrics_image_caption",
                    wrap_explicit=False,
                )

        side_left = left + image_width + 18
        side_box = (side_left, lower_top, right - side_left, lower_bottom - lower_top)
        self.add_panel(slide, side_box, "white", "line", radius_px=9, line_width_px=1.4, shadow=True)
        header_height = self.add_module_header(
            slide,
            side_box,
            str(s.get("conditions_title", "实验条件")),
            icon="alert",
        )
        content_top = lower_top + header_height + 12
        conditions = s.get("conditions") or s.get("bullets") or []
        note = str(s.get("conditions_note", "")).strip()
        note_size = int(round(self.type_size("micro", 16)))
        note_lines: list[str] = []
        note_height = 70 if note else 0
        if note:
            note_size, note_lines = self.measure(
                note,
                "body",
                max(40, int(side_box[2] - 54)),
                180,
                note_size,
                12,
                max_lines=2,
                line_ratio=1.0,
            )
            note_height = max(
                70,
                int(round(note_size * 1.0 * max(1, len(note_lines)) + 24)),
            )
        conditions_bottom = lower_bottom - 16 - note_height
        if conditions:
            if isinstance(conditions[0], dict):
                columns = 2
                rows = max(1, math.ceil(len(conditions) / columns))
                cell_gap = 12
                cell_width = (side_box[2] - 32 - cell_gap) / columns
                defense_grid = self.profile == "defense"
                if defense_grid:
                    row_height = min(108, max(88, (conditions_bottom - content_top) // rows))
                else:
                    row_height = min(96, max(58, (conditions_bottom - content_top) // rows))
                for index, condition in enumerate(conditions):
                    row = index // columns
                    column = index % columns
                    cell_x = side_left + 16 + column * (cell_width + cell_gap)
                    cell_y = content_top + row * row_height
                    label = str(condition.get("label", ""))
                    value = str(condition.get("value", ""))
                    if label:
                        label_height = 38 if defense_grid else 40
                        label_size = 32 if defense_grid else self.type_size("micro", 15)
                        self.add_text(
                            slide,
                            (cell_x, cell_y + 2, cell_width, label_height),
                            label,
                            label_size,
                            "muted",
                            family=self.fonts["title"],
                            line_ratio=1.0,
                            name=f"condition_label_{index + 1}",
                            wrap_explicit=not defense_grid,
                        )
                    value_top = cell_y + (42 if defense_grid else 42)
                    value_height = max(40 if defense_grid else 26, row_height - (46 if defense_grid else 42))
                    self.add_fit_text(
                        slide,
                        (
                            cell_x,
                            value_top,
                            cell_width,
                            value_height,
                        ),
                        value,
                        "title",
                        38 if defense_grid else 22,
                        38 if defense_grid else 14,
                        "blue_dark",
                        max_lines=1 if defense_grid else 2,
                        line_ratio=1.0 if defense_grid else 1.15,
                        name=f"condition_value_{index + 1}",
                    )
                    if not defense_grid:
                        self.add_line(
                            slide,
                            (cell_x, cell_y + row_height - 5),
                            (cell_x + cell_width, cell_y + row_height - 5),
                            "line",
                            1,
                        )
            else:
                self.add_bullets(
                    slide,
                    conditions,
                    (side_left + 16, content_top, side_box[2] - 32, max(26, conditions_bottom - content_top)),
                    max_size=20,
                    min_size=14,
                    gap_px=6,
                )
        if note:
            note_box = (
                side_left + 14,
                lower_bottom - note_height,
                side_box[2] - 30,
                note_height - 12,
            )
            self.add_panel(
                slide,
                note_box,
                "blue_wash",
                "line",
                radius_px=6,
                line_width_px=1,
                gradient=False,
            )
            self.add_rect(slide, (note_box[0], note_box[1], 5, note_box[3]), "red")
            self.add_text(
                slide,
                (note_box[0] + 16, note_box[1] + 8, note_box[2] - 24, note_box[3] - 12),
                "\n".join(note_lines) if note_lines else note,
                note_size,
                "ink",
                line_ratio=1.0,
                name="conditions_note",
                wrap_explicit=False,
            )
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_balance(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s)
        gap = 24
        rail_width = int((right - left) * 0.255)
        center_left = left + rail_width + gap
        center_right = right - rail_width - gap
        self.draw_balance_rail(slide, s.get("left") or {}, (left, body_top, rail_width, bottom - body_top), 0)
        self.draw_balance_rail(slide, s.get("right") or {}, (right - rail_width, body_top, rail_width, bottom - body_top), 1)
        center_box = (center_left, body_top, center_right - center_left, bottom - body_top)
        chart_data = s.get("chart")
        if chart_data:
            self.add_chart_card(
                slide,
                center_box,
                chart_data,
                str(s.get("image_label", "图版")),
                str(s.get("image_caption", "")),
                str(s.get("image_badge", "")),
                name="balance_chart",
            )
        else:
            self.add_figure_card(slide, center_box, s, name="balance_figure_card")
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def add_figure_card(self, slide, box: tuple[float, float, float, float], s: dict, name: str = "figure_card") -> None:
        left, top, width, height = box
        bottom = top + height
        self.add_panel(
            slide,
            box,
            "white",
            "line",
            radius_px=8,
            line_width_px=1.2,
            shadow=True,
            name=name,
            gradient=False,
        )
        label_text = str(s.get("image_label", "")).strip()
        caption_text = str(s.get("image_caption", "")).strip()
        badge_text = str(s.get("image_badge", "")).strip()
        chrome_size = int(round(clamp_defense_size_px(20, self.profile)))
        inside_header = self.profile == "defense"
        header_bottom = top + 14
        measure = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
        if label_text:
            label_width = min(
                width - 36,
                max(200, int(text_width(measure, label_text, self.renderer.fonts.get(TextStyle(bold=True), chrome_size))) + 40),
            )
            label_height = max(34, int(round(chrome_size * 1.2)))
            label_top = top + 10 if inside_header else top - 15
            self.add_panel(
                slide,
                (left + 18, label_top, label_width, label_height),
                str(self.panel_cfg.get("label_fill", "blue_pale")),
                "line",
                radius_px=5,
                line_width_px=1,
                shadow=True,
                gradient=False,
                name="figure_label_fill",
            )
            self.add_text(
                slide,
                (left + 30, label_top + 3, label_width - 24, label_height - 6),
                label_text,
                chrome_size,
                str(self.panel_cfg.get("label_text", "blue_dark")),
                bold=True,
                family=self.fonts["title"],
                valign="middle",
                line_ratio=1.0,
                wrap_explicit=False,
                name="figure_label",
            )
            header_bottom = max(header_bottom, label_top + label_height + 10)
        if badge_text:
            badge_size = max(16, int(round(chrome_size * 0.85)))
            badge_width = max(
                70,
                int(text_width(measure, badge_text, self.renderer.fonts.get(TextStyle(bold=True), badge_size))) + 26,
            )
            badge_height = max(28, int(round(badge_size * 1.45)))
            badge_top = top + 10 if inside_header else top - 15
            self.add_panel(
                slide,
                (left + width - 18 - badge_width, badge_top, badge_width, badge_height),
                "blue_pale",
                None,
                radius_px=4,
                line_width_px=0,
            )
            self.add_text(
                slide,
                (left + width - 18 - badge_width, badge_top + 2, badge_width, badge_height - 4),
                badge_text,
                badge_size,
                "blue_dark",
                bold=True,
                align="center",
                valign="middle",
                line_ratio=1.0,
                wrap_explicit=False,
                name="figure_badge",
            )
            header_bottom = max(header_bottom, badge_top + badge_height + 10)
        caption_height = 0
        if caption_text:
            caption_size = int(round(self.type_size("micro", 16)))
            caption_size, caption_lines = self.measure(
                caption_text,
                "body",
                max(40, width - 32),
                200,
                caption_size,
                12,
                max_lines=2,
                line_ratio=1.0,
            )
            caption_height = max(36, int(round(caption_size * 1.05 * max(1, len(caption_lines)) + 14)))
            self.add_panel(
                slide,
                (left, bottom - caption_height, width, caption_height),
                "blue_wash",
                None,
                radius_px=0,
                line_width_px=0,
                gradient=False,
            )
            self.add_text(
                slide,
                (left + 16, bottom - caption_height + 6, width - 32, caption_height - 10),
                "\n".join(caption_lines) if caption_lines else caption_text,
                caption_size,
                "muted",
                line_ratio=1.0,
                wrap_explicit=False,
                name="figure_caption",
            )
        image_path = resolve_path(s.get("image"), self.deck_path.parent)
        if image_path and image_path.exists():
            self.add_picture_fit(
                slide,
                image_path,
                (left + 14, header_bottom, width - 28, max(40, bottom - header_bottom - caption_height - 10)),
                name="balance_figure",
                alt_text=self.image_alt(s),
            )
        else:
            self.add_text(
                slide,
                (left + 18, header_bottom, width - 36, max(40, bottom - header_bottom - caption_height - 10)),
                "图片缺失",
                chrome_size,
                "muted",
                align="center",
                valign="middle",
                name="balance_figure_missing",
            )

    def draw_balance_rail(self, slide, rail: dict, box: tuple[float, float, float, float], index: int) -> None:
        left, top, width, height = box
        bottom = top + height
        self.add_panel(slide, box, "blue_wash" if index == 0 else "white", "line", radius_px=8, line_width_px=1.4, shadow=True)
        header_height = self.add_module_header(
            slide,
            box,
            str(rail.get("heading", "")),
            icon=str(rail.get("icon", "check" if index == 0 else "alert")),
        )
        items = rail.get("items") or []
        if not items:
            self.warn(str(self.current_deck_slide.get("id", "")), "Balance rail has no items.")
            return
        header_bottom = top + header_height
        item_height = (bottom - header_bottom - 8) / len(items)
        for item_index, item in enumerate(items):
            item_top = header_bottom + item_index * item_height
            if item_index:
                self.add_line(
                    slide,
                    (left + 18, item_top),
                    (left + width - 18, item_top),
                    "line",
                    1,
                )
            badge_size = 66 if item_height >= 180 else 54
            badge_left = left + 24
            badge_top = item_top + 24
            self.add_ellipse(
                slide,
                (badge_left, badge_top, badge_size, badge_size),
                "blue_wash",
                "line",
                2,
            )
            inner = badge_size * 0.58
            self._add_icon(
                slide,
                str(item.get("icon", "check" if index == 0 else "alert")),
                (
                    badge_left + (badge_size - inner) / 2,
                    badge_top + (badge_size - inner) / 2,
                    inner,
                    inner,
                ),
                "blue",
                2,
            )
            text_left = badge_left + badge_size + 20
            text_width_available = left + width - text_left - 20
            heading = str(item.get("heading", ""))
            heading_size, _ = self.measure(
                heading,
                "title",
                max(40, int(text_width_available)),
                60,
                36,
                24,
                max_lines=1,
                line_ratio=1.0,
                bold=True,
            )
            heading_height = max(46, int(round(heading_size * 1.35)))
            self.add_text(
                slide,
                (text_left, item_top + 22, text_width_available, heading_height),
                heading,
                heading_size,
                "blue_dark",
                bold=True,
                family=self.fonts["title"],
                line_ratio=1.0,
                wrap_explicit=False,
                name="balance_item_heading",
            )
            body_top = item_top + 22 + heading_height + 2
            body_height = max(28, item_top + item_height - body_top - 12)
            body = str(item.get("body", "")).strip()
            if body:
                self.add_fit_text(
                    slide,
                    (text_left, body_top, text_width_available, body_height),
                    body,
                    "body",
                    24,
                    18,
                    "ink",
                    max_lines=3,
                    line_ratio=1.2,
                    name="balance_item_body",
                )
            elif item.get("bullets"):
                self.add_bullets(
                    slide,
                    item.get("bullets") or [],
                    (text_left, body_top, text_width_available, body_height),
                    max_size=24,
                    min_size=18,
                    gap_px=4,
                )

    def export(self, pptx_name: str = "defense_deck.pptx") -> dict:
        self.presentation = Presentation()
        self.presentation.slide_width = px(self.width)
        self.presentation.slide_height = px(self.height)
        slides = self.selected_slides()
        for index, s in enumerate(slides, start=1):
            slide = self.new_slide(index)
            layout = s.get("layout")
            self.current_deck_slide = s
            if layout not in {"cover", "outline", "section"}:
                self.add_header(slide, s)
            if layout == "cover":
                self.draw_cover(slide, s)
            elif layout == "outline":
                self.draw_outline(slide, s)
            elif layout == "section":
                self.draw_section(slide, s)
            elif layout == "grid":
                self.draw_grid(slide, s)
            elif layout == "pipeline":
                self.draw_pipeline(slide, s)
            elif layout == "compare":
                self.draw_compare(slide, s)
            elif layout == "timeline":
                self.draw_timeline(slide, s)
            elif layout == "hub":
                self.draw_hub(slide, s)
            elif layout == "references":
                self.draw_references(slide, s)
            elif layout == "quote":
                self.draw_quote(slide, s)
            elif layout == "closing":
                self.draw_closing(slide, s)
            elif layout == "bullets":
                self.draw_bullets(slide, s)
            elif layout == "figure":
                self.draw_figure(slide, s)
            elif layout == "figure_flow":
                self.draw_figure_flow(slide, s)
            elif layout == "concept":
                self.draw_concept(slide, s)
            elif layout == "formula":
                self.draw_formula(slide, s)
            elif layout == "metrics":
                self.draw_metrics(slide, s)
            elif layout == "balance":
                self.draw_balance(slide, s)
            else:
                self.warn(str(s.get("id", index)), f"Unsupported native layout: {layout}")
                self.add_text(slide, (200, 200, self.width - 400, 100), f"暂不支持的可编辑版式：{layout}", 32, "red")
            speech = str(s.get("speech", "")).strip()
            if speech:
                slide.notes_slide.notes_text_frame.text = speech
            finalize_slide_accessibility(slide)
            chart_count = sum(1 for shape in slide.shapes if getattr(shape, "has_chart", False))
            self.native_chart_count += chart_count
            self.slide_stats.append(
                {
                    "page": index,
                    "layout": layout,
                    "shape_count": len(slide.shapes),
                    "text_shape_count": sum(1 for shape in slide.shapes if shape.has_text_frame and shape.text.strip()),
                    "picture_count": sum(1 for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE),
                    "table_count": sum(1 for shape in slide.shapes if shape.has_table),
                    "chart_count": chart_count,
                }
            )
        meta = self.deck.get("meta", {})
        core = self.presentation.core_properties
        core.title = str(meta.get("title", ""))
        core.author = str(meta.get("presenter") or "GYSoaring")
        core.subject = str(meta.get("report_type") or "Academic presentation")
        core.keywords = ", ".join(str(section.get("title", "")) for section in self.deck.get("sections", []))
        core.language = str(meta.get("language") or "zh-CN")
        core.comments = "Generated by academic-ppt-skill"
        pptx_path = self.output_dir / pptx_name
        self.presentation.save(str(pptx_path))
        report = {
            "deck": str(self.deck_path),
            "theme": self.theme["name"],
            "pptx": str(pptx_path),
            "slide_count": len(slides),
            "text_shape_count": self.text_shape_count,
            "picture_count": self.picture_count,
            "native_table_count": self.native_table_count,
            "native_chart_count": self.native_chart_count,
            "formula_native_count": self.formula_native,
            "formula_fallback_count": self.formula_fallback,
            "fonts": self.fonts,
            "slides": self.slide_stats,
            "issues": self.issues,
        }
        write_json(self.output_dir / "pptx_report.json", report)
        return report


def width_fix(value: float) -> float:
    return max(1.0, float(value))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export an editable native PowerPoint deck.")
    parser.add_argument("--deck", required=True)
    parser.add_argument("--theme", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--pptx-name", default="defense_deck.pptx")
    parser.add_argument("--profile", choices=["document", "projection", "lecture", "defense"], default=None)
    parser.add_argument("--sample", type=int, default=None, help="Export only the first N slides as a visual sample.")
    parser.add_argument("--pages", default=None, help="Comma-separated 1-based slide numbers, e.g. 1,4,7.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    pages = None
    if args.pages:
        pages = [int(value) for value in str(args.pages).replace("，", ",").split(",") if value.strip()]
    exporter = EditableDeckExporter(
        Path(args.deck),
        Path(args.theme),
        Path(args.out),
        profile=args.profile,
        pages=pages,
        sample=args.sample,
    )
    report = exporter.export(args.pptx_name)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
