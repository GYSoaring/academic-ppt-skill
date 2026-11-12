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
    fit_text_block,
    load_json,
    load_theme,
    mix_color,
    resolve_path,
    text_width,
    write_json,
)
from design import Run, TextStyle, parse_markup
from render_deck_next import DeckRenderer


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
    def __init__(self, deck_path: Path, theme_path: Path, output_dir: Path) -> None:
        self.deck_path = deck_path.resolve()
        self.theme_path = theme_path.resolve()
        self.output_dir = output_dir.resolve()
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.asset_dir = self.output_dir / "assets"
        self.asset_dir.mkdir(parents=True, exist_ok=True)
        self.deck = load_json(self.deck_path)
        self.theme = load_theme(self.theme_path)
        self.renderer = DeckRenderer(self.deck_path, self.theme_path, self.output_dir)
        self.width = int(self.theme["canvas"]["width"])
        self.height = int(self.theme["canvas"]["height"])
        self.colors = self.theme["_colors"]
        self.layout = self.theme["layout"]
        self.type_scale = self.theme.get("type_scale", {})
        self.panel_radius = int(self.layout.get("panel_radius", 8))
        self.panel_shadow_alpha = int(self.layout.get("panel_shadow_alpha", 20))
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

    def warn(self, page: str, message: str) -> None:
        self.issues.append({"page": page, "message": message})

    def color(self, name: str) -> tuple[int, int, int]:
        return self.colors[name]

    def rgb(self, name: str) -> RGBColor:
        return RGBColor(*self.color(name))

    def type_size(self, name: str, fallback: int) -> int:
        return int(self.type_scale.get(name, fallback))

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
            "body": str(configured.get("body") or family_from_path(body_path, "Microsoft YaHei")),
            "display": str(configured.get("display") or family_from_path(display_path, "STZhongsong")),
            "formula": str(configured.get("formula") or "Cambria Math"),
            "latin": str(configured.get("latin") or family_from_path(latin_path, "Arial")),
        }

    def _prepare_backgrounds(self) -> None:
        for index, slide in enumerate(self.deck.get("slides", []), start=1):
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
        run.font.name = family
        run.font.size = pt(size_px)
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
    ):
        left, top, width, height = box
        shape = slide.shapes.add_textbox(px(left), px(top), px(width), px(height))
        if name:
            shape.name = name
        self.text_shape_count += 1
        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = px(margin_px)
        tf.margin_right = px(margin_px)
        tf.margin_top = 0
        tf.margin_bottom = 0
        tf.vertical_anchor = anchor_value(valign)
        tf.auto_size = None
        paragraphs = str(text).split("\n") if str(text) else [""]
        font_family = family or self.fonts["body"]
        for paragraph_index, paragraph_text in enumerate(paragraphs):
            paragraph = tf.paragraphs[0] if paragraph_index == 0 else tf.add_paragraph()
            paragraph.alignment = align_value(align)
            paragraph.line_spacing = line_ratio
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
    ) -> tuple[int, list[str]]:
        draw = ImageDraw.Draw(Image.new("RGB", (4, 4), (255, 255, 255)))
        fonts = self.renderer.theme["fonts"].get(kind) or self.renderer.theme["fonts"]["body"]
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
    ):
        left, top, width, height = box
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(left), px(top), px(width), px(height))
        if name:
            shape.name = name
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
            tail = parse_xml(f'<a:tailEnd xmlns:a="{NS_A}" type="triangle" w="sm" len="sm"/>')
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
    ):
        left, top, width, height = box
        shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, px(left), px(top), px(width), px(height))
        if name:
            shape.name = name
        if fill_name:
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
    ):
        if len(points) < 3:
            return None
        x0, y0 = points[0]
        builder = slide.shapes.build_freeform(x0, y0, scale=EMU_PER_PX)
        builder.add_line_segments(points[1:], close=True)
        shape = builder.convert_to_shape()
        if name:
            shape.name = name
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
        self.picture_count += 1
        return shape

    def add_corner_marks(self, slide, box: tuple[float, float, float, float], color_name: str = "line") -> None:
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

        chart = graphic.chart
        chart.has_legend = False
        chart.has_title = False
        chart.font.name = self.fonts["body"]
        chart.font.size = pt(15)
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
        self.add_panel(slide, box, "white", "line", radius_px=12, line_width_px=1.2, shadow=True)
        if label:
            label_width = max(120.0, len(label) * 20.0 + 34.0)
            label_box = (left + 18, top - 15, min(width - 36, label_width), 34)
            self.add_panel(slide, label_box, "blue", None, radius_px=5, line_width_px=0, shadow=True)
            self.add_text(
                slide,
                (label_box[0] + 12, label_box[1] + 4, label_box[2] - 20, label_box[3] - 7),
                label,
                20,
                "white",
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
            self.add_picture_fit(slide, lockup, box, align=align, name=name)
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
            name=name,
        )

    def add_icon(
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

    def add_module_header(self, slide, box: tuple[float, float, float, float], heading: str, index: str = "", icon: str = "", height: int = 56) -> None:
        left, top, width, _ = box
        header = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(left + 2), px(top + 2), px(width - 4), px(height))
        header.adjustments[0] = min(0.5, self.panel_radius / max(1, height))
        header.fill.solid()
        header.fill.fore_color.rgb = self.rgb("blue")
        header.line.fill.background()
        header.shadow.inherit = False
        self.add_rect(slide, (left + 3, top + 3, width - 6, 3), "blue_mid")
        text_left = left + 16
        if index:
            self.add_text(slide, (left + 18, top + 12, 36, 32), index, 20, "blue_pale", bold=True)
            text_left = left + 56
        if heading:
            available = width - (text_left - left) - (56 if icon else 18)
            self.add_fit_text(
                slide,
                (text_left, top + 9, available, height - 12),
                heading,
                "title",
                self.type_size("block_title", 26),
                17,
                "white",
                bold=True,
                max_lines=1,
                line_ratio=1.0,
            )
        if icon:
            self.add_icon(slide, icon, (left + width - 48, top + 12, 32, 32), "white")

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
            paragraph.line_spacing = 1.18
            marker = paragraph.add_run()
            marker.text = "◆  "
            self.apply_font(marker, self.fonts["body"], chosen, self.rgb("blue_mid"), bold=True)
            run = paragraph.add_run()
            run.text = strip_markup(str(bullet))
            self.apply_font(run, self.fonts["body"], chosen, self.rgb("ink"))

    def add_output_strip(self, slide, box: tuple[float, float, float, float], label: str, text: str) -> None:
        left, top, width, height = box
        if not text or width < 80 or height < 28:
            return
        self.add_panel(slide, box, "blue_wash", "line", radius_px=4, line_width_px=1)
        label_width = min(max(58, len(strip_markup(label)) * 16 + 24), max(58, width // 3))
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
        left, _, right, _ = map(int, self.layout["content"])
        height = int(self.layout.get("lead_height", 88))
        panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(left), px(top), px(right - left), px(height))
        panel.adjustments[0] = min(0.5, self.panel_radius / max(1, height))
        panel.fill.gradient()
        panel.fill.gradient_angle = 0
        stops = panel.fill.gradient_stops
        stops[0].color.rgb = RGBColor(*mix_color(self.color("blue_wash"), self.color("white"), 0.12))
        stops[0].position = 0.0
        stops[1].color.rgb = self.rgb("white")
        stops[1].position = 1.0
        panel.line.color.rgb = self.rgb("line")
        panel.line.width = pt(1)
        panel.shadow.inherit = False
        label = str(self.layout.get("lead_label", "")).strip()
        if label:
            label_width = min(int(self.layout.get("lead_label_width", 168)), max(120, (right - left) // 3))
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
            self.add_text(slide, (left + 8, top + 8, label_width - 30, height - 16), label, 22, "white", bold=True, align="center", valign="middle", name="lead_label")
            text_left = left + label_width + 24
            text_width_available = right - text_left - 24
        else:
            self.add_rect(slide, (left, top + 1, 9, height - 2), "blue")
            self.add_chevrons(slide, left + 28, top + 26, 34, "blue_mid", count=2)
            text_left = left + 96
            text_width_available = right - left - 118
        self.add_fit_text(
            slide,
            (text_left, top + 13, text_width_available, height - 22),
            text,
            "body",
            self.type_size("lead", 31),
            22,
            "ink",
            max_lines=2,
            line_ratio=1.12,
            name="lead_claim",
        )
        return top + height + 20

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
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(left), px(top), px(right - left), px(height))
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
        self.add_fit_text(
            slide,
            (left + 88, top + 8, right - left - 112, height - 14),
            text,
            "body",
            self.type_size("takeaway", 29),
            19,
            "white",
            max_lines=2,
            line_ratio=1.08,
            accent_color=mix_color(self.color("white"), self.color("red"), 0.58),
            name="takeaway",
        )

    def add_footnote(self, slide, text: str) -> None:
        if not text:
            return
        left = int(self.layout["content"][0])
        right = int(self.layout["content"][2])
        top = int(self.layout["takeaway_top"]) + int(self.layout["takeaway_height"]) + 8
        self.add_fit_text(slide, (left, top, right - left, 22), text, "body", 15, 11, "muted", max_lines=1, line_ratio=1.0, name="footnote")

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
        self.add_polygon(slide, first, "blue")
        self.add_polygon(slide, second, "blue_mid")
        eyebrow = str(s.get("eyebrow", "")).strip()
        chapter_x = float(self.layout["chapter_x"])
        title_left = float(self.layout["title_x"])
        title = str(s.get("title", "")).strip()
        title_width = float(self.theme["brand"]["header_box"][0] - title_left - 54)
        title_size, _ = self.measure(
            title,
            "title",
            int(title_width),
            62,
            self.type_size("header_title", 48),
            34,
            max_lines=1,
            line_ratio=1.0,
        )
        if eyebrow:
            eyebrow_width_available = max(24.0, title_left - chapter_x - 18)
            draw = ImageDraw.Draw(Image.new("RGB", (4, 4)))
            font = self.renderer.fonts.get(TextStyle(bold=True), title_size)
            ew = text_width(draw, eyebrow, font)
            eyebrow_size = title_size
            if ew > eyebrow_width_available:
                self.warn(str(s.get("id", "")), f"Header eyebrow is too long at title size: {eyebrow}")
                eyebrow_size = max(18, int(title_size * eyebrow_width_available / max(1, ew)))
                font = self.renderer.fonts.get(TextStyle(bold=True), eyebrow_size)
                ew = text_width(draw, eyebrow, font)
            self.add_text(
                slide,
                (chapter_x, float(self.layout["title_y"]) - 4, eyebrow_width_available, 62),
                eyebrow,
                eyebrow_size,
                "blue",
                bold=True,
                valign="top",
                line_ratio=1.0,
                name="chapter",
            )
            divider_x = chapter_x + ew + 18
            if divider_x < title_left - 20:
                divider_top = float(self.layout["title_y"]) + 4
                self.add_line(slide, (divider_x, divider_top), (divider_x, divider_top + eyebrow_size * 0.92), "line", 2)
        self.add_text(
            slide,
            (title_left, float(self.layout["title_y"]) - 4, title_width, 62),
            title,
            title_size,
            "blue_dark",
            bold=True,
            family=self.fonts["title"],
            line_ratio=1.0,
            name="header_title",
        )
        divider_y = float(self.layout["divider_y"])
        self.add_rect(slide, (24, divider_y, self.width - 48, 4), "blue")
        self.add_rect(slide, (24, divider_y + 6, self.width - 48, 1), "line")
        box = tuple(self.theme["brand"]["header_box"])
        self.add_brand(slide, (box[0], box[1], box[2] - box[0], box[3] - box[1]), name="brand_lockup")
        self.add_rect(slide, (24, float(self.layout["footer_y"]), self.width - 48, 3), "blue")

    def add_sidebar(self, slide, active_section: str | None) -> None:
        sections = self.deck.get("sections", [])
        if not sections:
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
                self.add_polygon(slide, points, "blue", name="sidebar_active")
            else:
                self.add_panel(slide, (left, row_top, right - left - 8, row_height), "white", "line", radius_px=6, line_width_px=1.4)
                self.add_rect(slide, (left, row_top + 14, 4, row_height - 28), "blue_pale")
            self.add_fit_text(
                slide,
                (left + 12, row_top + 8, right - left - 26, row_height - 16),
                str(section.get("title", "")),
                "title",
                self.type_size("sidebar", 29),
                17,
                "white" if active else "blue_dark",
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
        slide = self.presentation.slides.add_slide(self.presentation.slide_layouts[6])
        background = self.cover_background if index == 1 and self.deck["slides"][index - 1].get("layout") == "cover" else self.background_paths[index - 1]
        picture = slide.shapes.add_picture(str(background), 0, 0, px(self.width), px(self.height))
        picture.name = "background_texture"
        self.picture_count += 1
        return slide

    def draw_cover(self, slide, s: dict) -> None:
        if str(self.theme.get("cover", {}).get("style", "centered")) != "editorial":
            self.draw_cover_centered(slide, s)
            return
        cover = self.theme["cover"]
        meta = self.deck.get("meta", {})
        accent_x = float(cover.get("accent_x", 124))
        title_width = float(cover.get("title_width", 1092))
        self.add_rect(slide, (0, 0, 12, self.height), "blue")
        brand_box = tuple(cover.get("brand_box", self.theme["brand"]["cover_box"]))
        self.add_brand(slide, (brand_box[0], brand_box[1], brand_box[2] - brand_box[0], brand_box[3] - brand_box[1]), align="left", name="cover_brand")
        panel_left = float(cover.get("panel_left", self.width * 0.615))
        self.add_line(slide, (panel_left, 210), (panel_left, 930), "line", 2)
        if self.watermark_path is not None:
            wb = tuple(cover.get("watermark_box", (1288, 228, 1964, 862)))
            self.add_picture_fit(slide, self.watermark_path, (wb[0], wb[1], wb[2] - wb[0], wb[3] - wb[1]), name="cover_watermark")
        kicker = str(meta.get("report_type") or meta.get("subtitle") or "学术报告").strip()
        kicker_y = float(cover.get("kicker_y", 252))
        self.add_rect(slide, (accent_x, kicker_y + 1, 7, 38), "blue")
        self.add_text(slide, (accent_x + 24, kicker_y - 2, title_width - 40, 44), kicker, self.type_size("cover_kicker", 27), "blue", bold=True, name="cover_kicker")
        title = str(s.get("title") or meta.get("title") or "").strip()
        title_lines = s.get("title_lines") or meta.get("title_lines")
        if isinstance(title_lines, list) and title_lines:
            title = "\n".join(str(line) for line in title_lines)
        title_size, _ = self.measure(title, "display", int(title_width), 270, min(66, self.type_size("cover_title", 86)), 44, max_lines=3, line_ratio=1.10)
        title_top = float(cover.get("title_y", 304))
        title_shape = self.add_text(slide, (accent_x, title_top, title_width, 276), title, title_size, "blue", family=self.fonts["display"], line_ratio=1.10, name="cover_title")
        title_bottom = title_top + title_size * 1.22 * max(1, len(title.split("\n"))) + 12
        subtitle = str(s.get("subtitle") or meta.get("subtitle") or "").strip()
        subtitle_bottom = title_bottom
        if subtitle and subtitle != kicker:
            subtitle_y = title_bottom + float(cover.get("subtitle_gap", 22))
            subtitle_size, _ = self.measure(subtitle, "body", int(title_width), 96, self.type_size("cover_subtitle", 34), 22, max_lines=2, line_ratio=1.2)
            self.add_text(slide, (accent_x, subtitle_y, title_width, 96), subtitle, subtitle_size, "muted", line_ratio=1.2, name="cover_subtitle")
            subtitle_bottom = subtitle_y + subtitle_size * 1.2 * max(1, math.ceil(len(subtitle) / 32))
        rule_y = subtitle_bottom + float(cover.get("rule_gap", 34))
        self.add_rect(slide, (accent_x, rule_y, title_width, 4), "blue")
        self.add_rect(slide, (accent_x, rule_y + 10, title_width, 1), "line")
        info_y = max(float(cover.get("info_y", 708)), rule_y + float(cover.get("info_gap", 34)))
        info_items = [
            ("汇报人", str(meta.get("presenter") or "—")),
            ("指导教师", str(meta.get("advisor") or "—")),
            ("日期", str(meta.get("date") or "—")),
        ]
        info_x = [accent_x, accent_x + 386, accent_x + 748]
        for x, (label, value) in zip(info_x, info_items):
            self.add_text(slide, (x, info_y, 310, 24), label, 18, "muted", line_ratio=1.0)
            self.add_text(slide, (x, info_y + 26, 330, 42), value, 28, "ink", line_ratio=1.0)
        department = str(meta.get("department") or meta.get("institution") or "").strip()
        if department:
            self.add_text(slide, (accent_x, info_y + 78, 700, 28), department, 20, "muted", line_ratio=1.0)
        cover_line = str(meta.get("cover_line") or "").strip()
        if cover_line:
            band_y = float(cover.get("bottom_band_y", 1002))
            band_height = float(cover.get("bottom_band_height", 92))
            self.add_rect(slide, (0, band_y, self.width, band_height), "blue_dark")
            self.add_rect(slide, (0, band_y, 10, band_height), "red")
            self.add_text(slide, (accent_x, band_y + 17, self.width - accent_x * 2, band_height - 22), cover_line, 26, "white", valign="middle", line_ratio=1.0)
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
        self.add_text(slide, (title_left, float(cover.get("title_y", 352)), title_width, 136), title, title_size, "blue", family=self.fonts["display"], align="center", line_ratio=1.10, name="cover_title")
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
        self.add_line(slide, (spine_x, top + 15), (spine_x, last_row_bottom - 15), "line", 2)
        for index, section in enumerate(sections):
            row_top = top + index * (row_height + gap)
            self.add_ellipse(slide, (spine_x - 6, row_top + row_height / 2 - 6, 12, 12), "blue_mid", None)
            self.add_panel(slide, (left, row_top, width, row_height), "white", "line", radius_px=9, line_width_px=1.6, shadow=True)
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
        self.add_header(slide, s)
        number = str(s.get("number", s.get("id", ""))).strip()[:2]
        self.add_text(slide, (176, 260, 300, 220), number, 210, "blue_pale", bold=True, line_ratio=1.0)
        self.add_rect(slide, (470, 292, 4, 464), "blue")
        title = str(s.get("section_title", s.get("title", "")))
        self.add_fit_text(slide, (520, 326, 1400, 180), title, "title", 78, 48, "blue_dark", bold=True, max_lines=2, line_ratio=1.05)
        summary = str(s.get("summary", ""))
        if summary:
            self.add_fit_text(slide, (524, 496, 1320, 140), summary, "body", 30, 20, "ink", max_lines=3, line_ratio=1.2)
        keywords = s.get("keywords") or []
        if keywords:
            x = 524
            for keyword in keywords[:4]:
                tag_width = max(92, len(str(keyword)) * 24 + 34)
                self.add_panel(slide, (x, 680, tag_width, 42), "blue_wash", "line", radius_px=5, line_width_px=1)
                self.add_text(slide, (x, 680, tag_width, 42), str(keyword), 20, "blue_dark", bold=True, align="center", valign="middle", line_ratio=1.0)
                x += tag_width + 14
        self.add_rect(slide, (176, 876, 1694, 4), "blue")

    def draw_grid(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s.get("section"))
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
        self.add_module_header(slide, box, str(block.get("heading", "")), index=f"{index:02d}", icon=str(block.get("icon", "")))
        content_top = top + 70
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
                self.add_picture_fit(slide, image_path, image_box, name="block_image")
                text_width_available = (image_box[0] - text_left - 14)
            else:
                image_height = max(50, (content_bottom - content_top) * 0.38)
                self.add_picture_fit(slide, image_path, (left + 14, content_top, width - 28, image_height), name="block_image")
                content_top += image_height + 10
        stat = str(block.get("stat", "")).strip()
        if stat:
            self.add_text(slide, (text_left, content_top, text_width_available, 54), stat, 46, "blue", bold=True, line_ratio=1.0)
            content_top += 58
        formula = str(block.get("formula", "")).strip()
        formula_latex = str(block.get("formula_latex", "")).strip()
        if formula or formula_latex:
            formula_box = (text_left, content_top, text_width_available, 58)
            self.add_panel(slide, formula_box, "blue_wash", "line", radius_px=5, line_width_px=1)
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
            self.add_panel(slide, box, "blue_wash", "line", radius_px=6, line_width_px=1)
        pad = max(10, min(18, int(width * 0.018)))
        inner_left, inner_top = left + pad, top + pad
        inner_width, inner_height = width - pad * 2, height - pad * 2
        if label:
            label_width = max(64.0, len(label) * 15.0 + 22.0)
            self.add_panel(slide, (inner_left, inner_top, label_width, 25), "blue", None, radius_px=3, line_width_px=0)
            self.add_text(
                slide,
                (inner_left + 10, inner_top + 3, label_width - 16, 20),
                label,
                13,
                "white",
                bold=True,
                valign="middle",
                line_ratio=1.0,
                name="diagram_label",
            )
            inner_top += 34
            inner_height -= 34
        note = str(diagram.get("note", "")).strip()
        if note:
            inner_height -= 30
        kind = str(diagram.get("kind", "sequence")).strip().lower()
        items = [item for item in (diagram.get("items") or []) if isinstance(item, dict)]
        highlight = int(diagram.get("highlight", -1))
        inner_right = inner_left + inner_width
        inner_bottom = inner_top + inner_height

        if kind == "sequence":
            items = items[:5]
            count = max(1, len(items))
            gap = max(20, min(34, int(inner_width * 0.035)))
            node_width = max(64, (inner_width - gap * (count - 1)) // count)
            node_height = min(112, max(72, int(inner_height * 0.56)))
            node_top = inner_top + max(0, (inner_height - node_height) // 2)
            for index, item in enumerate(items):
                x = inner_left + index * (node_width + gap)
                is_focus = index == highlight
                self.add_panel(
                    slide,
                    (x, node_top, node_width, node_height),
                    "blue" if is_focus else "white",
                    "blue" if is_focus else "line",
                    radius_px=5,
                    line_width_px=2 if is_focus else 1,
                )
                self.add_text(
                    slide,
                    (x + 10, node_top + 9, 34, 22),
                    f"{index + 1:02d}",
                    13,
                    "white" if is_focus else "blue_dark",
                    bold=True,
                    line_ratio=1.0,
                    name="diagram_index",
                )
                self.add_fit_text(
                    slide,
                    (x + 10, node_top + 40, node_width - 20, 30),
                    str(item.get("title") or item.get("label") or ""),
                    "title",
                    18,
                    12,
                    "white" if is_focus else "blue_dark",
                    bold=True,
                    align="center",
                    max_lines=2,
                    line_ratio=1.05,
                    name="diagram_node_title",
                )
                body = str(item.get("body", "")).strip()
                if body:
                    self.add_fit_text(
                        slide,
                        (x + 10, node_top + 72, node_width - 20, node_height - 80),
                        body,
                        "body",
                        14,
                        10,
                        "blue_pale" if is_focus else "muted",
                        align="center",
                        max_lines=2,
                        line_ratio=1.08,
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
            root_width = int(inner_width * 0.31)
            root_height = max(100, inner_height - 36)
            root_top = inner_top + (inner_height - root_height) / 2
            self.add_panel(slide, (inner_left, root_top, root_width, root_height), "blue", "blue", radius_px=5, line_width_px=1)
            self.add_fit_text(
                slide,
                (inner_left + 12, root_top + 18, root_width - 24, max(48, root_height * 0.4)),
                str(root.get("title") or diagram.get("center") or "共同问题"),
                "title",
                24,
                14,
                "white",
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
                    (inner_left + 12, root_top + root_height * 0.46, root_width - 24, root_height * 0.46),
                    root_body,
                    "body",
                    16,
                    11,
                    "blue_pale",
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
                )
                self.add_text(
                    slide,
                    (item_left + 10, y + 9, 36, 22),
                    f"{index + 1:02d}",
                    13,
                    "blue_dark",
                    bold=True,
                    line_ratio=1.0,
                    name="diagram_index",
                )
                self.add_fit_text(
                    slide,
                    (item_left + 52, y + 7, item_width - 62, 28),
                    str(item.get("title") or item.get("label") or ""),
                    "title",
                    17,
                    12,
                    "blue_dark",
                    bold=True,
                    max_lines=1,
                    name="diagram_branch_title",
                )
                body = str(item.get("body", "")).strip()
                if body:
                    self.add_fit_text(
                        slide,
                        (item_left + 52, y + 35, item_width - 62, item_height - 42),
                        body,
                        "body",
                        14,
                        10,
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
                    "blue" if is_focus else ("white" if index % 2 == 0 else "blue_wash"),
                    "blue_mid" if is_focus else "line",
                    radius_px=4,
                    line_width_px=1,
                )
                self.add_rect(slide, (inner_left + indent, y, 5, row_height), "red" if is_focus else "blue")
                self.add_text(
                    slide,
                    (inner_left + indent + 12, y + max(7, (row_height - 22) / 2), 34, 22),
                    f"{index + 1:02d}",
                    13,
                    "white" if is_focus else "blue_dark",
                    bold=True,
                    line_ratio=1.0,
                    name="diagram_index",
                )
                text_left = inner_left + indent + 64
                title_width = (inner_right - text_left) * 0.34
                self.add_fit_text(
                    slide,
                    (text_left, y + 4, title_width, row_height - 8),
                    str(item.get("title") or item.get("label") or ""),
                    "title",
                    17,
                    12,
                    "white" if is_focus else "blue_dark",
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
                        "blue_pale" if is_focus else "muted",
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
                    "blue" if is_focus else ("white" if row == column else "blue_wash"),
                    "red" if is_focus else "line",
                    radius_px=4,
                    line_width_px=2 if is_focus else 1,
                )
                self.add_text(
                    slide,
                    (x + 9, y + 8, 36, 22),
                    f"{index + 1:02d}",
                    13,
                    "white" if is_focus else "blue_dark",
                    bold=True,
                    line_ratio=1.0,
                    name="diagram_index",
                )
                self.add_fit_text(
                    slide,
                    (x + 34, y + 7, cell_width - 44, 32),
                    str(item.get("title") or item.get("label") or ""),
                    "title",
                    16,
                    11,
                    "white" if is_focus else "blue_dark",
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
                        "blue_pale" if is_focus else "ink",
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
                    "blue" if is_focus else "white",
                    "red" if is_focus else "line",
                    radius_px=5,
                    line_width_px=2 if is_focus else 1,
                )
                self.add_fit_text(
                    slide,
                    (point[0] - node_width / 2 + 8, point[1] - 23, node_width - 16, 46),
                    str(item.get("title") or item.get("label") or ""),
                    "title",
                    16,
                    11,
                    "white" if is_focus else "blue_dark",
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
                (inner_left, inner_top + inner_height + 5, inner_width, 24),
                note,
                "body",
                13,
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
        transform = etree.XSLT(etree.parse(str(self.xsl_path)))
        result = transform(mathml)
        root = result.getroot()
        if root is None:
            raise RuntimeError("LaTeX to OMML conversion returned no root")
        self.formula_cache[latex] = root
        return root

    def draw_concept(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s.get("section"))
        diagram = s.get("diagram") or {}
        panels = s.get("panels") or []
        gap = 18
        if panels:
            diagram_ratio = max(0.48, min(0.68, float(s.get("diagram_width_ratio", 0.58))))
            diagram_width = (right - left) * diagram_ratio
            panel_width = right - left - diagram_width - gap
        else:
            diagram_width = right - left
            panel_width = 0
        self.draw_diagram_native(
            slide,
            (left, body_top, diagram_width, bottom - body_top),
            diagram,
            label=str(s.get("diagram_label", "机制示意")),
            frame=True,
        )
        if panels:
            panel_left = left + diagram_width + gap
            panel_gap = 12
            panel_height = (bottom - body_top - panel_gap * (len(panels) - 1)) / max(1, len(panels))
            for index, panel in enumerate(panels):
                panel_top = body_top + index * (panel_height + panel_gap)
                box = (panel_left, panel_top, panel_width, panel_height)
                self.add_panel(slide, box, "white" if index % 2 == 0 else "blue_wash", "line", radius_px=self.panel_radius, line_width_px=1)
                self.add_module_header(slide, box, str(panel.get("heading", "")), index=f"{index + 1:02d}", icon=str(panel.get("icon", "")), height=48)
                content_top = panel_top + 58
                body = str(panel.get("body", "")).strip()
                bullets = panel.get("bullets") or []
                if body:
                    body_size, _ = self.measure(body, "body", int(panel_width - 28), max(40, int(panel_height - 70)), self.type_size("body", 19), self.type_size("small", 13), max_lines=4 if not bullets else 3, line_ratio=1.2)
                    self.add_text(slide, (panel_left + 14, content_top, panel_width - 28, max(36, panel_top + panel_height - content_top - 10)), body, body_size, "ink", line_ratio=1.2)
                    content_top += min(panel_top + panel_height - content_top - 10, max(34, body_size * 1.3 * max(1, math.ceil(len(strip_markup(body)) / max(8, (panel_width - 28) / max(1, body_size)))))) + 8
                if bullets:
                    self.add_bullets(slide, bullets, (panel_left + 14, content_top, panel_width - 28, max(24, panel_top + panel_height - content_top - 10)), max_size=self.type_size("bullet", 16), min_size=self.type_size("micro", 11), gap_px=4)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_formula(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s.get("section"))
        gap = 18
        left_ratio = max(0.52, min(0.70, float(s.get("formula_width_ratio", 0.62))))
        left_width = (right - left) * left_ratio
        right_left = left + left_width + gap
        panel_height = int(s.get("formula_panel_height", 236))
        formula_box = (left, body_top, left_width, panel_height)
        self.add_panel(slide, formula_box, "blue_wash", "line", radius_px=self.panel_radius, line_width_px=1)
        self.add_rect(slide, (left, body_top, 6, panel_height), "blue")
        label = str(s.get("formula_label", "核心关系"))
        label_width = max(80.0, len(label) * 16.0 + 28.0)
        self.add_panel(slide, (left + 18, body_top + 14, label_width, 30), "blue", None, radius_px=3, line_width_px=0)
        self.add_text(slide, (left + 32, body_top + 19, label_width - 24, 22), label, 16, "white", bold=True, valign="middle", line_ratio=1.0, name="formula_label")
        formula = str(s.get("formula_latex") or s.get("formula") or "").strip()
        condition = str(s.get("formula_condition", "")).strip()
        math_bottom = body_top + panel_height - (58 if condition else 22)
        self.add_formula(slide, formula, (left + 28, body_top + 58, left_width - 56, max(50, math_bottom - body_top - 58)), 46)
        if condition:
            self.add_panel(slide, (left + 18, body_top + panel_height - 42, left_width - 36, 32), "blue_pale", "line", radius_px=3, line_width_px=1)
            self.add_fit_text(slide, (left + 28, body_top + panel_height - 37, left_width - 56, 24), condition, "body", 14, 10, "ink", max_lines=1, line_ratio=1.05, name="formula_condition")

        variables_top = body_top + panel_height + 18
        variables = [item for item in (s.get("variables") or []) if isinstance(item, dict)]
        self.add_text(slide, (left, variables_top, left_width, 28), str(s.get("variables_title", "符号与条件")), 18, "blue_dark", bold=True, line_ratio=1.0, name="formula_variables_title")
        self.add_line(slide, (left, variables_top + 30), (left + left_width, variables_top + 30), "line", 2)
        if variables:
            columns = 2 if len(variables) > 2 else 1
            column_gap = 18
            column_width = (left_width - column_gap * (columns - 1)) / columns
            rows = math.ceil(len(variables) / columns)
            row_height = max(54, min(74, (bottom - variables_top - 50) / max(1, rows)))
            for index, item in enumerate(variables):
                column = index % columns
                row = index // columns
                x = left + column * (column_width + column_gap)
                y = variables_top + 42 + row * row_height
                self.add_line(slide, (x, y + row_height - 6), (x + column_width, y + row_height - 6), "line", 1)
                symbol = str(item.get("symbol", "")).strip()
                meaning = str(item.get("meaning", "")).strip()
                self.add_text(slide, (x, y, column_width, 32), symbol, 26, "blue", family=self.fonts["formula"], line_ratio=1.0, name="formula_symbol")
                self.add_fit_text(slide, (x, y + 32, column_width, max(20, row_height - 42)), meaning, "body", 16, 11, "ink", max_lines=2, line_ratio=1.15, name="formula_meaning")

        panels = s.get("panels") or []
        if panels:
            panel_gap = 12
            panel_height_right = (bottom - body_top - panel_gap * (len(panels) - 1)) / max(1, len(panels))
            panel_width = right - right_left
            for index, panel in enumerate(panels):
                panel_top = body_top + index * (panel_height_right + panel_gap)
                box = (right_left, panel_top, panel_width, panel_height_right)
                self.add_panel(slide, box, "white" if index % 2 == 0 else "blue_wash", "line", radius_px=self.panel_radius, line_width_px=1)
                self.add_module_header(slide, box, str(panel.get("heading", "")), index=f"{index + 1:02d}", icon=str(panel.get("icon", "")), height=48)
                content_top = panel_top + 58
                body = str(panel.get("body", "")).strip()
                bullets = panel.get("bullets") or []
                if body:
                    body_size, _ = self.measure(body, "body", int(panel_width - 28), max(36, int(panel_height_right - 72)), self.type_size("body", 18), self.type_size("small", 12), max_lines=4 if not bullets else 3, line_ratio=1.18)
                    self.add_text(slide, (right_left + 14, content_top, panel_width - 28, max(32, panel_top + panel_height_right - content_top - 10)), body, body_size, "ink", line_ratio=1.18)
                    content_top += min(panel_top + panel_height_right - content_top - 10, max(30, body_size * 1.25 * max(1, math.ceil(len(strip_markup(body)) / max(8, (panel_width - 28) / max(1, body_size)))))) + 8
                if bullets:
                    self.add_bullets(slide, bullets, (right_left + 14, content_top, panel_width - 28, max(22, panel_top + panel_height_right - content_top - 10)), max_size=self.type_size("bullet", 16), min_size=self.type_size("micro", 11), gap_px=4)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_pipeline(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s.get("section"))
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
            self.add_ellipse(slide, (center - 29, line_top - 29, 58, 58), "blue", "white", 4)
            self.add_text(slide, (center - 24, line_top - 15, 48, 30), str(index + 1), 24, "white", bold=True, align="center", valign="middle", line_ratio=1.0)
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
        self.add_sidebar(slide, s.get("section"))
        columns = s.get("columns") or []
        rows = s.get("rows") or []
        if not columns:
            self.warn(str(s.get("id", "")), "Compare table has no columns.")
            return
        table_theme = self.theme["table"]
        header_height = int(table_theme.get("header_height", 66))
        row_height = min(112, max(80, (bottom - body_top - header_height - 4) // max(1, len(rows))))
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
        table.first_row = False
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
            cell.fill.fore_color.rgb = self.rgb("blue_dark")
            cell.margin_left = px(10)
            cell.margin_right = px(10)
            cell.margin_top = px(4)
            cell.margin_bottom = px(4)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            self.add_cell_text(cell, str(column), 25, "white", bold=True, align="center")
            set_cell_border(cell, self.rgb("white"), 0.7)
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
                    23 if column_index == 0 else 22,
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
            self.add_bullets(slide, notes[:half], (left, note_top, note_width, 54), max_size=17, min_size=12, gap_px=3)
            self.add_bullets(slide, notes[half:], (left + note_width + 14, note_top, note_width, 54), max_size=17, min_size=12, gap_px=3)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def add_cell_text(self, cell, text: str, size_px: int, color_name: str, bold: bool = False, align: str = "left") -> None:
        tf = cell.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        paragraph = tf.paragraphs[0]
        paragraph.alignment = align_value(align)
        paragraph.line_spacing = 1.08
        run = paragraph.add_run()
        run.text = strip_markup(text)
        self.apply_font(run, self.fonts["title"] if bold else self.fonts["body"], size_px, self.rgb(color_name), bold=bold)

    def draw_timeline(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s.get("section"))
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
                self.add_panel(slide, (x + 14, milestone_top, card_width - 28, milestone_height), "blue_wash", "line", radius_px=4, line_width_px=1)
                self.add_panel(slide, (x + 14, milestone_top, 74, milestone_height), "blue", None, radius_px=4, line_width_px=0)
                self.add_text(slide, (x + 20, milestone_top + 18, 64, 24), "里程碑", 14, "white", bold=True, line_ratio=1.0)
                self.add_text(slide, (x + 96, milestone_top + 9, card_width - 112, milestone_height - 18), milestone, 16, "blue_dark", valign="middle", line_ratio=1.08)
        if timeline_note and card_bottom + 18 < bottom:
            note_top = card_bottom + 18
            self.add_panel(slide, (left + 30, note_top, right - left - 60, note_height), "blue_wash", "line", radius_px=self.panel_radius, line_width_px=1)
            self.add_icon(slide, "alert", (left + 46, note_top + 9, 36, 38), "blue")
            self.add_text(slide, (left + 92, note_top + 7, right - left - 160, note_height - 14), timeline_note, 18, "blue_dark", valign="middle", line_ratio=1.05)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_hub(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s.get("section"))
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
        self.add_sidebar(slide, s.get("section"))
        references = s.get("references") or []
        if not references:
            references = [{"authors": "", "title": "参考文献待补充", "venue": "", "doi": ""}]
        column_gap = 44
        column_width = (right - left - column_gap) / 2
        split = math.ceil(len(references) / 2)
        columns = [(references[:split], left), (references[split:], left + column_width + column_gap)]
        divider_x = left + column_width + column_gap / 2
        self.add_line(slide, (divider_x, body_top + 6), (divider_x, bottom - 6), "line", 1)
        for column_index, (items, x) in enumerate(columns):
            if not items:
                continue
            start_number = 1 + (0 if column_index == 0 else split)
            self.add_text(slide, (x, body_top, column_width, 24), f"{s.get('reference_label', '参考文献')}  {start_number:02d}–{start_number + len(items) - 1:02d}", 16, "muted", line_ratio=1.0)
            self.add_line(slide, (x, body_top + 26), (x + column_width, body_top + 26), "line", 1)
            list_top = body_top + 40
            item_gap = 16
            item_height = (bottom - list_top - item_gap * (len(items) - 1)) / len(items)
            for item_index, reference in enumerate(items):
                item_top = list_top + item_index * (item_height + item_gap)
                number = start_number + item_index
                self.add_text(slide, (x, item_top + 2, 44, 32), f"{number:02d}", 28, "blue_pale", bold=True, line_ratio=1.0)
                text_left = x + 52
                text_width_available = column_width - 52
                authors = str(reference.get("authors", "")).strip()
                cursor = item_top
                if authors:
                    self.add_text(slide, (text_left, cursor, text_width_available, 24), authors, 18, "muted", line_ratio=1.0)
                    cursor += 24
                title = str(reference.get("title", "")).strip()
                title_size, _ = self.measure(title, "title", int(text_width_available), 58, 20, 14, max_lines=2, line_ratio=1.12)
                self.add_text(slide, (text_left, cursor, text_width_available, 58), title, title_size, "blue_dark", bold=False, line_ratio=1.12)
                cursor += min(58, title_size * 1.12 * max(1, math.ceil(len(title) / max(14, text_width_available / max(1, title_size))))) + 4
                venue = str(reference.get("venue", "")).strip()
                if venue:
                    self.add_text(slide, (text_left, cursor, text_width_available, 22), venue, 17, "ink", line_ratio=1.0)
                    cursor += 23
                doi = str(reference.get("doi", "")).strip()
                if doi:
                    self.add_text(slide, (text_left, cursor, text_width_available, 20), doi, 15, "muted", line_ratio=1.0)
                if item_index < len(items) - 1:
                    self.add_line(slide, (x, item_top + item_height - 4), (x + column_width, item_top + item_height - 4), "line", 1)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_closing(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s.get("section"))
        points = s.get("points") or ["结论待补充"]
        row_height = 104
        gap = 14
        points_bottom = body_top + len(points) * (row_height + gap) - gap
        for index, point in enumerate(points):
            row_top = body_top + index * (row_height + gap)
            self.add_panel(slide, (left, row_top, right - left, row_height), "blue_wash" if index == 0 else "white", "line", radius_px=8, line_width_px=1.4, shadow=True)
            self.add_rect(slide, (left, row_top, 7, row_height), "red" if index == 0 else "blue")
            self.add_text(slide, (left + 24, row_top + 20, 80, 44), f"{index + 1:02d}", 44, "blue_pale", bold=True, line_ratio=1.0)
            self.add_fit_text(slide, (left + 108, row_top + 16, right - left - 136, row_height - 28), str(point), "body", self.type_size("lead", 27), 17, "ink", max_lines=2, line_ratio=1.1)
        cards = []
        if s.get("boundary") or s.get("boundary_points"):
            cards.append(("边界", "alert", str(s.get("boundary", "")), s.get("boundary_points") or [], "blue_wash"))
        if s.get("next") or s.get("next_points"):
            cards.append(("下一步", "arrow", str(s.get("next", "")), s.get("next_points") or [], "white"))
        cards_top = points_bottom + 16
        cards_bottom = min(bottom - 120, cards_top + 190)
        if cards and cards_bottom - cards_top > 80:
            gap_cards = 16
            card_width = (right - left - gap_cards * (len(cards) - 1)) / len(cards)
            for index, (title, icon, body, points, fill) in enumerate(cards):
                x = left + index * (card_width + gap_cards)
                self.add_panel(slide, (x, cards_top, card_width, cards_bottom - cards_top), fill, "line", radius_px=8, line_width_px=1.2, shadow=True)
                self.add_icon(slide, icon, (x + 18, cards_top + 18, 32, 32), "red" if index == 0 else "blue")
                self.add_text(slide, (x + 62, cards_top + 18, card_width - 78, 28), title, 20, "blue_dark", bold=True, line_ratio=1.0)
                cursor = cards_top + 56
                if body:
                    self.add_fit_text(slide, (x + 18, cursor, card_width - 36, 54), body, "body", 18, 13, "ink", max_lines=2, line_ratio=1.1)
                    cursor += 54
                if points:
                    self.add_bullets(slide, points, (x + 18, cursor, card_width - 36, max(24, cards_bottom - cursor - 12)), max_size=16, min_size=11, gap_px=3)
        closing = str(s.get("closing", "谢谢各位老师")).strip()
        closing_y = min(cards_bottom + 42 if cards and cards_bottom - cards_top > 80 else bottom - 62, bottom - 52)
        self.add_line(slide, (left, closing_y - 24), (right, closing_y - 24), "line", 1)
        self.add_fit_text(slide, (left + 40, closing_y, right - left - 80, 58), closing, "title", 48, 28, "blue", bold=True, align="center", max_lines=1, line_ratio=1.0)

    def draw_bullets(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s.get("section"))
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
            self.add_module_header(slide, box, str(column.get("heading", "")), index=f"{index + 1:02d}", icon=str(column.get("icon", "")))
            content_top = body_top + 68
            body = str(column.get("body", "")).strip()
            if body:
                self.add_fit_text(slide, (x + 18, content_top, width - 36, 100), body, "body", self.type_size("body", 22), self.type_size("small", 16), max_lines=3, line_ratio=1.12)
                content_top += 112
            self.add_bullets(slide, column.get("bullets") or [], (x + 18, content_top, width - 36, bottom - content_top - 18), max_size=self.type_size("bullet", 21), min_size=self.type_size("micro", 14))
            x += width + gap
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_figure(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s.get("section"))
        image_path = resolve_path(s.get("image"), self.deck_path.parent)
        side = str(s.get("image_side", "left"))
        gap = 20
        total_width = right - left
        image_width = round(total_width * 0.66)
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
            self.add_panel(slide, image_box, "white", "line", radius_px=8, line_width_px=1.2, shadow=True)
            if image_path and image_path.exists():
                self.add_picture_fit(
                    slide,
                    image_path,
                    (image_box[0] + 14, body_top + 14, image_width - 28, bottom - body_top - 28),
                    name="hero_figure",
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
        gap = 14
        panel_height = (panel_height_total - gap * (len(panels) - 1)) / len(panels)
        for index, panel in enumerate(panels):
            current_top = panel_top + index * (panel_height + gap)
            box = (panel_left, current_top, panel_width, panel_height)
            self.add_panel(slide, box, "white", "line", radius_px=8, line_width_px=1.2, shadow=True)
            self.add_module_header(slide, box, str(panel.get("heading", "")), index=f"{index + 1:02d}", icon=str(panel.get("icon", "")), height=46)
            self.add_fit_text(
                slide,
                (panel_left + 16, current_top + 58, panel_width - 32, panel_height - 72),
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
        self.add_sidebar(slide, s.get("section"))
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
                (x + 74, body_top + 14, flow_width - 86, 38),
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
                (x + 74, body_top + 52, flow_width - 86, 46),
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
        image_width = round((right - left) * 0.60)
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
            self.add_panel(slide, figure_box, "white", "line", radius_px=8, line_width_px=1.2, shadow=True)
            if image_path and image_path.exists():
                self.add_picture_fit(
                    slide,
                    image_path,
                    (left + 14, lower_top + 14, image_width - 28, bottom - lower_top - 28),
                    name="flow_figure",
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
        gap = 14
        panel_height = (bottom - lower_top - gap * (len(panels) - 1)) / len(panels)
        for index, panel in enumerate(panels):
            current_top = lower_top + index * (panel_height + gap)
            box = (panel_left, current_top, panel_width, panel_height)
            self.add_panel(slide, box, "white", "line", radius_px=8, line_width_px=1.2, shadow=True)
            self.add_module_header(slide, box, str(panel.get("heading", "")), index=f"{index + 1:02d}", icon=str(panel.get("icon", "")), height=46)
            body = str(panel.get("body", ""))
            bullets = panel.get("bullets") or []
            self.add_fit_text(
                slide,
                (panel_left + 16, current_top + 58, panel_width - 32, 68),
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
                    (panel_left + 16, current_top + 128, panel_width - 32, max(24, panel_height - 140)),
                    max_size=16,
                    min_size=11,
                    gap_px=3,
                )
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_metrics(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s.get("section"))
        metrics = s.get("metrics") or []
        gap = 18
        card_width = (right - left - gap * (len(metrics) - 1)) / max(1, len(metrics))
        for index, metric in enumerate(metrics):
            x = left + index * (card_width + gap)
            box = (x, body_top, card_width, 176)
            self.add_panel(slide, box, "blue_wash" if index == 0 else "white", "line", radius_px=10, line_width_px=1.4, shadow=True)
            self.add_rect(slide, (x + 2, body_top + 2, card_width - 4, 5), "blue" if index == 0 else "blue_mid")
            self.add_text(slide, (x + 20, body_top + 20, card_width - 40, 62), str(metric.get("value", "")), 58, "blue", bold=True, line_ratio=1.0)
            self.add_text(slide, (x + 20, body_top + 90, card_width - 40, 30), str(metric.get("label", "")), 22, "blue_dark", bold=True, line_ratio=1.0)
            self.add_text(slide, (x + 20, body_top + 126, card_width - 40, 30), str(metric.get("note", "")), 16, "muted", line_ratio=1.0)
        lower_top = body_top + 198
        conditions = s.get("conditions") or []
        if conditions:
            self.add_text(slide, (left, lower_top, 260, 26), str(s.get("conditions_title", "实验条件")), 18, "muted", bold=True, line_ratio=1.0)
            cursor = lower_top + 34
            for condition in conditions:
                if isinstance(condition, dict):
                    label = str(condition.get("label", ""))
                    value = str(condition.get("value", ""))
                    self.add_text(slide, (left, cursor, 180, 22), label, 16, "muted", line_ratio=1.0)
                    self.add_text(slide, (left + 180, cursor, 200, 22), value, 16, "ink", line_ratio=1.0)
                else:
                    self.add_text(slide, (left, cursor, 420, 22), str(condition), 16, "ink", line_ratio=1.0)
                cursor += 28
        image_path = resolve_path(s.get("image"), self.deck_path.parent)
        evidence_box = (left + 460, lower_top, right - left - 460, bottom - lower_top - 10)
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
            self.add_picture_fit(slide, image_path, evidence_box, name="metrics_image")
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_balance(self, slide, s: dict) -> None:
        left, top, right, bottom = self.content_bounds(s)
        body_top = self.add_lead(slide, s.get("lead", ""), top, str(s.get("id", "")))
        self.add_sidebar(slide, s.get("section"))
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
            self.add_panel(slide, center_box, "white", "line", radius_px=8, line_width_px=1.2, shadow=True)
            image_path = resolve_path(s.get("image"), self.deck_path.parent)
            if image_path and image_path.exists():
                self.add_picture_fit(slide, image_path, (center_left + 18, body_top + 38, center_right - center_left - 36, bottom - body_top - 74), name="balance_figure")
            self.add_text(slide, (center_left + 20, bottom - 28, center_right - center_left - 40, 20), str(s.get("image_label", "系统结构")), 16, "muted", align="center", line_ratio=1.0)
        self.add_takeaway(slide, s.get("takeaway", ""), str(s.get("id", "")))
        self.add_footnote(slide, s.get("footnote", ""))

    def draw_balance_rail(self, slide, rail: dict, box: tuple[float, float, float, float], index: int) -> None:
        left, top, width, height = box
        self.add_panel(slide, box, "blue_wash" if index == 0 else "white", "line", radius_px=8, line_width_px=1.4, shadow=True)
        self.add_text(slide, (left + 14, top + 14, width - 28, 42), str(rail.get("heading", "")), 18, "blue_dark", bold=True, line_ratio=1.0)
        items = rail.get("items") or []
        item_top = top + 68
        item_height = (height - 80) / max(1, len(items))
        for index_item, item in enumerate(items):
            y = item_top + index_item * item_height
            self.add_text(slide, (left + 14, y, width - 28, 28), str(item.get("heading", "")), 19, "blue_dark", bold=True, line_ratio=1.0)
            self.add_fit_text(slide, (left + 14, y + 28, width - 28, max(22, item_height - 34)), str(item.get("body", "")), "body", 16, 11, max_lines=3, line_ratio=1.1)

    def export(self, pptx_name: str = "defense_deck.pptx") -> dict:
        self.presentation = Presentation()
        self.presentation.slide_width = px(self.width)
        self.presentation.slide_height = px(self.height)
        slides = self.deck.get("slides", [])
        for index, s in enumerate(slides, start=1):
            slide = self.new_slide(index)
            layout = s.get("layout")
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
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    exporter = EditableDeckExporter(Path(args.deck), Path(args.theme), Path(args.out))
    report = exporter.export(args.pptx_name)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
