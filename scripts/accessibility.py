from __future__ import annotations

from pathlib import Path
from typing import Iterable

from PIL import Image
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from pptx.oxml.ns import _nsmap


NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
ADEC_NS = "http://schemas.microsoft.com/office/drawing/2017/decorative"
DECORATIVE_EXT_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"
_nsmap.setdefault("adec", ADEC_NS)


def _cnvpr(shape):
    """Return the non-visual properties element for a PowerPoint shape."""
    element = getattr(shape, "_element", None)
    if element is None:
        return None
    return element.find(".//" + qn("p:cNvPr"))


def get_alt_text(shape) -> str:
    cnvpr = _cnvpr(shape)
    return (cnvpr.get("descr") or "").strip() if cnvpr is not None else ""


def set_alt_text(shape, text: str) -> None:
    cnvpr = _cnvpr(shape)
    if cnvpr is not None:
        cnvpr.set("descr", str(text).strip())


def mark_decorative(shape) -> None:
    """Use PowerPoint's native 'mark as decorative' extension."""
    cnvpr = _cnvpr(shape)
    if cnvpr is None:
        return
    ext_lst = cnvpr.find(qn("a:extLst"))
    if ext_lst is None:
        ext_lst = parse_xml(f'<a:extLst xmlns:a="{NS_A}"/>')
        cnvpr.append(ext_lst)
    for ext in ext_lst.findall(qn("a:ext")):
        if ext.get("uri") == DECORATIVE_EXT_URI:
            ext_lst.remove(ext)
            break
    ext = parse_xml(
        f'<a:ext xmlns:a="{NS_A}" uri="{DECORATIVE_EXT_URI}">'
        f'<adec:decorative xmlns:adec="{ADEC_NS}" val="1"/>'
        "</a:ext>"
    )
    ext_lst.append(ext)


def is_decorative(shape) -> bool:
    cnvpr = _cnvpr(shape)
    if cnvpr is None:
        return False
    return bool(cnvpr.xpath(".//adec:decorative[@val='1']"))


def shape_has_visible_text(shape) -> bool:
    return bool(getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip())


def shape_needs_alt_text(shape) -> bool:
    if shape_has_visible_text(shape):
        return False
    if getattr(shape, "has_table", False) or getattr(shape, "has_chart", False):
        return False
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return False
    return True


def finalize_slide_accessibility(slide) -> None:
    """Mark non-text decoration without erasing semantic image descriptions."""
    for shape in slide.shapes:
        if shape_has_visible_text(shape):
            continue
        if getattr(shape, "has_table", False):
            continue
        if getattr(shape, "has_chart", False):
            if not get_alt_text(shape):
                set_alt_text(shape, "图表：请参见页内标题与数据说明")
            continue
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            if is_decorative(shape):
                continue
            alt = get_alt_text(shape)
            if not alt or alt.lower().endswith((".png", ".jpg", ".jpeg", ".webp", ".emf", ".wmf")):
                mark_decorative(shape)
            continue
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        mark_decorative(shape)


def _relative_luminance(rgb: tuple[int, int, int]) -> float:
    channels = []
    for value in rgb:
        channel = value / 255.0
        channels.append(channel / 12.92 if channel <= 0.04045 else ((channel + 0.055) / 1.055) ** 2.4)
    return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]


def contrast_ratio(foreground: tuple[int, int, int], background: tuple[int, int, int]) -> float:
    first = _relative_luminance(foreground)
    second = _relative_luminance(background)
    lighter, darker = max(first, second), min(first, second)
    return (lighter + 0.05) / (darker + 0.05)


def image_looks_blank(path: str | Path, threshold: float = 3.0) -> bool:
    with Image.open(path) as image:
        grayscale = image.convert("L").resize((256, 144))
        pixels = list(grayscale.getdata())
    mean = sum(pixels) / max(1, len(pixels))
    variance = sum((value - mean) ** 2 for value in pixels) / max(1, len(pixels))
    return variance < threshold


def summarize_alt_text(shapes: Iterable) -> dict[str, int]:
    counts = {"pictures": 0, "with_alt": 0, "decorative": 0, "filename_alt": 0}
    for shape in shapes:
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        counts["pictures"] += 1
        if is_decorative(shape):
            counts["decorative"] += 1
            continue
        alt = get_alt_text(shape)
        if alt:
            counts["with_alt"] += 1
            if Path(alt).suffix.lower() in {".png", ".jpg", ".jpeg", ".webp", ".emf", ".wmf"}:
                counts["filename_alt"] += 1
    return counts
