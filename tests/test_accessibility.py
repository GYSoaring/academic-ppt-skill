from __future__ import annotations

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from accessibility import (
    contrast_ratio,
    get_alt_text,
    is_decorative,
    mark_decorative,
    set_alt_text,
)


def test_alt_text_round_trip() -> None:
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 100, 100)
    set_alt_text(shape, "系统结构示意图")
    assert get_alt_text(shape) == "系统结构示意图"


def test_decorative_round_trip() -> None:
    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    shape = slide.shapes.add_textbox(0, 0, 100, 100)
    assert not is_decorative(shape)
    mark_decorative(shape)
    assert is_decorative(shape)


def test_contrast_ratio_is_symmetric() -> None:
    dark = (0, 0, 0)
    light = (255, 255, 255)
    assert round(contrast_ratio(dark, light), 2) == 21.0
    assert round(contrast_ratio(light, dark), 2) == 21.0
