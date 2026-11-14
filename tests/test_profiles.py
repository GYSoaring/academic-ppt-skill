from __future__ import annotations

from common import PROJECTION_TYPE_SCALE, apply_type_profile


def test_projection_profile_applies_overrides() -> None:
    theme = {"type_scale": {"body": 22, "small": 17}}
    apply_type_profile(theme, "projection")
    assert theme["active_type_profile"] == "projection"
    assert theme["type_scale"]["body"] == PROJECTION_TYPE_SCALE["body"]
    assert theme["type_scale"]["small"] == PROJECTION_TYPE_SCALE["small"]


def test_document_profile_keeps_theme_values() -> None:
    theme = {"type_scale": {"body": 22, "small": 17}}
    apply_type_profile(theme, "document")
    assert theme["active_type_profile"] == "document"
    assert theme["type_scale"]["body"] == 22
    assert theme["type_scale"]["small"] == 17
