from __future__ import annotations

import json
from pathlib import Path

import pytest
from PIL import Image, ImageDraw

from build_brand_lockup import build, white_to_alpha


def find_test_font() -> str | None:
    candidates = [
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return candidate
    return None


def make_emblem(path: Path) -> None:
    image = Image.new("RGBA", (240, 240), (255, 255, 255, 255))
    draw = ImageDraw.Draw(image)
    draw.ellipse((20, 20, 220, 220), outline=(91, 11, 99, 255), width=16)
    draw.ellipse((100, 100, 140, 140), fill=(91, 11, 99, 255))
    image.save(path)


def test_white_to_alpha_removes_background() -> None:
    image = Image.new("RGBA", (40, 40), (255, 255, 255, 255))
    draw = ImageDraw.Draw(image)
    draw.ellipse((8, 8, 31, 31), fill=(91, 11, 99, 255))
    result = white_to_alpha(image)
    assert result.getpixel((0, 0))[3] == 0
    assert result.getpixel((20, 20))[3] == 255


def test_build_generates_emblem_and_lockup(work_tmp: Path) -> None:
    font = find_test_font()
    if not font:
        pytest.skip("No test font available.")
    theme_source = Path(__file__).resolve().parents[1] / "assets" / "themes" / "academic_blue" / "theme.json"
    theme = json.loads(theme_source.read_text(encoding="utf-8"))
    theme_path = work_tmp / "theme.json"
    theme_path.write_text(json.dumps(theme, ensure_ascii=False), encoding="utf-8")
    emblem_path = work_tmp / "emblem.png"
    make_emblem(emblem_path)
    report = build(
        theme_path,
        "Test University",
        str(emblem_path),
        latin="TEST UNIVERSITY",
        font=font,
        remove_white=True,
        update_theme=True,
    )
    assert Path(report["emblem"]).exists()
    assert Path(report["lockup"]).exists()
    updated = json.loads(theme_path.read_text(encoding="utf-8"))
    assert updated["brand"]["lockup"] == "brand_lockup.png"
    assert updated["brand"]["emblem"] == "emblem_mark.png"
    assert updated["brand"]["fallback_text"] == "Test University"
