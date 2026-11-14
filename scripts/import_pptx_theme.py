from __future__ import annotations

import argparse
import colorsys
import json
import re
import zipfile
from pathlib import Path

from lxml import etree

from common import hex_rgb, load_json, mix_color, write_json


SKILL_ROOT = Path(__file__).resolve().parent.parent
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Import colors and fonts from an existing PPTX theme.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--out", required=True, help="theme.json path or output directory.")
    parser.add_argument("--name", default="")
    parser.add_argument(
        "--base-theme",
        default=str(SKILL_ROOT / "assets" / "themes" / "academic_blue" / "theme.json"),
        help="Base theme whose layout geometry is reused.",
    )
    return parser.parse_args()


def rgb_hex(value: str, fallback: str) -> str:
    value = (value or "").strip().lstrip("#")
    if re.fullmatch(r"[0-9a-fA-F]{6}", value):
        return f"#{value.upper()}"
    return fallback


def to_hex(value: tuple[int, int, int]) -> str:
    return "#" + "".join(f"{channel:02X}" for channel in value)


def mix(first: str, second: str, amount: float) -> str:
    return to_hex(mix_color(hex_rgb(first), hex_rgb(second), amount))


def is_reddish(value: str) -> bool:
    red, green, blue = (channel / 255 for channel in hex_rgb(value))
    hue, saturation, _ = colorsys.rgb_to_hsv(red, green, blue)
    return saturation > 0.28 and (hue < 0.075 or hue > 0.94)


def font_candidates(family: str) -> list[str]:
    value = (family or "").casefold()
    candidates: list[str] = []
    if any(token in value for token in ("yahei", "微软雅黑")):
        candidates += ["C:/Windows/Fonts/msyhbd.ttc", "C:/Windows/Fonts/msyh.ttc"]
    if any(token in value for token in ("simhei", "黑体")):
        candidates.append("C:/Windows/Fonts/simhei.ttf")
    if any(token in value for token in ("simsun", "宋体")):
        candidates.append("C:/Windows/Fonts/simsun.ttc")
    if any(token in value for token in ("stsong", "华文宋体")):
        candidates.append("C:/Windows/Fonts/STSONG.TTF")
    if any(token in value for token in ("stzhongs", "华文中宋")):
        candidates.append("C:/Windows/Fonts/STZHONGS.TTF")
    if any(token in value for token in ("deng", "等线")):
        candidates.append("C:/Windows/Fonts/Deng.ttf")
    if "noto sans sc" in value or "noto sans cjk" in value:
        candidates += [
            "C:/Windows/Fonts/Noto Sans SC (TrueType).otf",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        ]
    if "noto serif" in value:
        candidates += [
            "C:/Windows/Fonts/Noto Serif SC (TrueType).otf",
            "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
        ]
    if "pingfang" in value:
        candidates.append("/System/Library/Fonts/PingFang.ttc")
    if "arial" in value:
        candidates.append("C:/Windows/Fonts/arial.ttf")
    if "times" in value:
        candidates.append("C:/Windows/Fonts/times.ttf")
    if "cambria" in value:
        candidates.append("C:/Windows/Fonts/cambria.ttc")
    if "calibri" in value:
        candidates += ["C:/Windows/Fonts/calibrib.ttf", "C:/Windows/Fonts/calibri.ttf"]
    return candidates


def read_pptx_theme(pptx_path: Path) -> dict:
    with zipfile.ZipFile(pptx_path) as archive:
        theme_parts = sorted(
            name
            for name in archive.namelist()
            if name.startswith("ppt/theme/theme") and name.endswith(".xml")
        )
        if not theme_parts:
            raise ValueError("PPTX has no ppt/theme/theme*.xml part.")
        root = etree.fromstring(archive.read(theme_parts[0]))
    colors: dict[str, str] = {}
    scheme = root.find(".//a:clrScheme", NS)
    if scheme is not None:
        for child in scheme:
            name = etree.QName(child).localname
            srgb = child.find("a:srgbClr", NS)
            sys_color = child.find("a:sysClr", NS)
            value = srgb.get("val") if srgb is not None else sys_color.get("lastClr") if sys_color is not None else None
            if value:
                colors[name] = rgb_hex(value, "#000000")
    fonts: dict[str, str] = {}
    font_scheme = root.find(".//a:fontScheme", NS)
    if font_scheme is not None:
        for role in ("majorFont", "minorFont"):
            node = font_scheme.find(f"a:{role}", NS)
            if node is None:
                continue
            latin = node.find("a:latin", NS)
            east_asian = node.find("a:ea", NS)
            selected = (east_asian.get("typeface") if east_asian is not None else "") or (
                latin.get("typeface") if latin is not None else ""
            )
            if selected:
                fonts[role] = selected
            if latin is not None and latin.get("typeface"):
                fonts[f"{role}_latin"] = latin.get("typeface")
    return {"theme_part": theme_parts[0], "colors": colors, "fonts": fonts}


def build_theme(pptx_path: Path, base_theme_path: Path, name: str) -> dict:
    source = read_pptx_theme(pptx_path)
    base = load_json(base_theme_path)
    palette = source["colors"]
    paper = palette.get("lt1", "#F9FBFD")
    ink = palette.get("dk1", "#172B3F")
    accent1 = palette.get("accent1", "#174A7E")
    accent2 = palette.get("accent2", "#B3261E")
    accent3 = palette.get("accent3", "#2C735C")
    theme_name = re.sub(r"[^\w.-]+", "_", name or pptx_path.stem).strip("_").lower() or "imported_theme"
    theme = base
    theme["name"] = theme_name
    theme["display_name"] = name or pptx_path.stem
    theme["colors"] = {
        "paper": paper,
        "paper_alt": mix(accent1, paper, 0.9),
        "white": paper,
        "ink": ink,
        "muted": mix(ink, paper, 0.55),
        "blue": accent1,
        "blue_dark": mix(accent1, "#000000", 0.28),
        "blue_mid": accent2 if not is_reddish(accent2) else mix(accent1, paper, 0.35),
        "blue_bright": mix(accent1, paper, 0.2),
        "blue_pale": mix(accent1, paper, 0.82),
        "blue_wash": mix(accent1, paper, 0.93),
        "line": mix(accent1, paper, 0.62),
        "red": accent2 if is_reddish(accent2) else "#B3261E",
        "green": accent3 if not is_reddish(accent3) else "#1F7A5C",
    }
    major = source["fonts"].get("majorFont", "")
    minor = source["fonts"].get("minorFont", "")
    major_latin = source["fonts"].get("majorFont_latin", "")
    minor_latin = source["fonts"].get("minorFont_latin", "")
    theme["pptx_fonts"] = {
        "title": major or "Microsoft YaHei",
        "body": minor or "Microsoft YaHei",
        "display": major or "Microsoft YaHei",
        "latin": major_latin or minor_latin or "Arial",
        "formula": "Cambria Math",
    }
    theme["fonts"]["title"] = list(
        dict.fromkeys(font_candidates(major) + base.get("fonts", {}).get("title", []))
    )
    theme["fonts"]["body"] = list(
        dict.fromkeys(font_candidates(minor) + base.get("fonts", {}).get("body", []))
    )
    theme["fonts"]["display"] = list(
        dict.fromkeys(font_candidates(major) + base.get("fonts", {}).get("display", []))
    )
    theme["fonts"]["latin"] = list(
        dict.fromkeys(
            font_candidates(major_latin or minor_latin)
            + base.get("fonts", {}).get("latin", [])
        )
    )
    theme["brand"]["lockup"] = None
    theme["brand"]["emblem"] = None
    theme["brand"]["fallback_text"] = name or pptx_path.stem
    theme["imported_from"] = {
        "pptx": str(pptx_path),
        "theme_part": source["theme_part"],
        "major_font": major,
        "minor_font": minor,
        "palette": palette,
    }
    return theme


def main() -> None:
    args = parse_args()
    pptx_path = Path(args.pptx).resolve()
    output = Path(args.out).resolve()
    output_path = output if output.suffix.lower() == ".json" else output / "theme.json"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    theme = build_theme(pptx_path, Path(args.base_theme).resolve(), args.name)
    write_json(output_path, theme)
    source_note = [
        f"# Imported Theme: {theme['display_name']}",
        "",
        f"- Source PPTX: `{pptx_path}`",
        f"- Theme part: `{theme['imported_from']['theme_part']}`",
        f"- Major font: `{theme['imported_from']['major_font']}`",
        f"- Minor font: `{theme['imported_from']['minor_font']}`",
        "",
        "This is a generated theme draft. Review the palette, font availability, cover identity, and school-brand rules before using it for a public deck.",
        "",
    ]
    (output_path.parent / "SOURCE.md").write_text("\n".join(source_note), encoding="utf-8")
    summary = {
        "out": str(output_path),
        "name": theme["name"],
        "colors": theme["colors"],
        "pptx_fonts": theme["pptx_fonts"],
        "theme_part": theme["imported_from"]["theme_part"],
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
