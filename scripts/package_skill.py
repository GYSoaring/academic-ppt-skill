from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path


INCLUDE_DIRS = ("scripts", "assets", "references")
INCLUDE_FILES = (
    "SKILL.md",
    "README.md",
    "CHANGELOG.md",
    "LICENSE",
    "requirements.txt",
)
EXCLUDE_PARTS = {
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".ruff_cache",
    "work",
    "outputs",
    "tests",
    "examples",
    ".git",
    ".github",
}
EXCLUDE_SUFFIXES = {".pyc", ".pyo", ".pptx", ".zip"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Package the runtime-only academic-ppt-skill tree as a zip.")
    parser.add_argument("--out", default="dist/academic-ppt-skill.zip")
    parser.add_argument("--version", default="0.1.0")
    parser.add_argument("--root", default=None)
    return parser.parse_args()


def should_include(path: Path) -> bool:
    if any(part in EXCLUDE_PARTS for part in path.parts):
        return False
    if path.suffix.lower() in EXCLUDE_SUFFIXES:
        return False
    return path.is_file()


def collect(root: Path) -> list[Path]:
    files: list[Path] = []
    for name in INCLUDE_FILES:
        candidate = root / name
        if candidate.exists() and should_include(candidate):
            files.append(candidate)
    for directory in INCLUDE_DIRS:
        base = root / directory
        if not base.exists():
            continue
        for candidate in sorted(base.rglob("*")):
            if should_include(candidate):
                files.append(candidate)
    return files


def main() -> None:
    args = parse_args()
    root = Path(args.root).resolve() if args.root else Path(__file__).resolve().parent.parent
    out_path = Path(args.out)
    if not out_path.is_absolute():
        out_path = root / out_path
    out_path.parent.mkdir(parents=True, exist_ok=True)
    files = collect(root)
    if not files:
        raise SystemExit("No runtime files found.")
    prefix = f"academic-ppt-skill-v{args.version}"
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in files:
            archive.write(path, f"{prefix}/{path.relative_to(root).as_posix()}")
    report = {
        "zip": str(out_path),
        "version": args.version,
        "file_count": len(files),
        "size_bytes": out_path.stat().st_size,
        "files": [path.relative_to(root).as_posix() for path in files],
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
