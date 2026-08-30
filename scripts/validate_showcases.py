"""Validate checked-in showcase reproducibility, assets and public hygiene."""

from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
SHOWCASE = ROOT / "showcase"
EXPECTED_DEMOS = {
    "deck-pipeline": {"deck.compile", "document.inspect", "slide.render"},
    "template-gallery": {"deck.compile", "document.inspect", "slide.render"},
    "image-rich-deck": {"deck.compile", "document.inspect", "slide.render"},
    "production-dashboard": {"office-generate-production-dashboard", "batch.convert"},
    "word-executive-brief": {"document.inspect", "batch.convert"},
    "global-text-replace": {"batch.replace_text", "slide.render"},
    "batch-to-pdf": {"batch.convert"},
}
ABSOLUTE_PATH = re.compile(r"(?:[A-Za-z]:[\\/]|/home/|/Users/|\\\\[^\\]+\\)")


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def validate_demo(name: str, expected_capabilities: set[str]) -> list[str]:
    errors: list[str] = []
    demo = SHOWCASE / name
    metadata_path = demo / "metadata.json"
    if not metadata_path.is_file():
        return [f"{name}: missing metadata.json"]
    raw = metadata_path.read_text(encoding="utf-8")
    if ABSOLUTE_PATH.search(raw):
        errors.append(f"{name}: metadata contains an absolute/local path")
    try:
        metadata = json.loads(raw)
    except json.JSONDecodeError as error:
        return [f"{name}: invalid metadata JSON: {error}"]

    if metadata.get("schema_version") != "dcc-mcp-showcase/1.0":
        errors.append(f"{name}: unsupported schema_version")
    if set(metadata.get("capabilities", [])) != expected_capabilities:
        errors.append(f"{name}: capability set does not match the gallery contract")
    if not metadata.get("verification"):
        errors.append(f"{name}: verification list is empty")
    if not metadata.get("reproduce"):
        errors.append(f"{name}: reproduce list is empty")

    referenced = [
        *metadata.get("inputs", []),
        *metadata.get("artifacts", []),
        *metadata.get("previews", []),
        "transcript.json",
    ]
    if len(referenced) != len(set(referenced)):
        errors.append(f"{name}: duplicate referenced file")
    hashes = metadata.get("sha256", {})
    for relative in referenced:
        path = demo / relative
        if not path.is_file():
            errors.append(f"{name}: missing referenced file {relative}")
            continue
        if path.stat().st_size > 3 * 1024 * 1024:
            errors.append(f"{name}: {relative} exceeds 3 MB")
        expected = hashes.get(relative)
        if expected != digest(path):
            errors.append(f"{name}: checksum mismatch for {relative}")

    previews = metadata.get("previews", [])
    if previews != ["preview.png"]:
        errors.append(f"{name}: exactly one preview.png is required")
    elif (demo / "preview.png").is_file():
        with Image.open(demo / "preview.png") as image:
            if image.size != (1600, 900):
                errors.append(f"{name}: preview must be 1600x900, got {image.size}")
            if image.format != "PNG":
                errors.append(f"{name}: preview must be PNG")
    return errors


def main() -> int:
    errors: list[str] = []
    if not (ROOT / "docs" / "images" / "office-suite-showcase.webp").is_file():
        errors.append("missing docs/images/office-suite-showcase.webp")
    elif (ROOT / "docs" / "images" / "office-suite-showcase.webp").stat().st_size > 500 * 1024:
        errors.append("office-suite-showcase.webp exceeds 500 KB")

    actual = {path.name for path in SHOWCASE.iterdir() if path.is_dir()}
    if actual != set(EXPECTED_DEMOS):
        errors.append(f"showcase directory set mismatch: expected {sorted(EXPECTED_DEMOS)}, got {sorted(actual)}")
    for name, capabilities in EXPECTED_DEMOS.items():
        errors.extend(validate_demo(name, capabilities))

    readme = (SHOWCASE / "README.md").read_text(encoding="utf-8")
    for name in EXPECTED_DEMOS:
        if f"./{name}/preview.png" not in readme or f"./{name}/metadata.json" not in readme:
            errors.append(f"showcase/README.md does not link preview and metadata for {name}")

    if errors:
        print("Showcase validation failed:", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1
    print(f"Validated {len(EXPECTED_DEMOS)} reproducible showcases and the suite hero image.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
