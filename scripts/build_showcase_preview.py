"""Compose real Office renders into a GitHub-friendly 16:9 preview.

This script never invents product UI. It only arranges PNG pages/slides that
were rendered from the checked-in showcase artifacts.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageOps

CANVAS = (1600, 900)
BACKGROUND = (8, 19, 36)
CARD = (24, 42, 66)
MARGIN = 42
GAP = 24


def compose_preview(inputs: list[Path], output: Path, columns: int | None = None) -> None:
    if not inputs:
        raise ValueError("at least one input image is required")
    for item in inputs:
        if not item.is_file():
            raise FileNotFoundError(item)

    count = len(inputs)
    columns = columns or (1 if count == 1 else 2 if count <= 4 else 3)
    columns = max(1, min(columns, count))
    rows = (count + columns - 1) // columns
    cell_width = (CANVAS[0] - 2 * MARGIN - (columns - 1) * GAP) // columns
    cell_height = (CANVAS[1] - 2 * MARGIN - (rows - 1) * GAP) // rows

    canvas = Image.new("RGB", CANVAS, BACKGROUND)
    for index, source in enumerate(inputs):
        row, column = divmod(index, columns)
        left = MARGIN + column * (cell_width + GAP)
        top = MARGIN + row * (cell_height + GAP)
        card = Image.new("RGB", (cell_width, cell_height), CARD)
        with Image.open(source) as image:
            rendered = ImageOps.contain(
                image.convert("RGB"),
                (cell_width - 20, cell_height - 20),
                Image.Resampling.LANCZOS,
            )
        x = (cell_width - rendered.width) // 2
        y = (cell_height - rendered.height) // 2
        card.paste(rendered, (x, y))
        canvas.paste(card, (left, top))

    output.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output, "PNG", optimize=True)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", action="append", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--columns", type=int)
    args = parser.parse_args()
    compose_preview(args.input, args.output, args.columns)


if __name__ == "__main__":
    main()
