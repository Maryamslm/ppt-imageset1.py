"""
resize_images_grid.py
─────────────────────
Resizes and repositions all pictures on a chosen slide into a uniform grid
that fills the slide area, with optional padding between cells.

Requirements:
    pip install python-pptx

Usage:
    python resize_images_grid.py presentation.pptx
    python resize_images_grid.py presentation.pptx --slide 1
    python resize_images_grid.py presentation.pptx --slide 1 --cols 8 --padding 5
    python resize_images_grid.py presentation.pptx --slide 1 --margin 20 --padding 8
"""

import argparse
import math
import copy
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Emu, Pt
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    raise SystemExit("❌  python-pptx not found. Run:  pip install python-pptx")


# ─── helpers ──────────────────────────────────────────────────────────────────

def pt_to_emu(pt: float) -> int:
    """Points → EMU  (1 pt = 12700 EMU)"""
    return int(pt * 12700)


def px_to_emu(px: float, dpi: int = 96) -> int:
    """Pixels → EMU  (1 inch = 914400 EMU)"""
    return int(px / dpi * 914400)


def is_picture(shape) -> bool:
    """Return True for picture shapes (handles group members too)."""
    return shape.shape_type == MSO_SHAPE_TYPE.PICTURE


def collect_pictures(slide):
    """Collect all top-level picture shapes from a slide."""
    pics = []
    for shape in slide.shapes:
        if is_picture(shape):
            pics.append(shape)
    return pics


def best_grid(n: int, slide_w: int, slide_h: int) -> tuple[int, int]:
    """
    Choose (cols, rows) so the grid fills the slide with as square cells as
    possible.  Tries every column count 1..n and picks the one that minimises
    |cell_w/cell_h - 1|  (i.e. cells closest to square).
    """
    best = (n, 1)
    best_score = float("inf")
    for cols in range(1, n + 1):
        rows = math.ceil(n / cols)
        cell_w = slide_w / cols
        cell_h = slide_h / rows
        ratio = cell_w / cell_h
        score = abs(ratio - 1.0)          # 0 = perfect square
        if score < best_score:
            best_score = score
            best = (cols, rows)
    return best


# ─── core ─────────────────────────────────────────────────────────────────────

def arrange_grid(
    pptx_path: str,
    slide_index: int = 0,       # 0-based
    cols: int | None = None,    # None → auto
    margin_pt: float = 0.0,     # outer margin in points
    padding_pt: float = 4.0,    # gap between cells in points
    output_path: str | None = None,
):
    prs = Presentation(pptx_path)

    if slide_index >= len(prs.slides):
        raise ValueError(
            f"Slide {slide_index + 1} does not exist "
            f"(presentation has {len(prs.slides)} slide(s))."
        )

    slide = prs.slides[slide_index]
    slide_w = prs.slide_width    # EMU
    slide_h = prs.slide_height   # EMU

    pics = collect_pictures(slide)
    if not pics:
        print("⚠️  No pictures found on that slide.")
        return

    n = len(pics)
    print(f"✅  Found {n} picture(s) on slide {slide_index + 1}.")

    # Convert margin / padding from points to EMU
    margin  = pt_to_emu(margin_pt)
    padding = pt_to_emu(padding_pt)

    # Usable area
    usable_w = slide_w - 2 * margin
    usable_h = slide_h - 2 * margin

    # Determine grid dimensions
    if cols is None:
        auto_cols, auto_rows = best_grid(n, usable_w, usable_h)
        cols = auto_cols
    rows = math.ceil(n / cols)

    print(f"📐  Grid: {cols} col(s) × {rows} row(s)  (padding={padding_pt}pt, margin={margin_pt}pt)")

    # Cell size (each image gets this bounding box)
    cell_w = (usable_w - padding * (cols - 1)) // cols
    cell_h = (usable_h - padding * (rows - 1)) // rows

    print(f"🖼️   Cell size: {cell_w/914400*2.54:.1f} cm × {cell_h/914400*2.54:.1f} cm  "
          f"({cell_w} × {cell_h} EMU)")

    # Reposition and resize each picture
    for i, pic in enumerate(pics):
        row = i // cols
        col = i % cols

        left = margin + col * (cell_w + padding)
        top  = margin + row * (cell_h + padding)

        pic.left   = int(left)
        pic.top    = int(top)
        pic.width  = int(cell_w)
        pic.height = int(cell_h)

    # Save
    if output_path is None:
        p = Path(pptx_path)
        output_path = str(p.with_stem(p.stem + "_grid"))

    prs.save(output_path)
    print(f"\n💾  Saved → {output_path}")


# ─── CLI ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Arrange all images on a PowerPoint slide into a uniform grid."
    )
    parser.add_argument("pptx", help="Path to the .pptx file")
    parser.add_argument(
        "--slide", type=int, default=1,
        help="Slide number to process (1-based, default: 1)"
    )
    parser.add_argument(
        "--cols", type=int, default=None,
        help="Number of columns (default: auto-calculated for squarest cells)"
    )
    parser.add_argument(
        "--padding", type=float, default=4.0,
        help="Gap between images in points (default: 4)"
    )
    parser.add_argument(
        "--margin", type=float, default=0.0,
        help="Outer margin in points (default: 0)"
    )
    parser.add_argument(
        "--output", type=str, default=None,
        help="Output .pptx path (default: <original>_grid.pptx)"
    )

    args = parser.parse_args()

    arrange_grid(
        pptx_path   = args.pptx,
        slide_index = args.slide - 1,   # convert to 0-based
        cols        = args.cols,
        margin_pt   = args.margin,
        padding_pt  = args.padding,
        output_path = args.output,
    )


if __name__ == "__main__":
    main()
