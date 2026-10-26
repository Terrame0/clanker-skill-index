#!/usr/bin/env python3
"""Check a deck for the defects a render pass will not reliably show.

    python3 audit.py deck.pptx [--max-notes-missing 0]

Reports shapes that fall outside the slide, text that cannot fit its box,
fonts the deck asks for, and slides without speaker notes.
"""

import argparse
import math
import subprocess
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu, Pt


def font_file(name, bold=False, italic=False):
    style = "Bold" if bold else "Regular"
    if italic:
        style = "Bold Italic" if bold else "Italic"
    try:
        out = subprocess.run(
            ["fc-match", "-f", "%{file}", f"{name}:style={style}"],
            capture_output=True, text=True, timeout=10, check=False,
        ).stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return None
    return out or None


_cache = {}


def measure(text, size_pt, name, bold=False, italic=False):
    from PIL import ImageFont
    key = (name, bold, italic)
    if key not in _cache:
        path = font_file(name, bold, italic)
        _cache[key] = ImageFont.truetype(path, 100) if path else None
    font = _cache[key]
    if font is None:
        return 0.0
    return font.getlength(text) * size_pt / 100


def para_metrics(paragraph):
    text = "".join(run.text for run in paragraph.runs)
    sizes = [run.font.size.pt for run in paragraph.runs if run.font.size]
    size = max(sizes) if sizes else 18.0
    name = next((run.font.name for run in paragraph.runs if run.font.name), "Calibri")
    bold = any(run.font.bold for run in paragraph.runs)
    italic = any(run.font.italic for run in paragraph.runs)
    spacing = paragraph.line_spacing or 1.0
    before = paragraph.space_before.pt if paragraph.space_before else 0.0
    after = paragraph.space_after.pt if paragraph.space_after else 0.0
    return text, size, name, bold, italic, spacing, before + after


def overflow(shape):
    tf = shape.text_frame
    avail_w = shape.width - (tf.margin_left or 0) - (tf.margin_right or 0)
    avail_h = shape.height - (tf.margin_top or 0) - (tf.margin_bottom or 0)
    if avail_w <= 0 or avail_h <= 0:
        return None

    wrap = tf.word_wrap is not False
    total_h = 0.0
    widest = 0.0
    for paragraph in tf.paragraphs:
        text, size, name, bold, italic, spacing, extra = para_metrics(paragraph)
        if not text.strip():
            total_h += size * 1.2 * spacing
            continue
        width = measure(text, size, name, bold, italic)
        widest = max(widest, width)
        per_line = max(1, math.ceil(width / (avail_w / 12700))) if wrap else 1
        total_h += (per_line * size * 1.2 * spacing + extra) * 12700

    if not wrap and widest > avail_w / 12700 + 1:
        return ("horizontal", widest / 72, avail_w / 914400)
    if total_h > avail_h * 1.15:
        return ("vertical", total_h / 914400, avail_h / 914400)
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck", type=Path)
    ap.add_argument("--max-notes-missing", type=int, default=0)
    ap.add_argument("--tolerance", type=float, default=0.02)
    args = ap.parse_args()

    prs = Presentation(args.deck)
    W, H = prs.slide_width, prs.slide_height
    fonts = set()
    problems = []
    missing_notes = []

    for index, slide in enumerate(prs.slides, 1):
        notes = slide.notes_slide.notes_text_frame.text.strip()
        if not notes:
            missing_notes.append(index)

        for shape in slide.shapes:
            if shape.left is None:
                continue
            if (shape.left < 0 or shape.top < 0
                    or shape.left + shape.width > W + Emu(1)
                    or shape.top + shape.height > H + Emu(1)):
                problems.append((index, shape.name, "outside the slide"))

            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.font.name:
                            fonts.add(run.font.name)
                result = overflow(shape)
                if result:
                    kind, need, have = result
                    problems.append((index, shape.name,
                                     f"text overflow ({kind}): needs {need:.2f} in, has {have:.2f} in"))

            if getattr(shape, "has_table", False):
                table = shape.table
                widths = [col.width for col in table.columns]
                for r, row in enumerate(table.rows):
                    for c, cell in enumerate(row.cells):
                        for paragraph in cell.text_frame.paragraphs:
                            if not paragraph.runs:
                                continue
                            text = "".join(run.text for run in paragraph.runs)
                            size = next((run.font.size.pt for run in paragraph.runs
                                         if run.font.size), 12.0)
                            name = next((run.font.name for run in paragraph.runs
                                         if run.font.name), "Calibri")
                            width = measure(text, size, name)
                            available = widths[c] / 12700 - 6
                            if width > available:
                                problems.append((index, f"table r{r}c{c}",
                                                 f"cell text wider than column: "
                                                 f"{width:.2f} in > {available:.2f} in"))

    print(f"slides: {len(prs.slides)}")
    print(f"slides without notes: {len(missing_notes)}"
          + (f" {missing_notes}" if missing_notes else ""))
    print(f"fonts: {', '.join(sorted(fonts)) if fonts else '(none set)'}")

    if problems:
        print(f"\nproblems: {len(problems)}")
        for index, name, message in problems:
            print(f"  slide {index}: {name}: {message}")
    else:
        print("\nno geometry or overflow problems found")

    return 1 if (problems or len(missing_notes) > args.max_notes_missing) else 0


if __name__ == "__main__":
    sys.exit(main())
