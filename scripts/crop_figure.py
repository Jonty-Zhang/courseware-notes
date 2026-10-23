#!/usr/bin/env python3
"""Crop figure regions from slide pages at print resolution.

A spec file lists one crop per line:
    <name> <page> <x0> <y0> <x1> <y1>
where the box is given as fractions (0-1) of the page width/height, read off the
rendered page or contact sheet. Lines starting with # are ignored.

    crop_figure.py deck.pdf crops.txt --out figs/ [--dpi 220] [--pad 0.01]

Each crop is written to <out>/<name>.jpg (or .png with --format png), and a review sheet of all crops
(<out>/_review.png) is written so the boxes can be checked in one look.
"""
import argparse
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw


def render_page(pdf, page, dpi, tmp):
    stem = Path(tmp) / f"p{page}"
    subprocess.run(["pdftoppm", "-r", str(dpi), "-f", str(page), "-l", str(page), "-png",
                    "-singlefile", str(pdf), str(stem)], check=True)
    return Image.open(f"{stem}.png")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("spec")
    ap.add_argument("--out", required=True)
    ap.add_argument("--dpi", type=int, default=220)
    ap.add_argument("--pad", type=float, default=0.01)
    ap.add_argument("--format", choices=["jpg", "png"], default="jpg",
                    help="jpg keeps the docx small: slide backgrounds are often textured")
    ap.add_argument("--max-width", type=int, default=1600)
    args = ap.parse_args()

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    crops = []
    with tempfile.TemporaryDirectory() as tmp:
        cache = {}
        for line in Path(args.spec).read_text(encoding="utf-8").splitlines():
            line = line.split("#", 1)[0].strip()
            if not line:
                continue
            name, page, *box = line.split()
            page = int(page)
            x0, y0, x1, y1 = (float(v) for v in box)
            if page not in cache:
                cache[page] = render_page(args.pdf, page, args.dpi, tmp)
            img = cache[page]
            w, h = img.size
            p = args.pad
            bx = (max(0, int((x0 - p) * w)), max(0, int((y0 - p) * h)),
                  min(w, int((x1 + p) * w)), min(h, int((y1 + p) * h)))
            crop = img.crop(bx).convert("RGB")
            if crop.width > args.max_width:
                crop = crop.resize((args.max_width, int(crop.height * args.max_width / crop.width)),
                                   Image.LANCZOS)
            if args.format == "jpg":
                crop.save(out / f"{name}.jpg", quality=88, optimize=True)
            else:
                crop.save(out / f"{name}.png")
            crops.append((name, crop))

    if crops:
        thumb_w = 420
        rows = (len(crops) + 3) // 4
        cell_h = 340
        sheet = Image.new("RGB", (thumb_w * 4, cell_h * rows), "white")
        draw = ImageDraw.Draw(sheet)
        for i, (name, c) in enumerate(crops):
            r, col = divmod(i, 4)
            c = c.copy()
            c.thumbnail((thumb_w - 10, cell_h - 30))
            sheet.paste(c, (col * thumb_w + 5, r * cell_h + 25))
            draw.text((col * thumb_w + 5, r * cell_h + 5), name, fill="red")
        sheet.save(out / "_review.png")
    print(f"{len(crops)} crops -> {out}")


if __name__ == "__main__":
    sys.exit(main())
