#!/usr/bin/env python3
"""Combine several cropped figures into one labelled grid (a comparison figure).

    compose_figures.py out.png --cols 3 a.png:光滑面 b.png:圆柱铰链 ...

Each tile is scaled to the same height and centred in its cell, with its label
underneath. Use it when several small diagrams are best read side by side, such
as the force diagrams of different constraint types.
"""
import argparse
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def load_font(size):
    for p in ["/System/Library/Fonts/STHeiti Medium.ttc", "/System/Library/Fonts/PingFang.ttc",
              "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"]:
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("out")
    ap.add_argument("items", nargs="+", help="path:label")
    ap.add_argument("--cols", type=int, default=3)
    ap.add_argument("--tile-h", type=int, default=360)
    ap.add_argument("--cell-w", type=int, default=560)
    args = ap.parse_args()

    tiles = []
    for item in args.items:
        path, _, label = item.partition(":")
        img = Image.open(path).convert("RGB")
        scale = min(args.tile_h / img.height, (args.cell_w - 20) / img.width)
        tiles.append((img.resize((int(img.width * scale), int(img.height * scale))), label))

    font = load_font(34)
    label_h = 56
    rows = (len(tiles) + args.cols - 1) // args.cols
    cell_h = args.tile_h + label_h + 20
    sheet = Image.new("RGB", (args.cell_w * args.cols, cell_h * rows), "white")
    draw = ImageDraw.Draw(sheet)
    for i, (img, label) in enumerate(tiles):
        r, c = divmod(i, args.cols)
        x0, y0 = c * args.cell_w, r * cell_h
        sheet.paste(img, (x0 + (args.cell_w - img.width) // 2, y0 + 10 + (args.tile_h - img.height) // 2))
        tw = draw.textlength(label, font=font)
        draw.text((x0 + (args.cell_w - tw) / 2, y0 + args.tile_h + 22), label, fill="black", font=font)
    sheet.save(args.out)
    print(f"{len(tiles)} tiles -> {args.out}")


if __name__ == "__main__":
    sys.exit(main())
