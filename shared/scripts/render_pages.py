#!/usr/bin/env python3
"""Render a slide PDF for visual reading.

Outputs (under --out):
  pages/p001.png ...   every page at --dpi, for cropping figures and close reading
  sheets/s01.png ...   contact sheets, --grid pages per image with page labels,
                       so a deck can be read in a few image views
  text/p001.txt ...    text layer per page (often garbled for formulas; index only)
  manifest.json        page count, page size, chars of text per page

Requires poppler (pdftoppm, pdftotext, pdfinfo) and Pillow.
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def page_count(pdf):
    out = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, check=True).stdout
    for line in out.splitlines():
        if line.startswith("Pages:"):
            return int(line.split()[1])
    raise RuntimeError("cannot read page count")


def load_font(size):
    for p in ["/System/Library/Fonts/PingFang.ttc", "/System/Library/Fonts/Helvetica.ttc",
              "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]:
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def make_sheets(pages, sheet_dir, grid, sheet_width):
    cols = 2 if grid >= 2 else 1
    rows = (grid + cols - 1) // cols
    cell_w = sheet_width // cols
    font = load_font(28)
    sheets = []
    for start in range(0, len(pages), grid):
        chunk = pages[start:start + grid]
        imgs = [Image.open(p) for p in chunk]
        ratio = imgs[0].height / imgs[0].width
        cell_h = int(cell_w * ratio)
        label_h = 36
        sheet = Image.new("RGB", (cell_w * cols, (cell_h + label_h) * rows), "white")
        draw = ImageDraw.Draw(sheet)
        for i, (img, path) in enumerate(zip(imgs, chunk)):
            r, c = divmod(i, cols)
            x, y = c * cell_w, r * (cell_h + label_h)
            draw.rectangle([x, y, x + cell_w, y + label_h], fill=(30, 30, 30))
            draw.text((x + 10, y + 3), f"p{int(path.stem[1:])}", fill="white", font=font)
            sheet.paste(img.resize((cell_w - 4, cell_h - 4)), (x + 2, y + label_h + 2))
        name = sheet_dir / f"s{start // grid + 1:02d}.png"
        sheet.save(name)
        sheets.append(name.name)
    return sheets


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--out", required=True)
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--grid", type=int, default=4, help="pages per contact sheet (1, 2, 4 or 6)")
    ap.add_argument("--sheet-width", type=int, default=1800)
    args = ap.parse_args()

    pdf = Path(args.pdf).resolve()
    out = Path(args.out)
    page_dir, sheet_dir, text_dir = out / "pages", out / "sheets", out / "text"
    for d in (page_dir, sheet_dir, text_dir):
        d.mkdir(parents=True, exist_ok=True)

    n = page_count(pdf)
    subprocess.run(["pdftoppm", "-r", str(args.dpi), "-png", str(pdf), str(page_dir / "raw")], check=True)
    pages = []
    for f in sorted(page_dir.glob("raw-*.png")):
        num = int(f.stem.split("-")[1])
        dst = page_dir / f"p{num:03d}.png"
        f.rename(dst)
        pages.append(dst)
    pages.sort()

    chars = {}
    for i in range(1, n + 1):
        txt = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), "-layout", str(pdf), "-"],
                             capture_output=True, text=True).stdout
        (text_dir / f"p{i:03d}.txt").write_text(txt, encoding="utf-8")
        chars[i] = len(txt.strip())

    sheets = make_sheets(pages, sheet_dir, args.grid, args.sheet_width)
    w, h = Image.open(pages[0]).size
    manifest = {"pdf": str(pdf), "pages": n, "page_px": [w, h], "dpi": args.dpi,
                "grid": args.grid, "sheets": sheets, "text_chars": chars,
                "low_text_pages": [i for i, c in chars.items() if c < 20]}
    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{pdf.name}: {n} pages -> {len(sheets)} sheets in {sheet_dir}")


if __name__ == "__main__":
    sys.exit(main())
