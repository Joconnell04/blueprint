#!/usr/bin/env python3
"""Render the Blueprint extension icon: assets/icon.png, 512x512, PNG with transparency.

A dark rounded square with a 3x2 grid of thin rounded rectangles in the accent. No text,
because the icon is shown at 32 px in Claude Desktop's Settings.

    python3 assets/make_icon.py [--out assets/icon.png] [--size 512]

Uses Pillow when it is installed, otherwise its own rasteriser and PNG writer, so the icon
can be rebuilt on a machine with nothing but the standard library.
"""

from __future__ import annotations

import argparse
import os
import struct
import zlib

BG = (0x17, 0x17, 0x1A)       # tokens.css --bg0
ACCENT = (0x6D, 0x8B, 0xFF)   # tokens.css --accent

# Geometry at 512 px, scaled for any other size.
CORNER = 112.0                # background corner radius
GRID = (72.0, 150.0, 440.0, 362.0)   # x0, y0, x1, y1 of the block grid
GUTTER = 24.0
STROKE = 9.0
BLOCK_CORNER = 14.0


def blocks(size: int) -> list[tuple[float, float, float, float]]:
    """The 3x2 grid, in pixels at `size`."""
    k = size / 512.0
    x0, y0, x1, y1 = (v * k for v in GRID)
    gut = GUTTER * k
    w = (x1 - x0 - 2 * gut) / 3.0
    h = (y1 - y0 - gut) / 2.0
    out = []
    for row in range(2):
        for col in range(3):
            bx = x0 + col * (w + gut)
            by = y0 + row * (h + gut)
            out.append((bx, by, bx + w, by + h))
    return out


# ---------------------------------------------------------------- Pillow path

def render_pillow(size: int, path: str) -> None:
    from PIL import Image, ImageDraw

    ss = 4  # supersample, then downscale for clean edges
    big = size * ss
    img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    k = size / 512.0
    d.rounded_rectangle([0, 0, big - 1, big - 1], radius=CORNER * k * ss, fill=BG + (255,))
    for (x0, y0, x1, y1) in blocks(size):
        d.rounded_rectangle(
            [x0 * ss, y0 * ss, x1 * ss, y1 * ss],
            radius=BLOCK_CORNER * k * ss,
            outline=ACCENT + (255,),
            width=max(1, round(STROKE * k * ss)),
        )
    img.resize((size, size), Image.LANCZOS).save(path, "PNG")


# ------------------------------------------------- standard-library fallback

def in_rounded(x: float, y: float, x0: float, y0: float, x1: float, y1: float, r: float) -> bool:
    if x < x0 or x > x1 or y < y0 or y > y1:
        return False
    cx = min(max(x, x0 + r), x1 - r)
    cy = min(max(y, y0 + r), y1 - r)
    dx, dy = x - cx, y - cy
    return dx * dx + dy * dy <= r * r


def render_plain(size: int, path: str) -> None:
    k = size / 512.0
    corner = CORNER * k
    stroke = STROKE * k
    brad = BLOCK_CORNER * k
    rects = blocks(size)
    ss = 4                     # sub-samples per axis
    off = [(i + 0.5) / ss for i in range(ss)]
    n = float(ss * ss)

    # Start from the flat background colour with zero alpha, then fill in what is covered.
    rows = [bytearray(size * 4) for _ in range(size)]
    for y in range(size):
        row = rows[y]
        # Background alpha: only the four corners need sampling; everything else is inside or out.
        near_corner = y < corner or y >= size - corner
        for x in range(size):
            i = x * 4
            if near_corner and (x < corner or x >= size - corner):
                hits = 0
                for dy in off:
                    for dx in off:
                        if in_rounded(x + dx, y + dy, 0.0, 0.0, float(size), float(size), corner):
                            hits += 1
                a = int(round(255 * hits / n))
            else:
                a = 255
            if a:
                row[i], row[i + 1], row[i + 2], row[i + 3] = BG[0], BG[1], BG[2], a

    # The block outlines: sample only inside each rect's bounding box.
    for (x0, y0, x1, y1) in rects:
        ix0, iy0 = max(0, int(x0) - 1), max(0, int(y0) - 1)
        ix1, iy1 = min(size, int(x1) + 2), min(size, int(y1) + 2)
        ix0i, iy0i = x0 + stroke, y0 + stroke
        ix1i, iy1i = x1 - stroke, y1 - stroke
        irad = max(0.0, brad - stroke)
        for y in range(iy0, iy1):
            row = rows[y]
            for x in range(ix0, ix1):
                hits = 0
                for dy in off:
                    for dx in off:
                        px, py = x + dx, y + dy
                        if in_rounded(px, py, x0, y0, x1, y1, brad) and not in_rounded(px, py, ix0i, iy0i, ix1i, iy1i, irad):
                            hits += 1
                if not hits:
                    continue
                f = hits / n
                i = x * 4
                for c in range(3):
                    row[i + c] = int(round(row[i + c] * (1 - f) + ACCENT[c] * f))
                row[i + 3] = max(row[i + 3], int(round(255 * f)))

    raw = bytearray()
    for row in rows:
        raw.append(0)          # filter type 0
        raw += row
    write_png(path, size, size, bytes(raw))


def write_png(path: str, w: int, h: int, raw: bytes) -> None:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0)   # 8-bit RGBA
    with open(path, "wb") as f:
        f.write(b"\x89PNG\r\n\x1a\n")
        f.write(chunk(b"IHDR", ihdr))
        f.write(chunk(b"IDAT", zlib.compress(raw, 9)))
        f.write(chunk(b"IEND", b""))


def main() -> None:
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(here, "icon.png"))
    ap.add_argument("--size", type=int, default=512)
    args = ap.parse_args()
    try:
        render_pillow(args.size, args.out)
        how = "pillow"
    except ImportError:
        render_plain(args.size, args.out)
        how = "stdlib"
    print(f"wrote {args.out} ({args.size}x{args.size}, {how}, {os.path.getsize(args.out)} bytes)")


if __name__ == "__main__":
    main()
