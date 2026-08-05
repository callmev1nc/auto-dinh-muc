"""Build script: crop logo.png down to the Ecolar wordmark+bee and de-background it.

- Crops above the white gap that separates the wordmark+bee from the baked tagline.
- Auto-trims the remaining padding (3px margin).
- Removes the opaque white *background* via flood-fill from the borders (preserves
  enclosed white such as the bee's body), so the mark floats cleanly on the frosted header.
- Writes api/_logo_mark.png (reference) and regenerates api/_logo_data.py (base64).

Run:  python build_logo.py
"""
import base64
import io
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent
SRC = ROOT.parent / "logo.png"            # e:/Work/Thai-work/logo.png
OUT_PNG = ROOT / "api" / "_logo_mark.png"
OUT_PY = ROOT / "api" / "_logo_data.py"

MARGIN = 3
WHITE_THRESH = 245        # a row/col with all channels >= this is "empty"
FLOOD_THRESH = 56         # flood-fill tolerance (catches anti-alias fringe, stops at saturated mark)


def main() -> None:
    src = Image.open(SRC).convert("RGB")
    arr = np.array(src)
    nonwhite = (arr < WHITE_THRESH).any(axis=2)

    rows = np.where(nonwhite.any(axis=1))[0]
    # split the top block (wordmark+bee) from the tagline via the largest internal row gap
    prev = int(rows[0])
    best, gap_before = 0, int(rows[-1])
    for r in rows[1:]:
        r = int(r)
        if r - prev > best:
            best, gap_before = r - prev, prev
        prev = r

    top0, top1 = int(rows[0]), gap_before
    sub = nonwhite[top0 : top1 + 1, :]
    cols = np.where(sub.any(axis=0))[0]
    left, right = int(cols[0]), int(cols[-1])

    box = (
        max(left - MARGIN, 0),
        max(top0 - MARGIN, 0),
        min(right + MARGIN + 1, arr.shape[1]),
        min(top1 + MARGIN + 1, arr.shape[0]),
    )
    cropped = src.crop(box)
    w, h = cropped.size

    # de-background: flood-fill a sentinel from every corner, only true background is connected
    rgb = cropped.copy()
    sentinel = (255, 0, 255)
    for corner in [(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)]:
        ImageDraw.floodfill(rgb, corner, sentinel, thresh=FLOOD_THRESH)
    ca = np.array(rgb)
    bg = (ca[:, :, 0] == 255) & (ca[:, :, 1] == 0) & (ca[:, :, 2] == 255)

    rgba = np.array(cropped.convert("RGBA"))
    rgba[bg, 3] = 0  # transparent background
    out = Image.fromarray(rgba)

    buf = io.BytesIO()
    out.save(buf, format="PNG", optimize=True)
    png_bytes = buf.getvalue()
    out.save(OUT_PNG)  # reference asset

    b64 = base64.b64encode(png_bytes).decode("ascii")
    OUT_PY.write_text('LOGO_B64 = "%s"\n' % b64)

    kept_white = int(((rgba[:, :, :3] >= 240).all(axis=2) & (rgba[:, :, 3] > 0)).sum())
    print(f"cropped mark: {w}x{h}px  (was 378x138)")
    print(f"background pixels made transparent: {int(bg.sum())}")
    print(f"enclosed white kept (e.g. bee body): {kept_white}")
    print(f"png bytes: {len(png_bytes)}  base64 chars: {len(b64)}")
    print(f"saved: {OUT_PNG}")
    print(f"saved: {OUT_PY}")


if __name__ == "__main__":
    main()
