# -*- coding: utf-8 -*-
"""Render the composed listing assets to exact-size, alpha-free PNGs (and JPEGs)."""
import pathlib, sys
from PIL import Image
from shoot import shoot, BASE, OUT

JOBS = [
    ("shot-1.html",  "screenshot-1-render.png",   1280, 800),
    ("shot-2.html",  "screenshot-2-editor.png",   1280, 800),
    ("shot-3.html",  "screenshot-3-diagrams.png", 1280, 800),
    ("shot-4.html",  "screenshot-4-search.png",   1280, 800),
    ("shot-5.html",  "screenshot-5-themes.png",   1280, 800),
    ("marquee.html", "promo-marquee-1400x560.png", 1400, 560),
    ("tile.html",    "promo-small-440x280.png",    440, 280),
]

if __name__ == "__main__":
    only = sys.argv[1:] or None
    for src, dst, w, h in JOBS:
        if only and not any(o in src for o in only):
            continue
        shoot(BASE / src, OUT / dst, w, h, budget=8000)

    print("\nverification:")
    for _, dst, w, h in JOBS:
        p = OUT / dst
        if not p.exists():
            continue
        im = Image.open(p)
        # Chrome Web Store: 24-bit PNG, no alpha channel.
        if im.mode != "RGB":
            im = im.convert("RGB")
            im.save(p, "PNG", optimize=True)
            im = Image.open(p)
        ok = "OK " if (im.size == (w, h) and im.mode == "RGB") else "BAD"
        print(f"  {ok} {p.name:<32} {im.size[0]}x{im.size[1]}  {im.mode}  "
              f"{p.stat().st_size//1024} KB")
