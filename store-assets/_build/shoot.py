# -*- coding: utf-8 -*-
"""Headless-Chrome screenshot driver. Exact pixel sizes, DPR 1, no alpha."""
import subprocess, sys, pathlib, shutil

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
BASE = pathlib.Path(__file__).resolve().parent
RAW = BASE / "raw"
OUT = BASE.parent          # finished assets land in store-assets/
RAW.mkdir(exist_ok=True)
OUT.mkdir(exist_ok=True)


def shoot(page: pathlib.Path, dest: pathlib.Path, w: int, h: int, budget: int = 12000):
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        dest.unlink()
    cmd = [
        CHROME,
        "--headless",
        "--disable-gpu",
        "--hide-scrollbars",
        "--force-device-scale-factor=1",
        "--disable-lcd-text",
        f"--window-size={w},{h}",
        f"--virtual-time-budget={budget}",
        "--run-all-compositor-stages-before-draw",
        "--user-data-dir=" + str(BASE / "cdp"),
        "--screenshot=" + str(dest),
        page.as_uri(),
    ]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if not dest.exists():
        print("FAILED", page.name, r.stdout[-800:], r.stderr[-800:])
        sys.exit(1)
    print(f"  {dest.name}  {w}x{h}")


if __name__ == "__main__":
    app = BASE / "app"
    W, H = 1152, 640
    jobs = [
        ("app-view-dark.html", "view-dark.png", W, H),
        ("app-split-light.html", "split-light.png", W, H),
        ("app-diagram-dark.html", "diagram-dark.png", W, H),
        ("app-search-light.html", "search-light.png", W, H),
        ("app-theme-dark.html", "theme-dark.png", W, H),
        ("popup-dark.html", "popup-dark.png", 280, 760),
    ]
    only = sys.argv[1:] or None
    for src, dst, w, h in jobs:
        if only and not any(o in src for o in only):
            continue
        shoot(app / src, RAW / dst, w, h)

    # Trim the popup capture down to its real content height.
    from PIL import Image
    p = RAW / "popup-dark.png"
    if p.exists():
        im = Image.open(p).convert("RGB")
        bg = im.getpixel((im.width - 2, im.height - 2))
        last = 0
        px = im.load()
        for y in range(im.height):
            row_has_content = any(px[x, y] != bg for x in range(0, im.width, 2))
            if row_has_content:
                last = y
        im.crop((0, 0, im.width, min(im.height, last + 15))).save(p)
        print(f"  popup trimmed to {Image.open(p).size}")
