# -*- coding: utf-8 -*-
"""
Regenerate every Chrome Web Store listing image from the live extension source.

    python store-assets/_build/regenerate.py

Requires: Pillow (`pip install pillow`) and Google Chrome installed at the usual
Windows location (edit CHROME in shoot.py otherwise).

Pipeline
  1. copy the extension's runtime files into _build/app/
  2. patch content.js so it renders outside the extension sandbox
  3. drive the real UI in headless Chrome -> _build/raw/*.png
  4. compose headline + browser frame around each capture
  5. render the composites to exact-size, alpha-free 24-bit PNGs in store-assets/
"""
import pathlib, shutil, subprocess, sys

BUILD = pathlib.Path(__file__).resolve().parent
REPO = BUILD.parent.parent
APP = BUILD / "app"

RUNTIME_FILES = ["content.css", "bridge.js", "mermaid.min.js", "docx-export.js",
                 "popup.css", "popup.js", "popup.html"]
RUNTIME_DIRS = ["katex", "icons"]


def prepare():
    APP.mkdir(parents=True, exist_ok=True)
    for name in RUNTIME_FILES:
        shutil.copy2(REPO / name, APP / name)
    for name in RUNTIME_DIRS:
        dst = APP / name
        if dst.exists():
            shutil.rmtree(dst)
        shutil.copytree(REPO / name, dst)

    # content.js assumes it was injected by the extension into a plain-text .md
    # page. Two surgical edits let it run inside the harness instead.
    src = (REPO / "content.js").read_text(encoding="utf-8")
    patched = src.replace(
        "function isMarkdownPage() {",
        "function isMarkdownPage() {\n    if (window.__MD_HARNESS__) return true;", 1)
    patched = patched.replace(
        "function getFilename() {",
        "function getFilename() {\n    if (window.__MD_FILENAME__) return window.__MD_FILENAME__;", 1)
    if patched == src:
        sys.exit("content.js no longer matches the expected patch points — "
                 "update regenerate.py before rebuilding.")
    (APP / "content.mv.js").write_text(patched, encoding="utf-8")
    print("prepared _build/app/ from extension source")


def run(script):
    print(f"\n--- {script} ---")
    r = subprocess.run([sys.executable, str(BUILD / script)], cwd=BUILD)
    if r.returncode:
        sys.exit(r.returncode)


if __name__ == "__main__":
    prepare()
    run("gen_harness.py")
    run("shoot.py")
    # extra capture: the narrower window used as the marquee artwork
    sys.path.insert(0, str(BUILD))
    from shoot import shoot, RAW
    shoot(APP / "app-view-dark.html", RAW / "marquee-app.png", 680, 540)
    run("comp.py")
    run("build.py")
