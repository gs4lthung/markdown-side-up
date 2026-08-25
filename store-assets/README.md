# Chrome Web Store listing assets

All files are **24-bit PNG, colour-type 2 (no alpha channel)** at the exact
dimensions the store form requires.

| File | Size | Store field |
| :--- | :--- | :--- |
| `screenshot-1-render.png` | 1280×800 | Screenshots (min 1, max 5) |
| `screenshot-2-editor.png` | 1280×800 | " |
| `screenshot-3-diagrams.png` | 1280×800 | " |
| `screenshot-4-search.png` | 1280×800 | " |
| `screenshot-5-themes.png` | 1280×800 | " |
| `promo-small-440x280.png` | 440×280 | Small promo tile |
| `promo-marquee-1400x560.png` | 1400×560 | Marquee promo tile |

The screenshots are not mockups — each one is the real extension UI rendered by
the actual `content.js`, captured in headless Chrome, then composed onto a
branded backdrop.

## Regenerating

```bash
pip install pillow
python store-assets/_build/regenerate.py
```

The script copies the extension's runtime files into `_build/app/`, patches two
lines of `content.js` so it renders outside the extension sandbox, screenshots
the result, and rebuilds every asset in place.

- Demo documents and the headline/subhead copy live in `_build/gen_harness.py`
  and `_build/comp.py`.
- If `regenerate.py` aborts with "content.js no longer matches the expected
  patch points", the `isMarkdownPage()` / `getFilename()` signatures changed —
  update the two `.replace()` calls in `_build/regenerate.py`.

`_build/app/` and `_build/raw/` are intermediates and can be deleted at any time.
