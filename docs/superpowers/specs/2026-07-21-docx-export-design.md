# Export to DOCX — Design

**Date:** 2026-07-21
**Status:** Approved (pending implementation plan)
**Component:** Markdown Viewer browser extension (MV3, vanilla JS, no build step)

## Summary

Add a feature that exports the currently viewed Markdown document to a real,
editable Microsoft Word `.docx` file. The file is a genuine Office Open XML
(OOXML) package generated entirely inside the extension — no external libraries,
no build step. It opens cleanly in Word, Google Docs, and LibreOffice with
native headings, bold/italic/strike, tables, lists, blockquotes, links, and
embedded images. Mermaid diagrams and KaTeX math are rasterized to PNG and
embedded so they look exactly like the on-screen page.

## Goals

- Produce a **valid, editable `.docx`** (OOXML), not the legacy `.doc`
  HTML-hack format.
- **High visual fidelity**: diagrams and math embedded as images; syntax-
  highlighted code keeps its colors.
- **Zero external dependencies and no build step**, consistent with how the
  project already *vendors* KaTeX/Mermaid rather than bundling.
- **No silent failures**: any element that cannot be embedded degrades to a
  visible text fallback, and the user is told.

## Non-goals (YAGNI)

- PDF or HTML export.
- Page headers/footers, a Word TOC field, or custom style themes.
- DEFLATE compression of the ZIP (stored/uncompressed is fine for v1; PNG/JPEG
  media is already compressed). May be added later as a pure enhancement via
  `CompressionStream('deflate-raw')`.
- Re-importing `.docx` back into the editor.
- A multi-format export menu — the button is a single DOCX action.

## Decisions (from brainstorming)

1. **Real, editable `.docx`** — hand-rolled OOXML, not the `.doc` HTML hack.
2. **Generation mechanism** — hand-rolled OOXML string builders packed with a
   small *stored* (uncompressed) ZIP writer + CRC32. No `docx`/JSZip library.
3. **Rich content** — Mermaid diagrams and math **embedded as images**; code
   blocks kept as monospace text on a light-gray shaded background with syntax
   colors preserved.

## Architecture

### Module boundary

A new self-contained file **`docx-export.js`** is added to the manifest's
`content_scripts[0].js` array **before** `content.js`. Files in the same
content-script entry share one isolated-world global, so the module exposes a
single entry point:

```js
window.MdDocx = { export({ container, filename, title }) }  // async → triggers download
```

`content.js` remains the orchestrator: it owns the toolbar button and the
keyboard shortcut, and calls `window.MdDocx.export(...)`. This keeps the
already-large `content.js` (~2047 lines) from growing by ~600 lines and gives
the exporter a clean, independently understandable boundary.

`web_accessible_resources` already lists `katex/fonts/*`, which the exporter
needs for font inlining (see Math). No manifest permission changes are required
beyond adding `docx-export.js` to the content-script list.

### Internal units of `docx-export.js`

Each unit has one job and a well-defined interface:

1. **Stored-ZIP writer** (inlined) — `crc32(bytes)` and
   `zipStore(files[]) → Uint8Array`. Writes local file headers, the central
   directory, and the end-of-central-directory record for *stored* entries only.
   ~50 lines, no dependencies. Asserts each entry's CRC32 and sizes so a bad
   part throws before a corrupt file is emitted.
2. **OOXML part builders** — pure string builders returning XML text:
   `contentTypesXml`, `relsXml`, `documentXml(bodyXml)`, `stylesXml`,
   `numberingXml`, `docPropsXml`, `documentRelsXml(relationships)`.
3. **DOM walker** — `blocksToOoxml(container)` for block-level elements and
   `inlineToRuns(node)` for inline runs. The translation core.
4. **Media collector** — accumulates
   `{ id, path, bytes, contentType, wEmu, hEmu }` for every embedded raster and
   hands relationship IDs back to the walker so drawings reference the right
   media part.

### DOCX package layout produced

```
[Content_Types].xml
_rels/.rels
docProps/core.xml
word/document.xml
word/_rels/document.xml.rels
word/styles.xml
word/numbering.xml
word/media/img1.png, img2.png, …
```

## Source of truth: the live rendered DOM

The exporter walks the **live `#md-content .md-body`** node. All the hard
rendering — tables, nested lists, syntax-highlighted code, KaTeX HTML, Mermaid
SVG — is already computed there, so the renderer is the single source of truth
and no Markdown is re-parsed for export.

- If the editor is in `edit`/`split` mode, `content.js` **flushes the pending
  preview render** before calling the exporter, so the exported DOM matches the
  current editor text (unsaved edits are included).
- Export always emits a **light-themed** document regardless of the on-screen
  dark/light theme (the Word norm).

## Block-level mapping (`.md-body` children → WordprocessingML)

| DOM | OOXML |
|---|---|
| `h1`–`h6` | `w:p` with `Heading1`–`Heading6` paragraph style |
| `p` | normal `w:p`, inline runs |
| `ul` / `ol` (including nested) | list paragraphs referencing `numbering.xml`; `w:ilvl` per nesting depth; bullets vs decimal; each top-level `ol` restarts numbering |
| task-list `li` | checkbox glyph `☐`/`☑` prefix run + item text |
| `blockquote` | `Quote` style — left indent + left border, italic |
| `pre > code` (code block) | one `w:p` per line, `Consolas`, light-gray paragraph shading; syntax colors from child `<span>` classes mapped to colored `w:r` runs |
| `table` | `w:tbl` with borders; header row bold and shaded |
| `hr` | empty `w:p` with a bottom border |
| `img`, `.mermaid svg`, `.math-inline` / `.math-block` | image drawing (see Rich content) |

## Inline mapping (`inlineToRuns`)

Walks text and inline nodes into `w:r` runs, accumulating run properties so
nested formatting composes (e.g. bold **and** italic):

- `strong` / `b` → bold
- `em` / `i` → italic
- `del` / `s` → strike
- `code` → `Consolas` run with light-gray shading
- `a[href]` → `w:hyperlink` with an external relationship
- `br` → `w:br`
- inline `img` / inline math → inline drawing

All text is XML-escaped.

## Rich content → embedded images

A single helper `rasterize(node) → { pngBytes, wPx, hPx }` draws the node to an
offscreen `<canvas>` at **2× scale** for crispness, then
`canvas.toBlob('image/png')` → `arrayBuffer`. Each raster is registered with the
media collector and referenced by a `w:drawing` sized in EMUs (`px × 9525`),
width-capped to the ~6.0 in content area with aspect ratio preserved.

- **Mermaid** (`.mermaid > svg`): serialize the SVG (stamp explicit
  `width`/`height` from `getBoundingClientRect`, add a white background rect),
  wrap as a `Blob`/`data:image/svg+xml`, draw to canvas. The SVG is same-origin
  (generated inline), so the canvas is not tainted and PNG export succeeds.
- **KaTeX math** (`.math-inline` / `.math-block` → `.katex`): wrap the rendered
  `.katex` HTML in `<svg><foreignObject>`, with the **KaTeX fonts inlined as
  `@font-face` data-URIs** (fetched once via `chrome.runtime.getURL(
  'katex/fonts/…')`, base64-encoded, and cached). Inlining fonts avoids external
  references and keeps the canvas clean. Display math becomes a block image;
  inline math becomes an inline image aligned to the text baseline.
- **Raster images** (`img`): reuse the existing `fetch(src) → blob →
  arrayBuffer` pattern already present in `content.js`. Detect true pixel size
  via `createImageBitmap`. PNG/JPEG bytes embed directly (no re-encode);
  GIF/WebP/SVG/BMP are normalized to PNG via canvas. `[Content_Types].xml`
  declares every media extension used.

## Fallback ladder (no silent failures)

Every raster is wrapped in try/catch and degrades *visibly*:

- **Math fails** → emit the original TeX from
  `.katex-mathml annotation[encoding="application/x-tex"]` (always present in
  KaTeX output) as a `Consolas` run.
- **Diagram fails** → emit the Mermaid source in a code block.
- **Remote image fails** (CORS/offline) → emit an italic placeholder
  `[image: <alt> — could not embed]`.
- The export still completes. A summary toast reports any skips, e.g.
  *"Exported — 1 image couldn't be embedded."*

## UI & trigger

- One **Export** button added to the top bar `#md-bar .bar-right` (visible in
  all modes, unlike the editor-only format bar), placed left of Search. A
  Word-style document icon with the label **DOCX**.
- Keyboard shortcut **Ctrl+Shift+E**.

Async flow: on click → disable the button and swap to a spinner →
`await window.MdDocx.export(...)` (rasterization is async and can take a moment
on diagram-heavy documents) → download via an anchor
(`application/vnd.openxmlformats-officedocument.wordprocessingml.document`,
filename `<basename>.docx`) → restore the button → success toast. The filename
derives from `getFilename()` with the extension swapped to `.docx`.

## Error handling

- Whole-export failure (unexpected throw) is caught at the top: the button is
  restored, a red toast shows *"Export failed — see console,"* and the full
  error is logged. The button never stays stuck spinning.
- An empty document still produces a valid minimal `.docx`.
- The stored-ZIP writer validates each entry's CRC32 and sizes; a bad part
  throws before download rather than emitting a corrupt file.

## Testing

Manual verification fits this context (browser extension, binary output, no
existing automated test harness):

1. **Round-trip fixtures** — export `ALL-FEATURES-TEST.md`, `math-test.md`,
   `emoji-test.md`, and a Mermaid-heavy document; open each in **Word, Google
   Docs, and LibreOffice**; confirm no repair prompt and that
   headings/tables/lists/code/images/math/diagrams render correctly.
2. **Validity** — the OOXML opens without Word's "unreadable content" repair
   dialog (the strictest real-world check).
3. **Fallbacks** — a document with a broken remote image and a deliberately
   malformed `$…$` confirms graceful degradation and the summary toast.
4. **Fixture** — a short `docx-export-test.md` checklist enumerating every
   supported element, committed alongside for repeatable manual runs.

## Files touched

- **New:** `docx-export.js` (the exporter module).
- **New:** `docx-export-test.md` (manual test fixture).
- **Modified:** `manifest.json` — add `docx-export.js` to
  `content_scripts[0].js` before `content.js`.
- **Modified:** `content.js` — add the Export button to `#md-bar .bar-right`,
  wire the click handler and Ctrl+Shift+E, flush the preview before export, and
  call `window.MdDocx.export(...)`.
- **Modified:** `content.css` — styles for the Export button and spinner state.
