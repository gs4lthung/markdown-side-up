"use strict";
// docx-export.js — Export the rendered Markdown document to a real .docx (OOXML).
// Loaded as a content script before content.js; also require()-able in Node for
// unit-testing the pure core. No DOM/chrome access at module top level.

// ── CRC32 (IEEE 802.3) ───────────────────────────────────────────────────────
const _CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++)
    c = _CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// ── Stored (uncompressed) ZIP writer ─────────────────────────────────────────
function zipStore(files) {
  const enc = new TextEncoder();
  const locals = [];
  const centrals = [];
  let offset = 0;

  for (const f of files) {
    if (!(f.bytes instanceof Uint8Array)) throw new Error('zipStore: ' + f.name + ' bytes must be Uint8Array');
    if (!f.name) throw new Error('zipStore: every entry needs a name');
    const nameBytes = enc.encode(f.name);
    const data = f.bytes;
    const crc = crc32(data);
    const size = data.length;

    const local = new Uint8Array(30 + nameBytes.length + size);
    const ldv = new DataView(local.buffer);
    ldv.setUint32(0, 0x04034b50, true); // local file header signature
    ldv.setUint16(4, 20, true); // version needed
    ldv.setUint16(6, 0, true); // flags
    ldv.setUint16(8, 0, true); // compression: 0 = stored
    ldv.setUint16(10, 0, true); // mod time
    ldv.setUint16(12, 0x21, true); // mod date (1980-01-01)
    ldv.setUint32(14, crc, true);
    ldv.setUint32(18, size, true); // compressed size
    ldv.setUint32(22, size, true); // uncompressed size
    ldv.setUint16(26, nameBytes.length, true);
    ldv.setUint16(28, 0, true); // extra length
    local.set(nameBytes, 30);
    local.set(data, 30 + nameBytes.length);
    locals.push(local);

    const central = new Uint8Array(46 + nameBytes.length);
    const cdv = new DataView(central.buffer);
    cdv.setUint32(0, 0x02014b50, true); // central dir signature
    cdv.setUint16(4, 20, true); // version made by
    cdv.setUint16(6, 20, true); // version needed
    cdv.setUint16(8, 0, true); // flags
    cdv.setUint16(10, 0, true); // compression
    cdv.setUint16(12, 0, true); // mod time
    cdv.setUint16(14, 0x21, true); // mod date
    cdv.setUint32(16, crc, true);
    cdv.setUint32(20, size, true);
    cdv.setUint32(24, size, true);
    cdv.setUint16(28, nameBytes.length, true);
    cdv.setUint16(30, 0, true); // extra len
    cdv.setUint16(32, 0, true); // comment len
    cdv.setUint16(34, 0, true); // disk number start
    cdv.setUint16(36, 0, true); // internal attrs
    cdv.setUint32(38, 0, true); // external attrs
    cdv.setUint32(42, offset, true); // local header offset
    central.set(nameBytes, 46);
    centrals.push(central);

    offset += local.length;
  }

  const cdSize = centrals.reduce((s, c) => s + c.length, 0);
  const cdOffset = offset;
  const eocd = new Uint8Array(22);
  const edv = new DataView(eocd.buffer);
  edv.setUint32(0, 0x06054b50, true); // EOCD signature
  edv.setUint16(4, 0, true); // disk number
  edv.setUint16(6, 0, true); // disk with CD
  edv.setUint16(8, files.length, true); // entries this disk
  edv.setUint16(10, files.length, true); // total entries
  edv.setUint32(12, cdSize, true);
  edv.setUint32(16, cdOffset, true);
  edv.setUint16(20, 0, true); // comment len

  let total = offset + cdSize + eocd.length;
  const out = new Uint8Array(total);
  let p = 0;
  for (const l of locals) {
    out.set(l, p);
    p += l.length;
  }
  for (const c of centrals) {
    out.set(c, p);
    p += c.length;
  }
  out.set(eocd, p);
  return out;
}

// ── XML helpers ──────────────────────────────────────────────────────────────
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

function escapeXml(s) {
  return String(s)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "") // strip chars illegal in XML 1.0 (avoids repair prompt)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// ── OOXML parts ──────────────────────────────────────────────────────────────
function contentTypesXml(exts) {
  const media = [];
  if (exts.has("png"))
    media.push('<Default Extension="png" ContentType="image/png"/>');
  if (exts.has("jpeg"))
    media.push('<Default Extension="jpeg" ContentType="image/jpeg"/>');
  return (
    XML_DECL +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    media.join("") +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    "</Types>"
  );
}

function rootRelsXml() {
  return (
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    "</Relationships>"
  );
}

function documentRelsXml(rels) {
  const fixed =
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>';
  const dyn = rels
    .map(
      (r) =>
        `<Relationship Id="${r.id}" Type="${r.type}" Target="${escapeXml(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`,
    )
    .join("");
  return (
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    fixed +
    dyn +
    "</Relationships>"
  );
}

function corePropsXml(title) {
  return (
    XML_DECL +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/">' +
    `<dc:title>${escapeXml(title)}</dc:title>` +
    "<dc:creator>Markdown Viewer</dc:creator>" +
    "</cp:coreProperties>"
  );
}

function documentXml(bodyXml) {
  return (
    XML_DECL +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    "<w:body>" +
    bodyXml +
    "<w:sectPr>" +
    '<w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    "</w:sectPr>" +
    "</w:body></w:document>"
  );
}

function stylesXml() {
  const heading = (id, name, lvl, halfPt, color) =>
    `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>` +
    `<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${lvl}"/></w:pPr>` +
    `<w:rPr><w:b/><w:color w:val="${color}"/><w:sz w:val="${halfPt}"/><w:szCs w:val="${halfPt}"/></w:rPr></w:style>`;
  const border = 'w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"';
  return (
    XML_DECL +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    heading("Heading1", "heading 1", 0, 48, "111111") +
    heading("Heading2", "heading 2", 1, 36, "111111") +
    heading("Heading3", "heading 3", 2, 28, "111111") +
    heading("Heading4", "heading 4", 3, 24, "111111") +
    heading("Heading5", "heading 5", 4, 22, "333333") +
    heading("Heading6", "heading 6", 5, 22, "555555") +
    '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="12" w:color="D0D7DE"/></w:pBdr><w:ind w:left="480"/></w:pPr>' +
    '<w:rPr><w:i/><w:color w:val="57606A"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:spacing w:after="0"/><w:contextualSpacing/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="CodeBlock"><w:name w:val="Code Block"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F6F8FA"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="CodeChar"><w:name w:val="Code Char"/>' +
    '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:shd w:val="clear" w:color="auto" w:fill="EFF1F3"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0969DA"/><w:u w:val="single"/></w:rPr></w:style>' +
    '<w:style w:type="table" w:default="1" w:styleId="TableGrid"><w:name w:val="Table Grid"/>' +
    `<w:tblPr><w:tblBorders><w:top ${border}/><w:left ${border}/><w:bottom ${border}/><w:right ${border}/><w:insideH ${border}/><w:insideV ${border}/></w:tblBorders></w:tblPr></w:style>` +
    "</w:styles>"
  );
}

function numberingXml(orderedNumIds) {
  const bulletChars = ["•", "◦", "▪", "•", "◦", "▪", "•", "◦", "▪"];
  let bulletLvls = "",
    decimalLvls = "";
  for (let i = 0; i < 9; i++) {
    const ind = 720 * (i + 1);
    bulletLvls +=
      `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${bulletChars[i]}"/>` +
      `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${ind}" w:hanging="360"/></w:pPr></w:lvl>`;
    decimalLvls +=
      `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%${i + 1}."/>` +
      `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${ind}" w:hanging="360"/></w:pPr></w:lvl>`;
  }
  const ordered = orderedNumIds
    .map(
      (id) =>
        `<w:num w:numId="${id}"><w:abstractNumId w:val="1"/>` +
        `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`,
    )
    .join("");
  return (
    XML_DECL +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:abstractNum w:abstractNumId="0">${bulletLvls}</w:abstractNum>` +
    `<w:abstractNum w:abstractNumId="1">${decimalLvls}</w:abstractNum>` +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    ordered +
    "</w:numbering>"
  );
}

// ── Package assembler ────────────────────────────────────────────────────────
function assembleDocx({ bodyXml, rels, media, exts, orderedNumIds, title }) {
  const enc = new TextEncoder();
  const files = [
    { name: "[Content_Types].xml", bytes: enc.encode(contentTypesXml(exts)) },
    { name: "_rels/.rels", bytes: enc.encode(rootRelsXml()) },
    { name: "docProps/core.xml", bytes: enc.encode(corePropsXml(title)) },
    { name: "word/document.xml", bytes: enc.encode(documentXml(bodyXml)) },
    {
      name: "word/_rels/document.xml.rels",
      bytes: enc.encode(documentRelsXml(rels)),
    },
    { name: "word/styles.xml", bytes: enc.encode(stylesXml()) },
    {
      name: "word/numbering.xml",
      bytes: enc.encode(numberingXml(orderedNumIds)),
    },
  ];
  for (const m of media) files.push({ name: "word/" + m.name, bytes: m.bytes });
  return zipStore(files);
}

// ── EMU / sizing ─────────────────────────────────────────────────────────────
const MAX_IMG_W_EMU = 5943600; // 6.5in content width
function emuFromPx(px) {
  return Math.round(px * 9525);
}

// ── Images ───────────────────────────────────────────────────────────────────
function extFromMime(mime) {
  if (/png/i.test(mime)) return "png";
  if (/jpe?g/i.test(mime)) return "jpeg";
  return null;
}

async function bitmapToPng(blob) {
  const bmp = await createImageBitmap(blob);
  const canvas = document.createElement("canvas");
  canvas.width = bmp.width;
  canvas.height = bmp.height;
  const ctx2d = canvas.getContext("2d");
  ctx2d.drawImage(bmp, 0, 0);
  const png = await new Promise((res, rej) =>
    canvas.toBlob(
      (b) => (b ? res(b) : rej(new Error("toBlob failed"))),
      "image/png",
    ),
  );
  return {
    bytes: new Uint8Array(await png.arrayBuffer()),
    w: bmp.width,
    h: bmp.height,
  };
}

async function imgToBytes(imgEl) {
  const src = imgEl.currentSrc || imgEl.getAttribute("src") || "";
  if (!src) return null;
  try {
    const resp = await fetch(src);
    if (!resp.ok) throw new Error("HTTP " + resp.status);
    const blob = await resp.blob();
    let ext = extFromMime(blob.type);
    if (ext === "png" || ext === "jpeg") {
      const bmp = await createImageBitmap(blob);
      const bytes = new Uint8Array(await blob.arrayBuffer());
      const dims = { w: bmp.width, h: bmp.height };
      return { bytes, ext, wPx: dims.w, hPx: dims.h };
    }
    // gif/webp/svg/bmp → normalize to PNG
    const norm = await bitmapToPng(blob);
    return { bytes: norm.bytes, ext: "png", wPx: norm.w, hPx: norm.h };
  } catch (e) {
    return null;
  }
}

// Scale (wPx,hPx) down so width never exceeds the content area; return EMUs.
function fitEmu(wPx, hPx) {
  let cx = emuFromPx(wPx),
    cy = emuFromPx(hPx);
  if (cx > MAX_IMG_W_EMU) {
    cy = Math.round(cy * (MAX_IMG_W_EMU / cx));
    cx = MAX_IMG_W_EMU;
  }
  return { cx: Math.max(1, cx), cy: Math.max(1, cy) };
}

function drawingXml(rId, cx, cy, docPrId, inline) {
  const anchor = inline ? "wp:inline" : "wp:inline"; // block images are inline in their own paragraph
  return (
    `<w:r><w:drawing><${anchor} distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${docPrId}" name="Picture ${docPrId}"/>` +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:nvPicPr><pic:cNvPr id="${docPrId}" name="Picture ${docPrId}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    `</pic:pic></a:graphicData></a:graphic></${anchor}></w:drawing></w:r>`
  );
}

async function imageRun(imgEl, ctx) {
  const data = await imgToBytes(imgEl);
  if (!data) {
    ctx.skipped++;
    const alt = imgEl.getAttribute("alt") || "image";
    return runXml(`[image: ${alt} — could not embed]`, {
      i: true,
      color: "999999",
    });
  }
  const rId = ctx.addImage(data.bytes, data.ext);
  const { cx, cy } = fitEmu(data.wPx, data.hPx);
  return drawingXml(rId, cx, cy, ctx.newDocPrId(), true);
}

// ── Build context (accumulates rels, media, numbering, ids, skip count) ──────
function createCtx() {
  return {
    rels: [], // {id, type, target, external}
    media: [], // {name, bytes}
    exts: new Set(), // 'png' | 'jpeg'
    orderedNumIds: [], // numIds for ordered lists (restart numbering)
    skipped: 0, // count of items that fell back to text
    _rid: 3, // rId1=styles, rId2=numbering are fixed
    _mediaN: 0,
    _docPr: 0,
    _numId: 1, // numId 1 = bullets; newOrderedNumId() returns 2, 3, …
    addImage(bytes, ext) {
      const id = "rId" + this._rid++;
      const name = `media/img${++this._mediaN}.${ext}`;
      this.media.push({ name, bytes });
      this.exts.add(ext);
      this.rels.push({
        id,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        target: name,
      });
      return id;
    },
    addHyperlink(url) {
      const id = "rId" + this._rid++;
      this.rels.push({
        id,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: url,
        external: true,
      });
      return id;
    },
    newOrderedNumId() {
      const n = ++this._numId;
      this.orderedNumIds.push(n);
      return n;
    },
    newDocPrId() {
      return ++this._docPr;
    },
  };
}

// ── Run / paragraph helpers ──────────────────────────────────────────────────
// rpr: { b, i, strike, code, color } → run properties XML
function rprXml(rpr) {
  if (!rpr) return "";
  let x = "";
  if (rpr.code) x += '<w:rStyle w:val="CodeChar"/>';
  else if (rpr.link) x += '<w:rStyle w:val="Hyperlink"/>';
  if (rpr.b) x += "<w:b/>";
  if (rpr.i) x += "<w:i/>";
  if (rpr.strike) x += "<w:strike/>";
  if (rpr.color) x += `<w:color w:val="${rpr.color}"/>`;
  return x ? `<w:rPr>${x}</w:rPr>` : "";
}

function runXml(text, rpr) {
  if (!text) return "";
  return `<w:r>${rprXml(rpr)}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

function paragraph(styleId, innerXml, extraPPr) {
  const pPr =
    (styleId ? `<w:pStyle w:val="${styleId}"/>` : "") + (extraPPr || "");
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${innerXml}</w:p>`;
}

// ── Inline walker: DOM inline nodes → concatenated <w:r> runs ─────────────────
async function inlineToRuns(node, ctx, rpr) {
  let out = "";
  for (const child of node.childNodes) {
    if (child.nodeType === 3) {
      // text
      out += runXml(child.nodeValue, rpr);
      continue;
    }
    if (child.nodeType !== 1) continue; // skip comments etc.
    const tag = child.tagName.toLowerCase();
    if (tag === "br") {
      out += "<w:r><w:br/></w:r>";
      continue;
    }
    if (tag === "strong" || tag === "b") {
      out += await inlineToRuns(child, ctx, { ...rpr, b: true });
      continue;
    }
    if (tag === "em" || tag === "i") {
      out += await inlineToRuns(child, ctx, { ...rpr, i: true });
      continue;
    }
    if (tag === "del" || tag === "s" || tag === "strike") {
      out += await inlineToRuns(child, ctx, { ...rpr, strike: true });
      continue;
    }
    if (tag === "code") {
      out += await inlineToRuns(child, ctx, { ...rpr, code: true });
      continue;
    }
    if (tag === "a") {
      const href = child.getAttribute("href") || "";
      if (/^https?:|^mailto:/i.test(href)) {
        const id = ctx.addHyperlink(href);
        out += `<w:hyperlink r:id="${id}">${await inlineToRuns(child, ctx, { ...rpr, link: true })}</w:hyperlink>`;
      } else {
        out += await inlineToRuns(child, ctx, rpr); // internal anchors: keep text only
      }
      continue;
    }
    if (tag === 'img') { out += await imageRun(child, ctx); continue; }        // defensive: raw inline <img>
    if (child.classList && child.classList.contains('img-download-btn')) { continue; } // skip UI chrome
    if (child.classList && child.classList.contains('math-inline')) { out += await mathRun(child, ctx, false); continue; }
    if (child.classList && child.classList.contains('math-block')) { out += await mathRun(child, ctx, false); continue; } // a block $$…$$ div can land mid-line; emit inline so we never nest <w:p> inside a run slot
    if (child.classList && child.classList.contains('math-error')) { out += runXml(child.textContent || '', { code: true, color: 'CF222E' }); continue; }
    // Inline <img> (Task 9) and inline math (Task 11) branches get inserted here,
    // BEFORE this generic recurse fallback. Both walkers are async — always await.
    out += await inlineToRuns(child, ctx, rpr);
  }
  return out;
}

// ── Lists ────────────────────────────────────────────────────────────────────
async function listToOoxml(listEl, ctx, ilvl, numId) {
  const ordered = listEl.tagName.toLowerCase() === "ol";
  // Top-level ordered list gets its own numId so numbering restarts at 1.
  const myNumId = ordered ? (numId != null ? numId : ctx.newOrderedNumId()) : 1;
  let out = "";
  for (const li of listEl.children) {
    if (li.tagName.toLowerCase() !== "li") continue;

    // Task-list item: <li class="task-item"><label><input ...> text</label>…</li>
    const isTask = li.classList.contains("task-item");
    const checkbox = li.querySelector(
      ':scope > label > input[type="checkbox"]',
    );
    const glyph = isTask ? (checkbox && checkbox.checked ? "☑ " : "☐ ") : "";

    // Inline content of this item = the <li>/<label> minus nested lists.
    const source = li.querySelector(":scope > label") || li;
    let runs = "";
    if (glyph) runs += runXml(glyph, null);
    for (const node of source.childNodes) {
      if (node.nodeType === 1) {
        const t = node.tagName.toLowerCase();
        if (t === "ul" || t === "ol" || t === "input") continue; // nested lists / the checkbox handled separately
      }
      runs += await inlineToRuns({ childNodes: [node] }, ctx, null);
    }

    const numPr = `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${isTask ? 1 : myNumId}"/></w:numPr>`;
    // Task items look best without a bullet marker → use ilvl but suppress via a plain paragraph.
    if (isTask) {
      out += paragraph(
        "ListParagraph",
        runs,
        `<w:ind w:left="${720 * (ilvl + 1)}" w:hanging="360"/>`,
      );
    } else {
      out += paragraph("ListParagraph", runs, numPr);
    }

    // Recurse into nested lists.
    for (const child of li.children) {
      const ct = child.tagName.toLowerCase();
      if (ct === "ul" || ct === "ol")
        out += await listToOoxml(child, ctx, ilvl + 1, null);
    }
  }
  return out;
}

// ── Tables ───────────────────────────────────────────────────────────────────
async function tableToOoxml(tableEl, ctx) {
  const rows = [...tableEl.querySelectorAll("tr")];
  // CT_Tbl requires <w:tblGrid> (one <w:gridCol/> per column) BETWEEN tblPr and the
  // rows. Omitting it triggers Word's "unreadable content — recover?" repair prompt
  // on every table.
  const firstRow = rows[0];
  const colCount = firstRow ? firstRow.querySelectorAll("th, td").length : 1;
  let gridCols = "";
  for (let i = 0; i < colCount; i++) gridCols += "<w:gridCol/>";

  let trs = "";
  for (const tr of rows) {
    // for...of (not .forEach) so we can await inlineToRuns per cell
    let tcs = "";
    for (const cell of tr.querySelectorAll("th, td")) {
      const isHeader = cell.tagName.toLowerCase() === "th";
      const align = (
        cell.style.textAlign ||
        cell.getAttribute("align") ||
        ""
      ).toLowerCase();
      const jc =
        align === "center"
          ? '<w:jc w:val="center"/>'
          : align === "right"
            ? '<w:jc w:val="right"/>'
            : "";
      const shd = isHeader
        ? '<w:shd w:val="clear" w:color="auto" w:fill="F0F0F0"/>'
        : "";
      const runs =
        (await inlineToRuns(cell, ctx, isHeader ? { b: true } : null)) || "";
      const p = `<w:p>${jc ? `<w:pPr>${jc}</w:pPr>` : ""}${runs}</w:p>`;
      tcs += `<w:tc><w:tcPr>${shd}</w:tcPr>${p}</w:tc>`;
    }
    trs += `<w:tr>${tcs}</w:tr>`;
  }
  return (
    '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>' +
    '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
    "</w:tblPr>" +
    "<w:tblGrid>" +
    gridCols +
    "</w:tblGrid>" +
    trs +
    "</w:tbl>" +
    "<w:p/>" // spacer: Word merges adjacent <w:tbl> with nothing between; also satisfies "table not last body element"
  );
}

// ── Code blocks (syntax-highlighted) ─────────────────────────────────────────
// Maps the renderer's sh-* token classes to GitHub-light hex colors.
const SH_COLORS = {
  "sh-kw": "CF222E",
  "sh-str": "0A3069",
  "sh-cmt": "6E7781",
  "sh-num": "0550AE",
  "sh-fn": "8250DF",
  "sh-tag": "116329",
  "sh-attr": "953800",
  "sh-prop": "8250DF",
  "sh-key": "0550AE",
};
const CODE_DEFAULT = "24292F";

function codeBlockToOoxml(wrapEl, ctx) {
  const codeEl = wrapEl.querySelector("pre > code");
  if (!codeEl) return "";
  // Build a flat list of {text, color} runs from the highlighted markup,
  // then split into lines on '\n' so each line is its own shaded paragraph.
  const runs = [];
  (function walk(node, color) {
    for (const child of node.childNodes) {
      if (child.nodeType === 3) {
        runs.push({ text: child.nodeValue, color });
        continue;
      }
      if (child.nodeType !== 1) continue;
      let c = color;
      for (const cls of child.classList)
        if (SH_COLORS[cls]) {
          c = SH_COLORS[cls];
          break;
        }
      walk(child, c);
    }
  })(codeEl, CODE_DEFAULT);

  // Split runs across newlines into per-line arrays.
  const lines = [[]];
  for (const r of runs) {
    const parts = r.text.split("\n");
    for (let i = 0; i < parts.length; i++) {
      if (i > 0) lines.push([]);
      if (parts[i])
        lines[lines.length - 1].push({ text: parts[i], color: r.color });
    }
  }
  // Drop a trailing empty line (from a final newline in the code text).
  if (lines.length > 1 && lines[lines.length - 1].length === 0) lines.pop();

  let out = "";
  for (const line of lines) {
    let inner = "";
    for (const seg of line) {
      inner += runXml(seg.text, {
        code: true,
        color: seg.color === CODE_DEFAULT ? null : seg.color,
      });
    }
    out += paragraph("CodeBlock", inner);
  }
  return out;
}

// ── SVG → PNG rasterization ──────────────────────────────────────────────────
function loadImage(url) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => res(img);
    img.onerror = () => rej(new Error("image load failed"));
    img.src = url;
  });
}

async function svgToPng(svgEl, scale) {
  scale = scale || 2;
  const rect = svgEl.getBoundingClientRect();
  const w = Math.max(
    1,
    Math.ceil(rect.width || parseFloat(svgEl.getAttribute("width")) || 300),
  );
  const h = Math.max(
    1,
    Math.ceil(rect.height || parseFloat(svgEl.getAttribute("height")) || 200),
  );
  const clone = svgEl.cloneNode(true);
  clone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  clone.setAttribute("width", w);
  clone.setAttribute("height", h);
  const svgStr = new XMLSerializer().serializeToString(clone);
  const url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgStr);
  const img = await loadImage(url);
  const canvas = document.createElement("canvas");
  canvas.width = w * scale;
  canvas.height = h * scale;
  const c = canvas.getContext("2d");
  c.fillStyle = "#ffffff";
  c.fillRect(0, 0, canvas.width, canvas.height);
  c.drawImage(img, 0, 0, canvas.width, canvas.height);
  const png = await new Promise((res, rej) =>
    canvas.toBlob(
      (b) => (b ? res(b) : rej(new Error("toBlob failed"))),
      "image/png",
    ),
  );
  return { bytes: new Uint8Array(await png.arrayBuffer()), wPx: w, hPx: h };
}

async function diagramBlock(el, ctx) {
  const svg = el.querySelector("svg");
  if (svg) {
    try {
      const { bytes, wPx, hPx } = await svgToPng(svg, 2);
      const rId = ctx.addImage(bytes, "png");
      const { cx, cy } = fitEmu(wPx, hPx);
      return `<w:p>${drawingXml(rId, cx, cy, ctx.newDocPrId(), true)}</w:p>`;
    } catch (e) {
      // fall through to source fallback
    }
  }
  // Fallback: render the mermaid source as a code block.
  ctx.skipped++;
  const source = el.getAttribute("data-diagram") || el.textContent || "";
  let out = "";
  for (const line of String(source).split("\n"))
    out += paragraph("CodeBlock", runXml(line, { code: true }));
  return out;
}

// ── KaTeX math rasterization ─────────────────────────────────────────────────
function bytesToBase64(bytes) {
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(bin);
}

let _katexCssCache = null;
async function fetchKatexCss() {
  if (_katexCssCache != null) return _katexCssCache;
  const resp = await fetch(chrome.runtime.getURL("katex/katex.min.css"));
  _katexCssCache = await resp.text();
  return _katexCssCache;
}

const _fontUriCache = new Map();
async function fontDataUri(file) {
  if (_fontUriCache.has(file)) return _fontUriCache.get(file);
  const resp = await fetch(chrome.runtime.getURL("katex/fonts/" + file));
  const bytes = new Uint8Array(await resp.arrayBuffer());
  const uri = "data:font/woff2;base64," + bytesToBase64(bytes);
  _fontUriCache.set(file, uri);
  return uri;
}

// Build inlined CSS: layout rules verbatim + @font-face (used families only) with data-URI woff2.
async function katexInlineCss(katexEl) {
  const raw = await fetchKatexCss();
  const used = new Set([
    "KaTeX_Main",
    "KaTeX_Math",
    "KaTeX_Size1",
    "KaTeX_Size2",
  ]);
  katexEl.querySelectorAll("*").forEach((n) => {
    const ff = getComputedStyle(n).fontFamily || "";
    ff.split(",").forEach((f) => {
      const name = f.trim().replace(/["']/g, "");
      if (name.indexOf("KaTeX_") === 0) used.add(name);
    });
  });
  const faceRe = /@font-face\s*\{[^}]*\}/g;
  const layout = raw.replace(faceRe, "");
  const faces = raw.match(faceRe) || [];
  let inlined = "";
  for (const block of faces) {
    const famM = block.match(/font-family\s*:\s*([^;}]+)/i);
    if (!famM) continue;
    const fam = famM[1].trim().replace(/["']/g, "");
    if (!used.has(fam)) continue;
    const woff2M = block.match(
      /url\(\s*['"]?fonts\/([^)'"]+\.woff2)['"]?\s*\)/i,
    );
    if (!woff2M) {
      inlined += block;
      continue;
    }
    const uri = await fontDataUri(woff2M[1]);
    const wM = block.match(/font-weight\s*:\s*([^;}]+)/i);
    const sM = block.match(/font-style\s*:\s*([^;}]+)/i);
    inlined +=
      `@font-face{font-family:${fam};src:url(${uri}) format("woff2");` +
      `font-weight:${wM ? wM[1].trim() : "normal"};font-style:${sM ? sM[1].trim() : "normal"}}`;
  }
  return inlined + layout;
}

async function mathToPng(katexEl) {
  const css = await katexInlineCss(katexEl);
  const rect = katexEl.getBoundingClientRect();
  const pad = 2;
  const w = Math.max(1, Math.ceil(rect.width) + pad * 2);
  const h = Math.max(1, Math.ceil(rect.height) + pad * 2);
  const inner = new XMLSerializer().serializeToString(katexEl);
  const html = `<div xmlns="http://www.w3.org/1999/xhtml" style="display:inline-block;padding:${pad}px;color:#111111;background:#ffffff;">${inner}</div>`;
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">` +
    `<foreignObject width="${w}" height="${h}"><style>${css}</style>${html}</foreignObject></svg>`;
  const url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
  const img = await loadImage(url);
  const scale = 2;
  const canvas = document.createElement("canvas");
  canvas.width = w * scale;
  canvas.height = h * scale;
  const c = canvas.getContext("2d");
  c.fillStyle = "#ffffff";
  c.fillRect(0, 0, canvas.width, canvas.height);
  c.drawImage(img, 0, 0, canvas.width, canvas.height);
  const png = await new Promise((res, rej) =>
    canvas.toBlob(
      (b) => (b ? res(b) : rej(new Error("toBlob failed"))),
      "image/png",
    ),
  );
  return { bytes: new Uint8Array(await png.arrayBuffer()), wPx: w, hPx: h };
}

async function mathRun(mathEl, ctx, display) {
  // KaTeX failure renders <span class="math-error">tex</span> (no .katex, no annotation) —
  // emit it as red monospace text directly; don't waste a rasterize attempt on it.
  const errEl = mathEl.classList.contains("math-error")
    ? mathEl
    : mathEl.querySelector(".math-error");
  if (errEl) {
    const runs = runXml(errEl.textContent || "", { code: true, color: "CF222E" });
    return display ? paragraph("CodeBlock", runs) : runs;
  }
  const katexEl = mathEl.querySelector(".katex") || mathEl;
  try {
    const { bytes, wPx, hPx } = await mathToPng(katexEl);
    const rId = ctx.addImage(bytes, "png");
    const { cx, cy } = fitEmu(wPx, hPx);
    const run = drawingXml(rId, cx, cy, ctx.newDocPrId(), true);
    return display
      ? `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${run}</w:p>`
      : run;
  } catch (e) {
    ctx.skipped++;
    const ann = mathEl.querySelector(
      'annotation[encoding="application/x-tex"]',
    );
    const tex = ann ? ann.textContent : mathEl.textContent || "";
    const runs = runXml(tex, { code: true });
    return display ? paragraph("CodeBlock", runs) : runs;
  }
}

// ── Block walker: children of .md-body → body XML ────────────────────────────
async function blocksToOoxml(container, ctx) {
  let out = "";
  for (const el of container.children) {
    const tag = el.tagName.toLowerCase();
    if (/^h[1-6]$/.test(tag)) {
      out += paragraph("Heading" + tag[1], await inlineToRuns(el, ctx, null));
    } else if (tag === "p") {
      out += paragraph(null, await inlineToRuns(el, ctx, null));
    } else if (tag === 'blockquote') {
      // Emit each child block as a Quote-styled paragraph (await — inlineToRuns is async).
      // v1 limitation: nested lists/code/tables inside a quote are flattened to Quote
      // paragraph text rather than reproduced as nested structure — but never dropped.
      const kids = el.children.length ? [...el.children] : [el];
      for (const child of kids) out += paragraph('Quote', await inlineToRuns(child, ctx, null));
    } else if (tag === 'hr') {
      out += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="D0D7DE"/></w:pBdr></w:pPr></w:p>';
      // ── Later tasks insert their block branches here, before the final else ──
    } else if (tag === 'ul' || tag === 'ol') {
      out += await listToOoxml(el, ctx, 0, null);
    } else if (tag === 'table') {
      out += await tableToOoxml(el, ctx);
    } else if (el.classList.contains('cb-wrap')) {
      out += codeBlockToOoxml(el, ctx);
    } else if (el.classList.contains('img-wrap')) {
      const img = el.querySelector('img');
      out += img ? `<w:p>${await imageRun(img, ctx)}</w:p>` : '';
    } else if (el.classList.contains('mermaid-wrap') || el.classList.contains('mermaid-pending') || el.classList.contains('mermaid-error')) {
      out += await diagramBlock(el, ctx);
    } else if (el.classList.contains("math-block")) {
      out += await mathRun(el, ctx, true);
    } else {
      // Fallback for unrecognized blocks: render their text as a paragraph.
      out += paragraph(null, await inlineToRuns(el, ctx, null));
    }
  }
  return out;
}

// ── Orchestrator ─────────────────────────────────────────────────────────────
async function exportDocx({ container, filename, title }) {
  const ctx = createCtx();
  const bodyXml = await blocksToOoxml(container, ctx);
  const bytes = assembleDocx({
    bodyXml,
    rels: ctx.rels,
    media: ctx.media,
    exts: ctx.exts,
    orderedNumIds: ctx.orderedNumIds,
    title: title || filename,
  });
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  return { skipped: ctx.skipped };
}

// ── Environment exports ──────────────────────────────────────────────────────
if (typeof module !== "undefined" && module.exports) {
  module.exports = { crc32, zipStore, escapeXml, contentTypesXml, rootRelsXml, documentRelsXml, stylesXml, numberingXml, documentXml, corePropsXml, assembleDocx };
}

if (typeof window !== "undefined") {
  window.MdDocx = { export: exportDocx };
}
