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

// ── Environment exports ──────────────────────────────────────────────────────
if (typeof module !== "undefined" && module.exports) {
  module.exports = { crc32, zipStore, escapeXml, contentTypesXml, rootRelsXml, documentRelsXml, stylesXml, numberingXml, documentXml, corePropsXml };
}
