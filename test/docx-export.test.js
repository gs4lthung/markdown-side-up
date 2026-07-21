"use strict";
const test = require("node:test");
const assert = require("node:assert");
const { crc32, zipStore } = require("../docx-export.js");

const enc = (s) => new TextEncoder().encode(s);

// A minimal stored-ZIP reader — TEST-ONLY, proves zipStore output is a real ZIP.
function readZip(buf) {
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  // Find End Of Central Directory (0x06054b50), scanning from the end.
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (dv.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  assert.ok(eocd >= 0, "EOCD signature present");
  const count = dv.getUint16(eocd + 10, true);
  let cd = dv.getUint32(eocd + 16, true);
  const out = {};
  for (let n = 0; n < count; n++) {
    assert.strictEqual(
      dv.getUint32(cd, true),
      0x02014b50,
      "central dir signature",
    );
    const crc = dv.getUint32(cd + 16, true) >>> 0;
    const nameLen = dv.getUint16(cd + 28, true);
    const extraLen = dv.getUint16(cd + 30, true);
    const cmtLen = dv.getUint16(cd + 32, true);
    const lho = dv.getUint32(cd + 42, true);
    const name = new TextDecoder().decode(
      buf.subarray(cd + 46, cd + 46 + nameLen),
    );
    // Read the local header to recover the data bytes.
    assert.strictEqual(
      dv.getUint32(lho, true),
      0x04034b50,
      "local header signature",
    );
    const lNameLen = dv.getUint16(lho + 26, true);
    const lExtraLen = dv.getUint16(lho + 28, true);
    const size = dv.getUint32(lho + 22, true); // compressed size (== uncompressed, stored)
    const dataStart = lho + 30 + lNameLen + lExtraLen;
    out[name] = { crc, bytes: buf.subarray(dataStart, dataStart + size) };
    cd += 46 + nameLen + extraLen + cmtLen;
  }
  return out;
}

test("crc32 known vectors", () => {
  assert.strictEqual(crc32(enc("")) >>> 0, 0x00000000);
  assert.strictEqual(crc32(enc("123456789")) >>> 0, 0xcbf43926);
});

test("zipStore round-trips names and bytes", () => {
  const files = [
    { name: "a.txt", bytes: enc("hello") },
    { name: "dir/b.xml", bytes: enc("<x/>") },
  ];
  const zip = zipStore(files);
  const read = readZip(zip);
  assert.deepStrictEqual(Object.keys(read).sort(), ["a.txt", "dir/b.xml"]);
  assert.strictEqual(new TextDecoder().decode(read["a.txt"].bytes), "hello");
  assert.strictEqual(read["dir/b.xml"].crc, crc32(enc("<x/>")) >>> 0);
});

const {
  escapeXml,
  contentTypesXml,
  rootRelsXml,
  documentRelsXml,
  stylesXml,
  numberingXml,
  documentXml,
  corePropsXml,
} = require("../docx-export.js");

test("escapeXml handles the five predefined entities", () => {
  assert.strictEqual(
    escapeXml(`<a b="c" d='e' & f>`),
    "&lt;a b=&quot;c&quot; d=&apos;e&apos; &amp; f&gt;",
  );
});

test("contentTypesXml declares media extensions only when used", () => {
  const none = contentTypesXml(new Set());
  assert.ok(!none.includes("image/png"));
  const png = contentTypesXml(new Set(["png", "jpeg"]));
  assert.ok(png.includes('Extension="png" ContentType="image/png"'));
  assert.ok(png.includes('Extension="jpeg" ContentType="image/jpeg"'));
  assert.ok(png.includes("/word/document.xml"));
});

test("documentRelsXml always includes styles + numbering and appends dynamic rels", () => {
  const xml = documentRelsXml([
    {
      id: "rId3",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      target: "media/img1.png",
    },
    {
      id: "rId4",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target: "https://example.com",
      external: true,
    },
  ]);
  assert.ok(xml.includes('Id="rId1"') && xml.includes("styles.xml"));
  assert.ok(xml.includes('Id="rId2"') && xml.includes("numbering.xml"));
  assert.ok(xml.includes('Id="rId3"') && xml.includes("media/img1.png"));
  assert.ok(xml.includes('TargetMode="External"'));
});

test("numberingXml emits one <w:num> per ordered list plus the shared bullet num", () => {
  const xml = numberingXml([2, 3]);
  assert.ok(xml.includes('w:numId="1"')); // bullets
  assert.ok(xml.includes('w:numId="2"')); // first ordered
  assert.ok(xml.includes('w:numId="3"')); // second ordered
  assert.ok(xml.includes("startOverride")); // restart each ordered list
  assert.ok(
    xml.includes('w:abstractNumId="0"') && xml.includes('w:abstractNumId="1"'),
  );
});

test("documentXml wraps the body and includes a sectPr", () => {
  const xml = documentXml("<w:p/>");
  assert.ok(xml.includes("<w:body><w:p/>"));
  assert.ok(xml.includes("w:pgSz"));
  assert.ok(xml.startsWith("<?xml"));
});

test("stylesXml and corePropsXml are non-empty and well-formed at the root", () => {
  assert.ok(stylesXml().includes('w:styleId="Heading1"'));
  assert.ok(stylesXml().includes('w:styleId="CodeBlock"'));
  assert.ok(corePropsXml("My & Title").includes("My &amp; Title"));
});

test("stylesXml respects OOXML child-element ordering (xsd:sequence)", () => {
  const s = stylesXml();
  // Heading rPr: w:color must precede w:sz.
  assert.ok(s.includes('<w:color w:val="111111"/><w:sz'), "heading color before sz");
  // Quote pPr: w:pBdr must precede w:ind.
  const q = s.indexOf('w:styleId="Quote"');
  assert.ok(s.indexOf("<w:pBdr>", q) < s.indexOf("<w:ind ", q), "quote pBdr before ind");
  // CodeBlock pPr: w:shd must precede w:spacing.
  const cb = s.indexOf('w:styleId="CodeBlock"');
  assert.ok(s.indexOf("<w:shd ", cb) < s.indexOf("<w:spacing ", cb), "codeblock shd before spacing");
});

const { assembleDocx } = require("../docx-export.js");

test("assembleDocx produces a valid ZIP containing all required parts", () => {
  const bytes = assembleDocx({
    bodyXml: '<w:p><w:r><w:t xml:space="preserve">Hi</w:t></w:r></w:p>',
    rels: [],
    media: [],
    exts: new Set(),
    orderedNumIds: [],
    title: "Doc",
  });
  const read = readZip(bytes);
  for (const part of [
    "[Content_Types].xml",
    "_rels/.rels",
    "docProps/core.xml",
    "word/document.xml",
    "word/_rels/document.xml.rels",
    "word/styles.xml",
    "word/numbering.xml",
  ]) {
    assert.ok(read[part], "missing part: " + part);
  }
  const doc = new TextDecoder().decode(read["word/document.xml"].bytes);
  assert.ok(doc.includes('<w:t xml:space="preserve">Hi</w:t>'));
});

test("assembleDocx includes media files and their content-type declarations", () => {
  const bytes = assembleDocx({
    bodyXml: "<w:p/>",
    rels: [
      {
        id: "rId3",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        target: "media/img1.png",
      },
    ],
    media: [{ name: "media/img1.png", bytes: new Uint8Array([1, 2, 3]) }],
    exts: new Set(["png"]),
    orderedNumIds: [],
    title: "Doc",
  });
  const read = readZip(bytes);
  assert.ok(read["word/media/img1.png"], "media file present");
  const ct = new TextDecoder().decode(read["[Content_Types].xml"].bytes);
  assert.ok(ct.includes("image/png"));
  const rels = new TextDecoder().decode(
    read["word/_rels/document.xml.rels"].bytes,
  );
  assert.ok(rels.includes("rId3") && rels.includes("media/img1.png"));
});
