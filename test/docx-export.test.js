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
