// Opens a .zip in the browser and hands back the Genetec files inside it, so a
// whole week of exports can be dropped as one zip.
//
// Sizes and offsets come from the zip's central directory, which every zip
// tool writes the same way (Windows, Mac and phone zips all work). Files are
// unpacked with the browser's own deflate support, and every file is checked
// against the size and CRC stored in the zip, so a damaged or cut off zip
// stops with a message instead of feeding part of a day's swipes into a
// report.

export class ZipError extends Error {
  constructor(code, message) {
    super(message);
    this.name = "ZipError";
    this.code = code;
  }
}

const END_SIG = 0x06054b50;
const DIR_SIG = 0x02014b50;
const LOCAL_SIG = 0x04034b50;

let crcTable = null;
function crc32(bytes) {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = crcTable[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

async function inflateRaw(bytes) {
  let stream;
  try {
    stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  } catch (_) {
    throw new ZipError("browser", "This browser can't open .zip files.");
  }
  try {
    return new Uint8Array(await new Response(stream).arrayBuffer());
  } catch (_) {
    throw new ZipError("damaged", "The zip is damaged.");
  }
}

// Names flagged as UTF 8 are read that way. Older zips use the DOS code page;
// plain names like the Genetec exports read the same either way.
const utf8 = new TextDecoder("utf-8");
const legacy = new TextDecoder("windows-1252");

// Every entry in the zip with its sizes, CRC and where its header sits.
export function listZip(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const len = bytes.byteLength;
  if (len < 22) throw new ZipError("notzip", "Not a zip file.");
  // The end record is 22 bytes plus a comment of up to 65,535 bytes
  let end = -1;
  for (let i = len - 22, stop = Math.max(0, len - 22 - 0xffff); i >= stop; i--) {
    if (view.getUint32(i, true) === END_SIG) { end = i; break; }
  }
  if (end < 0) throw new ZipError("notzip", "Not a zip file.");
  const count = view.getUint16(end + 10, true);
  const dirSize = view.getUint32(end + 12, true);
  const dirStart = view.getUint32(end + 16, true);
  if (count === 0xffff || dirSize === 0xffffffff || dirStart === 0xffffffff) {
    throw new ZipError("toobig", "The zip is too large to open here.");
  }
  if (dirStart + dirSize > end) throw new ZipError("damaged", "The zip is damaged.");
  const entries = [];
  let p = dirStart;
  for (let i = 0; i < count; i++) {
    if (p + 46 > end || view.getUint32(p, true) !== DIR_SIG) throw new ZipError("damaged", "The zip is damaged.");
    const flags = view.getUint16(p + 8, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const noteLen = view.getUint16(p + 32, true);
    const entry = {
      flags,
      method: view.getUint16(p + 10, true),
      crc: view.getUint32(p + 16, true),
      packedSize: view.getUint32(p + 20, true),
      size: view.getUint32(p + 24, true),
      offset: view.getUint32(p + 42, true),
      name: (flags & 0x800 ? utf8 : legacy).decode(bytes.subarray(p + 46, p + 46 + nameLen)),
    };
    if (entry.packedSize === 0xffffffff || entry.size === 0xffffffff || entry.offset === 0xffffffff) {
      throw new ZipError("toobig", "The zip is too large to open here.");
    }
    entries.push(entry);
    p += 46 + nameLen + extraLen + noteLen;
  }
  return entries;
}

// Unpacks one entry and checks it against the size and CRC the zip recorded.
export async function readEntry(bytes, entry) {
  if (entry.flags & 1) throw new ZipError("locked", "The zip is password protected.");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const p = entry.offset;
  if (p + 30 > bytes.byteLength || view.getUint32(p, true) !== LOCAL_SIG) {
    throw new ZipError("damaged", "The zip is damaged.");
  }
  const start = p + 30 + view.getUint16(p + 26, true) + view.getUint16(p + 28, true);
  if (start + entry.packedSize > bytes.byteLength) throw new ZipError("damaged", "The zip is damaged.");
  const packed = bytes.subarray(start, start + entry.packedSize);
  let data;
  if (entry.method === 0) data = packed;
  else if (entry.method === 8) data = await inflateRaw(packed);
  else throw new ZipError("method", "The zip uses a compression type the site can't read.");
  if (data.length !== entry.size || crc32(data) !== entry.crc) {
    throw new ZipError("damaged", "The zip is damaged.");
  }
  return data;
}

// The files in a zip whose names match `wanted`, as File objects named
// without their folder path. Mac zips carry "__MACOSX" copies and "._"
// shadow files, and an open workbook can leave a "~$" lock file; none of
// those are exports, so they are skipped. A zip inside the zip is opened too.
export async function filesFromZip(file, wanted, depth = 0) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const out = [];
  for (const entry of listZip(bytes)) {
    const parts = entry.name.split(/[\\/]/);
    const base = parts[parts.length - 1];
    if (!base) continue;                                   // a folder
    if (parts.includes("__MACOSX") || /^(\.|~\$)/.test(base)) continue;
    const inner = /\.zip$/i.test(base);
    if (inner ? depth >= 1 : !wanted.test(base)) continue;
    const data = await readEntry(bytes, entry);
    const f = new File([data], base);
    if (inner) out.push(...await filesFromZip(f, wanted, depth + 1));
    else out.push(f);
  }
  return out;
}
