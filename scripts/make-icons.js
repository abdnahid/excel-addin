// Generates the placeholder ribbon/store icons (assets/icon-*.png): a teal rounded
// square with white ⇄ arrows. Pure Node, no image libraries. Replace with designed
// artwork before publishing to AppSource if you like.
const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const SIZES = [16, 32, 64, 80, 128];
const BG = [15, 108, 90];
const FG = [255, 255, 255];
const SS = 4; // supersampling per axis for anti-aliasing

// Shape tests in unit coordinates (0..1).
function inRoundedSquare(x, y) {
  const r = 0.2;
  const cx = Math.min(Math.max(x, r), 1 - r);
  const cy = Math.min(Math.max(y, r), 1 - r);
  return (x - cx) ** 2 + (y - cy) ** 2 <= r * r;
}

function inArrow(x, y, cy, pointsRight) {
  const u = pointsRight ? x : 1 - x;
  const shaft = u >= 0.2 && u <= 0.62 && Math.abs(y - cy) <= 0.055;
  const head = u >= 0.58 && u <= 0.82 && Math.abs(y - cy) <= (0.82 - u) * 0.62;
  return shaft || head;
}

function pixel(x, y) {
  if (!inRoundedSquare(x, y)) return null;
  return inArrow(x, y, 0.36, true) || inArrow(x, y, 0.64, false) ? FG : BG;
}

function crc32(buf) {
  let c;
  let crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

function png(size) {
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let py = 0; py < size; py++) {
    raw[py * (size * 4 + 1)] = 0; // filter: none
    for (let px = 0; px < size; px++) {
      let r = 0, g = 0, b = 0, a = 0;
      for (let sy = 0; sy < SS; sy++) {
        for (let sx = 0; sx < SS; sx++) {
          const c = pixel((px + (sx + 0.5) / SS) / size, (py + (sy + 0.5) / SS) / size);
          if (c) {
            r += c[0];
            g += c[1];
            b += c[2];
            a++;
          }
        }
      }
      const o = py * (size * 4 + 1) + 1 + px * 4;
      raw[o] = a ? Math.round(r / a) : 0;
      raw[o + 1] = a ? Math.round(g / a) : 0;
      raw[o + 2] = a ? Math.round(b / a) : 0;
      raw[o + 3] = Math.round((a / (SS * SS)) * 255);
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

const outDir = path.join(__dirname, "..", "assets");
fs.mkdirSync(outDir, { recursive: true });
for (const size of SIZES) {
  fs.writeFileSync(path.join(outDir, `icon-${size}.png`), png(size));
  console.log(`assets/icon-${size}.png`);
}
