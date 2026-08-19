import fs from 'fs';
import path from 'path';
import zlib from 'zlib';

/**
 * Creates an uncompressed / deflate PNG buffer from raw RGBA pixel data.
 */
function createPng(width, height, rgbaBuffer) {
  // PNG signature
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

  // IHDR chunk
  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(width, 0);
  ihdrData.writeUInt32BE(height, 4);
  ihdrData.writeUInt8(8, 8); // 8 bit depth
  ihdrData.writeUInt8(6, 9); // RGBA color type
  ihdrData.writeUInt8(0, 10); // Compression method (deflate)
  ihdrData.writeUInt8(0, 11); // Filter method (standard)
  ihdrData.writeUInt8(0, 12); // Interlace method (none)

  const ihdrChunk = createChunk('IHDR', ihdrData);

  // IDAT chunk (scanlines with filter byte 0)
  const rawScanlines = Buffer.alloc(height * (width * 4 + 1));
  for (let y = 0; y < height; y++) {
    const rowOffset = y * (width * 4 + 1);
    rawScanlines[rowOffset] = 0; // Filter byte: None
    rgbaBuffer.copy(rawScanlines, rowOffset + 1, y * width * 4, (y + 1) * width * 4);
  }

  const compressedData = zlib.deflateSync(rawScanlines);
  const idatChunk = createChunk('IDAT', compressedData);

  // IEND chunk
  const iendChunk = createChunk('IEND', Buffer.alloc(0));

  return Buffer.concat([signature, ihdrChunk, idatChunk, iendChunk]);
}

function createChunk(type, data) {
  const length = data.length;
  const chunk = Buffer.alloc(4 + 4 + length + 4);
  chunk.writeUInt32BE(length, 0);
  chunk.write(type, 4, 4, 'ascii');
  data.copy(chunk, 8);

  const crc = crc32(chunk.subarray(4, 8 + length));
  chunk.writeInt32BE(crc, 8 + length);
  return chunk;
}

// CRC32 table & calculation
const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  crcTable[n] = c;
}

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc = crcTable[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) | 0;
}

/**
 * Draws a beautiful rounded #0F6CBD icon with white math Pi (π) symbol.
 */
function renderPiIcon(size) {
  const buffer = Buffer.alloc(size * size * 4);

  // #0F6CBD = R: 15, G: 108, B: 189
  const bgR = 15;
  const bgG = 108;
  const bgB = 189;

  const cornerRadius = size * 0.22;
  const center = size / 2;

  // Normalized Pi stroke coordinates in [0, 1] relative to icon bounding box
  // Pi has:
  // 1. Horizontal top bar: curve extending from left to right with slight serif
  // 2. Left vertical leg: curving down and left
  // 3. Right vertical leg: curving down and hook right
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      const idx = (y * size + x) * 4;

      // Check rounded rectangle distance
      const dx = Math.max(0, Math.abs(x + 0.5 - center) - (center - cornerRadius));
      const dy = Math.max(0, Math.abs(y + 0.5 - center) - (center - cornerRadius));
      const distToCorner = Math.sqrt(dx * dx + dy * dy);

      let alpha = 0;
      if (distToCorner <= cornerRadius - 0.5) {
        alpha = 1;
      } else if (distToCorner <= cornerRadius + 0.5) {
        alpha = Math.max(0, Math.min(1, cornerRadius + 0.5 - distToCorner));
      }

      if (alpha <= 0) {
        // Transparent
        buffer[idx] = 0;
        buffer[idx + 1] = 0;
        buffer[idx + 2] = 0;
        buffer[idx + 3] = 0;
        continue;
      }

      // Inside rounded rect: check if on Pi glyph
      const nx = (x + 0.5) / size;
      const ny = (y + 0.5) / size;

      const isPi = computePiCoverage(nx, ny, size);

      if (isPi > 0) {
        // Blend white (#FFFFFF) over #0F6CBD
        const piAlpha = Math.min(1, isPi);
        const r = Math.round(bgR * (1 - piAlpha) + 255 * piAlpha);
        const g = Math.round(bgG * (1 - piAlpha) + 255 * piAlpha);
        const b = Math.round(bgB * (1 - piAlpha) + 255 * piAlpha);
        buffer[idx] = r;
        buffer[idx + 1] = g;
        buffer[idx + 2] = b;
        buffer[idx + 3] = Math.round(alpha * 255);
      } else {
        buffer[idx] = bgR;
        buffer[idx + 1] = bgG;
        buffer[idx + 2] = bgB;
        buffer[idx + 3] = Math.round(alpha * 255);
      }
    }
  }

  return createPng(size, heightOrWidth(size), buffer);
}

function heightOrWidth(s) {
  return s;
}

/**
 * High-quality procedural subpixel Pi (π) distance field calculation.
 */
function computePiCoverage(nx, ny, size) {
  // Pi bounding box: x in [0.20, 0.80], y in [0.24, 0.76]
  const topY = 0.32;
  const barHeight = Math.max(0.065, 1.8 / size);
  const leftX = 0.38;
  const rightX = 0.62;
  const legWidth = Math.max(0.065, 1.8 / size);
  const bottomY = 0.72;

  let minDist = 999;

  // 1. Top bar: segment from (0.22, 0.32) to (0.78, 0.32) with slight arch
  // y_bar(nx) = topY - 0.03 * sin((nx - 0.22) / 0.56 * PI)
  if (nx >= 0.20 && nx <= 0.80) {
    const arch = 0.02 * Math.sin(Math.PI * (nx - 0.20) / 0.60);
    const dTop = Math.abs(ny - (topY - arch));
    minDist = Math.min(minDist, dTop - barHeight / 2);
  } else if (nx < 0.20 && nx >= 0.17 && ny >= topY - 0.03 && ny <= topY + 0.06) {
    // Left serif curve
    const d = Math.hypot(nx - 0.20, ny - topY);
    minDist = Math.min(minDist, d - barHeight / 2);
  }

  // 2. Left leg: from (0.38, topY) down to (0.32, bottomY)
  if (ny >= topY - 0.01 && ny <= bottomY + 0.03) {
    const t = Math.max(0, Math.min(1, (ny - topY) / (bottomY - topY)));
    // curves slightly left towards bottom
    const legX = leftX - 0.06 * Math.pow(t, 1.5);
    const dLeft = Math.hypot(nx - legX, 0);
    minDist = Math.min(minDist, dLeft - legWidth / 2);
  }

  // 3. Right leg: from (0.62, topY) down to (0.64, bottomY) with a right hook
  if (ny >= topY - 0.01 && ny <= bottomY) {
    const t = Math.max(0, Math.min(1, (ny - topY) / (bottomY - topY)));
    const legX = rightX + 0.02 * t;
    const dRight = Math.hypot(nx - legX, 0);
    minDist = Math.min(minDist, dRight - legWidth / 2);
  }
  // Hook on right leg bottom: curves up to (0.75, 0.66)
  if (ny >= bottomY - 0.06 && ny <= bottomY + 0.06 && nx >= rightX && nx <= 0.76) {
    const hookCenterY = bottomY - 0.02;
    const hookCenterX = 0.65;
    const dHook = Math.abs(Math.hypot(nx - hookCenterX, ny - hookCenterY) - 0.07);
    if (nx > hookCenterX && ny >= hookCenterY - 0.04) {
      minDist = Math.min(minDist, dHook - legWidth / 2);
    }
  }

  // Subpixel antialiasing
  const pixelDist = minDist * size;
  if (pixelDist <= -0.5) return 1.0;
  if (pixelDist >= 0.5) return 0.0;
  return 0.5 - pixelDist;
}

// Generate all standard icon sizes
const sizes = [16, 32, 64, 80, 128];
const assetsDir = path.resolve('public/assets');

if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

sizes.forEach((sz) => {
  const png = renderPiIcon(sz);
  const filePath = path.join(assetsDir, `icon-${sz}.png`);
  fs.writeFileSync(filePath, png);
  console.log(`Generated: ${filePath} (${sz}x${sz})`);
});
