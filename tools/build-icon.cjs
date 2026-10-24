// Optional asset-generation step: Node.js + sharp. Normal builds use the committed ICO.
// Usage: node tools/build-icon.cjs
const fs = require('node:fs');
const path = require('node:path');
const sharp = require('sharp');

async function main() {
  const assets = path.join(__dirname, '..', 'assets');
  const source = path.join(assets, 'DualSlide.svg');
  const sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256];
  const frames = [];
  for (const size of sizes) {
    // Rasterize the vector at each target resolution for clean small-size edges.
    frames.push(await sharp(source, { density: 384 }).resize(size, size).png().toBuffer());
  }
  const header = Buffer.alloc(6 + 16 * sizes.length);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(sizes.length, 4);
  let offset = header.length;
  frames.forEach((frame, i) => {
    const entry = 6 + 16 * i;
    header[entry] = sizes[i] === 256 ? 0 : sizes[i];
    header[entry + 1] = header[entry];
    header.writeUInt16LE(1, entry + 4);
    header.writeUInt16LE(32, entry + 6);
    header.writeUInt32LE(frame.length, entry + 8);
    header.writeUInt32LE(offset, entry + 12);
    offset += frame.length;
  });
  fs.writeFileSync(path.join(assets, 'DualSlide.ico'), Buffer.concat([header, ...frames]));
  console.log('Generated assets/DualSlide.ico from SVG (' + sizes.join(', ') + ' px).');
}
main().catch(error => { console.error(error); process.exitCode = 1; });
