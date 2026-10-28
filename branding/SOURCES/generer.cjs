const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
sharp.concurrency(1);
const root = path.resolve(__dirname, '..');

// Contour du logo choisi par l'utilisateur, repris de la référence sans épaississement.
const contour = fs.readFileSync(path.join(__dirname, 'korecto-contour.txt'), 'utf8').trim();
// Keep the approved outline unchanged; center its visible bounds in the tile.
// The source outline uses only absolute M/L coordinates and Z.
if (/[^MLZ\d.\s-]/i.test(contour)) throw Error('Commandes de contour non prises en charge.');
const coordinates = contour.match(/-?\d+(?:\.\d+)?/g).map(Number);
const xs = [], ys = [];
for (let i = 0; i < coordinates.length; i += 2) { xs.push(coordinates[i]); ys.push(coordinates[i + 1]); }
const dx = 64 - (Math.min(...xs) + Math.max(...xs)) / 2;
const dy = 64 - (Math.min(...ys) + Math.max(...ys)) / 2;
const mark = (color) => `<path transform="translate(${dx.toFixed(3)} ${dy.toFixed(3)})" fill="${color}" d="${contour}"/>`;
const svg = (variant) => `<svg xmlns="http://www.w3.org/2000/svg" width="1024" height="1024" viewBox="0 0 128 128">${variant === 'couleur' ? '<rect x="4" y="4" width="120" height="120" rx="28" fill="#087F83"/>' : ''}${mark(variant === 'sombre' ? '#18282B' : '#FFFFFF')}</svg>`;
const put = (file, data) => fs.writeFileSync(path.join(root, file), data);

function ico(frames) {
  const header = Buffer.alloc(6 + 16 * frames.length);
  header.writeUInt16LE(1, 2); header.writeUInt16LE(frames.length, 4);
  let offset = header.length;
  frames.forEach(({size, data}, i) => {
    const at = 6 + i * 16;
    header[at] = header[at + 1] = size === 256 ? 0 : size;
    header.writeUInt16LE(1, at + 4); header.writeUInt16LE(32, at + 6);
    header.writeUInt32LE(data.length, at + 8); header.writeUInt32LE(offset, at + 12);
    offset += data.length;
  });
  return Buffer.concat([header, ...frames.map(f => f.data)]);
}

async function render(variant, size) {
  return sharp(Buffer.from(svg(variant))).resize(size, size).png().toBuffer();
}

(async () => {
  for (const variant of ['couleur', 'clair', 'sombre']) put(`SOURCES/korecto-${variant}.svg`, svg(variant));
  const sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256];
  const frames = [];
  for (const size of sizes) frames.push({size, data: await render('couleur', size)});
  put('APPLICATION/Korecto.ico', ico(frames));
  for (const size of [16, 32, 48, 64, 128, 256, 512, 1024]) put(`APPLICATION/korecto-${size}.png`, await render('couleur', size));
  put('WEB/favicon.ico', ico(frames.filter(f => [16, 32, 48].includes(f.size))));
  for (const size of [16, 32]) put(`WEB/favicon-${size}.png`, await render('couleur', size));
  put('WEB/apple-touch-icon.png', await render('couleur', 180));
  for (const variant of ['couleur', 'clair', 'sombre']) {
    const tray = [];
    for (const size of [16, 20, 24, 32, 48]) {
      const data = await render(variant, size);
      put(`NOTIFICATION/korecto-${variant}-${size}.png`, data);
      tray.push({size, data});
    }
    put(`NOTIFICATION/korecto-${variant}.ico`, ico(tray));
  }

  // Planche de contrôle : les tailles 16 et 32 sont affichées sans agrandissement.
  const labels = `<svg xmlns="http://www.w3.org/2000/svg" width="1000" height="460"><rect width="1000" height="460" rx="24" fill="#F3F7F7"/><text x="40" y="53" fill="#18282B" font-family="Segoe UI,Arial" font-size="27" font-weight="600">Korecto</text><text x="40" y="82" fill="#5A6A6D" font-family="Segoe UI,Arial" font-size="15">K + coche de correction</text><text x="362" y="55" fill="#18282B" font-family="Segoe UI,Arial" font-size="17">Application / favicon</text><text x="362" y="92" fill="#5A6A6D" font-family="Segoe UI,Arial" font-size="14">16 px</text><text x="442" y="92" fill="#5A6A6D" font-family="Segoe UI,Arial" font-size="14">32 px</text><text x="542" y="92" fill="#5A6A6D" font-family="Segoe UI,Arial" font-size="14">64 px</text><text x="362" y="218" fill="#18282B" font-family="Segoe UI,Arial" font-size="17">Notification · fond clair</text><rect x="342" y="240" width="288" height="82" rx="12" fill="#FFFFFF"/><text x="680" y="218" fill="#18282B" font-family="Segoe UI,Arial" font-size="17">Notification · fond sombre</text><rect x="660" y="240" width="298" height="82" rx="12" fill="#18282B"/><text x="40" y="413" fill="#5A6A6D" font-family="Segoe UI,Arial" font-size="15">PNG transparents · ICO Windows multitailles · SVG modifiables</text></svg>`;
  const layers = [{input: await render('couleur', 256), left: 35, top: 112}];
  for (const [size, left, top] of [[16, 362, 112], [32, 442, 104], [64, 542, 98]]) layers.push({input: await render('couleur', size), left, top});
  for (const [variant, base] of [['sombre', 362], ['clair', 680]]) {
    for (const [size, x, y] of [[16, 0, 273], [32, 80, 265], [48, 170, 257]]) layers.push({input: await render(variant, size), left: base + x, top: y});
  }
  put('APERCU.png', await sharp(Buffer.from(labels)).composite(layers).png().toBuffer());

  let checked = 0;
  for (const folder of ['APPLICATION', 'WEB', 'NOTIFICATION']) {
    for (const file of fs.readdirSync(path.join(root, folder))) {
      const data = fs.readFileSync(path.join(root, folder, file));
      if (file.endsWith('.png')) {
        const info = await sharp(data).metadata();
        if (!info.hasAlpha || info.width !== info.height) throw Error(`PNG invalide : ${file}`);
        const stats = await sharp(data).stats();
        if (stats.channels[3].min !== 0 || stats.channels[3].max !== 255) throw Error(`Transparence invalide : ${file}`);
      } else if (file.endsWith('.ico')) {
        if (data.readUInt16LE(2) !== 1) throw Error(`ICO invalide : ${file}`);
        for (let i = 0; i < data.readUInt16LE(4); i++) {
          const at = 6 + 16 * i;
          const offset = data.readUInt32LE(at + 12), length = data.readUInt32LE(at + 8);
          const info = await sharp(data.subarray(offset, offset + length)).metadata();
          if (info.width !== (data[at] || 256)) throw Error(`Trame ICO invalide : ${file}`);
        }
      }
      checked++;
    }
  }
  // Update the embedded Windows assets from the same generated artwork.
  const assets = path.resolve(root, '..', 'src', 'assets');
  fs.copyFileSync(path.join(root, 'APPLICATION/Korecto.ico'), path.join(assets, 'Korecto.ico'));
  fs.copyFileSync(path.join(root, 'APPLICATION/korecto-128.png'), path.join(assets, 'Korecto-128.png'));
  fs.copyFileSync(path.join(root, 'NOTIFICATION/korecto-couleur.ico'), path.join(assets, 'Korecto-notification.ico'));
  console.log(`Centrage du K : translation (${dx.toFixed(3)}, ${dy.toFixed(3)}) dans le repère 128 × 128.`);
  console.log(`${checked} exports vérifiés. Aperçu : ${path.join(root, 'APERCU.png')}`);
})().catch(error => { console.error(error); process.exit(1); });
