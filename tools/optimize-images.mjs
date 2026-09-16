/**
 * assets/img/ 配下の取り込み画像を Web 用に縮小する（元サイズのままだと1枚数百KBあるため）。
 *
 *   node tools/optimize-images.mjs           … 幅 640px / WebP に変換
 *   node tools/optimize-images.mjs --width 900
 *
 * 変換後は .webp に置き換わるので content/*.json の参照も書き換える。
 */
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import sharp from 'sharp';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const IMG_ROOT = path.join(ROOT, 'assets', 'img');
const CONTENT = path.join(ROOT, 'content');

const widthArg = process.argv.indexOf('--width');
const WIDTH = widthArg > -1 ? Number(process.argv[widthArg + 1]) : 640;

async function walk(dir) {
  const out = [];
  for (const e of await fs.readdir(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) out.push(...(await walk(p)));
    else out.push(p);
  }
  return out;
}

async function main() {
  const files = (await walk(IMG_ROOT)).filter((f) => /\.(jpe?g|png)$/i.test(f));
  if (!files.length) console.log('変換対象の画像はありません。参照の点検だけ行います。');

  const renames = new Map();
  let before = 0;
  let after = 0;

  for (const f of files) {
    const stat = await fs.stat(f);
    before += stat.size;
    const dest = f.replace(/\.(jpe?g|png)$/i, '.webp');
    // パス指定だと Windows でファイルハンドルが残り unlink に失敗するのでバッファ経由で扱う
    const buf = await fs.readFile(f);
    await fs.writeFile(dest, await sharp(buf).resize({ width: WIDTH, withoutEnlargement: true }).webp({ quality: 76 }).toBuffer());
    after += (await fs.stat(dest)).size;
    await fs.unlink(f);
    renames.set(path.relative(ROOT, f).replace(/\\/g, '/'), path.relative(ROOT, dest).replace(/\\/g, '/'));
    process.stdout.write('.');
  }

  let repaired = 0;
  for (const name of await fs.readdir(CONTENT)) {
    if (!name.endsWith('.json')) continue;
    const p = path.join(CONTENT, name);
    let json = await fs.readFile(p, 'utf8');
    const before = json;

    for (const [from, to] of renames) {
      json = json.split(from).join(to);
    }

    /* 途中で中断した実行の取りこぼし対策：
       元ファイルが既に無く .webp だけが残っている参照も貼り替える。
       これをやらないと JSON が存在しない画像を指したままになる。 */
    const stale = json.match(/assets\/img\/[^"]+\.(?:jpe?g|png)/g) || [];
    for (const ref of new Set(stale)) {
      const webp = ref.replace(/\.(jpe?g|png)$/i, '.webp');
      if (existsSync(path.join(ROOT, ref))) continue; // 元ファイルが残っているなら触らない
      if (!existsSync(path.join(ROOT, webp))) continue;
      json = json.split(ref).join(webp);
      repaired++;
    }

    if (json !== before) await fs.writeFile(p, json);
  }
  if (repaired) console.log(`\n過去の実行で取りこぼしていた参照 ${repaired}件を貼り替えました。`);

  const mb = (n) => (n / 1024 / 1024).toFixed(1) + 'MB';
  console.log(`\n${files.length}枚を幅${WIDTH}px/WebPへ変換： ${mb(before)} → ${mb(after)}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
