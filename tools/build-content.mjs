/**
 * tools/.cache/*.raw.json → content/note.json を生成する。
 * ペライチCDN上の画像も assets/img/ へ取り込み、参照をローカルパスに書き換える。
 *
 *   node tools/build-content.mjs              … 画像も取り込む
 *   node tools/build-content.mjs --no-images  … JSONだけ（画像はCDN参照のまま）
 *   node tools/build-content.mjs --force      … 既存の content/note.json を上書き
 *
 * ※ content/note.json は管理画面での編集対象。上書きすると編集内容が失われるため
 *   既存ファイルがある場合は --force が必要。
 */
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const CACHE = path.join(ROOT, 'tools', '.cache');
const IMG_DIR = path.join(ROOT, 'assets', 'img', 'note');
const OUT = path.join(ROOT, 'content', 'note.json');

const WITH_IMAGES = !process.argv.includes('--no-images');
const FORCE = process.argv.includes('--force');

const imageCache = new Map();

async function localizeImage(url) {
  if (!url || !WITH_IMAGES) return url;
  if (!/^https?:\/\//.test(url)) return url;
  if (imageCache.has(url)) return imageCache.get(url);

  const ext = (url.match(/\.(jpe?g|png|gif|webp|svg)(?:\?|$)/i) || [, 'jpg'])[1].toLowerCase();
  const name = crypto.createHash('sha1').update(url).digest('hex').slice(0, 16) + '.' + ext;
  const dest = path.join(IMG_DIR, name);
  const rel = `assets/img/note/${name}`;

  if (!existsSync(dest)) {
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      await fs.writeFile(dest, Buffer.from(await res.arrayBuffer()));
      process.stdout.write('.');
    } catch (e) {
      console.warn(`\n  画像取得失敗 ${url}: ${e.message}`);
      imageCache.set(url, url);
      return url;
    }
  }
  imageCache.set(url, rel);
  return rel;
}

/** 「お試し0円 / 200円 / 前半0円+後半500円」→ 表示用ラベルと無料判定 */
function normalizePrice(price) {
  const p = String(price || '').trim();
  const free = /^(お試し)?0円$|^無料$|^FREE$|^基本0円$|^本文0円$/i.test(p);
  return { price: p, free };
}

async function main() {
  if (existsSync(OUT) && !FORCE) {
    console.error(`content/note.json は既に存在します。上書きするなら --force を付けてください。`);
    process.exit(1);
  }

  const raw = JSON.parse(await fs.readFile(path.join(CACHE, 'note.raw.json'), 'utf8'));
  await fs.mkdir(IMG_DIR, { recursive: true });
  if (WITH_IMAGES) process.stdout.write('画像を取り込み中 ');

  const categories = [];
  for (const cat of raw.categories) {
    const groups = [];
    for (const g of cat.groups) {
      const items = [];
      for (const it of g.items) {
        const { price, free } = normalizePrice(it.price);
        items.push({
          id: it.id,
          title: it.title,
          summary: it.summary,
          date: it.date,
          era: it.era,
          eraNote: it.eraNote,
          price,
          free,
          href: it.href,
          image: await localizeImage(it.image),
          published: true,
        });
      }
      groups.push({ title: g.title, subtitle: g.subtitle, items });
    }
    categories.push({ id: cat.id, label: cat.label, lead: cat.lead, groups });
  }

  const magazines = [];
  for (const m of raw.magazines) {
    const items = [];
    for (const it of m.items) {
      const { price, free } = normalizePrice(it.price);
      items.push({
        id: it.id,
        label: it.label,
        title: it.title,
        price,
        free,
        href: it.href,
        image: await localizeImage(it.image),
        published: true,
      });
    }
    magazines.push({ id: m.id, title: m.title, items });
  }

  const subscriptions = [];
  for (const s of raw.subscriptions) {
    const items = [];
    for (const it of s.items) {
      const { price, free } = normalizePrice(it.price);
      items.push({
        id: it.id,
        label: it.label,
        title: it.title,
        summary: it.summary,
        price,
        free,
        href: it.href,
        image: await localizeImage(it.image),
        published: true,
      });
    }
    subscriptions.push({ title: s.title, items });
  }

  const out = {
    _comment: '管理画面（/admin/）から編集できます。手で書き換えても構いません。',
    meta: {
      title: '.note ｜ バックナンバー',
      description:
        'note.com で公開している記事を、時系列別・分野別（History / Finance / Insurance）にまとめたバックナンバー一覧です。',
    },
    hero: {
      eyebrow: 'BACK NUMBER — バックナンバー',
      title: '.note',
      lead: raw.page.lead,
    },
    caution: raw.page.caution,
    priceLegend: [
      { label: 'お試し0円', note: '期間限定を含む無料公開のお試し記事です。' },
      { label: '100〜900円', note: '前半は無料、中盤から後半にかけて有料の記事です。' },
      { label: '1,000円〜', note: '専門的な内容、または有料セミナー・有料相談でお伝えする内容を文章化したものです。' },
    ],
    categories,
    magazines,
    subscriptions,
  };

  await fs.writeFile(OUT, JSON.stringify(out, null, 2) + '\n');
  const n = categories.reduce((a, c) => a + c.groups.reduce((b, g) => b + g.items.length, 0), 0);
  const nm = magazines.reduce((a, m) => a + m.items.length, 0);
  console.log(`\ncontent/note.json を書き出しました（記事 ${n} / バックナンバー ${nm} / 画像 ${imageCache.size}）`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
