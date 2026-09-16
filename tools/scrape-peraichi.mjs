/**
 * ペライチ（wize.uno）の3ページから掲載内容を抽出し content/*.json を生成する移行スクリプト。
 *
 *   node tools/scrape-peraichi.mjs            … 実サイトを取得して生成
 *   node tools/scrape-peraichi.mjs --dry      … 生成せず件数だけ表示
 *
 * ペライチは「ブロック」単位のHTMLを吐くため、data-structure 属性で種別を判定して拾っている。
 * 移行が終われば不要。以降の更新は admin/ の管理画面から行う。
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DRY = process.argv.includes('--dry');

const SOURCES = {
  home: 'https://wize.uno/',
  note: 'https://wize.uno/.note',
  seminar: 'https://wize.uno/seminar',
};

/* ---------- 最小限のHTMLユーティリティ ---------- */

const decode = (s) =>
  s
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#0?39;/g, "'")
    .replace(/&#x27;/gi, "'");

/** タグを落として1行テキストに */
const text = (html) =>
  decode(String(html || '').replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|li)>/gi, '\n').replace(/<[^>]+>/g, ''))
    .split('\n')
    .map((l) => l.replace(/\s+/g, ' ').trim())
    .filter(Boolean)
    .join('\n');

/** タグを落として改行も潰す */
const line = (html) => text(html).replace(/\n+/g, ' ').trim();

const firstMatch = (html, re) => (String(html).match(re) || [])[1] || '';

const absUrl = (u) => {
  if (!u) return '';
  if (u.startsWith('//')) return 'https:' + u;
  return u;
};

/** ペライチのセクションブロックへ分割 */
const splitBlocks = (html) =>
  html
    .split(/(?=<div class="pera1-section block")/)
    .slice(1)
    .map((raw) => ({
      raw,
      structure: firstMatch(raw, /data-structure="([^"]+)"/),
      id: firstMatch(raw, /id="(section-\d+)"/),
    }));

/** ブロック内の h2（見出し） */
const heading = (raw) => line(firstMatch(raw, /<h2[^>]*>([\s\S]*?)<\/h2>/));

/** ブロック内の最初の本文テキスト要素 */
const bodyText = (raw) => {
  const m = raw.match(/<div[^>]*data-structure="e-text"[^>]*>([\s\S]*?)<\/div>\s*<\/div>/);
  return m ? text(m[1]) : '';
};

/** ブロック内のリンク（ボタン/カード）をすべて */
const links = (raw) => {
  const out = [];
  const re = /<a\b([^>]*)>([\s\S]*?)<\/a>/gi;
  let m;
  while ((m = re.exec(raw))) {
    const href = decode(firstMatch(m[1], /href="([^"]*)"/));
    if (!href || href.startsWith('#') || href.startsWith('javascript')) continue;
    out.push({ href, label: line(m[2]) });
  }
  return out;
};

/** ブロック内の画像 */
const images = (raw) => {
  const out = [];
  const re = /<img\b[^>]*src="([^"]+)"[^>]*>/gi;
  let m;
  while ((m = re.exec(raw))) out.push(absUrl(decode(m[1])));
  return out;
};

/** 「note.comで閲覧 (200円)」のようなボタン文言から価格だけ取り出す */
const priceOf = (label) => {
  const m = String(label).match(/[（(]([^（()）]*?[0-9０-９][^（()）]*?)[)）]/);
  if (m) return m[1].trim();
  const m2 = String(label).match(/(お試し0円|無料|FREE)/i);
  return m2 ? m2[1] : '';
};

/** 本文末尾の「Write.2022/01/15」等を日付として切り出す */
const splitDate = (body) => {
  const lines = body.split('\n');
  let date = '';
  const kept = lines.filter((l) => {
    const m = l.match(/^\s*Write[.:：]?\s*(\d{4}[\/.年]\d{1,2}[\/.月]?\d{0,2}日?)/i);
    if (m) {
      date = m[1].replace(/[年月]/g, '/').replace(/日$/, '').replace(/\/$/, '');
      return false;
    }
    return true;
  });
  return { date, body: kept.join('\n') };
};

const slug = (s, i) =>
  (String(s).match(/[A-Za-z0-9]+/g) || []).join('-').toLowerCase().slice(0, 40) || `item-${i}`;

async function fetchPage(url) {
  const res = await fetch(url, { headers: { 'user-agent': 'wize-site-migration/1.0' } });
  if (!res.ok) throw new Error(`${url} -> HTTP ${res.status}`);
  return res.text();
}

/* ---------- .note ページ ---------- */

const ARTICLE = /^b-article-img-(left|right)/;
const CATEGORY_BLOCK = 'b-main-set-2--left';
const GROUP_BLOCKS = ['b-heading-double-bdr--updown', 'b-heading-lowerpage-title', 'b-heading-accent'];
const SUB_BLOCK = 'b-heading-has-subtitle';

function parseNote(html) {
  const blocks = splitBlocks(html);
  const page = { title: '.note', lead: '', caution: { title: '', body: '' } };
  const categories = [];
  const magazines = [];
  const subscriptions = [];

  let cat = null;
  let group = null;
  let sub = null;

  for (const b of blocks) {
    const h = heading(b.raw);

    if (b.structure === CATEGORY_BLOCK) {
      if (h === '.note' || !page.title) {
        // 先頭のヒーロー
        if (h === '.note') {
          page.lead = bodyText(b.raw);
          continue;
        }
      }
      if (h === 'WITHOVER News') {
        cat = null;
        group = null;
        sub = null;
        continue;
      }
      cat = { id: h.replace(/^\./, '') || `cat-${categories.length}`, label: h, lead: bodyText(b.raw), groups: [] };
      categories.push(cat);
      group = null;
      sub = null;
      continue;
    }

    if (b.structure === 'b-sentence-caution') {
      page.caution = { title: h, body: bodyText(b.raw) };
      continue;
    }

    if (GROUP_BLOCKS.includes(b.structure) && cat) {
      group = { title: h, subtitle: '', items: [] };
      cat.groups.push(group);
      sub = null;
      continue;
    }

    if (b.structure === SUB_BLOCK && cat) {
      sub = { title: h, note: bodyText(b.raw) };
      continue;
    }

    if (ARTICLE.test(b.structure)) {
      const raw = b.raw;
      const linkList = links(raw).filter((l) => /note\.com/.test(l.href));
      const primary = linkList[0] || {};
      const { date, body } = splitDate(bodyText(raw));
      const item = {
        id: slug(primary.href || h, cat ? cat.groups.length : 0),
        title: h,
        summary: body,
        date,
        price: priceOf(primary.label),
        href: primary.href || '',
        image: images(raw)[0] || '',
        era: sub ? sub.title : '',
        eraNote: sub ? sub.note : '',
        extraLinks: linkList.slice(1).map((l) => ({ label: l.label, href: l.href, price: priceOf(l.label) })),
      };
      if (!cat) continue;
      if (!group) {
        group = { title: '', subtitle: '', items: [] };
        cat.groups.push(group);
      }
      group.items.push(item);
      continue;
    }

    // 年別バックナンバー（4カラムカード）
    if (b.structure === 'b-cards--4col') {
      const cards = parseCards(b.raw);
      if (/マガジン|定期購読|メンバーシップ/.test(h)) {
        subscriptions.push({ title: h, items: cards });
      } else {
        magazines.push({ id: slug(h, magazines.length), title: h, items: cards });
      }
      continue;
    }
  }

  return { page, categories, magazines, subscriptions };
}

/** カードブロック（3col/4col）からカードを取り出す */
function parseCards(raw) {
  const cards = raw.split(/(?=<div data-structure="m-card")/).slice(1);
  const out = [];
  cards.forEach((col, i) => {
    const headingHtml =
      firstMatch(col, /<h3[^>]*>([\s\S]*?)<\/h3>/) ||
      firstMatch(col, /<div[^>]*data-structure="e-heading"[^>]*>([\s\S]*?)<\/div>/);
    const lines = text(headingHtml).split('\n').filter(Boolean);
    const l = links(col)[0];
    const body = bodyText(col);
    if (!lines.length && !l) return;
    out.push({
      id: slug(l?.href || lines[0], i),
      label: lines[0] || '',
      title: lines.slice(1).join(' ') || lines[0] || '',
      summary: body,
      price: priceOf(l?.label),
      href: l?.href || '',
      image: images(col)[0] || '',
    });
  });
  return out;
}

/* ---------- seminar ページ ---------- */

function parseSeminar(html) {
  const blocks = splitBlocks(html);
  const sections = [];
  let current = null;

  for (const b of blocks) {
    const h = heading(b.raw);
    const body = bodyText(b.raw);
    const ls = links(b.raw).filter((l) => /udemy\.com|note\.com|wize\.uno\/reserve/.test(l.href));

    if (/^b-heading|^b-main-set|^b-premium-heading/.test(b.structure)) {
      current = { id: slug(h, sections.length), title: h, lead: body, items: [] };
      sections.push(current);
      continue;
    }

    if (!ls.length && !body) continue;
    if (!current) {
      current = { id: `section-${sections.length}`, title: '', lead: '', items: [] };
      sections.push(current);
    }
    if (ls.length) {
      ls.forEach((l, i) => {
        current.items.push({
          id: slug(l.href, i),
          title: h || l.label.split('\n')[0] || '',
          summary: body,
          price: priceOf(l.label),
          ctaLabel: l.label.replace(/\n/g, ' '),
          href: l.href,
          image: images(b.raw)[0] || '',
        });
      });
    } else if (h || body) {
      current.items.push({ id: slug(h, current.items.length), title: h, summary: body, price: '', href: '', image: images(b.raw)[0] || '' });
    }
  }
  return { sections: sections.filter((s) => s.title || s.items.length) };
}

/* ---------- トップページ ---------- */

function parseHome(html) {
  const blocks = splitBlocks(html);
  return {
    blocks: blocks
      .map((b) => ({
        structure: b.structure,
        heading: heading(b.raw),
        body: bodyText(b.raw),
        links: links(b.raw),
        images: images(b.raw),
      }))
      .filter((b) => b.heading || b.body || b.links.length),
  };
}

/* ---------- 実行 ---------- */

const outDir = path.join(ROOT, 'content');
const rawDir = path.join(ROOT, 'tools', '.cache');

async function main() {
  await fs.mkdir(rawDir, { recursive: true });
  const pages = {};
  for (const [key, url] of Object.entries(SOURCES)) {
    process.stdout.write(`fetch ${url} ... `);
    const html = await fetchPage(url);
    await fs.writeFile(path.join(rawDir, `${key}.html`), html);
    pages[key] = html;
    console.log(`${html.length.toLocaleString()} bytes`);
  }

  const note = parseNote(pages.note);
  const seminar = parseSeminar(pages.seminar);
  const home = parseHome(pages.home);

  const articleCount = note.categories.reduce((n, c) => n + c.groups.reduce((m, g) => m + g.items.length, 0), 0);
  console.log(`\n.note      : ${note.categories.length} カテゴリ / ${articleCount} 記事 / ${note.magazines.length} 年分のバックナンバー`);
  console.log(`seminar    : ${seminar.sections.length} セクション / ${seminar.sections.reduce((n, s) => n + s.items.length, 0)} 項目`);
  console.log(`home       : ${home.blocks.length} ブロック`);

  if (DRY) return;

  await fs.mkdir(outDir, { recursive: true });
  await fs.writeFile(path.join(rawDir, 'note.raw.json'), JSON.stringify(note, null, 2));
  await fs.writeFile(path.join(rawDir, 'seminar.raw.json'), JSON.stringify(seminar, null, 2));
  await fs.writeFile(path.join(rawDir, 'home.raw.json'), JSON.stringify(home, null, 2));
  console.log(`\n生の抽出結果を tools/.cache/*.raw.json に保存しました。`);
  console.log(`content/*.json への反映は tools/build-content.mjs で行います。`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
