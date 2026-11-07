/**
 * ペライチに置いていた法務2ページを content/legal.json へ取り込む。
 *
 *   node tools/import-legal.mjs            … 取得して書き出す
 *   node tools/import-legal.mjs --dry      … 書き出さず内容だけ表示
 *
 * 法務文書なので、本文は一字も変えずにそのまま移す。
 * 見出し（h2）・段落（p）・箇条書き（ul/li、入れ子あり）の構造だけを読み取る。
 *
 * 移設が済めば不要。以降の修正は管理画面の「法務ページ」タブから行う。
 */
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'content', 'legal.json');
const DRY = process.argv.includes('--dry');
const FORCE = process.argv.includes('--force');

const SOURCES = {
  privacy: {
    url: 'https://peraichi.com/privacy_policies/view/xcfnirpkrd8',
    slug: 'privacy',
    description: '合同会社WIZEのプライバシーポリシー（個人情報の取扱いについて）です。',
  },
  terms: {
    url: 'https://peraichi.com/terms_for_customer/view/xcfnirpkrd8',
    slug: 'terms',
    description: '合同会社WIZEの決済に関する利用規約です。',
  },
};

const decode = (s) =>
  String(s)
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#0?39;/g, "'")
    .replace(/&#x27;/gi, "'");

/** タグを落として1行に。リンクは [文字](URL) として残す */
const text = (html) =>
  decode(
    String(html || '')
      .replace(/<a\b[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, (m, href, inner) => {
        const label = inner.replace(/<[^>]+>/g, '').trim();
        return label && label !== href ? `[${label}](${href})` : href;
      })
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, '')
  )
    .split('\n')
    .map((l) => l.replace(/[ \t　]+/g, ' ').trim())
    .filter(Boolean)
    .join('\n');

/**
 * 本文を、見出し・段落・箇条書きの並びに読み替える。
 * ul は入れ子になることがあるので、開いた ul を積み上げて追う。
 */
function parseBody(html) {
  const blocks = [];
  const stack = []; // 開いている ul のネスト
  const re = /<(\/?)(h2|h3|p|ul|li)\b[^>]*>/gi;

  let m;
  let cursor = 0;
  let pending = null; // いま中身を集めている li / p / h2

  const flush = (end) => {
    if (!pending) return;
    const inner = html.slice(pending.from, end);
    const value = text(inner.replace(/<ul\b[\s\S]*$/i, '')); // 入れ子の ul は別で拾う
    if (value) {
      if (pending.tag === 'li') {
        const target = stack[stack.length - 1];
        if (target) target.items.push(value);
      } else {
        const kind =
          pending.tag === 'h2' ? 'heading' : pending.tag === 'h3' ? 'subheading' : 'paragraph';
        blocks.push({ type: kind, text: value });
      }
    }
    pending = null;
  };

  while ((m = re.exec(html))) {
    const closing = m[1] === '/';
    const tag = m[2].toLowerCase();
    const at = m.index;

    if (!closing) {
      if (tag === 'ul') {
        flush(at);
        const list = { type: 'list', items: [] };
        const parent = stack[stack.length - 1];
        if (parent) parent.items.push(list); // 入れ子
        else blocks.push(list);
        stack.push(list);
      } else {
        flush(at);
        pending = { tag, from: re.lastIndex };
      }
    } else {
      if (tag === 'ul') {
        flush(at);
        stack.pop();
      } else {
        flush(at);
      }
    }
    cursor = re.lastIndex;
  }
  flush(html.length);

  return blocks;
}

async function grab(key) {
  const src = SOURCES[key];
  process.stdout.write(`取得 ${src.url} ... `);
  const res = await fetch(src.url, { headers: { 'user-agent': 'wize-site-migration/1.0' } });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const html = await res.text();
  console.log('OK');

  const sec = html.match(/<section class="paymentTos">([\s\S]*?)<\/section>/);
  if (!sec) throw new Error(`${key}: 本文のまとまりが見つかりません`);

  const title = text((sec[1].match(/<h1[^>]*>([\s\S]*?)<\/h1>/) || [, ''])[1]);
  const body = (sec[1].match(/<div class="paymentTos_body">([\s\S]*)$/) || [, sec[1]])[1];

  return {
    slug: src.slug,
    meta: { title: `${title} ｜ WIZE LLC`, description: src.description },
    title,
    blocks: parseBody(body),
  };
}

async function main() {
  if (existsSync(OUT) && !FORCE && !DRY) {
    console.error('content/legal.json は既にあります。作り直すなら --force を付けてください。');
    process.exit(1);
  }

  const out = {
    _comment:
      '法務ページの本文。ペライチに置いていたものをそのまま移しました。管理画面の「法務ページ」タブから編集できます。blocks の type は heading（見出し）/ paragraph（段落）/ list（箇条書き、入れ子可）です。',
    privacy: await grab('privacy'),
    terms: await grab('terms'),
  };

  const count = (d) => {
    let h = 0, sh = 0, p = 0, li = 0;
    const walk = (bs) => bs.forEach((b) => {
      if (b.type === 'heading') h++;
      else if (b.type === 'subheading') sh++;
      else if (b.type === 'paragraph') p++;
      else if (b.type === 'list') b.items.forEach((i) => (typeof i === 'string' ? li++ : walk([i])));
    });
    walk(d.blocks);
    return `条${h} / 号${sh} / 段落${p} / 箇条書き${li}`;
  };

  console.log(`\nprivacy : ${out.privacy.title} — ${count(out.privacy)}`);
  console.log(`terms   : ${out.terms.title} — ${count(out.terms)}`);

  if (DRY) {
    console.log('\n--- privacy の先頭5ブロック ---');
    console.log(JSON.stringify(out.privacy.blocks.slice(0, 5), null, 1));
    return;
  }

  await fs.writeFile(OUT, JSON.stringify(out, null, 2) + '\n');
  console.log(`\ncontent/legal.json を書き出しました。`);
}

main().catch((e) => {
  console.error(e.message || e);
  process.exit(1);
});
