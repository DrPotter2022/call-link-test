/**
 * 内容を埋め込んだ完成HTMLを、ブラウザを使わずファイルへ書き出す。
 * 管理画面の「HTMLを書き出す」と同じ結果を、ダウンロードを介さずに得られる。
 *
 *   node tools/export-html.mjs                  … dist/site/ へ（assets を参照するページ）
 *   node tools/export-html.mjs --standalone     … dist/standalone/ へ（1ファイル完結）
 *   node tools/export-html.mjs --inplace        … リポジトリ内のHTMLを直接更新
 *   node tools/export-html.mjs --pages home,note
 *
 * ページの組み立てには、サイト本体と同じ assets/js/*.js をそのまま使う。
 * 書き出し結果とサイトの表示がずれないようにするため。
 */
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

const argv = process.argv.slice(2);
const has = (f) => argv.includes(f);
const valueOf = (f) => {
  const i = argv.indexOf(f);
  return i > -1 ? argv[i + 1] : null;
};

const STANDALONE = has('--standalone');
const INPLACE = has('--inplace');
const ONLY = (valueOf('--pages') || '').split(',').map((s) => s.trim()).filter(Boolean);

const PAGE_DEFS = [
  { key: 'home', out: 'index.html', src: 'index.html', base: '', current: 'home', label: 'トップ' },
  { key: 'note', out: 'note/index.html', src: 'note/index.html', base: '../', current: 'note', label: '.note' },
  { key: 'seminar', out: 'seminar/index.html', src: 'seminar/index.html', base: '../', current: 'seminar', label: 'セミナー' },
];

const read = (rel) => fs.readFile(path.join(ROOT, rel), 'utf8');
const readJSON = async (rel) => JSON.parse(await read(rel));

const esc = (s) =>
  String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const MIME = {
  '.webp': 'image/webp', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
  '.gif': 'image/gif', '.svg': 'image/svg+xml',
};

/* ---------- ブラウザ用スクリプトを Node で動かすための最小の器 ---------- */

async function loadRenderers() {
  const sandbox = {
    console,
    setTimeout,
    clearTimeout,
    localStorage: { getItem: () => null, setItem() {}, removeItem() {} },
    fetch: async () => { throw new Error('書き出し中はネットワークを使いません'); },
    matchMedia: () => ({ matches: false }),
    document: {
      currentScript: null,
      getElementById: () => null,
      querySelector: () => null,
      querySelectorAll: () => [],
      addEventListener() {},
      documentElement: { classList: { add() {} } },
      createElement: () => ({ setAttribute() {}, appendChild() {} }),
    },
    IntersectionObserver: function () { this.observe = () => {}; this.unobserve = () => {}; },
  };
  sandbox.window = sandbox;
  sandbox.global = sandbox;
  vm.createContext(sandbox);

  for (const f of ['assets/js/site.js', 'assets/js/page-home.js', 'assets/js/page-note.js', 'assets/js/page-seminar.js']) {
    vm.runInContext(await read(f), sandbox, { filename: f });
  }
  if (!sandbox.WIZE || !sandbox.WIZE.pages.home) throw new Error('レンダラを読み込めませんでした');
  return sandbox.WIZE;
}

/* ---------- 画像の埋め込み ---------- */

const dataUriCache = new Map();

async function toDataUri(rel) {
  const clean = rel.replace(/^(\.\.\/)+/, '');
  if (dataUriCache.has(clean)) return dataUriCache.get(clean);
  const abs = path.join(ROOT, clean);
  if (!existsSync(abs)) {
    console.warn(`  画像が見つかりません: ${clean}`);
    dataUriCache.set(clean, null);
    return null;
  }
  const buf = await fs.readFile(abs);
  const mime = MIME[path.extname(abs).toLowerCase()] || 'application/octet-stream';
  const uri = `data:${mime};base64,${buf.toString('base64')}`;
  dataUriCache.set(clean, uri);
  return uri;
}

/** 掲載データ内の assets/img 参照をすべて data URI に置き換えた複製を返す */
async function embedImages(data) {
  let json = JSON.stringify(data);
  const refs = new Set(json.match(/(?:\.\.\/)*assets\/img\/[^"\\]+/g) || []);
  for (const ref of refs) {
    const uri = await toDataUri(ref);
    if (uri) json = json.split(`"${ref}"`).join(JSON.stringify(uri));
  }
  return { data: JSON.parse(json), count: refs.size };
}

/* ---------- 1ページ分の組み立て ---------- */

/** 単体HTML同士をつなぐ対応表（"note/" → "note.html"） */
function linkMap(site, pages) {
  const siteUrl = ((site.brand && site.brand.siteUrl) || 'https://www.wize.uno/').replace(/\/?$/, '/');
  const map = {};
  for (const d of PAGE_DEFS) {
    const from = d.key === 'home' ? '' : d.key + '/';
    map[from] = pages.includes(d.key) ? d.key + '.html' : siteUrl + from;
  }
  return map;
}

/** 焼き付けたHTML側のリンクも同じ対応表で貼り替える（JS無効でも辿れるように） */
function flattenLinks(html, map) {
  for (const [from, to] of Object.entries(map)) {
    html = html.split(`href="${from}"`).join(`href="${to}"`);
  }
  return html;
}

/**
 * テンプレート内のマーカーの「間」を差し替える。
 * マーカーを残すので、書き出し済みのHTMLに対して何度でも実行できる（--inplace 用）。
 */
function splice(tpl, parts) {
  // 置換文字列の $ が特別扱いされないよう、正規表現ではなく位置で切り貼りする
  const between = (html, name, content) => {
    const start = `<!--wize:${name}:start-->`;
    const end = `<!--wize:${name}:end-->`;
    const i = html.indexOf(start);
    const j = html.indexOf(end, i);
    if (i < 0 || j < 0) throw new Error(`テンプレートに wize:${name} マーカーがありません`);
    return html.slice(0, i + start.length) + content + html.slice(j);
  };
  let html = between(tpl, 'header', parts.header);
  html = between(html, 'footer', parts.footer);
  html = between(html, 'page', parts.page);
  // 中身が入ったので、外側の要素にもスタイル用のクラスを付けておく
  return html
    .replace('<header data-chrome="header">', '<header data-chrome="header" class="site">')
    .replace('<footer data-chrome="footer">', '<footer data-chrome="footer" class="site">');
}

async function buildPage(W, def, site, pages) {
  const tpl = await read(def.src);
  let data = await readJSON(`content/${def.key}.json`);
  let embedded = 0;

  if (STANDALONE) {
    const r = await embedImages(data);
    data = r.data;
    embedded = r.count;
  }

  const base = STANDALONE ? '' : def.base;
  const parts = W.withBase(base, () => ({
    header: W.headerHTML(site, def.current),
    footer: W.footerHTML(site),
    page: W.pages[def.key](data),
  }));

  let html = splice(tpl, parts);

  if (data.meta) {
    if (data.meta.title) html = html.replace(/<title>[\s\S]*?<\/title>/, `<title>${esc(data.meta.title)}</title>`);
    if (data.meta.description) {
      html = html.replace(/(<meta name="description" content=")[^"]*(">)/, `$1${esc(data.meta.description)}$2`);
    }
  }

  if (STANDALONE) {
    const [css, siteJs, pageJs] = await Promise.all([
      read('assets/css/site.css'),
      read('assets/js/site.js'),
      read(`assets/js/page-${def.key}.js`),
    ]);

    html = html.replace(/[ \t]*<link rel="stylesheet" href="[^"]*site\.css">\n?/, `<style>\n${css}\n</style>\n`);

    const map = linkMap(site, pages);
    const payload = JSON.stringify({ site, [def.key]: data }).replace(/</g, '\\u003c');
    const block =
      `<script>window.WIZE_EMBEDDED = ${payload};\n` +
      `window.WIZE_LINKMAP = ${JSON.stringify(map)};<\/script>\n` +
      `<script>\n${siteJs}\n<\/script>\n` +
      `<script>\n${pageJs}\n<\/script>\n`;

    html = html.replace(
      /[ \t]*<script src="[^"]*site\.js"><\/script>\s*<script src="[^"]*page-[a-z]+\.js"><\/script>\n?/,
      block
    );
    html = flattenLinks(html, map);
  }

  return { html, embedded };
}

/* ---------- 実行 ---------- */

async function main() {
  const W = await loadRenderers();
  const site = await readJSON('content/site.json');

  const defs = PAGE_DEFS.filter((d) => !ONLY.length || ONLY.includes(d.key));
  if (!defs.length) throw new Error(`--pages の指定が不正です（使えるのは ${PAGE_DEFS.map((d) => d.key).join(', ')}）`);
  const pages = defs.map((d) => d.key);

  const outDir = INPLACE ? ROOT : path.join(ROOT, 'dist', STANDALONE ? 'standalone' : 'site');
  if (INPLACE && STANDALONE) throw new Error('--inplace と --standalone は同時に使えません');

  console.log(`書き出し先: ${path.relative(ROOT, outDir) || '.'}（${STANDALONE ? '単体HTML' : 'サイト用'}）\n`);

  for (const def of defs) {
    const { html, embedded } = await buildPage(W, def, site, pages);
    const rel = STANDALONE ? `${def.key}.html` : def.out;
    const dest = path.join(outDir, rel);
    await fs.mkdir(path.dirname(dest), { recursive: true });
    await fs.writeFile(dest, html);
    const kb = (Buffer.byteLength(html) / 1024).toFixed(0);
    console.log(`  ${def.label.padEnd(8, '　')} → ${rel}  ${kb}KB${embedded ? `（画像${embedded}枚を埋め込み）` : ''}`);
  }

  console.log(`\n${defs.length}ページを書き出しました。`);
  if (!INPLACE) console.log(`そのまま公開するなら： node tools/export-html.mjs --inplace`);
}

main().catch((e) => {
  console.error('\n' + (e.stack || e.message));
  process.exit(1);
});
