/**
 * このサイトを手元で確認するための簡易サーバー。
 *
 *   node tools/serve.mjs
 *   node tools/serve.mjs --port 8080
 *
 * Node に最初から入っている機能だけで動くので、npx も npm も要りません。
 * （PowerShell の実行ポリシーで npx.ps1 が止められる環境でも動きます）
 *
 * 止めるときは Ctrl + C。
 */
import http from 'node:http';
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

const argv = process.argv.slice(2);
const portArg = argv.indexOf('--port');
const PORT = portArg > -1 ? Number(argv[portArg + 1]) : 4173;

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.webp': 'image/webp',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.ico': 'image/x-icon',
  '.txt': 'text/plain; charset=utf-8',
  '.xml': 'application/xml; charset=utf-8',
  '.woff2': 'font/woff2',
};

/** URL を、必ず公開フォルダの内側のパスへ解決する */
function resolvePath(urlPath) {
  let p;
  try {
    p = decodeURIComponent(urlPath.split('?')[0].split('#')[0]);
  } catch {
    return null;
  }

  const abs = path.resolve(ROOT, '.' + p);
  // 上位フォルダへの脱出を許さない
  if (abs !== ROOT && !abs.startsWith(ROOT + path.sep)) return null;

  if (p.endsWith('/')) return path.join(abs, 'index.html');
  if (existsSync(abs) && !path.extname(abs)) {
    const asDir = path.join(abs, 'index.html');
    if (existsSync(asDir)) return asDir;
  }
  return abs;
}

const server = http.createServer(async (req, res) => {
  const file = resolvePath(req.url || '/');
  const send = (code, body, type) => {
    res.writeHead(code, {
      'Content-Type': type || 'text/html; charset=utf-8',
      'Cache-Control': 'no-store',   // 編集がすぐ反映されるように
    });
    res.end(body);
  };

  if (!file) return send(400, 'Bad request', 'text/plain; charset=utf-8');

  try {
    const body = await fs.readFile(file);
    send(200, body, TYPES[path.extname(file).toLowerCase()] || 'application/octet-stream');
  } catch {
    // 見つからないときは本番（GitHub Pages）と同じく 404.html を返す
    try {
      send(404, await fs.readFile(path.join(ROOT, '404.html')));
    } catch {
      send(404, 'Not found', 'text/plain; charset=utf-8');
    }
  }
});

server.on('error', (e) => {
  if (e.code === 'EADDRINUSE') {
    console.error(`\nポート ${PORT} は既に使われています。`);
    console.error(`別の番号で起動してください：  node tools/serve.mjs --port 4174\n`);
  } else {
    console.error(e.message);
  }
  process.exit(1);
});

server.listen(PORT, () => {
  console.log('');
  console.log('  WIZE サイトをローカルで公開しました。Ctrl + C で停止します。');
  console.log('');
  console.log(`    サイト      http://localhost:${PORT}/`);
  console.log(`    .note       http://localhost:${PORT}/note/`);
  console.log(`    セミナー    http://localhost:${PORT}/seminar/`);
  console.log(`    管理画面    http://localhost:${PORT}/admin/`);
  console.log('');
});
