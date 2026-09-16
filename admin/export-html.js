/* ============================================================
   管理画面：HTML書き出し
   編集中の内容を埋め込んだ完成HTMLを作り、ZIPでダウンロードする。

   ・サイト用   … assets/ を参照する通常のページ。リポジトリにそのまま置ける
   ・単体HTML   … CSS・JS・掲載内容・画像をすべて埋め込んだ1ファイル完結版
   ============================================================ */
(function (global) {
  'use strict';

  var PAGE_DEFS = [
    { key: 'home', out: 'index.html', src: '../index.html', base: '', current: 'home', label: 'トップ' },
    { key: 'note', out: 'note/index.html', src: '../note/index.html', base: '../', current: 'note', label: '.note' },
    { key: 'seminar', out: 'seminar/index.html', src: '../seminar/index.html', base: '../', current: 'seminar', label: 'セミナー' },
  ];

  /* ---------- 取得ユーティリティ ---------- */

  function fetchText(url) {
    return fetch(url, { cache: 'no-store' }).then(function (r) {
      if (!r.ok) throw new Error(url + ' を読み込めません（HTTP ' + r.status + '）');
      return r.text();
    });
  }

  function fetchDataUri(url) {
    return fetch(url, { cache: 'no-store' })
      .then(function (r) {
        if (!r.ok) throw new Error(url + ' HTTP ' + r.status);
        return r.blob();
      })
      .then(function (blob) {
        return new Promise(function (resolve, reject) {
          var fr = new FileReader();
          fr.onload = function () { resolve(fr.result); };
          fr.onerror = reject;
          fr.readAsDataURL(blob);
        });
      });
  }

  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /* ---------- 画像の埋め込み ---------- */

  /** 掲載データを走査して assets/img を指す文字列を集める */
  function collectImages(node, out) {
    if (typeof node === 'string') {
      if (/^(?:\.\.\/)*assets\/img\//.test(node)) out[node] = true;
    } else if (Array.isArray(node)) {
      node.forEach(function (v) { collectImages(v, out); });
    } else if (node && typeof node === 'object') {
      Object.keys(node).forEach(function (k) { collectImages(node[k], out); });
    }
    return out;
  }

  /** 集めたパスを data URI に置き換えたデータの複製を返す */
  function embedImages(data, cache, onProgress) {
    var paths = Object.keys(collectImages(data, {}));
    if (!paths.length) return Promise.resolve(data);

    var i = 0;
    function step() {
      if (i >= paths.length) {
        // まとめて置換（JSON文字列上で行うのが確実）
        var json = JSON.stringify(data);
        paths.forEach(function (p) {
          if (cache[p]) json = json.split('"' + p + '"').join(JSON.stringify(cache[p]));
        });
        return Promise.resolve(JSON.parse(json));
      }
      var p = paths[i++];
      if (onProgress) onProgress(i, paths.length);
      if (cache[p]) return step();
      return fetchDataUri('../' + p.replace(/^(\.\.\/)+/, ''))
        .then(function (uri) { cache[p] = uri; })
        .catch(function () { /* 取得できない画像はパスのまま残す */ })
        .then(step);
    }
    return step();
  }

  /* ---------- 単体HTML同士のリンクをつなぎ直す ---------- */

  /**
   * 単体HTMLはZIP内で home.html / note.html / seminar.html と平らに並ぶため、
   * "note/" のようなサイト内リンクをファイル名に置き換える。
   * 今回書き出さないページは、公開サイトのURLへ逃がす。
   */
  function linkMap(site, pages) {
    var siteUrl = ((site.brand && site.brand.siteUrl) || 'https://www.wize.uno/').replace(/\/?$/, '/');
    var map = {};
    PAGE_DEFS.forEach(function (d) {
      var from = d.key === 'home' ? '' : d.key + '/';
      map[from] = !pages || pages.indexOf(d.key) > -1 ? d.key + '.html' : siteUrl + from;
    });
    return map;
  }

  /** 焼き付けたHTML側のリンクも同じ対応表で貼り替える（JS無効でも辿れるように） */
  function flattenLinks(html, map) {
    Object.keys(map).forEach(function (from) {
      html = html.split('href="' + from + '"').join('href="' + map[from] + '"');
    });
    return html;
  }

  /* ---------- テンプレートへの流し込み ---------- */

  /**
   * マーカーの「間」だけを差し替える。マーカーは残すので、
   * 書き出し済みのHTMLをもう一度テンプレートにしても壊れない。
   * 置換文字列の $ が特別扱いされないよう、正規表現ではなく位置で切り貼りする。
   */
  function splice(tpl, parts) {
    function between(html, name, content) {
      var start = '<!--wize:' + name + ':start-->';
      var end = '<!--wize:' + name + ':end-->';
      var i = html.indexOf(start);
      var j = html.indexOf(end, i);
      if (i < 0 || j < 0) throw new Error('テンプレートに wize:' + name + ' マーカーがありません');
      return html.slice(0, i + start.length) + content + html.slice(j);
    }
    var html = between(tpl, 'header', parts.header);
    html = between(html, 'footer', parts.footer);
    html = between(html, 'page', parts.page);
    return html
      .replace('<header data-chrome="header">', '<header data-chrome="header" class="site">')
      .replace('<footer data-chrome="footer">', '<footer data-chrome="footer" class="site">');
  }

  /* ---------- 1ページ分のHTMLを作る ---------- */

  function buildPage(def, opts) {
    var W = global.WIZE;
    var site = opts.getData('site');
    var standalone = opts.standalone;
    var base = standalone ? '' : def.base;

    // 単体HTMLでは、年の切り替えなど後から描画される分も含めて
    // 画像をすべて data URI にしてから組み立てる
    var prepare = standalone
      ? embedImages(opts.getData(def.key), opts.assetCache, function (i, n) {
          opts.onProgress(def.label + ' の画像を埋め込み中… ' + i + '/' + n);
        })
      : Promise.resolve(opts.getData(def.key));

    return Promise.all([fetchText(def.src), prepare]).then(function (pair) {
      var tpl = pair[0];
      var data = pair[1];

      var parts = W.withBase(base, function () {
        return {
          header: W.headerHTML(site, def.current),
          footer: W.footerHTML(site),
          page: W.pages[def.key](data),
        };
      });

      var html = splice(tpl, parts);

      if (data.meta) {
        if (data.meta.title) {
          html = html.replace(/<title>[\s\S]*?<\/title>/, '<title>' + esc(data.meta.title) + '</title>');
        }
        if (data.meta.description) {
          html = html.replace(/(<meta name="description" content=")[^"]*(">)/, '$1' + esc(data.meta.description) + '$2');
        }
      }

      if (!standalone) return html;

      /* --- 単体HTML：CSS・JS・データ・画像をすべて埋め込む --- */
      return Promise.all([
        fetchText('../assets/css/site.css'),
        fetchText('../assets/js/site.js'),
        fetchText('../assets/js/page-' + def.key + '.js'),
      ]).then(function (r) {
        var css = r[0], siteJs = r[1], pageJs = r[2];

        html = html.replace(/[ \t]*<link rel="stylesheet" href="[^"]*site\.css">\n?/, '<style>\n' + css + '\n</style>\n');

        var embedded = {};
        embedded.site = site;
        embedded[def.key] = data;
        // </script> がデータ中に現れてもスクリプトが切れないようにする
        var json = JSON.stringify(embedded).replace(/</g, '\\u003c');
        var map = linkMap(site, opts.pages);

        var block =
          '<script>window.WIZE_EMBEDDED = ' + json + ';\n' +
          'window.WIZE_LINKMAP = ' + JSON.stringify(map) + ';<\/script>\n' +
          '<script>\n' + siteJs + '\n<\/script>\n' +
          '<script>\n' + pageJs + '\n<\/script>\n';

        html = html.replace(
          /[ \t]*<script src="[^"]*site\.js"><\/script>\s*<script src="[^"]*page-[a-z]+\.js"><\/script>\n?/,
          block
        );

        return flattenLinks(html, map);
      });
    });
  }

  /* ---------- 最小限のZIP書き出し（無圧縮） ---------- */

  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    var c = 0xffffffff;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  function zip(files) {
    var enc = new TextEncoder();
    var chunks = [];
    var central = [];
    var offset = 0;

    function u32(v) { return [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff]; }
    function u16(v) { return [v & 0xff, (v >>> 8) & 0xff]; }

    files.forEach(function (f) {
      var name = enc.encode(f.name);
      var data = enc.encode(f.text);
      var crc = crc32(data);
      // 汎用ビット11を立てて、ファイル名がUTF-8であることを示す
      var header = [].concat(
        [0x50, 0x4b, 0x03, 0x04], u16(20), u16(0x0800), u16(0), u16(0), u16(0),
        u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0)
      );
      chunks.push(new Uint8Array(header), name, data);

      central.push({ name: name, crc: crc, size: data.length, offset: offset });
      offset += header.length + name.length + data.length;
    });

    var centralStart = offset;
    central.forEach(function (c) {
      var rec = [].concat(
        [0x50, 0x4b, 0x01, 0x02], u16(20), u16(20), u16(0x0800), u16(0), u16(0), u16(0),
        u32(c.crc), u32(c.size), u32(c.size),
        u16(c.name.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(c.offset)
      );
      chunks.push(new Uint8Array(rec), c.name);
      offset += rec.length + c.name.length;
    });

    var eocd = [].concat(
      [0x50, 0x4b, 0x05, 0x06], u16(0), u16(0),
      u16(files.length), u16(files.length),
      u32(offset - centralStart), u32(centralStart), u16(0)
    );
    chunks.push(new Uint8Array(eocd));

    return new Blob(chunks, { type: 'application/zip' });
  }

  /**
   * 自動ダウンロードを試みる。
   * 書き出しには数秒かかるため、この時点ではクリック操作（ユーザー操作）から
   * 離れており、ブラウザによっては自動ダウンロードが黙って無視される。
   * 成否は判定できないので、呼び出し側で必ず手動リンクも用意すること。
   */
  function tryAutoSave(url, filename, parent) {
    try {
      var a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.style.display = 'none';
      // モーダル表示中は dialog の外が inert になるため、開いていれば中に入れる
      (parent || document.querySelector('dialog[open]') || document.body).appendChild(a);
      a.click();
      setTimeout(function () { a.remove(); }, 2000);
      return true;
    } catch (e) {
      return false;
    }
  }

  /* ---------- 公開API ---------- */

  /**
   * @param {object} opts
   *   standalone {boolean}  1ファイル完結版にするか
   *   pages      {string[]} 書き出すページのキー。省略時は全部
   *   getData    {function} キーを渡すと編集中のデータを返す
   *   onProgress {function} 進捗メッセージ
   */
  function exportHTML(opts) {
    var defs = PAGE_DEFS.filter(function (d) {
      return !opts.pages || opts.pages.indexOf(d.key) > -1;
    });
    var assetCache = {};
    var out = [];

    var chain = Promise.resolve();
    defs.forEach(function (def) {
      chain = chain.then(function () {
        opts.onProgress(def.label + ' を書き出し中…');
        return buildPage(def, {
          standalone: opts.standalone,
          getData: opts.getData,
          pages: opts.pages,
          assetCache: assetCache,
          onProgress: opts.onProgress,
        }).then(function (html) {
          out.push({ name: opts.standalone ? def.key + '.html' : def.out, text: html });
        });
      });
    });

    return chain.then(function () {
      opts.onProgress('ZIPを作成中…');
      var stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      var filename = 'wize-html-' + (opts.standalone ? 'standalone-' : 'site-') + stamp + '.zip';
      var blob = zip(out);
      return { blob: blob, filename: filename, count: out.length, files: out };
    });
  }

  /** 1ページ分のHTMLを文字列で得る（動作確認・別用途向け） */
  function buildOne(key, opts) {
    var def = PAGE_DEFS.filter(function (d) { return d.key === key; })[0];
    if (!def) return Promise.reject(new Error('不明なページ: ' + key));
    return buildPage(def, {
      standalone: !!opts.standalone,
      getData: opts.getData,
      pages: opts.pages,
      assetCache: opts.assetCache || {},
      onProgress: opts.onProgress || function () {},
    });
  }

  global.WIZE_EXPORT = {
    exportHTML: exportHTML,
    buildOne: buildOne,
    zip: zip,
    tryAutoSave: tryAutoSave,
    PAGE_DEFS: PAGE_DEFS,
  };
})(window);
