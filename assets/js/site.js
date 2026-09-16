/* ============================================================
   WIZE LLC — 共通スクリプト
   ・content/*.json を読み込んでページを組み立てる
   ・管理画面（/admin/）の下書きが localStorage にあればそれを優先して表示する
   ============================================================ */
(function (global) {
  'use strict';

  /* サイトのルートを自分の script src から求める（/note/ 等の下層でも動くように） */
  var BASE = (function () {
    var s = document.currentScript && document.currentScript.src;
    if (!s) return './';
    return s.replace(/assets\/js\/site\.js.*$/, '');
  })();

  var DRAFT_PREFIX = 'wize_draft_';

  /* 静的書き出しのときだけ、書き出す先の階層に合わせて基準パスを差し替える */
  var baseOverride = null;

  function url(p) {
    var s = String(p == null ? '' : p);
    // 絶対URL・data: は基準パスを足さずそのまま使う（単体HTMLの埋め込み画像など）
    if (/^(https?:|data:|blob:|\/\/)/i.test(s)) return s;
    return (baseOverride === null ? BASE : baseOverride) + s.replace(/^\//, '');
  }

  /** base を一時的に差し替えて fn を実行する（管理画面のHTML書き出し用） */
  function withBase(base, fn) {
    var prev = baseOverride;
    baseOverride = base;
    try {
      return fn();
    } finally {
      baseOverride = prev;
    }
  }

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  /** 改行を <br> に。テキストはエスケープ済みで扱う */
  function nl2br(s) {
    return esc(s).replace(/\n/g, '<br>');
  }

  /**
   * 見出し等で最小限の装飾だけ許可する。
   * <em>（朱のマーカー）と <strong> のみ復活させ、それ以外のタグは無効化したまま。
   */
  function rich(s) {
    return nl2br(s)
      .replace(/&lt;(\/?)(em|strong)&gt;/g, '<$1$2>');
  }

  function readDraft(name) {
    try {
      var raw = localStorage.getItem(DRAFT_PREFIX + name);
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null;
    }
  }

  var loaded = {};

  /** content/<name>.json を読む。管理画面の下書きがあればそちらを優先 */
  function load(name) {
    if (loaded[name]) return loaded[name];

    // 単体HTMLとして書き出された場合、掲載内容はページ内に埋め込まれている
    if (global.WIZE_EMBEDDED && global.WIZE_EMBEDDED[name]) {
      loaded[name] = Promise.resolve(global.WIZE_EMBEDDED[name]);
      return loaded[name];
    }

    var draft = readDraft(name);
    if (draft) {
      markDraft();
      loaded[name] = Promise.resolve(draft);
      return loaded[name];
    }
    loaded[name] = fetch(url('content/' + name + '.json'), { cache: 'no-store' })
      .then(function (r) {
        if (!r.ok) throw new Error(name + '.json: HTTP ' + r.status);
        return r.json();
      })
      .catch(function (e) {
        console.error('コンテンツを読み込めませんでした', e);
        return null;
      });
    return loaded[name];
  }

  var draftMarked = false;
  function markDraft() {
    if (draftMarked) return;
    draftMarked = true;
    document.addEventListener('DOMContentLoaded', function () {
      var a = document.createElement('a');
      a.id = 'draft-badge';
      a.href = url('admin/');
      a.textContent = '下書きを表示中 — 管理画面へ';
      document.body.appendChild(a);
    });
  }

  /* ---------- 共通の外枠（ヘッダー・フッター） ---------- */

  /** ヘッダーの中身をHTML文字列で作る（静的書き出しからも使う） */
  function headerHTML(site, current) {
    if (!site) return '';
    return (
        '<div class="wrap nav">' +
        '<a class="logo" href="' + url('') + '">' +
        '<span class="logo-jp">' + esc(site.brand.nameJa) + '</span>' +
        '<span class="logo-en">' + esc(site.brand.nameEn) + '</span>' +
        '</a>' +
        '<button class="nav-toggle" type="button" aria-expanded="false" aria-controls="nav-links">MENU</button>' +
        '<ul class="nav-links" id="nav-links">' +
        (site.nav || [])
          .map(function (n) {
            var here = n.id && n.id === current;
            return (
              '<li><a href="' + esc(resolveHref(n.href)) + '"' + (here ? ' aria-current="page"' : '') + '>' + esc(n.label) + '</a></li>'
            );
          })
          .join('') +
        (site.navCta && site.navCta.label
          ? '<li><a class="nav-cta" href="' + esc(resolveHref(site.navCta.href)) + '">' + esc(site.navCta.label) + '</a></li>'
          : '') +
        '</ul>' +
        '</div>'
    );
  }

  /** フッターの中身をHTML文字列で作る（静的書き出しからも使う） */
  function footerHTML(site) {
    if (!site) return '';
    var f = site.footer || {};
    var c = f.company || {};
    return (
        '<div class="wrap">' +
        '<div class="f-grid">' +
        '<a class="logo" href="' + url('') + '">' +
        '<span class="logo-jp">' + esc(site.brand.nameJa) + '</span>' +
        '<span class="logo-en">' + esc(site.brand.nameEn) + '</span>' +
        '</a>' +
        '<nav>' +
        (f.links || [])
          .map(function (l) {
            var ext = /^https?:/.test(l.href) ? ' target="_blank" rel="noopener"' : '';
            return '<a href="' + esc(resolveHref(l.href)) + '"' + ext + '>' + esc(l.label) + '</a>';
          })
          .join('') +
        '</nav>' +
        '</div>' +
        /* 事業者情報（OPERATOR）は既定では出さない。決済はUdemy側で完結し、
           特定商取引法に基づく表記もUdemy側にあるため。
           出す必要が生じたら site.json の footer.company に値を入れる。 */
        (c && c.name
          ? '<div class="operator">' +
            '<p class="op-label">' + esc(f.companyLabel || 'OPERATOR — 運営') + '</p>' +
            '<dl>' +
            [
              ['事業者名', c.name],
              ['代表者', c.representative],
              ['所在地', c.address],
              ['お問い合わせ', [c.email, c.tel].filter(Boolean).join('　')],
              ['営業時間', c.hours],
            ]
              .filter(function (r) {
                return r[1];
              })
              .map(function (r) {
                return '<div class="op-row"><dt>' + esc(r[0]) + '</dt><dd>' + esc(r[1]) + '</dd></div>';
              })
              .join('') +
            '</dl>' +
            '</div>'
          : '') +
        '<div class="legal">' +
        (f.legal ? '<p>' + nl2br(f.legal) + '</p>' : '') +
        '<p class="f-bottom">' +
        '<span>' + esc(f.copyright || '') + '</span>' +
        /* 管理画面への入口。site.json で footer.adminLink を false にすると消えます。
           URLを隠しても守りにはならない（守っているのはパスワードとGitHubトークン）ので、
           迷わず入れることを優先しています。 */
        (f.adminLink === false ? '' : '<a class="f-admin" href="' + url('admin/') + '">管理</a>') +
        '</p>' +
        '</div>' +
        '</div>'
    );
  }

  /**
   * 掲載内容を読み込めなかったときの案内。
   * file:// で直接開くと fetch がブラウザに禁じられるため、まず起きるのはこれ。
   * 黙って白紙になると原因が分からないので、開き方を説明する。
   */
  function renderLoadError(name) {
    var el = document.getElementById('page');
    if (!el) return;
    var isFile = location.protocol === 'file:';

    el.innerHTML =
      '<section><div class="wrap" style="max-width:720px">' +
      '<p class="eyebrow">' + (isFile ? 'LOCAL FILE — 開き方' : 'ERROR — 読み込めません') + '</p>' +
      '<h2>' +
      (isFile ? 'このページは、サーバー経由で<br>開く必要があります。' : '掲載内容を読み込めませんでした。') +
      '</h2>' +
      (isFile
        ? '<p class="lead">掲載内容は <code>content/' + esc(name) + '.json</code> から読み込んでいますが、' +
          'ファイルを直接開いた状態（file://）では、ブラウザがその読み込みを禁止します。' +
          'これは安全のための仕組みで、設定では変えられません。</p>' +
          '<div class="callout" style="margin-top:28px">' +
          '<h3>かんたんな開き方</h3>' +
          '<p>コマンドプロンプト（または PowerShell）で次を実行し、表示されたURLを開いてください。</p>' +
          '<pre class="code-block">cd C:\\Users\\with_\\wize-site\nnode tools/serve.mjs</pre>' +
          '<p>サイト → <strong>http://localhost:4173/</strong><br>' +
          '管理画面 → <strong>http://localhost:4173/admin/</strong></p>' +
          '</div>'
        : '<p class="lead">時間をおいて再度お試しください。何度も表示される場合は、' +
          '<code>content/' + esc(name) + '.json</code> が壊れていないかご確認ください。</p>') +
      '</div></section>';
  }

  /** ページ上のヘッダー・フッターへ流し込み、メニューの開閉を配線する */
  function renderChrome(site, current) {
    if (!site) return;
    var header = document.querySelector('[data-chrome="header"]');
    var footer = document.querySelector('[data-chrome="footer"]');

    if (header) {
      header.className = 'site';
      header.innerHTML = headerHTML(site, current);
      var toggle = header.querySelector('.nav-toggle');
      var links = header.querySelector('.nav-links');
      toggle.addEventListener('click', function () {
        var open = links.classList.toggle('is-open');
        toggle.setAttribute('aria-expanded', String(open));
        toggle.textContent = open ? 'CLOSE' : 'MENU';
      });
    }

    if (footer) {
      footer.className = 'site';
      footer.innerHTML = footerHTML(site);
    }
  }

  /** content 側では "note/" のようにルート相対で書く。ここで実パスに直す */
  function resolveHref(href) {
    if (href == null) return '#';
    /* 単体HTMLとして書き出された場合、ページ同士は home.html / note.html … と
       平らに並ぶ。書き出し時に埋め込んだ対応表があればそれを優先する。
       これが無いと、描画のたびにサイト内リンクが元の "note/" に戻ってしまう。 */
    var map = global.WIZE_LINKMAP;
    if (map && Object.prototype.hasOwnProperty.call(map, href)) return map[href];
    if (href === '') return url(''); // 空文字はサイトのトップを指す
    if (/^(https?:|mailto:|tel:|#)/.test(href)) return href;
    return url(href);
  }

  /* ---------- スクロール出現 ---------- */

  var revealFailsafe;
  function observeReveal(root) {
    // 隠すのは JS が動いているときだけ（CSS 側は .js-reveal 配下でしか opacity を下げない）
    document.documentElement.classList.add('js-reveal');

    var targets = (root || document).querySelectorAll('.reveal:not(.on)');
    if (!targets.length) return;
    var reduced = global.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (reduced || !('IntersectionObserver' in global)) {
      Array.prototype.forEach.call(targets, function (el) {
        el.classList.add('on');
      });
      return;
    }

    /* 保険：非表示タブ等で IntersectionObserver が発火しないまま終わると
       本文が見えなくなるため、一定時間後に残りを強制的に表示する。 */
    clearTimeout(revealFailsafe);
    revealFailsafe = setTimeout(function () {
      document.querySelectorAll('.reveal:not(.on)').forEach(function (el) {
        el.classList.add('on');
      });
    }, 2500);
    var io = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (e) {
          if (e.isIntersecting) {
            e.target.classList.add('on');
            io.unobserve(e.target);
          }
        });
      },
      { threshold: 0.12 }
    );
    Array.prototype.forEach.call(targets, function (el) {
      io.observe(el);
    });
  }

  /* ---------- 価格ボタン ---------- */

  function priceLink(item, fallbackLabel) {
    if (!item.href) return '';
    var label = item.price || fallbackLabel || '記事を読む';
    return (
      '<a class="price' + (item.free ? ' is-free' : '') + '" href="' + esc(item.href) + '" target="_blank" rel="noopener">' +
      esc(label) +
      '</a>'
    );
  }

  function setMeta(meta) {
    if (!meta) return;
    if (meta.title) document.title = meta.title;
    if (meta.description) {
      var m = document.querySelector('meta[name="description"]');
      if (m) m.setAttribute('content', meta.description);
    }
  }

  global.WIZE = {
    BASE: BASE,
    url: url,
    withBase: withBase,
    esc: esc,
    nl2br: nl2br,
    rich: rich,
    load: load,
    pages: {}, // 各ページの render(data) を page-*.js が登録する
    headerHTML: headerHTML,
    footerHTML: footerHTML,
    renderChrome: renderChrome,
    renderLoadError: renderLoadError,
    resolveHref: resolveHref,
    observeReveal: observeReveal,
    priceLink: priceLink,
    setMeta: setMeta,
    DRAFT_PREFIX: DRAFT_PREFIX,
  };
})(window);
