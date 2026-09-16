/* .note（バックナンバー）ページ。掲載内容は content/note.json */
(function () {
  'use strict';
  var W = window.WIZE;
  var esc = W.esc;
  var br = W.nl2br;

  var state = { category: 'all', query: '', year: null };
  var data = null;

  /* ---------- 記事カード ---------- */

  function articleCard(it, ref) {
    var hidden = it.published === false;
    if (hidden && !W.editMode) return '';
    return (
      '<article class="article reveal' + (hidden ? ' is-unpublished' : '') + '"' +
      ' data-ref="' + esc(ref) + '"' +
      ' data-title="' + esc((it.title || '') + ' ' + (it.summary || '')) + '">' +
      (it.image
        ? '<div class="article-thumb"><img src="' + esc(W.url(it.image)) + '" alt="" loading="lazy" decoding="async"></div>'
        : '<div class="article-thumb"></div>') +
      '<div class="article-body">' +
      '<div class="article-meta">' +
      (it.date ? '<span class="date">' + esc(it.date) + '</span>' : '') +
      (it.era ? '<span class="era">' + esc(it.era) + '</span>' : '') +
      '</div>' +
      '<h3>' +
      (it.href ? '<a href="' + esc(it.href) + '" target="_blank" rel="noopener">' + esc(it.title) + '</a>' : esc(it.title)) +
      '</h3>' +
      (it.summary ? '<p class="summary">' + br(it.summary) + '</p>' : '') +
      '<div class="article-foot">' +
      W.priceLink(it) +
      (it.summary && it.summary.length > 90 ? '<button class="more" type="button">全文を読む</button>' : '') +
      '</div>' +
      '</div>' +
      '</article>'
    );
  }

  /* ---------- 分野セクション ---------- */

  function categorySection(cat, ci) {
    var groups = (cat.groups || [])
      .map(function (g, gi) {
        var lastEra = null;
        var items = (g.items || [])
          .map(function (it, ii) {
            var head = '';
            if (it.era && it.era !== lastEra) {
              lastEra = it.era;
              head =
                '<h4 class="era-head">' + esc(it.era) + (it.eraNote ? '<small>' + esc(it.eraNote) + '</small>' : '') + '</h4>';
            }
            return head + articleCard(it, 'categories.' + ci + '.groups.' + gi + '.items.' + ii);
          })
          .join('');
        if (!items) return '';
        return (g.title ? '<h3 class="group-head">' + esc(g.title) + '</h3>' : '') + items;
      })
      .join('');

    return (
      '<section class="category" id="' + esc(cat.id) + '" data-category="' + esc(cat.id) + '" style="padding:0">' +
      '<div class="cat-head reveal">' +
      '<h2>' + esc(cat.label) + '</h2>' +
      (cat.lead ? '<p class="lead">' + br(cat.lead) + '</p>' : '') +
      '</div>' +
      groups +
      '</section>'
    );
  }

  /* ---------- 年別バックナンバー ---------- */

  function issueCard(it, ref) {
    var hidden = it.published === false;
    if (hidden && !W.editMode) return '';
    return (
      '<a class="issue' + (hidden ? ' is-unpublished' : '') + '"' +
      ' data-ref="' + esc(ref) + '"' +
      ' href="' + esc(it.href || '#') + '"' + (it.href ? ' target="_blank" rel="noopener"' : '') + '>' +
      (it.image ? '<div class="issue-thumb"><img src="' + esc(W.url(it.image)) + '" alt="" loading="lazy" decoding="async"></div>' : '') +
      '<div class="issue-body">' +
      (it.label ? '<span class="label">' + esc(it.label) + '</span>' : '') +
      '<span class="title">' + esc(it.title) + '</span>' +
      (it.price ? '<span class="price' + (it.free ? ' is-free' : '') + '">' + esc(it.price) + '</span>' : '') +
      '</div>' +
      '</a>'
    );
  }

  /** 選択中の年の号を、データ上の位置つきで組み立てる */
  function magazineItems(mags) {
    var mi = -1;
    for (var i = 0; i < mags.length; i++) if (mags[i].id === state.year) { mi = i; break; }
    if (mi < 0) mi = 0;
    if (!mags[mi]) return '';
    return mags[mi].items
      .map(function (it, ii) { return issueCard(it, 'magazines.' + mi + '.items.' + ii); })
      .join('');
  }

  function magazineSection(mags) {
    if (!mags || !mags.length) return '';
    if (!state.year) state.year = mags[mags.length - 1].id;
    return (
      '<section id="backnumber" style="background:var(--shiro);border-top:1px solid var(--line)">' +
      '<div class="wrap">' +
      '<div class="reveal">' +
      '<p class="eyebrow">BACK NUMBER — 年別バックナンバー</p>' +
      '<h2>WITHOVER News ／ WIZE News</h2>' +
      '<p class="lead">セミナー受講生・クライアントへ配信している継続学習用メールマガジンのバックナンバーです。</p>' +
      '</div>' +
      '<div class="year-tabs" id="year-tabs">' +
      mags
        .map(function (m) {
          return (
            '<button class="year-tab" type="button" data-year="' + esc(m.id) + '" aria-pressed="' + (m.id === state.year) + '">' +
            esc(m.title) +
            '</button>'
          );
        })
        .join('') +
      '</div>' +
      '<div class="grid grid-4" id="issue-grid">' +
      magazineItems(mags) +
      '</div>' +
      '</div>' +
      '</section>'
    );
  }

  function subscriptionSection(subs) {
    if (!subs || !subs.length) return '';
    return subs
      .map(function (s, si) {
        return (
          '<section>' +
          '<div class="wrap">' +
          '<div class="reveal">' +
          '<p class="eyebrow">MAGAZINE — マガジン・定期購読</p>' +
          '<h2>' + esc(s.title) + '</h2>' +
          '</div>' +
          '<div class="grid grid-4 reveal" style="margin-top:32px">' +
          s.items.map(function (it, ii) { return issueCard(it, 'subscriptions.' + si + '.items.' + ii); }).join('') +
          '</div>' +
          '</div>' +
          '</section>'
        );
      })
      .join('');
  }

  /* ---------- ページ全体 ---------- */

  function render(d) {
    var h = d.hero || {};
    return [
      '<section class="hero">',
      h.tate ? '<p class="tate">' + esc(h.tate) + '</p>' : '',
      '<div class="wrap">',
      '<p class="eyebrow">' + esc(h.eyebrow) + '</p>',
      '<h1>' + W.rich(h.title) + '</h1>',
      '<p>' + br(h.lead) + '</p>',
      '</div>',
      '</section>',

      '<section class="section--tight" style="background:var(--shiro);border-top:1px solid var(--line);border-bottom:1px solid var(--line)">',
      '<div class="wrap">',
      d.caution && d.caution.title
        ? '<div class="callout reveal"><h3>' + esc(d.caution.title) + '</h3><p>' + br(d.caution.body) + '</p></div>'
        : '',
      '<div class="reveal"><p class="eyebrow">PRICING — リンク先の記事と価格設定</p>',
      '<div class="legend">' +
        (d.priceLegend || [])
          .map(function (l) {
            return '<div><span class="legend-label">' + esc(l.label) + '</span><p>' + br(l.note) + '</p></div>';
          })
          .join('') +
        '</div></div>',
      '</div>',
      '</section>',

      '<section style="padding-top:0">',
      '<div class="wrap">',
      '<div class="filters" id="filters">',
      '<button class="filter" type="button" data-cat="all" aria-pressed="true">すべて</button>',
      (d.categories || [])
        .map(function (c) {
          return '<button class="filter" type="button" data-cat="' + esc(c.id) + '" aria-pressed="false">' + esc(c.label) + '</button>';
        })
        .join(''),
      '<input class="search" id="search" type="search" placeholder="キーワードで絞り込む（例：NISA、火災保険、為替）" aria-label="記事を検索">',
      '</div>',
      '<p class="result-count" id="result-count"></p>',
      '<div id="categories">' + (d.categories || []).map(categorySection).join('') + '</div>',
      '</div>',
      '</section>',

      magazineSection(d.magazines),
      subscriptionSection(d.subscriptions),
    ].join('');
  }

  /* ---------- 絞り込み ---------- */

  function applyFilter() {
    var q = state.query.trim().toLowerCase();
    var shown = 0;
    document.querySelectorAll('#categories .category').forEach(function (sec) {
      var catOk = state.category === 'all' || sec.dataset.category === state.category;
      var visibleInSec = 0;
      sec.querySelectorAll('.article').forEach(function (a) {
        var hit = catOk && (!q || a.dataset.title.toLowerCase().indexOf(q) > -1);
        a.classList.toggle('is-hidden', !hit);
        if (hit) {
          visibleInSec++;
          shown++;
        }
      });
      sec.classList.toggle('is-hidden', visibleInSec === 0);
      // 見出しは中身が残っているときだけ出す
      sec.querySelectorAll('.group-head, .era-head').forEach(function (hd) {
        var next = hd.nextElementSibling;
        var any = false;
        while (next && !next.classList.contains('group-head')) {
          if (hd.classList.contains('era-head') && next.classList.contains('era-head')) break;
          if (next.classList.contains('article') && !next.classList.contains('is-hidden')) {
            any = true;
            break;
          }
          next = next.nextElementSibling;
        }
        hd.classList.toggle('is-hidden', !any);
      });
    });
    var counter = document.getElementById('result-count');
    if (counter) counter.textContent = shown + ' 件を表示中';
  }

  /* ---------- 起動 ---------- */

  W.pages.note = function (d) {
    state.year = null; // 書き出しのたびに既定の年へ戻す
    return render(d);
  };

  /* 編集モード（assets/js/admin-inline.js）から使う出入口 */
  W.notePage = {
    getData: function () { return data; },
    getState: function () { return state; },
    rerender: function () {
      if (!data) return;
      document.getElementById('page').innerHTML = render(data);
      // 描き直しで作り直される操作部に、いまの検索語・分野を戻す
      var s = document.getElementById('search');
      if (s) s.value = state.query;
      document.querySelectorAll('.filter').forEach(function (b) {
        b.setAttribute('aria-pressed', String(b.dataset.cat === state.category));
      });
      applyFilter();
      W.observeReveal();
    },
    magazineItems: function () { return magazineItems(data.magazines || []); },
  };

  // 管理画面から読み込まれたときは、描画せず render だけ提供する
  if (!document.getElementById('page')) return;

  Promise.all([W.load('site'), W.load('note')]).then(function (r) {
    W.renderChrome(r[0], 'note');
    data = r[1];
    if (!data) return W.renderLoadError('note');
    W.setMeta(data.meta);
    document.getElementById('page').innerHTML = render(data);
    applyFilter();
    W.observeReveal();

    var page = document.getElementById('page');

    page.addEventListener('click', function (e) {
      var f = e.target.closest('.filter');
      if (f) {
        state.category = f.dataset.cat;
        page.querySelectorAll('.filter').forEach(function (b) {
          b.setAttribute('aria-pressed', String(b === f));
        });
        applyFilter();
        W.observeReveal();
        return;
      }

      var m = e.target.closest('.more');
      if (m) {
        var art = m.closest('.article');
        var open = art.classList.toggle('is-open');
        m.textContent = open ? '閉じる' : '全文を読む';
        return;
      }

      var y = e.target.closest('.year-tab');
      if (y) {
        state.year = y.dataset.year;
        page.querySelectorAll('.year-tab').forEach(function (b) {
          b.setAttribute('aria-pressed', String(b === y));
        });
        document.getElementById('issue-grid').innerHTML = magazineItems(data.magazines || []);
        if (W.notePage.afterRender) W.notePage.afterRender();
      }
    });

    /* 検索欄は #page ごと描き直されることがある（編集モードでの更新など）ので、
       要素に直接ではなく #page 側で受ける。 */
    var timer;
    page.addEventListener('input', function (e) {
      if (!e.target.matches('#search')) return;
      var value = e.target.value;
      clearTimeout(timer);
      timer = setTimeout(function () {
        state.query = value;
        applyFilter();
        W.observeReveal();
      }, 120);
    });
  });
})();
