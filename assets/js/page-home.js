/* トップページの組み立て。掲載内容は content/home.json + content/site.json */
(function () {
  'use strict';
  var W = window.WIZE;
  var esc = W.esc;
  var br = W.nl2br;

  function cta(c, cls) {
    if (!c || !c.label) return '';
    var ext = /^https?:/.test(c.href) ? ' target="_blank" rel="noopener"' : '';
    return '<a class="btn ' + cls + '" href="' + esc(W.resolveHref(c.href)) + '"' + ext + '>' + esc(c.label) + '</a>';
  }

  function hanko(main, sub) {
    if (!main) return '';
    return '<div class="hanko" aria-hidden="true">' + esc(main) + (sub ? '<small>' + esc(sub) + '</small>' : '') + '</div>';
  }

  function render(d) {
    var h = d.hero || {};
    var p = d.problem || {};
    var cr = d.creed || {};
    var co = d.contents || {};
    var bu = d.business || {};
    var re = d.recruit || {};
    var ct = d.cta || {};

    return [
      /* ---- ヒーロー ---- */
      '<section class="hero hero--split">',
      h.tate ? '<p class="tate">' + esc(h.tate) + '</p>' : '',
      '<div class="wrap">',
      '<div>',
      '<p class="eyebrow">' + esc(h.eyebrow) + '</p>',
      '<h1>' + W.rich(h.title) + '</h1>',
      '<p>' + br(h.body) + '</p>',
      h.note ? '<p class="note-text" style="margin-top:16px">' + br(h.note) + '</p>' : '',
      '<div class="hero-ctas">' + cta(h.primaryCta, 'btn-primary') + cta(h.secondaryCta, 'btn-ghost') + '</div>',
      '</div>',
      '<div class="reveal" style="display:flex;justify-content:center">' + hanko(h.hankoMain, h.hankoSub) + '</div>',
      '</div>',
      '</section>',

      /* ---- 問題提起 ---- */
      '<section style="background:var(--shiro);border-top:1px solid var(--line);border-bottom:1px solid var(--line)">',
      '<div class="wrap grid grid-2" style="align-items:center;gap:56px">',
      '<div class="reveal">',
      '<p class="eyebrow">' + esc(p.eyebrow) + '</p>',
      '<h2>' + br(p.title) + '</h2>',
      '<p class="lead">' + br(p.body) + '</p>',
      '</div>',
      '<div class="reveal">',
      '<div class="card is-focus" style="box-shadow:none;border-style:dashed">',
      '<p class="tag">' + esc(p.cardTag) + '</p>',
      '<h3>' + esc(p.cardTitle) + '</h3>',
      '<ul>' + (p.cardPoints || []).map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul>',
      '</div>',
      '</div>',
      '</div>',
      '</section>',

      /* ---- ミッション ---- */
      '<section>',
      '<div class="wrap wrap--narrow reveal">',
      '<p class="eyebrow">' + esc(cr.eyebrow) + '</p>',
      '<h2>' + br(cr.title) + '</h2>',
      '<div class="creed" style="margin-top:26px">',
      '<span class="creed-mark">' + esc(cr.mark) + '</span>',
      '<p>' + br(cr.body) + '</p>',
      '</div>',
      '</div>',
      '</section>',

      /* ---- 提供しているもの ---- */
      '<section style="background:var(--shiro);border-top:1px solid var(--line)">',
      '<div class="wrap">',
      '<div class="reveal">',
      '<p class="eyebrow">' + esc(co.eyebrow) + '</p>',
      '<h2>' + br(co.title) + '</h2>',
      '<p class="lead">' + br(co.lead) + '</p>',
      '</div>',
      '<div class="grid grid-3 reveal" style="margin-top:40px">',
      (co.cards || [])
        .map(function (c) {
          return (
            '<div class="card' + (c.focus ? ' is-focus' : '') + '">' +
            '<p class="tag">' + esc(c.tag) + '</p>' +
            '<h3>' + esc(c.title) + '</h3>' +
            '<p>' + br(c.body) + '</p>' +
            '<div class="card-foot">' + cta({ label: c.ctaLabel, href: c.ctaHref }, c.focus ? 'btn-primary btn-sm' : 'btn-ghost btn-sm') + '</div>' +
            '</div>'
          );
        })
        .join(''),
      '</div>',
      '</div>',
      '</section>',

      /* ---- 事業内容 ---- */
      '<section>',
      '<div class="wrap">',
      '<div class="reveal">',
      '<p class="eyebrow">' + esc(bu.eyebrow) + '</p>',
      '<h2>' + br(bu.title) + '</h2>',
      '</div>',
      '<div class="grid grid-2 reveal" style="margin-top:36px">',
      (bu.items || [])
        .map(function (b) {
          return (
            '<div class="card">' +
            '<span style="font-family:var(--disp);font-weight:800;font-size:34px;color:var(--ai);line-height:1;border-bottom:2px solid var(--ai);display:inline-block;padding-bottom:6px">' +
            esc(b.kanji) + '</span>' +
            '<h3 style="margin-top:14px">' + esc(b.title) + '</h3>' +
            '<p class="note-text" style="margin-bottom:12px">' + esc(b.for) + '</p>' +
            '<p>' + br(b.body) + '</p>' +
            '</div>'
          );
        })
        .join(''),
      '</div>',
      '</div>',
      '</section>',

      /* ---- 同業の方へ ---- */
      re.enabled === false
        ? ''
        : [
            '<section class="pledge">',
            '<div class="wrap wrap--narrow reveal">',
            '<p class="eyebrow">' + esc(re.eyebrow) + '</p>',
            '<h2>' + br(re.title) + '</h2>',
            '<p class="lead">' + br(re.body) + '</p>',
            '<div style="margin-top:30px;display:flex;gap:18px;align-items:center;flex-wrap:wrap">',
            re.ctaHref ? cta({ label: re.ctaLabel, href: re.ctaHref }, 'btn-primary') : '',
            re.status
              ? '<span class="note-text" style="color:#C4D2DC;border:1px dashed #4C6478;padding:10px 18px">' + esc(re.status) + '</span>'
              : '',
            '</div>',
            '</div>',
            '</section>',
          ].join(''),

      /* ---- CTA ---- */
      '<section class="cta">',
      '<div class="wrap">',
      '<p class="eyebrow eyebrow--center reveal">' + esc(ct.eyebrow) + '</p>',
      '<h2 class="reveal">' + br(ct.title) + '</h2>',
      '<p class="reveal">' + br(ct.body) + '</p>',
      '<div class="cta-ctas reveal">' + cta(ct.primaryCta, 'btn-primary') + cta(ct.secondaryCta, 'btn-ghost') + '</div>',
      ct.note ? '<p class="note-text reveal" style="margin-top:24px">' + br(ct.note) + '</p>' : '',
      '</div>',
      '</section>',
    ].join('');
  }

  W.pages.home = render;

  // 管理画面から読み込まれたときは、描画せず render だけ提供する
  if (!document.getElementById('page')) return;

  Promise.all([W.load('site'), W.load('home')]).then(function (r) {
    var site = r[0];
    var home = r[1];
    W.renderChrome(site, 'home');
    if (!home) return W.renderLoadError('home');
    W.setMeta(home.meta);
    document.getElementById('page').innerHTML = render(home);
    W.observeReveal();
  });
})();
