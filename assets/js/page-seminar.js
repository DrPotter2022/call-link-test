/* セミナーページ。掲載内容は content/seminar.json */
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

  function lecture(l) {
    if (l.published === false) return '';
    return (
      '<div class="lecture reveal">' +
      '<span class="lecture-code">' + esc(l.code) + '</span>' +
      '<h4>' + esc(l.title) + '</h4>' +
      (l.duration ? '<span class="duration">⏲ ' + esc(l.duration) + '</span>' : '') +
      (l.body ? '<p class="lecture-body">' + br(l.body) + '</p>' : '') +
      '<div class="lecture-foot">' +
      (l.href
        ? '<a class="btn btn-ghost btn-sm" href="' + esc(l.href) + '" target="_blank" rel="noopener">' +
          esc(l.ctaLabel || 'Udemyで受講する') +
          '</a>'
        : '<span class="note-text">準備中</span>') +
      (l.price ? '<span class="price' + (l.free ? ' is-free' : '') + '">' + esc(l.price) + '</span>' : '') +
      '</div>' +
      '</div>'
    );
  }

  function course(c) {
    var lectures = (c.lectures || []).map(lecture).join('');
    return (
      '<div class="course' + (c.focus ? ' is-focus' : '') + ' reveal" id="' + esc(c.id) + '" style="margin-bottom:28px">' +
      '<span class="course-tag">' + esc(c.tag) + '</span>' +
      '<h3>' + esc(c.title) + '</h3>' +
      (c.for ? '<p class="course-for">' + esc(c.for) + '</p>' : '') +
      (c.lead ? '<p class="course-lead">' + br(c.lead) + '</p>' : '') +
      (c.lessons && c.lessons.length
        ? '<ol class="lessons">' + c.lessons.map(function (l) { return '<li>' + esc(l) + '</li>'; }).join('') + '</ol>'
        : '') +
      (lectures
        ? '<p class="eyebrow" style="margin:32px 0 18px">' + esc(c.lecturesTitle || '講座') + '</p>' +
          '<div class="grid grid-2">' + lectures + '</div>'
        : '') +
      '</div>'
    );
  }

  function render(d) {
    var h = d.hero || {};
    var m = d.message || {};
    var v = d.voices || {};
    var ct = d.cta || {};

    return [
      '<section class="hero hero--split">',
      h.tate ? '<p class="tate">' + esc(h.tate) + '</p>' : '',
      '<div class="wrap">',
      '<div>',
      '<p class="eyebrow">' + esc(h.eyebrow) + '</p>',
      '<h1>' + W.rich(h.title) + '</h1>',
      '<p>' + br(h.body) + '</p>',
      '<div class="hero-ctas">' + cta(h.primaryCta, 'btn-primary') + cta(h.secondaryCta, 'btn-ghost') + '</div>',
      '</div>',
      '<div class="reveal" style="display:flex;justify-content:center">' + hanko(h.hankoMain, h.hankoSub) + '</div>',
      '</div>',
      '</section>',

      /* コース一覧 */
      '<section id="courses" style="background:var(--shiro);border-top:1px solid var(--line);border-bottom:1px solid var(--line)">',
      '<div class="wrap">',
      '<div class="reveal"><p class="eyebrow">COURSES — 難易度別のコース</p>',
      '<h2>それぞれのテーマに合わせて、難易度別に。</h2></div>',
      '<div class="year-tabs" style="margin-top:24px">' +
        (d.courses || [])
          .map(function (c) {
            return '<a class="year-tab" href="#' + esc(c.id) + '" style="text-decoration:none">' + esc(c.tag) + '</a>';
          })
          .join('') +
        '</div>',
      '<div style="margin-top:12px">' + (d.courses || []).map(course).join('') + '</div>',
      '</div>',
      '</section>',

      /* 講師メッセージ */
      '<section class="pledge">',
      '<div class="wrap">',
      '<div class="reveal" style="max-width:760px;margin:0 auto">',
      '<p class="eyebrow">' + esc(m.eyebrow) + '</p>',
      '<h2>' + br(m.title) + '</h2>',
      '</div>',
      '<div class="pledge-paper reveal">',
      '<p style="font-size:14.5px;line-height:2;white-space:pre-line;color:var(--sumi-2)">' + br(m.body) + '</p>',
      '</div>',
      '</div>',
      '</section>',

      /* 受講者の声 */
      '<section>',
      '<div class="wrap">',
      '<div class="reveal">',
      '<p class="eyebrow">' + esc(v.eyebrow) + '</p>',
      '<h2>' + br(v.title) + '</h2>',
      '</div>',
      '<div class="voices reveal">' +
        (v.items || [])
          .map(function (x) {
            return '<div class="voice"><span>' + esc(x) + '</span></div>';
          })
          .join('') +
        '</div>',
      '</div>',
      '</section>',

      /* CTA */
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

  W.pages.seminar = render;

  // 管理画面から読み込まれたときは、描画せず render だけ提供する
  if (!document.getElementById('page')) return;

  Promise.all([W.load('site'), W.load('seminar')]).then(function (r) {
    W.renderChrome(r[0], 'seminar');
    var d = r[1];
    if (!d) return W.renderLoadError('seminar');
    W.setMeta(d.meta);
    document.getElementById('page').innerHTML = render(d);
    W.observeReveal();
  });
})();
