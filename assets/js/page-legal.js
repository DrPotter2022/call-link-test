/* 法務ページ（プライバシーポリシー／決済に関する利用規約）。
   掲載内容は content/legal.json。どちらを出すかは #page の data-legal で決める。 */
(function () {
  'use strict';
  var W = window.WIZE;
  var esc = W.esc;

  /**
   * 本文中の [文字](URL) をリンクにする。
   * 元のページにあったリンクを保つためで、それ以外のHTMLは書けない。
   */
  function body(s) {
    return W.nl2br(s).replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g, function (_, label, href) {
      return '<a href="' + esc(href) + '" target="_blank" rel="noopener">' + esc(label) + '</a>';
    });
  }

  function list(items, depth) {
    return (
      '<ul class="legal-list' + (depth ? ' is-nested' : '') + '">' +
      items
        .map(function (it) {
          if (typeof it === 'string') return '<li>' + body(it) + '</li>';
          if (it && it.type === 'list') return '<li class="is-holder">' + list(it.items || [], (depth || 0) + 1) + '</li>';
          return '';
        })
        .join('') +
      '</ul>'
    );
  }

  function render(d) {
    if (!d) return '';
    var blocks = (d.blocks || [])
      .map(function (b) {
        if (b.type === 'heading') return '<h2 class="legal-h">' + esc(b.text) + '</h2>';
        if (b.type === 'subheading') return '<h3 class="legal-sub">' + esc(b.text) + '</h3>';
        if (b.type === 'paragraph') return '<p class="legal-p">' + body(b.text) + '</p>';
        if (b.type === 'list') return list(b.items || [], 0);
        return '';
      })
      .join('');

    return (
      '<section class="legal">' +
      '<div class="wrap wrap--narrow">' +
      '<p class="eyebrow">' + esc(d.eyebrow || 'LEGAL') + '</p>' +
      '<h1 class="legal-title">' + esc(d.title) + '</h1>' +
      '<div class="legal-body">' + blocks + '</div>' +
      '</div>' +
      '</section>'
    );
  }

  /* 書き出し用。legal.json のうち1件分を受け取る */
  W.pages.privacy = render;
  W.pages.terms = render;

  // 管理画面から読み込まれたときは、描画せず render だけ提供する
  var page = document.getElementById('page');
  if (!page || !page.dataset.legal) return;

  var which = page.dataset.legal;

  Promise.all([W.load('site'), W.load('legal')]).then(function (r) {
    W.renderChrome(r[0], which);
    var all = r[1];
    if (!all || !all[which]) return W.renderLoadError('legal');
    W.setMeta(all[which].meta);
    page.innerHTML = render(all[which]);
    W.observeReveal();
  });
})();
