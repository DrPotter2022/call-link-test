/* ============================================================
   .note ページの編集モード
   管理者としてログインしている場合だけ読み込まれ、
   記事とメールマガジンのバックナンバーを、その場で追加・削除・編集できる。

   変更は管理画面と同じ下書き（localStorage の wize_draft_note）に入る。
   実際のサイトに反映するには、管理画面から「GitHubへ公開」する必要がある。

   ※ 静的サイトなので、この編集モードはあくまで操作の入口です。
     公開にはGitHubのアクセストークンが要り、そちらが本当の防御になっています。
   ============================================================ */
(function (global) {
  'use strict';

  var W = global.WIZE;
  var DRAFT_KEY = 'wize_draft_note';
  var SESSION_KEY = 'wize_admin_session';
  var MODE_KEY = 'wize_edit_mode';

  var esc = W.esc;

  /* 項目名の日本語ラベル */
  var LABELS = {
    title: 'タイトル', label: '号数・ラベル', summary: '本文', date: '執筆日',
    era: '時代・テーマの見出し', eraNote: '時代の補足', price: '価格の表示',
    free: '無料として扱う（朱色で表示）', href: 'リンク先URL', image: '画像のパス',
    published: 'サイトに表示する', id: 'ID（変更しないでください）',
  };

  /* 入力欄を大きくする項目 */
  var LONG = ['summary'];

  /* 編集画面に出す順番。ここに無いキーは後ろへ回す */
  var ORDER = ['label', 'title', 'summary', 'date', 'era', 'eraNote', 'price', 'free', 'href', 'image', 'published', 'id'];

  function session() {
    try {
      var s = JSON.parse(localStorage.getItem(SESSION_KEY) || 'null');
      if (!s || !s.until || Date.now() > s.until) return null;
      return s;
    } catch (e) { return null; }
  }

  /* ---------- データの出し入れ ---------- */

  function data() { return W.notePage.getData(); }

  /** "magazines.3.items.7" のような位置指定を、配列と添字に解く */
  function resolve(ref) {
    var parts = String(ref).split('.');
    var node = data();
    for (var i = 0; i < parts.length - 1; i++) {
      var k = parts[i];
      node = Array.isArray(node) ? node[Number(k)] : node[k];
      if (node == null) return null;
    }
    var idx = Number(parts[parts.length - 1]);
    return { list: node, index: idx, item: node[idx] };
  }

  var saveTimer;
  function save(immediate) {
    clearTimeout(saveTimer);
    var go = function () {
      try {
        localStorage.setItem(DRAFT_KEY, JSON.stringify(data()));
        setStatus('下書きに保存しました', 'ok');
      } catch (e) {
        setStatus('保存できません（容量超過の可能性）', 'err');
      }
    };
    if (immediate) go(); else saveTimer = setTimeout(go, 400);
  }

  function refresh() {
    W.notePage.rerender();
    decorate();
  }

  /* ---------- 画面下の操作バー ---------- */

  var bar, statusEl;

  function setStatus(msg, kind) {
    if (!statusEl) return;
    statusEl.textContent = msg || '';
    statusEl.className = 'wa-status' + (kind ? ' is-' + kind : '');
    if (kind === 'ok') {
      setTimeout(function () {
        if (statusEl.textContent === msg) statusEl.textContent = '';
      }, 2500);
    }
  }

  function buildBar(who) {
    bar = document.createElement('div');
    bar.id = 'wize-admin-bar';
    bar.innerHTML =
      '<span class="wa-who">' + esc(who || '管理者') + '</span>' +
      '<label class="wa-toggle"><input type="checkbox" id="wa-mode"><span>編集モード</span></label>' +
      '<span class="wa-status" id="wa-status"></span>' +
      '<span class="wa-spacer"></span>' +
      '<button type="button" class="wa-btn" id="wa-discard">下書きを破棄</button>' +
      '<a class="wa-btn is-primary" href="' + esc(W.url('admin/')) + '">管理画面で公開する</a>';
    document.body.appendChild(bar);
    statusEl = bar.querySelector('#wa-status');

    var toggle = bar.querySelector('#wa-mode');
    toggle.checked = localStorage.getItem(MODE_KEY) === '1';
    applyMode(toggle.checked);
    toggle.addEventListener('change', function () {
      localStorage.setItem(MODE_KEY, toggle.checked ? '1' : '0');
      applyMode(toggle.checked);
      refresh();
    });

    bar.querySelector('#wa-discard').addEventListener('click', function () {
      if (!localStorage.getItem(DRAFT_KEY)) return setStatus('下書きはありません');
      if (!confirm('このページの下書きを破棄して、公開中の内容に戻します。よろしいですか？')) return;
      localStorage.removeItem(DRAFT_KEY);
      location.reload();
    });
  }

  function applyMode(on) {
    W.editMode = on;
    document.body.classList.toggle('wize-edit', on);
  }

  /* ---------- カードに操作ボタンを付ける ---------- */

  function toolbar(ref, kind) {
    return (
      '<div class="wa-tools" data-for="' + esc(ref) + '">' +
      '<button type="button" data-act="edit" title="編集">編集</button>' +
      '<button type="button" data-act="up" title="前へ">↑</button>' +
      '<button type="button" data-act="down" title="後へ">↓</button>' +
      '<button type="button" data-act="dup" title="複製">複製</button>' +
      '<button type="button" data-act="del" class="is-danger" title="削除">削除</button>' +
      '</div>'
    );
  }

  function decorate() {
    if (!W.editMode) return;

    document.querySelectorAll('[data-ref]').forEach(function (el) {
      if (el.querySelector(':scope > .wa-tools')) return;
      var wrap = document.createElement('div');
      wrap.innerHTML = toolbar(el.dataset.ref);
      el.appendChild(wrap.firstChild);
      if (el.classList.contains('is-unpublished')) {
        var badge = document.createElement('span');
        badge.className = 'wa-badge';
        badge.textContent = '非公開';
        el.appendChild(badge);
      }
    });

    // 年別バックナンバーに「号を追加」
    var grid = document.getElementById('issue-grid');
    if (grid && !document.getElementById('wa-add-issue')) {
      var b = document.createElement('button');
      b.type = 'button';
      b.id = 'wa-add-issue';
      b.className = 'wa-add';
      b.textContent = '＋ この年に号を追加';
      grid.parentNode.insertBefore(b, grid.nextSibling);
      b.addEventListener('click', addIssue);
    }

    // 分野ごとに「記事を追加」
    document.querySelectorAll('#categories .category').forEach(function (sec) {
      if (sec.querySelector(':scope > .wa-add')) return;
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'wa-add';
      b.textContent = '＋ この分野に記事を追加';
      b.addEventListener('click', function () { addArticle(sec.dataset.category); });
      sec.appendChild(b);
    });
  }

  /* ---------- 追加 ---------- */

  function blankLike(sample, fallback) {
    var src = sample || fallback;
    var out = {};
    Object.keys(src).forEach(function (k) {
      var v = src[k];
      if (k === 'id') out[k] = 'new-' + Date.now().toString(36);
      else if (typeof v === 'boolean') out[k] = k === 'published' ? true : false;
      else if (typeof v === 'number') out[k] = 0;
      else if (Array.isArray(v)) out[k] = [];
      else out[k] = '';
    });
    return out;
  }

  var ISSUE_SHAPE = { id: '', label: '', title: '', price: '', free: false, href: '', image: '', published: true };
  var ARTICLE_SHAPE = { id: '', title: '', summary: '', date: '', era: '', eraNote: '', price: '', free: false, href: '', image: '', published: true };

  function addIssue() {
    var d = data();
    var st = W.notePage.getState();
    var mi = (d.magazines || []).findIndex(function (m) { return m.id === st.year; });
    if (mi < 0) mi = 0;
    var mag = (d.magazines || [])[mi];
    if (!mag) return setStatus('追加先の年が見つかりません', 'err');

    var item = blankLike(mag.items[0], ISSUE_SHAPE);
    item.label = '新しい号';
    item.title = '（タイトルを入力してください）';
    mag.items.push(item);
    save(true);
    refresh();
    openEditor('magazines.' + mi + '.items.' + (mag.items.length - 1));
  }

  function addArticle(catId) {
    var d = data();
    var cat = (d.categories || []).filter(function (c) { return c.id === catId; })[0];
    if (!cat) return setStatus('追加先の分野が見つかりません', 'err');
    var ci = d.categories.indexOf(cat);
    var gi = Math.max(0, cat.groups.length - 1);
    if (!cat.groups.length) cat.groups.push({ title: '', subtitle: '', items: [] });
    var group = cat.groups[gi];

    var item = blankLike(group.items[0], ARTICLE_SHAPE);
    item.title = '（タイトルを入力してください）';
    group.items.push(item);
    save(true);
    refresh();
    openEditor('categories.' + ci + '.groups.' + gi + '.items.' + (group.items.length - 1));
  }

  /* ---------- 並べ替え・複製・削除 ---------- */

  function act(ref, action) {
    var r = resolve(ref);
    if (!r || !r.item) return setStatus('対象が見つかりません', 'err');
    var list = r.list;
    var i = r.index;

    if (action === 'up' && i > 0) list.splice(i - 1, 0, list.splice(i, 1)[0]);
    else if (action === 'down' && i < list.length - 1) list.splice(i + 1, 0, list.splice(i, 1)[0]);
    else if (action === 'dup') {
      var copy = JSON.parse(JSON.stringify(r.item));
      if (copy.id) copy.id = copy.id + '-copy';
      list.splice(i + 1, 0, copy);
    } else if (action === 'del') {
      var name = r.item.title || r.item.label || 'この項目';
      if (!confirm('「' + name + '」を削除します。よろしいですか？\n（下書きの破棄で元に戻せます）')) return;
      list.splice(i, 1);
    } else return;

    save(true);
    refresh();
    setStatus('更新しました', 'ok');
  }

  /* ---------- 編集ダイアログ ---------- */

  var dlg;

  function fieldHTML(key, value) {
    var label = LABELS[key] || key;
    var id = 'wa-f-' + key;

    if (typeof value === 'boolean') {
      return (
        '<div class="wa-field wa-inline">' +
        '<input id="' + id + '" type="checkbox" data-key="' + esc(key) + '"' + (value ? ' checked' : '') + '>' +
        '<label for="' + id + '">' + esc(label) + '</label>' +
        '</div>'
      );
    }

    var input = LONG.indexOf(key) > -1
      ? '<textarea id="' + id + '" data-key="' + esc(key) + '" rows="5">' + esc(value) + '</textarea>'
      : '<input id="' + id + '" type="text" data-key="' + esc(key) + '" value="' + esc(value) + '">';

    return (
      '<div class="wa-field">' +
      '<label for="' + id + '">' + esc(label) + '</label>' +
      input +
      (key === 'image' && value ? '<img class="wa-thumb" src="' + esc(W.url(value)) + '" alt="">' : '') +
      (key === 'href' ? '<p class="wa-help">note.com の記事URLなど。空欄にするとボタンが出ません。</p>' : '') +
      (key === 'price' ? '<p class="wa-help">「200円」「お試し0円」のように、表示したい文字をそのまま入れます。</p>' : '') +
      '</div>'
    );
  }

  function openEditor(ref) {
    var r = resolve(ref);
    if (!r || !r.item) return setStatus('対象が見つかりません', 'err');
    var item = r.item;

    var keys = Object.keys(item).sort(function (a, b) {
      var ia = ORDER.indexOf(a), ib = ORDER.indexOf(b);
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
    });

    if (!dlg) {
      dlg = document.createElement('dialog');
      dlg.id = 'wa-dialog';
      document.body.appendChild(dlg);
    }

    dlg.innerHTML =
      '<form method="dialog" class="wa-dlg">' +
      '<h2>' + (item.label || item.title ? esc(item.label || item.title).slice(0, 40) : '項目') + ' を編集</h2>' +
      '<p class="wa-ref">' + esc(ref) + '</p>' +
      '<div class="wa-fields">' + keys.map(function (k) { return fieldHTML(k, item[k]); }).join('') + '</div>' +
      '<div class="wa-dlg-actions">' +
      '<button type="button" class="wa-btn" data-close>閉じる</button>' +
      '<button type="button" class="wa-btn is-primary" data-apply>反映する</button>' +
      '</div>' +
      '</form>';

    dlg.showModal();

    dlg.querySelector('[data-close]').onclick = function () { dlg.close(); };
    dlg.querySelector('[data-apply]').onclick = function () {
      dlg.querySelectorAll('[data-key]').forEach(function (el) {
        var k = el.dataset.key;
        item[k] = el.type === 'checkbox' ? el.checked : el.value;
      });
      save(true);
      dlg.close();
      refresh();
      setStatus('反映しました', 'ok');
    };
  }

  /* ---------- 起動 ---------- */

  function start() {
    var s = session();
    if (!s) return;
    if (!W.notePage) return; // このページには編集対象がない

    buildBar(s.name ? s.name + ' としてログイン中' : '管理者としてログイン中');
    // 年を切り替えた直後にも操作ボタンを付け直す
    W.notePage.afterRender = decorate;
    decorate();

    document.addEventListener('click', function (e) {
      var btn = e.target.closest('.wa-tools button');
      if (!btn) return;
      e.preventDefault();
      e.stopPropagation();
      var ref = btn.closest('.wa-tools').dataset.for;
      if (btn.dataset.act === 'edit') openEditor(ref);
      else act(ref, btn.dataset.act);
    }, true);

    // 編集モード中は、カードのリンクで外部サイトへ飛ばないようにする
    document.addEventListener('click', function (e) {
      if (!W.editMode) return;
      var a = e.target.closest('a.issue, .article h3 a, .article .price');
      if (a) { e.preventDefault(); }
    }, true);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', waitForPage);
  else waitForPage();

  /* page-note.js の描画が終わってから割り込む */
  function waitForPage() {
    var tries = 0;
    (function poll() {
      if (W.notePage && W.notePage.getData()) return start();
      if (tries++ > 60) return;
      setTimeout(poll, 100);
    })();
  }
})(window);
