/* ============================================================
   WIZE 管理画面
   content/*.json をフォームで編集し、
   ・下書き … localStorage（プレビューに即反映）
   ・書き出し … JSONファイルをダウンロード
   ・公開 … GitHub Contents API でコミット
   ============================================================ */
(function () {
  'use strict';

  var FILES = [
    { key: 'site', label: '共通設定' },
    { key: 'home', label: 'トップ' },
    { key: 'note', label: '.note（記事）' },
    { key: 'seminar', label: 'セミナー' },
  ];

  var DRAFT_PREFIX = 'wize_draft_';
  var GH_KEY = 'wize_github_config';

  /* 項目名の日本語ラベル。無いものはキーをそのまま表示する */
  var LABELS = {
    brand: 'ブランド', nameJa: '表示名（日本語）', nameEn: '表示名（英字）', concept: 'コンセプト',
    nav: 'ヘッダーのメニュー', navCta: 'ヘッダーのボタン', footer: 'フッター', links: 'リンク',
    companyLabel: '運営情報の見出し', company: '会社情報', name: '事業者名', representative: '代表者',
    address: '所在地', email: 'メールアドレス', tel: '電話番号', hours: '営業時間',
    legal: '注意書き（免責）', copyright: 'コピーライト',
    meta: 'ページ情報（検索結果に出る文言）', title: 'タイトル', description: '説明文',
    hero: '冒頭（ヒーロー）', eyebrow: '小見出し（英字ラベル）', tate: '縦書きの一言',
    body: '本文', lead: 'リード文', note: '補足', summary: '本文',
    primaryCta: '主ボタン', secondaryCta: '副ボタン', ctaLabel: 'ボタンの文言', ctaHref: 'ボタンのリンク先',
    label: '表示名', href: 'リンク先', id: 'ID（変更しないでください）',
    hankoMain: '印章の文字', hankoSub: '印章の下の文字',
    problem: '問題提起', cardTag: 'カードのラベル', cardTitle: 'カードの見出し', cardPoints: 'カードの箇条書き',
    creed: 'ミッション', mark: '罫線上のラベル',
    contents: '提供しているもの', cards: 'カード', tag: 'ラベル', focus: '強調表示する',
    business: '事業内容', items: '項目', kanji: '漢字1文字', for: '対象',
    recruit: '同業の方へ', status: '募集状況', enabled: 'このセクションを表示する',
    cta: '最後の行動喚起（CTA）',
    caution: '留意事項', priceLegend: '価格の説明', categories: '分野', groups: 'グループ',
    magazines: '年別バックナンバー', subscriptions: 'マガジン・定期購読',
    date: '執筆日', era: '時代・テーマの見出し', eraNote: '時代の補足', price: '価格の表示',
    free: '無料として扱う（朱色で表示）', image: '画像のパス', published: '公開する',
    courses: 'コース', lessons: 'レッスン', lectures: '講座', lecturesTitle: '講座一覧の見出し',
    code: '講座番号', duration: '収録時間', message: '講師メッセージ', voices: '受講者の声',
    subtitle: 'サブタイトル', _comment: 'メモ（サイトには表示されません）',
  };

  /* 一覧の見出しに使う候補 */
  var TITLE_KEYS = ['title', 'label', 'name', 'code', 'tag'];

  var state = {
    current: 'site',
    published: {},   // 公開中の内容（比較用）
    working: {},     // 編集中の内容
    dirty: {},
  };

  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function labelOf(key) {
    return LABELS[key] || key;
  }

  function status(msg, kind) {
    var el = $('#status');
    el.textContent = msg || '';
    el.className = 'status' + (kind ? ' ' + kind : '');
    if (msg && kind === 'ok') setTimeout(function () { if (el.textContent === msg) el.textContent = ''; }, 4000);
  }

  function clone(v) {
    return JSON.parse(JSON.stringify(v));
  }

  /* ---------- 読み込み ---------- */

  function loadFile(key) {
    return fetch('../content/' + key + '.json', { cache: 'no-store' })
      .then(function (r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.json();
      });
  }

  function loadAll() {
    return Promise.all(
      FILES.map(function (f) {
        return loadFile(f.key).then(function (json) {
          state.published[f.key] = json;
          var draft = null;
          try {
            var raw = localStorage.getItem(DRAFT_PREFIX + f.key);
            if (raw) draft = JSON.parse(raw);
          } catch (e) { /* 壊れた下書きは無視する */ }
          state.working[f.key] = draft || clone(json);
          state.dirty[f.key] = !!draft;
        });
      })
    );
  }

  function saveDraft(key) {
    try {
      localStorage.setItem(DRAFT_PREFIX + key, JSON.stringify(state.working[key]));
      state.dirty[key] = true;
      renderTabs();
      status('下書きを保存しました', 'ok');
    } catch (e) {
      status('下書きを保存できません（容量超過の可能性）', 'err');
    }
  }

  /* ---------- 値の読み書き（パス指定） ---------- */

  function getAt(obj, path) {
    return path.reduce(function (o, k) { return o == null ? o : o[k]; }, obj);
  }

  function setAt(obj, path, value) {
    var last = path[path.length - 1];
    var parent = path.slice(0, -1).reduce(function (o, k) { return o[k]; }, obj);
    parent[last] = value;
  }

  /* ---------- フォームの組み立て ---------- */

  var uid = 0;
  function nextId() { return 'f' + (++uid); }

  function isLongText(key, value) {
    return String(value).length > 60 || String(value).indexOf('\n') > -1 ||
      ['body', 'lead', 'summary', 'legal', 'description', 'note'].indexOf(key) > -1;
  }

  function titleOf(item, index) {
    for (var i = 0; i < TITLE_KEYS.length; i++) {
      var v = item[TITLE_KEYS[i]];
      if (typeof v === 'string' && v.trim()) return v.replace(/\n/g, ' ').slice(0, 70);
    }
    return '項目 ' + (index + 1);
  }

  /** 値の型に応じてフォーム部品を返す */
  function field(key, value, path) {
    var id = nextId();
    var lbl = '<label for="' + id + '">' + esc(labelOf(key)) + '<span class="key">' + esc(key) + '</span></label>';

    if (typeof value === 'boolean') {
      return (
        '<div class="field inline" data-path="' + esc(JSON.stringify(path)) + '">' +
        '<input id="' + id + '" type="checkbox" data-type="boolean"' + (value ? ' checked' : '') + '>' +
        '<label for="' + id + '">' + esc(labelOf(key)) + '</label>' +
        '</div>'
      );
    }

    if (typeof value === 'number') {
      return (
        '<div class="field" data-path="' + esc(JSON.stringify(path)) + '">' + lbl +
        '<input id="' + id + '" type="number" data-type="number" value="' + esc(value) + '">' +
        '</div>'
      );
    }

    if (typeof value === 'string') {
      var isImage = key === 'image' && value;
      var input = isLongText(key, value)
        ? '<textarea id="' + id + '" data-type="string" rows="4">' + esc(value) + '</textarea>'
        : '<input id="' + id + '" type="text" data-type="string" value="' + esc(value) + '">';
      return (
        '<div class="field" data-path="' + esc(JSON.stringify(path)) + '">' + lbl + input +
        (isImage ? '<img class="thumb" style="margin-top:8px;width:120px;height:80px" src="' + esc('../' + value) + '" alt="">' : '') +
        (key === 'href' ? '<p class="help">「note/」「seminar/」のようにサイト内の場所、または https:// から始まるURLを入力します。空欄にするとボタンは表示されません。</p>' : '') +
        '</div>'
      );
    }

    if (Array.isArray(value)) {
      if (!value.length || typeof value[0] !== 'object') return stringArray(key, value, path);
      return objectArray(key, value, path);
    }

    if (value && typeof value === 'object') {
      return (
        '<fieldset data-group="' + esc(key) + '"><legend>' + esc(labelOf(key)) + '</legend>' +
        Object.keys(value).map(function (k) { return field(k, value[k], path.concat(k)); }).join('') +
        '</fieldset>'
      );
    }

    return '';
  }

  function stringArray(key, arr, path) {
    return (
      '<div class="field"><label>' + esc(labelOf(key)) + '<span class="key">' + esc(key) + '</span></label>' +
      '<div class="list" data-array="' + esc(JSON.stringify(path)) + '" data-kind="string">' +
      '<div class="rows">' +
      arr.map(function (v, i) {
        return (
          '<div class="row-item" data-index="' + i + '">' +
          (String(v).length > 50
            ? '<textarea data-type="string" data-path="' + esc(JSON.stringify(path.concat(i))) + '" rows="2">' + esc(v) + '</textarea>'
            : '<input type="text" data-type="string" data-path="' + esc(JSON.stringify(path.concat(i))) + '" value="' + esc(v) + '">') +
          '<button type="button" class="btn small ghost move-up" title="上へ">↑</button>' +
          '<button type="button" class="btn small ghost move-down" title="下へ">↓</button>' +
          '<button type="button" class="btn small danger remove">削除</button>' +
          '</div>'
        );
      }).join('') +
      '</div>' +
      '<button type="button" class="btn small ghost list-add add-string" style="border-color:var(--ai-line);color:var(--ai)">＋ 追加</button>' +
      '</div></div>'
    );
  }

  function objectArray(key, arr, path) {
    return (
      '<div class="field"><label>' + esc(labelOf(key)) + '<span class="key">' + esc(key) + '（' + arr.length + '件）</span></label>' +
      '<div class="list" data-array="' + esc(JSON.stringify(path)) + '" data-kind="object">' +
      arr.map(function (item, i) {
        var hidden = item.published === false;
        return (
          '<details class="list-item" data-index="' + i + '">' +
          '<summary>' +
          '<span class="idx">' + String(i + 1).padStart(2, '0') + '</span>' +
          (item.image ? '<img class="thumb" src="' + esc('../' + item.image) + '" alt="" loading="lazy">' : '') +
          '<span class="ttl">' + esc(titleOf(item, i)) + '</span>' +
          (hidden ? '<span class="off">非公開</span>' : '') +
          '</summary>' +
          '<div class="list-body">' +
          Object.keys(item).map(function (k) { return field(k, item[k], path.concat(i, k)); }).join('') +
          '</div>' +
          '<div class="item-tools">' +
          '<button type="button" class="move-up">↑ 上へ</button>' +
          '<button type="button" class="move-down">↓ 下へ</button>' +
          '<button type="button" class="duplicate">複製</button>' +
          '<button type="button" class="del remove">削除</button>' +
          '</div>' +
          '</details>'
        );
      }).join('') +
      '<button type="button" class="btn small ghost list-add add-object" style="border-color:var(--ai-line);color:var(--ai)">＋ 追加</button>' +
      '</div></div>'
    );
  }

  /* ---------- 描画 ---------- */

  function renderTabs() {
    $('#tabs').innerHTML = FILES.map(function (f) {
      return (
        '<button type="button" role="tab" data-key="' + f.key + '" aria-selected="' + (f.key === state.current) + '">' +
        esc(f.label) + (state.dirty[f.key] ? '<span class="badge">● 未公開</span>' : '') +
        '</button>'
      );
    }).join('');
  }

  function renderForm() {
    uid = 0;
    var data = state.working[state.current];
    var form = $('#form');
    form.innerHTML = Object.keys(data)
      .map(function (k) { return field(k, data[k], [k]); })
      .join('');
    renderOutline();
  }

  function renderOutline() {
    var groups = $$('#form > fieldset, #form > .field');
    $('#outline').innerHTML = groups
      .map(function (el, i) {
        var lg = el.querySelector(':scope > legend') || el.querySelector(':scope > label');
        if (!lg) return '';
        el.id = el.id || 'sec' + i;
        return '<a href="#' + el.id + '" data-target="' + el.id + '">' + esc(lg.textContent.replace(/\s+/g, ' ').trim()) + '</a>';
      })
      .join('');
  }

  /* ---------- イベント ---------- */

  function pathOf(el) {
    var holder = el.dataset.path ? el : el.closest('[data-path]');
    return holder ? JSON.parse(holder.dataset.path) : null;
  }

  function onInput(e) {
    var el = e.target;
    if (!el.dataset.type) return;
    var path = pathOf(el);
    if (!path) return;
    var v;
    if (el.dataset.type === 'boolean') v = el.checked;
    else if (el.dataset.type === 'number') v = Number(el.value);
    else v = el.value;
    setAt(state.working[state.current], path, v);
    scheduleSave();
  }

  var saveTimer;
  function scheduleSave() {
    clearTimeout(saveTimer);
    status('編集中…');
    saveTimer = setTimeout(function () { saveDraft(state.current); }, 700);
  }

  /** 配列の要素をひな形から作る（同じ配列の1件目をコピーして空にする） */
  function blankLike(sample) {
    if (!sample) return {};
    var out = {};
    Object.keys(sample).forEach(function (k) {
      var v = sample[k];
      if (typeof v === 'string') out[k] = k === 'id' ? 'new-' + Date.now().toString(36) : '';
      else if (typeof v === 'boolean') out[k] = k === 'published' ? true : false;
      else if (typeof v === 'number') out[k] = 0;
      else if (Array.isArray(v)) out[k] = [];
      else if (v && typeof v === 'object') out[k] = blankLike(v);
      else out[k] = null;
    });
    return out;
  }

  function onClick(e) {
    var btn = e.target.closest('button');
    if (!btn) return;
    var list = btn.closest('.list');
    if (!list) return;
    var path = JSON.parse(list.dataset.array);
    var arr = getAt(state.working[state.current], path);
    var holder = btn.closest('.list-item, .row-item');
    var idx = holder ? Number(holder.dataset.index) : -1;

    if (btn.classList.contains('add-string')) {
      arr.push('');
    } else if (btn.classList.contains('add-object')) {
      arr.push(blankLike(arr[0]));
    } else if (btn.classList.contains('move-up') && idx > 0) {
      arr.splice(idx - 1, 0, arr.splice(idx, 1)[0]);
    } else if (btn.classList.contains('move-down') && idx > -1 && idx < arr.length - 1) {
      arr.splice(idx + 1, 0, arr.splice(idx, 1)[0]);
    } else if (btn.classList.contains('duplicate') && idx > -1) {
      var copy = clone(arr[idx]);
      if (copy.id) copy.id = copy.id + '-copy';
      arr.splice(idx + 1, 0, copy);
    } else if (btn.classList.contains('remove') && idx > -1) {
      if (!confirm('この項目を削除します。よろしいですか？')) return;
      arr.splice(idx, 1);
    } else {
      return;
    }

    e.preventDefault();
    saveDraft(state.current);
    renderForm();
  }

  /* ---------- 書き出し ---------- */

  function download() {
    var changed = FILES.filter(function (f) { return state.dirty[f.key]; });
    if (!changed.length) return status('変更はありません', 'ok');

    // 1ファイルずつ落とすとブラウザに「複数ダウンロード」として止められるため、
    // まとめてZIPにして1回のダウンロードにする
    var files = changed.map(function (f) {
      return { name: 'content/' + f.key + '.json', text: JSON.stringify(state.working[f.key], null, 2) + '\n' };
    });
    var blob = window.WIZE_EXPORT.zip(files);
    var url = URL.createObjectURL(blob);
    var stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    window.WIZE_EXPORT.tryAutoSave(url, 'wize-content-' + stamp + '.zip');
    setTimeout(function () { URL.revokeObjectURL(url); }, 30000);
    status(changed.length + ' 件をZIPで書き出しました', 'ok');
  }

  /* ---------- 絞り込み ---------- */

  /** この項目が検索語に当たるか。見出しに加えて、入力欄の中身（URL等）も見る */
  function itemMatches(li, q) {
    var head = li.querySelector(':scope > summary .ttl');
    if (head && head.textContent.indexOf(q) > -1) return true;

    var inputs = li.querySelectorAll('input[type="text"], textarea');
    for (var i = 0; i < inputs.length; i++) {
      // 入れ子になった下位項目の欄は、その項目自身の当たりとして数える
      if (inputs[i].closest('.list-item') !== li) continue;
      if (String(inputs[i].value).indexOf(q) > -1) return true;
    }
    return false;
  }

  /**
   * 項目一覧（.note の記事など）を検索語で絞り込む。
   * 記事は 分野 → グループ → 項目 と3階層あり、既定では閉じているため、
   * 当たった項目は上位ごと開いて、その場で編集できるようにする。
   */
  function filterEverything(q) {
    var items = $$('#form .list-item');

    // 目次側
    $$('#outline a').forEach(function (a) {
      a.style.display = !q || a.textContent.indexOf(q) > -1 ? '' : 'none';
    });

    if (!q) {
      items.forEach(function (li) {
        li.classList.remove('is-filtered-out', 'is-hit');
        li.open = false;
      });
      $$('#form fieldset').forEach(function (f) { f.classList.remove('is-filtered-out'); });
      setCount('');
      return;
    }

    // まず自分自身が当たるかを判定
    var hits = 0;
    items.forEach(function (li) {
      var hit = itemMatches(li, q);
      li.classList.toggle('is-hit', hit);
      if (hit) hits++;
    });

    // 子孫に当たりがあれば、親も残して開く
    items.forEach(function (li) {
      var keep = li.classList.contains('is-hit') || li.querySelector('.list-item.is-hit');
      li.classList.toggle('is-filtered-out', !keep);
      li.open = keep;
    });

    // 中身が全部消えたまとまりは隠す
    $$('#form fieldset').forEach(function (f) {
      var hasList = f.querySelector('.list-item');
      if (!hasList) return;
      f.classList.toggle('is-filtered-out', !f.querySelector('.list-item.is-hit'));
    });

    setCount(hits + ' 件');
  }

  function setCount(text) {
    var el = $('#jump-count');
    if (el) el.textContent = text;
  }

  /* ---------- HTML 書き出し ---------- */

  var lastExportUrl = null;

  function openHTMLDialog() {
    $('#html-warn').textContent = '';
    $('#html-go').disabled = false;
    clearExportResult();
    $('#html-dialog').showModal();
  }

  function clearExportResult() {
    var box = $('#html-result');
    box.innerHTML = '';
    box.hidden = true;
    if (lastExportUrl) {
      URL.revokeObjectURL(lastExportUrl);
      lastExportUrl = null;
    }
  }

  function showExportResult(res) {
    var box = $('#html-result');
    lastExportUrl = URL.createObjectURL(res.blob);

    var mb = (res.blob.size / 1048576).toFixed(res.blob.size > 1048576 ? 1 : 2);
    box.hidden = false;
    box.innerHTML =
      '<p class="done">' + res.count + ' ページ（' + mb + 'MB）の書き出しが終わりました。</p>' +
      '<a class="btn primary dl" id="html-dl">ZIPをダウンロード</a>' +
      '<ul class="filelist">' +
      res.files.map(function (f) { return '<li>' + esc(f.name) + '</li>'; }).join('') +
      '</ul>' +
      '<p class="tip">自動で保存されなかった場合は、上のボタンを押してください。' +
      'それでも保存できないときは、ブラウザのダウンロード設定（自動ダウンロードのブロック）をご確認ください。</p>';

    var dl = $('#html-dl');
    dl.href = lastExportUrl;
    dl.setAttribute('download', res.filename);

    // 自動保存も試みる（ユーザー操作から時間が経つと無視するブラウザがある）
    window.WIZE_EXPORT.tryAutoSave(lastExportUrl, res.filename, $('#html-dialog'));
  }

  function runHTMLExport() {
    var mode = $('input[name="html-mode"]:checked').value;
    var pages = $$('.html-page:checked').map(function (c) { return c.value; });
    var warn = $('#html-warn');

    if (!pages.length) {
      warn.textContent = '書き出すページを1つ以上選んでください。';
      return;
    }

    clearExportResult();
    $('#html-go').disabled = true;
    warn.textContent = '準備中…';

    window.WIZE_EXPORT.exportHTML({
      standalone: mode === 'standalone',
      pages: pages,
      getData: function (key) { return state.working[key]; },
      onProgress: function (msg) {
        warn.textContent = msg;
        status(msg);
      },
    })
      .then(function (res) {
        warn.textContent = '';
        showExportResult(res);
        status(res.count + ' ページを書き出しました', 'ok');
      })
      .catch(function (e) {
        warn.textContent = '書き出しに失敗しました：' + (e.message || e);
        status('書き出しに失敗しました', 'err');
      })
      .then(function () {
        $('#html-go').disabled = false;
      });
  }

  /* ---------- GitHub 公開 ---------- */

  function b64(str) {
    var bytes = new TextEncoder().encode(str);
    var bin = '';
    bytes.forEach(function (b) { bin += String.fromCharCode(b); });
    return btoa(bin);
  }

  function ghConfig() {
    try { return JSON.parse(localStorage.getItem(GH_KEY) || '{}'); } catch (e) { return {}; }
  }

  function openPublish() {
    var changed = FILES.filter(function (f) { return state.dirty[f.key]; });
    var cfg = ghConfig();
    $('#gh-repo').value = cfg.repo || '';
    $('#gh-branch').value = cfg.branch || 'main';
    $('#gh-token').value = cfg.token || '';
    $('#gh-remember').checked = !!cfg.token;
    $('#gh-warn').textContent = changed.length
      ? '公開する対象：' + changed.map(function (f) { return f.label; }).join('、')
      : '未公開の変更はありません。';
    $('#gh-go').disabled = !changed.length;
    $('#publish-dialog').showModal();
  }

  async function ghPut(repo, branch, token, filePath, content, message) {
    var api = 'https://api.github.com/repos/' + repo + '/contents/' + filePath;
    var headers = {
      Authorization: 'Bearer ' + token,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
    };

    var sha;
    var head = await fetch(api + '?ref=' + encodeURIComponent(branch), { headers: headers });
    if (head.ok) sha = (await head.json()).sha;
    else if (head.status !== 404) throw new Error(filePath + ' の取得に失敗（HTTP ' + head.status + '）');

    var res = await fetch(api, {
      method: 'PUT',
      headers: Object.assign({ 'Content-Type': 'application/json' }, headers),
      body: JSON.stringify({ message: message, content: b64(content), branch: branch, sha: sha }),
    });
    if (!res.ok) {
      var body = await res.text();
      throw new Error(filePath + ' の書き込みに失敗（HTTP ' + res.status + '）' + body.slice(0, 200));
    }
  }

  async function publish() {
    var repo = $('#gh-repo').value.trim();
    var branch = $('#gh-branch').value.trim() || 'main';
    var token = $('#gh-token').value.trim();
    var message = $('#gh-message').value.trim() || 'サイト掲載内容の更新';
    var warn = $('#gh-warn');

    if (!/^[\w.-]+\/[\w.-]+$/.test(repo)) return (warn.textContent = 'リポジトリは owner/repo の形式で入力してください。');
    if (!token) return (warn.textContent = 'アクセストークンを入力してください。');

    if ($('#gh-remember').checked) {
      localStorage.setItem(GH_KEY, JSON.stringify({ repo: repo, branch: branch, token: token }));
    } else {
      localStorage.setItem(GH_KEY, JSON.stringify({ repo: repo, branch: branch }));
    }

    var changed = FILES.filter(function (f) { return state.dirty[f.key]; });
    $('#gh-go').disabled = true;
    warn.textContent = '公開中…';

    try {
      for (var i = 0; i < changed.length; i++) {
        var f = changed[i];
        warn.textContent = '公開中… (' + (i + 1) + '/' + changed.length + ') ' + f.label;
        await ghPut(repo, branch, token, 'content/' + f.key + '.json', JSON.stringify(state.working[f.key], null, 2) + '\n', message + '：' + f.label);
        state.published[f.key] = clone(state.working[f.key]);
        localStorage.removeItem(DRAFT_PREFIX + f.key);
        state.dirty[f.key] = false;
      }
      renderTabs();
      $('#publish-dialog').close();
      status('公開しました。GitHub Pages の反映まで数十秒かかります', 'ok');
    } catch (e) {
      warn.textContent = String(e.message || e);
    } finally {
      $('#gh-go').disabled = false;
    }
  }

  /* ---------- 起動 ---------- */

  function switchTo(key) {
    state.current = key;
    renderTabs();
    renderForm();
    window.scrollTo(0, 0);
  }

  function init() {
    renderTabs();
    renderForm();

    $('#tabs').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-key]');
      if (b) switchTo(b.dataset.key);
    });

    var form = $('#form');
    form.addEventListener('input', onInput);
    form.addEventListener('change', onInput);
    form.addEventListener('click', onClick);

    var jumpTimer;
    $('#jump').addEventListener('input', function () {
      var q = this.value.trim();
      clearTimeout(jumpTimer);
      jumpTimer = setTimeout(function () { filterEverything(q); }, 120);
    });

    $('#outline').addEventListener('click', function (e) {
      var a = e.target.closest('a[data-target]');
      if (!a) return;
      e.preventDefault();
      var t = document.getElementById(a.dataset.target);
      if (t) t.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });

    $('#btn-download').addEventListener('click', download);
    $('#btn-html').addEventListener('click', openHTMLDialog);
    $('#html-go').addEventListener('click', runHTMLExport);
    $('#btn-publish').addEventListener('click', openPublish);
    $('#gh-go').addEventListener('click', publish);

    $('#btn-reload').addEventListener('click', function () {
      if (!confirm('このタブの下書きを破棄し、公開中の内容に戻します。よろしいですか？')) return;
      localStorage.removeItem(DRAFT_PREFIX + state.current);
      state.working[state.current] = clone(state.published[state.current]);
      state.dirty[state.current] = false;
      renderTabs();
      renderForm();
      status('公開中の内容に戻しました', 'ok');
    });
  }

  window.WIZE_AUTH.require()
    .then(function (session) {
      if (session && !session.skipped) {
        var who = session.name ? session.name + ' でログイン中' : 'ログイン中';
        var bar = $('.bar-actions');
        var out = document.createElement('button');
        out.type = 'button';
        out.className = 'btn ghost';
        out.textContent = 'ログアウト';
        out.title = who;
        out.addEventListener('click', function () { window.WIZE_AUTH.logout(); });
        bar.appendChild(out);
      }
      return loadAll();
    })
    .then(init)
    .catch(function (e) {
      $('#form').innerHTML =
        '<fieldset><legend>読み込みエラー</legend><p class="hint">content/*.json を読み込めませんでした：' +
        esc(e.message) +
        '</p><p class="hint">管理画面はローカルファイル（file://）では動きません。GitHub Pages 上か、ローカルの簡易サーバー経由で開いてください。</p></fieldset>';
    });
})();
