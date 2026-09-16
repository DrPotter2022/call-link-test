/* ============================================================
   管理画面のログイン
   パスワードは Googleスプレッドシートで管理し、一致した場合だけ画面を表示する。

   照合の方法は2通り（content/admin.json で設定）
   ・gasEndpoint … Apps Script のウェブアプリに照合させる（推奨）
                   パスワードはGoogle側でのみ突き合わせ、シートの中身は外に出ない
   ・sheetCsvUrl … 公開したCSVをブラウザで読んで突き合わせる（簡単だが、
                   URLを知っている人はパスワードを読めてしまう）

   ※ 静的サイトである以上、この画面はあくまで「入口の鍵」です。
     ブラウザの開発者ツールを使える人は表示自体は回避できます。
     実際にサイトを書き換えるにはGitHubのアクセストークンが別途必要で、
     そちらが本当の防御になっています。
   ============================================================ */
(function (global) {
  'use strict';

  var CONFIG_URL = '../content/admin.json';
  var SESSION_KEY = 'wize_admin_session';

  function $(s) { return document.querySelector(s); }

  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /* ---------- セッション ---------- */

  /* ログイン状態は localStorage に持つ。
     管理画面とサイト本体（.note の編集モード）は別タブで開くことが多く、
     sessionStorage だとタブごとに分かれて毎回入力させることになるため。
     期限を持たせ、ログアウトで消す。 */
  function readSession() {
    try {
      var raw = localStorage.getItem(SESSION_KEY);
      if (!raw) return null;
      var s = JSON.parse(raw);
      if (!s.until || Date.now() > s.until) {
        localStorage.removeItem(SESSION_KEY);
        return null;
      }
      return s;
    } catch (e) {
      return null;
    }
  }

  function writeSession(name, hours) {
    try {
      localStorage.setItem(
        SESSION_KEY,
        JSON.stringify({ name: name || '', until: Date.now() + (hours || 12) * 3600 * 1000 })
      );
    } catch (e) { /* 保存できなくてもログイン自体は続行する */ }
  }

  function logout() {
    try { localStorage.removeItem(SESSION_KEY); } catch (e) {}
    location.reload();
  }

  /* ---------- CSV ---------- */

  /** 引用符つきCSVを行×列に分解する */
  function parseCSV(text) {
    var rows = [];
    var row = [];
    var cell = '';
    var quoted = false;

    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (quoted) {
        if (ch === '"') {
          if (text[i + 1] === '"') { cell += '"'; i++; }
          else quoted = false;
        } else cell += ch;
        continue;
      }
      if (ch === '"') quoted = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\r') { /* 無視 */ }
      else if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (c) { return String(c).trim() !== ''; }); });
  }

  var FALSY = ['false', 'no', '0', '×', 'x', '無効', '停止'];

  /** 見出し行から「パスワード」「名前」「有効」の列位置を割り出す */
  function findColumns(header) {
    var idx = { pass: -1, name: -1, enabled: -1 };
    header.forEach(function (h, i) {
      var k = String(h).trim().toLowerCase();
      if (idx.pass < 0 && (k.indexOf('パスワード') > -1 || k.indexOf('password') > -1 || k === 'pw')) idx.pass = i;
      if (idx.name < 0 && (k.indexOf('名前') > -1 || k.indexOf('名称') > -1 || k.indexOf('name') > -1 || k.indexOf('担当') > -1)) idx.name = i;
      if (idx.enabled < 0 && (k.indexOf('有効') > -1 || k.indexOf('enabled') > -1 || k.indexOf('active') > -1)) idx.enabled = i;
    });
    // 見出しが読み取れない場合は「名前 / パスワード / 有効」の並びとみなす
    if (idx.pass < 0) { idx.name = 0; idx.pass = 1; idx.enabled = 2; return { cols: idx, headerIsData: true }; }
    return { cols: idx, headerIsData: false };
  }

  function verifyByCsv(url, password) {
    return fetch(url, { cache: 'no-store' })
      .then(function (r) {
        if (!r.ok) throw new Error('スプレッドシートを読み込めませんでした（HTTP ' + r.status + '）');
        return r.text();
      })
      .then(function (text) {
        if (/^\s*</.test(text)) {
          throw new Error('CSVではなくHTMLが返りました。「ウェブに公開」でCSV形式のURLになっているか確認してください。');
        }
        var rows = parseCSV(text);
        if (!rows.length) throw new Error('スプレッドシートが空です。');

        var found = findColumns(rows[0]);
        var cols = found.cols;
        var body = found.headerIsData ? rows : rows.slice(1);

        for (var i = 0; i < body.length; i++) {
          var r = body[i];
          var pass = String(r[cols.pass] == null ? '' : r[cols.pass]).trim();
          if (!pass || pass !== password) continue;
          if (cols.enabled > -1) {
            var en = String(r[cols.enabled] == null ? '' : r[cols.enabled]).trim().toLowerCase();
            if (en && FALSY.indexOf(en) > -1) continue;
          }
          return { ok: true, name: cols.name > -1 ? String(r[cols.name] || '').trim() : '' };
        }
        return { ok: false };
      });
  }

  function verifyByGas(url, password) {
    // text/plain にしておくと事前確認（プリフライト）が発生せず、Apps Script でも通る
    return fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: password,
    })
      .then(function (r) {
        if (!r.ok) throw new Error('照合サーバーに接続できませんでした（HTTP ' + r.status + '）');
        return r.text();
      })
      .then(function (text) {
        var j;
        try {
          j = JSON.parse(text);
        } catch (e) {
          // Apps Script 側で例外が起きるとHTMLのエラー画面が返る。理由を取り出して見せる。
          var m = text.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').match(/(エラー|Error)[^。]{0,160}/);
          throw new Error(
            'Apps Script でエラーが起きています。' +
            (m ? '\n' + m[0].trim() : '') +
            '\ntools/gas-auth.gs を貼り直して、デプロイを更新してください。'
          );
        }
        if (j.message && !j.ok) {
          // スクリプト側が理由を返してきた場合（シートが見つからない等）
          if (/見つかりません|接続できません|空です/.test(j.message)) throw new Error(j.message);
        }
        return { ok: !!j.ok, name: j.name || '' };
      })
      .catch(function (e) {
        if (e instanceof TypeError) {
          /* Apps Script は中でエラーが起きるとHTMLのエラー画面を返すが、
             そこにはCORSの許可が付かないため、ブラウザからは中身すら読めない。
             アクセス権の問題と見分けがつかないので、両方を案内する。 */
          throw new Error(
            '照合サーバーから返事がありません。次のどちらかです。\n' +
            '(1) ウェブアプリの「アクセスできるユーザー」が「全員」になっていない\n' +
            '(2) Apps Script の中でエラーが起きている\n' +
            'ウェブアプリのURLをブラウザで直接開くと、どちらか分かります。'
          );
        }
        throw e;
      });
  }

  function verify(auth, password) {
    if (auth.gasEndpoint) return verifyByGas(auth.gasEndpoint, password);
    if (auth.sheetCsvUrl) return verifyByCsv(auth.sheetCsvUrl, password);
    return Promise.reject(new Error('unconfigured'));
  }

  /* ---------- ログイン画面 ---------- */

  function lockScreen(auth) {
    return new Promise(function (resolve) {
      var configured = !!(auth.gasEndpoint || auth.sheetCsvUrl);

      var el = document.createElement('div');
      el.className = 'lock';
      el.innerHTML =
        '<div class="lock-panel">' +
        '<p class="lock-eyebrow">WIZE — CONTENT EDITOR</p>' +
        '<h1>管理画面</h1>' +
        (configured
          ? '<p class="lock-lead">パスワードを入力してください。</p>' +
            '<form id="lock-form" autocomplete="off">' +
            '<input id="lock-pw" type="password" placeholder="パスワード" autocomplete="current-password" aria-label="パスワード">' +
            '<button class="btn primary" id="lock-go" type="submit">開く</button>' +
            '</form>' +
            '<p class="lock-msg" id="lock-msg" role="status"></p>' +
            (auth.hint ? '<p class="lock-hint">' + esc(auth.hint) + '</p>' : '')
          : '<p class="lock-lead">パスワードの参照先がまだ設定されていません。</p>' +
            '<div class="lock-setup">' +
            '<p>Googleスプレッドシートでパスワードを管理する設定が必要です。' +
            '<code>content/admin.json</code> の <code>gasEndpoint</code> または <code>sheetCsvUrl</code> にURLを入れてください。</p>' +
            '<p>手順は <code>README.md</code> の「管理画面にパスワードをかける」にあります。</p>' +
            '<p>設定が終わるまで管理画面を使いたい場合は、<code>content/admin.json</code> の ' +
            '<code>"enabled"</code> を <code>false</code> にしてください。</p>' +
            '</div>') +
        '</div>';

      document.body.appendChild(el);
      document.body.classList.add('is-locked');

      if (!configured) return; // 解決しないので、この画面のまま止まる

      var form = el.querySelector('#lock-form');
      var input = el.querySelector('#lock-pw');
      var btn = el.querySelector('#lock-go');
      var msg = el.querySelector('#lock-msg');
      input.focus();

      form.addEventListener('submit', function (e) {
        e.preventDefault();
        var pw = input.value;
        if (!pw) { msg.textContent = 'パスワードを入力してください。'; return; }

        btn.disabled = true;
        msg.className = 'lock-msg';
        msg.textContent = '照合中…';

        verify(auth, pw)
          .then(function (r) {
            if (!r.ok) {
              msg.className = 'lock-msg is-error';
              msg.textContent = 'パスワードが一致しません。';
              input.select();
              return;
            }
            writeSession(r.name, auth.sessionHours);
            el.remove();
            document.body.classList.remove('is-locked');
            resolve(r);
          })
          .catch(function (err) {
            msg.className = 'lock-msg is-error';
            msg.textContent = err.message === 'unconfigured' ? 'パスワードの参照先が設定されていません。' : err.message;
          })
          .then(function () {
            btn.disabled = false;
          });
      });
    });
  }

  /* ---------- 公開API ---------- */

  /** 認証が済むまで解決しない Promise を返す */
  function require() {
    return fetch(CONFIG_URL, { cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : {}; })
      .catch(function () { return {}; })
      .then(function (cfg) {
        var auth = (cfg && cfg.auth) || {};
        if (auth.enabled === false) {
          /* パスワードを使わない設定のときも印は残す。
             こうしないと /note/ の編集モードが管理者だと判断できない。 */
          writeSession('', auth.sessionHours);
          return { ok: true, name: '', skipped: true };
        }

        var s = readSession();
        if (s) return { ok: true, name: s.name, resumed: true };

        return lockScreen(auth);
      });
  }

  global.WIZE_AUTH = { require: require, logout: logout, readSession: readSession };
})(window);
