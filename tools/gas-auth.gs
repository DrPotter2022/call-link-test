/**
 * WIZE 管理画面 — パスワード照合用の Google Apps Script
 *
 * スプレッドシートでパスワードを管理し、照合だけをGoogle側で行う。
 * 管理画面へはパスワードそのものを渡さないので、シートの中身が外部に漏れない。
 *
 * ── 設置手順 ─────────────────────────────
 * 1. パスワード管理用のスプレッドシートを作り、1行目を見出しにする
 *
 *      A列: 名前        B列: パスワード      C列: 有効
 *      和田            ********           TRUE
 *      スタッフ         ********           TRUE
 *
 *    （見出しは「名前」「パスワード」「有効」を含んでいれば順不同で構いません）
 *
 * 2. そのスプレッドシートで 拡張機能 → Apps Script を開く
 * 3. エディタの中身をすべて消してから、このファイルを「まるごと」貼り付ける
 *    ★ 先頭の var SHEET_NAME / var SPREADSHEET_ID の行まで必ず含めてください
 * 4. 右上の「デプロイ」→「新しいデプロイ」
 *      種類            ウェブアプリ
 *      次のユーザーとして実行   自分
 *      アクセスできるユーザー   全員          ← ここが「全員」でないと管理画面から届きません
 * 5. 発行された「ウェブアプリのURL」を、content/admin.json の gasEndpoint に貼る
 *
 * コードを直したら、そのつど「デプロイ」→「デプロイを管理」→ 鉛筆 →
 * バージョン「新バージョン」→「デプロイ」で反映してください。URLは変わりません。
 *
 * 設置できたか確かめるには、ウェブアプリのURLをそのままブラウザで開きます。
 *   {"ok":false,"message":"WIZE admin auth endpoint", ...}
 * と表示されれば動いています（sheet の欄にシート名が出ます）。
 * ───────────────────────────────────────
 */

/** 参照するシート名。空にすると、いちばん左のシートを使います。 */
var SHEET_NAME = '';

/** スプレッドシートに紐づかない「スタンドアロン」で作った場合だけ、シートのIDを入れます。
 *  （URLの /d/ と /edit の間の長い文字列。通常は空のままで構いません） */
var SPREADSHEET_ID = '';

/* ============================================================
   ここから下は、通常は触る必要がありません
   ============================================================ */

/** 管理画面からの照合リクエスト（本文にパスワードが入っている） */
function doPost(e) {
  try {
    var password = e && e.postData ? String(e.postData.contents || '') : '';
    return respond(check(password));
  } catch (err) {
    // 何かあってもHTMLのエラー画面ではなくJSONで返す（管理画面側で理由が見えるように）
    return respond({ ok: false, message: String(err && err.message ? err.message : err) });
  }
}

/** ブラウザで直接URLを開いたときの確認用。パスワードは受け付けない。 */
function doGet() {
  var info = { ok: false, message: 'WIZE admin auth endpoint' };
  try {
    var sheet = getSheet();
    info.sheet = sheet.getName();
    info.rows = Math.max(0, sheet.getLastRow() - 1);
  } catch (err) {
    info.error = String(err && err.message ? err.message : err);
  }
  return respond(info);
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 設定に応じてシートを取り出す。指定が無ければ、いちばん左のシート。 */
function getSheet() {
  var id = typeof SPREADSHEET_ID === 'string' ? SPREADSHEET_ID : '';
  var book = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();

  if (!book) {
    throw new Error('スプレッドシートに接続できません。スプレッドシートの「拡張機能 → Apps Script」から作るか、SPREADSHEET_ID を設定してください。');
  }

  var name = typeof SHEET_NAME === 'string' ? SHEET_NAME : '';
  var sheet = name ? book.getSheetByName(name) : book.getSheets()[0];

  if (!sheet) {
    var names = book.getSheets().map(function (s) { return s.getName(); }).join('、');
    throw new Error('シート「' + name + '」が見つかりません。このファイルにあるのは：' + names);
  }
  return sheet;
}

function check(password) {
  if (!password) return { ok: false };

  var values = getSheet().getDataRange().getValues();
  if (!values.length) return { ok: false, message: 'シートが空です' };

  var cols = findColumns(values[0]);
  var rows = cols.headerIsData ? values : values.slice(1);

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var pass = String(row[cols.pass] == null ? '' : row[cols.pass]).trim();
    if (!pass || pass !== password) continue;

    if (cols.enabled > -1) {
      var en = String(row[cols.enabled] == null ? '' : row[cols.enabled]).trim().toLowerCase();
      if (en && ['false', 'no', '0', '×', 'x', '無効', '停止'].indexOf(en) > -1) continue;
    }
    return { ok: true, name: cols.name > -1 ? String(row[cols.name] || '').trim() : '' };
  }
  return { ok: false };
}

/** 見出し行から各列の位置を割り出す。見出しが無ければ 名前/パスワード/有効 の並びとみなす */
function findColumns(header) {
  var idx = { pass: -1, name: -1, enabled: -1 };

  header.forEach(function (h, i) {
    var k = String(h).trim().toLowerCase();
    if (idx.pass < 0 && (k.indexOf('パスワード') > -1 || k.indexOf('password') > -1 || k === 'pw')) idx.pass = i;
    if (idx.name < 0 && (k.indexOf('名前') > -1 || k.indexOf('名称') > -1 || k.indexOf('name') > -1 || k.indexOf('担当') > -1)) idx.name = i;
    if (idx.enabled < 0 && (k.indexOf('有効') > -1 || k.indexOf('enabled') > -1 || k.indexOf('active') > -1)) idx.enabled = i;
  });

  if (idx.pass < 0) return { pass: 1, name: 0, enabled: 2, headerIsData: true };
  idx.headerIsData = false;
  return idx;
}

/** 設置確認用。エディタ上で実行するとログに結果が出る。 */
function testCheck() {
  Logger.log(getSheet().getName());
  Logger.log(check('ここにテストしたいパスワード'));
}
