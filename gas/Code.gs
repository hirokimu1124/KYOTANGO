/**
 * 画面モニター メッセージ記録用（スプレッドシートにバインドして使う）
 *
 * 使い方：
 * 1. 新しいスプレッドシートを作る →「拡張機能」→「Apps Script」→ このコードを貼り付けて保存
 * 2. 関数「setup」を選んで実行（初回は権限の許可が必要）→ 実行ログに合言葉が出る
 * 3.「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *      次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 * 4. 表示されたウェブアプリURLと合言葉を、先生用ページの「📝 記録設定」に入力
 */

const SHEET_NAME = 'ログ';
const HEADER = ['日時', 'クラスコード', '生徒名', '送信者', '内容', '生徒ID', 'メッセージID'];

function setup() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('LOG_TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    props.setProperty('LOG_TOKEN', token);
  }
  getSheet_();
  Logger.log('合言葉：' + token);
  return token;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('画面モニター')
    .addItem('合言葉を表示', 'showToken')
    .addToUi();
}

function showToken() {
  const token = PropertiesService.getScriptProperties().getProperty('LOG_TOKEN') || setup();
  SpreadsheetApp.getUi().alert('合言葉：' + token);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME, 0);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold').setBackground('#e8eefc');
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(3, 140);
    sh.setColumnWidth(5, 420);
    sh.getRange('E:E').setWrap(true);
  }
  return sh;
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty('LOG_TOKEN');
    if (!token || body.token !== token) return json_({ ok: false, error: '合言葉が違います' });

    const rows = (body.rows || []).slice(0, 500);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = getSheet_();
      // 再送による重複を防ぐ（直近1000行のメッセージIDと照合）
      const last = sh.getLastRow();
      const n = Math.min(1000, Math.max(0, last - 1));
      const seen = new Set(n ? sh.getRange(last - n + 1, 7, n, 1).getValues().map(r => String(r[0])) : []);
      const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
      const out = [];
      for (const r of rows) {
        const id = String(r.id || '');
        if (id && seen.has(id)) continue;
        seen.add(id);
        out.push([
          Utilities.formatDate(new Date(Number(r.ts) || Date.now()), tz, 'yyyy/MM/dd HH:mm:ss'),
          clean_(r.room, 20),
          clean_(r.name, 60),
          clean_(r.from, 20),
          clean_(r.text, 2000),
          clean_(r.sid, 40),
          id,
        ]);
      }
      if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, HEADER.length).setValues(out);
      return json_({ ok: true, saved: out.length });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// 数式として解釈されないように先頭の = + - @ を無害化
function clean_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
