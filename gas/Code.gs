/**
 * 画面モニター 記録用（スプレッドシートにバインドして使う）
 *
 * はじめて使うとき：
 * 1. 新しいスプレッドシートを作る →「拡張機能」→「Apps Script」→ このコードを貼り付けて保存
 * 2. 左の「プロジェクトの設定」（歯車）で「「appsscript.json」マニフェスト ファイルをエディタで表示する」をオン →
 *    エディタに出た appsscript.json の中身を、GitHub の gas/appsscript.json に置き換えて保存
 * 3. 関数「setup」を選んで実行（初回は権限の許可が必要）→ 実行ログに合言葉が出る
 * 4.「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *      次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 * 5. 表示されたウェブアプリURLと合言葉を、先生用ページの「📝 記録設定」に入力
 * 6. （自動更新を使う場合）https://script.google.com/home/usersettings で「Google Apps Script API」をオン
 *
 * setup を実行すると、次の2つが自動で動くようになる（スプレッドシートのメニュー「画面モニター」からも実行できる）
 *  - 入退室まとめ：1時間ごとに「ログ」から「入退室まとめ」シートを作り直す
 *  - 自動更新：毎日1回、GitHub の最新の Code.gs / appsscript.json を取り込み、同じウェブアプリURLのまま新しい版にする
 */

const SCRIPT_VERSION = '2026-11-04.2040';
const SHEET_NAME = 'ログ';
const SUMMARY_SHEET = '入退室まとめ';
const HEADER = ['日時', 'クラスコード', '生徒名', '送信者', '内容', '生徒ID', 'メッセージID', 'IPアドレス'];
const REPO_RAW = 'https://raw.githubusercontent.com/hirokimu1124/KYOTANGO/main/gas/';

function setup() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('LOG_TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    props.setProperty('LOG_TOKEN', token);
  }
  getSheet_();
  setupAutomation();
  Logger.log('合言葉：' + token);
  // 自動更新の準備ができていれば、ここで最新版を取り込み、ウェブアプリも最新の版に切り替える
  try { Logger.log('自動更新：' + runUpdate_().msg); } catch (e) { Logger.log('自動更新：' + e); }
  return token;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('画面モニター')
    .addItem('合言葉を表示', 'showToken')
    .addItem('入退室まとめを今すぐ作る', 'buildSummaryNow')
    .addSeparator()
    .addItem('最新版を確認して更新する', 'checkUpdateNow')
    .addItem('自動化の状態を見る', 'showAutomationStatus')
    .addItem('自動化を設定し直す', 'setupAutomationNow')
    .addItem('自動更新を止める／再開する', 'toggleAutoUpdate')
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
  // 以前の版で作ったシート（7列）には「IPアドレス」の見出しを足す
  if (sh.getLastColumn() < HEADER.length) {
    sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold').setBackground('#e8eefc');
    sh.setColumnWidth(HEADER.length, 130);
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
          clean_(r.room, 40),
          clean_(r.name, 60),
          clean_(r.from, 20),
          clean_(r.text, 2000),
          clean_(r.sid, 40),
          id,
          clean_(r.ip, 64),
        ]);
      }
      if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, HEADER.length).setValues(out);
      return json_({ ok: true, saved: out.length, version: SCRIPT_VERSION });
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

// =====================================================================
// 自動化（時間主導のトリガー）
// =====================================================================

function setupAutomation() {
  const want = { buildSummary: 'hourly', autoUpdate: 'daily' };
  for (const t of ScriptApp.getProjectTriggers()) {
    if (want[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger('buildSummary').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('autoUpdate').timeBased().everyDays(1).atHour(3).create();
  PropertiesService.getScriptProperties().setProperty('AUTOMATION_SET', new Date().toISOString());
}

function setupAutomationNow() {
  setupAutomation();
  SpreadsheetApp.getUi().alert('自動化を設定しました。\n・入退室まとめ：1時間ごと\n・自動更新：毎日 午前3時ごろ');
}

function showAutomationStatus() {
  const p = PropertiesService.getScriptProperties();
  const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  const lines = [
    `この記録用コードの版：${SCRIPT_VERSION}`,
    `入退室まとめ：${triggers.includes('buildSummary') ? '1時間ごとに作成' : '未設定'}（最後：${p.getProperty('SUMMARY_AT') || 'まだ'}）`,
    `自動更新：${triggers.includes('autoUpdate') ? (p.getProperty('AUTO_UPDATE') === 'off' ? '止めています' : '毎日確認') : '未設定'}`,
    `最後の確認：${p.getProperty('UPDATE_AT') || 'まだ'}`,
    `結果：${p.getProperty('UPDATE_RESULT') || '—'}`,
  ];
  if (!triggers.includes('buildSummary') || !triggers.includes('autoUpdate')) lines.push('', '「自動化を設定し直す」を押すと設定できます。');
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

function toggleAutoUpdate() {
  const p = PropertiesService.getScriptProperties();
  const off = p.getProperty('AUTO_UPDATE') === 'off';
  p.setProperty('AUTO_UPDATE', off ? 'on' : 'off');
  SpreadsheetApp.getUi().alert(off ? '自動更新を再開しました。' : '自動更新を止めました（手動の「最新版を確認して更新する」は使えます）。');
}

// =====================================================================
// 入退室まとめ（「ログ」の 送信者＝入退室 の行から作る）
// =====================================================================

const SUMMARY_HEADER = ['日付', '学級', '名前', '区分', '最初の入室', '最後の退室', '滞在合計（分）',
  '入室回数', '再接続', 'IPアドレス', '回線', '備考'];

function buildSummaryNow() {
  const n = buildSummary();
  SpreadsheetApp.getUi().alert(`「${SUMMARY_SHEET}」を作りました（${n}行）。`);
}

function buildSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(SHEET_NAME);
  if (!log || log.getLastRow() < 2) return 0;
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const values = log.getRange(2, 1, log.getLastRow() - 1, HEADER.length).getValues();
  const rows = summarizeRows_(values, tz);

  let sh = ss.getSheetByName(SUMMARY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SUMMARY_SHEET, 1);
    sh.setFrozenRows(1);
    sh.setColumnWidth(3, 200);
    sh.setColumnWidth(10, 220);
    sh.setColumnWidth(12, 260);
  }
  sh.clearContents();
  sh.getRange(1, 1, 1, SUMMARY_HEADER.length).setValues([SUMMARY_HEADER]).setFontWeight('bold').setBackground('#e8eefc');
  if (rows.length) sh.getRange(2, 1, rows.length, SUMMARY_HEADER.length).setValues(rows);
  PropertiesService.getScriptProperties().setProperty('SUMMARY_AT',
    Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm'));
  return rows.length;
}

// ログの値（2次元配列）からまとめの行を作る。シートに触れないので単体で確かめられる
function summarizeRows_(values, tz) {
  const fmt = (d, f) => Utilities.formatDate(d, tz, f);
  const toDate = v => {
    if (v instanceof Date) return v;
    const m = String(v).match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/);
    return m ? new Date(+m[1], m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)) : null;
  };
  const groups = new Map();
  for (const r of values) {
    if (String(r[3]) !== '入退室') continue;
    const at = toDate(r[0]);
    if (!at) continue;
    const day = fmt(at, 'yyyy/MM/dd');
    const room = String(r[1]);
    const name = String(r[2]);
    const sid = String(r[5]);
    const text = String(r[4]);
    let kind;
    if (/^先生（.*・主担当）$/.test(name)) kind = '主担当';
    else if (/副担当として参加しようとした）$/.test(name)) kind = '参加拒否';
    else if (/^先生（.*・副担当）$/.test(name)) kind = '副担当';
    else kind = '生徒';
    const key = [day, room, kind, sid || name].join('|');
    let g = groups.get(key);
    if (!g) {
      g = { day, room, name, kind, first: null, last: null, stay: 0, joins: 0, rejoins: 0, ips: [], open: false, denied: [] };
      groups.set(key, g);
    }
    g.name = name; // 同じ生徒IDで名前を直した場合は最後の名前
    const ipm = String(r[7] || '') || ((text.match(/IP ([0-9A-Fa-f.:]+)/) || [])[1] || '');
    if (ipm && !g.ips.includes(ipm)) g.ips.push(ipm);
    if (/^入室/.test(text)) {
      g.joins++; g.open = true;
      if (!g.first || at < g.first) g.first = at;
    } else if (/^再接続/.test(text)) {
      g.rejoins++; g.open = true;
    } else if (/^(退室|記録終了)/.test(text)) {
      g.open = false;
      if (!g.last || at > g.last) g.last = at;
      const sm = text.match(/滞在 (?:(\d+)時間)?(\d+)分/);
      if (sm) g.stay += (+(sm[1] || 0)) * 60 + (+sm[2]);
    } else if (/^参加を拒否/.test(text)) {
      g.joins++;
      if (!g.first || at < g.first) g.first = at;
      const why = (text.match(/^参加を拒否：([^（]+)/) || [])[1];
      if (why && !g.denied.includes(why)) g.denied.push(why);
    }
  }

  // 日付・学級ごとに、生徒でいちばん多いIP（学校の回線など）を求め、それと違う回線の生徒に目印をつける
  const major = new Map();
  for (const g of groups.values()) {
    if (g.kind !== '生徒') continue;
    const k = g.day + '|' + g.room;
    const m = major.get(k) || new Map();
    for (const ip of g.ips) m.set(ip, (m.get(ip) || 0) + 1);
    major.set(k, m);
  }
  const majorIp = new Map();
  for (const [k, m] of major) {
    const top = [...m.entries()].sort((a, b) => b[1] - a[1]);
    if (top.length && top[0][1] >= 2 && (top.length === 1 || top[0][1] > top[1][1])) majorIp.set(k, top[0][0]);
  }

  const order = { '主担当': 0, '副担当': 1, '生徒': 2, '参加拒否': 3 };
  const natural = (a, b) => a.localeCompare(b, 'ja', { numeric: true });
  const list = [...groups.values()].sort((a, b) =>
    b.day.localeCompare(a.day) || natural(a.room, b.room) || order[a.kind] - order[b.kind] || natural(a.name, b.name));

  return list.map(g => {
    const mip = majorIp.get(g.day + '|' + g.room);
    let line = '';
    if (g.kind === '生徒' && mip && g.ips.length) line = g.ips.includes(mip) ? '多数と同じ' : '別の回線';
    const notes = [];
    if (g.kind === '参加拒否') notes.push('参加を断った：' + g.denied.join('、'));
    else if (g.open) notes.push('退室の記録なし（最後は接続中のまま）');
    if (g.ips.length > 1) notes.push('IPが途中で変わった');
    return [
      g.day, g.room, g.name, g.kind,
      g.first ? fmt(g.first, 'HH:mm:ss') : '',
      g.last ? fmt(g.last, 'HH:mm:ss') : '',
      g.kind === '参加拒否' ? '' : g.stay,
      g.joins, g.rejoins,
      g.ips.join('、') || '不明',
      line,
      notes.join('／'),
    ];
  });
}

// =====================================================================
// 自動更新（GitHub の最新版を取り込み、ウェブアプリのURLを変えずに新しい版にする）
//   前提：appsscript.json に script.projects / script.deployments / script.external_request の権限、
//         https://script.google.com/home/usersettings で「Google Apps Script API」がオン
// =====================================================================

function autoUpdate() {
  if (PropertiesService.getScriptProperties().getProperty('AUTO_UPDATE') === 'off') return '自動更新は止めています';
  const r = runUpdate_();
  // 失敗が始まったときだけエラーにする（Google から「スクリプトのエラー」の通知メールが届く）
  if (r.failed && r.firstFailure) throw new Error(r.msg);
  return r.msg;
}

function checkUpdateNow() {
  SpreadsheetApp.getUi().alert(runUpdate_().msg);
}

function runUpdate_() {
  const p = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  let msg;
  try {
    msg = updateFromGitHub_();
  } catch (err) {
    msg = '更新できませんでした：' + String(err && err.message || err);
    if (/403|PERMISSION_DENIED|has not been used|disabled/i.test(msg)) {
      msg += '\n\n→ https://script.google.com/home/usersettings で「Google Apps Script API」をオンにし、' +
        'appsscript.json が最新か（権限 script.projects・script.deployments が入っているか）を確認してください。';
    }
  }
  const failed = /^更新できませんでした/.test(msg);
  const firstFailure = failed && !/^更新できませんでした/.test(p.getProperty('UPDATE_RESULT') || '');
  p.setProperty('UPDATE_AT', Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm'));
  p.setProperty('UPDATE_RESULT', msg.slice(0, 500));
  return { msg, failed, firstFailure };
}

function updateFromGitHub_() {
  const fetchText = name => {
    const r = UrlFetchApp.fetch(REPO_RAW + name + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error(`GitHub から ${name} を取得できません（${r.getResponseCode()}）`);
    return r.getContentText('UTF-8');
  };
  const code = fetchText('Code.gs');
  const manifest = fetchText('appsscript.json');
  const newVer = (code.match(/const SCRIPT_VERSION = '([^']+)'/) || [])[1];
  // 取得したものが本物の記録用コードかを確かめる（途中で切れたファイル等を入れない）
  if (!newVer || !/function doPost\(/.test(code) || !/function updateFromGitHub_\(/.test(code)) {
    throw new Error('取得したコードが記録用のコードではありません');
  }
  const man = JSON.parse(manifest);
  if (!man.webapp || !Array.isArray(man.oauthScopes)) throw new Error('取得した appsscript.json が正しくありません');

  const scriptId = ScriptApp.getScriptId();
  const api = (method, path, payload) => {
    const r = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + scriptId + path, {
      method, contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: payload ? JSON.stringify(payload) : undefined,
    });
    const c = r.getResponseCode(), t = r.getContentText();
    if (c >= 300) throw new Error(`Apps Script API ${method.toUpperCase()} ${path} → ${c} ${t.slice(0, 300)}`);
    return t ? JSON.parse(t) : {};
  };

  const files = (api('get', '/content').files || []).map(f => ({ name: f.name, type: f.type, source: f.source }));
  const main = files.find(f => f.type === 'SERVER_JS' && /function doPost\(/.test(f.source));
  if (!main) throw new Error('このプロジェクトに記録用のコード（doPost）が見つかりません');
  let mf = files.find(f => f.type === 'JSON' && f.name === 'appsscript');
  const sameManifest = mf && JSON.stringify(JSON.parse(mf.source)) === JSON.stringify(man);
  const props = PropertiesService.getScriptProperties();
  // コードが最新でも、ウェブアプリがその版で公開されていなければ切り替える（手で貼り替えた直後など）
  if (main.source === code && sameManifest && props.getProperty('DEPLOYED_VERSION') === newVer) {
    return `最新です（版 ${newVer}）`;
  }

  if (main.source !== code || !sameManifest) {
    main.source = code;
    if (mf) mf.source = manifest; else files.push({ name: 'appsscript', type: 'JSON', source: manifest });
    api('put', '/content', { files });
  }
  const desc = `自動更新 ${newVer}`;
  const version = api('post', '/versions', { description: desc });
  const deployments = api('get', '/deployments').deployments || [];
  const web = deployments.filter(d => d.deploymentConfig && d.deploymentConfig.versionNumber &&
    (d.entryPoints || []).some(e => e.entryPointType === 'WEB_APP'));
  for (const d of web) {
    api('put', '/deployments/' + d.deploymentId, {
      deploymentConfig: { scriptId, versionNumber: version.versionNumber, manifestFileName: 'appsscript', description: desc },
    });
  }
  if (web.length) props.setProperty('DEPLOYED_VERSION', newVer);
  else return `コードを ${newVer} にしました。ウェブアプリがまだないので、「デプロイ」→「新しいデプロイ」で公開してください。`;
  return `更新しました：${SCRIPT_VERSION} → ${newVer}（コードの版 ${version.versionNumber}、ウェブアプリ ${web.length} 件を切り替え。URLは変わりません）`;
}
