/**
 * 明和商会 社有車ログ — Google Apps Script（受け付け窓口）
 *
 * スプレッドシートの「拡張機能 → Apps Script」に貼り付けて使う。
 * 画面（GitHub Pages）からの記録をここで受け取り、シートに書き込む。
 *
 * シート構成
 *   使用者 … A列「名前」に社員名を1行ずつ
 *   車両   … A列「車種」、B列「ナンバー」（任意）。車種名は重複させない（例：プロボックス①）
 *   ログ   … アプリが自動で書き込む。手で触らない
 *   端末   … スマホを本人専用に登録した記録。アプリが自動で書き込む
 */

const SHEET = { people: '使用者', cars: '車両', logs: 'ログ', devices: '端末' };
const DEVICE_HEADERS = ['登録日時', '名前', '端末ID', '端末の種類'];
const LOG_HEADERS = ['ID', '日付', '車両', 'ナンバー', '使用者', '使用開始', '返却', '使用時間(分)', '返却者', '開始ms', '返却ms'];
// ログシートの列番号（1始まり）
const COL = { id: 1, date: 2, car: 3, plate: 4, user: 5, start: 6, end: 7, minutes: 8, returnedBy: 9, startMs: 10, endMs: 11 };
// 「取り消す」を受け付ける時間
const UNDO_LIMIT_MS = 10 * 60 * 1000;
// 状態のキャッシュ秒数（シートを手で編集した内容はこの秒数以内に反映される）
const CACHE_SEC = 30;

// ===== 初回セットアップ（エディタから一度だけ実行する） =====
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');
  ensureSheet_(ss, SHEET.people, ['名前']);
  ensureSheet_(ss, SHEET.cars, ['車種', 'ナンバー']);
  const logs = ensureSheet_(ss, SHEET.logs, LOG_HEADERS);
  // 人が見なくてよい列は隠す
  logs.hideColumns(COL.id);
  logs.hideColumns(COL.startMs, 2);
  // 空の「シート1」があれば片付ける（中身がある場合は残す）
  const blank = ss.getSheetByName('シート1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  ensureSheet_(ss, SHEET.devices, DEVICE_HEADERS);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PASSKEY')) props.setProperty('PASSKEY', newPasskey_());
  if (!props.getProperty('ADMIN_CODE')) props.setProperty('ADMIN_CODE', newAdminCode_());
  Logger.log('セットアップ完了。合言葉：' + props.getProperty('PASSKEY') + '　管理者コード：' + props.getProperty('ADMIN_CODE'));
}

// 管理者コード（スマホの名前登録を解除するときに使う6桁の数字）を確認する。未設定なら作る
function showAdminCode() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_CODE')) props.setProperty('ADMIN_CODE', newAdminCode_());
  Logger.log('管理者コード：' + props.getProperty('ADMIN_CODE'));
}

// 管理者コードを作り直す（コードが漏れたとき用）
function resetAdminCode() {
  const code = newAdminCode_();
  PropertiesService.getScriptProperties().setProperty('ADMIN_CODE', code);
  Logger.log('新しい管理者コード：' + code);
}

function newAdminCode_() {
  return String(100000 + Math.floor(Math.random() * 900000));
}

// 合言葉を確認する
function showPasskey() {
  Logger.log('合言葉：' + PropertiesService.getScriptProperties().getProperty('PASSKEY'));
}

// 合言葉を作り直す（リンクが外部に漏れたとき用。Mattermost のピン止めリンクも差し替えること）
function resetPasskey() {
  const key = newPasskey_();
  PropertiesService.getScriptProperties().setProperty('PASSKEY', key);
  Logger.log('新しい合言葉：' + key);
}

function newPasskey_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ===== 窓口 =====
// 動作確認用：ブラウザでURLを開くとこの文字が出る
function doGet() {
  return ContentService.createTextOutput('社有車ログ API は動いています');
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  const key = PropertiesService.getScriptProperties().getProperty('PASSKEY');
  if (!key || req.key !== key) return json_({ ok: false, error: 'auth' });

  try {
    switch (req.action) {
      case 'state': return json_({ ok: true, state: getState_() });
      case 'logs': return json_({ ok: true, logs: getLogs_(req.month) });
      case 'start': return json_(withLock_(() => start_(req.carId, req.userName)));
      case 'return': return json_(withLock_(() => return_(req.carId, req.userName)));
      case 'undo': return json_(withLock_(() => undo_(req.type, req.logId)));
      case 'register': return json_(withLock_(() => register_(req.userName, req.deviceId, req.ua)));
      case 'admin': return json_(checkAdmin_(req.code));
      default: return json_({ ok: false, error: 'bad_request' });
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server', message: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 2人が同時に同じ車を押しても二重に記録されないよう、書き込みは1件ずつ処理する
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'busy' };
  try { return fn(); }
  finally { lock.releaseLock(); }
}

// ===== 読み込み =====
function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。setup を実行してください');
  return sh;
}

function readPeople_() {
  const sh = sheet_(SHEET.people);
  if (sh.getLastRow() < 2) return [];
  const seen = {};
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(n => n && !seen[n] && (seen[n] = true))
    .map(n => ({ id: n, name: n }));
}

function readCars_() {
  const sh = sheet_(SHEET.cars);
  if (sh.getLastRow() < 2) return [];
  const seen = {};
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .map(r => ({ id: String(r[0]).trim(), name: String(r[0]).trim(), plate: String(r[1]).trim() }))
    .filter(c => c.id && !seen[c.id] && (seen[c.id] = true));
}

function readLogs_() {
  const sh = sheet_(SHEET.logs);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2,
    id: String(r[COL.id - 1]),
    carId: String(r[COL.car - 1]),
    carName: String(r[COL.car - 1]),
    plate: String(r[COL.plate - 1]),
    userName: String(r[COL.user - 1]),
    returnedBy: String(r[COL.returnedBy - 1] || ''),
    startMs: Number(r[COL.startMs - 1]) || 0,
    endMs: Number(r[COL.endMs - 1]) || null
  })).filter(l => l.id && l.startMs);
}

// 車両ごとの「使用中／空き」はログの返却欄が空かどうかで判断する
function computeState_() {
  const status = {};
  readLogs_().forEach(l => {
    const cur = status[l.carId];
    if (!l.endMs) {
      status[l.carId] = { inUse: true, userName: l.userName, sinceMs: l.startMs, logId: l.id };
    } else if (!cur || !cur.inUse) {
      if (!cur || (cur.lastReturnMs || 0) < l.endMs) status[l.carId] = { inUse: false, lastReturnMs: l.endMs, lastUserName: l.userName };
    }
  });
  return saveState_({ people: readPeople_(), cars: readCars_(), status: status });
}

// 状態をキャッシュに入れて返す（他の人のスマホからの読み込みも速くなる）
function saveState_(state) {
  state.now = Date.now();
  CacheService.getScriptCache().put('state', JSON.stringify(state), CACHE_SEC);
  return state;
}

function getState_() {
  const cached = CacheService.getScriptCache().get('state');
  return cached ? JSON.parse(cached) : computeState_();
}

// month は "2026-10" の形。日本時間でその月に開始した記録を新しい順に返す
function getLogs_(month) {
  if (!/^\d{4}-\d{2}$/.test(month || '')) throw new Error('month の形式が正しくありません');
  const [y, m] = month.split('-').map(Number);
  const off = 9 * 3600000;
  const from = Date.UTC(y, m - 1, 1) - off, to = Date.UTC(y, m, 1) - off;
  return readLogs_()
    .filter(l => l.startMs >= from && l.startMs < to)
    .sort((a, b) => b.startMs - a.startMs)
    .map(l => { const o = Object.assign({}, l); delete o.row; return o; });
}

function findLogRow_(logId) {
  const sh = sheet_(SHEET.logs);
  const last = sh.getLastRow();
  if (last < 2 || !logId) return 0;
  const ids = sh.getRange(2, COL.id, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(logId)) return i + 2;
  return 0;
}

// ===== 書き込み =====
function start_(carId, userName) {
  const state = computeState_();
  const car = state.cars.find(c => c.id === carId);
  if (!car) return { ok: false, error: 'unknown_car' };
  if (!state.people.some(p => p.name === userName)) return { ok: false, error: 'unknown_user' };
  const cur = state.status[carId];
  if (cur && cur.inUse) return { ok: false, error: 'in_use', userName: cur.userName, state: state };

  const now = new Date();
  const id = Utilities.getUuid();
  const sh = sheet_(SHEET.logs);
  sh.appendRow([id, now, car.name, car.plate, userName, now, '', '', '', now.getTime(), '']);
  const row = sh.getLastRow();
  sh.getRange(row, COL.date).setNumberFormat('yyyy/mm/dd (ddd)');
  sh.getRange(row, COL.start, 1, 2).setNumberFormat('HH:mm');
  // シートを読み直さず、手元の状態を更新して返す（応答を速くするため）
  state.status[carId] = { inUse: true, userName: userName, sinceMs: now.getTime(), logId: id };
  return { ok: true, logId: id, at: now.getTime(), state: saveState_(state) };
}

function return_(carId, userName) {
  const state = computeState_();
  const cur = state.status[carId];
  if (!cur || !cur.inUse) return { ok: false, error: 'not_in_use', state: state };
  const row = findLogRow_(cur.logId);
  if (!row) return { ok: false, error: 'not_in_use', state: state };

  const now = new Date();
  const minutes = Math.max(0, Math.round((now.getTime() - cur.sinceMs) / 60000));
  const sh = sheet_(SHEET.logs);
  sh.getRange(row, COL.end, 1, 3).setValues([[now, minutes, userName || '']]);
  sh.getRange(row, COL.endMs).setValue(now.getTime());
  state.status[carId] = { inUse: false, lastReturnMs: now.getTime(), lastUserName: cur.userName };
  return { ok: true, logId: cur.logId, at: now.getTime(), minutes: minutes, state: saveState_(state) };
}

// スマホを本人専用にした記録を残す（わざと他人の名前で登録した場合も後から確認できる）
function register_(userName, deviceId, ua) {
  if (!readPeople_().some(p => p.name === userName)) return { ok: false, error: 'unknown_user' };
  const sh = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET.devices, DEVICE_HEADERS);
  sh.appendRow([new Date(), userName, String(deviceId || '').slice(0, 64), String(ua || '').slice(0, 200)]);
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('yyyy/mm/dd HH:mm');
  return { ok: true };
}

// 管理者コードの確認。総当たりを防ぐため、10分間に10回間違えたら受け付けない
function checkAdmin_(code) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('adminFails') || 0);
  if (fails >= 10) return { ok: false, error: 'admin_locked' };
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_CODE');
  if (!real || String(code) !== real) {
    cache.put('adminFails', String(fails + 1), 600);
    return { ok: false, error: 'admin_auth' };
  }
  return { ok: true };
}

function undo_(type, logId) {
  const row = findLogRow_(logId);
  if (!row) return { ok: false, error: 'undo_expired', state: computeState_() };
  const sh = sheet_(SHEET.logs);
  const r = sh.getRange(row, 1, 1, LOG_HEADERS.length).getValues()[0];
  const nowMs = Date.now();
  const startMs = Number(r[COL.startMs - 1]), endMs = Number(r[COL.endMs - 1]) || 0;

  if (type === 'start') {
    if (endMs || nowMs - startMs > UNDO_LIMIT_MS) return { ok: false, error: 'undo_expired', state: computeState_() };
    sh.deleteRow(row);
  } else if (type === 'return') {
    if (!endMs || nowMs - endMs > UNDO_LIMIT_MS) return { ok: false, error: 'undo_expired', state: computeState_() };
    // 返却を取り消す間に別の人が使い始めていたら戻さない
    const st = computeState_().status[String(r[COL.car - 1])];
    if (st && st.inUse) return { ok: false, error: 'in_use', userName: st.userName, state: computeState_() };
    sh.getRange(row, COL.end, 1, 3).clearContent();
    sh.getRange(row, COL.endMs).clearContent();
  } else {
    return { ok: false, error: 'bad_request' };
  }
  return { ok: true, state: computeState_() };
}
