/**
 * 明和商会 社有車ログ — Google Apps Script（受け付け窓口）
 *
 * スプレッドシートの「拡張機能 → Apps Script」に貼り付けて使う。
 * 画面（GitHub Pages）からの記録をここで受け取り、シートに書き込む。
 *
 * シート構成
 *   使用者 … A列「名前」に社員名を1行ずつ
 *   車両   … A列「車種」、B列「ナンバー」（任意）。車種名は重複させない（例：プロボックス①）
 *            C列「点検予定日」、D列「車検予定」（年と月）、E列「お知らせ」（例：給油必要）は任意。アプリの車両カードに表示される
 *   ログ   … アプリが自動で書き込む。手で触らない
 *   端末   … スマホを本人専用に登録した記録。アプリが自動で書き込む
 *   予約   … アプリが自動で書き込む。取り消した予約も「取消日時」を入れて残す
 */

const SHEET = { people: '使用者', cars: '車両', logs: 'ログ', devices: '端末', reserves: '予約' };
// 予約シートの列。日付は "2026-11-02" の文字、枠は 午前／午後／終日
const RES_HEADERS = ['ID', '日付', '枠', '車両', '使用者', '登録日時', '取消日時', '取消者'];
const RES_COL = { id: 1, date: 2, slot: 3, car: 4, user: 5, created: 6, canceledAt: 7, canceledBy: 8 };
const SLOT_NAME = { am: '午前', pm: '午後', day: '終日' };
// 何日先まで予約できるか（今日を0日目として）
const RESERVE_DAYS = 7;
const DEVICE_HEADERS = ['登録日時', '名前', '端末ID', '端末の種類'];
// 車両シートの列。C〜E列はカードに表示する予定・お知らせ（空欄なら表示しない）
// F列「給油メモ」は返却時の「ガソリン少なめ」でアプリが書き込む（管理者が消してもよい）
const CAR_HEADERS = ['車種', 'ナンバー', '点検予定日', '車検予定', 'お知らせ', '給油メモ', 'オイルランプ'];
const CAR_FUEL_COL = 6;
const CAR_OIL_COL = 7;
const LOG_HEADERS = ['ID', '日付', '車両', 'ナンバー', '使用者', '使用開始', '返却', '使用時間(分)', '返却者', '開始ms', '返却ms', 'ガソリン少なめ', 'オイルランプ点灯'];
// ログシートの列番号（1始まり）
const COL = { id: 1, date: 2, car: 3, plate: 4, user: 5, start: 6, end: 7, minutes: 8, returnedBy: 9, startMs: 10, endMs: 11, fuelLow: 12, oilLamp: 13 };
// 「取り消す」を受け付ける時間
const UNDO_LIMIT_MS = 10 * 60 * 1000;
// 状態のキャッシュ秒数（シートを手で編集した内容はこの秒数以内に反映される）
const CACHE_SEC = 30;

// ===== 初回セットアップ（エディタから一度だけ実行する） =====
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');
  ensureSheet_(ss, SHEET.people, ['名前']);
  // 既存のシートに、後から増えた列の見出しを足す（データはそのまま）
  addMissingHeaders_(ensureSheet_(ss, SHEET.cars, CAR_HEADERS), CAR_HEADERS);
  const logs = ensureSheet_(ss, SHEET.logs, LOG_HEADERS);
  addMissingHeaders_(logs, LOG_HEADERS);
  // 人が見なくてよい列は隠す
  logs.hideColumns(COL.id);
  logs.hideColumns(COL.startMs, 2);
  // 空の「シート1」があれば片付ける（中身がある場合は残す）
  const blank = ss.getSheetByName('シート1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  ensureSheet_(ss, SHEET.devices, DEVICE_HEADERS);
  const res = ensureSheet_(ss, SHEET.reserves, RES_HEADERS);
  addMissingHeaders_(res, RES_HEADERS);
  res.hideColumns(RES_COL.id);

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

function addMissingHeaders_(sh, headers) {
  const head = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach((h, i) => { if (!head[i]) sh.getRange(1, i + 1).setValue(h).setFontWeight('bold'); });
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
      case 'logs': {
        // ログは管理者だけが見られる
        const admin = checkAdmin_(req.code);
        if (!admin.ok) return json_(admin);
        return json_({ ok: true, logs: getLogs_(req.month) });
      }
      case 'start': return json_(withLock_(() => start_(req.carId, req.userName)));
      case 'return': return json_(withLock_(() => return_(req.carId, req.userName, req.code, req.fuelLow, req.oilLamp)));
      case 'undo': return json_(withLock_(() => undo_(req.type, req.logId, req.prevFuel, req.prevOil)));
      case 'register': return json_(withLock_(() => register_(req.userName, req.deviceId, req.ua)));
      case 'admin': return json_(checkAdmin_(req.code));
      case 'reserve': return json_(withLock_(() => reserve_(req.carId, req.userName, req.date, req.slot)));
      case 'cancelReserve': return json_(withLock_(() => cancelReserve_(req.id, req.userName, req.code)));
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
  return sh.getRange(2, 1, sh.getLastRow() - 1, CAR_HEADERS.length).getValues()
    .map(r => ({
      id: String(r[0]).trim(), name: String(r[0]).trim(), plate: String(r[1]).trim(),
      inspect: toYMD_(r[2]), shaken: toYM_(r[3]), note: String(r[4] || '').trim(),
      fuel: String(r[CAR_FUEL_COL - 1] || '').trim(),
      oil: String(r[CAR_OIL_COL - 1] || '').trim()
    }))
    .filter(c => c.id && !seen[c.id] && (seen[c.id] = true));
}

// 点検予定日を "2026-11-05" の形にそろえる（日付セル・"2026/11/5"・"11/5" に対応。年がなければ今年）
function toYMD_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  const s = String(v || '');
  let m = s.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (m) return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
  m = s.match(/^\s*(\d{1,2})\D+(\d{1,2})/);
  if (m) return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy') + '-' + pad2_(m[1]) + '-' + pad2_(m[2]);
  return '';
}

// 車検予定を "2027-03" の形にそろえる（日付セル・"2027/3"・"2027年3月" に対応）
function toYM_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM');
  const m = String(v || '').match(/(\d{4})\D+(\d{1,2})/);
  return m ? m[1] + '-' + pad2_(m[2]) : '';
}

function pad2_(n) { return ('0' + Number(n)).slice(-2); }

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
    endMs: Number(r[COL.endMs - 1]) || null,
    fuelLow: String(r[COL.fuelLow - 1] || '') === '○',
    oilLamp: String(r[COL.oilLamp - 1] || '') === '○'
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
  return saveState_({ people: readPeople_(), cars: readCars_(), status: status, reserves: readReserves_() });
}

// ===== 予約 =====
// 日本時間の今日を "2026-11-02" の形で。days を足した日も出せる
function ymdJst_(days) {
  return Utilities.formatDate(new Date(Date.now() + (days || 0) * 86400000), 'Asia/Tokyo', 'yyyy-MM-dd');
}
// 午前は12時まで、午後は12時から
function nowSlot_() {
  return Number(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'H')) < 12 ? 'am' : 'pm';
}
const SLOT_KEY = { '午前': 'am', '午後': 'pm', '終日': 'day' };
// 終日は午前・午後のどちらとも重なる
const slotsOverlap_ = (a, b) => a === b || a === 'day' || b === 'day';

// 今日以降の、取り消されていない予約（予約シートがまだなければ空）
function readReserves_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.reserves);
  if (!sh || sh.getLastRow() < 2) return [];
  const today = ymdJst_(0);
  return sh.getRange(2, 1, sh.getLastRow() - 1, RES_HEADERS.length).getValues()
    .map(r => ({
      id: String(r[RES_COL.id - 1]), date: toYMD_(r[RES_COL.date - 1]), slot: SLOT_KEY[String(r[RES_COL.slot - 1]).trim()] || '',
      carId: String(r[RES_COL.car - 1]).trim(), userName: String(r[RES_COL.user - 1]).trim(),
      canceled: !!r[RES_COL.canceledAt - 1]
    }))
    .filter(v => v.id && v.date && v.slot && !v.canceled && v.date >= today)
    .map(v => ({ id: v.id, date: v.date, slot: v.slot, carId: v.carId, userName: v.userName }))
    .sort((a, b) => (a.date + a.slot).localeCompare(b.date + b.slot));
}

function reserve_(carId, userName, date, slot) {
  const state = computeState_();
  const car = state.cars.find(c => c.id === carId);
  if (!car) return { ok: false, error: 'unknown_car' };
  if (!state.people.some(p => p.name === userName)) return { ok: false, error: 'unknown_user' };
  if (!SLOT_NAME[slot] || !/^\d{4}-\d{2}-\d{2}$/.test(date || '')) return { ok: false, error: 'bad_request' };
  // 予約できるのは今日から RESERVE_DAYS 日先まで。今日の午前は12時を過ぎたら予約できない
  if (date < ymdJst_(0) || date > ymdJst_(RESERVE_DAYS)) return { ok: false, error: 'out_of_range', state: state };
  if (date === ymdJst_(0) && nowSlot_() === 'pm' && slot !== 'pm') return { ok: false, error: 'out_of_range', state: state };
  const same = state.reserves.filter(v => v.date === date && slotsOverlap_(v.slot, slot));
  // 同じ車・同じ時間に先約がある
  const taken = same.find(v => v.carId === carId);
  if (taken) return { ok: false, error: 'reserved', userName: taken.userName, state: state };
  // 1人が同じ時間に予約できるのは1台まで
  const mine = same.find(v => v.userName === userName);
  if (mine) return { ok: false, error: 'already_reserved', carName: mine.carId, state: state };

  const id = Utilities.getUuid();
  const sh = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET.reserves, RES_HEADERS);
  // 日付は文字として残す（' を付けないとシートが日付に変換する）
  sh.appendRow([id, "'" + date, SLOT_NAME[slot], car.name, userName, new Date(), '', '']);
  sh.getRange(sh.getLastRow(), RES_COL.created).setNumberFormat('yyyy/mm/dd HH:mm');
  state.reserves = state.reserves.concat([{ id: id, date: date, slot: slot, carId: carId, userName: userName }])
    .sort((a, b) => (a.date + a.slot).localeCompare(b.date + b.slot));
  return { ok: true, id: id, state: saveState_(state) };
}

// 予約を取り消せるのは予約した本人だけ。ほかの人の予約は管理者コードがあれば取り消せる
function cancelReserve_(id, userName, code) {
  const state = computeState_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.reserves);
  const v = state.reserves.find(x => x.id === String(id || ''));
  if (!sh || !v) return { ok: false, error: 'not_found', state: state };
  if (v.userName !== userName) {
    if (!code) return { ok: false, error: 'not_owner', userName: v.userName, state: state };
    const admin = checkAdmin_(code);
    if (!admin.ok) return admin;
  }
  const ids = sh.getRange(2, RES_COL.id, sh.getLastRow() - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === v.id) {
      sh.getRange(i + 2, RES_COL.canceledAt, 1, 2).setValues([[new Date(), userName || '']]);
      sh.getRange(i + 2, RES_COL.canceledAt).setNumberFormat('yyyy/mm/dd HH:mm');
      break;
    }
  }
  state.reserves = state.reserves.filter(x => x.id !== v.id);
  return { ok: true, state: saveState_(state) };
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
  // 1人が同時に使えるのは1台まで
  const using = Object.keys(state.status).find(id => state.status[id].inUse && state.status[id].userName === userName);
  if (using) return { ok: false, error: 'already_using', carName: using, state: state };

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

// 返却できるのは使い始めた本人だけ。本人以外は管理者コードがあれば代理で返却できる
// fuelLow: true=ガソリン少なめ／false=問題なし（給油メモを消す）／未指定=給油メモを変えない（代理返却）
// oilLamp: オイルランプ点灯。fuelLow と同じ考え方
function return_(carId, userName, code, fuelLow, oilLamp) {
  const state = computeState_();
  const cur = state.status[carId];
  if (!cur || !cur.inUse) return { ok: false, error: 'not_in_use', state: state };
  if (cur.userName !== userName) {
    if (!code) return { ok: false, error: 'not_owner', userName: cur.userName, state: state };
    const admin = checkAdmin_(code);
    if (!admin.ok) return admin;
  }
  const row = findLogRow_(cur.logId);
  if (!row) return { ok: false, error: 'not_in_use', state: state };

  const now = new Date();
  const minutes = Math.max(0, Math.round((now.getTime() - cur.sinceMs) / 60000));
  const sh = sheet_(SHEET.logs);
  sh.getRange(row, COL.end, 1, 3).setValues([[now, minutes, userName || '']]);
  sh.getRange(row, COL.endMs).setValue(now.getTime());
  // 返却時の報告（ガソリン少なめ・オイルランプ点灯）をログと車両シートに書く
  const report = (flag, logCol, carCol, key) => {
    if (flag !== true && flag !== false) return;
    sh.getRange(row, logCol).setValue(flag ? '○' : '');
    const memo = flag ? Utilities.formatDate(now, 'Asia/Tokyo', 'M/d') + ' ' + userName + 'さんが報告' : '';
    setCarCell_(carId, carCol, memo);
    const car = state.cars.find(c => c.id === carId);
    if (car) car[key] = memo;
  };
  report(fuelLow, COL.fuelLow, CAR_FUEL_COL, 'fuel');
  report(oilLamp, COL.oilLamp, CAR_OIL_COL, 'oil');
  state.status[carId] = { inUse: false, lastReturnMs: now.getTime(), lastUserName: cur.userName };
  return { ok: true, logId: cur.logId, at: now.getTime(), minutes: minutes, state: saveState_(state) };
}

// 車両シートの報告欄（F列 給油メモ・G列 オイルランプ）を書き換える
function setCarCell_(carId, col, memo) {
  const sh = sheet_(SHEET.cars);
  if (sh.getLastRow() < 2) return;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === carId) { sh.getRange(i + 2, col).setValue(memo); return; }
  }
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

// prevFuel / prevOil: 返却を取り消すとき、返却前の給油メモ・オイルランプ欄に戻すための値（画面が覚えていたもの）
function undo_(type, logId, prevFuel, prevOil) {
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
    sh.getRange(row, COL.endMs, 1, 3).clearContent();
    if (typeof prevFuel === 'string') setCarCell_(String(r[COL.car - 1]), CAR_FUEL_COL, prevFuel.slice(0, 100));
    if (typeof prevOil === 'string') setCarCell_(String(r[COL.car - 1]), CAR_OIL_COL, prevOil.slice(0, 100));
  } else {
    return { ok: false, error: 'bad_request' };
  }
  return { ok: true, state: computeState_() };
}
