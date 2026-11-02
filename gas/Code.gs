/**
 * 受付アルバイト相談窓口（Google Apps Script / スプレッドシート連携）
 *
 * 上位アルバイトが、受付アルバイトの気になった言動を文章で記録する。点数化はしない。
 * 画面は GitHub Pages の index.html。この GAS は doPost で JSON API を提供する。
 *
 * シート構成（setup() が作成）
 *   記録 : 送信日時 / 大会名 / 入力者 / 対象スタッフ / 言動 / 備考 / 優先度 / 対応メモ / リクエストID
 *          （優先度・対応メモは管理者がシフト決定時に手入力する欄。アプリは書き込まない）
 *   大会 : 大会名 / 表示（チェックが入った大会だけが画面の選択肢に出る）
 *   設定 : 項目 / 値（PIN）
 */

const TZ = 'Asia/Tokyo';
const MAX_LEN = 200;
const SH = { LOG: '記録', TOURN: '大会', CONF: '設定' };
const HEAD = {
  [SH.LOG]: ['送信日時', '大会名', '入力者', '対象スタッフ', '言動', '備考', '優先度', '対応メモ', 'リクエストID'],
  [SH.TOURN]: ['大会名', '表示'],
  [SH.CONF]: ['項目', '値'],
};

/* ---------- 初期設定（GASエディタで1回実行） ---------- */

/**
 * スプレッドシートがなければ作成し、3シートの見出し・書式とPINを用意する。
 * 何度実行してもよい（既存のデータとPINは消さない）。
 */
function setup() {
  const p = PropertiesService.getScriptProperties();
  let book;
  const id = p.getProperty('SPREADSHEET_ID');
  if (id) book = SpreadsheetApp.openById(id);
  else { book = SpreadsheetApp.create('受付アルバイト相談窓口'); p.setProperty('SPREADSHEET_ID', book.getId()); }
  book.setSpreadsheetTimeZone(TZ);
  Object.keys(HEAD).forEach(function (name) {
    const sh = book.getSheetByName(name) || book.insertSheet(name);
    const h = HEAD[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#f1f3f4');
    sh.setFrozenRows(1);
  });
  const def = book.getSheetByName('シート1') || book.getSheetByName('Sheet1');
  if (def && book.getSheets().length > 3) book.deleteSheet(def);
  const log = book.getSheetByName(SH.LOG);
  log.getRange(2, 1, log.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  log.getRange(1, 2, log.getMaxRows(), 8).setNumberFormat('@').setWrap(true);
  if (log.setColumnWidths) { log.setColumnWidths(5, 2, 320); log.setColumnWidth(4, 120); }
  const tour = book.getSheetByName(SH.TOURN);
  tour.getRange(2, 2, tour.getMaxRows() - 1, 1).insertCheckboxes();
  const conf = book.getSheetByName(SH.CONF);
  conf.getRange(1, 1, conf.getMaxRows(), 2).setNumberFormat('@');
  const rows = conf.getRange(2, 1, Math.max(conf.getLastRow() - 1, 1), 2).getValues();
  if (!rows.some(function (r) { return str_(r[0]) === 'PIN'; })) {
    conf.getRange(2, 1, 1, 2).setValues([['PIN', String(Math.floor(1000 + Math.random() * 9000))]]);
  }
  Logger.log('スプレッドシート：' + book.getUrl());
  Logger.log('PIN：' + pin_());
}

/* ---------- 画面からの API ---------- */

const API = { getConfig: getConfig, login: login, submit: submit };

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = API[req.fn];
    if (!fn) throw new Error('不明な操作です');
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('受付アルバイト相談窓口 API').setMimeType(ContentService.MimeType.TEXT);
}

/** 表示中の大会名の一覧 */
function getConfig() {
  return { tournaments: tournaments_() };
}

/** 名前とPINでログイン（PINの確認のみ。記録は残さない） */
function login(pin, name) {
  checkPin_(pin);
  if (!str_(name)) throw new Error('名前を入力してください');
  return { tournaments: tournaments_() };
}

/** 相談内容を記録する */
function submit(pin, name, tournament, target, action, note, reqId) {
  checkPin_(pin);
  name = str_(name); tournament = str_(tournament); target = str_(target);
  action = str_(action); note = str_(note); reqId = str_(reqId);
  if (!name) throw new Error('名前を入力してください');
  if (!tournament) throw new Error('大会名を選んでください');
  if (!target) throw new Error('対象スタッフを入力してください');
  if (!action) throw new Error('言動を入力してください');
  if (len_(target) > 50) throw new Error('対象スタッフは50字以内で入力してください');
  if (len_(action) > MAX_LEN) throw new Error('言動は' + MAX_LEN + '字以内で入力してください');
  if (len_(note) > MAX_LEN) throw new Error('備考は' + MAX_LEN + '字以内で入力してください');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('混み合っています。もう一度送信してください');
  try {
    const sh = sheet_(SH.LOG);
    const last = sh.getLastRow();
    if (reqId && last >= 2) {
      const ids = sh.getRange(2, 9, last - 1, 1).getValues();
      if (ids.some(function (r) { return str_(r[0]) === reqId; })) return { ok: true, duplicate: true }; // 二重送信
    }
    sh.getRange(last + 1, 1, 1, HEAD[SH.LOG].length).setValues([[new Date(), tournament, name, target, action, note, '', '', reqId]]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 内部処理 ---------- */

function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('初期設定が済んでいません（GASで setup を実行してください）');
  return SpreadsheetApp.openById(id);
}
function sheet_(name) {
  const sh = book_().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません（GASで setup を実行してください）');
  return sh;
}
function pin_() {
  const sh = sheet_(SH.CONF);
  const last = sh.getLastRow();
  if (last < 2) return '';
  const r = sh.getRange(2, 1, last - 1, 2).getValues().filter(function (x) { return str_(x[0]) === 'PIN'; })[0];
  return r ? str_(r[1]) : '';
}
function checkPin_(pin) {
  const p = pin_();
  if (!p || p !== str_(pin)) throw new Error('PINが正しくありません');
}
function tournaments_() {
  const sh = sheet_(SH.TOURN);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 2).getValues()
    .filter(function (r) { return str_(r[0]) && r[1] === true; })
    .map(function (r) { return str_(r[0]); });
}
function str_(v) { return v == null ? '' : String(v).trim(); }
function len_(s) { return Array.from(String(s)).length; }
