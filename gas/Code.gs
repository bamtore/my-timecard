/**
 * 勤怠打刻アプリ - API（Google Apps Script）
 *  - 出勤 / 休憩開始 / 休憩終了 / 退勤 を打刻
 *  - スプレッドシート（打刻ログ・勤怠記録）とGoogleカレンダーへ自動同期
 *
 * Copyright (c) 2026 BM
 * Released under the MIT License. See LICENSE.
 */

const CONFIG = {
  CALENDAR_ID: '',          // 空ならデフォルトカレンダー。専用カレンダーを使う場合はIDを入れる
  SPREADSHEET_ID: '',       // 空ならこのスクリプトが紐づくスプレッドシート
  LOG_SHEET: '打刻ログ',
  DAILY_SHEET: '勤怠記録',
  TZ: 'Asia/Tokyo',
  WORK_TITLE_ACTIVE: '🟢 勤務中',
  BREAK_TITLE: '☕ 休憩',
};

const STATE_KEY = 'KINTAI_STATE';
const ACTION_LABEL = { in: '出勤', breakStart: '休憩開始', breakEnd: '休憩終了', out: '退勤' };

// ---------- 初期設定（エディタから1回だけ実行） ----------

/**
 * APIトークンを発行し、シートを作成する。
 * 実行ログに表示されたトークンをスマホ側の設定画面に入力する。
 * トークンを作り直したい場合は regenerateToken を実行する。
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('API_TOKEN');
  if (!token) {
    token = newToken_();
    props.setProperty('API_TOKEN', token);
  }
  getDailySheet_();
  getLogSheet_();
  getCalendar_();
  Logger.log('記録先スプレッドシート: ' + getSpreadsheet_().getUrl());
  Logger.log('APIトークン: ' + token);
}

function regenerateToken() {
  const token = newToken_();
  PropertiesService.getScriptProperties().setProperty('API_TOKEN', token);
  Logger.log('新しいAPIトークン: ' + token);
}

function newToken_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

// ---------- Web API ----------

function doGet() {
  return ContentService.createTextOutput('勤怠API は稼働中です');
}

/**
 * リクエスト: POST（Content-Type: text/plain）本文 JSON {token, action}
 * action: status | in | breakStart | breakEnd | out
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!checkToken_(body.token)) return json_({ ok: false, error: '認証に失敗しました。トークンを確認してください' });
    const data = body.action === 'status' ? getStatus() : punch(body.action);
    return json_({ ok: true, data: data });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function checkToken_(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  if (!expected || typeof token !== 'string' || token.length !== expected.length) return false;
  let diff = 0;
  for (let i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ token.charCodeAt(i);
  return diff === 0;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 画面表示用の現在状態 */
function getStatus() {
  return buildResponse_(loadState_(), null);
}

/** 打刻 */
function punch(action) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const state = loadState_();
    const now = new Date();
    let message;

    switch (action) {
      case 'in':
        assertStatus_(state, ['off'], 'すでに出勤済みです');
        message = clockIn_(state, now);
        break;
      case 'breakStart':
        assertStatus_(state, ['working'], '勤務中のみ休憩を開始できます');
        message = breakStart_(state, now);
        break;
      case 'breakEnd':
        assertStatus_(state, ['break'], '休憩中ではありません');
        message = breakEnd_(state, now);
        break;
      case 'out':
        assertStatus_(state, ['working', 'break'], '出勤していません');
        if (state.status === 'break') breakEnd_(state, now); // 休憩中なら自動で休憩終了
        message = clockOut_(state, now);
        break;
      default:
        throw new Error('不明な操作です: ' + action);
    }

    saveState_(state);
    appendLog_(now, action);
    return buildResponse_(state, message);
  } finally {
    lock.releaseLock();
  }
}

// ---------- 各打刻処理 ----------

function clockIn_(state, now) {
  const sheet = getDailySheet_();
  sheet.appendRow([now, now, '', '', '', '']);
  const row = sheet.getLastRow();
  sheet.getRange(row, 1).setNumberFormat('yyyy/mm/dd (ddd)');
  sheet.getRange(row, 2, 1, 2).setNumberFormat('HH:mm');
  sheet.getRange(row, 4, 1, 2).setNumberFormat('[h]:mm');

  const event = getCalendar_().createEvent(
    CONFIG.WORK_TITLE_ACTIVE,
    now,
    new Date(now.getTime() + 60 * 1000),
    { description: '勤務中（退勤時に終了時刻が更新されます）' }
  );

  Object.assign(state, {
    status: 'working',
    start: now.getTime(),
    row: row,
    workEventId: event.getId(),
    breakMinutes: 0,
    breakStart: null,
  });
  return fmtTime_(now) + ' に出勤しました';
}

function breakStart_(state, now) {
  state.status = 'break';
  state.breakStart = now.getTime();
  return fmtTime_(now) + ' から休憩を開始しました';
}

function breakEnd_(state, now) {
  const start = new Date(state.breakStart);
  const minutes = Math.round((now - start) / 60000);
  state.breakMinutes = (state.breakMinutes || 0) + minutes;
  state.breakStart = null;
  state.status = 'working';

  getCalendar_().createEvent(CONFIG.BREAK_TITLE, start, now);
  getDailySheet_().getRange(state.row, 4).setValue(state.breakMinutes / 1440);
  return `休憩を終了しました（${minutes}分）`;
}

function clockOut_(state, now) {
  const start = new Date(state.start);
  const workMinutes = Math.max(0, Math.round((now - start) / 60000) - state.breakMinutes);
  const label = fmtDuration_(workMinutes);

  const sheet = getDailySheet_();
  sheet.getRange(state.row, 3).setValue(now);
  sheet.getRange(state.row, 4).setValue(state.breakMinutes / 1440);
  sheet.getRange(state.row, 5).setValue(workMinutes / 1440);

  const event = getCalendar_().getEventById(state.workEventId);
  if (event) {
    event.setTime(start, now);
    event.setTitle(`💼 勤務 ${label}`);
    event.setDescription(
      `出勤 ${fmtTime_(start)} / 退勤 ${fmtTime_(now)}\n休憩 ${fmtDuration_(state.breakMinutes)}\n実働 ${label}`
    );
  }

  Object.keys(state).forEach((k) => delete state[k]);
  state.status = 'off';
  return `${fmtTime_(now)} に退勤しました。おつかれさまでした！（実働 ${label}）`;
}

// ---------- シート・カレンダー ----------

/**
 * 保存先スプレッドシートを取得する。優先順位:
 *  1. CONFIG.SPREADSHEET_ID
 *  2. このスクリプトが紐づくスプレッドシート（「拡張機能 → Apps Script」から作った場合）
 *  3. 以前に自動作成したスプレッドシート
 *  4. どれも無ければ「勤怠記録」を新規作成
 */
function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('SPREADSHEET_ID');
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      // 削除された等で開けない場合は作り直す
    }
  }

  const ss = SpreadsheetApp.create('勤怠記録');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function getSheet_(name, headers) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getDailySheet_() {
  return getSheet_(CONFIG.DAILY_SHEET, ['日付', '出勤', '退勤', '休憩', '実働', '備考']);
}

function getLogSheet_() {
  return getSheet_(CONFIG.LOG_SHEET, ['日時', '種別']);
}

function appendLog_(now, action) {
  const sheet = getLogSheet_();
  sheet.appendRow([now, ACTION_LABEL[action]]);
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('yyyy/mm/dd HH:mm:ss');
}

function getCalendar_() {
  const cal = CONFIG.CALENDAR_ID
    ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID)
    : CalendarApp.getDefaultCalendar();
  if (!cal) throw new Error('カレンダーが見つかりません。CONFIG.CALENDAR_ID を確認してください');
  return cal;
}

// ---------- 状態管理 ----------

function loadState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(STATE_KEY);
  return raw ? JSON.parse(raw) : { status: 'off' };
}

function saveState_(state) {
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(state));
}

function assertStatus_(state, allowed, msg) {
  if (allowed.indexOf(state.status) === -1) throw new Error(msg);
}

/** 打刻状態がおかしくなった時にエディタから手動実行してリセット */
function resetState() {
  PropertiesService.getScriptProperties().deleteProperty(STATE_KEY);
}

// ---------- レスポンス ----------

function buildResponse_(state, message) {
  return {
    status: state.status,
    start: state.start || null,
    breakStart: state.breakStart || null,
    breakMinutes: state.breakMinutes || 0,
    serverTime: Date.now(),
    message: message,
    logs: recentLogs_(),
  };
}

function recentLogs_() {
  const sheet = getLogSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const n = Math.min(8, last - 1);
  return sheet
    .getRange(last - n + 1, 1, n, 2)
    .getValues()
    .reverse()
    .map(([d, type]) => ({
      time: Utilities.formatDate(new Date(d), CONFIG.TZ, 'M/d HH:mm'),
      type: type,
    }));
}

function fmtTime_(d) {
  return Utilities.formatDate(d, CONFIG.TZ, 'HH:mm');
}

function fmtDuration_(min) {
  return `${Math.floor(min / 60)}時間${min % 60}分`;
}
