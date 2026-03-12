/**
 * 共通ユーティリティ（定数・ヘルパー関数）
 */

// シート名
const SHEET_NAMES = {
  FORM_RESPONSE: 'フォームの回答',
  CHILD_MASTER: '児童マスタ',
  MONTHLY_SUMMARY: '月別集計',
  ALLOCATION: '振り分け記録',
  VISIT_CALENDAR: '来館カレンダー',
  CONFIRMED_VISITS: '確定来館記録',
  CHILD_VIEW: '児童別ビュー',
};

// 児童マスタの列インデックス（1始まり）
const MASTER_COL = {
  NO: 1,
  NAME: 2,
  PARENT_NAME: 3,
  PARENT_EMAIL: 4,
  MONTHLY_QUOTA: 5,
  MEDICAL_TYPE: 6,
  STAFF: 7,
  ENROLLMENT: 8,
  VISIT_DAYS: 9,
  PRIORITY: 10,
};

// フォームの回答の列インデックス（1始まり）
const FORM_COL = {
  TIMESTAMP: 1,
  RECORD_DATE: 2,
  STAFF_NAME: 3,
  CHILD_NAME: 4,
  CHECK_IN: 5,
  CHECK_OUT: 6,
  TEMPERATURE: 7,
  MEAL: 8,
  BATH: 9,
  SLEEP: 10,
  BOWEL: 11,
  MEDICINE: 12,
  NOTES: 13,
};

// 月別集計の列インデックス（1始まり）
const SUMMARY_COL = {
  NO: 1,
  NAME: 2,
  QUOTA: 3,
  VISITS: 4,
  REMAINING: 5,
  USAGE_RATE: 6,
  STAFF: 7,
  ENROLLMENT: 8,
};

// 来館カレンダーのレイアウト定数
const CALENDAR_LAYOUT = {
  HEADER_ROW: 3,       // ヘッダー行（日付 | 児童1 | 児童2 | ... | 日計）
  DATA_START_ROW: 4,   // データ開始行
  DATE_COL: 1,         // 日付列
  CHILD_START_COL: 2,  // 児童列の開始
};

// 振り分け記録の列インデックス（1始まり）
const ALLOCATION_COL = {
  TARGET_MONTH: 1,
  CHILD_NAME: 2,
  ALLOCATION_DATE: 3,
  STAFF_NAME: 4,
  CHECK_IN: 5,
  CHECK_OUT: 6,
  TEMPERATURE: 7,
  MEAL: 8,
  BATH: 9,
  SLEEP: 10,
  BOWEL: 11,
  MEDICINE: 12,
  NOTES: 13,
  EXECUTED_AT: 14,
};

// 振り分け記録の列数
const ALLOCATION_COL_COUNT = 14;

// 振り分け補完用デフォルト値
const ALLOCATION_DEFAULTS = {
  CHECK_IN: '09:00',
  CHECK_OUT: '17:00',
  TEMPERATURE: 36.5,
  MEAL: '○',
  BATH: '○',
  SLEEP: '○',
  BOWEL: '○',
  MEDICINE: '○',
  NOTES: '特になし',
};

// 1日あたりの最大来館数
const MAX_VISITS_PER_DAY = 7;

// 曜日文字列 → 数値変換マップ（Date.getDay()に対応）
const DAY_OF_WEEK_MAP = {
  '日': 0, '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6,
};

// 確定来館記録の列インデックス（1始まり）
const CONFIRMED_COL = {
  RECORD_DATE: 1,
  CHILD_NAME: 2,
  DATA_TYPE: 3,
  STAFF_NAME: 4,
  CHECK_IN: 5,
  CHECK_OUT: 6,
  TEMPERATURE: 7,
  MEAL: 8,
  BATH: 9,
  SLEEP: 10,
  BOWEL: 11,
  MEDICINE: 12,
  NOTES: 13,
};

// メール本文テンプレート
var EMAIL_TEMPLATE = [
  '{保護者名} 様',
  '',
  'いつもお世話になっております。{施設名}です。',
  '',
  '{日付}の{児童名}さんの来館記録をお知らせいたします。',
  '',
  '■ 来館記録',
  '入所時間: {入所時間}',
  '退所時間: {退所時間}',
  '体温: {体温}℃',
  '食事: {食事}',
  '入浴: {入浴}',
  '睡眠: {睡眠}',
  '便: {便}',
  '服薬: {服薬}',
  '',
  '■ 連絡事項',
  '{連絡事項}',
  '',
  '担当スタッフ: {スタッフ名}',
].join('\n');

// 確定来館記録のデータ開始行（ヘッダーが1行目）
const CONFIRMED_DATA_START_ROW = 2;

// 月別集計シートのデータ開始行（ヘッダーが2行目）
const SUMMARY_DATA_START_ROW = 3;

// 児童別ビューの来館履歴開始行
const CHILD_VIEW_HISTORY_START_ROW = 10;

/**
 * シートを名前で取得する
 * @param {string} name シート名
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('シート「' + name + '」が見つかりません');
  }
  return sheet;
}

/**
 * 児童マスタの全データを取得する（ヘッダー除く）
 * @returns {Array<Array>} 児童マスタデータ
 */
function getChildMasterData() {
  const sheet = getSheet(SHEET_NAMES.CHILD_MASTER);
  const data = sheet.getDataRange().getValues();
  // ヘッダー行を除く
  return data.slice(1);
}

/**
 * 入所状況が「稼働」または「休止」の児童を取得する（「退所」を除外）
 * @returns {Array<Array>} 在籍中の児童データ
 */
function getActiveChildren() {
  const data = getChildMasterData();
  return data.filter(function(row) {
    var status = row[MASTER_COL.ENROLLMENT - 1];
    return status === '稼働' || status === '休止';
  });
}

/**
 * フォームの回答から指定年月のデータを取得する
 * @param {number} year 年
 * @param {number} month 月（1-12）
 * @returns {Array<Array>} 該当月のフォーム回答データ
 */
function getFormResponsesByMonth(year, month) {
  const sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
  const data = sheet.getDataRange().getValues();
  // ヘッダー行を除く
  var responses = data.slice(1);
  return responses.filter(function(row) {
    var recordDate = new Date(row[FORM_COL.RECORD_DATE - 1]);
    return recordDate.getFullYear() === year && (recordDate.getMonth() + 1) === month;
  });
}

/**
 * 振り分け記録から指定年月のデータを取得する
 * @param {number} year 年
 * @param {number} month 月（1-12）
 * @returns {Array<Array>} 該当月の振り分けデータ
 */
function getAllocationsByMonth(year, month) {
  var sheet;
  try {
    sheet = getSheet(SHEET_NAMES.ALLOCATION);
  } catch (e) {
    Logger.log('振り分け記録シートが存在しません: ' + e.message);
    return [];
  }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return [];
  }
  var records = data.slice(1);
  return records.filter(function(row) {
    var targetMonth = new Date(row[ALLOCATION_COL.TARGET_MONTH - 1]);
    return targetMonth.getFullYear() === year && (targetMonth.getMonth() + 1) === month;
  });
}

/**
 * 年月文字列をパースする（例: "2026年3月" → {year: 2026, month: 3}）
 * @param {string} yearMonthStr 年月文字列
 * @returns {{year: number, month: number}}
 */
function parseYearMonth(yearMonthStr) {
  // Date型の場合は直接year/monthを取得
  if (yearMonthStr instanceof Date) {
    return {
      year: yearMonthStr.getFullYear(),
      month: yearMonthStr.getMonth() + 1,
    };
  }
  var match = String(yearMonthStr).match(/(\d{4})年(\d{1,2})月/);
  if (!match) {
    throw new Error('年月の形式が不正です: ' + yearMonthStr);
  }
  return {
    year: parseInt(match[1], 10),
    month: parseInt(match[2], 10),
  };
}

/**
 * 年月ドロップダウン用の選択肢を生成する
 * フォームの回答に存在する年月 + 当月・翌月をユニークに返す（昇順）
 * @returns {Array<string>} 年月文字列の配列
 */
function generateYearMonthOptions() {
  // キーとして "YYYY-MM" 形式で管理し、重複排除
  var ymSet = {};

  // 当月・翌月を常に含める
  var now = new Date();
  for (var i = 0; i <= 1; i++) {
    var d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    var key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
    ymSet[key] = true;
  }

  // フォームの回答から年月を抽出
  try {
    var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var recordDate = data[r][FORM_COL.RECORD_DATE - 1];
      if (recordDate instanceof Date && !isNaN(recordDate.getTime())) {
        var key = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM');
        ymSet[key] = true;
      }
    }
  } catch (e) {
    Logger.log('フォームの回答シートの読み取りに失敗: ' + e.message);
  }

  // キーを昇順ソートして表示文字列に変換
  var keys = Object.keys(ymSet).sort();
  return keys.map(function(key) {
    var parts = key.split('-');
    return parseInt(parts[0], 10) + '年' + parseInt(parts[1], 10) + '月';
  });
}

/**
 * 児童名のドロップダウン選択肢を児童マスタから取得する
 * @returns {Array<string>} 児童名の配列
 */
function getChildNameOptions() {
  var children = getActiveChildren();
  return children.map(function(row) {
    return row[MASTER_COL.NAME - 1];
  });
}

/**
 * 確定来館記録から指定年月のデータを取得する
 * @param {number} year 年
 * @param {number} month 月（1-12）
 * @returns {Array<Array>} 該当月の確定来館記録データ
 */
function getConfirmedVisitsByMonth(year, month) {
  var sheet;
  try {
    sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  } catch (e) {
    Logger.log('確定来館記録シートが存在しません: ' + e.message);
    return [];
  }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return [];
  }
  var records = data.slice(1);
  return records.filter(function(row) {
    var recordDate = new Date(row[CONFIRMED_COL.RECORD_DATE - 1]);
    return recordDate.getFullYear() === year && (recordDate.getMonth() + 1) === month;
  });
}
