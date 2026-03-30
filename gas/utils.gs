/**
 * 共通ユーティリティ（定数・ヘルパー関数）
 */

// シート名
const SHEET_NAMES = {
  FORM_RESPONSE: 'フォームの回答',
  CHILD_MASTER: '児童マスタ',
  MONTHLY_SUMMARY: '月別集計',
  VISIT_CALENDAR: '来館カレンダー',
  CONFIRMED_VISITS: '確定来館記録',
  CHILD_VIEW: '児童別ビュー',
  LOG: 'ログ',
  SETTINGS: '設定',
  BOUNCE_LOG: 'バウンスログ',
};

// 設定シートの行インデックス（C列=デフォルト値）
const SETTINGS_ROW = {
  MAX_VISITS_PER_DAY: 2,  // 1日最大来館数
  CHECK_IN: 3,            // 入所時間
  CHECK_OUT: 4,           // 退所時間
  BUSINESS_DAYS: 5,       // 営業日
  TEMPERATURE: 6,         // 体温
  MEAL: 7,                // 食事
  BATH: 8,                // 入浴
  SLEEP: 9,               // 睡眠
  BOWEL: 10,              // 便
  MEDICINE: 11,           // 服薬
  NOTES: 12,              // 連絡事項
  ERROR_EMAIL: 13,        // エラー通知先メール
  EMAIL_SUBJECT: 14,      // メール件名
  EMAIL_BODY: 15,         // メール本文
  DUMMY_STAFF_NAME: 16,   // 固定スタッフ名（振り分け・スタッフ2補完用）
};

// 児童マスタの列インデックス（1始まり）
const MASTER_COL = {
  NO: 1,
  NAME: 2,
  PARENT_NAME: 3,
  PARENT_EMAIL: 4,
  STAFF: 5,
  MEDICAL_TYPE: 6,
  MEDICAL_SUPPORT: 7,   // 重度支援（あり/なし）
  PRIORITY: 8,          // 重度支援区分（区分1〜区分5）
  ANNUAL_QUOTA: 9,      // 年間利用枠（例: 180）。空欄の場合は上限なし
  MONTHLY_QUOTA: 10,
  ENROLLMENT: 11,
  VISIT_DAYS: 12,
  NON_VISIT_DAYS: 13,   // 非来館曜日
  DEPARTURE_DATE: 14,   // 退所日
  MASTER_NOTES: 15,     // 備考
};

// フォームの回答の列インデックス（1始まり）
const FORM_COL = {
  TIMESTAMP: 1,
  RECORD_DATE: 2,
  STAFF_NAME: 3,    // スタッフ1
  STAFF_NAME_2: 4,  // スタッフ2（任意）
  CHILD_NAME: 5,
  CHECK_IN: 6,
  CHECK_OUT: 7,
  TEMPERATURE: 8,
  MEAL: 9,
  BATH: 10,
  SLEEP: 11,
  BOWEL: 12,
  MEDICINE: 13,
  NOTES: 14,
  EMAIL_SENT: 15,
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
  HEADER_ROW: 3,       // ヘッダー行（日付 | 曜日 | 児童1 | 児童2 | ... | 日計）
  DATA_START_ROW: 4,   // データ開始行
  DATE_COL: 1,         // 日付列
  DOW_COL: 2,          // 曜日列
  CHILD_START_COL: 3,  // 児童列の開始
};

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
  STAFF_NAME: 4,    // スタッフ1
  STAFF_NAME_2: 5,  // スタッフ2（任意）
  CHECK_IN: 6,
  CHECK_OUT: 7,
  TEMPERATURE: 8,
  MEAL: 9,
  BATH: 10,
  SLEEP: 11,
  BOWEL: 12,
  MEDICINE: 13,
  NOTES: 14,
};

// メール本文テンプレート
var EMAIL_TEMPLATE = [
  '{保護者名} 様',
  '',
  'いつもお世話になっております。',
  'テスト施設です。',
  '',
  '{日付}の{児童名}さんの来館記録をお知らせいたします。',
  '',
  '■ 来館記録',
  '・入所時間: {入所時間}',
  '・退所時間: {退所時間}',
  '・体温: {体温}℃',
  '・食事: {食事}',
  '  入浴: {入浴}',
  '・睡眠: {睡眠}',
  '・便: {便}',
  '・服薬: {服薬}',
  '',
  ' ■ 連絡事項',
  '・{連絡事項}',
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
    var departureDate = row[MASTER_COL.DEPARTURE_DATE - 1];
    return (status === '稼働' || status === '休止') && !departureDate;
  });
}

/**
 * フォームの回答から指定年のデータを取得する
 * @param {number} year 年
 * @returns {Array<Array>} 該当年のフォーム回答データ
 */
function getFormResponsesByYear(year) {
  const sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(function(row) {
    var recordDate = new Date(row[FORM_COL.RECORD_DATE - 1]);
    return recordDate.getFullYear() === year;
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
 * フォームの回答から対象年を取得する
 * データがない場合は現在の年を返す
 * @returns {number} 対象年
 */
function getTargetYearFromFormResponses_() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.FORM_RESPONSE);
    if (!sheet) return new Date().getFullYear();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return new Date().getFullYear();
    var years = {};
    data.slice(1).forEach(function(row) {
      var recordDate = new Date(row[FORM_COL.RECORD_DATE - 1]);
      if (!isNaN(recordDate.getTime())) {
        var y = recordDate.getFullYear();
        years[y] = (years[y] || 0) + 1;
      }
    });
    var yearList = Object.keys(years).map(Number).sort(function(a, b) { return b - a; });
    return yearList.length > 0 ? yearList[0] : new Date().getFullYear();
  } catch (e) {
    return new Date().getFullYear();
  }
}

/**
 * 年月ドロップダウン用の選択肢を生成する
 * フォームの回答から対象年を取得し、1月〜12月を昇順で返す
 * @returns {Array<string>} 年月文字列の配列
 */
function generateYearMonthOptions() {
  var year = getTargetYearFromFormResponses_();
  var options = [];
  for (var m = 1; m <= 12; m++) {
    options.push(year + '年' + m + '月');
  }
  return options;
}

/**
 * 月別集計用の選択肢を生成する
 * 年全体オプション（YYYY年）+ 月オプション（1〜12月）
 * @returns {Array<string>} 選択肢の配列
 */
function generateMonthlySummaryOptions() {
  var year = getTargetYearFromFormResponses_();
  var options = [year + '年'];
  for (var m = 1; m <= 12; m++) {
    options.push(year + '年' + m + '月');
  }
  return options;
}

/**
 * 来館カレンダー用の選択肢を生成する
 * 年全体オプション（YYYY年）+ 月オプション（1〜12月）
 * @returns {Array<string>} 選択肢の配列
 */
function generateCalendarOptions() {
  var year = getTargetYearFromFormResponses_();
  var options = [year + '年'];
  for (var m = 1; m <= 12; m++) {
    options.push(year + '年' + m + '月');
  }
  return options;
}

/**
 * 児童別ビュー用の選択肢を生成する
 * すべて・年全体（YYYY年）+ 月オプション（1〜12月）
 * @returns {Array<string>} 選択肢の配列
 */
function generateChildViewOptions() {
  var year = getTargetYearFromFormResponses_();
  var options = ['すべて', year + '年'];
  for (var m = 1; m <= 12; m++) {
    options.push(year + '年' + m + '月');
  }
  return options;
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
 * 児童別ビュー用: 全児童名を取得する（在籍中を先頭、退所済みを後半に配置）
 * @returns {Array<string>} 児童名の配列
 */
function getAllChildNameOptions() {
  var allChildren = getChildMasterData();
  var active = [];
  var inactive = [];

  allChildren.forEach(function(row) {
    var name = row[MASTER_COL.NAME - 1];
    var status = row[MASTER_COL.ENROLLMENT - 1];
    if (name) {
      if (status === '稼働' || status === '休止') {
        active.push(name);
      } else {
        inactive.push(name);
      }
    }
  });

  return active.concat(inactive);
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

/**
 * 確定来館記録から指定年のデータを取得する
 * @param {number} year 年
 * @returns {Array<Array>} 該当年の確定来館記録データ
 */
function getConfirmedVisitsByYear(year) {
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
    return recordDate.getFullYear() === year;
  });
}

/**
 * 確定来館記録から全期間のデータを取得する
 * @returns {Array<Array>} 全期間の確定来館記録データ
 */
function getAllConfirmedVisits() {
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
  return data.slice(1);
}

/**
 * 入所日時・退所日時から宿泊日数を返す
 * 同日=1、1泊2日=2、2泊3日=3
 * @param {Date} checkIn 入所日時
 * @param {Date} checkOut 退所日時
 * @returns {number} 宿泊日数（最小1）
 */
function calcStayDays_(checkIn, checkOut) {
  if (!(checkIn instanceof Date) || !(checkOut instanceof Date)) return 1;
  var inMid = new Date(checkIn.getFullYear(), checkIn.getMonth(), checkIn.getDate());
  var outMid = new Date(checkOut.getFullYear(), checkOut.getMonth(), checkOut.getDate());
  var diff = Math.round((outMid - inMid) / 86400000);
  return Math.max(1, diff + 1);
}

/**
 * 入所日時・退所日時からスティ期間の各日の Date 配列を返す
 * @param {Date} checkIn 入所日時
 * @param {Date} checkOut 退所日時
 * @returns {Array<Date>} 各宿泊日の Date 配列
 */
function expandStayToDates_(checkIn, checkOut) {
  var n = calcStayDays_(checkIn, checkOut);
  var base = (checkIn instanceof Date) ? checkIn : new Date(checkIn);
  var result = [];
  for (var i = 0; i < n; i++) {
    result.push(new Date(base.getFullYear(), base.getMonth(), base.getDate() + i));
  }
  return result;
}

/**
 * 設定シートのC列から指定行の値を取得する
 * @param {number} row 行番号
 * @returns {*} 設定値（シートが存在しない場合はnull）
 */
function getSettingValue_(row) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
    if (!sheet) return null;
    return sheet.getRange(row, 3).getValue();
  } catch (e) {
    return null;
  }
}

/**
 * 設定シートから営業日をDay番号配列で取得する
 * 未設定の場合は空配列を返す（全曜日を対象とみなす）
 * @returns {Array<number>} Day番号の配列（0=日, 1=月, ... 6=土）
 */
function getBusinessDays() {
  var val = getSettingValue_(SETTINGS_ROW.BUSINESS_DAYS);
  if (!val) return [];
  return parseVisitDays_(String(val));
}

/**
 * 年間利用枠を月別平日数の比率で按分し、月次利用枠を算出する
 * 月ごとの平日数（月〜金）の差によって 14〜16 程度のばらつきが生まれ、
 * 12ヶ月合計は annualQuota と近似するが必ずしも一致しない（要件通り）
 * @param {number} annualQuota 年間利用枠
 * @param {number} year 年
 * @param {number} month 月（1-12）
 * @returns {number} 月次利用枠
 */
function calcMonthlyQuota_(annualQuota, year, month) {
  var weekdaysThisMonth = countWeekdaysInMonth_(year, month);
  var totalWeekdaysInYear = 0;
  for (var m = 1; m <= 12; m++) {
    totalWeekdaysInYear += countWeekdaysInMonth_(year, m);
  }
  if (totalWeekdaysInYear <= 0) return Math.floor(annualQuota / 12);
  return Math.round(annualQuota * weekdaysThisMonth / totalWeekdaysInYear);
}

/**
 * 指定年月の平日数（月〜金）を返す
 * @param {number} year 年
 * @param {number} month 月（1-12）
 * @returns {number} 平日数
 */
function countWeekdaysInMonth_(year, month) {
  var count = 0;
  var daysInMonth = new Date(year, month, 0).getDate();
  for (var d = 1; d <= daysInMonth; d++) {
    var dow = new Date(year, month - 1, d).getDay();
    if (dow >= 1 && dow <= 5) count++;
  }
  return count;
}

/**
 * "YYYY年" 形式から年を抽出する（"YYYY年M月" にはマッチしない）
 * @param {string} str 入力文字列
 * @returns {number|null} 年（マッチしない場合null）
 */
function parseYearOnly_(str) {
  if (!str || str instanceof Date) return null;
  var match = String(str).match(/^(\d{4})年$/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * 設定シートから固定スタッフ名を取得する（振り分け・スタッフ2補完用）
 * @returns {string} 固定スタッフ名（未設定の場合は空文字）
 */
function getDummyStaffName_() {
  return String(getSettingValue_(SETTINGS_ROW.DUMMY_STAFF_NAME) || '');
}

/**
 * エラーをログシートに記録する
 * @param {string} functionName エラーが発生した関数名
 * @param {Error} error エラーオブジェクト
 */
function logError_(functionName, error) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.LOG);

    // ログシートが存在しない場合は自動作成
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAMES.LOG);
      sheet.getRange(1, 1, 1, 4).setValues([['タイムスタンプ', '関数名', 'エラーメッセージ', 'スタックトレース']]);
      sheet.getRange(1, 1, 1, 4)
        .setBackground('#4285F4')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 200);
      sheet.setColumnWidth(3, 400);
      sheet.setColumnWidth(4, 400);
    }

    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    var message = error.message || String(error);
    var stack = error.stack || '';

    sheet.appendRow([timestamp, functionName, message, stack]);
  } catch (e) {
    Logger.log('ログ出力に失敗: ' + e.message);
  }
}
