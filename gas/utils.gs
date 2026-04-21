/**
 * 共通ユーティリティ（定数・ヘルパー関数）
 */

// スタッフ用_回答シートのスプレッドシートID（メール送信済フラグ書き込み先）
const STAFF_SHEET_ID = '1vj9pU6IMSkQmV7L6aNhKxFaPzj55p2NmO1wOnyaaywc';

// シート名
const SHEET_NAMES = {
  FORM_RESPONSE: 'フォームの回答',
  CHILD_MASTER: '児童マスタ',
  STAFF_MASTER: 'スタッフマスタ',
  MONTHLY_SUMMARY: '月別集計',
  VISIT_CALENDAR: '来館カレンダー',
  CONFIRMED_VISITS: '確定来館記録',
  CHILD_VIEW: '児童別ビュー',
  LOG: 'ログ',
  SETTINGS: '設定',
  BOUNCE_LOG: 'バウンスログ',
};

// 設定シートの行インデックス（C列=デフォルト値）。実シートの並びに合わせる
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
  DUMMY_STAFF_NAME: 13,   // 固定スタッフ（振り分け・スタッフ2補完用）
  ERROR_EMAIL: 14,        // エラー通知先メール
  EMAIL_SUBJECT: 15,      // メール件名
  EMAIL_BODY: 16,         // メール本文
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
  VISIT_DAYS: 11,       // 来館曜日
  NON_VISIT_DAYS: 12,   // 非来館曜日
  ENROLLMENT: 13,       // 入所状況（稼働/休止/退所）
  DEPARTURE_STATUS: 14, // 退所状況（別施設移動/別施設移動無）
  MASTER_NOTES: 15,     // 備考
};

// フォームの回答シートのデータ開始行（1行目=ヘッダー）
const FORM_DATA_START_ROW = 2;

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

// 振り分け補完用デフォルト値（設定シート未設定時のフォールバック）
const ALLOCATION_DEFAULTS = {
  CHECK_IN: '17:00',
  CHECK_OUT: '08:00',
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

// 確定来館記録のレイアウト定数（1行目=年、2行目=月、3行目=ヘッダー、4行目〜=データ）
const CONFIRMED_HEADER_ROW = 3;
const CONFIRMED_DATA_START_ROW = 4;

// 月別集計のレイアウト定数（1行目=年、2行目=月、3行目=ヘッダー、4行目〜=データ）
const SUMMARY_HEADER_ROW = 3;
const SUMMARY_DATA_START_ROW = 4;

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
    return (status === '稼働' || status === '休止');
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
 * 入所日時〜退所日時の滞在期間が対象月と重なるレコードも返す（月またぎ連泊対応）
 * @param {number} year 年
 * @param {number} month 月（1-12）
 * @returns {Array<Array>} 該当月のフォーム回答データ
 */
function getFormResponsesByMonth(year, month) {
  const sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
  const data = sheet.getDataRange().getValues();
  var responses = data.slice(1);
  var monthStart = new Date(year, month - 1, 1);
  var monthEnd = new Date(year, month, 0); // 月末日

  return responses.filter(function(row) {
    var checkIn = row[FORM_COL.CHECK_IN - 1];
    var checkOut = row[FORM_COL.CHECK_OUT - 1];

    // 入所/退所日時が両方フル日時（1900年以降）なら滞在期間と対象月の重なりで判定
    if (checkIn instanceof Date && checkIn.getFullYear() >= 1900) {
      var stayStart = new Date(checkIn.getFullYear(), checkIn.getMonth(), checkIn.getDate());
      var stayEnd = (checkOut instanceof Date && checkOut.getFullYear() >= 1900)
        ? new Date(checkOut.getFullYear(), checkOut.getMonth(), checkOut.getDate())
        : stayStart;
      return stayStart <= monthEnd && stayEnd >= monthStart;
    }

    // 時刻のみ・不正値の場合は記録日で判定
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
 * フォームの回答から対象年のリストを降順で取得する
 * データがない場合は現在の年1件のみを返す
 * @returns {Array<number>} 年の配列（降順）
 */
function collectYearsFromFormResponses_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.FORM_RESPONSE);
  if (!sheet) return [new Date().getFullYear()];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [new Date().getFullYear()];
  var yearsSet = {};
  data.slice(1).forEach(function(row) {
    var d = new Date(row[FORM_COL.RECORD_DATE - 1]);
    if (!isNaN(d.getTime())) yearsSet[d.getFullYear()] = true;
  });
  var years = Object.keys(yearsSet).map(Number);
  if (years.length === 0) return [new Date().getFullYear()];
  return years.sort(function(a, b) { return b - a; });
}

/**
 * 年ドロップダウンの選択肢を生成する
 * @returns {Array<string>} ["すべて", "2026年", "2025年", ...]
 */
function generateYearOptions() {
  var years = collectYearsFromFormResponses_();
  return ['すべて'].concat(years.map(function(y) { return y + '年'; }));
}

/**
 * 月ドロップダウンの選択肢を生成する
 * @returns {Array<string>} ["すべて", "1月", ..., "12月"]
 */
function generateMonthOptions() {
  var options = ['すべて'];
  for (var m = 1; m <= 12; m++) options.push(m + '月');
  return options;
}

/**
 * 年ドロップダウン値をパースする
 * @param {*} str 値（"すべて"・"2026年"・Date等）
 * @returns {number|null} 年（「すべて」または不正値は null）
 */
function parseYearOption_(str) {
  if (!str || str === 'すべて') return null;
  if (str instanceof Date) return str.getFullYear();
  var match = String(str).match(/^(\d{4})年$/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * 月ドロップダウン値をパースする
 * @param {*} str 値（"すべて"・"3月"等）
 * @returns {number|null} 月（1-12、「すべて」は null）
 */
function parseMonthOption_(str) {
  if (!str || str === 'すべて') return null;
  var match = String(str).match(/^(\d{1,2})月$/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * 年・月ドロップダウン値から期間スコープを構築する
 * @param {*} yearStr B1の値
 * @param {*} monthStr B2の値
 * @returns {{type: string, year?: number, month?: number}}
 */
function buildScope_(yearStr, monthStr) {
  var year = parseYearOption_(yearStr);
  var month = parseMonthOption_(monthStr);
  if (year !== null && month !== null) return { type: 'month', year: year, month: month };
  if (year !== null) return { type: 'year', year: year };
  if (month !== null) return { type: 'month_all_years', month: month };
  return { type: 'all' };
}

/**
 * ログ・メッセージ用にスコープを文字列化する
 */
function describeScope_(scope) {
  if (scope.type === 'month') return scope.year + '年' + scope.month + '月';
  if (scope.type === 'year') return scope.year + '年';
  if (scope.type === 'month_all_years') return '全期間の' + scope.month + '月';
  return '全期間';
}

/**
 * Dateが期間スコープに含まれるか判定する
 * @param {Date} date 判定対象の日付
 * @param {{type: string, year?: number, month?: number}} scope
 * @returns {boolean}
 */
function matchesScope_(date, scope) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return false;
  if (scope.type === 'month') {
    return date.getFullYear() === scope.year && (date.getMonth() + 1) === scope.month;
  }
  if (scope.type === 'year') {
    return date.getFullYear() === scope.year;
  }
  if (scope.type === 'month_all_years') {
    return (date.getMonth() + 1) === scope.month;
  }
  return true; // all
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
  if (data.length < CONFIRMED_DATA_START_ROW) {
    return [];
  }
  var records = data.slice(CONFIRMED_DATA_START_ROW - 1);
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
  if (data.length < CONFIRMED_DATA_START_ROW) {
    return [];
  }
  var records = data.slice(CONFIRMED_DATA_START_ROW - 1);
  return records.filter(function(row) {
    var recordDate = new Date(row[CONFIRMED_COL.RECORD_DATE - 1]);
    return recordDate.getFullYear() === year;
  });
}

/**
 * 確定来館記録からスコープに合致するデータを取得する
 * @param {{type: string, year?: number, month?: number}} scope
 * @returns {Array<Array>} 該当する確定来館記録データ
 */
function getConfirmedVisitsByScope(scope) {
  var all = getAllConfirmedVisits();
  if (scope.type === 'all') return all;
  return all.filter(function(row) {
    return matchesScope_(new Date(row[CONFIRMED_COL.RECORD_DATE - 1]), scope);
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
  if (data.length < CONFIRMED_DATA_START_ROW) {
    return [];
  }
  return data.slice(CONFIRMED_DATA_START_ROW - 1);
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
 * 基準日（記録日）と入所/退所日時から宿泊日の Date 配列を返す
 * - 入所/退所が時刻のみ（1899/12/30基準）の場合でも基準日を軸に正しい日付を生成する
 * - 入所日時がフルDate、かつ退所日時もフルDateなら両者の日付差で展開
 * @param {Date|string} recordDate 基準日（フォームの「記録日」列）
 * @param {Date} checkIn 入所日時
 * @param {Date} checkOut 退所日時
 * @returns {Array<Date>} 各宿泊日の Date 配列
 */
function expandStayToDates_(recordDate, checkIn, checkOut) {
  var base = (recordDate instanceof Date) ? recordDate : new Date(recordDate);
  if (isNaN(base.getTime())) {
    // recordDate が使えない場合は checkIn にフォールバック（後方互換）
    if (checkIn instanceof Date && !isNaN(checkIn.getTime())) {
      base = checkIn;
    } else {
      return [];
    }
  }

  // 入所日時・退所日時が両方フルDate（1900年以降）なら差分で日数算出
  // それ以外（時刻のみ等）は 1 日として扱う
  var n = 1;
  if (checkIn instanceof Date && checkOut instanceof Date
      && checkIn.getFullYear() >= 1900 && checkOut.getFullYear() >= 1900) {
    var inMid = new Date(checkIn.getFullYear(), checkIn.getMonth(), checkIn.getDate());
    var outMid = new Date(checkOut.getFullYear(), checkOut.getMonth(), checkOut.getDate());
    var diff = Math.round((outMid - inMid) / 86400000);
    n = Math.max(1, diff + 1);
  }

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
 * 設定シートから固定スタッフ名を取得する（振り分け・スタッフ2補完用）
 * @returns {string} 固定スタッフ名（未設定の場合は空文字）
 */
function getDummyStaffName_() {
  return String(getSettingValue_(SETTINGS_ROW.DUMMY_STAFF_NAME) || '');
}

/**
 * 振り分け補完値を設定シートから取得する
 * 設定シート未設定の項目は ALLOCATION_DEFAULTS にフォールバックする
 * @returns {Object} 補完値オブジェクト
 */
function getAllocationDefaultsFromSettings_() {
  function pick(row, fallback) {
    var v = getSettingValue_(row);
    return (v === null || v === '' || typeof v === 'undefined') ? fallback : v;
  }
  return {
    CHECK_IN: pick(SETTINGS_ROW.CHECK_IN, ALLOCATION_DEFAULTS.CHECK_IN),
    CHECK_OUT: pick(SETTINGS_ROW.CHECK_OUT, ALLOCATION_DEFAULTS.CHECK_OUT),
    TEMPERATURE: pick(SETTINGS_ROW.TEMPERATURE, ALLOCATION_DEFAULTS.TEMPERATURE),
    MEAL: pick(SETTINGS_ROW.MEAL, ALLOCATION_DEFAULTS.MEAL),
    BATH: pick(SETTINGS_ROW.BATH, ALLOCATION_DEFAULTS.BATH),
    SLEEP: pick(SETTINGS_ROW.SLEEP, ALLOCATION_DEFAULTS.SLEEP),
    BOWEL: pick(SETTINGS_ROW.BOWEL, ALLOCATION_DEFAULTS.BOWEL),
    MEDICINE: pick(SETTINGS_ROW.MEDICINE, ALLOCATION_DEFAULTS.MEDICINE),
    NOTES: pick(SETTINGS_ROW.NOTES, ALLOCATION_DEFAULTS.NOTES),
  };
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
