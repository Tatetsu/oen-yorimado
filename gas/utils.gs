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
  NOTES_MASTER: '定型文マスタ',
  ALLOWED_USERS: '許可ユーザー',
};

// 許可ユーザーシートの列インデックス（1始まり）
// レイアウト: Row 1〜5=マニュアル/編集画面URL/TOKEN/ランダム文字列など（ユーザー編集領域）,
//             Row 6=ヘッダー, Row 7〜=データ
// 編集画面URL は B2 セルに記載される想定。
// D列(URL)はARRAYFORMULAで自動生成される（B2のURLとC列のトークンを結合）
const ALLOWED_USERS_COL = {
  EMAIL: 1,
  NAME: 2,
  TOKEN: 3,
  URL: 4,
  ACTIVE: 5,
  NOTE: 6,
};
const ALLOWED_USERS_HEADER_ROW = 6;
const ALLOWED_USERS_DATA_START_ROW = 7;
const ALLOWED_USERS_BASE_URL_CELL = 'B2'; // 編集画面URLのセル参照

// 設定シートの行インデックス（C列=デフォルト値）。実シートの並びに合わせる
const SETTINGS_ROW = {
  MAX_VISITS_PER_DAY: 2,  // 1日最大来館数
  CHECK_IN: 3,            // 入所時間
  CHECK_OUT: 4,           // 退所時間
  BUSINESS_DAYS: 5,       // 営業日
  TEMPERATURE: 6,         // 体温
  MEAL_DINNER: 7,         // 夕食
  MEAL_BREAKFAST: 8,      // 朝食
  MEAL_LUNCH: 9,          // 昼食
  BATH: 10,               // 入浴
  SLEEP: 11,              // 睡眠
  BOWEL: 12,              // 便
  MEDICINE_MORNING: 13,   // 服薬（朝）
  MEDICINE_NIGHT: 14,     // 服薬（夜）
  NOTES: 15,              // 連絡事項
  DUMMY_STAFF_NAME: 16,   // 固定スタッフ（振り分け・スタッフ2補完用）
  ERROR_EMAIL: 17,        // エラー通知先メール（カンマ区切り）
  EMAIL_SUBJECT: 18,      // メール件名
  EMAIL_BODY: 19,         // メール本文
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
// 行政様式の実績記録票に合わせ、食事は夕/朝/昼、服薬は夜/朝に分離する
// CHECK_OUT は「退所予定日時」として扱う（連泊初日・中日は空欄可）
// OVERNIGHT_FLAG: 連泊フラグ（true=連泊、false/空欄=単泊）。その他連絡事項の直後（メール送信の直前）
// 確定来館記録の列順と2列目以降が一致するレイアウト
const FORM_COL = {
  TIMESTAMP: 1,
  RECORD_DATE: 2,
  STAFF_NAME: 3,           // スタッフ1
  STAFF_NAME_2: 4,         // スタッフ2(任意)
  CHILD_NAME: 5,
  CHECK_IN: 6,             // 入所日時(連泊・最終日は空欄可)
  CHECK_OUT: 7,            // 退所予定日時(連泊・初日/中日は空欄可)
  TEMPERATURE: 8,
  MEAL_DINNER: 9,          // 夕食
  MEAL_BREAKFAST: 10,      // 朝食
  MEAL_LUNCH: 11,          // 昼食
  BATH: 12,
  SLEEP: 13,
  BOWEL: 14,
  MEDICINE_NIGHT: 15,      // 服薬(夜)
  MEDICINE_MORNING: 16,    // 服薬(朝)
  NOTES: 17,
  OVERNIGHT_FLAG: 18,      // 連泊フラグ(true/false)
  EMAIL_SENT: 19,          // メール送信済(システム管理・初回送信時に追加)
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
// 食事・服薬は分離項目それぞれに同じデフォルト値を適用する
const ALLOCATION_DEFAULTS = {
  CHECK_IN: '17:00',
  CHECK_OUT: '08:00',
  TEMPERATURE: 36.5,
  MEAL_DINNER: '○',
  MEAL_BREAKFAST: '○',
  MEAL_LUNCH: '−',
  BATH: '○',
  SLEEP: '○',
  BOWEL: '○',
  MEDICINE_NIGHT: '○',
  MEDICINE_MORNING: '○',
  NOTES: '特になし',
};

// 1日あたりの最大来館数（設定シート未設定時のフォールバック）
const DEFAULT_MAX_VISITS_PER_DAY = 7;

// 曜日文字列 → 数値変換マップ（Date.getDay()に対応）
const DAY_OF_WEEK_MAP = {
  '日': 0, '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6,
};

// 確定来館記録の列インデックス（1始まり）
// FORM_COL と同じ粒度で食事・服薬を分離する
// 1列目=DATA_TYPE はフォームの TIMESTAMP と同位置。
// CHECK_OUT の直後に PICKUP_OUTBOUND/PICKUP_RETURN（送迎の往/復）が入る
//   → 月別の単純合計で送迎加算カウントが取れる（月またぎ・連泊にも対応）
// OVERNIGHT_FLAG: 連泊フラグ（入所日と退所日が異なる行で true）
// STAY_PK: 宿泊単位のユニークキー（児童名|入所日時ISO）。最終列・既定で非表示。
const CONFIRMED_COL = {
  DATA_TYPE: 1,
  RECORD_DATE: 2,
  STAFF_NAME: 3,           // スタッフ1
  STAFF_NAME_2: 4,         // スタッフ2（任意）
  CHILD_NAME: 5,
  CHECK_IN: 6,             // 入所日時（ペアリング後の論理1宿泊の値）
  CHECK_OUT: 7,            // 退所予定日時（ペアリング後の論理1宿泊の値）
  PICKUP_OUTBOUND: 8,      // 送迎(往): 記録日が入所日と一致 → 1
  PICKUP_RETURN: 9,        // 送迎(復): 記録日が退所予定日と一致 → 1
  TEMPERATURE: 10,
  MEAL_DINNER: 11,         // 夕食
  MEAL_BREAKFAST: 12,      // 朝食
  MEAL_LUNCH: 13,          // 昼食
  BATH: 14,
  SLEEP: 15,
  BOWEL: 16,
  MEDICINE_NIGHT: 17,      // 服薬(夜)
  MEDICINE_MORNING: 18,    // 服薬(朝)
  NOTES: 19,
  OVERNIGHT_FLAG: 20,      // 連泊フラグ（true/false）
  STAY_PK: 21,             // 宿泊PK（児童名|入所日時ISO）
};

// メール件名のデフォルト値（設定シート未設定時のフォールバック）
const DEFAULT_EMAIL_SUBJECT = '【テスト施設　来館記録のお知らせ】';

// メール本文テンプレートのデフォルト値（設定シート未設定時のフォールバック）
const DEFAULT_EMAIL_TEMPLATE = [
  '{保護者名} 様',
  '',
  'いつもお世話になっております。',
  'Yorimadoです。',
  '',
  '{日付}の{児童名}さんの来館記録をお知らせいたします。',
  '',
  '■ 来館記録',
  '・入所時間: {入所時間}',
  '・退所時間: {退所時間}',
  '・体温: {体温}℃',
  '・夕食: {夕食}',
  '・朝食: {朝食}',
  '・昼食: {昼食}',
  '・入浴: {入浴}',
  '・睡眠: {睡眠}',
  '・便: {便}',
  '・服薬(夜): {服薬(夜)}',
  '・服薬(朝): {服薬(朝)}',
  '',
  '■ 連絡事項',
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

// ========================================
// 汎用ユーティリティ（書式・検証・トリガー・日付）
// ========================================

/**
 * ヘッダー行に共通書式（青背景・白太字）と行固定をまとめて適用する
 * 色を変えたい場合は options で上書きする（バウンスログの赤ヘッダー等）
 * @param {GoogleAppsScript.Spreadsheet.Range} range 対象セル範囲
 * @param {number} [frozenRows] 指定があれば setFrozenRows する
 * @param {{bgColor?:string, fontColor?:string, horizontalAlignment?:string}} [options]
 */
function styleSheetHeader_(range, frozenRows, options) {
  options = options || {};
  var bg = options.bgColor || '#4285F4';
  var fc = options.fontColor || '#FFFFFF';
  range.setBackground(bg).setFontColor(fc).setFontWeight('bold');
  if (options.horizontalAlignment) {
    range.setHorizontalAlignment(options.horizontalAlignment);
  }
  if (frozenRows && frozenRows > 0) {
    range.getSheet().setFrozenRows(frozenRows);
  }
}

/**
 * 列幅をマップで一括設定する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} widthMap {列番号: 幅px} 例: {1:40, 2:100}
 */
function setColumnWidths_(sheet, widthMap) {
  Object.keys(widthMap).forEach(function(col) {
    sheet.setColumnWidth(parseInt(col, 10), widthMap[col]);
  });
}

/**
 * ドロップダウン（requireValueInList）のデータ検証を設定する
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {Array} options 選択肢
 * @param {boolean} [showDropdown=true] セル内に▼を表示するか
 */
function setListValidation_(range, options, showDropdown) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, showDropdown !== false)
    .build();
  range.setDataValidation(rule);
}

/**
 * Date値が有効か判定する（Date型かつNaNでない）
 * @param {*} d
 * @returns {boolean}
 */
function isValidDate_(d) {
  return (d instanceof Date) && !isNaN(d.getTime());
}

/**
 * Date を指定フォーマットで文字列化する（不正値は空文字を返す）
 * @param {*} date Date値
 * @param {string} [format='yyyy/MM/dd'] フォーマット
 * @param {string} [tz] タイムゾーン（省略時は Session.getScriptTimeZone()）
 * @returns {string}
 */
function formatDateYMD_(date, format, tz) {
  if (!isValidDate_(date)) return '';
  return Utilities.formatDate(date, tz || Session.getScriptTimeZone(), format || 'yyyy/MM/dd');
}

/**
 * 日付を YYYY-MM-DD 形式のキー文字列にする（マップのキー用）
 * @param {Date} date
 * @returns {string}
 */
function formatDateKey_(date) {
  var y = date.getFullYear();
  var m = ('0' + (date.getMonth() + 1)).slice(-2);
  var d = ('0' + date.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
}

/**
 * 児童マスタを「児童名 → 行データ」のマップに索引化する
 * @param {string} [filter] 'active'=稼働/休止のみ、未指定=全件
 * @returns {Object}
 */
function buildChildNameToRowMap_(filter) {
  var data = (filter === 'active') ? getActiveChildren() : getChildMasterData();
  var map = {};
  data.forEach(function(row) {
    var name = row[MASTER_COL.NAME - 1];
    if (name) map[name] = row;
  });
  return map;
}

/**
 * 時間ベースのトリガーを再作成する（同名ハンドラの既存トリガーは削除）
 * @param {string} handlerName 関数名
 * @param {{everyDays?:number, onMonthDay?:number, atHour?:number}} schedule
 *   - everyDays: 毎N日 / onMonthDay: 毎月N日（どちらか片方）
 *   - atHour: 実行時刻（時）
 */
function setupTimeTrigger_(handlerName, schedule) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
    }
  });
  var builder = ScriptApp.newTrigger(handlerName).timeBased();
  if (schedule.onMonthDay) {
    builder = builder.onMonthDay(schedule.onMonthDay);
  } else {
    builder = builder.everyDays(schedule.everyDays || 1);
  }
  if (typeof schedule.atHour === 'number') {
    builder = builder.atHour(schedule.atHour);
  }
  builder.create();
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
 * 入所日時〜退所日時の滞在期間が対象年と重なるレコードも返す（年またぎ連泊対応）
 * 児童名+入所日時でペアリングした論理1宿泊が対象年と重なる場合に構成全レコードを返す。
 * @param {number} year 年
 * @returns {Array<Array>} 該当年のフォーム回答データ
 */
function getFormResponsesByYear(year) {
  var allResponses = getFormResponsesAll_();
  var stays = pairStayRecords_(allResponses);
  var yearStart = new Date(year, 0, 1);
  var yearEnd = new Date(year, 11, 31);

  var includeRows = [];
  var seen = {};

  stays.forEach(function(stay) {
    var stayStart = stay.checkIn ? new Date(stay.checkIn.getFullYear(), stay.checkIn.getMonth(), stay.checkIn.getDate()) : null;
    var stayEnd = stay.checkOut ? new Date(stay.checkOut.getFullYear(), stay.checkOut.getMonth(), stay.checkOut.getDate()) : null;

    var overlaps = false;
    if (stayStart && stayEnd) {
      overlaps = stayStart <= yearEnd && stayEnd >= yearStart;
    } else if (stayStart) {
      overlaps = stayStart.getFullYear() === year;
    } else if (stayEnd) {
      overlaps = stayEnd.getFullYear() === year;
    } else {
      overlaps = stay.recordDate.getFullYear() === year;
    }

    if (overlaps) {
      stay.sourceRows.forEach(function(row) {
        var key = row[FORM_COL.TIMESTAMP - 1] + '|' + row[FORM_COL.CHILD_NAME - 1];
        if (!seen[key]) {
          seen[key] = true;
          includeRows.push(row);
        }
      });
    }
  });

  return includeRows;
}

/**
 * フォームの回答から指定年月のデータを取得する
 * 入所日時〜退所日時の滞在期間が対象月と重なるレコードも返す（月またぎ連泊対応）
 * 連泊レコード（入所のみ/退所のみ/両方空欄）は児童名でペアリングし、
 * ペアリング後の論理1宿泊が対象月と重なる場合に構成全レコードを返す。
 * @param {number} year 年
 * @param {number} month 月（1-12）
 * @returns {Array<Array>} 該当月のフォーム回答データ
 */
function getFormResponsesByMonth(year, month) {
  var allResponses = getFormResponsesAll_();
  var stays = pairStayRecords_(allResponses);
  var monthStart = new Date(year, month - 1, 1);
  var monthEnd = new Date(year, month, 0);

  var includeRows = [];
  var seen = {};

  stays.forEach(function(stay) {
    // ペアリング後の宿泊期間が対象月と重なるか判定
    var stayStart = stay.checkIn ? new Date(stay.checkIn.getFullYear(), stay.checkIn.getMonth(), stay.checkIn.getDate()) : null;
    var stayEnd = stay.checkOut ? new Date(stay.checkOut.getFullYear(), stay.checkOut.getMonth(), stay.checkOut.getDate()) : null;

    var overlaps = false;
    if (stayStart && stayEnd) {
      overlaps = stayStart <= monthEnd && stayEnd >= monthStart;
    } else if (stayStart) {
      overlaps = stayStart <= monthEnd && stayStart >= monthStart;
    } else if (stayEnd) {
      overlaps = stayEnd <= monthEnd && stayEnd >= monthStart;
    } else {
      // 両方空欄（連泊中日のみで構成）→ 記録日で判定
      overlaps = stay.recordDate.getFullYear() === year && (stay.recordDate.getMonth() + 1) === month;
    }

    if (overlaps) {
      stay.sourceRows.forEach(function(row) {
        var key = row[FORM_COL.TIMESTAMP - 1] + '|' + row[FORM_COL.CHILD_NAME - 1];
        if (!seen[key]) {
          seen[key] = true;
          includeRows.push(row);
        }
      });
    }
  });

  return includeRows;
}

/**
 * フォームの回答から全データを取得する（ヘッダー除く）
 * @returns {Array<Array>} 全フォーム回答データ
 */
function getFormResponsesAll_() {
  var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
  var data = sheet.getDataRange().getValues();
  return data.slice(1);
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
 * - 入所/退所が両方フルDateなら入所日付を基準に滞在カレンダー全日を展開（月またぎ対応）
 * - 入所のみフルDateなら入所日のみ返す
 * - 両方時刻のみなら recordDate を基準に1日として扱う（後方互換）
 * @param {Date|string} recordDate 基準日（フォームの「記録日」列）
 * @param {Date} checkIn 入所日時
 * @param {Date} checkOut 退所日時
 * @returns {Array<Date>} 各宿泊日の Date 配列
 */
function expandStayToDates_(recordDate, checkIn, checkOut) {
  var hasCheckIn = (checkIn instanceof Date) && checkIn.getFullYear() >= 1900;
  var hasCheckOut = (checkOut instanceof Date) && checkOut.getFullYear() >= 1900;

  // 両方フルDate: 入所日基準で滞在カレンダー全日を展開
  if (hasCheckIn && hasCheckOut) {
    var base = new Date(checkIn.getFullYear(), checkIn.getMonth(), checkIn.getDate());
    var outMid = new Date(checkOut.getFullYear(), checkOut.getMonth(), checkOut.getDate());
    var diff = Math.round((outMid - base) / 86400000);
    var n = Math.max(1, diff + 1);
    var result = [];
    for (var i = 0; i < n; i++) {
      result.push(new Date(base.getFullYear(), base.getMonth(), base.getDate() + i));
    }
    return result;
  }

  // 入所のみフルDate: その日付のみ返す
  if (hasCheckIn && !hasCheckOut) {
    return [new Date(checkIn.getFullYear(), checkIn.getMonth(), checkIn.getDate())];
  }

  // 退所のみフルDate: その日付のみ返す
  if (!hasCheckIn && hasCheckOut) {
    return [new Date(checkOut.getFullYear(), checkOut.getMonth(), checkOut.getDate())];
  }

  // 両方時刻のみ・不正値: recordDate を基準に1日として扱う（後方互換）
  var base2 = (recordDate instanceof Date) ? recordDate : new Date(recordDate);
  if (isNaN(base2.getTime())) return [];
  return [new Date(base2.getFullYear(), base2.getMonth(), base2.getDate())];
}

/**
 * フォーム回答を「児童名+入所日時」のユニークキーで論理1宿泊にペアリングする
 *
 * 新仕様（ユニーク宿泊キー方式）:
 *   - 全レコードに入所日時・退所日時を記入する運用
 *   - 児童名+入所日時が同じレコードは同一宿泊（中日の様子記録は同じ入退所を共有）
 *   - 状態機械によるペアリングは不要
 *   - プライマリ行 = 記録日と入所日が一致する行（無ければグループ内で記録日が最も早い行）
 *
 * @param {Array<Array>} formResponses フォーム回答データ（ヘッダー除く）
 * @returns {Array<{
 *   childName: string,
 *   checkIn: Date|null,
 *   checkOut: Date|null,
 *   recordDate: Date,
 *   isOvernight: boolean,
 *   sourceRows: Array<Array>,
 *   primaryRow: Array,
 *   issues: Array<string>,
 *   stayPk: string
 * }>} 論理1宿泊のリスト
 */
function pairStayRecords_(formResponses) {
  var groups = {};

  formResponses.forEach(function(row) {
    var name = row[FORM_COL.CHILD_NAME - 1];
    if (!name) return;

    var checkIn = row[FORM_COL.CHECK_IN - 1];
    var checkOut = row[FORM_COL.CHECK_OUT - 1];
    var recordDate = new Date(row[FORM_COL.RECORD_DATE - 1]);
    var hasCheckIn = (checkIn instanceof Date) && checkIn.getFullYear() >= 1900;
    var hasCheckOut = (checkOut instanceof Date) && checkOut.getFullYear() >= 1900;

    // 入所日時が無いレコードは記録日+タイムスタンプで独立した1宿泊として扱う（移行データ吸収用）
    var key;
    if (hasCheckIn) {
      key = buildStayPk_(name, checkIn);
    } else {
      var ts = row[FORM_COL.TIMESTAMP - 1];
      key = name + '|NOIN|' +
        (recordDate instanceof Date ? recordDate.getTime() : String(recordDate)) + '|' +
        (ts instanceof Date ? ts.getTime() : String(ts || ''));
    }

    if (!groups[key]) {
      groups[key] = {
        childName: name,
        checkIn: hasCheckIn ? checkIn : null,
        checkOut: hasCheckOut ? checkOut : null,
        recordDate: recordDate,
        sourceRows: [row],
        issues: [],
      };
    } else {
      var g = groups[key];
      g.sourceRows.push(row);
      // 入退所日時が後続レコードで埋まる場合は採用（運用上同一値であるべきだが、片方欠損レコード救済）
      if (!g.checkIn && hasCheckIn) g.checkIn = checkIn;
      if (!g.checkOut && hasCheckOut) g.checkOut = checkOut;
    }
  });

  var stays = Object.keys(groups).map(function(key) {
    var g = groups[key];
    var primary = pickPrimaryRow_(g.sourceRows, g.checkIn);
    var primaryRecordDate = new Date(primary[FORM_COL.RECORD_DATE - 1]);
    var isOvernight = false;
    if (g.checkIn && g.checkOut) {
      var inDay = new Date(g.checkIn.getFullYear(), g.checkIn.getMonth(), g.checkIn.getDate());
      var outDay = new Date(g.checkOut.getFullYear(), g.checkOut.getMonth(), g.checkOut.getDate());
      isOvernight = inDay.getTime() !== outDay.getTime();
    }
    return {
      childName: g.childName,
      checkIn: g.checkIn,
      checkOut: g.checkOut,
      recordDate: isValidDate_(primaryRecordDate) ? primaryRecordDate : g.recordDate,
      isOvernight: isOvernight,
      sourceRows: g.sourceRows,
      primaryRow: primary,
      issues: g.issues,
      stayPk: g.checkIn ? buildStayPk_(g.childName, g.checkIn) : '',
    };
  });

  stays.sort(function(a, b) {
    var ta = a.recordDate ? a.recordDate.getTime() : 0;
    var tb = b.recordDate ? b.recordDate.getTime() : 0;
    var t = ta - tb;
    return t !== 0 ? t : a.childName.localeCompare(b.childName);
  });
  return stays;
}

/**
 * グループ内のソース行から「プライマリ行」を選ぶ
 * - 記録日 == 入所日（日付一致）の行を優先
 * - 該当無しの場合は記録日が最も早い行を採用
 * @param {Array<Array>} rows ソース行
 * @param {Date|null} checkIn 入所日時
 * @returns {Array} プライマリ行
 */
function pickPrimaryRow_(rows, checkIn) {
  if (rows.length === 1) return rows[0];
  if (checkIn instanceof Date) {
    var inKey = formatDateKey_(new Date(checkIn.getFullYear(), checkIn.getMonth(), checkIn.getDate()));
    for (var i = 0; i < rows.length; i++) {
      var rd = new Date(rows[i][FORM_COL.RECORD_DATE - 1]);
      if (isValidDate_(rd) && formatDateKey_(rd) === inKey) return rows[i];
    }
  }
  // 記録日昇順で先頭
  var sorted = rows.slice().sort(function(a, b) {
    var ra = new Date(a[FORM_COL.RECORD_DATE - 1]);
    var rb = new Date(b[FORM_COL.RECORD_DATE - 1]);
    return (ra.getTime() || 0) - (rb.getTime() || 0);
  });
  return sorted[0];
}

/**
 * 旧API互換: pairStayRecords_ のラッパー
 * 既存の呼び出し側を順次差し替えるための暫定。Phase 3 で削除予定。
 * @deprecated pairStayRecords_ を直接呼び出すこと
 */
function pairOvernightRecords_(formResponses) {
  return pairStayRecords_(formResponses);
}

/**
 * 宿泊PK文字列を生成する
 * 形式: "<児童名>|<入所日時ISO>"（例: "aさん|2026-04-30T21:00:00"）
 * 入所日時はタイムゾーン依存を避けるためスクリプトTZでフォーマット
 * @param {string} childName 児童名
 * @param {Date} checkIn 入所日時
 * @returns {string} 宿泊PK
 */
function buildStayPk_(childName, checkIn) {
  if (!childName || !(checkIn instanceof Date) || isNaN(checkIn.getTime())) return '';
  var iso = Utilities.formatDate(checkIn, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  return String(childName) + '|' + iso;
}

/**
 * フォーム/確定来館記録の行が連泊扱いかを判定する
 * 新仕様: 入所日と退所日が異なる行を連泊とみなす（旧 OVERNIGHT_FLAG 列の値は補助として使用）
 * @param {Array} row フォーム回答 or 確定来館記録の1行
 * @param {boolean} [fromConfirmed=false] true=CONFIRMED_COL基準、false=FORM_COL基準
 * @returns {boolean}
 */
function isOvernightRow_(row, fromConfirmed) {
  var checkInCol = fromConfirmed ? CONFIRMED_COL.CHECK_IN : FORM_COL.CHECK_IN;
  var checkOutCol = fromConfirmed ? CONFIRMED_COL.CHECK_OUT : FORM_COL.CHECK_OUT;
  var checkIn = row[checkInCol - 1];
  var checkOut = row[checkOutCol - 1];
  if ((checkIn instanceof Date) && (checkOut instanceof Date) && checkIn.getFullYear() >= 1900 && checkOut.getFullYear() >= 1900) {
    var inDay = new Date(checkIn.getFullYear(), checkIn.getMonth(), checkIn.getDate());
    var outDay = new Date(checkOut.getFullYear(), checkOut.getMonth(), checkOut.getDate());
    if (inDay.getTime() !== outDay.getTime()) return true;
  }
  // フォールバック: 旧 OVERNIGHT_FLAG 列の値（移行期データ向け）
  var flagCol = fromConfirmed ? CONFIRMED_COL.OVERNIGHT_FLAG : FORM_COL.OVERNIGHT_FLAG;
  var v = row[flagCol - 1];
  if (v === true) return true;
  if (v === false || v === '' || v == null) return false;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'on' || s === '1' || s === '連泊' || s === '○' || s === 'はい';
}

/**
 * 設定シートのB列（値）から指定行の値を取得する
 * 列レイアウト: A=項目名 / B=値 / C=備考（setupSettingsSheet_ と対応）
 * @param {number} row 行番号
 * @returns {*} 設定値（シートが存在しない場合はnull）
 */
function getSettingValue_(row) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
    if (!sheet) return null;
    return sheet.getRange(row, 2).getValue();
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
 * 食事（夕/朝/昼）・服薬（朝/夜）は設定シートの個別行を読む。
 * 設定シート未設定の項目は ALLOCATION_DEFAULTS にフォールバックする。
 *
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
    MEAL_DINNER: pick(SETTINGS_ROW.MEAL_DINNER, ALLOCATION_DEFAULTS.MEAL_DINNER),
    MEAL_BREAKFAST: pick(SETTINGS_ROW.MEAL_BREAKFAST, ALLOCATION_DEFAULTS.MEAL_BREAKFAST),
    MEAL_LUNCH: pick(SETTINGS_ROW.MEAL_LUNCH, ALLOCATION_DEFAULTS.MEAL_LUNCH),
    BATH: pick(SETTINGS_ROW.BATH, ALLOCATION_DEFAULTS.BATH),
    SLEEP: pick(SETTINGS_ROW.SLEEP, ALLOCATION_DEFAULTS.SLEEP),
    BOWEL: pick(SETTINGS_ROW.BOWEL, ALLOCATION_DEFAULTS.BOWEL),
    MEDICINE_NIGHT: pick(SETTINGS_ROW.MEDICINE_NIGHT, ALLOCATION_DEFAULTS.MEDICINE_NIGHT),
    MEDICINE_MORNING: pick(SETTINGS_ROW.MEDICINE_MORNING, ALLOCATION_DEFAULTS.MEDICINE_MORNING),
    NOTES: pick(SETTINGS_ROW.NOTES, ALLOCATION_DEFAULTS.NOTES),
  };
}

/**
 * 設定シートから1日最大来館数を取得する
 * 未設定・不正値の場合は DEFAULT_MAX_VISITS_PER_DAY を返す
 * @returns {number} 1日最大来館数
 */
function getMaxVisitsPerDay_() {
  var v = getSettingValue_(SETTINGS_ROW.MAX_VISITS_PER_DAY);
  var n = parseInt(v, 10);
  return (isNaN(n) || n <= 0) ? DEFAULT_MAX_VISITS_PER_DAY : n;
}

/**
 * 設定シートからメール件名を取得する
 * 未設定の場合は DEFAULT_EMAIL_SUBJECT を返す
 * @returns {string} メール件名
 */
function getEmailSubjectFromSettings_() {
  var v = getSettingValue_(SETTINGS_ROW.EMAIL_SUBJECT);
  var s = (v === null || typeof v === 'undefined') ? '' : String(v).trim();
  return s || DEFAULT_EMAIL_SUBJECT;
}

/**
 * 設定シートからメール本文テンプレートを取得する
 * 未設定の場合は DEFAULT_EMAIL_TEMPLATE を返す
 * @returns {string} メール本文テンプレート
 */
function getEmailBodyFromSettings_() {
  var v = getSettingValue_(SETTINGS_ROW.EMAIL_BODY);
  var s = (v === null || typeof v === 'undefined') ? '' : String(v);
  return s.trim() ? s : DEFAULT_EMAIL_TEMPLATE;
}

/**
 * 設定シートからエラー通知先メール（カンマ区切り）を配列で取得する
 * @returns {Array<string>} メールアドレスの配列（空要素除く）
 */
function getErrorEmailsFromSettings_() {
  var v = getSettingValue_(SETTINGS_ROW.ERROR_EMAIL);
  if (!v) return [];
  return String(v).split(/[,、，\s]+/).map(function(e) { return e.trim(); }).filter(function(e) { return !!e; });
}

/**
 * 定型文マスタシートから連絡事項のレパートリーを取得する
 * @returns {Array<string>} 定型文の配列（ヘッダー除く、空行除く）
 */
function getNotesMasterData_() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NOTES_MASTER);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var notes = [];
    for (var i = 1; i < data.length; i++) {
      var note = String(data[i][0] || '').trim();
      if (note) notes.push(note);
    }
    return notes;
  } catch (e) {
    return [];
  }
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
      var logHeaderRange = sheet.getRange(1, 1, 1, 4);
      logHeaderRange.setValues([['タイムスタンプ', '関数名', 'エラーメッセージ', 'スタックトレース']]);
      styleSheetHeader_(logHeaderRange, 1);
      setColumnWidths_(sheet, { 1: 160, 2: 200, 3: 400, 4: 400 });
    }

    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    var message = error.message || String(error);
    var stack = error.stack || '';

    sheet.appendRow([timestamp, functionName, message, stack]);
  } catch (e) {
    Logger.log('ログ出力に失敗: ' + e.message);
  }
}

/**
 * WebView アクセストークンを検証する
 * @param {string} token クエリパラメータ / クライアントから渡されたトークン
 * @returns {{email: string, name: string}|null} 有効ならユーザー情報、無効なら null
 */
function validateToken_(token) {
  if (!token) return null;
  var normalized = String(token).trim();
  if (!normalized) return null;

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ALLOWED_USERS);
    if (!sheet) return null;
    var lastRow = sheet.getLastRow();
    if (lastRow < ALLOWED_USERS_DATA_START_ROW) return null;

    var numRows = lastRow - ALLOWED_USERS_DATA_START_ROW + 1;
    var values = sheet.getRange(ALLOWED_USERS_DATA_START_ROW, 1, numRows, ALLOWED_USERS_COL.NOTE).getValues();
    for (var i = 0; i < values.length; i++) {
      var rowToken = String(values[i][ALLOWED_USERS_COL.TOKEN - 1] || '').trim();
      var active = values[i][ALLOWED_USERS_COL.ACTIVE - 1];
      if (rowToken && rowToken === normalized && active === true) {
        return {
          email: String(values[i][ALLOWED_USERS_COL.EMAIL - 1] || ''),
          name: String(values[i][ALLOWED_USERS_COL.NAME - 1] || ''),
        };
      }
    }
  } catch (e) {
    logError_('validateToken_', e);
  }
  return null;
}

