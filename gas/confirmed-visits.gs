/**
 * F-03: 確定来館記録生成
 * フォームの回答（実データ）を確定来館記録シートに書き込む
 * 振り分けデータは allocation.gs が直接管理するため、ここでは実データのみ扱う
 */

/**
 * 確定来館記録を再生成する
 * - (year, month) 指定: 対象月の実データのみ洗い替え（他月・振り分け行は保持）
 * - (year) 指定:        対象年の実データのみ洗い替え（他年・振り分け行は保持）
 * 振り分け行は allocation.gs が直接管理するため、ここでは触らない
 * @param {number} year 対象年
 * @param {number} [month] 対象月 1-12（省略時は年単位）
 */
function updateConfirmedVisits(year, month) {
  if (year === undefined || year === null) {
    throw new Error('updateConfirmedVisits: year は必須です');
  }
  var sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  var filterByMonth = (month !== undefined && month !== null);
  var colCount = CONFIRMED_COL.STAY_PK; // 列数=末尾(STAY_PK=19)

  // 既存データを全件取得
  var lastRow = sheet.getLastRow();
  var existingData = [];
  if (lastRow >= CONFIRMED_DATA_START_ROW) {
    existingData = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, colCount).getValues();
  }

  // 既存データを「残す行」と「置き換える行」に分離
  // 不正な日付（pre-1900）の行は実データ・振り分け問わず除去する
  var keepRows = existingData.filter(function(row) {
    var recordDate = new Date(row[CONFIRMED_COL.RECORD_DATE - 1]);
    if (isNaN(recordDate.getTime()) || recordDate.getFullYear() < 1900) return false;
    // 振り分け行は常に保持
    if (row[CONFIRMED_COL.DATA_TYPE - 1] === '振り分け') return true;
    // 実データ行はスコープに該当するもののみ除去
    if (filterByMonth) {
      return !(recordDate.getFullYear() === year && (recordDate.getMonth() + 1) === month);
    }
    return recordDate.getFullYear() !== year;
  });

  // 全期間の回答からペアリング → 月指定で絞り込み
  var allResponses = getFormResponsesAll_();
  var allStays = pairStayRecords_(allResponses);

  var newRows = [];
  allStays.forEach(function(stay) {
    // ペアリング後の論理1宿泊から、滞在カレンダー全日に展開
    var checkIn = stay.checkIn;
    var checkOut = stay.checkOut;
    var recordDate = stay.recordDate;
    var primary = stay.primaryRow;

    // フル日時が両方そろっている場合はそのまま、片方欠けている場合はフォールバック
    var checkInFull = checkIn;
    var checkOutFull = checkOut;
    if (!(checkInFull instanceof Date) || !(checkOutFull instanceof Date)) {
      // 時刻のみ・欠損のフォールバック（後方互換）
      var baseDate = (recordDate instanceof Date) ? recordDate : new Date(recordDate);
      if (!(checkInFull instanceof Date)) {
        checkInFull = checkIn ? toDateTimeOnDate_(baseDate, checkIn) : null;
      }
      if (!(checkOutFull instanceof Date)) {
        checkOutFull = checkOut ? toDateTimeOnDate_(baseDate, checkOut) : null;
      }
      if (checkInFull && checkOutFull && checkOutFull.getTime() <= checkInFull.getTime()) {
        checkOutFull = new Date(checkOutFull.getTime() + 24 * 60 * 60 * 1000);
      }
    }

    var stayDates = expandStayToDates_(recordDate, checkInFull, checkOutFull);
    stayDates.forEach(function(stayDate) {
      // スコープ外の日付は除外
      if (filterByMonth) {
        if (stayDate.getFullYear() !== year || (stayDate.getMonth() + 1) !== month) return;
      } else {
        if (stayDate.getFullYear() !== year) return;
      }
      newRows.push([
        '実データ',                                 // データ区分
        stayDate,                                  // 記録日（宿泊日ごとに展開）
        primary[FORM_COL.STAFF_NAME - 1],         // スタッフ1
        primary[FORM_COL.STAFF_NAME_2 - 1],       // スタッフ2（任意・空欄の場合あり）
        primary[FORM_COL.CHILD_NAME - 1],         // 児童名
        checkInFull,                               // 入所日時（ペアリング後の値）
        checkOutFull,                              // 退所予定日時（ペアリング後の値）
        primary[FORM_COL.TEMPERATURE - 1],        // 体温
        primary[FORM_COL.MEAL_DINNER - 1],        // 夕食
        primary[FORM_COL.MEAL_BREAKFAST - 1],     // 朝食
        primary[FORM_COL.MEAL_LUNCH - 1],         // 昼食
        primary[FORM_COL.BATH - 1],               // 入浴
        primary[FORM_COL.SLEEP - 1],              // 睡眠
        primary[FORM_COL.BOWEL - 1],              // 便
        primary[FORM_COL.MEDICINE_NIGHT - 1],     // 服薬(夜)
        primary[FORM_COL.MEDICINE_MORNING - 1],   // 服薬(朝)
        primary[FORM_COL.NOTES - 1],              // その他連絡事項
        stay.isOvernight ? true : false,          // 連泊フラグ
        buildStayPk_(stay.childName, checkInFull),// 宿泊PK
      ]);
    });
  });

  var allData = keepRows.concat(newRows);

  // 既存データをクリア（ヘッダーは残す）
  if (lastRow >= CONFIRMED_DATA_START_ROW) {
    sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, colCount).clearContent();
  }

  if (allData.length === 0) {
    Logger.log('確定来館記録: 書き込むデータがありません');
    return;
  }

  // 日付昇順 → 児童名昇順でソート
  allData.sort(function(a, b) {
    var dateCompare = new Date(a[CONFIRMED_COL.RECORD_DATE - 1]) - new Date(b[CONFIRMED_COL.RECORD_DATE - 1]);
    if (dateCompare !== 0) return dateCompare;
    return String(a[CONFIRMED_COL.CHILD_NAME - 1]).localeCompare(String(b[CONFIRMED_COL.CHILD_NAME - 1]));
  });

  // 書き込み
  sheet.getRange(CONFIRMED_DATA_START_ROW, 1, allData.length, colCount).setValues(allData);

  // 記録日列の表示形式
  sheet.getRange(CONFIRMED_DATA_START_ROW, CONFIRMED_COL.RECORD_DATE, allData.length, 1)
    .setNumberFormat('yyyy/mm/dd');

  // 入所日時・退所日時列の表示形式
  sheet.getRange(CONFIRMED_DATA_START_ROW, CONFIRMED_COL.CHECK_IN, allData.length, 2)
    .setNumberFormat('yyyy/mm/dd H:mm');

  var scopeMsg = filterByMonth ? (year + '年' + month + '月分') : (year + '年分');
  Logger.log('確定来館記録を更新しました（' + scopeMsg + '）: ' + allData.length + '件');
}
