/**
 * F-03: 確定来館記録生成
 * フォームの回答（実データ）を確定来館記録シートに書き込む
 * 振り分けデータは allocation.gs が直接管理するため、ここでは実データのみ扱う
 */

/**
 * 確定来館記録を再生成する
 * 年月指定がある場合はその月の実データのみ洗い替えし、他の月はそのまま保持する
 * 年月指定がない場合は全期間の実データを洗い替えする（従来動作）
 * @param {number} [year] 対象年（省略時は全期間）
 * @param {number} [month] 対象月 1-12（省略時は全期間）
 */
function updateConfirmedVisits(year, month) {
  var sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  var filterByMonth = (year !== undefined && month !== undefined);
  var colCount = CONFIRMED_COL.OVERNIGHT_FLAG; // 列数=末尾(OVERNIGHT_FLAG=18)

  // 既存データを全件取得
  var lastRow = sheet.getLastRow();
  var existingData = [];
  if (lastRow >= CONFIRMED_DATA_START_ROW) {
    existingData = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, colCount).getValues();
  }

  // 既存データを「残す行」と「置き換える行」に分離
  // 不正な日付（pre-1900）の行は実データ・振り分け問わず除去する
  var keepRows = [];
  if (filterByMonth) {
    // 月指定あり: 対象月の実データ行だけを除去し、それ以外は保持
    keepRows = existingData.filter(function(row) {
      var recordDate = new Date(row[CONFIRMED_COL.RECORD_DATE - 1]);
      if (isNaN(recordDate.getTime()) || recordDate.getFullYear() < 1900) return false;
      if (row[CONFIRMED_COL.DATA_TYPE - 1] === '振り分け') return true;
      return !(recordDate.getFullYear() === year && (recordDate.getMonth() + 1) === month);
    });
  } else {
    // 月指定なし（従来動作）: 振り分け行のみ保持（不正日付の振り分けは除去）
    keepRows = existingData.filter(function(row) {
      if (row[CONFIRMED_COL.DATA_TYPE - 1] !== '振り分け') return false;
      var recordDate = new Date(row[CONFIRMED_COL.RECORD_DATE - 1]);
      return !isNaN(recordDate.getTime()) && recordDate.getFullYear() >= 1900;
    });
  }

  // 全期間の回答からペアリング → 月指定で絞り込み
  var allResponses = getFormResponsesAll_();
  var allStays = pairOvernightRecords_(allResponses);

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
      // 月指定がある場合、対象月外の日付は除外
      if (filterByMonth && (stayDate.getFullYear() !== year || (stayDate.getMonth() + 1) !== month)) {
        return;
      }
      newRows.push([
        stayDate,                                  // 記録日（宿泊日ごとに展開）
        primary[FORM_COL.CHILD_NAME - 1],         // 児童名
        '実データ',                                 // データ区分
        primary[FORM_COL.STAFF_NAME - 1],         // スタッフ1
        primary[FORM_COL.STAFF_NAME_2 - 1],       // スタッフ2（任意・空欄の場合あり）
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
    var dateCompare = new Date(a[0]) - new Date(b[0]);
    if (dateCompare !== 0) return dateCompare;
    return String(a[1]).localeCompare(String(b[1]));
  });

  // 書き込み
  sheet.getRange(CONFIRMED_DATA_START_ROW, 1, allData.length, colCount).setValues(allData);

  // 記録日列の表示形式
  sheet.getRange(CONFIRMED_DATA_START_ROW, 1, allData.length, 1)
    .setNumberFormat('yyyy/mm/dd');

  // 入所日時・退所日時列の表示形式
  sheet.getRange(CONFIRMED_DATA_START_ROW, CONFIRMED_COL.CHECK_IN, allData.length, 2)
    .setNumberFormat('yyyy/mm/dd H:mm');

  var scopeMsg = filterByMonth ? (year + '年' + month + '月分') : '全期間';
  Logger.log('確定来館記録を更新しました（' + scopeMsg + '）: ' + allData.length + '件');
}
