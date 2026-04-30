/**
 * F-03: 確定来館記録生成
 * フォーム回答（実データ）を確定来館記録シートに書き込む
 * 振り分けデータは allocation.gs が直接管理するため、ここでは実データのみ扱う
 *
 * ロジック: フォーム1行 = 1宿泊（入所〜退所が完結）。各行の入退所日付の差から
 * 滞在期間を展開し、対象月内の各日付を1行として書き出す。
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
  var colCount = CONFIRMED_COL.NOTES; // 列数=末尾(NOTES=21)

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

  // フォーム回答を1行ずつ展開
  var allResponses = getFormResponsesAll_();

  var newRows = [];
  allResponses.forEach(function(row) {
    var childName = row[FORM_COL.CHILD_NAME - 1];
    if (!childName) return;

    var checkIn = row[FORM_COL.CHECK_IN - 1];
    var checkOut = row[FORM_COL.CHECK_OUT - 1];
    var hasIn = (checkIn instanceof Date) && checkIn.getFullYear() >= 1900;
    var hasOut = (checkOut instanceof Date) && checkOut.getFullYear() >= 1900;
    var recordDate = getRowRecordDate_(row);
    if (!hasIn && !hasOut && !recordDate) return; // 全て無効ならスキップ

    var checkInFull = hasIn ? checkIn : null;
    var checkOutFull = hasOut ? checkOut : null;
    var stayDates = expandStayToDates_(recordDate, checkInFull, checkOutFull);
    var checkInKey = checkInFull ? formatDateKey_(checkInFull) : null;
    var checkOutKey = checkOutFull ? formatDateKey_(checkOutFull) : null;

    stayDates.forEach(function(stayDate) {
      // スコープ外の日付は除外
      if (filterByMonth) {
        if (stayDate.getFullYear() !== year || (stayDate.getMonth() + 1) !== month) return;
      } else {
        if (stayDate.getFullYear() !== year) return;
      }
      var stayDateKey = formatDateKey_(stayDate);
      var isCheckInDay = (checkInKey && stayDateKey === checkInKey);
      var isCheckOutDay = (checkOutKey && stayDateKey === checkOutKey);
      // 入所日のみ → 往=1 / 退所日のみ → 復=1 / 中日（どちらでもない）→ 両方1
      // 同日入退所（単日）→ 両条件マッチで両方1
      var pickupOutbound, pickupReturn;
      if (!isCheckInDay && !isCheckOutDay) {
        pickupOutbound = 1;
        pickupReturn = 1;
      } else {
        pickupOutbound = isCheckInDay ? 1 : '';
        pickupReturn = isCheckOutDay ? 1 : '';
      }
      newRows.push([
        '実データ',                              // データ区分
        stayDate,                               // 利用日（滞在日ごとに展開）
        row[FORM_COL.STAFF_NAME - 1],          // スタッフ1
        row[FORM_COL.STAFF_NAME_2 - 1],        // スタッフ2（任意）
        childName,                              // 児童名
        checkInFull,                            // 入所日時
        checkOutFull,                           // 退所予定日時
        pickupOutbound,                         // 送迎(往)
        pickupReturn,                           // 送迎(復)
        row[FORM_COL.TEMPERATURE - 1],         // 体温
        row[FORM_COL.MEAL_DINNER - 1],         // 夕食
        row[FORM_COL.MEAL_BREAKFAST - 1],      // 朝食
        row[FORM_COL.MEAL_LUNCH - 1],          // 昼食
        row[FORM_COL.BATH - 1],                // 入浴
        row[FORM_COL.SLEEP_ONSET - 1],         // 入眠時刻
        row[FORM_COL.SLEEP_CHECK_4AM - 1],     // 朝4時チェック
        row[FORM_COL.WAKE_UP - 1],             // 起床時刻
        row[FORM_COL.BOWEL - 1],               // 便
        row[FORM_COL.MEDICINE_NIGHT - 1],      // 服薬(夜)
        row[FORM_COL.MEDICINE_MORNING - 1],    // 服薬(朝)
        row[FORM_COL.NOTES - 1],               // その他連絡事項
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
    return String(a[CONFIRMED_COL.CHILD_NAME - 1] || '').localeCompare(String(b[CONFIRMED_COL.CHILD_NAME - 1] || ''));
  });

  // 書き込み
  sheet.getRange(CONFIRMED_DATA_START_ROW, 1, allData.length, colCount).setValues(allData);

  // 利用日列の表示形式
  sheet.getRange(CONFIRMED_DATA_START_ROW, CONFIRMED_COL.RECORD_DATE, allData.length, 1)
    .setNumberFormat('yyyy/mm/dd');

  // 入所日時・退所日時列の表示形式
  sheet.getRange(CONFIRMED_DATA_START_ROW, CONFIRMED_COL.CHECK_IN, allData.length, 2)
    .setNumberFormat('yyyy/mm/dd H:mm');

  var scopeMsg = filterByMonth ? (year + '年' + month + '月分') : (year + '年分');
  Logger.log('確定来館記録を更新しました（' + scopeMsg + '）: ' + allData.length + '件');
}
