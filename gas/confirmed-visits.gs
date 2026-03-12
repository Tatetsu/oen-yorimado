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

  // 既存データを全件取得
  var lastRow = sheet.getLastRow();
  var existingData = [];
  if (lastRow >= CONFIRMED_DATA_START_ROW) {
    existingData = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, 13).getValues();
  }

  // 既存データを「残す行」と「置き換える行」に分離
  var keepRows = [];
  if (filterByMonth) {
    // 月指定あり: 対象月の実データ行だけを除去し、それ以外は保持
    keepRows = existingData.filter(function(row) {
      if (row[CONFIRMED_COL.DATA_TYPE - 1] === '振り分け') return true; // 振り分け行は常に保持
      var recordDate = new Date(row[CONFIRMED_COL.RECORD_DATE - 1]);
      // 対象月以外の実データ行は保持
      return !(recordDate.getFullYear() === year && (recordDate.getMonth() + 1) === month);
    });
  } else {
    // 月指定なし（従来動作）: 振り分け行のみ保持
    keepRows = existingData.filter(function(row) {
      return row[CONFIRMED_COL.DATA_TYPE - 1] === '振り分け';
    });
  }

  // フォームの回答から実データを取得（月指定あり→対象月のみ、なし→全件）
  var formData = filterByMonth ? getFormResponsesByMonth(year, month) : getFormResponsesAll_();
  var newRows = [];
  formData.forEach(function(row) {
    newRows.push([
      row[FORM_COL.RECORD_DATE - 1],     // 記録日
      row[FORM_COL.CHILD_NAME - 1],      // 児童名
      '実データ',                          // データ区分
      row[FORM_COL.STAFF_NAME - 1],      // スタッフ名
      row[FORM_COL.CHECK_IN - 1],        // 入所時間
      row[FORM_COL.CHECK_OUT - 1],       // 退所時間
      row[FORM_COL.TEMPERATURE - 1],     // 体温
      row[FORM_COL.MEAL - 1],            // 食事
      row[FORM_COL.BATH - 1],            // 入浴
      row[FORM_COL.SLEEP - 1],           // 睡眠
      row[FORM_COL.BOWEL - 1],           // 便
      row[FORM_COL.MEDICINE - 1],        // 服薬
      row[FORM_COL.NOTES - 1],           // その他連絡事項
    ]);
  });

  var allData = keepRows.concat(newRows);

  // 既存データをクリア（ヘッダーは残す）
  if (lastRow >= CONFIRMED_DATA_START_ROW) {
    sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, 13).clearContent();
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
  sheet.getRange(CONFIRMED_DATA_START_ROW, 1, allData.length, 13).setValues(allData);

  // 記録日列の表示形式
  sheet.getRange(CONFIRMED_DATA_START_ROW, 1, allData.length, 1)
    .setNumberFormat('yyyy/mm/dd');

  // 入所時間・退所時間列の表示形式
  sheet.getRange(CONFIRMED_DATA_START_ROW, CONFIRMED_COL.CHECK_IN, allData.length, 2)
    .setNumberFormat('H:mm');

  var scopeMsg = filterByMonth ? (year + '年' + month + '月分') : '全期間';
  Logger.log('確定来館記録を更新しました（' + scopeMsg + '）: ' + allData.length + '件');
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
