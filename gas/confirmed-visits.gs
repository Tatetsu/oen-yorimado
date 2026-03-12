/**
 * F-03: 確定来館記録生成
 * フォームの回答（実データ）+ 振り分け記録を統合して確定来館記録シートに書き込む
 */

/**
 * 確定来館記録を全月分再生成する
 * フォームの回答と振り分け記録を統合し、確定来館記録シートに書き込む
 */
function updateConfirmedVisits() {
  var sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);

  // 既存データをクリア（ヘッダーは残す）
  var lastRow = sheet.getLastRow();
  if (lastRow >= CONFIRMED_DATA_START_ROW) {
    sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, 13).clearContent();
  }

  var allData = [];

  // フォームの回答から実データを取得
  var formData = getFormResponsesAll_();
  formData.forEach(function(row) {
    allData.push([
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

  // 振り分け記録からデータを取得（補完データ含む）
  var allocationData = getAllocationsAll_();
  allocationData.forEach(function(row) {
    // 旧形式（4列）データはスキップ（再振り分けが必要）
    if (row.length < ALLOCATION_COL_COUNT) {
      Logger.log('警告: 振り分け記録に旧形式データがあります（列数: ' + row.length + '）。再振り分けしてください。');
      allData.push([
        row[2] || '',  // 振り分け日（旧形式の3列目）
        row[1] || '',  // 児童名（旧形式の2列目）
        '振り分け',
        '', '', '', '', '', '', '', '', '', '',
      ]);
      return;
    }
    allData.push([
      row[ALLOCATION_COL.ALLOCATION_DATE - 1],  // 記録日（振り分け日）
      row[ALLOCATION_COL.CHILD_NAME - 1],       // 児童名
      '振り分け',                                 // データ区分
      row[ALLOCATION_COL.STAFF_NAME - 1] || '',  // スタッフ名
      row[ALLOCATION_COL.CHECK_IN - 1] || '',    // 入所時間
      row[ALLOCATION_COL.CHECK_OUT - 1] || '',   // 退所時間
      row[ALLOCATION_COL.TEMPERATURE - 1] || '', // 体温
      row[ALLOCATION_COL.MEAL - 1] || '',        // 食事
      row[ALLOCATION_COL.BATH - 1] || '',        // 入浴
      row[ALLOCATION_COL.SLEEP - 1] || '',       // 睡眠
      row[ALLOCATION_COL.BOWEL - 1] || '',       // 便
      row[ALLOCATION_COL.MEDICINE - 1] || '',    // 服薬
      row[ALLOCATION_COL.NOTES - 1] || '',       // その他連絡事項
    ]);
  });

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

  Logger.log('確定来館記録を更新しました: ' + allData.length + '件');
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
 * 振り分け記録から全データを取得する（ヘッダー除く）
 * @returns {Array<Array>} 全振り分けデータ
 */
function getAllocationsAll_() {
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
  return data.slice(1);
}
