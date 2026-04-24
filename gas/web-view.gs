/**
 * GAS Webビュー: フォームの回答 閲覧・編集
 * フォーム誤送信の修正・削除に対応（追加は行わない）
 *
 * アクセス制御: クエリパラメータ `?t=<token>` を「許可ユーザー」シートで検証。
 * 無効・退所済み（active=FALSE）の場合は拒否ページを返す。
 */

function doGet(e) {
  var token = (e && e.parameter && e.parameter.t) || '';
  var user = validateToken_(token);
  if (!user) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8">' +
      '<meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<title>アクセス不可</title></head>' +
      '<body style="font-family:\'Hiragino Sans\',\'Meiryo\',sans-serif;padding:40px;text-align:center;color:#333;">' +
      '<h2 style="color:#ea4335;">アクセス権限がありません</h2>' +
      '<p>このURLはアクセスが許可されていません。</p>' +
      '<p>管理者までお問い合わせください。</p>' +
      '</body></html>'
    ).setTitle('アクセス不可');
  }

  var template = HtmlService.createTemplateFromFile('index');
  template.accessToken = token;
  template.userName = user.name;
  return template.evaluate()
    .setTitle('フォーム回答の修正 | 来館管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * API 用のトークン再検証。不正な場合は例外を投げて google.script.run のエラーハンドラへ到達させる。
 * @param {string} token
 */
function requireValidToken_(token) {
  var user = validateToken_(token);
  if (!user) throw new Error('アクセス権限がありません');
  return user;
}

/**
 * 初期データ（年リスト・児童名リスト）を返す
 */
function getInitialDataWeb(token) {
  requireValidToken_(token);
  var years = {};
  var childNames = {};
  try {
    var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
    var data = sheet.getDataRange().getValues();
    data.slice(FORM_DATA_START_ROW - 1).forEach(function(row) {
      var d = new Date(row[FORM_COL.RECORD_DATE - 1]);
      if (isValidDate_(d)) years[d.getFullYear()] = true;
      var name = String(row[FORM_COL.CHILD_NAME - 1] || '').trim();
      if (name) childNames[name] = true;
    });
  } catch (e) {
    logError_('getInitialDataWeb', e);
  }

  var yearList = Object.keys(years).map(Number).sort(function(a, b) { return b - a; });
  if (!yearList.length) yearList = [new Date().getFullYear()];

  return {
    years: yearList,
    children: Object.keys(childNames).sort(),
  };
}

/**
 * フォームの回答データを取得する（行番号付き）
 */
function getFormResponsesWeb(token, mode, year, month) {
  requireValidToken_(token);
  var sheet;
  try {
    sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
  } catch (e) {
    return [];
  }

  var tz = Session.getScriptTimeZone();
  var allData = sheet.getDataRange().getValues();
  var result = [];

  for (var i = FORM_DATA_START_ROW - 1; i < allData.length; i++) {
    var row = allData[i];
    var dateVal = row[FORM_COL.RECORD_DATE - 1];
    if (!dateVal) continue;
    var d = new Date(dateVal);
    if (!isValidDate_(d)) continue;

    var rowYear = d.getFullYear();
    var rowMonth = d.getMonth() + 1;
    var matches = (mode === 'year')
      ? (rowYear === year)
      : (rowYear === year && rowMonth === month);
    if (!matches) continue;

    var tsVal = row[FORM_COL.TIMESTAMP - 1];
    var checkInVal = row[FORM_COL.CHECK_IN - 1];
    var checkOutVal = row[FORM_COL.CHECK_OUT - 1];

    result.push({
      rowIndex: i + 1,
      timestamp: formatDtDisplay_(tsVal, tz),
      recordDateDisplay: formatDateYMD_(d, 'yyyy/MM/dd', tz),
      recordDateInput: formatDateYMD_(d, 'yyyy-MM-dd', tz),
      childName: String(row[FORM_COL.CHILD_NAME - 1] || ''),
      staffName: String(row[FORM_COL.STAFF_NAME - 1] || ''),
      staffName2: String(row[FORM_COL.STAFF_NAME_2 - 1] || ''),
      checkInDisplay: formatDtDisplay_(checkInVal, tz),
      checkInInput: formatDtInput_(checkInVal, tz),
      checkOutDisplay: formatDtDisplay_(checkOutVal, tz),
      checkOutInput: formatDtInput_(checkOutVal, tz),
      temperature: (row[FORM_COL.TEMPERATURE - 1] !== '' && row[FORM_COL.TEMPERATURE - 1] !== null)
        ? String(row[FORM_COL.TEMPERATURE - 1])
        : '',
      mealDinner: String(row[FORM_COL.MEAL_DINNER - 1] || ''),
      mealBreakfast: String(row[FORM_COL.MEAL_BREAKFAST - 1] || ''),
      mealLunch: String(row[FORM_COL.MEAL_LUNCH - 1] || ''),
      bath: String(row[FORM_COL.BATH - 1] || ''),
      sleep: String(row[FORM_COL.SLEEP - 1] || ''),
      bowel: String(row[FORM_COL.BOWEL - 1] || ''),
      medicineNight: String(row[FORM_COL.MEDICINE_NIGHT - 1] || ''),
      medicineMorning: String(row[FORM_COL.MEDICINE_MORNING - 1] || ''),
      notes: String(row[FORM_COL.NOTES - 1] || ''),
      isOvernight: isOvernightRow_(row),
    });
  }

  return result;
}

function formatDtDisplay_(val, tz) {
  if (!isValidDate_(val)) return '';
  return formatDateYMD_(val, 'yyyy/MM/dd HH:mm', tz);
}

function formatDtInput_(val, tz) {
  if (!isValidDate_(val)) return '';
  return formatDateYMD_(val, "yyyy-MM-dd'T'HH:mm", tz);
}

/**
 * フォームの回答の1行を更新する
 */
function updateFormResponseWeb(token, rowIndex, data) {
  try {
    requireValidToken_(token);
    var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);

    var recordDate = new Date(data.recordDateInput);
    var checkIn = data.checkInInput ? new Date(data.checkInInput) : '';
    var checkOut = data.checkOutInput ? new Date(data.checkOutInput) : '';

    var values = [[
      recordDate,
      data.staffName,
      data.staffName2 || '',
      data.childName,
      checkIn,
      checkOut,
      data.isOvernight ? true : false,
      (data.temperature !== '' && data.temperature !== null) ? parseFloat(data.temperature) : '',
      data.mealDinner || '',
      data.mealBreakfast || '',
      data.mealLunch || '',
      data.bath,
      data.sleep,
      data.bowel,
      data.medicineNight || '',
      data.medicineMorning || '',
      data.notes || '',
    ]];

    // RECORD_DATE(col 2) から NOTES(col 18) までの 17 列を書き込む（新列順: 連泊はCHECK_OUT直後）
    var writeWidth = FORM_COL.NOTES - FORM_COL.RECORD_DATE + 1;
    sheet.getRange(rowIndex, FORM_COL.RECORD_DATE, 1, writeWidth).setValues(values);
    sheet.getRange(rowIndex, FORM_COL.RECORD_DATE, 1, 1).setNumberFormat('yyyy/mm/dd');
    if (checkIn instanceof Date) {
      sheet.getRange(rowIndex, FORM_COL.CHECK_IN, 1, 2).setNumberFormat('yyyy/mm/dd H:mm');
    }

    return { success: true };
  } catch (e) {
    logError_('updateFormResponseWeb', e);
    return { success: false, error: e.message };
  }
}

/**
 * フォームの回答の1行を削除する
 * 行物理削除は書式を崩すため、後続行の値を1行ずつ上にシフトし、
 * 末尾行の値だけクリアする方式で「値詰め」を行う（書式は保持）。
 */
function deleteFormResponseWeb(token, rowIndex) {
  try {
    requireValidToken_(token);
    var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (rowIndex > lastRow || lastCol === 0) return { success: true };

    if (rowIndex < lastRow) {
      // 後続行の値を1行上にシフト
      var shiftRange = sheet.getRange(rowIndex + 1, 1, lastRow - rowIndex, lastCol);
      var shiftValues = shiftRange.getValues();
      sheet.getRange(rowIndex, 1, shiftValues.length, lastCol).setValues(shiftValues);
    }
    // 末尾行の値のみクリア（書式は維持）
    sheet.getRange(lastRow, 1, 1, lastCol).clearContent();
    return { success: true };
  } catch (e) {
    logError_('deleteFormResponseWeb', e);
    return { success: false, error: e.message };
  }
}
