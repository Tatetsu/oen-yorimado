/**
 * GAS Webビュー: フォームの回答 閲覧・編集
 * フォーム誤送信の修正・削除に対応（追加は行わない）
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('フォーム回答の修正 | 来館管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 初期データ（年リスト・児童名リスト）を返す
 */
function getInitialDataWeb() {
  var years = {};
  var childNames = {};
  try {
    var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
    var data = sheet.getDataRange().getValues();
    data.slice(FORM_DATA_START_ROW - 1).forEach(function(row) {
      var d = new Date(row[FORM_COL.RECORD_DATE - 1]);
      if (!isNaN(d.getTime())) years[d.getFullYear()] = true;
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
function getFormResponsesWeb(mode, year, month) {
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
    if (isNaN(d.getTime())) continue;

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
      recordDateDisplay: Utilities.formatDate(d, tz, 'yyyy/MM/dd'),
      recordDateInput: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
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
      meal: String(row[FORM_COL.MEAL - 1] || ''),
      bath: String(row[FORM_COL.BATH - 1] || ''),
      sleep: String(row[FORM_COL.SLEEP - 1] || ''),
      bowel: String(row[FORM_COL.BOWEL - 1] || ''),
      medicine: String(row[FORM_COL.MEDICINE - 1] || ''),
      notes: String(row[FORM_COL.NOTES - 1] || ''),
    });
  }

  return result;
}

function formatDtDisplay_(val, tz) {
  if (!val || !(val instanceof Date) || isNaN(val.getTime())) return '';
  return Utilities.formatDate(val, tz, 'yyyy/MM/dd HH:mm');
}

function formatDtInput_(val, tz) {
  if (!val || !(val instanceof Date) || isNaN(val.getTime())) return '';
  return Utilities.formatDate(val, tz, "yyyy-MM-dd'T'HH:mm");
}

/**
 * フォームの回答の1行を更新する
 */
function updateFormResponseWeb(rowIndex, data) {
  try {
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
      (data.temperature !== '' && data.temperature !== null) ? parseFloat(data.temperature) : '',
      data.meal,
      data.bath,
      data.sleep,
      data.bowel,
      data.medicine,
      data.notes || '',
    ]];

    sheet.getRange(rowIndex, FORM_COL.RECORD_DATE, 1, 13).setValues(values);
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
 */
function deleteFormResponseWeb(rowIndex) {
  try {
    var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    logError_('deleteFormResponseWeb', e);
    return { success: false, error: e.message };
  }
}
