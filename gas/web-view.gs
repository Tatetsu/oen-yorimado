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
 * 初期データ（年リスト・児童名リスト・スタッフ名リスト・フォーム選択肢）を返す
 * - children: FORM_RESPONSE シートから既出児童名を抽出
 * - staffNames: スタッフマスタ（B列）から取得（空なら FORM_RESPONSE の既出値で補完）
 * - formChoices: 連携フォームのプルダウン正規値（取得失敗時は空オブジェクト、フロントでハードコード値にフォールバック）
 */
function getInitialDataWeb(token) {
  requireValidToken_(token);
  var years = {};
  var childNames = {};
  var staffFromResponses = {};
  try {
    var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
    var data = sheet.getDataRange().getValues();
    data.slice(FORM_DATA_START_ROW - 1).forEach(function(row) {
      var d = getRowRecordDate_(row);
      if (d) years[d.getFullYear()] = true;
      var name = String(row[FORM_COL.CHILD_NAME - 1] || '').trim();
      if (name) childNames[name] = true;
      var s1 = String(row[FORM_COL.STAFF_NAME - 1] || '').trim();
      if (s1) staffFromResponses[s1] = true;
      var s2 = String(row[FORM_COL.STAFF_NAME_2 - 1] || '').trim();
      if (s2) staffFromResponses[s2] = true;
    });
  } catch (e) {
    logError_('getInitialDataWeb', e);
  }

  var yearList = Object.keys(years).map(Number).sort(function(a, b) { return b - a; });
  if (!yearList.length) yearList = [new Date().getFullYear()];

  var staffNames = [];
  try {
    staffNames = getStaffNamesFromMaster_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (e) {
    logError_('getInitialDataWeb.staff', e);
  }
  if (!staffNames.length) {
    staffNames = Object.keys(staffFromResponses).sort();
  }

  return {
    years: yearList,
    children: Object.keys(childNames).sort(),
    staffNames: staffNames,
    formChoices: getFormChoicesFromLinkedForm_(),
  };
}

/**
 * 連携Googleフォームから各プルダウン項目の選択肢を取得する。
 * 取得失敗時や未リンク時は空オブジェクトを返す（フロント側がハードコード値にフォールバック）。
 */
function getFormChoicesFromLinkedForm_() {
  var result = {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formUrl = ss.getFormUrl();
    if (!formUrl) return result;
    var form = FormApp.openByUrl(formUrl);
    // フォーム質問タイトル → フロント側キー（半角/全角カッコ両対応）
    var titleMap = {
      '夕食': 'mealDinner',
      '朝食': 'mealBreakfast',
      '昼食': 'mealLunch',
      '入浴': 'bath',
      '入眠時刻': 'sleepOnset',
      '朝4時チェック': 'sleepCheck4am',
      '起床時刻': 'wakeUp',
      '便': 'bowel',
      '服薬(夜)': 'medicineNight',
      '服薬（夜）': 'medicineNight',
      '服薬(朝)': 'medicineMorning',
      '服薬（朝）': 'medicineMorning',
    };
    form.getItems().forEach(function(item) {
      var title = item.getTitle();
      var key = titleMap[title];
      if (!key) return;
      var type = item.getType();
      var choices = [];
      if (type === FormApp.ItemType.LIST) {
        choices = item.asListItem().getChoices().map(function(c) { return c.getValue(); });
      } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
        choices = item.asMultipleChoiceItem().getChoices().map(function(c) { return c.getValue(); });
      }
      if (choices.length) result[key] = choices;
    });
  } catch (e) {
    logError_('getFormChoicesFromLinkedForm_', e);
  }
  return result;
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
    var d = getRowRecordDate_(row);
    if (!d) continue;

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
      sleepOnset: String(row[FORM_COL.SLEEP_ONSET - 1] || ''),
      sleepCheck4am: String(row[FORM_COL.SLEEP_CHECK_4AM - 1] || ''),
      wakeUp: String(row[FORM_COL.WAKE_UP - 1] || ''),
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
 * 「記録日」列(2列目) はフォームから廃止済みのため書き換え対象外。
 * STAFF_NAME(3列目)〜NOTES(19列目)の17列のみ更新する。
 */
function updateFormResponseWeb(token, rowIndex, data) {
  try {
    requireValidToken_(token);
    var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);

    var checkIn = data.checkInInput ? new Date(data.checkInInput) : '';
    var checkOut = data.checkOutInput ? new Date(data.checkOutInput) : '';

    var values = [[
      data.staffName,
      data.staffName2 || '',
      data.childName,
      checkIn,
      checkOut,
      (data.temperature !== '' && data.temperature !== null) ? parseFloat(data.temperature) : '',
      data.mealDinner || '',
      data.mealBreakfast || '',
      data.mealLunch || '',
      data.bath,
      data.sleepOnset || '',
      data.sleepCheck4am || '',
      data.wakeUp || '',
      data.bowel,
      data.medicineNight || '',
      data.medicineMorning || '',
      data.notes || '',
    ]];

    // STAFF_NAME(col 3) から NOTES(col 19) までの 17 列を書き込む
    var writeWidth = FORM_COL.NOTES - FORM_COL.STAFF_NAME + 1;
    sheet.getRange(rowIndex, FORM_COL.STAFF_NAME, 1, writeWidth).setValues(values);
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
