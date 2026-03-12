/**
 * F-07: 保護者向け来館報告メール送信
 *
 * TODO: 同日に兄弟（同一保護者メールアドレス）が来館した場合、
 *       個別メールではなく1通にまとめるかどうか検討・実装する
 */

/**
 * 前日の来館記録を保護者にメール送信する（自動トリガー用）
 */
function sendDailyVisitReports() {
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  sendVisitReportsByDate_(yesterday);
}

/**
 * 手動実行用: ダイアログで対象日付を指定してメール送信する
 */
function sendVisitReportsManual() {
  var ui = SpreadsheetApp.getUi();

  // デフォルト値として前日の日付を設定
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var defaultDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  var response = ui.prompt(
    '来館報告メール送信',
    '対象日付を入力してください（例: ' + defaultDate + '）',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var inputDate = response.getResponseText().trim();
  if (!inputDate) {
    inputDate = defaultDate;
  }

  // 日付パース（yyyy/MM/dd または yyyy-MM-dd）
  var targetDate = parseDateInput_(inputDate);
  if (!targetDate) {
    ui.alert('日付の形式が不正です。yyyy/MM/dd の形式で入力してください。');
    return;
  }

  var formattedDate = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  var confirm = ui.alert(
    '確認',
    formattedDate + ' の来館記録をメール送信しますか？',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    return;
  }

  sendVisitReportsByDate_(targetDate);
  ui.alert('メール送信処理が完了しました。詳細はログを確認してください。');
}

/**
 * 指定日の来館記録を保護者にメール送信する
 * @param {Date} targetDate 対象日付
 */
function sendVisitReportsByDate_(targetDate) {
  var tz = Session.getScriptTimeZone();
  var targetDateStr = Utilities.formatDate(targetDate, tz, 'yyyy-MM-dd');
  Logger.log('来館報告メール送信開始: 対象日 = ' + targetDateStr);

  // フォームの回答から対象日のレコードを抽出
  var records = getFormResponsesByDate_(targetDate);
  if (records.length === 0) {
    Logger.log('対象日の来館記録がありません: ' + targetDateStr);
    return;
  }

  // 児童マスタを取得してマップ化（児童名 → 行データ）
  var childMasterMap = buildChildMasterMap_();

  // スクリプトプロパティから施設名を取得
  var props = PropertiesService.getScriptProperties();
  var facilityName = props.getProperty('FACILITY_NAME') || '施設';
  var senderName = props.getProperty('EMAIL_SENDER_NAME') || facilityName;

  var sentCount = 0;
  var skipCount = 0;
  var errorCount = 0;

  for (var i = 0; i < records.length; i++) {
    var record = records[i];
    var childName = record[FORM_COL.CHILD_NAME - 1];
    var masterRow = childMasterMap[childName];

    if (!masterRow) {
      Logger.log('児童マスタに該当なし（スキップ）: ' + childName);
      skipCount++;
      continue;
    }

    var parentEmail = masterRow[MASTER_COL.PARENT_EMAIL - 1];
    if (!parentEmail || String(parentEmail).trim() === '') {
      Logger.log('保護者メールアドレスが未設定（スキップ）: ' + childName);
      skipCount++;
      continue;
    }

    try {
      var emailData = buildEmailData_(record, masterRow, targetDate, facilityName);
      MailApp.sendEmail({
        to: String(parentEmail).trim(),
        subject: emailData.subject,
        body: emailData.body,
        name: senderName,
      });
      Logger.log('メール送信成功: ' + childName + ' → ' + parentEmail);
      sentCount++;
    } catch (error) {
      Logger.log('メール送信エラー: ' + childName + ' → ' + parentEmail + ' / ' + error.message);
      errorCount++;
    }
  }

  Logger.log('来館報告メール送信完了: 送信=' + sentCount + '件, スキップ=' + skipCount + '件, エラー=' + errorCount + '件');
}

/**
 * フォームの回答から指定日のデータを取得する
 * @param {Date} targetDate 対象日付
 * @returns {Array<Array>} 該当日のフォーム回答データ
 */
function getFormResponsesByDate_(targetDate) {
  var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
  var data = sheet.getDataRange().getValues();
  var responses = data.slice(1);
  var tz = Session.getScriptTimeZone();
  var targetStr = Utilities.formatDate(targetDate, tz, 'yyyy-MM-dd');

  return responses.filter(function(row) {
    var recordDate = row[FORM_COL.RECORD_DATE - 1];
    if (!(recordDate instanceof Date)) {
      return false;
    }
    var rowStr = Utilities.formatDate(recordDate, tz, 'yyyy-MM-dd');
    return rowStr === targetStr;
  });
}

/**
 * 児童マスタデータを児童名→行データのマップに変換する
 * @returns {Object} 児童名をキーとするマップ
 */
function buildChildMasterMap_() {
  var masterData = getChildMasterData();
  var map = {};
  for (var i = 0; i < masterData.length; i++) {
    var name = masterData[i][MASTER_COL.NAME - 1];
    map[name] = masterData[i];
  }
  return map;
}

/**
 * メールの件名と本文を組み立てる
 * @param {Array} record フォーム回答の1行
 * @param {Array} masterRow 児童マスタの1行
 * @param {Date} targetDate 対象日付
 * @param {string} facilityName 施設名
 * @returns {{subject: string, body: string}}
 */
function buildEmailData_(record, masterRow, targetDate, facilityName) {
  var tz = Session.getScriptTimeZone();
  var dateStr = Utilities.formatDate(targetDate, tz, 'M月d日');
  var childName = record[FORM_COL.CHILD_NAME - 1];
  var parentName = masterRow[MASTER_COL.PARENT_NAME - 1] || '';
  var staffName = record[FORM_COL.STAFF_NAME - 1] || '';

  var subject = '【' + facilityName + '】' + dateStr + ' ' + childName + 'さんの来館記録';

  var body = EMAIL_TEMPLATE
    .replace(/{保護者名}/g, parentName)
    .replace(/{施設名}/g, facilityName)
    .replace(/{日付}/g, dateStr)
    .replace(/{児童名}/g, childName)
    .replace(/{入所時間}/g, formatTime_(record[FORM_COL.CHECK_IN - 1]))
    .replace(/{退所時間}/g, formatTime_(record[FORM_COL.CHECK_OUT - 1]))
    .replace(/{体温}/g, record[FORM_COL.TEMPERATURE - 1] || '')
    .replace(/{食事}/g, record[FORM_COL.MEAL - 1] || '')
    .replace(/{入浴}/g, record[FORM_COL.BATH - 1] || '')
    .replace(/{睡眠}/g, record[FORM_COL.SLEEP - 1] || '')
    .replace(/{便}/g, record[FORM_COL.BOWEL - 1] || '')
    .replace(/{服薬}/g, record[FORM_COL.MEDICINE - 1] || '')
    .replace(/{連絡事項}/g, record[FORM_COL.NOTES - 1] || '特になし')
    .replace(/{スタッフ名}/g, staffName);

  return { subject: subject, body: body };
}

/**
 * 時刻値をHH:mm形式の文字列にフォーマットする
 * @param {*} value 時刻値（Date型またはテキスト）
 * @returns {string} フォーマット済み時刻文字列
 */
function formatTime_(value) {
  if (!value) {
    return '';
  }
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  return String(value);
}

/**
 * 日付入力文字列をパースする（yyyy/MM/dd または yyyy-MM-dd）
 * @param {string} input 日付文字列
 * @returns {Date|null} パース結果（不正な場合はnull）
 */
function parseDateInput_(input) {
  var match = String(input).match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!match) {
    return null;
  }
  var date = new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10));
  if (isNaN(date.getTime())) {
    return null;
  }
  return date;
}

/**
 * メール送信の時間トリガーを設定する（毎朝8時）
 */
function setupEmailTrigger() {
  // 既存のsendDailyVisitReportsトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendDailyVisitReports') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎朝8時のトリガーを作成
  ScriptApp.newTrigger('sendDailyVisitReports')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('メール送信トリガーを設定しました（毎朝8時）');
  SpreadsheetApp.getUi().alert('メール送信トリガーを設定しました（毎朝8時）');
}
