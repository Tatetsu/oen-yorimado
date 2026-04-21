/**
 * フォームのドロップダウンをマスタデータに同期する
 */

// フォームの質問タイトル（実際の質問文と異なる場合は修正すること）
var FORM_TITLE_STAFF1 = 'スタッフ1';
var FORM_TITLE_STAFF2 = 'スタッフ2';
var FORM_TITLE_CHILD = '児童名';

/**
 * フォームのドロップダウンをマスタデータで更新する
 * 手動実行・トリガーから呼び出す
 */
function syncFormDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formUrl = ss.getFormUrl();
  if (!formUrl) {
    Logger.log('リンクされたフォームが見つかりません');
    return;
  }

  var form = FormApp.openByUrl(formUrl);
  var staffNames = getStaffNamesFromMaster_(ss);
  var childNames = getChildNamesFromMaster_(ss);

  if (!staffNames.length || !childNames.length) {
    Logger.log('マスタデータが空のため同期をスキップしました');
    return;
  }

  form.getItems().forEach(function(item) {
    var title = item.getTitle();
    if (title === FORM_TITLE_STAFF1 || title === FORM_TITLE_STAFF2) {
      setListChoices_(item, staffNames);
    } else if (title === FORM_TITLE_CHILD) {
      setListChoices_(item, childNames);
    }
  });

  Logger.log('同期完了: スタッフ ' + staffNames.length + '名 / 児童 ' + childNames.length + '名');
}

/**
 * スタッフマスタからスタッフ名一覧を取得する（B列）
 */
function getStaffNamesFromMaster_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.STAFF_MASTER);
  if (!sheet) {
    Logger.log('シートが見つかりません: ' + SHEET_NAMES.STAFF_MASTER);
    return [];
  }
  var data = sheet.getDataRange().getValues();
  var names = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][1] || '').trim();
    if (name) names.push(name);
  }
  return names;
}

/**
 * 児童マスタから児童名一覧を取得する（稼働・休止のみ）
 */
function getChildNamesFromMaster_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.CHILD_MASTER);
  if (!sheet) {
    Logger.log('シートが見つかりません: ' + SHEET_NAMES.CHILD_MASTER);
    return [];
  }
  var data = sheet.getDataRange().getValues();
  var names = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][MASTER_COL.NAME - 1] || '').trim();
    var status = String(data[i][MASTER_COL.ENROLLMENT - 1] || '').trim();
    if (name && (status === '稼働' || status === '休止')) {
      names.push(name);
    }
  }
  return names;
}

/**
 * フォームアイテムの選択肢を更新する（リスト・ラジオ・チェックボックス対応）
 */
function setListChoices_(item, choices) {
  var type = item.getType();
  if (type === FormApp.ItemType.LIST) {
    item.asListItem().setChoiceValues(choices);
  } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    item.asMultipleChoiceItem().setChoiceValues(choices);
  } else if (type === FormApp.ItemType.CHECKBOX) {
    item.asCheckboxItem().setChoiceValues(choices);
  }
}

/**
 * 時間ベーストリガーを設定する（毎日AM1時に自動同期）
 * 一度だけ手動で実行してください
 */
function setupFormSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncFormDropdowns') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('syncFormDropdowns')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
  Logger.log('トリガー設定完了（毎日AM1時）');
  SpreadsheetApp.getUi().alert('トリガーを設定しました（毎日AM1時に自動同期）');
}
