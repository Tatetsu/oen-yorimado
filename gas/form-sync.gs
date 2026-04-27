/**
 * フォームのドロップダウンをマスタデータに同期する
 */

// フォームの質問タイトル（実際の質問文と異なる場合は修正すること）
var FORM_TITLE_STAFF1 = 'スタッフ1';
var FORM_TITLE_STAFF2 = 'スタッフ2';
var FORM_TITLE_CHILD = '児童名';
var FORM_TITLE_CHECK_IN = '入所日時';
var FORM_TITLE_CHECK_OUT_OLD = '退所日時';
var FORM_TITLE_CHECK_OUT = '退所予定日時';
var FORM_TITLE_OVERNIGHT = '連泊';
var FORM_HELP_OVERNIGHT = '連泊の場合はチェック。初日は退所予定空欄／中日は両方空欄／最終日は入所空欄で送信してください。';

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
    } else if (title === FORM_TITLE_CHECK_OUT_OLD) {
      // 旧名「退所日時」→「退所予定日時」へリネーム（一度だけ実行される）
      item.setTitle(FORM_TITLE_CHECK_OUT);
    }
  });

  // 新仕様（ユニーク宿泊キー方式）では連泊フラグは不要のため、スキーマ保証呼び出しは無効化。
  // 関数本体は残してロールバック可能にしている（必要なら下記のコメントを外す）。
  // ensureOvernightSchema_(form);

  Logger.log('同期完了: スタッフ ' + staffNames.length + '名 / 児童 ' + childNames.length + '名');
}

/**
 * 連泊対応のフォームスキーマを保証する
 * - 入所日時・退所予定日時を「必須」から「任意」に変更
 * - 連泊フラグ（チェックボックス）を末尾に追加（既存なら何もしない）
 * @param {GoogleAppsScript.Forms.Form} form
 */
function ensureOvernightSchema_(form) {
  var hasOvernight = false;

  form.getItems().forEach(function(item) {
    var title = item.getTitle();
    var type = item.getType();
    if (title === FORM_TITLE_OVERNIGHT) {
      hasOvernight = true;
      return;
    }
    if (title === FORM_TITLE_CHECK_IN || title === FORM_TITLE_CHECK_OUT) {
      // 入所/退所予定の必須を解除
      if (type === FormApp.ItemType.DATETIME) {
        item.asDateTimeItem().setRequired(false);
      } else if (type === FormApp.ItemType.DATE) {
        item.asDateItem().setRequired(false);
      } else if (type === FormApp.ItemType.TIME) {
        item.asTimeItem().setRequired(false);
      }
    }
  });

  if (!hasOvernight) {
    var item = form.addCheckboxItem();
    item.setTitle(FORM_TITLE_OVERNIGHT)
        .setHelpText(FORM_HELP_OVERNIGHT)
        .setChoiceValues(['連泊'])
        .setRequired(false);
    Logger.log('連泊フラグをフォームに追加しました');
  }
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
  setupTimeTrigger_('syncFormDropdowns', { everyDays: 1, atHour: 1 });
  Logger.log('フォーム同期トリガー設定完了（毎日AM1時）');
}
