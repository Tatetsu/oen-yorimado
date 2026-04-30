/**
 * 旧連泊ロジック（空欄パターン方式）から新仕様（ユニーク宿泊キー方式）への
 * データマイグレーション。手動実行のみ。トリガー禁止。
 *
 * 処理内容:
 *   1. フォームの回答シートを走査
 *   2. 旧連泊空欄パターン（入所のみ・退所のみ・両方空欄）のレコードを児童名でペアリング
 *   3. ペアリング結果から各レコードに入所日時・退所日時を埋める
 *   4. 連泊フラグ列の値はそのまま残す（後方互換）
 *
 * 安全策:
 *   - ドライランモードあり（dryRun=true で書き込まずログ出力のみ）
 *   - 既に両方記入済みのレコードはスキップ
 *   - 警告がある stay（孤立等）は書き込みをスキップしてログ出力
 */

/**
 * 手動実行エントリーポイント（実書き込み）
 * メニューやエディタから実行する想定
 */
function migrateLegacyOvernightRecords() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    '旧連泊レコードのマイグレーション',
    'フォームの回答シートを新仕様（全レコード入退所両方記入）に変換します。\n' +
    'バックアップを取った上で実行してください。\n\n実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;

  var result = runMigrateLegacyOvernightRecords_(false);
  ui.alert('マイグレーション完了',
    '更新行数: ' + result.updatedCount + '\n' +
    'スキップ行数: ' + result.skippedCount + '\n' +
    '警告: ' + result.warnings.length + '件\n\n' +
    '詳細はログシートを確認してください。',
    ui.ButtonSet.OK);
}

/**
 * ドライラン（書き込みなし）。差分をログ出力のみ。
 */
function migrateLegacyOvernightRecordsDryRun() {
  var result = runMigrateLegacyOvernightRecords_(true);
  SpreadsheetApp.getUi().alert(
    'ドライラン完了',
    '更新予定行数: ' + result.updatedCount + '\n' +
    'スキップ行数: ' + result.skippedCount + '\n' +
    '警告: ' + result.warnings.length + '件\n\n' +
    '詳細はログを確認してください（書き込みは行っていません）。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * マイグレーション本体
 * @param {boolean} dryRun true=書き込みせずログのみ
 * @returns {{updatedCount: number, skippedCount: number, warnings: Array<string>}}
 */
function runMigrateLegacyOvernightRecords_(dryRun) {
  var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
  var lastRow = sheet.getLastRow();
  if (lastRow < FORM_DATA_START_ROW) {
    return { updatedCount: 0, skippedCount: 0, warnings: [] };
  }

  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(FORM_DATA_START_ROW, 1, lastRow - FORM_DATA_START_ROW + 1, lastCol).getValues();

  // 旧式の状態機械ペアリングを再現するため、フォーム由来の行を児童ごとに記録日順で走査する
  // pairLegacyOvernight_: 入所空欄/退所空欄/両方空欄のパターンから1宿泊にまとめる
  var stays = pairLegacyOvernight_(data);

  // sheetIndex: フォームシート上の行番号 → stay の checkIn/checkOut にマッピング
  var rowOverride = {}; // {sheetRowIndex: {checkIn, checkOut}}
  var warnings = [];

  stays.forEach(function(stay) {
    if (stay.warning) {
      warnings.push(stay.warning);
      return;
    }
    if (!stay.checkIn || !stay.checkOut) {
      warnings.push('未確定の宿泊（児童=' + stay.childName + ' 開始=' + (stay.checkIn || '?') + '）→ スキップ');
      return;
    }
    stay.dataRowIndices.forEach(function(idx) {
      rowOverride[idx] = { checkIn: stay.checkIn, checkOut: stay.checkOut };
    });
  });

  var updatedCount = 0;
  var skippedCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var sheetRowIndex = i + FORM_DATA_START_ROW;
    var checkIn = row[FORM_COL.CHECK_IN - 1];
    var checkOut = row[FORM_COL.CHECK_OUT - 1];
    var hasIn = (checkIn instanceof Date) && checkIn.getFullYear() >= 1900;
    var hasOut = (checkOut instanceof Date) && checkOut.getFullYear() >= 1900;

    if (hasIn && hasOut) {
      skippedCount++;
      continue;
    }

    var override = rowOverride[i];
    if (!override) {
      skippedCount++;
      continue;
    }

    if (!dryRun) {
      sheet.getRange(sheetRowIndex, FORM_COL.CHECK_IN).setValue(override.checkIn);
      sheet.getRange(sheetRowIndex, FORM_COL.CHECK_OUT).setValue(override.checkOut);
      sheet.getRange(sheetRowIndex, FORM_COL.CHECK_IN, 1, 2).setNumberFormat('yyyy/mm/dd H:mm');
    }
    updatedCount++;
    Logger.log('[migrate] row=' + sheetRowIndex + ' 児童=' + row[FORM_COL.CHILD_NAME - 1] +
      ' → 入=' + formatDateYMD_(override.checkIn, 'yyyy/MM/dd HH:mm') +
      ' 退=' + formatDateYMD_(override.checkOut, 'yyyy/MM/dd HH:mm'));
  }

  warnings.forEach(function(w) { Logger.log('[migrate-warn] ' + w); });
  Logger.log('[migrate] 完了: 更新=' + updatedCount + ' スキップ=' + skippedCount + ' 警告=' + warnings.length + (dryRun ? ' (dryRun)' : ''));

  return { updatedCount: updatedCount, skippedCount: skippedCount, warnings: warnings };
}

/**
 * 旧仕様（空欄パターン方式）でフォーム回答を1宿泊にまとめる
 * - 入所あり・退所空欄 → 連泊開始
 * - 両方空欄         → 連泊中日（直近の開始に紐付く）
 * - 入所空欄・退所あり → 連泊終了
 * - 入所あり・退所あり → 単泊（そのまま採用）
 *
 * 返り値の各 stay には dataRowIndices（data 配列内のインデックス）が入っており、
 * checkIn/checkOut が確定したらそれぞれを埋める。
 *
 * @param {Array<Array>} data フォームの回答行（ヘッダー除外済み、シート行順）
 * @returns {Array<{childName, checkIn, checkOut, dataRowIndices, warning}>}
 */
function pairLegacyOvernight_(data) {
  var byChild = {};
  data.forEach(function(row, idx) {
    var name = row[FORM_COL.CHILD_NAME - 1];
    if (!name) return;
    if (!byChild[name]) byChild[name] = [];
    byChild[name].push({ row: row, idx: idx });
  });

  var stays = [];
  Object.keys(byChild).forEach(function(name) {
    var records = byChild[name];
    records.sort(function(a, b) {
      var da = getRowRecordDate_(a.row);
      var db = getRowRecordDate_(b.row);
      var t = ((da && da.getTime()) || 0) - ((db && db.getTime()) || 0);
      return t !== 0 ? t : (a.idx - b.idx);
    });

    var openStay = null;
    records.forEach(function(rec) {
      var row = rec.row;
      var checkIn = row[FORM_COL.CHECK_IN - 1];
      var checkOut = row[FORM_COL.CHECK_OUT - 1];
      var hasIn = (checkIn instanceof Date) && checkIn.getFullYear() >= 1900;
      var hasOut = (checkOut instanceof Date) && checkOut.getFullYear() >= 1900;

      if (hasIn && hasOut) {
        if (openStay) {
          stays.push({
            childName: name,
            warning: '連泊終了レコード未送信（次の単泊で打ち切り）→ 児童=' + name,
            dataRowIndices: openStay.dataRowIndices,
          });
          openStay = null;
        }
        // 単泊は対象外（既に両方記入済み）
        return;
      }

      if (hasIn && !hasOut) {
        if (openStay) {
          stays.push({
            childName: name,
            warning: '連泊終了レコード未送信（次の連泊開始で打ち切り）→ 児童=' + name,
            dataRowIndices: openStay.dataRowIndices,
          });
        }
        openStay = {
          childName: name,
          checkIn: checkIn,
          checkOut: null,
          dataRowIndices: [rec.idx],
        };
        return;
      }

      if (!hasIn && hasOut) {
        if (openStay) {
          openStay.checkOut = checkOut;
          openStay.dataRowIndices.push(rec.idx);
          stays.push(openStay);
          openStay = null;
        } else {
          stays.push({
            childName: name,
            warning: '連泊開始なしの終了レコード（孤立）→ 児童=' + name,
            dataRowIndices: [rec.idx],
          });
        }
        return;
      }

      // 両方空欄（連泊中日）
      if (openStay) {
        openStay.dataRowIndices.push(rec.idx);
      } else {
        stays.push({
          childName: name,
          warning: '連泊開始なしの中日レコード（孤立）→ 児童=' + name,
          dataRowIndices: [rec.idx],
        });
      }
    });

    if (openStay) {
      stays.push({
        childName: name,
        warning: '連泊終了レコード未送信（オープンのまま）→ 児童=' + name,
        dataRowIndices: openStay.dataRowIndices,
      });
    }
  });

  return stays;
}
