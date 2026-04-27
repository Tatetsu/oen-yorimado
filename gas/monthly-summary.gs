/**
 * F-02: 月別集計更新
 * フォームの回答（実記録）から集計し、月別集計シートに値を書き込む
 * B1=対象年、B2=対象月 の組み合わせでスコープを決定する
 */

/**
 * 月別集計を更新する（メイン処理）
 */
function updateMonthlySummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = ss.getActiveSheet();
  try {
    var sheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
    var yearStr = sheet.getRange('B1').getValue();
    var monthStr = sheet.getRange('B2').getValue();
    var scope = buildScope_(yearStr, monthStr);

    ss.toast('月別集計を更新中...', '読み込み中', -1);

    writeMonthlySummaryByScope_(sheet, scope);

    ss.toast('月別集計の更新が完了しました', '完了', 3);
    Logger.log('月別集計を更新しました: ' + describeScope_(scope));
  } catch (error) {
    logError_('updateMonthlySummary', error);
  } finally {
    originalSheet.activate();
  }
}

/**
 * スコープに応じて月別集計を書き込む
 * ヘッダー書式は setup で設定済みのため触らず、データ行のみ差し替える
 */
function writeMonthlySummaryByScope_(sheet, scope) {
  var masterData = getChildMasterData();

  // データエリアをクリア（スタイルは維持するため clearContent のみ）
  var lastRow = sheet.getLastRow();
  if (lastRow >= SUMMARY_DATA_START_ROW) {
    sheet.getRange(SUMMARY_DATA_START_ROW, 1, lastRow - SUMMARY_DATA_START_ROW + 1, 6).clearContent();
  }

  if (masterData.length === 0) {
    Logger.log('表示対象の児童がいません');
    return;
  }

  // 利用枠は scope に応じて月間/年間を切替
  var useAnnualQuota = (scope.type === 'year' || scope.type === 'all');
  var quotaColIdx = useAnnualQuota ? MASTER_COL.ANNUAL_QUOTA - 1 : MASTER_COL.MONTHLY_QUOTA - 1;

  // フォーム回答から来館数を集計
  var formResponses = collectFormResponsesByScope_(scope);
  var visitCounts = countVisitsByScope_(formResponses, scope);

  var outputData = masterData.map(function(row) {
    var childName = row[MASTER_COL.NAME - 1];
    var quota = row[quotaColIdx] || 0;
    var visits = visitCounts[childName] || 0;
    var remaining = quota > 0 ? quota - visits : '';
    var usageRate = quota > 0 ? visits / quota : 0;
    return [
      row[MASTER_COL.NO - 1],
      childName,
      quota || '',
      visits,
      remaining,
      usageRate,
    ];
  });

  sheet.getRange(SUMMARY_DATA_START_ROW, 1, outputData.length, 6).setValues(outputData);
  sheet.getRange(SUMMARY_DATA_START_ROW, SUMMARY_COL.USAGE_RATE, outputData.length, 1)
    .setNumberFormat('0%');
}

/**
 * スコープに応じてフォーム回答を取得する
 * @param {{type: string, year?: number, month?: number}} scope
 * @returns {Array<Array>} フォーム回答データ
 */
function collectFormResponsesByScope_(scope) {
  if (scope.type === 'month') return getFormResponsesByMonth(scope.year, scope.month);
  if (scope.type === 'year') return getFormResponsesByYear(scope.year);
  // month_all_years / all は全データを渡し、カウント側でスコープ判定する
  return getFormResponsesAll_();
}

/**
 * スコープに合致する日数を児童名ごとに数える（連泊は宿泊カレンダー全日をカウント）
 * 連泊2レコード等はペアリングして1宿泊として扱い、開始日〜終了日を1カウントずつ加算する。
 * 同一日が複数の論理1宿泊で重複するケース（運用上想定外）は1日として数える。
 * @param {Array<Array>} formResponses フォーム回答データ
 * @param {{type: string, year?: number, month?: number}} scope
 * @returns {Object} {児童名: 日数}
 */
function countVisitsByScope_(formResponses, scope) {
  var counts = {};
  var seenByChild = {}; // {児童名: {日付キー: true}} 同日重複防止
  var stays = pairStayRecords_(formResponses);
  stays.forEach(function(stay) {
    var childName = stay.childName;
    if (!childName) return;
    if (!seenByChild[childName]) seenByChild[childName] = {};
    var stayDates = expandStayToDates_(stay.recordDate, stay.checkIn, stay.checkOut);
    stayDates.forEach(function(d) {
      if (!matchesScope_(d, scope)) return;
      var key = formatDateKey_(d);
      if (seenByChild[childName][key]) return;
      seenByChild[childName][key] = true;
      counts[childName] = (counts[childName] || 0) + 1;
    });
  });
  return counts;
}
