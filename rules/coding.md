# コーディング規約

## 基本方針
- コメントは日本語で記載する
- 関数にはJSDocコメントを付ける
- エラーハンドリングを実装する（握りつぶし禁止、Logger.logで記録）

## GASファイル構成

```
gas/
├── main.gs              # エントリポイント・トリガー管理・ボタン実行関数
├── setup.gs             # F-01: シート初期セットアップ
├── monthly-summary.gs   # F-02: 月別集計更新
├── confirmed-visits.gs  # F-03: 確定来館記録生成
├── child-view.gs        # F-04: 児童別ビュー更新
├── visit-calendar.gs    # 来館カレンダー（日×児童マトリクス表示）
├── allocation.gs        # F-05/F-06: 余りポイント振り分け
├── email.gs             # F-07: 保護者メール送信
└── utils.gs             # 共通ユーティリティ（定数・ヘルパー関数）
```

## 命名規則

| 対象 | 規則 | 例 |
|---|---|---|
| 関数 | camelCase | `updateMonthlySummary` |
| 定数 | UPPER_SNAKE_CASE | `SHEET_NAME_MASTER` |
| 変数 | camelCase | `childName` |
| ファイル | kebab-case | `monthly-summary.gs` |

## 定数管理（utils.gs）

シート名・列インデックスはマジックナンバーを避け、定数として一元管理する。

```javascript
// シート名
const SHEET_NAMES = {
  FORM_RESPONSE: 'フォームの回答',
  CHILD_MASTER: '児童マスタ',
  MONTHLY_SUMMARY: '月別集計',
  VISIT_CALENDAR: '来館カレンダー',
  CONFIRMED_VISITS: '確定来館記録',
  CHILD_VIEW: '児童別ビュー',
};

// 児童マスタの列インデックス（1始まり）
const MASTER_COL = {
  NO: 1,
  NAME: 2,
  PARENT_NAME: 3,
  PARENT_EMAIL: 4,
  MONTHLY_QUOTA: 5,
  MEDICAL_TYPE: 6,
  STAFF: 7,
  ENROLLMENT: 8,
  VISIT_DAYS: 9,
  PRIORITY: 10,
};
```

## コーディングパターン

### シートの取得
```javascript
/** @returns {GoogleAppsScript.Spreadsheet.Sheet} */
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(`シート「${name}」が見つかりません`);
  }
  return sheet;
}
```

### データの一括読み書き
- `getDataRange().getValues()` で一括取得し、ループ内で `getValue()` を使わない
- 書き込みも `setValues()` で一括書き込みする
- 理由: API呼び出し回数を減らしパフォーマンスを確保するため

## 検証・完了条件

GASコードを変更した際は、以下を確認してからタスク完了とする：
- Apps Scriptエディタでの構文チェック（保存時に自動）
- テストデータでの動作確認
- Logger.log での処理結果確認
- 既存シートのデータが破損していないことの確認

## Git運用
- コミットメッセージ形式: `[type]: 変更内容`
  - type: `feat` / `fix` / `docs` / `refactor` / `chore`
- GASコードはローカルにも保管し、Git管理する
