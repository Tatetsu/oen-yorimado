# トリガー一覧

## 概要

GASには4種類の時間ベーストリガーと `onEdit` / `onOpen` が設定される。時間ベーストリガーは `setupAllSheets` 実行時に一括登録される。

---

## 時間ベーストリガー

| 実行タイミング | 関数名 | 処理内容 | 設定関数 |
|-------------|-------|---------|---------|
| 毎日 AM1時 | `syncFormDropdowns` | フォームの児童名・スタッフプルダウンをマスタで同期 | `setupFormSyncTrigger()` |
| 毎月1日 AM3時 | `runMonthlyProcessAutomatic` | 前月を対象に振り分け・確定来館記録・来館カレンダーを自動実行 | `setupMonthlyProcessTrigger()` |
| 毎朝 AM8時 | `sendDailyVisitReports` | 前日の来館記録を保護者にメール送信 | `setupEmailTrigger()` |
| 毎日 AM9時 | `collectBounceEmails` | Gmail NDR を検索してバウンスログに記録、検出時は通知メール送信 | `setupBounceCheckTrigger()` |

> `runAllocationAutomatic`（`setupAllocationTrigger` で設定）は旧実装。現在は `runMonthlyProcessAutomatic` に統合済み。既存環境で両方が登録されている場合は重複実行になるため、`runAllocationAutomatic` 側は削除すること。

### 月初自動処理の対象月

`runMonthlyProcessAutomatic` は **実行日-1日** の月を対象にする。

```
実行: 毎月1日 AM3時
対象: getDate()-1 = 前月末日 → その月 = 前月
```

---

## onEdit トリガー（自動）

セル編集時に `onEdit(e)` が発火し、シート・セルに応じた処理を実行する。

| シート名 | 変更セル | 実行処理 |
|---------|---------|---------|
| 児童別ビュー | B1 / B2 / B3 | `updateChildView()` |
| 来館カレンダー | B1 / B2 | `updateVisitCalendar()` |
| 月別集計 | B1 / B2 | `updateMonthlySummary()` |
| 確定来館記録 | B1 / B2 | `filterConfirmedVisits_()` |

実装: `gas/main.gs` → `onEdit`

---

## カスタムメニュー（手動実行）

スプレッドシートを開いた際に「来館管理」メニューが追加される（`onOpen`）。

| メニュー項目 | 関数名 | 処理内容 |
|------------|-------|---------|
| 初期セットアップ | `setupAllSheets` | 全シートの初期作成・ヘッダー・書式設定、および全時間ベーストリガーの登録 |
| 月次一括処理 | `runMonthlyProcess` | 年月選択ダイアログ → 振り分け・確定来館記録・月別集計・来館カレンダー一括更新 |
| 確定来館記録を手動更新 | `updateConfirmedVisitsAndCalendar` | 実データで確定来館記録・来館カレンダーを再生成 |
| 来館報告メール手動送信 | `sendVisitReportsManual` | HTMLダイアログで日付選択 → 保護者メール送信 |
| ドロップダウンを更新 | `refreshDropdowns` | 全ビューシートの年月ドロップダウン + 児童マスタの重度支援区分 + フォーム（児童名・スタッフ1・スタッフ2）を更新 |
| バウンスメールを確認 | `collectBounceEmailsManual` | バウンスメールを検出してバウンスログシートに記録 |

---

## トリガー設定の初回手順

「来館管理」メニュー →「初期セットアップ」（`setupAllSheets`）を1回実行すれば、上記の時間ベーストリガー4件がすべて登録される。個別の `setup*Trigger()` を手動実行する必要はない。
