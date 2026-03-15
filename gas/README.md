# GAS ファイル構成

児童施設の来館記録管理システムを構成する Google Apps Script ファイルの一覧と役割。

## ファイル一覧

### `main.gs` — エントリポイント・トリガー管理

システム全体の起点となるファイル。

- **`onOpen()`** — スプレッドシートを開いた時にカスタムメニュー「来館管理」を追加
- **`runMonthlyProcess()`** — 月次一括処理（手動）。年月選択ダイアログを表示し、振り分け → 来館カレンダー更新を一括実行
- **`runMonthlyProcessAutomatic()`** — 月次一括処理（トリガー用）。前日の属する月を対象に自動実行
- **`onEdit(e)`** — セル編集トリガー。児童別ビュー・来館カレンダー・月別集計の B1/B2 セル変更を検知して自動更新
- **`onFormSubmit(e)`** — フォーム送信トリガー。月別集計・確定来館記録・来館カレンダーを自動更新
- **`refreshDropdowns()`** — 各シートとGoogleフォームの児童名・年月ドロップダウンを最新の児童マスタで更新
- **`setupFormSubmitTrigger()`** / **`setupMonthlyProcessTrigger()`** — トリガーの初回セットアップ用
- **`updateConfirmedVisitsAndCalendar()`** — 確定来館記録と来館カレンダーの手動更新

---

### `setup.gs` — シート初期セットアップ（F-01）

スプレッドシートの各シートを作成し、ヘッダー・書式・ドロップダウンを初期設定する。初回1回のみ実行。

- **`setupAllSheets()`** — 月別集計・確定来館記録・来館カレンダー・児童別ビューの4シートを一括セットアップ
- 各シートのヘッダー行、列幅、行固定、データバリデーションを設定

---

### `utils.gs` — 共通ユーティリティ

全ファイルから参照される定数・ヘルパー関数を定義。

- **定数**: シート名（`SHEET_NAMES`）、各シートの列インデックス（`MASTER_COL`, `FORM_COL`, `SUMMARY_COL`, `CONFIRMED_COL` 等）、レイアウト定数、振り分けデフォルト値、メールテンプレート
- **ヘルパー関数**: `getSheet()`, `getChildMasterData()`, `getActiveChildren()`, `getFormResponsesByMonth()`, `parseYearMonth()`, `generateYearMonthOptions()`, `getChildNameOptions()`, `getConfirmedVisitsByMonth()` 等

---

### `monthly-summary.gs` — 月別集計更新（F-02）

確定来館記録（実データ＋振り分け）から児童ごとの来館数を集計し、月別集計シートに書き込む。

- **`updateMonthlySummary()`** — B1セルの対象年月を参照し、各児童の No.・児童名・月間利用枠・来館数・残数・利用率を算出して出力

---

### `confirmed-visits.gs` — 確定来館記録生成（F-03）

フォームの回答（実データ）を確定来館記録シートに転記する。振り分けデータは `allocation.gs` が管理するため、ここでは実データのみ扱う。

- **`updateConfirmedVisits(year, month)`** — 指定月の実データを洗い替え（他の月や振り分け行は保持）。引数省略時は全期間の実データを洗い替え

---

### `child-view.gs` — 児童別ビュー更新（F-04）

選択された児童・年月の確定来館記録を児童別ビューシートに書き込む。印刷にも対応。

- **`updateChildView()`** — B1（児童名）・B2（年月/すべて）を参照し、基本情報と来館履歴を表示
- 基本情報: 保護者名・担当スタッフ・月間利用枠・医療型の有無・来館回数/残枠/利用率
- **`prepareChildViewForPrint()`** / **`restoreChildViewFromPrint()`** — 印刷モードの切り替え

---

### `allocation.gs` — 余りポイント自動振り分け（F-05 / F-06）

月間利用枠に対する残枠を算出し、未来館日に自動振り分けする。振り分け結果は確定来館記録シートに「データ区分=振り分け」として直接書き込む。

- **`allocateRemainingPoints_(year, month)`** — メインロジック。優先度順に児童をソートし、来館曜日優先→その他の日で均等分散しながら振り分け
- **`runAllocationManual()`** — 手動実行（F-06）。二重実行時は確認ダイアログを表示
- **`runAllocationAutomatic()`** — 月初自動実行（F-05）
- 補完データは実来館データの最頻値から算出（入退所時間、体温、食事等）

---

### `visit-calendar.gs` — 来館カレンダー更新

日×児童のマトリクス形式で月間の来館状況を表示する。

- **`updateVisitCalendar()`** — B1セルの対象年月を参照してカレンダーを生成
- `○` = 実データ、`△` = 振り分けで視覚的に区別
- 土日祝の行を色分け表示（祝日は Google カレンダーから取得）
- 下部にサマリ行（月計・枠・残）を出力

---

### `email.gs` — 保護者向け来館報告メール送信（F-07）

来館記録を保護者にメールで通知する。

- **`sendDailyVisitReports()`** — 前日の来館記録を自動送信（毎朝8時トリガー）
- **`sendVisitReportsManual()`** — 日付指定による手動送信
- 児童マスタの保護者メールアドレスに対し、テンプレートベースでメールを作成・送信
- スクリプトプロパティ `FACILITY_NAME`, `EMAIL_SENDER_NAME` で施設名・送信者名を設定

---

### `appsscript.json` — プロジェクト設定

GAS プロジェクトのマニフェストファイル。タイムゾーン（`Asia/Tokyo`）、ランタイム（V8）を定義。

## トリガー一覧

| トリガー | 実行タイミング | 関数 | 設定方法 |
|---------|--------------|------|---------|
| フォーム送信時 | フォーム回答追加時 | `onFormSubmit` | `setupFormSubmitTrigger()` を1回実行 |
| 月次一括処理 | 毎月1日 午前3時 | `runMonthlyProcessAutomatic` | `setupMonthlyProcessTrigger()` を1回実行 |
| 月初振り分け | 毎月1日 午前2時 | `runAllocationAutomatic` | `setupAllocationTrigger()` を1回実行 |
| メール送信 | 毎朝8時 | `sendDailyVisitReports` | `setupEmailTrigger()` を1回実行 |
| セル編集時 | シート編集時 | `onEdit` | シンプルトリガー（自動） |
| シート起動時 | スプレッドシート起動時 | `onOpen` | シンプルトリガー（自動） |

## スクリプトプロパティ

| キー | 用途 |
|------|------|
| `FORM_ID` | Google フォームの ID（ドロップダウン更新用） |
| `FORM_CHILD_NAME_QUESTION` | フォーム内の児童名質問タイトル（デフォルト: `児童名`） |
| `FACILITY_NAME` | メール本文に使用する施設名 |
| `EMAIL_SENDER_NAME` | メール送信者名 |
