# システム構成図

## 全体アーキテクチャ

```mermaid
graph TD
    A[スタッフ] -->|入力| B[Googleフォーム]
    B -->|自動反映| C[スプレッドシート<br/>フォームの回答]

    C --> D[GAS<br/>ビジネスロジック]
    E[スプレッドシート<br/>児童マスタ] --> D

    D -->|集計書き込み| F[スプレッドシート<br/>月別集計]
    D -->|振り分け書き込み| G[スプレッドシート<br/>振り分け記録]
    D -->|統合書き込み| H[スプレッドシート<br/>確定来館記録]
    H -->|参照・書き込み| I[スプレッドシート<br/>児童別ビュー]

    I -->|印刷| J[印刷出力]

    D -->|Phase 3| K[Gmail<br/>保護者メール送信]

    style D fill:#4285F4,color:#fff
    style B fill:#673AB7,color:#fff
    style K fill:#EA4335,color:#fff
```

## コンポーネント説明

| コンポーネント | 役割 | 使用技術 |
|---|---|---|
| Googleフォーム | スタッフの来館記録入力UI | Google Forms |
| スプレッドシート | データストア + 閲覧UI | Google Sheets |
| GAS | 集計・振り分け・データ統合のビジネスロジック | Google Apps Script (JavaScript) |
| Gmail | 保護者への来館報告メール送信（Phase 3） | Gmail API via GAS |

## GAS処理フロー

### フォーム送信時トリガー

```mermaid
sequenceDiagram
    participant Staff as スタッフ
    participant Form as Googleフォーム
    participant Sheet as フォームの回答
    participant GAS as GAS
    participant Summary as 月別集計
    participant Confirmed as 確定来館記録

    Staff->>Form: 来館記録入力
    Form->>Sheet: 自動反映
    Sheet->>GAS: onFormSubmitトリガー
    GAS->>GAS: 月別集計を再計算
    GAS->>Summary: 値を書き込み
    GAS->>GAS: 確定来館記録を再生成
    GAS->>Confirmed: 値を書き込み
```

### 月初自動振り分け（月初1日トリガー）

```mermaid
sequenceDiagram
    participant Trigger as 時間トリガー（毎月1日）
    participant GAS as GAS
    participant Master as 児童マスタ
    participant Form as フォームの回答
    participant Alloc as 振り分け記録
    participant Confirmed as 確定来館記録
    participant Summary as 月別集計

    Trigger->>GAS: 実行開始
    GAS->>Master: 児童情報取得（枠・曜日・優先度）
    GAS->>Form: 前月の来館データ取得
    GAS->>GAS: 残枠算出
    GAS->>GAS: 振り分けロジック実行
    GAS->>Alloc: 振り分け結果書き込み
    GAS->>Confirmed: 確定来館記録再生成
    GAS->>Summary: 月別集計更新
```

### 児童別ビュー更新（手動ボタン）

```mermaid
sequenceDiagram
    participant Staff as スタッフ
    participant View as 児童別ビュー
    participant GAS as GAS
    participant Master as 児童マスタ
    participant Confirmed as 確定来館記録

    Staff->>View: 児童名・年月を選択
    Staff->>View: 更新ボタン押下
    View->>GAS: 実行
    GAS->>Master: 基本情報取得
    GAS->>Confirmed: 該当データ抽出
    GAS->>View: 値を書き込み
```

## トリガー一覧

| トリガー種別 | タイミング | 実行関数 | Phase |
|---|---|---|---|
| onFormSubmit | フォーム送信時 | updateMonthlySummary, generateConfirmedRecords | 1 |
| 時間ベース | 毎月1日 | autoAllocatePoints | 2 |
| 時間ベース | 毎朝（時刻指定） | sendParentEmail | 3 |
| ボタン | 手動 | updateMonthlySummary | 1 |
| ボタン | 手動 | updateChildView | 1 |
| ボタン | 手動 | runAllocation | 2 |

## ファイル構成（GAS）

```
gas/
├── main.gs              # エントリポイント・トリガー管理
├── setup.gs             # F-01: シート初期セットアップ
├── monthly-summary.gs   # F-02: 月別集計更新
├── confirmed-visits.gs  # F-03: 確定来館記録生成
├── child-view.gs        # F-04: 児童別ビュー更新
├── allocation.gs        # F-05/F-06: 余りポイント振り分け
├── email.gs             # F-07: 保護者メール送信（Phase 3）
├── monthly-close.gs     # F-08: 月次確定処理（Phase 3）
└── utils.gs             # 共通ユーティリティ
```
