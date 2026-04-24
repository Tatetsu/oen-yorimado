# WebView アクセス制御 仕様書

## 概要

フォーム回答修正用の WebView (`gas/web-view.gs` + `gas/index.html`) は、スプレッドシート付随の Google Apps Script として Web アプリ公開される。顧客が個人Gmailのため Google Workspace ドメイン制限が使えないこと、およびスタッフの入退に応じて即時にアクセス権を付与・剥奪する必要があることから、**トークン付きURL方式** で認可する。

---

## 認可フロー

```
1. 管理者が「許可ユーザー」シートに行追加（email, 氏名, 有効=TRUE）
2. メニュー「来館管理 > 許可ユーザーのURL発行」実行
   → トークン未発行の行に UUID(ハイフン除去) 32文字を自動生成
   → 有効ユーザー分のアクセスURLをダイアログ表示
3. 管理者が各スタッフにアクセスURLを配布（LINE/Chatwork等のDM推奨）
4. スタッフが URL にアクセス
   → Google アカウントログイン (GAS ANYONE 設定)
   → doGet(e) がクエリパラメータ t を validateToken_ で検証
   → 有効なら index.html をテンプレートレンダリング（token を埋め込み）
   → 無効なら拒否HTMLを返す
5. 画面内の google.script.run 呼び出しは全て token 付きで実行
   → サーバー側 requireValidToken_ が再検証
6. 退所時: 管理者が「有効」列のチェックを外す → 以降のアクセスは即拒否
```

---

## 許可ユーザーシートの構造

| 列 | 項目 | 形式 | 例 |
|----|------|-----|----|
| A | メールアドレス | 文字列 | `tanaka@example.com` |
| B | 氏名 | 文字列 | 田中 |
| C | トークン | 32文字英数字 | 自動発行（UUID） |
| D | 有効 | チェックボックス | ☑ TRUE / ☐ FALSE |
| E | 備考 | 自由記入 | 入職日 2026/05/01 等 |

- **C列（トークン）は手動入力しない**。メニュー「許可ユーザーのURL発行」で空欄行に自動発行される。
- **D列（有効）を FALSE にすると即時にアクセス拒否**（退所スタッフ対応）。行を削除しても同じ効果だが、監査目的で履歴を残したい場合は FALSE 運用推奨。
- トークン漏洩時は該当行のC列を削除して再度メニュー実行で新規発行される。

---

## 実装ファイル対応

| 責務 | ファイル | 関数 |
|------|---------|------|
| シート定数 | `gas/utils.gs` | `SHEET_NAMES.ALLOWED_USERS`, `ALLOWED_USERS_COL`, `ALLOWED_USERS_DATA_START_ROW` |
| トークン検証 | `gas/utils.gs` | `validateToken_(token)` |
| トークン生成 | `gas/utils.gs` | `generateAccessToken_()` |
| シート初期化 | `gas/setup.gs` | `setupAllowedUsersSheet_(ss)` |
| WebView ゲート | `gas/web-view.gs` | `doGet(e)`, `requireValidToken_(token)` |
| API 再検証 | `gas/web-view.gs` | `getInitialDataWeb` / `getFormResponsesWeb` / `updateFormResponseWeb` / `deleteFormResponseWeb` の第1引数 |
| メニュー & URL発行 | `gas/main.gs` | `issueAllowedUserUrls()` |
| クライアント側 | `gas/index.html` | `ACCESS_TOKEN` 変数、各 `google.script.run.*(ACCESS_TOKEN, ...)` |

---

## セキュリティ設計

### 二重防御

1. **`doGet` レイヤ**: トークン検証に失敗すると index.html をそもそもレンダリングしない
2. **API レイヤ**: 各 `google.script.run` 呼び出しで `requireValidToken_` を再実行。退所即時反映（開いたままのタブは次の操作で弾かれる）

### トークン設計

- `Utilities.getUuid()` → 32文字の16進英数字（推測困難）
- URL に付与（GET parameter）。HTTPS 通信のため通信路での盗聴耐性あり
- ブラウザ履歴・URL共有で漏洩するリスクは残る → 配布時は本人DMで送付、流出疑いは該当行のトークン再発行で即無効化

### `appsscript.json` のアクセス設定

- `access: "ANYONE"` のまま（Googleアカウントログインは必須）
- `executeAs: "USER_DEPLOYING"` のまま（スプレッドシート操作権限を保つため）

この組み合わせで、デプロイURL直アクセス時も「Googleログイン → トークンチェック」の二段ゲートが働く。

---

## 運用手順

### 新規スタッフ追加

1. 「許可ユーザー」シートを開く
2. 末尾行に以下を入力
   - A: メールアドレス
   - B: 氏名
   - D: チェックON
3. メニュー「来館管理 > 許可ユーザーのURL発行」を実行
4. ダイアログからURLをコピーしてスタッフにDM送付

### 退所スタッフ排除

1. 「許可ユーザー」シートの該当行のD列（有効）チェックを外す
2. 以上。即時反映される（次回アクセス・次回API呼び出しから拒否）

### トークン漏洩時の再発行

1. 該当行のC列（トークン）を削除
2. メニュー「許可ユーザーのURL発行」を実行
3. 新しいURLを本人に配布

---

## 拒否時の挙動

| アクセスパターン | 挙動 |
|-----------------|------|
| `?t=` なしで直アクセス | 拒否HTML表示（データは一切返らない） |
| 不正なトークン | 拒否HTML表示 |
| 有効=FALSEのユーザーのトークン | 拒否HTML表示 |
| 有効なトークンで画面表示後、管理者が有効=FALSE化 | 次回の google.script.run 呼び出しで例外 → エラートースト表示 |

---

## 制限・注意点

- GAS Web アプリではブラウザ標準の Basic 認証は使用不可（レスポンスヘッダ操作不可のため）
- アクセスログは現時点で未実装（必要に応じて `validateToken_` に呼び出し時のタイムスタンプ記録を追加する）
- デプロイURLを更新した場合、既発行URLのホスト部分は無効になるため、メニューから再発行が必要
