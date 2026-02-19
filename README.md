# 日報投稿アプリ

HTML/CSS/JavaScript で作成した日報投稿アプリです。

## できること
- スタッフ画面
  - 最初はスタッフ名一覧
  - スタッフ名タップで日報履歴（`日付 / イベント会場 / 確認状態`）
  - `編集` と `内容確認` が可能
- 管理者画面（ログイン必須）
  - 未確認一覧: `スタッフ名 / 稼働日` + `日報を見る`
  - 日報を見る画面: `確認済みにする / 未分類 / フォルダ保存 / 内容確認 / 編集`
  - 確認済み一覧: スタッフ別表示
- Googleスプレッドシート同期
  - 保存/更新/確認状態変更/フォルダ変更/削除を自動同期
  - 全件再同期ボタンあり
  - シートから取得ボタンあり（他端末の更新を反映）

## 管理者ログイン
- `admin01 / admin1234`
- `sv01 / sv1234`

変更する場合は `app.js` の `ADMIN_USERS` を編集してください。

## 社内共有URL化（公開）
このアプリは静的ファイルなので、社内公開は以下のどれでも可能です。

1. GitHub Pages
- このフォルダをGitHubリポジトリへ push
- Settings → Pages → Deploy from branch
- 公開URLを社内共有

2. Vercel（推奨）
- Vercelで新規プロジェクト作成
- `日報レポート` フォルダをデプロイ
- 発行されたURLを共有

3. 社内Webサーバ
- `index.html / style.css / app.js` を社内サーバへ配置
- 社内URLを発行

## Googleスプレッドシート連携手順
### 1) Apps Scriptを作成
- Googleスプレッドシートを1つ作成
- 拡張機能 → Apps Script
- `google-apps-script/Code.gs` の内容を貼り付け

### 2) 連携キー設定（推奨）
- Apps Script の「プロジェクトの設定」または実行コードで
  - Script Property: `SYNC_TOKEN = 任意の秘密文字列`

### 3) Webアプリとしてデプロイ
- 「デプロイ」→「新しいデプロイ」→ 種別「ウェブアプリ」
- 実行ユーザー: 自分
- アクセス権: リンクを知っている全員（社内のみ運用なら社内権限に合わせて設定）
- 発行された `https://script.google.com/macros/s/.../exec` をコピー

### 4) アプリにURL設定
- 管理者でログイン
- 「Googleスプレッドシート連携」に
  - Apps Script WebアプリURL
  - 連携キー（SYNC_TOKEN）
- 「連携設定を保存」
- 必要なら「全件をシートへ再同期」
- 他端末で使う時は「シートから取得」を押すと最新が反映されます

## 他端末に反映されない時
- 原因: `localStorage` は端末ごとに別保存
- 対策:
  1. すべての端末で同じ公開URLを開く
  2. `app.js` の `DEFAULT_SYNC_CONFIG` に Apps Script URL/連携キーを設定して再デプロイ
  3. 「シートから取得」を実行（または30秒待つと自動反映）
- Apps Script側は `Code.gs` を最新版に更新し、**再デプロイ**してください（doGet追加のため）

## スタッフも自動同期させる設定
- `app.js` の先頭にある以下を設定してください:
  - `DEFAULT_SYNC_CONFIG.endpoint`
  - `DEFAULT_SYNC_CONFIG.token`
- これで管理者ではないスタッフ端末でも、起動時と30秒ごとにシートから自動取得されます。

## 保存先
- 日報: `localStorage` (`daily-report-app-v1`)
- 管理者セッション: `daily-report-admin-session-v1`
- シート連携設定: `daily-report-sync-config-v1`
