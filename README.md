# 家計簿レシート登録（GitHub Pages版）

## GitHubへアップするファイル
リポジトリ直下へ以下をアップロードします。

- `index.html`
- `manifest.webmanifest`
- `sw.js`
- `icon-192.png`
- `icon-512.png`

`apps-script/` フォルダはGoogle Apps Script用なので、GitHub Pagesでは実行されません。

## GitHub PagesをONにする
GitHub → リポジトリ → Settings → Pages

- Source: `Deploy from a branch`
- Branch: `main`
- Folder: `/ (root)`

保存後に表示されたURLをスマホで開きます。

## Apps Script側
1. 家計簿スプレッドシート → 拡張機能 → Apps Script
2. `apps-script/Code_GitHub対応.gs` を `Code.gs` に貼る
3. `apps-script/Index.html` をHTMLファイル `Index` に貼る
4. Apps Script → サービス → Google Drive API を追加
5. スプレッドシートを再読込 → `📸 レシート` → `初期設定`
6. Apps Script → デプロイ → 新しいデプロイ → ウェブアプリ
7. 発行された `/exec` URL をGitHub Pagesの初回設定画面に貼る

### 重要
GitHub Pages内にApps Script画面を表示するため、`doGet()` に
`setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)` を入れています。

## 既存のスプレッドシート入力
今まで通り直接入力できます。
この仕組みはレシート入力経路を追加するだけで、既存入力方法は変更しません。
