# TsukeTsuke — 大津家家計簿（GitHub Pages + Apps Script API）

## 構成
- `index.html` … 画面（GitHub Pages）。Apps Script を JSON API として fetch で呼ぶ
- `apps-script/Code_GitHub対応.gs` … Apps Script（家計簿スプシに紐づけ）。API・OCR・書き込み・キャッシュ

## Apps Script 側（更新時）
1. 家計簿スプレッドシート → 拡張機能 → Apps Script
2. `apps-script/Code_GitHub対応.gs` の中身を `コード.gs` に丸ごと貼り替え → 保存
3. `Index.html` は不要（あれば削除してOK）
4. デプロイ → デプロイを管理 → 鉛筆 → 「新バージョン」→ デプロイ（URLは変わらない）

初回のみ：サービスに Drive API を追加 → スプシで「📸 レシート → 初期設定」→ ウェブアプリとしてデプロイ（実行ユーザー：自分 / アクセス：全員）→ `/exec` URL を `index.html` の `DEFAULT_GAS_URL` に設定。

## GitHub Pages 側
`index.html` `manifest.webmanifest` `sw.js` `icon-*.png` をリポジトリ直下に置き、Settings → Pages で main / root を公開。
