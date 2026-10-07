# gcal-booking

FF14 登録サポート用の予約ページです。

## 構成

- `index.html` — GitHub Pagesで公開する予約画面
- `Code.gs` — Googleスプレッドシートに紐づけるApps Script

## Googleスプレッドシート

スプレッドシートを1つ作り、シート名を `予約管理` にします。

1行目の列は以下です。

|列|内容|
|---|---|
|A|キャラクター名|
|B|予約状態|
|C|予約日時|
|D|集合DC|
|E|予約時入力名|
|F|名前照合|
|G|対応状況|
|H|手動紐付け先|
|I|備考|

A列の2行目以降へ既存メンバーのキャラクターフルネームを入れます。

## Apps Script

1. スプレッドシートで「拡張機能」→「Apps Script」
2. `Code.gs` の内容を貼り付けて保存
3. `setupSheet` を一度実行して権限を許可
4. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」
5. 実行ユーザー: 自分
6. アクセスできるユーザー: 全員
7. 発行された `.../exec` URL をコピー
8. `index.html` の `API_URL` にそのURLを設定

## GitHub Pages

Repository Settings → Pages → Build and deployment → Deploy from a branch

- Branch: `main`
- Folder: `/(root)`

公開URL:

https://jogu6.github.io/gcal-booking/

## 名前照合

比較時のみ以下を行います。

- Unicode NFKCで全角英数字・全角スペース等を正規化
- 前後の空白を削除
- 連続空白を半角スペース1つに統一
- 英字を小文字化

自動一致しない予約は同じシート下部へ「⚠ 未照合 / 要確認」として追加されます。
H列へ正しいメンバー名を入力後、スプレッドシートの「予約管理」メニュー → 「手動紐付けを反映」で統合できます。
