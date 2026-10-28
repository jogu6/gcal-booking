# 掲示板への投稿API v1

更新日：2026/10/28（Markdownの取扱いに関する誤記を訂正）。Codex等の送信側実装者向け。Discord Webhook形式の一部を取り込むAPIであり、Discord APIの完全互換ではありません。

## 発行と管理

開発者で通常の掲示板を開き、掲示板の設定 → Webhookの設定から作成します。1掲示板に最大10件。表示名を自由に設定し、投稿者の自画像も選択できます。「受信URLをコピー」で得たURLを送信側のSecretへ保存してください。

開発 → Webhookのメンバー管理でも、対象だけの一覧・名前検索・名前/自画像/受信可否の変更・URL再発行・削除を行えます。ギルドマスターや通常メンバーにはこの管理権限がありません。投稿者はゲームに所属せず、通常メンバー一覧・イベント招待・人数には含まれません。名前と自画像は投稿時のものを過去コメントに保存します。

URL再発行は旧URLを即無効化します。新URLを送信側でも保存し直してください。URLの削除は過去コメントを消しません。停止/削除済み/期限切れ掲示板には投稿できません。システム管理のエラー報告/認証再発行依頼掲示板は対象外です。

## 接続

```text
POST https://<発行元ホスト>/hooks/board/<Webhook UUID>/<43文字トークン>
Content-Type: application/json
Idempotency-Key: <一つの通知イベントに固有のID>
```

URL自身が資格情報です。ユーザーのログイン、Cookie、ブラウザーのOriginヘッダーは不要です。URLを公開コード・HTML・実行ログ・エラー本文に出さないでください。対象ギルド/環境/掲示板は発行時に固定され、送信データから変更できません。

`Idempotency-Key`は推奨。英数字、`-_.:`の1〜128文字。UUIDを推奨します。再送時も同じIDと同じ内容を使い、予約を保存し直さず通知だけを再送してください。同じ本文でも別イベントは別IDです。省略したDiscord形式の送信も受付可能ですが、その再送を重複防止できません。重複判定は365日保持します。

## 本文とDiscord形式

最小例：

```json
{"content":"新規予約\nキャラクター：Test Player\n予約日時：2026-10-18 20:00"}
```

Discord形式の例：

```json
{
  "content": "**新しい予約があります**",
  "embeds": [{
    "title": "予約受付",
    "description": "担当者の確認をお願いします。",
    "fields": [{"name":"キャラクター","value":"Test Player","inline":true}],
    "image": {"url":"https://<公開ホスト>/image.png"}
  }]
}
```

本文はcontent、続いて各embedのtitle、description、fieldsの名前と値、footer.textの順に改行区切りで連結します。embedsは最大10件、各fieldsは最大25件です。

- 受信側はMarkdownをそのまま扱います（2026/10/28、運用者確認。旧版の「書式を除去する」という記載を訂正）。送信側は見出し・太字・箇条書きで本文を整形できます。入力された名前等の値だけをエスケープし、送信側が付けた書式を一括エスケープしないでください。HTMLは使用せず、画像は下記の専用項目で指定してください。
- 本文はUnicodeコードポイントで最大1,000文字。書式記号を含めるかの上限判定詳細は未確認のため、送信側は書式記号を含む本文を1,000文字以内に収めてください。JavaScriptでは`Array.from(text).length`で数えます。改行も1文字。上限超過は400となり、黙って切り捨てません。本文が空でも有効な画像があれば投稿できます。
- username/avatar_urlは受け付けますが無視し、開発者が設定した名前・自画像を使います。色、inline、メンション、tts、components、flags、thread_name、applied_tags、poll等の表示/動作は取り込みません。メンション先へ通知を強制しません。
- 未対応のトップレベル項目や不正な型は400。JSON全体はUTF-8で32KiBまで。`?wait=true`/`?wait=false`は許容しますが、どちらも保存確定後に結果を返します。それ以外のクエリ指定は404です。

## 画像

`attachments[].url`、`embeds[].image.url`、`embeds[].thumbnail.url`から、異なるURLの画像を最大2枚まで取り込みます。枚数はWebアプリの共通コメント設定に従います。同じURLは1枚にまとめます。

```json
{"content":"参考画像です","attachments":[{"url":"https://<公開ホスト>/photo.jpg"}]}
```

送信側でWebPへ変換する必要はありません。JPEG・PNG・WebP・HEIC/HEIFを読み取り、WebP本体とサムネイルに変換して既存の画像保存先へ格納します。アニメーション等は対象外です。

画像の入力上限は公開テスト256MiB（本番64MiB）、80,000,000画素。本体は長辺4096px以内、サムネイル384px以内。変換後の容量制限も通常の画像仕様に従います。メタデータは保存しません。multipart/form-dataによるファイル本体の送信は今回未対応です。

画像URLは認証なしで取得できる公開HTTPS（標準ポート443）に限定し、認証ヘッダー等は送信しません。私設/ループバック/リンクローカル等のアドレスや、そのアドレスへの転送は拒否します。URLの名前解決結果に固定して接続します。転送は最大3回、取得・変換全体は1要求45秒まで。変換はサービス内で同時1件、使用中は429です。画像が1枚でも失敗した場合は、その要求を成功扱いにしません。再送では同じ通知IDを使ってください。保存済みの同じ通知の再送は、画像URLが期限切れになっても再取得せず同じ投稿を返します。

## 応答と再送

|HTTP|意味|送信側の扱い|
|---|---|---|
|201|新しい投稿を保存|完了|
|200|同じID・同じ内容の保存済み投稿|完了|
|400|本文/JSON/文字数/画像数/IDなどの入力不正|内容を直して新しいイベントとして送信。元の要求を無限再送しない|
|404|削除済み/誤った/再発行前のURL、未対応クエリ|設定URLを確認|
|405|POST以外|メソッドを修正|
|409|同じ通知IDで本文/画像URLが異なる|通知IDの使い回しを修正|
|410|受信オフ、掲示板停止、期限切れ|開発者が設定と掲示板を確認|
|413|受信データが大きすぎる|容量を減らす|
|422|画像URL/画像形式/取得/変換の問題|公開範囲・形式・容量等を確認|
|429|回数制限または画像変換中|Retry-After以降に同じ通知IDで再送|
|503|メンテナンス/起動準備/一時障害等|Retry-Afterがあれば従い、同じ通知IDで間隔を空けて再送|

通常応答は `{"id":"投稿UUID","replayed":false}`。再送時は同じidと`replayed:true`。入力エラーは `{"code":"content_too_long","error":"..."}`等です。入口で返す413等はJSONとは限らないので、送信側は先にHTTP状態を調べてください。

1URLにつき1分間60投稿、接続元につき1分間120要求。429のRetry-Afterは60秒です。メンテナンス503は意図した受付停止で、開発者のエラー掲示板へ報告しません。HTTP成功は保存完了を示し、通知端末への配信完了は示しません。通知は通常のピン留め・通知オン/オフ・閲覧権限に従います。

ネットワーク切断などで結果不明なら、1秒/数秒単位の連続再送をせず、60秒程度から待ち時間を増やしてください。同じID・内容で再送すると、保存が済んでいた場合も重複コメントを作りません。

## Node.js送信例

```javascript
const url = process.env.GCAL_BOARD_WEBHOOK_URL;
const notificationId = savedEvent.id; // 送信前に永続化済みのイベントID
const response = await fetch(url, {
  method: 'POST',
  headers: {'Content-Type':'application/json', 'Idempotency-Key':notificationId},
  body: JSON.stringify({content:savedEvent.content}),
  redirect: 'error', signal: AbortSignal.timeout(60000)
});
if (!response.ok) {
  // URL自体や全リクエストをログへ出さない。状態とイベントIDだけ記録。
  throw new Error(`掲示板通知 HTTP ${response.status} / ${notificationId}`);
}
```

## Apps Script送信例

```javascript
const url = PropertiesService.getScriptProperties().getProperty('GCAL_BOARD_WEBHOOK_URL');
const response = UrlFetchApp.fetch(url, {
  method: 'post', contentType: 'application/json',
  headers: {'Idempotency-Key': savedEvent.id},
  payload: JSON.stringify({content: savedEvent.content}),
  followRedirects: false, muteHttpExceptions: true
});
const status = response.getResponseCode();
// 2xxなら完了。429/503/通信失敗なら予約を変更せず通知だけを再送。
```

gcal-bookingの実装はCode.gs＋BoardWebhook.gs、設定手順は同リポジトリREADMEにあります。予約処理中は非公開シートへ送信待ち記録を書き、ユーザーが有効化する1分間隔のApps Scriptトリガーが予約ロックの外で配信します。既存Discord送信は維持します。

参考一次資料：[Discord Execute Webhook](https://docs.discord.com/developers/resources/webhook#execute-webhook)、[Apps Script UrlFetchApp](https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app)、[Apps Scriptの割当制限](https://developers.google.com/apps-script/guides/services/quotas)。Google側の実行時間・回数・トリガー等の制限は利用アカウントに従います。
