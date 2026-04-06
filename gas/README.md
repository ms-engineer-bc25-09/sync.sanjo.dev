# GAS（Google Apps Script）

## セットアップ

```bash
npm install -g @google/clasp
clasp login
```

実行するとブラウザが自動で開くので、Googleアカウントでログインしてください。

```bash
cd gas
clasp clone <スクリプトID>
```

スクリプトIDはチームのNotionを確認してください。
参照先Notion：https://www.notion.so/ms-engineer/Clasp-ID-3268f7a0362880039fc8d60106855a7d?source=copy_link

## GASエディタへの反映（push）

ローカルで編集したコードをGASエディタに反映するには以下を実行してください。

```bash
cd gas
clasp push
```

## GASエディタからローカルへの取得（pull）

GASエディタ上で直接編集した場合はローカルに取得できます。

```bash
cd gas
clasp pull
```

## デプロイ

```bash
clasp push
```

pushしただけではデプロイは完了しません。
GASエディタから「デプロイ」→「デプロイを管理」→「新しいバージョンに更新」が必要です。

## Script Properties

この GAS では、機密情報や接続先設定を Script Properties に保存します。

主に使うキー:

1. `GEMINI_API_KEY`
2. `GEMINI_MODEL`
3. `SUPABASE_URL`
4. `SUPABASE_KEY`
5. `LINE_CHANNEL_ACCESS_TOKEN`

用途に応じて使うキー:

1. `LINE_NOTIFY_USER_ID`
2. `LINE_RICH_MENU_IMAGE_FILE_ID`

補足:

1. LINE メッセージ返信や push 通知には `LINE_CHANNEL_ACCESS_TOKEN` が必要です
2. Gemini による解析には `GEMINI_API_KEY` と `GEMINI_MODEL` が必要です
3. Supabase 連携には `SUPABASE_URL` と `SUPABASE_KEY` が必要です
4. Tally 通知を LINE に送る場合は `LINE_NOTIFY_USER_ID` を設定します
5. Messaging API リッチメニュー画像をアップロードする場合は `LINE_RICH_MENU_IMAGE_FILE_ID` を設定します

## LINE画像案件の基本フロー

LINE で図面や FAX 画像を送ると、GAS が AI で必要項目を抽出し、シートと Supabase に反映します。

基本の見せ方:

1. LINE で画像を送信
2. Gemini で顧客名 / 案件名 / 材質 / 数量 / 希望納期などを抽出
3. `案件台帳` に自動登録
4. `画像案件台帳` と `画像内部ログ` も同時に更新
5. Supabase `projects` / `project_items` にも保存
6. 必要時のみ `案件台帳` 上で担当者が修正

補助データ:

1. `画像案件台帳` は画像案件の一覧と案件台帳との対応確認用
2. `画像内部ログ` は再解析や監査用の内部ログ

## Supabase画像Webhook運用

Supabase `projects` テーブルの画像案件を GAS に送ると、画像系の補助ログを同期できます。`ledger_id` が入っているレコードは、`案件台帳` に登録済みの案件として扱います。

前提:

1. `clasp push` 後に GAS エディタで Web アプリを最新バージョンへ更新する
2. Supabase Database Webhook を `projects` テーブルに設定する
3. Webhook の送信先 URL は GAS の `/exec` URL にする
4. `DELETE` 以外のイベントで、画像 URL を含むレコードが対象になる
5. `flow_type = 'image'` のレコードを前提とする
6. `saved_image_url`、`drawing_url`、`original_image_url` のいずれかから画像 URL を取得できる必要がある

Webhook で GAS に渡る想定 payload:

```json
{
  "type": "UPDATE",
  "record": {
    "id": "<projects.id>",
    "flow_type": "image",
    "source_type": "line",
    "project_type": "受付メモ",
    "received_at": "2026-03-24T12:39:59.210553+00:00",
    "updated_at": "2026-03-24T13:11:43.422076+00:00",
    "original_file_name": "sample.jpg",
    "saved_image_url": "https://<project-ref>.supabase.co/storage/v1/object/public/inquiry-files/...",
    "drawing_url": "https://<project-ref>.supabase.co/storage/v1/object/public/inquiry-files/...",
    "processing_status": "reprocessed"
  }
}
```

重要:

1. `saved_image_url` / `drawing_url` にはダミー文字列ではなく、Supabase Storage 上の実在する公開 URL を入れる
2. 同じ `projects.id` は `画像案件台帳` / `画像内部ログ` の `ID` 列に入り、再送時は upsert される
3. `ledger_id` がある場合は、`案件台帳` 登録済みの状態として同期される
4. `DELETE` イベントは同期対象外
5. `ai_extracted_json` があれば再解析せずそれを優先し、なければ画像を再取得して Gemini 解析する

動作確認手順:

1. `projects` の対象レコードを `INSERT` または `UPDATE` する
2. Supabase SQL Editor で `net._http_response` を確認する
3. GAS 実行履歴で `doPost` が完了していることを確認する
4. スプレッドシートで `案件台帳` / `画像案件台帳` / `画像内部ログ` の関連を確認する

確認用 SQL:

```sql
select
  id,
  status_code,
  content,
  created
from net._http_response
order by created desc
limit 5;
```

成功時の目安:

1. `net._http_response.content` が `{"ok":true}`
2. `画像案件台帳` に対象 `ID` が 1 行だけ存在する
3. `画像内部ログ` に対象 `ID` が 1 行だけ存在する
4. `ledger_id` が設定済みなら、画像系シートでも登録済みステータスとして同期される
5. 同じ `projects.id` の再更新でも行数は増えず、既存行が更新される
6. 画像 URL が無い payload や `DELETE` イベントは処理されない

## Messaging API リッチメニュー

`gas/src/richmenu.js` に Messaging API 用のリッチメニュー定義があります。

使い方:

1. Script Properties に `LINE_RICH_MENU_IMAGE_FILE_ID` を設定
2. Script Properties に `LINE_CHANNEL_ACCESS_TOKEN` を設定
3. `clasp push`
4. GAS エディタで `createAndSetDefaultMessagingApiRichMenu_()` を実行

これで Messaging API 側の default rich menu を作成し、画像アップロードと適用まで行います。

## 注意

`.clasp.json`はGit管理されません（`.gitignore`に含まれています）。
clone後は各自で設定が必要です。
