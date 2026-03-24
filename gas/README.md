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

## Supabase画像Webhook運用

Supabase `projects` テーブルの画像案件を GAS に送って、画像解析結果を `画像案件台帳` / `画像内部ログ` に反映できます。

前提:

1. `clasp push` 後に GAS エディタで Web アプリを最新バージョンへ更新する
2. Supabase Database Webhook を `projects` テーブルに設定する
3. Webhook の送信先 URL は GAS の `/exec` URL にする
4. 対象イベントはまず `UPDATE` のみで運用開始し、必要なら `INSERT` も追加する
5. 対象レコードは `flow_type = 'image'` を前提とする

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

動作確認手順:

1. `projects` の対象レコードを `UPDATE` する
2. Supabase SQL Editor で `net._http_response` を確認する
3. GAS 実行履歴で `doPost` が完了していることを確認する
4. スプレッドシートで `画像案件台帳` / `画像内部ログ` の `ID` 列を確認する

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
4. 同じ `projects.id` の再更新でも行数は増えず、既存行が更新される

## Messaging API リッチメニュー

`gas/src/richmenu.js` に Messaging API 用のリッチメニュー定義があります。

使い方:

1. Script Properties に `LINE_RICH_MENU_IMAGE_FILE_ID` を設定
2. `clasp push`
3. GAS エディタで `createAndSetDefaultMessagingApiRichMenu_()` を実行

これで Messaging API 側の default rich menu を作成し、画像アップロードと適用まで行います。

## 注意

`.clasp.json`はGit管理されません（`.gitignore`に含まれています）。
clone後は各自で設定が必要です。
