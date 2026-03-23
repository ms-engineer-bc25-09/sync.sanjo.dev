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
