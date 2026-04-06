# sync.sanjo.dev

# Docker起動手順と開発ルール

この README では、ローカル開発の準備としてアプリを起動し、動作確認するところまでを説明します。

このプロジェクトでの Docker は、主に **開発環境をそろえてローカルで起動するため** に使います。
本番運用の説明ではなく、まずは開発用の起動手順として読んでください。

最初にやることは 3 つだけです。

1. `.env` ファイルを作る
2. Docker でアプリを起動する
3. `/health` にアクセスして成功を確認する

## 1. 環境変数ファイルを作成

このプロジェクトでは環境変数を `.env` ファイルで管理します。

リポジトリには `.env.example` が用意されています。

まずは `.env.example` をコピーして `.env` を作ってください。

### Mac / Linux

```bash
cp .env.example .env
```

### Windows (PowerShell)

```powershell
Copy-Item .env.example .env
```

環境変数はこのファイルで設定してください。

`.env` は **Git管理されません。**

最低限の見本は以下です。

```dotenv
PORT=3000
NGROK_AUTHTOKEN=
SUPABASE_URL=
SUPABASE_SERVICE_ROLE_KEY=
LINE_CHANNEL_ACCESS_TOKEN=
LINE_CHANNEL_SECRET=
VERTEX_PROJECT_ID=
```

補足:

- 詳細な環境変数一覧は `[.env.example](./.env.example)` を参照してください。
- `docker-compose.yml` は `env_file: .env` を読むため、**Docker 起動前に `.env` 作成が必要です。**
- `NGROK_AUTHTOKEN` は ngrok を使うときだけ必要です。使わないなら空のままで問題ありません。
- `SUPABASE_SERVICE_ROLE_KEY`、`LINE_CHANNEL_ACCESS_TOKEN`、`LINE_CHANNEL_SECRET`、`VERTEX_PROJECT_ID` は対応する機能を使うときに設定します。

---

## 2. Dockerでアプリを起動

このプロジェクトは **Dockerを使って開発用アプリを起動する前提**としています。  
ローカルで `npm install` や `npm run dev` を実行する必要はありません。

まずは次のコマンドを使ってください。

### アプリ起動

```bash
docker compose up --build
```

これは、必要なイメージを作成してからアプリを起動するコマンドです。
初回起動や Dockerfile を変更したあとはこちらを使うのが安全です。

または

```bash
docker compose up -d
```

こちらはバックグラウンドで起動するコマンドです。
画面にログを出しながら確認したい場合は、まず `docker compose up --build` を使ってください。

補足:

- 通常の開発では `app` サービスだけで十分です。
- `ngrok` は外部からローカル環境へアクセスしたいときだけ使う補助ツールです。
- 今回の開発で ngrok を準備していても、実際に使っていなければそのまま未使用で問題ありません。

起動後、以下のURLで動作確認できます。

```text
http://localhost:3000
http://localhost:3000/health
```

初心者向けには、まず `http://localhost:3000/health` だけ確認すれば十分です。

`/health` のレスポンス例:

```json
{
  "status": "ok",
  "service": "sync-sanjo-app",
  "timestamp": "2026-04-06T00:00:00.000Z"
}
```

`status` が `ok` なら起動成功です。

参考として、`/` では以下のようなレスポンスを返します。

```json
{
  "message": "Sync Sanjo API is running"
}
```

### アプリ停止

ターミナルで

```text
Ctrl + C
```

または別ターミナルで

```bash
docker compose down
```

### コンテナ状態確認

```bash
docker compose ps
```

### ログ確認

```bash
docker compose logs -f
```

### トラブルシューティング

Dockerが起動しない場合は、まず Docker Desktop が起動しているか確認してください。

`docker compose up --build` で失敗したら、次を順番に確認してください。

1. Docker Desktop が起動しているか
2. `.env` が作成されているか
3. `docker compose ps` で `app` コンテナが動いているか
4. `docker compose logs -f` でエラーが出ていないか

Dockerを完全リセットする場合:

```bash
docker compose down
docker compose build --no-cache
docker compose up
```

---

## 3. 重要な開発ルール

チーム開発で環境差分や事故を防ぐため、以下のルールを守ってください。

### ① npm install を実行しない

依存関係は **Docker内で自動インストールされます。**

そのため、まずは以下のコマンドをローカルで実行しないでください。

```
npm install
npm run dev
```

起動は必ず Docker を使用します。

```
docker compose up
```

---

### ② node_modules は Git管理しない

`node_modules` は Docker volume で管理されています。

```
app_node_modules:/app/node_modules
```

そのため Git に含めません。

---

### ③ .env は Gitにpushしない

`.env` には秘密情報が含まれるため
Gitに push してはいけません。

Git管理されるのは

```
.env.example
```

のみです。

---

### ④ PR前にDockerで動作確認する

Pull Request を出す前に必ず以下で起動確認をしてください。

```bash
docker compose up --build
```

---

### ⑤ コンテナが壊れた場合

次のコマンドで環境をリセットできます。

```bash
docker compose down
docker compose up --build
```
