# sync.sanjo.dev

# Docker起動手順と開発ルール

## 1. Dockerでアプリを起動

このプロジェクトは **Dockerでの起動を前提**としています。  
ローカルで `npm install` や `npm run dev` を実行する必要はありません。

### アプリ起動

```bash
docker compose up --build
```

または

```
docker compose up-d
```

起動後、以下のURLで動作確認できます。

```
http://localhost:3000
```

または

```
http://localhost:3000/health
```

以下のようなレスポンスが返れば成功です。

```json
{
  "status": "ok"
}
```

### アプリ停止

ターミナルで

```
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

Dockerが起動しない

まず、Docker Desktop が起動しているか確認してください。

Dockerを完全リセット

```
docker compose down
docker compose build --no-cache
docker compose up
```

---

## 2. 環境変数ファイルを作成

このプロジェクトでは環境変数を `.env` ファイルで管理します。

リポジトリには `.env.example` が用意されています。

### Mac / Linux

```bash
cp .env.example .env
```

### Windows (PowerShell)

```powershell
Copy-Item .env.example .env
```

`.env` は **Git管理されません。**

環境変数はこのファイルで設定してください。

例

```
PORT=3000
NGROK_AUTHTOKEN=
SUPABASE_URL=
SUPABASE_ANON_KEY=
LINE_CHANNEL_ACCESS_TOKEN=
LINE_CHANNEL_SECRET=
```

---

## 3. 重要な開発ルール

チーム開発で環境差分や事故を防ぐため、以下のルールを守ってください。

### ① npm install を実行しない

依存関係は **Docker内で自動インストールされます。**

そのため以下のコマンドは実行しないでください。

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
