// sync-sanjo/app/src/index.ts
// 役割:
// - Expressアプリを作る
// - JSONを受け取れるようにする
// - /health を登録する
// - デフォルトでは3000番ポートで起動する

import express, { Request, Response } from "express";
import { healthRouter } from "./routes/health.js";

const app = express();

// JSON形式のリクエストを読めるようにする
app.use(express.json());

// ルート確認用
app.get("/", (_req: Request, res: Response) => {
  res.json({
    message: "Sync Sanjo API is running",
  });
});

// ヘルスチェック用ルート
app.use("/health", healthRouter);

// まだ未作成のAPIの仮置きメモ
// 今後ここに app.use("/webhook/line", lineWebhookRouter) のように追加していく

const port = Number(process.env.PORT || 3000);

// サーバー起動
app.listen(port, "0.0.0.0", () => {
  console.log(`Sync Sanjo API started on port ${port}`);
});
