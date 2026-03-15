// sync-sanjo/app/src/routes/health.ts
// 役割:
// - コンテナが元気かどうかを確認するためのAPI
// - docker compose の healthcheck から呼ばれる

import { Router, Request, Response } from "express";

export const healthRouter = Router();

healthRouter.get("/", (_req: Request, res: Response) => {
  res.status(200).json({
    status: "ok",
    service: "sync-sanjo-app",
    timestamp: new Date().toISOString(),
  });
});
