# sync-sanjo/app/Dockerfile
# 開発用 / 本番用の両方に使えるようにマルチステージにしておく

# =========================
# Base
# =========================
FROM node:22-bookworm-slim AS base

# 作業ディレクトリ
WORKDIR /app

# package.json だけ先にコピーして依存関係をインストール
COPY package.json ./

# npm install を実行
RUN npm install

# =========================
# Development
# =========================
FROM base AS development

# ソースコードをコピー
COPY tsconfig.json ./
COPY src ./src

# 開発サーバーを起動
CMD ["npm", "run", "dev"]

# =========================
# Build
# =========================
FROM base AS build

COPY tsconfig.json ./
COPY src ./src

# TypeScript をビルド
RUN npm run build

# =========================
# Production
# =========================
FROM node:22-bookworm-slim AS production

WORKDIR /app

# 本番に必要なファイルだけコピー
COPY --from=base /app/node_modules ./node_modules
COPY --from=build /app/dist ./dist
COPY package.json ./

ENV NODE_ENV=production
EXPOSE 3000

CMD ["npm", "run", "start"]