/**
 * Upstash Redis クライアント
 *
 * Vercel Marketplace の Upstash 統合により、
 * 環境変数 KV_REST_API_URL / KV_REST_API_TOKEN が自動設定される。
 */

import { Redis } from "@upstash/redis";

let redis: Redis | null = null;

export function getRedis(): Redis {
  if (redis) return redis;

  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;

  if (!url || !token) {
    throw new Error(
      "Redis 環境変数が設定されていません。\n" +
        "Vercel ダッシュボード → Storage → Upstash Redis を追加してください。\n" +
        "KV_REST_API_URL と KV_REST_API_TOKEN が自動設定されます。"
    );
  }

  redis = new Redis({ url, token });
  return redis;
}
