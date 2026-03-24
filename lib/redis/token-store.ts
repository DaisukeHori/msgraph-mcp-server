/**
 * トークンストア
 *
 * Redis に保存されるデータ:
 *   msgraph:refresh_token    - AES-256-GCM 暗号化された refresh_token
 *   msgraph:mcp_api_key      - MCP クライアント用 API キー (平文、64文字 hex)
 *   msgraph:token_metadata   - メタ情報 (ユーザー名、最終更新日等)
 *   msgraph:lockout:{ip}     - ブルートフォース対策ロックアウト
 *   msgraph:last_cron        - 最終 Cron 実行日時
 */

import { randomBytes, timingSafeEqual } from "crypto";
import { getRedis } from "./client";
import { encrypt, decrypt } from "@/lib/crypto";

const KEY_PREFIX = "msgraph";

// ── Refresh Token ──

export async function saveRefreshToken(
  refreshToken: string,
  userName?: string,
  userEmail?: string
): Promise<void> {
  const redis = getRedis();
  const encrypted = encrypt(refreshToken);

  await redis.set(`${KEY_PREFIX}:refresh_token`, encrypted);
  await redis.set(`${KEY_PREFIX}:token_metadata`, JSON.stringify({
    userName: userName || "unknown",
    userEmail: userEmail || "unknown",
    updatedAt: new Date().toISOString(),
    createdAt: (await getTokenMetadata())?.createdAt || new Date().toISOString(),
  }));
}

export async function getRefreshToken(): Promise<string | null> {
  const redis = getRedis();
  const encrypted = await redis.get<string>(`${KEY_PREFIX}:refresh_token`);
  if (!encrypted) return null;
  try {
    return decrypt(encrypted);
  } catch {
    return null;
  }
}

export async function deleteRefreshToken(): Promise<void> {
  const redis = getRedis();
  await redis.del(`${KEY_PREFIX}:refresh_token`);
  await redis.del(`${KEY_PREFIX}:token_metadata`);
}

// ── Token Metadata ──

interface TokenMetadata {
  userName: string;
  userEmail: string;
  updatedAt: string;
  createdAt: string;
}

export async function getTokenMetadata(): Promise<TokenMetadata | null> {
  const redis = getRedis();
  const raw = await redis.get<string>(`${KEY_PREFIX}:token_metadata`);
  if (!raw) return null;
  try {
    return typeof raw === "string" ? JSON.parse(raw) : raw as unknown as TokenMetadata;
  } catch {
    return null;
  }
}

// ── MCP API Key ──

export async function generateMcpApiKey(): Promise<string> {
  const redis = getRedis();
  const key = randomBytes(32).toString("hex"); // 64 chars
  await redis.set(`${KEY_PREFIX}:mcp_api_key`, key);
  return key;
}

export async function getMcpApiKey(): Promise<string | null> {
  const redis = getRedis();
  return redis.get<string>(`${KEY_PREFIX}:mcp_api_key`);
}

export async function verifyMcpApiKey(provided: string): Promise<boolean> {
  const stored = await getMcpApiKey();
  if (!stored) return false;
  try {
    const a = Buffer.from(provided, "utf-8");
    const b = Buffer.from(stored, "utf-8");
    if (a.length !== b.length) return false;
    return timingSafeEqual(a, b);
  } catch {
    return false;
  }
}

export async function rotateMcpApiKey(): Promise<string> {
  // 古いキーは上書きされて即無効化
  return generateMcpApiKey();
}

// ── Admin Lockout (ブルートフォース対策) ──

const MAX_ATTEMPTS = 5;
const LOCKOUT_SECONDS = 900; // 15分

export async function checkLockout(ip: string): Promise<{ locked: boolean; remaining?: number }> {
  const redis = getRedis();
  const key = `${KEY_PREFIX}:lockout:${ip}`;
  const attempts = await redis.get<number>(key);
  if (attempts !== null && attempts >= MAX_ATTEMPTS) {
    const ttl = await redis.ttl(key);
    return { locked: true, remaining: ttl > 0 ? ttl : LOCKOUT_SECONDS };
  }
  return { locked: false };
}

export async function recordFailedAttempt(ip: string): Promise<number> {
  const redis = getRedis();
  const key = `${KEY_PREFIX}:lockout:${ip}`;
  const current = await redis.incr(key);
  if (current === 1) {
    await redis.expire(key, LOCKOUT_SECONDS);
  }
  return current;
}

export async function clearLockout(ip: string): Promise<void> {
  const redis = getRedis();
  await redis.del(`${KEY_PREFIX}:lockout:${ip}`);
}

// ── Cron ──

export async function recordCronExecution(): Promise<void> {
  const redis = getRedis();
  await redis.set(`${KEY_PREFIX}:last_cron`, new Date().toISOString());
}

export async function getLastCronExecution(): Promise<string | null> {
  const redis = getRedis();
  return redis.get<string>(`${KEY_PREFIX}:last_cron`);
}

// ── Admin Session ──

export function createAdminSessionToken(): string {
  return randomBytes(32).toString("hex");
}

export async function saveAdminSession(token: string): Promise<void> {
  const redis = getRedis();
  // 30分有効
  await redis.set(`${KEY_PREFIX}:admin_session:${token}`, "valid", { ex: 1800 });
}

export async function verifyAdminSession(token: string): Promise<boolean> {
  const redis = getRedis();
  const val = await redis.get<string>(`${KEY_PREFIX}:admin_session:${token}`);
  return val === "valid";
}

export async function deleteAdminSession(token: string): Promise<void> {
  const redis = getRedis();
  await redis.del(`${KEY_PREFIX}:admin_session:${token}`);
}
