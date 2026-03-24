import { describe, it, expect, beforeEach } from "vitest";
import { clearMockRedis } from "../setup";
import { POST as rotateKeyPost } from "@/app/api/auth/rotate-key/route";
import { POST as statusPost } from "@/app/api/auth/status/route";
import { NextRequest } from "next/server";
import {
  createAdminSessionToken,
  saveAdminSession,
  generateMcpApiKey,
  saveRefreshToken,
  recordCronExecution,
  verifyMcpApiKey,
} from "@/lib/redis/token-store";

function makeRequest(body: unknown): NextRequest {
  return new NextRequest("http://localhost/api/auth/test", {
    method: "POST",
    body: JSON.stringify(body),
    headers: { "content-type": "application/json" },
  });
}

describe("POST /api/auth/rotate-key", () => {
  beforeEach(() => {
    clearMockRedis();
  });

  it("R01: 有効なセッションで API キーをローテーション", async () => {
    const session = createAdminSessionToken();
    await saveAdminSession(session);
    const oldKey = await generateMcpApiKey();

    const res = await rotateKeyPost(makeRequest({ session }));
    expect(res.status).toBe(200);
    const data = await res.json();
    expect(data.success).toBe(true);
    expect(data.mcpApiKey).toBeDefined();
    expect(data.mcpApiKey).not.toBe(oldKey);

    // 古いキーは無効
    expect(await verifyMcpApiKey(oldKey)).toBe(false);
    // 新しいキーは有効
    expect(await verifyMcpApiKey(data.mcpApiKey)).toBe(true);
  });

  it("R02: セッションなしで 401", async () => {
    const res = await rotateKeyPost(makeRequest({}));
    expect(res.status).toBe(401);
  });

  it("R03: 無効なセッションで 401", async () => {
    const res = await rotateKeyPost(makeRequest({ session: "invalid" }));
    expect(res.status).toBe(401);
  });

  it("R04: ローテーション後の新キーの形式確認", async () => {
    const session = createAdminSessionToken();
    await saveAdminSession(session);
    await generateMcpApiKey();

    const res = await rotateKeyPost(makeRequest({ session }));
    const data = await res.json();
    expect(data.mcpApiKey.length).toBe(64);
    expect(/^[a-f0-9]+$/.test(data.mcpApiKey)).toBe(true);
  });

  it("R05: 連続ローテーションで毎回異なるキー", async () => {
    const session = createAdminSessionToken();
    await saveAdminSession(session);
    await generateMcpApiKey();

    const keys = new Set<string>();
    for (let i = 0; i < 5; i++) {
      const res = await rotateKeyPost(makeRequest({ session }));
      const data = await res.json();
      keys.add(data.mcpApiKey);
    }
    expect(keys.size).toBe(5);
  });
});

describe("POST /api/auth/status", () => {
  beforeEach(() => {
    clearMockRedis();
  });

  it("S01: 有効なセッションでステータスを返す（認証前）", async () => {
    const session = createAdminSessionToken();
    await saveAdminSession(session);

    const res = await statusPost(makeRequest({ session }));
    expect(res.status).toBe(200);
    const data = await res.json();
    expect(data.authenticated).toBe(false);
    expect(data.user).toBeNull();
    expect(data.mcpApiKeyConfigured).toBe(false);
  });

  it("S02: 認証後のステータス", async () => {
    const session = createAdminSessionToken();
    await saveAdminSession(session);
    await saveRefreshToken("rt_test", "堀大介", "hori@revol.co.jp");
    await generateMcpApiKey();
    await recordCronExecution();

    const res = await statusPost(makeRequest({ session }));
    const data = await res.json();
    expect(data.authenticated).toBe(true);
    expect(data.user.name).toBe("堀大介");
    expect(data.user.email).toBe("hori@revol.co.jp");
    expect(data.mcpApiKeyConfigured).toBe(true);
    expect(data.mcpApiKeyHint).toMatch(/^\*\*\*\*.{4}$/);
    expect(data.lastCronExecution).not.toBeNull();
    expect(data.cronSchedule).toContain("毎日");
  });

  it("S03: セッションなしで 401", async () => {
    const res = await statusPost(makeRequest({}));
    expect(res.status).toBe(401);
  });

  it("S04: 無効なセッションで 401", async () => {
    const res = await statusPost(makeRequest({ session: "expired" }));
    expect(res.status).toBe(401);
  });

  it("S05: API キーのヒントが末尾4文字", async () => {
    const session = createAdminSessionToken();
    await saveAdminSession(session);
    const key = await generateMcpApiKey();

    const res = await statusPost(makeRequest({ session }));
    const data = await res.json();
    expect(data.mcpApiKeyHint).toBe(`****${key.slice(-4)}`);
  });
});
