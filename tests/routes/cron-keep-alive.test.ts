import { describe, it, expect, vi, beforeEach } from "vitest";
import { clearMockRedis } from "../setup";
import { GET } from "@/app/api/cron/keep-alive/route";
import { NextRequest } from "next/server";
import { saveRefreshToken, getLastCronExecution, getRefreshToken } from "@/lib/redis/token-store";

function makeRequest(cronSecret?: string): NextRequest {
  const headers: Record<string, string> = {};
  if (cronSecret) {
    headers.authorization = `Bearer ${cronSecret}`;
  }
  return new NextRequest("http://localhost/api/cron/keep-alive", {
    method: "GET",
    headers,
  });
}

describe("GET /api/cron/keep-alive", () => {
  beforeEach(() => {
    clearMockRedis();
    vi.restoreAllMocks();
  });

  it("K01: 正しい CRON_SECRET でトークン更新成功", async () => {
    await saveRefreshToken("rt_old");
    vi.stubGlobal("fetch", vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_new", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ displayName: "堀大介", mail: "hori@revol.co.jp", userPrincipalName: "hori@revol.co.jp" }),
      })
    );

    const res = await GET(makeRequest("test-cron-secret"));
    expect(res.status).toBe(200);
    const data = await res.json();
    expect(data.success).toBe(true);
    expect(data.user).toBe("堀大介");
  });

  it("K02: 不正な CRON_SECRET で 401", async () => {
    const res = await GET(makeRequest("wrong-secret"));
    expect(res.status).toBe(401);
  });

  it("K03: CRON_SECRET なしで 401", async () => {
    const res = await GET(makeRequest());
    expect(res.status).toBe(401);
  });

  it("K04: refresh_token 未保存で失敗レスポンス", async () => {
    const res = await GET(makeRequest("test-cron-secret"));
    const data = await res.json();
    expect(data.success).toBe(false);
    expect(data.error).toContain("refresh_token");
  });

  it("K05: Cron 実行日時が記録される", async () => {
    await saveRefreshToken("rt_old");
    vi.stubGlobal("fetch", vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_new", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ displayName: "Test", mail: "test@test.com", userPrincipalName: "test@test.com" }),
      })
    );

    await GET(makeRequest("test-cron-secret"));
    const lastCron = await getLastCronExecution();
    expect(lastCron).not.toBeNull();
  });

  it("K06: 新しい refresh_token が保存される", async () => {
    await saveRefreshToken("rt_old");
    vi.stubGlobal("fetch", vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_new", refresh_token: "rt_updated", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ displayName: "Test", mail: "test@test.com", userPrincipalName: "test@test.com" }),
      })
    );

    await GET(makeRequest("test-cron-secret"));
    const storedToken = await getRefreshToken();
    expect(storedToken).toBe("rt_updated");
  });

  it("K07: Microsoft API エラーで失敗レスポンス", async () => {
    await saveRefreshToken("rt_expired");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValueOnce({
      ok: false,
      status: 401,
      statusText: "Unauthorized",
      json: () => Promise.resolve({ error_description: "Token expired" }),
    }));

    const res = await GET(makeRequest("test-cron-secret"));
    const data = await res.json();
    expect(data.success).toBe(false);
  });

  it("K08: tokenUpdatedAt がレスポンスに含まれる", async () => {
    await saveRefreshToken("rt_old");
    vi.stubGlobal("fetch", vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_new", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ displayName: "Test", mail: "test@test.com", userPrincipalName: "test@test.com" }),
      })
    );

    const res = await GET(makeRequest("test-cron-secret"));
    const data = await res.json();
    expect(data.tokenUpdatedAt).toBeDefined();
  });
});
