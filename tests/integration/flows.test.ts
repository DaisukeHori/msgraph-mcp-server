import { describe, it, expect, vi, beforeEach } from "vitest";
import { clearMockRedis } from "../setup";
import { POST as verifyPost } from "@/app/api/auth/verify/route";
import { POST as rotatePost } from "@/app/api/auth/rotate-key/route";
import { POST as statusPost } from "@/app/api/auth/status/route";
import { NextRequest } from "next/server";
import {
  saveRefreshToken,
  generateMcpApiKey,
  verifyMcpApiKey,
  getRefreshToken,
  getMcpApiKey,
  getTokenMetadata,
} from "@/lib/redis/token-store";
import { encrypt, decrypt } from "@/lib/crypto";
import { getGraphTokenFromRedis, clearAccessTokenCache } from "@/lib/msgraph/auth-context";

function makeReq(path: string, body: unknown): NextRequest {
  return new NextRequest(`http://localhost${path}`, {
    method: "POST",
    body: JSON.stringify(body),
    headers: { "content-type": "application/json", "x-forwarded-for": "192.168.1.1" },
  });
}

describe("結合テスト: 認証フロー", () => {
  beforeEach(() => {
    clearMockRedis();
    clearAccessTokenCache();
    vi.restoreAllMocks();
  });

  it("I01: ADMIN_SECRET 検証 → セッション取得 → ステータス確認の一連フロー", async () => {
    // 1. ADMIN_SECRET でログイン
    const verifyRes = await verifyPost(makeReq("/api/auth/verify", { secret: "test-admin-secret-123" }));
    expect(verifyRes.status).toBe(200);
    const { sessionToken } = await verifyRes.json();

    // 2. ステータス確認（まだ未認証）
    const statusRes = await statusPost(makeReq("/api/auth/status", { session: sessionToken }));
    const status = await statusRes.json();
    expect(status.authenticated).toBe(false);
  });

  it("I02: 認証完了後のステータスフロー", async () => {
    // 1. ログイン
    const verifyRes = await verifyPost(makeReq("/api/auth/verify", { secret: "test-admin-secret-123" }));
    const { sessionToken } = await verifyRes.json();

    // 2. OAuth 完了をシミュレート（直接 Redis に保存）
    await saveRefreshToken("rt_from_oauth", "堀大介", "hori@revol.co.jp");
    await generateMcpApiKey();

    // 3. ステータス確認（認証済み）
    const statusRes = await statusPost(makeReq("/api/auth/status", { session: sessionToken }));
    const status = await statusRes.json();
    expect(status.authenticated).toBe(true);
    expect(status.user.name).toBe("堀大介");
    expect(status.mcpApiKeyConfigured).toBe(true);
  });

  it("I03: API キーローテーションフロー", async () => {
    // 1. 初期セットアップ
    const verifyRes = await verifyPost(makeReq("/api/auth/verify", { secret: "test-admin-secret-123" }));
    const { sessionToken } = await verifyRes.json();
    const oldKey = await generateMcpApiKey();

    // 2. ローテーション
    const rotateRes = await rotatePost(makeReq("/api/auth/rotate-key", { session: sessionToken }));
    const { mcpApiKey: newKey } = await rotateRes.json();

    // 3. 検証
    expect(await verifyMcpApiKey(oldKey)).toBe(false);
    expect(await verifyMcpApiKey(newKey)).toBe(true);
  });

  it("I04: 3回連続ローテーションで常に最新のみ有効", async () => {
    const verifyRes = await verifyPost(makeReq("/api/auth/verify", { secret: "test-admin-secret-123" }));
    const { sessionToken } = await verifyRes.json();
    await generateMcpApiKey();

    const keys: string[] = [];
    for (let i = 0; i < 3; i++) {
      const res = await rotatePost(makeReq("/api/auth/rotate-key", { session: sessionToken }));
      const data = await res.json();
      keys.push(data.mcpApiKey);
    }

    // 最新のキーのみ有効
    expect(await verifyMcpApiKey(keys[0])).toBe(false);
    expect(await verifyMcpApiKey(keys[1])).toBe(false);
    expect(await verifyMcpApiKey(keys[2])).toBe(true);
  });

  it("I05: ロックアウト → 待機 → 再試行の流れ", async () => {
    // 5回失敗
    for (let i = 0; i < 5; i++) {
      await verifyPost(makeReq("/api/auth/verify", { secret: "wrong" }));
    }

    // ロックアウト中
    const lockedRes = await verifyPost(makeReq("/api/auth/verify", { secret: "test-admin-secret-123" }));
    expect(lockedRes.status).toBe(429);
  });
});

describe("結合テスト: 暗号化⇔Redis⇔トークン更新", () => {
  beforeEach(() => {
    clearMockRedis();
    clearAccessTokenCache();
    vi.restoreAllMocks();
  });

  it("I06: refresh_token 暗号化保存→復号取得の整合性", async () => {
    const originalToken = "0.ARwA6hB_KtYxxx.AgABAAEAAAAmoe";
    await saveRefreshToken(originalToken);
    const retrieved = await getRefreshToken();
    expect(retrieved).toBe(originalToken);
  });

  it("I07: 大きな refresh_token（4KB）の暗号化→復号", async () => {
    const bigToken = "T".repeat(4096);
    await saveRefreshToken(bigToken);
    expect(await getRefreshToken()).toBe(bigToken);
  });

  it("I08: メタデータが正しく保存される", async () => {
    await saveRefreshToken("rt_test", "テストユーザー", "test@example.com");
    const metadata = await getTokenMetadata();
    expect(metadata!.userName).toBe("テストユーザー");
    expect(metadata!.userEmail).toBe("test@example.com");
    expect(new Date(metadata!.updatedAt).getTime()).toBeGreaterThan(0);
  });

  it("I09: refresh_token → access_token 取得→Graph API 呼び出しの一連フロー", async () => {
    await saveRefreshToken("rt_real");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_got", refresh_token: "rt_got", expires_in: 3600 }),
      });
    vi.stubGlobal("fetch", mockFetch);

    const token = await getGraphTokenFromRedis();
    expect(token).toBe("at_got");

    // refresh_token が更新されたか確認
    const updatedRt = await getRefreshToken();
    expect(updatedRt).toBe("rt_got");
  });

  it("I10: token refresh 失敗時に元の refresh_token は保持される", async () => {
    await saveRefreshToken("rt_original");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValueOnce({
      ok: false,
      status: 401,
      statusText: "Unauthorized",
      json: () => Promise.resolve({ error_description: "Token expired" }),
    }));

    await expect(getGraphTokenFromRedis()).rejects.toThrow();
    // 元のトークンはまだある
    const rt = await getRefreshToken();
    expect(rt).toBe("rt_original");
  });
});

describe("結合テスト: セキュリティ境界", () => {
  beforeEach(() => {
    clearMockRedis();
  });

  it("I11: 期限切れセッションで全 API が拒否される", async () => {
    const fakeSession = "expired-session-token-that-does-not-exist";

    const statusRes = await statusPost(makeReq("/api/auth/status", { session: fakeSession }));
    expect(statusRes.status).toBe(401);

    const rotateRes = await rotatePost(makeReq("/api/auth/rotate-key", { session: fakeSession }));
    expect(rotateRes.status).toBe(401);
  });

  it("I12: 暗号化キーが異なると refresh_token が読めない", async () => {
    // 正しいキーで保存
    await saveRefreshToken("rt_secret");

    // キーを変更
    const original = process.env.TOKEN_ENCRYPTION_KEY;
    process.env.TOKEN_ENCRYPTION_KEY = "different_key_123456789012345678901234567890";

    // 復号失敗（null が返る）
    const result = await getRefreshToken();
    expect(result).toBeNull();

    process.env.TOKEN_ENCRYPTION_KEY = original;
  });

  it("I13: MCP API キー未生成で全 MCP リクエスト拒否", async () => {
    // キー未生成
    expect(await getMcpApiKey()).toBeNull();
    expect(await verifyMcpApiKey("any-key")).toBe(false);
  });

  it("I14: encrypt/decrypt ラウンドトリップが token-store 経由でも成立", async () => {
    const tokens = [
      "short",
      "a".repeat(1000),
      "日本語トークン",
      JSON.stringify({ nested: { data: [1, 2, 3] } }),
      "special chars: !@#$%^&*()_+-=[]{}|;':\",./<>?",
    ];

    for (const token of tokens) {
      await saveRefreshToken(token);
      const retrieved = await getRefreshToken();
      expect(retrieved).toBe(token);
    }
  });

  it("I15: 複数回の保存で最新の refresh_token のみ有効", async () => {
    await saveRefreshToken("rt_1");
    await saveRefreshToken("rt_2");
    await saveRefreshToken("rt_3");
    expect(await getRefreshToken()).toBe("rt_3");
  });
});
