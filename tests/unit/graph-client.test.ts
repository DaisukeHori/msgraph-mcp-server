import { describe, it, expect, vi, beforeEach } from "vitest";
import { clearMockRedis } from "../setup";
import { saveRefreshToken } from "@/lib/redis/token-store";
import { getGraphTokenFromRedis, clearAccessTokenCache } from "@/lib/msgraph/auth-context";
import {
  graphGet,
  graphPost,
  graphPatch,
  graphDelete,
  truncateResponse,
  handleToolError,
} from "@/lib/msgraph/graph-client";

describe("lib/msgraph/auth-context", () => {
  beforeEach(() => {
    clearMockRedis();
    clearAccessTokenCache();
    vi.restoreAllMocks();
  });

  it("A01: refresh_token 未保存で認証エラー", async () => {
    await expect(getGraphTokenFromRedis()).rejects.toThrow("認証されていません");
  });

  it("A02: refresh_token → access_token 取得成功", async () => {
    await saveRefreshToken("rt_stored");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve({
        access_token: "at_new",
        refresh_token: "rt_new",
        expires_in: 3600,
      }),
    }));

    const token = await getGraphTokenFromRedis();
    expect(token).toBe("at_new");
  });

  it("A03: access_token がキャッシュされる", async () => {
    await saveRefreshToken("rt_stored");
    const mockFetch = vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve({
        access_token: "at_cached",
        refresh_token: "rt_new",
        expires_in: 3600,
      }),
    });
    vi.stubGlobal("fetch", mockFetch);

    await getGraphTokenFromRedis();
    await getGraphTokenFromRedis();
    // fetch は1回だけ（2回目はキャッシュ）
    expect(mockFetch).toHaveBeenCalledTimes(1);
  });

  it("A04: clearAccessTokenCache でキャッシュがクリアされる", async () => {
    await saveRefreshToken("rt_stored");
    const mockFetch = vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve({
        access_token: "at_test",
        refresh_token: "rt_new",
        expires_in: 3600,
      }),
    });
    vi.stubGlobal("fetch", mockFetch);

    await getGraphTokenFromRedis();
    clearAccessTokenCache();
    await getGraphTokenFromRedis();
    expect(mockFetch).toHaveBeenCalledTimes(2);
  });
});

describe("lib/msgraph/graph-client utilities", () => {
  // ── truncateResponse ──
  it("G01: 短いテキストはそのまま返す", () => {
    expect(truncateResponse("hello")).toBe("hello");
  });

  it("G02: 25000文字を超えるとトランケートされる", () => {
    const long = "x".repeat(30000);
    const result = truncateResponse(long);
    expect(result.length).toBeLessThan(30000);
    expect(result).toContain("切り詰められました");
  });

  it("G03: ちょうど25000文字はトランケートされない", () => {
    const exact = "x".repeat(25000);
    expect(truncateResponse(exact)).toBe(exact);
  });

  // ── handleToolError ──
  it("G04: Error オブジェクトからメッセージ抽出", () => {
    const result = handleToolError(new Error("テストエラー"));
    expect(result).toBe("エラー: テストエラー");
  });

  it("G05: 文字列エラー", () => {
    const result = handleToolError("something went wrong");
    expect(result).toBe("エラー: something went wrong");
  });

  it("G06: null エラー", () => {
    const result = handleToolError(null);
    expect(result).toBe("エラー: null");
  });

  it("G07: undefined エラー", () => {
    const result = handleToolError(undefined);
    expect(result).toBe("エラー: undefined");
  });
});

describe("lib/msgraph/graph-client API calls", () => {
  beforeEach(() => {
    clearMockRedis();
    clearAccessTokenCache();
    vi.restoreAllMocks();
  });

  it("G08: graphGet が正しい URL で fetch を呼ぶ", async () => {
    await saveRefreshToken("rt_test");
    // 1st call: token refresh, 2nd call: actual API call
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: () => Promise.resolve({ displayName: "Test User" }),
      });
    vi.stubGlobal("fetch", mockFetch);

    const result = await graphGet<{ displayName: string }>("/me");
    expect(result.displayName).toBe("Test User");
    // 2nd call should be to graph.microsoft.com
    const graphCall = mockFetch.mock.calls[1];
    expect(graphCall[0]).toContain("graph.microsoft.com/v1.0/me");
  });

  it("G09: graphPost が POST メソッドで呼ぶ", async () => {
    await saveRefreshToken("rt_test");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: () => Promise.resolve({ id: "123" }),
      });
    vi.stubGlobal("fetch", mockFetch);

    await graphPost("/me/sendMail", { message: {} });
    const graphCall = mockFetch.mock.calls[1];
    expect(graphCall[1].method).toBe("POST");
  });

  it("G10: 204 レスポンスで undefined を返す", async () => {
    await saveRefreshToken("rt_test");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 204,
        json: () => Promise.reject("No content"),
      });
    vi.stubGlobal("fetch", mockFetch);

    const result = await graphDelete("/me/messages/123");
    expect(result).toBeUndefined();
  });

  it("G11: Graph API エラーで例外", async () => {
    await saveRefreshToken("rt_test");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: false,
        status: 404,
        json: () => Promise.resolve({ error: { code: "ResourceNotFound", message: "Not found" } }),
      });
    vi.stubGlobal("fetch", mockFetch);

    await expect(graphGet("/me/messages/nonexistent")).rejects.toThrow("ResourceNotFound");
  });
});
