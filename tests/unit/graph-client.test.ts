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

// ============================================================
// リトライロジック (Workbook API 用 504/503/429) と
// workbook-session-id / Prefer ヘッダーのテスト
// ============================================================

describe("graph-client retry & workbook session", () => {
  beforeEach(() => {
    clearMockRedis();
    clearAccessTokenCache();
    vi.restoreAllMocks();
  });

  it("R01: 504 エラーで自動リトライして成功", async () => {
    await saveRefreshToken("rt_test");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: false,
        status: 504,
        statusText: "Gateway Timeout",
        headers: new Headers(),
        json: () => Promise.resolve({}),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: () => Promise.resolve({ id: "recovered" }),
      });
    vi.stubGlobal("fetch", mockFetch);

    const result = await graphGet<{ id: string }>("/me/drive/items/abc/workbook/tables");
    expect(result.id).toBe("recovered");
    expect(mockFetch).toHaveBeenCalledTimes(3);
  });

  it("R02: 429 (rate limit) で Retry-After ヘッダー値秒数を尊重", async () => {
    await saveRefreshToken("rt_test");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: false,
        status: 429,
        statusText: "Too Many Requests",
        headers: new Headers({ "Retry-After": "1" }),
        json: () => Promise.resolve({}),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: () => Promise.resolve({ ok: true }),
      });
    vi.stubGlobal("fetch", mockFetch);

    const result = await graphGet<{ ok: boolean }>("/me/drive");
    expect(result.ok).toBe(true);
  });

  it("R03: 503 でリトライ上限到達 → エラー", async () => {
    await saveRefreshToken("rt_test");
    const errResponse = {
      ok: false,
      status: 503,
      statusText: "Service Unavailable",
      headers: new Headers(),
      json: () => Promise.resolve({}),
    };
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValue(errResponse);
    vi.stubGlobal("fetch", mockFetch);

    await expect(graphGet("/me/drive")).rejects.toThrow("503");
  }, 15000);

  it("R04: 404 はリトライしない（即エラー）", async () => {
    await saveRefreshToken("rt_test");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: false,
        status: 404,
        statusText: "Not Found",
        headers: new Headers(),
        json: () => Promise.resolve({ error: { code: "NotFound", message: "x" } }),
      });
    vi.stubGlobal("fetch", mockFetch);

    await expect(graphGet("/me/drive/items/zzz")).rejects.toThrow("NotFound");
    expect(mockFetch).toHaveBeenCalledTimes(2);
  });

  it("S01: workbook-session-id ヘッダーが付与される (graphGet)", async () => {
    await saveRefreshToken("rt_test");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: () => Promise.resolve({ value: [] }),
      });
    vi.stubGlobal("fetch", mockFetch);

    await graphGet(
      "/me/drive/items/abc/workbook/tables",
      undefined,
      { workbookSessionId: "session-xyz" }
    );
    const headers = mockFetch.mock.calls[1][1].headers;
    expect(headers["workbook-session-id"]).toBe("session-xyz");
  });

  it("S02: workbook-session-id 無指定ならヘッダー付与なし", async () => {
    await saveRefreshToken("rt_test");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: () => Promise.resolve({ value: [] }),
      });
    vi.stubGlobal("fetch", mockFetch);

    await graphGet("/me/drive/items/abc/workbook/tables");
    const headers = mockFetch.mock.calls[1][1].headers;
    expect(headers["workbook-session-id"]).toBeUndefined();
  });

  it("S03: graphPost にも workbook-session-id 渡せる", async () => {
    await saveRefreshToken("rt_test");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 201,
        json: () => Promise.resolve({ index: 0, values: [[1, 2]] }),
      });
    vi.stubGlobal("fetch", mockFetch);

    await graphPost(
      "/me/drive/items/abc/workbook/tables/T1/rows/add",
      { index: null, values: [[1, 2]] },
      undefined,
      { workbookSessionId: "session-post" }
    );
    const headers = mockFetch.mock.calls[1][1].headers;
    expect(headers["workbook-session-id"]).toBe("session-post");
    expect(mockFetch.mock.calls[1][1].method).toBe("POST");
  });

  it("S04: graphPatch にも workbook-session-id 渡せる", async () => {
    await saveRefreshToken("rt_test");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: () => Promise.resolve({ id: "ws1", name: "Sheet1" }),
      });
    vi.stubGlobal("fetch", mockFetch);

    await graphPatch(
      "/me/drive/items/abc/workbook/worksheets/Sheet1",
      { name: "Renamed" },
      { workbookSessionId: "session-patch" }
    );
    const headers = mockFetch.mock.calls[1][1].headers;
    expect(headers["workbook-session-id"]).toBe("session-patch");
    expect(mockFetch.mock.calls[1][1].method).toBe("PATCH");
  });

  it("S05: graphDelete にも workbook-session-id 渡せる", async () => {
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

    await graphDelete(
      "/me/drive/items/abc/workbook/tables/T1",
      { workbookSessionId: "session-del" }
    );
    const headers = mockFetch.mock.calls[1][1].headers;
    expect(headers["workbook-session-id"]).toBe("session-del");
    expect(mockFetch.mock.calls[1][1].method).toBe("DELETE");
  });

  it("P01: Prefer ヘッダーが付与される (respond-async 用)", async () => {
    await saveRefreshToken("rt_test");
    const mockFetch = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ access_token: "at_test", refresh_token: "rt_new", expires_in: 3600 }),
      })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        json: () => Promise.resolve({ id: "row-1" }),
      });
    vi.stubGlobal("fetch", mockFetch);

    await graphPost(
      "/me/drive/items/abc/workbook/tables/T1/rows",
      { values: [[1, 2]] },
      undefined,
      { prefer: "respond-async" }
    );
    const headers = mockFetch.mock.calls[1][1].headers;
    expect(headers["Prefer"]).toBe("respond-async");
  });
});
