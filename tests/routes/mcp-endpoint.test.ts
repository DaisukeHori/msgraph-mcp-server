import { describe, it, expect, vi, beforeEach } from "vitest";
import { clearMockRedis } from "../setup";
import { generateMcpApiKey } from "@/lib/redis/token-store";

// MCP handler はモックが必要（mcp-handler は外部ライブラリ）
vi.mock("mcp-handler", () => ({
  createMcpHandler: () => {
    return async (request: Request) => {
      return new Response(JSON.stringify({ jsonrpc: "2.0", result: "ok" }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    };
  },
}));

// ルートハンドラーを動的インポート（モック後）
const { GET, POST } = await import("@/app/api/[transport]/route");

function makeRequest(method: string, apiKey?: string, queryKey?: string): Request {
  let url = "http://localhost/api/mcp";
  if (queryKey) url += `?key=${queryKey}`;

  const headers: Record<string, string> = { "content-type": "application/json" };
  if (apiKey) headers.authorization = `Bearer ${apiKey}`;

  return new Request(url, { method, headers, body: method === "POST" ? "{}" : undefined });
}

describe("MCP Endpoint /api/[transport]", () => {
  beforeEach(() => {
    clearMockRedis();
  });

  it("M01: API キーなしで 401", async () => {
    const res = await POST(makeRequest("POST"));
    expect(res.status).toBe(401);
    const data = await res.json();
    expect(data.error.message).toContain("API キーが必要");
  });

  it("M02: 不正な API キーで 401", async () => {
    await generateMcpApiKey();
    const res = await POST(makeRequest("POST", "wrong-key"));
    expect(res.status).toBe(401);
    const data = await res.json();
    expect(data.error.message).toContain("無効です");
  });

  it("M03: 正しい API キーで 200", async () => {
    const key = await generateMcpApiKey();
    const res = await POST(makeRequest("POST", key));
    expect(res.status).toBe(200);
  });

  it("M04: GET リクエストも API キー検証される", async () => {
    const res = await GET(makeRequest("GET"));
    expect(res.status).toBe(401);
  });

  it("M05: GET + 正しい API キーで 200", async () => {
    const key = await generateMcpApiKey();
    const res = await GET(makeRequest("GET", key));
    expect(res.status).toBe(200);
  });

  it("M06: クエリパラメータ ?key= でも認証できる", async () => {
    const key = await generateMcpApiKey();
    const res = await POST(makeRequest("POST", undefined, key));
    expect(res.status).toBe(200);
  });

  it("M07: ローテーション後の古いキーで 401", async () => {
    const oldKey = await generateMcpApiKey();
    await generateMcpApiKey(); // ローテーション

    const res = await POST(makeRequest("POST", oldKey));
    expect(res.status).toBe(401);
  });

  it("M08: Bearer プレフィックスなしで 401", async () => {
    const key = await generateMcpApiKey();
    const req = new Request("http://localhost/api/mcp", {
      method: "POST",
      headers: { authorization: key, "content-type": "application/json" },
      body: "{}",
    });
    const res = await POST(req);
    expect(res.status).toBe(401);
  });

  it("M09: Bearer の大文字小文字は無視される", async () => {
    const key = await generateMcpApiKey();
    const req = new Request("http://localhost/api/mcp", {
      method: "POST",
      headers: { authorization: `bearer ${key}`, "content-type": "application/json" },
      body: "{}",
    });
    const res = await POST(req);
    expect(res.status).toBe(200);
  });

  it("M10: JSONRPC エラー形式でレスポンスが返る", async () => {
    const res = await POST(makeRequest("POST"));
    const data = await res.json();
    expect(data.jsonrpc).toBe("2.0");
    expect(data.error.code).toBe(-32001);
  });
});
