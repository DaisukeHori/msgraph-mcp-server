/**
 * MCP エンドポイント
 *
 * /api/mcp  → Streamable HTTP (メイン)
 * /api/sse  → SSE (後方互換)
 *
 * 認証: MCP API キー (Redis に保存済み)
 *   - Authorization: Bearer <MCP_API_KEY>
 *   - または ?key=<MCP_API_KEY>
 *
 * MCP API キーは /auth ページで発行される。
 * 検証後、Redis の refresh_token → access_token で Graph API を呼ぶ。
 */

import { createMcpHandler } from "mcp-handler";
import { registerAllTools } from "@/lib/mcp/server";
import { verifyMcpApiKey } from "@/lib/redis/token-store";

const mcpHandler = createMcpHandler(
  (server) => {
    registerAllTools(server);
  },
  {},
  {
    basePath: "/api",
    maxDuration: 60,
    verboseLogs: process.env.NODE_ENV === "development",
  }
);

function extractBearerToken(request: Request): string | undefined {
  const h = request.headers.get("authorization") || "";
  return h.match(/^Bearer\s+(.+)$/i)?.[1];
}

function extractQueryParam(request: Request, param: string): string | undefined {
  try {
    return new URL(request.url).searchParams.get(param) || undefined;
  } catch {
    return undefined;
  }
}

async function handler(request: Request): Promise<Response> {
  const apiKey = extractBearerToken(request) || extractQueryParam(request, "key");

  if (!apiKey) {
    return new Response(
      JSON.stringify({
        jsonrpc: "2.0",
        error: {
          code: -32001,
          message:
            "認証エラー: MCP API キーが必要です。\n" +
            "Authorization: Bearer <MCP_API_KEY> ヘッダーを設定してください。\n" +
            "API キーは /auth ページで発行されます。",
        },
      }),
      { status: 401, headers: { "Content-Type": "application/json" } }
    );
  }

  const valid = await verifyMcpApiKey(apiKey);
  if (!valid) {
    return new Response(
      JSON.stringify({
        jsonrpc: "2.0",
        error: {
          code: -32001,
          message:
            "認証エラー: MCP API キーが無効です。\n" +
            "/auth ページで正しいキーを確認するか、新しいキーを発行してください。",
        },
      }),
      { status: 401, headers: { "Content-Type": "application/json" } }
    );
  }

  return mcpHandler(request);
}

export { handler as GET, handler as POST };
