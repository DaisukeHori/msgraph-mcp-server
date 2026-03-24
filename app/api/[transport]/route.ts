/**
 * MCP エンドポイント
 *
 * /api/mcp  → Streamable HTTP (メイン)
 * /api/sse  → SSE (後方互換)
 *
 * AUTH_MODE による認証分岐:
 *
 *  "token" (Vercel デフォルト):
 *    本人の Microsoft Graph アクセストークンを Bearer Token で渡す。
 *    Claude.ai Web では ?token=<TOKEN> クエリでも可。
 *    /me/ で本人のデータにアクセス。
 *
 *  "delegated" (ローカル stdio 用):
 *    MSAL Device Code Flow で認証。Vercel では使用不可。
 *
 *  "client_credentials" (管理者/自動化):
 *    Azure AD クライアント資格情報。/me/ 使用不可。
 */

import { createMcpHandler } from "mcp-handler";
import { registerAllTools } from "@/lib/mcp/server";
import { authStorage, getAuthMode } from "@/lib/msgraph/auth-context";

const mcpHandler = createMcpHandler(
  (server) => { registerAllTools(server); },
  {},
  { basePath: "/api", maxDuration: 60, verboseLogs: process.env.NODE_ENV === "development" }
);

function extractBearerToken(request: Request): string | undefined {
  const h = request.headers.get("authorization") || "";
  return h.match(/^Bearer\s+(.+)$/i)?.[1];
}

function extractQueryToken(request: Request, param: string): string | undefined {
  try { return new URL(request.url).searchParams.get(param) || undefined; }
  catch { return undefined; }
}

async function handler(request: Request): Promise<Response> {
  const mode = getAuthMode();
  const bearerToken = extractBearerToken(request);

  if (mode === "token" || mode === "delegated") {
    // Bearer Token or ?token= で Graph アクセストークンを取得
    const graphToken = bearerToken || extractQueryToken(request, "token");
    if (graphToken) {
      return authStorage.run({ graphAccessToken: graphToken }, () => mcpHandler(request));
    }
    // トークンなしの場合はそのまま実行（エラーメッセージが返る）
    return mcpHandler(request);
  }

  if (mode === "client_credentials") {
    return mcpHandler(request);
  }

  return mcpHandler(request);
}

export { handler as GET, handler as POST };
