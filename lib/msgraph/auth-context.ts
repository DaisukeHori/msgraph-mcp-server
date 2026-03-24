/**
 * リクエストスコープの認証コンテキスト
 *
 * AsyncLocalStorage を使って、リクエストごとの認証情報を
 * ツールハンドラーまで伝播する。
 *
 * AUTH_MODE による分岐:
 *  - "graph_token" (デフォルト): Bearer Token に Microsoft Graph アクセストークンを直接渡す
 *  - "client_credentials": Azure AD クライアント資格情報で自動取得
 *  - "api_key": MCP_API_KEY でサーバー認証 + クライアント資格情報
 */

import { AsyncLocalStorage } from "node:async_hooks";

// ── 認証モード定義 ──

export type AuthMode = "graph_token" | "client_credentials" | "api_key";

export function getAuthMode(): AuthMode {
  const mode = process.env.AUTH_MODE?.toLowerCase() || "graph_token";
  if (mode === "client_credentials") return "client_credentials";
  if (mode === "api_key") return "api_key";
  return "graph_token";
}

// ── AsyncLocalStorage ──

interface AuthContext {
  graphAccessToken?: string;
}

export const authStorage = new AsyncLocalStorage<AuthContext>();

// ── クライアント資格情報キャッシュ ──

let cachedToken: string | null = null;
let cachedTokenExpiry = 0;

/**
 * Azure AD クライアント資格情報フローでトークンを取得する
 */
async function acquireTokenByClientCredentials(): Promise<string> {
  const now = Date.now();
  if (cachedToken && cachedTokenExpiry > now + 60_000) {
    return cachedToken;
  }

  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
  const tenantId = process.env.MICROSOFT_TENANT_ID;

  if (!clientId || !clientSecret || !tenantId) {
    throw new Error(
      "Azure AD 設定エラー: MICROSOFT_CLIENT_ID, MICROSOFT_CLIENT_SECRET, " +
        "MICROSOFT_TENANT_ID の環境変数がすべて必要です。" +
        "Azure Portal > Entra ID > アプリの登録 でアプリを作成し、" +
        "クライアントシークレットを発行してください。"
    );
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(
      `Azure AD トークン取得失敗 [${response.status}]: ${
        (err as { error_description?: string }).error_description || response.statusText
      }`
    );
  }

  const data = (await response.json()) as {
    access_token: string;
    expires_in: number;
  };

  cachedToken = data.access_token;
  cachedTokenExpiry = now + data.expires_in * 1000;

  return cachedToken;
}

/**
 * 現在のリクエストスコープから Microsoft Graph Access Token を取得する。
 *
 * AUTH_MODE に応じて取得元が変わる:
 *  - graph_token: Bearer Token のみ（必須）
 *  - client_credentials: Azure AD クライアント資格情報フロー
 *  - api_key: クライアント資格情報フロー（API キー検証済み前提）
 */
export async function getGraphToken(): Promise<string> {
  const mode = getAuthMode();

  if (mode === "client_credentials" || mode === "api_key") {
    return acquireTokenByClientCredentials();
  }

  // graph_token モード: Bearer Token から取得
  const ctx = authStorage.getStore();
  if (ctx?.graphAccessToken) {
    return ctx.graphAccessToken;
  }

  throw new Error(
    "Microsoft Graph Access Token が見つかりません。\n" +
      "MCP クライアントの設定で Authorization: Bearer <your-graph-token> " +
      "ヘッダーを指定してください。\n" +
      "トークンは Azure Portal > Entra ID > アプリの登録 で取得できます。\n" +
      "または AUTH_MODE=client_credentials に切り替えて、" +
      "MICROSOFT_CLIENT_ID / MICROSOFT_CLIENT_SECRET / MICROSOFT_TENANT_ID を設定してください。"
  );
}
