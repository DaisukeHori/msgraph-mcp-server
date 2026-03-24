/**
 * リクエストスコープの認証コンテキスト
 *
 * AUTH_MODE による分岐:
 *
 *  "delegated" (デフォルト / ローカル stdio 向け):
 *    MSAL Device Code Flow で本人としてサインイン。
 *    初回のみブラウザで認証 → トークンはファイルにキャッシュ → 自動リフレッシュ。
 *    /me/ エンドポイントで自分のメール・予定・ファイルを操作。
 *
 *  "token" (Vercel / リモート向け):
 *    Bearer Token で Microsoft Graph アクセストークンを直接渡す。
 *    Claude.ai Web 等では ?token=<TOKEN> クエリも可。
 *    /me/ エンドポイントが使え、本人としてアクセス。
 *
 *  "client_credentials" (管理者 / 自動化向け):
 *    Azure AD クライアント資格情報フローでアプリとして認証。
 *    /me/ は使えない → /users/{userId}/ が必要。
 *    全ユーザーのデータにアクセス可能（管理者権限）。
 */

import { AsyncLocalStorage } from "node:async_hooks";

// ── 認証モード定義 ──

export type AuthMode = "delegated" | "token" | "client_credentials";

export function getAuthMode(): AuthMode {
  const mode = process.env.AUTH_MODE?.toLowerCase() || "delegated";
  if (mode === "token") return "token";
  if (mode === "client_credentials") return "client_credentials";
  return "delegated";
}

// ── リクエストスコープ認証コンテキスト (Vercel 用) ──

interface AuthContext {
  graphAccessToken?: string;
}

export const authStorage = new AsyncLocalStorage<AuthContext>();

// ── MSAL Delegated Flow (ローカル stdio 用) ──

let msalAccessToken: string | null = null;
let msalTokenExpiry = 0;

// MSAL は動的インポート (Vercel では不要)
let msalModule: typeof import("@azure/msal-node") | null = null;
let msalInstance: InstanceType<typeof import("@azure/msal-node").PublicClientApplication> | null = null;

const TOKEN_CACHE_PATH = (() => {
  const home = process.env.HOME || process.env.USERPROFILE || ".";
  return `${home}/.msgraph-mcp-token-cache.json`;
})();

const DELEGATED_SCOPES = [
  "User.Read",
  "Mail.Read", "Mail.ReadWrite", "Mail.Send",
  "Calendars.Read", "Calendars.ReadWrite",
  "Team.ReadBasic.All", "Channel.ReadBasic.All",
  "ChannelMessage.Read.All", "ChannelMessage.Send",
  "Chat.Read", "Chat.ReadWrite", "ChatMessage.Read", "ChatMessage.Send",
  "Files.Read.All", "Files.ReadWrite.All",
  "Sites.Read.All", "Sites.ReadWrite.All",
  "User.ReadBasic.All",
  "offline_access",
];

async function getMsalInstance() {
  if (msalInstance) return msalInstance;

  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const tenantId = process.env.MICROSOFT_TENANT_ID || "common";

  if (!clientId) {
    throw new Error(
      "MICROSOFT_CLIENT_ID が設定されていません。\n\n" +
      "Azure Portal でアプリを登録してください:\n" +
      "  1. https://portal.azure.com → Microsoft Entra ID → アプリの登録 → 新規登録\n" +
      "  2. アプリケーション (クライアント) ID をコピー → MICROSOFT_CLIENT_ID に設定\n" +
      "  3. 認証 → 詳細設定 → 「パブリック クライアント フローを許可する」を「はい」に\n" +
      "  4. API のアクセス許可 → Microsoft Graph → 委任されたアクセス許可を追加"
    );
  }

  // 動的インポート
  if (!msalModule) {
    msalModule = await import("@azure/msal-node");
  }

  const fs = await import("fs");

  const cachePlugin = {
    beforeCacheAccess: async (ctx: { tokenCache: { deserialize: (data: string) => void } }) => {
      if (fs.existsSync(TOKEN_CACHE_PATH)) {
        ctx.tokenCache.deserialize(fs.readFileSync(TOKEN_CACHE_PATH, "utf-8"));
      }
    },
    afterCacheAccess: async (ctx: { cacheHasChanged: boolean; tokenCache: { serialize: () => string } }) => {
      if (ctx.cacheHasChanged) {
        fs.writeFileSync(TOKEN_CACHE_PATH, ctx.tokenCache.serialize(), "utf-8");
      }
    },
  };

  msalInstance = new msalModule.PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: { cachePlugin },
  });

  return msalInstance;
}

/**
 * MSAL Device Code Flow で委任トークンを取得する（ローカル用）
 * 初回: ブラウザで認証 → 以降: キャッシュから自動リフレッシュ
 */
async function acquireTokenDelegated(): Promise<string> {
  const now = Date.now();
  if (msalAccessToken && msalTokenExpiry > now + 60_000) {
    return msalAccessToken;
  }

  const pca = await getMsalInstance();
  const cache = pca.getTokenCache();
  const accounts = await cache.getAllAccounts();

  // サイレント取得を試行
  if (accounts.length > 0) {
    try {
      const result = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: DELEGATED_SCOPES.filter(s => s !== "offline_access"),
      });
      msalAccessToken = result.accessToken;
      msalTokenExpiry = result.expiresOn?.getTime() ?? now + 3600_000;
      return msalAccessToken;
    } catch {
      console.error("[auth] サイレント取得失敗、Device Code Flow を開始します...");
    }
  }

  // Device Code Flow
  const result = await pca.acquireTokenByDeviceCode({
    scopes: DELEGATED_SCOPES.filter(s => s !== "offline_access"),
    deviceCodeCallback: (response) => {
      console.error("");
      console.error("═".repeat(60));
      console.error("🔐 Microsoft アカウントへのサインインが必要です");
      console.error("═".repeat(60));
      console.error(response.message);
      console.error("═".repeat(60));
      console.error("");
    },
  });

  if (!result) throw new Error("Device Code 認証に失敗しました");
  msalAccessToken = result.accessToken;
  msalTokenExpiry = result.expiresOn?.getTime() ?? now + 3600_000;
  return msalAccessToken;
}

// ── Client Credentials Flow ──

let ccToken: string | null = null;
let ccTokenExpiry = 0;

async function acquireTokenClientCredentials(): Promise<string> {
  const now = Date.now();
  if (ccToken && ccTokenExpiry > now + 60_000) return ccToken;

  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
  const tenantId = process.env.MICROSOFT_TENANT_ID;

  if (!clientId || !clientSecret || !tenantId) {
    throw new Error(
      "client_credentials モードでは以下の環境変数がすべて必要です:\n" +
      "  MICROSOFT_CLIENT_ID, MICROSOFT_CLIENT_SECRET, MICROSOFT_TENANT_ID\n\n" +
      "⚠️ 注意: このモードでは /me/ エンドポイントは使えません。\n" +
      "本人のメール等を操作するには AUTH_MODE=delegated を使用してください。"
    );
  }

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const response = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: body.toString() }
  );

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(`Azure AD トークン取得失敗: ${(err as { error_description?: string }).error_description || response.statusText}`);
  }

  const data = (await response.json()) as { access_token: string; expires_in: number };
  ccToken = data.access_token;
  ccTokenExpiry = now + data.expires_in * 1000;
  return ccToken;
}

// ── メインエクスポート ──

/**
 * 現在のモードに応じて Microsoft Graph Access Token を取得する
 */
export async function getGraphToken(): Promise<string> {
  const mode = getAuthMode();

  switch (mode) {
    case "delegated":
      return acquireTokenDelegated();

    case "token": {
      const ctx = authStorage.getStore();
      if (ctx?.graphAccessToken) return ctx.graphAccessToken;
      throw new Error(
        "Microsoft Graph Access Token が見つかりません。\n\n" +
        "MCP クライアントで Authorization: Bearer <TOKEN> を設定してください。\n" +
        "トークンの取得方法:\n" +
        "  1. https://developer.microsoft.com/graph/graph-explorer でサインイン\n" +
        "  2. Access Token をコピー\n" +
        "  3. MCP クライアントの Bearer Token に設定\n\n" +
        "または AUTH_MODE=delegated に切り替えて Device Code Flow を使用してください。"
      );
    }

    case "client_credentials":
      return acquireTokenClientCredentials();

    default:
      throw new Error(`不明な AUTH_MODE: ${mode}`);
  }
}

/**
 * 認証キャッシュをクリア（ローカル delegated モード用）
 */
export async function clearAuthCache(): Promise<void> {
  const fs = await import("fs");
  if (fs.existsSync(TOKEN_CACHE_PATH)) {
    fs.unlinkSync(TOKEN_CACHE_PATH);
  }
  msalAccessToken = null;
  msalTokenExpiry = 0;
  msalInstance = null;
  console.error("[auth] 認証キャッシュをクリアしました");
}
