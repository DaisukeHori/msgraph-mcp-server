/**
 * Microsoft OAuth Authorization Code Flow
 *
 * /auth ページからの OAuth フローを処理する。
 * - 認可 URL 生成
 * - 認可コード → トークン交換
 * - refresh_token → access_token 更新
 */

import { randomBytes, createHash } from "crypto";

const SCOPES = [
  "User.Read",
  "Mail.Read", "Mail.ReadWrite", "Mail.Send",
  "Mail.Read.Shared", "Mail.ReadWrite.Shared", "Mail.Send.Shared",
  "Calendars.Read", "Calendars.ReadWrite",
  "Calendars.Read.Shared", "Calendars.ReadWrite.Shared",
  "Team.ReadBasic.All", "Channel.ReadBasic.All",
  "ChannelMessage.Read.All", "ChannelMessage.Send",
  "Chat.Read", "Chat.ReadWrite", "ChatMessage.Read", "ChatMessage.Send",
  "Files.Read.All", "Files.ReadWrite.All",
  "Sites.Read.All", "Sites.ReadWrite.All",
  "User.ReadBasic.All",
  "offline_access",
];

interface OAuthConfig {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  redirectUri: string;
}

function getOAuthConfig(baseUrl: string): OAuthConfig {
  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
  const tenantId = process.env.MICROSOFT_TENANT_ID;

  if (!clientId || !clientSecret || !tenantId) {
    throw new Error(
      "OAuth 設定エラー: MICROSOFT_CLIENT_ID, MICROSOFT_CLIENT_SECRET, " +
        "MICROSOFT_TENANT_ID が必要です。"
    );
  }

  return {
    clientId,
    clientSecret,
    tenantId,
    redirectUri: `${baseUrl}/api/auth/callback`,
  };
}

/**
 * PKCE code_verifier と code_challenge を生成
 */
function generatePKCE(): { verifier: string; challenge: string } {
  const verifier = randomBytes(32).toString("base64url");
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  return { verifier, challenge };
}

/**
 * Microsoft OAuth 認可 URL を生成
 */
export function generateAuthUrl(baseUrl: string): {
  url: string;
  state: string;
  codeVerifier: string;
} {
  const config = getOAuthConfig(baseUrl);
  const state = randomBytes(16).toString("hex");
  const pkce = generatePKCE();

  const params = new URLSearchParams({
    client_id: config.clientId,
    response_type: "code",
    redirect_uri: config.redirectUri,
    response_mode: "query",
    scope: SCOPES.join(" "),
    state,
    code_challenge: pkce.challenge,
    code_challenge_method: "S256",
    prompt: "consent",
  });

  const url = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/authorize?${params.toString()}`;

  return { url, state, codeVerifier: pkce.verifier };
}

/**
 * 認可コードを access_token + refresh_token に交換
 */
export async function exchangeCodeForTokens(
  code: string,
  codeVerifier: string,
  baseUrl: string
): Promise<{
  accessToken: string;
  refreshToken: string;
  expiresIn: number;
}> {
  const config = getOAuthConfig(baseUrl);

  const body = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    code,
    redirect_uri: config.redirectUri,
    grant_type: "authorization_code",
    code_verifier: codeVerifier,
    scope: SCOPES.join(" "),
  });

  const response = await fetch(
    `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    }
  );

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(
      `トークン交換失敗 [${response.status}]: ${
        (err as { error_description?: string }).error_description || response.statusText
      }`
    );
  }

  const data = (await response.json()) as {
    access_token: string;
    refresh_token: string;
    expires_in: number;
  };

  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresIn: data.expires_in,
  };
}

/**
 * refresh_token を使って新しい access_token + refresh_token を取得
 */
export async function refreshAccessToken(refreshToken: string): Promise<{
  accessToken: string;
  refreshToken: string;
  expiresIn: number;
}> {
  const clientId = process.env.MICROSOFT_CLIENT_ID!;
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET!;
  const tenantId = process.env.MICROSOFT_TENANT_ID!;

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    refresh_token: refreshToken,
    grant_type: "refresh_token",
    scope: SCOPES.join(" "),
  });

  const response = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    }
  );

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(
      `トークン更新失敗 [${response.status}]: ${
        (err as { error_description?: string }).error_description || response.statusText
      }`
    );
  }

  const data = (await response.json()) as {
    access_token: string;
    refresh_token: string;
    expires_in: number;
  };

  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresIn: data.expires_in,
  };
}

/**
 * access_token でユーザープロフィールを取得
 */
export async function getUserProfile(accessToken: string): Promise<{
  displayName: string;
  mail: string;
  userPrincipalName: string;
}> {
  const response = await fetch(
    "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName",
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );

  if (!response.ok) {
    throw new Error(`プロフィール取得失敗 [${response.status}]`);
  }

  return response.json() as Promise<{
    displayName: string;
    mail: string;
    userPrincipalName: string;
  }>;
}
