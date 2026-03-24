/**
 * 認証コンテキスト
 *
 * MCP リクエスト時:
 *   1. Bearer Token から MCP API キーを検証
 *   2. Redis から暗号化された refresh_token を取得・復号
 *   3. Microsoft に refresh_token を送って access_token を取得
 *   4. 新しい refresh_token を暗号化して Redis に上書き
 *   5. access_token で Graph API を呼ぶ
 *
 * stdio (ローカル) の場合は従来の MSAL Device Code Flow を使用。
 */

import { getRefreshToken, saveRefreshToken } from "@/lib/redis/token-store";
import { refreshAccessToken } from "./oauth";

// ── access_token キャッシュ (メモリ内、リクエスト間で再利用) ──
let cachedAccessToken: string | null = null;
let cachedTokenExpiry = 0;

/**
 * Redis の refresh_token を使って access_token を取得する (Vercel 用)
 * access_token は短期間メモリにキャッシュ
 */
export async function getGraphTokenFromRedis(): Promise<string> {
  const now = Date.now();

  // キャッシュが有効ならそれを返す（50分以内）
  if (cachedAccessToken && cachedTokenExpiry > now + 60_000) {
    return cachedAccessToken;
  }

  const storedRefreshToken = await getRefreshToken();
  if (!storedRefreshToken) {
    throw new Error(
      "認証されていません。\n" +
        "https://your-app.vercel.app/auth にアクセスして Microsoft アカウントでログインしてください。"
    );
  }

  // refresh_token → 新しい access_token + refresh_token
  const result = await refreshAccessToken(storedRefreshToken);

  // 新しい refresh_token を Redis に保存（90日カウンターリセット）
  await saveRefreshToken(result.refreshToken);

  // access_token をキャッシュ
  cachedAccessToken = result.accessToken;
  cachedTokenExpiry = now + result.expiresIn * 1000;

  return cachedAccessToken;
}

/**
 * キャッシュをクリア（テスト・デバッグ用）
 */
export function clearAccessTokenCache(): void {
  cachedAccessToken = null;
  cachedTokenExpiry = 0;
}
