/**
 * GET /api/auth/callback?code=<code>&state=<state>
 *
 * Microsoft OAuth コールバック。
 * 認可コードをトークンに交換し、refresh_token を Redis に保存。
 * MCP API キーを生成して /auth ページにリダイレクト。
 */

import { NextRequest, NextResponse } from "next/server";
import { getRedis } from "@/lib/redis/client";
import { exchangeCodeForTokens, getUserProfile } from "@/lib/msgraph/oauth";
import {
  saveRefreshToken,
  generateMcpApiKey,
  getMcpApiKey,
  verifyAdminSession,
} from "@/lib/redis/token-store";

export async function GET(request: NextRequest) {
  try {
    const code = request.nextUrl.searchParams.get("code");
    const state = request.nextUrl.searchParams.get("state");
    const error = request.nextUrl.searchParams.get("error");
    const errorDescription = request.nextUrl.searchParams.get("error_description");

    // Microsoft がエラーを返した場合
    if (error) {
      const authUrl = new URL("/auth", request.nextUrl.origin);
      authUrl.searchParams.set("error", errorDescription || error);
      return NextResponse.redirect(authUrl);
    }

    if (!code || !state) {
      return NextResponse.json({ error: "code と state が必要です" }, { status: 400 });
    }

    // Redis から state を検証
    const redis = getRedis();
    const stateData = await redis.get<string>(`msgraph:oauth_state:${state}`);
    if (!stateData) {
      const authUrl = new URL("/auth", request.nextUrl.origin);
      authUrl.searchParams.set("error", "OAuth state が無効または期限切れです。再度やり直してください。");
      return NextResponse.redirect(authUrl);
    }

    // state を使い捨て削除
    await redis.del(`msgraph:oauth_state:${state}`);

    const { codeVerifier, sessionToken } = typeof stateData === "string"
      ? JSON.parse(stateData)
      : stateData;

    // セッション再検証
    const sessionValid = await verifyAdminSession(sessionToken);
    if (!sessionValid) {
      const authUrl = new URL("/auth", request.nextUrl.origin);
      authUrl.searchParams.set("error", "管理セッションが期限切れです。管理パスワードから再認証してください。");
      return NextResponse.redirect(authUrl);
    }

    // 認可コード → トークン交換
    const baseUrl = `${request.nextUrl.protocol}//${request.nextUrl.host}`;
    const tokens = await exchangeCodeForTokens(code, codeVerifier, baseUrl);

    // ユーザープロフィールを取得
    const profile = await getUserProfile(tokens.accessToken);

    // refresh_token を暗号化して Redis に保存
    await saveRefreshToken(
      tokens.refreshToken,
      profile.displayName,
      profile.mail || profile.userPrincipalName
    );

    // MCP API キーが未生成なら生成
    const existingKey = await getMcpApiKey();
    const mcpApiKey = existingKey || await generateMcpApiKey();

    // /auth ページに成功パラメータ付きでリダイレクト
    const authUrl = new URL("/auth", request.nextUrl.origin);
    authUrl.searchParams.set("success", "true");
    authUrl.searchParams.set("session", sessionToken);
    // API キーは初回のみ URL に含める（画面に1回だけ表示）
    if (!existingKey) {
      authUrl.searchParams.set("newKey", mcpApiKey);
    }

    return NextResponse.redirect(authUrl);
  } catch (error) {
    const authUrl = new URL("/auth", request.nextUrl.origin);
    authUrl.searchParams.set("error", error instanceof Error ? error.message : "トークン交換に失敗しました");
    return NextResponse.redirect(authUrl);
  }
}
