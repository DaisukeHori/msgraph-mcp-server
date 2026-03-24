/**
 * GET /api/auth/login?session=<token>
 *
 * ADMIN_SECRET 認証済みセッションを検証後、
 * Microsoft OAuth 認可 URL にリダイレクト。
 * state と code_verifier を Redis に保存。
 */

import { NextRequest, NextResponse } from "next/server";
import { verifyAdminSession } from "@/lib/redis/token-store";
import { generateAuthUrl } from "@/lib/msgraph/oauth";
import { getRedis } from "@/lib/redis/client";

export async function GET(request: NextRequest) {
  try {
    const sessionToken = request.nextUrl.searchParams.get("session");

    if (!sessionToken) {
      return NextResponse.json({ error: "セッショントークンが必要です" }, { status: 401 });
    }

    const valid = await verifyAdminSession(sessionToken);
    if (!valid) {
      return NextResponse.json(
        { error: "セッションが無効または期限切れです。管理パスワードから再認証してください。" },
        { status: 401 }
      );
    }

    // OAuth 認可 URL を生成
    const baseUrl = `${request.nextUrl.protocol}//${request.nextUrl.host}`;
    const { url, state, codeVerifier } = generateAuthUrl(baseUrl);

    // state → code_verifier + session のマッピングを Redis に保存 (10分有効)
    const redis = getRedis();
    await redis.set(
      `msgraph:oauth_state:${state}`,
      JSON.stringify({ codeVerifier, sessionToken }),
      { ex: 600 }
    );

    return NextResponse.redirect(url);
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "内部エラー" },
      { status: 500 }
    );
  }
}
