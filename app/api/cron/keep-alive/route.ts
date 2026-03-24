/**
 * GET /api/cron/keep-alive
 *
 * Vercel Cron Job — 毎日 03:00 UTC に自動実行。
 * refresh_token を使って access_token を取得し、
 * /me にアクセスしてトークンの有効性を確認。
 * 新しい refresh_token を Redis に保存して 90 日カウンターをリセット。
 *
 * CRON_SECRET ヘッダーで認証（Vercel が自動付与）。
 */

import { NextRequest, NextResponse } from "next/server";
import {
  getRefreshToken,
  saveRefreshToken,
  recordCronExecution,
  getTokenMetadata,
} from "@/lib/redis/token-store";
import { refreshAccessToken, getUserProfile } from "@/lib/msgraph/oauth";

export async function GET(request: NextRequest) {
  try {
    // Vercel Cron 認証
    const authHeader = request.headers.get("authorization");
    const cronSecret = process.env.CRON_SECRET;

    if (cronSecret && authHeader !== `Bearer ${cronSecret}`) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    // refresh_token を取得
    const storedRefreshToken = await getRefreshToken();
    if (!storedRefreshToken) {
      console.error("[cron] refresh_token が見つかりません。/auth で認証してください。");
      return NextResponse.json({
        success: false,
        error: "refresh_token が未設定",
      });
    }

    // refresh_token → 新しい access_token + refresh_token
    const tokens = await refreshAccessToken(storedRefreshToken);

    // /me にアクセスしてトークンの有効性を確認
    const profile = await getUserProfile(tokens.accessToken);

    // 新しい refresh_token を保存（90日カウンターリセット）
    await saveRefreshToken(
      tokens.refreshToken,
      profile.displayName,
      profile.mail || profile.userPrincipalName
    );

    // Cron 実行記録
    await recordCronExecution();

    const metadata = await getTokenMetadata();

    console.log(
      `[cron] トークン更新成功: ${profile.displayName} (${profile.mail || profile.userPrincipalName})`
    );

    return NextResponse.json({
      success: true,
      user: profile.displayName,
      email: profile.mail || profile.userPrincipalName,
      tokenUpdatedAt: metadata?.updatedAt,
    });
  } catch (error) {
    console.error("[cron] トークン更新失敗:", error);
    return NextResponse.json({
      success: false,
      error: error instanceof Error ? error.message : "不明なエラー",
    });
  }
}
