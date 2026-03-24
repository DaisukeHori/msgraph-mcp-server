/**
 * POST /api/auth/status
 *
 * 認証ステータスを返す（管理セッション必須）
 */

import { NextRequest, NextResponse } from "next/server";
import {
  verifyAdminSession,
  getTokenMetadata,
  getMcpApiKey,
  getLastCronExecution,
  getRefreshToken,
} from "@/lib/redis/token-store";

export async function POST(request: NextRequest) {
  try {
    const body = await request.json().catch(() => ({})) as { session?: string };
    const sessionToken = body.session;

    if (!sessionToken) {
      return NextResponse.json({ error: "セッショントークンが必要です" }, { status: 401 });
    }

    const valid = await verifyAdminSession(sessionToken);
    if (!valid) {
      return NextResponse.json({ error: "セッションが無効または期限切れです" }, { status: 401 });
    }

    const metadata = await getTokenMetadata();
    const hasRefreshToken = !!(await getRefreshToken());
    const mcpApiKey = await getMcpApiKey();
    const lastCron = await getLastCronExecution();

    return NextResponse.json({
      authenticated: hasRefreshToken,
      user: metadata
        ? { name: metadata.userName, email: metadata.userEmail }
        : null,
      tokenUpdatedAt: metadata?.updatedAt || null,
      tokenCreatedAt: metadata?.createdAt || null,
      mcpApiKeyConfigured: !!mcpApiKey,
      mcpApiKeyHint: mcpApiKey ? `****${mcpApiKey.slice(-4)}` : null,
      lastCronExecution: lastCron,
      cronSchedule: "毎日 03:00 UTC",
    });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "内部エラー" },
      { status: 500 }
    );
  }
}
