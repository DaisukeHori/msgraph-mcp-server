/**
 * POST /api/auth/rotate-key
 *
 * MCP API キーをローテーション。
 * 管理セッション必須。古いキーは即無効化される。
 */

import { NextRequest, NextResponse } from "next/server";
import { verifyAdminSession, rotateMcpApiKey } from "@/lib/redis/token-store";

export async function POST(request: NextRequest) {
  try {
    const body = await request.json().catch(() => ({})) as { session?: string };
    const sessionToken = body.session;

    if (!sessionToken) {
      return NextResponse.json({ error: "セッショントークンが必要です" }, { status: 401 });
    }

    const valid = await verifyAdminSession(sessionToken);
    if (!valid) {
      return NextResponse.json(
        { error: "セッションが無効または期限切れです" },
        { status: 401 }
      );
    }

    const newKey = await rotateMcpApiKey();

    return NextResponse.json({
      success: true,
      mcpApiKey: newKey,
      message: "新しい MCP API キーを発行しました。古いキーは即座に無効化されています。",
    });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "内部エラー" },
      { status: 500 }
    );
  }
}
