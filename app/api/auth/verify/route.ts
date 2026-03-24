/**
 * POST /api/auth/verify
 *
 * ADMIN_SECRET を検証し、成功したらセッショントークンを返す。
 * ブルートフォース対策: 5回失敗で15分ロック。
 */

import { NextRequest, NextResponse } from "next/server";
import { timingSafeEqual } from "crypto";
import {
  checkLockout,
  recordFailedAttempt,
  clearLockout,
  createAdminSessionToken,
  saveAdminSession,
} from "@/lib/redis/token-store";

function getClientIp(request: NextRequest): string {
  return (
    request.headers.get("x-forwarded-for")?.split(",")[0]?.trim() ||
    request.headers.get("x-real-ip") ||
    "unknown"
  );
}

export async function POST(request: NextRequest) {
  try {
    const ip = getClientIp(request);

    // ロックアウトチェック
    const lockout = await checkLockout(ip);
    if (lockout.locked) {
      return NextResponse.json(
        { error: `ロックアウト中です。${lockout.remaining}秒後に再試行してください。` },
        { status: 429 }
      );
    }

    const body = await request.json().catch(() => ({})) as { secret?: string };
    const provided = body.secret || "";
    const expected = process.env.ADMIN_SECRET || "";

    if (!expected) {
      return NextResponse.json(
        { error: "ADMIN_SECRET 環境変数が設定されていません。" },
        { status: 500 }
      );
    }

    // タイミングセーフ比較
    let valid = false;
    try {
      const a = Buffer.from(provided, "utf-8");
      const b = Buffer.from(expected, "utf-8");
      if (a.length === b.length) {
        valid = timingSafeEqual(a, b);
      }
    } catch {
      valid = false;
    }

    if (!valid) {
      const attempts = await recordFailedAttempt(ip);
      const remaining = 5 - attempts;
      return NextResponse.json(
        {
          error: "管理パスワードが正しくありません。",
          remaining: remaining > 0 ? remaining : 0,
        },
        { status: 401 }
      );
    }

    // 成功 → ロックアウトカウンターをクリア
    await clearLockout(ip);

    // セッショントークンを生成して Redis に保存
    const sessionToken = createAdminSessionToken();
    await saveAdminSession(sessionToken);

    return NextResponse.json({ sessionToken });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "内部エラー" },
      { status: 500 }
    );
  }
}
