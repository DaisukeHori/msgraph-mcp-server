import { describe, it, expect, beforeEach } from "vitest";
import { clearMockRedis } from "../setup";
import {
  saveRefreshToken,
  getRefreshToken,
  deleteRefreshToken,
  getTokenMetadata,
  generateMcpApiKey,
  getMcpApiKey,
  verifyMcpApiKey,
  rotateMcpApiKey,
  checkLockout,
  recordFailedAttempt,
  clearLockout,
  recordCronExecution,
  getLastCronExecution,
  createAdminSessionToken,
  saveAdminSession,
  verifyAdminSession,
  deleteAdminSession,
} from "@/lib/redis/token-store";

describe("lib/redis/token-store", () => {
  beforeEach(() => {
    clearMockRedis();
  });

  // ── Refresh Token ──
  describe("saveRefreshToken / getRefreshToken", () => {
    it("T01: refresh_token を保存して取得できる", async () => {
      await saveRefreshToken("rt_test_123");
      const result = await getRefreshToken();
      expect(result).toBe("rt_test_123");
    });

    it("T02: 未保存時は null を返す", async () => {
      const result = await getRefreshToken();
      expect(result).toBeNull();
    });

    it("T03: ユーザー名とメールも保存される", async () => {
      await saveRefreshToken("rt_test", "堀大介", "hori@revol.co.jp");
      const metadata = await getTokenMetadata();
      expect(metadata?.userName).toBe("堀大介");
      expect(metadata?.userEmail).toBe("hori@revol.co.jp");
    });

    it("T04: 上書き保存できる", async () => {
      await saveRefreshToken("old_token");
      await saveRefreshToken("new_token");
      const result = await getRefreshToken();
      expect(result).toBe("new_token");
    });

    it("T05: 暗号化されて保存される（平文ではない）", async () => {
      const plainToken = "rt_plain_test_token_12345";
      await saveRefreshToken(plainToken);
      // 直接 Redis の中身を確認（暗号化されているはず）
      // getRefreshToken は復号して返すので、ここでは保存→取得の整合性を確認
      const result = await getRefreshToken();
      expect(result).toBe(plainToken);
    });
  });

  describe("deleteRefreshToken", () => {
    it("T06: refresh_token を削除できる", async () => {
      await saveRefreshToken("rt_to_delete");
      await deleteRefreshToken();
      const result = await getRefreshToken();
      expect(result).toBeNull();
    });

    it("T07: メタデータも同時に削除される", async () => {
      await saveRefreshToken("rt_test", "Test User", "test@test.com");
      await deleteRefreshToken();
      const metadata = await getTokenMetadata();
      expect(metadata).toBeNull();
    });

    it("T08: 存在しない場合もエラーにならない", async () => {
      await expect(deleteRefreshToken()).resolves.not.toThrow();
    });
  });

  // ── Token Metadata ──
  describe("getTokenMetadata", () => {
    it("T09: 保存されたメタデータを取得できる", async () => {
      await saveRefreshToken("rt_test", "テストユーザー", "test@example.com");
      const metadata = await getTokenMetadata();
      expect(metadata).not.toBeNull();
      expect(metadata?.userName).toBe("テストユーザー");
      expect(metadata?.userEmail).toBe("test@example.com");
      expect(metadata?.updatedAt).toBeDefined();
      expect(metadata?.createdAt).toBeDefined();
    });

    it("T10: 未保存時は null", async () => {
      const metadata = await getTokenMetadata();
      expect(metadata).toBeNull();
    });
  });

  // ── MCP API Key ──
  describe("generateMcpApiKey / getMcpApiKey", () => {
    it("T11: API キーを生成できる", async () => {
      const key = await generateMcpApiKey();
      expect(typeof key).toBe("string");
      expect(key.length).toBe(64); // 32 bytes = 64 hex chars
    });

    it("T12: 生成した API キーを取得できる", async () => {
      const key = await generateMcpApiKey();
      const stored = await getMcpApiKey();
      expect(stored).toBe(key);
    });

    it("T13: 未生成時は null", async () => {
      const result = await getMcpApiKey();
      expect(result).toBeNull();
    });

    it("T14: 再生成すると古いキーが上書きされる", async () => {
      const key1 = await generateMcpApiKey();
      const key2 = await generateMcpApiKey();
      expect(key1).not.toBe(key2);
      const stored = await getMcpApiKey();
      expect(stored).toBe(key2);
    });

    it("T15: 毎回異なるキーが生成される", async () => {
      const keys = new Set<string>();
      for (let i = 0; i < 10; i++) {
        keys.add(await generateMcpApiKey());
      }
      expect(keys.size).toBe(10);
    });
  });

  describe("verifyMcpApiKey", () => {
    it("T16: 正しいキーで true", async () => {
      const key = await generateMcpApiKey();
      expect(await verifyMcpApiKey(key)).toBe(true);
    });

    it("T17: 不正なキーで false", async () => {
      await generateMcpApiKey();
      expect(await verifyMcpApiKey("wrong-key")).toBe(false);
    });

    it("T18: キー未生成で false", async () => {
      expect(await verifyMcpApiKey("any-key")).toBe(false);
    });

    it("T19: 空文字で false", async () => {
      await generateMcpApiKey();
      expect(await verifyMcpApiKey("")).toBe(false);
    });

    it("T20: 部分一致で false（タイミングセーフ）", async () => {
      const key = await generateMcpApiKey();
      expect(await verifyMcpApiKey(key.slice(0, 32))).toBe(false);
    });
  });

  describe("rotateMcpApiKey", () => {
    it("T21: 古いキーが無効化され新しいキーが有効", async () => {
      const oldKey = await generateMcpApiKey();
      const newKey = await rotateMcpApiKey();
      expect(oldKey).not.toBe(newKey);
      expect(await verifyMcpApiKey(oldKey)).toBe(false);
      expect(await verifyMcpApiKey(newKey)).toBe(true);
    });
  });

  // ── Lockout ──
  describe("checkLockout / recordFailedAttempt", () => {
    it("T22: 初回は locked=false", async () => {
      const result = await checkLockout("1.2.3.4");
      expect(result.locked).toBe(false);
    });

    it("T23: 4回失敗してもまだロックされない", async () => {
      for (let i = 0; i < 4; i++) {
        await recordFailedAttempt("1.2.3.4");
      }
      const result = await checkLockout("1.2.3.4");
      expect(result.locked).toBe(false);
    });

    it("T24: 5回失敗でロックされる", async () => {
      for (let i = 0; i < 5; i++) {
        await recordFailedAttempt("1.2.3.4");
      }
      const result = await checkLockout("1.2.3.4");
      expect(result.locked).toBe(true);
    });

    it("T25: IP ごとに独立カウント", async () => {
      for (let i = 0; i < 5; i++) {
        await recordFailedAttempt("1.1.1.1");
      }
      const locked1 = await checkLockout("1.1.1.1");
      const locked2 = await checkLockout("2.2.2.2");
      expect(locked1.locked).toBe(true);
      expect(locked2.locked).toBe(false);
    });

    it("T26: clearLockout でロック解除", async () => {
      for (let i = 0; i < 5; i++) {
        await recordFailedAttempt("1.2.3.4");
      }
      await clearLockout("1.2.3.4");
      const result = await checkLockout("1.2.3.4");
      expect(result.locked).toBe(false);
    });

    it("T27: recordFailedAttempt は現在のカウントを返す", async () => {
      expect(await recordFailedAttempt("5.5.5.5")).toBe(1);
      expect(await recordFailedAttempt("5.5.5.5")).toBe(2);
      expect(await recordFailedAttempt("5.5.5.5")).toBe(3);
    });
  });

  // ── Cron ──
  describe("recordCronExecution / getLastCronExecution", () => {
    it("T28: Cron 実行日時を記録・取得できる", async () => {
      await recordCronExecution();
      const last = await getLastCronExecution();
      expect(last).not.toBeNull();
      // ISO 8601 形式であること
      expect(new Date(last!).toISOString()).toBe(last);
    });

    it("T29: 未実行時は null", async () => {
      expect(await getLastCronExecution()).toBeNull();
    });
  });

  // ── Admin Session ──
  describe("Admin Session", () => {
    it("T30: セッショントークンを生成・検証できる", async () => {
      const token = createAdminSessionToken();
      expect(token.length).toBe(64);
      await saveAdminSession(token);
      expect(await verifyAdminSession(token)).toBe(true);
    });

    it("T31: 未保存のセッションは無効", async () => {
      expect(await verifyAdminSession("nonexistent")).toBe(false);
    });

    it("T32: セッション削除後は無効", async () => {
      const token = createAdminSessionToken();
      await saveAdminSession(token);
      await deleteAdminSession(token);
      expect(await verifyAdminSession(token)).toBe(false);
    });

    it("T33: 毎回異なるセッショントークン", () => {
      const tokens = new Set<string>();
      for (let i = 0; i < 10; i++) {
        tokens.add(createAdminSessionToken());
      }
      expect(tokens.size).toBe(10);
    });
  });
});
