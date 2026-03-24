import { describe, it, expect, vi, beforeEach } from "vitest";
import { generateAuthUrl, exchangeCodeForTokens, refreshAccessToken, getUserProfile } from "@/lib/msgraph/oauth";

describe("lib/msgraph/oauth", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  // ── generateAuthUrl ──
  describe("generateAuthUrl", () => {
    it("O01: 認可 URL を生成する", () => {
      const result = generateAuthUrl("https://my-app.vercel.app");
      expect(result.url).toContain("login.microsoftonline.com");
      expect(result.url).toContain("oauth2/v2.0/authorize");
      expect(result.state).toBeDefined();
      expect(result.codeVerifier).toBeDefined();
    });

    it("O02: URL にクライアント ID が含まれる", () => {
      const result = generateAuthUrl("https://my-app.vercel.app");
      expect(result.url).toContain("client_id=test-client-id");
    });

    it("O03: URL にリダイレクト URI が含まれる", () => {
      const result = generateAuthUrl("https://my-app.vercel.app");
      expect(result.url).toContain(encodeURIComponent("https://my-app.vercel.app/api/auth/callback"));
    });

    it("O04: URL に PKCE code_challenge が含まれる", () => {
      const result = generateAuthUrl("https://my-app.vercel.app");
      expect(result.url).toContain("code_challenge=");
      expect(result.url).toContain("code_challenge_method=S256");
    });

    it("O05: state は毎回異なる", () => {
      const a = generateAuthUrl("https://test.app");
      const b = generateAuthUrl("https://test.app");
      expect(a.state).not.toBe(b.state);
    });

    it("O06: codeVerifier は毎回異なる", () => {
      const a = generateAuthUrl("https://test.app");
      const b = generateAuthUrl("https://test.app");
      expect(a.codeVerifier).not.toBe(b.codeVerifier);
    });

    it("O07: URL に offline_access スコープが含まれる", () => {
      const result = generateAuthUrl("https://test.app");
      expect(result.url).toContain("offline_access");
    });

    it("O08: URL に prompt=consent が含まれる", () => {
      const result = generateAuthUrl("https://test.app");
      expect(result.url).toContain("prompt=consent");
    });

    it("O09: テナント ID が URL に含まれる", () => {
      const result = generateAuthUrl("https://test.app");
      expect(result.url).toContain("test-tenant-id");
    });

    it("O10: CLIENT_ID 未設定でエラー", () => {
      const original = process.env.MICROSOFT_CLIENT_ID;
      delete process.env.MICROSOFT_CLIENT_ID;
      expect(() => generateAuthUrl("https://test.app")).toThrow("OAuth 設定エラー");
      process.env.MICROSOFT_CLIENT_ID = original;
    });
  });

  // ── exchangeCodeForTokens ──
  describe("exchangeCodeForTokens", () => {
    it("O11: 成功レスポンスからトークンを返す", async () => {
      vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
        ok: true,
        json: () => Promise.resolve({
          access_token: "at_test",
          refresh_token: "rt_test",
          expires_in: 3600,
        }),
      }));

      const result = await exchangeCodeForTokens("code123", "verifier123", "https://test.app");
      expect(result.accessToken).toBe("at_test");
      expect(result.refreshToken).toBe("rt_test");
      expect(result.expiresIn).toBe(3600);
    });

    it("O12: エラーレスポンスで例外", async () => {
      vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
        ok: false,
        status: 400,
        statusText: "Bad Request",
        json: () => Promise.resolve({ error_description: "Invalid code" }),
      }));

      await expect(exchangeCodeForTokens("bad", "bad", "https://test.app")).rejects.toThrow("トークン交換失敗");
    });
  });

  // ── refreshAccessToken ──
  describe("refreshAccessToken", () => {
    it("O13: refresh_token で新しいトークンを取得する", async () => {
      vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
        ok: true,
        json: () => Promise.resolve({
          access_token: "new_at",
          refresh_token: "new_rt",
          expires_in: 3600,
        }),
      }));

      const result = await refreshAccessToken("old_rt");
      expect(result.accessToken).toBe("new_at");
      expect(result.refreshToken).toBe("new_rt");
    });

    it("O14: 失敗時にエラー", async () => {
      vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
        ok: false,
        status: 401,
        statusText: "Unauthorized",
        json: () => Promise.resolve({ error_description: "Token expired" }),
      }));

      await expect(refreshAccessToken("expired_rt")).rejects.toThrow("トークン更新失敗");
    });
  });

  // ── getUserProfile ──
  describe("getUserProfile", () => {
    it("O15: プロフィール情報を取得する", async () => {
      vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
        ok: true,
        json: () => Promise.resolve({
          displayName: "堀大介",
          mail: "hori@revol.co.jp",
          userPrincipalName: "hori@revol.co.jp",
        }),
      }));

      const result = await getUserProfile("at_test");
      expect(result.displayName).toBe("堀大介");
      expect(result.mail).toBe("hori@revol.co.jp");
    });
  });
});
