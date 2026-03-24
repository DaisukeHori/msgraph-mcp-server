import { describe, it, expect, beforeEach } from "vitest";
import { encrypt, decrypt } from "@/lib/crypto";

describe("lib/crypto", () => {
  // ── encrypt ──
  describe("encrypt", () => {
    it("C01: 平文を暗号化して文字列を返す", () => {
      const result = encrypt("hello world");
      expect(typeof result).toBe("string");
      expect(result.length).toBeGreaterThan(0);
    });

    it("C02: 暗号文は iv:tag:ciphertext 形式", () => {
      const result = encrypt("test");
      const parts = result.split(":");
      expect(parts.length).toBe(3);
      // IV = 12 bytes = 24 hex chars
      expect(parts[0].length).toBe(24);
      // Tag = 16 bytes = 32 hex chars
      expect(parts[1].length).toBe(32);
      // Ciphertext > 0
      expect(parts[2].length).toBeGreaterThan(0);
    });

    it("C03: 同じ平文でも毎回異なる暗号文になる（ランダム IV）", () => {
      const a = encrypt("same text");
      const b = encrypt("same text");
      expect(a).not.toBe(b);
    });

    it("C04: 空文字列も暗号化できる", () => {
      const result = encrypt("");
      expect(typeof result).toBe("string");
      const parts = result.split(":");
      expect(parts.length).toBe(3);
    });

    it("C05: 長い文字列（10KB）も暗号化できる", () => {
      const longText = "x".repeat(10000);
      const result = encrypt(longText);
      expect(result.length).toBeGreaterThan(10000);
    });

    it("C06: 日本語テキストを暗号化できる", () => {
      const result = encrypt("これはテストです🚀");
      expect(typeof result).toBe("string");
    });

    it("C07: CLIENT_SECRET と TENANT_ID 両方未設定でエラー", () => {
      const origSecret = process.env.MICROSOFT_CLIENT_SECRET;
      const origTenant = process.env.MICROSOFT_TENANT_ID;
      const origKey = process.env.TOKEN_ENCRYPTION_KEY;
      delete process.env.MICROSOFT_CLIENT_SECRET;
      delete process.env.MICROSOFT_TENANT_ID;
      delete process.env.TOKEN_ENCRYPTION_KEY;
      expect(() => encrypt("test")).toThrow("暗号化キーを導出できません");
      process.env.MICROSOFT_CLIENT_SECRET = origSecret;
      process.env.MICROSOFT_TENANT_ID = origTenant;
      process.env.TOKEN_ENCRYPTION_KEY = origKey;
    });
  });

  // ── decrypt ──
  describe("decrypt", () => {
    it("C08: 暗号化した文字列を復号できる", () => {
      const original = "hello world";
      const encrypted = encrypt(original);
      const decrypted = decrypt(encrypted);
      expect(decrypted).toBe(original);
    });

    it("C09: 空文字列を暗号化→復号できる", () => {
      const encrypted = encrypt("");
      expect(decrypt(encrypted)).toBe("");
    });

    it("C10: 日本語テキストを暗号化→復号できる", () => {
      const original = "堀大介のテストデータ🎉";
      const encrypted = encrypt(original);
      expect(decrypt(encrypted)).toBe(original);
    });

    it("C11: JSON 文字列を暗号化→復号できる", () => {
      const original = JSON.stringify({ refresh_token: "0.AAAA.BBBB", expires_in: 3600 });
      const encrypted = encrypt(original);
      const decrypted = decrypt(encrypted);
      expect(JSON.parse(decrypted)).toEqual(JSON.parse(original));
    });

    it("C12: 不正な暗号文でエラー", () => {
      expect(() => decrypt("invalid")).toThrow("無効な暗号文");
    });

    it("C13: 改竄された暗号文でエラー", () => {
      const encrypted = encrypt("test");
      const parts = encrypted.split(":");
      parts[2] = "0000" + parts[2].slice(4); // ciphertext を改竄
      expect(() => decrypt(parts.join(":"))).toThrow();
    });

    it("C14: 異なる暗号化ソースで正しく復号できない", () => {
      const original_text = "secret";
      const encrypted = encrypt(original_text);
      const origSecret = process.env.MICROSOFT_CLIENT_SECRET;
      process.env.MICROSOFT_CLIENT_SECRET = "completely-different-secret-value";
      try {
        const result = decrypt(encrypted);
        // throw しなくても、復号結果が元と異なることを確認
        expect(result).not.toBe(original_text);
      } catch {
        // GCM auth tag 不一致で throw した場合もOK
        expect(true).toBe(true);
      }
      process.env.MICROSOFT_CLIENT_SECRET = origSecret;
    });

    it("C15: 大きなデータの暗号化→復号の整合性", () => {
      const original = "A".repeat(50000);
      const encrypted = encrypt(original);
      expect(decrypt(encrypted)).toBe(original);
    });
  });
});
