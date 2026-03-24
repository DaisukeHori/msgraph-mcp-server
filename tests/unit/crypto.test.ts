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

    it("C07: TOKEN_ENCRYPTION_KEY 未設定でエラー", () => {
      const original = process.env.TOKEN_ENCRYPTION_KEY;
      delete process.env.TOKEN_ENCRYPTION_KEY;
      expect(() => encrypt("test")).toThrow("TOKEN_ENCRYPTION_KEY");
      process.env.TOKEN_ENCRYPTION_KEY = original;
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

    it("C14: 異なる暗号化キーで復号不可", () => {
      const encrypted = encrypt("secret");
      const original = process.env.TOKEN_ENCRYPTION_KEY;
      process.env.TOKEN_ENCRYPTION_KEY = "different-key-for-testing-purposes-here";
      expect(() => decrypt(encrypted)).toThrow();
      process.env.TOKEN_ENCRYPTION_KEY = original;
    });

    it("C15: 大きなデータの暗号化→復号の整合性", () => {
      const original = "A".repeat(50000);
      const encrypted = encrypt(original);
      expect(decrypt(encrypted)).toBe(original);
    });
  });
});
