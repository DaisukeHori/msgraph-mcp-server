/**
 * AES-256-GCM 暗号化/復号
 *
 * 暗号化キーは MICROSOFT_CLIENT_SECRET + MICROSOFT_TENANT_ID から
 * HKDF で自動導出される。ユーザーが別途キーを設定する必要はない。
 *
 * オプションで TOKEN_ENCRYPTION_KEY を明示指定すれば、そちらが優先される。
 */

import { randomBytes, createCipheriv, createDecipheriv, hkdfSync } from "crypto";

const ALGORITHM = "aes-256-gcm";
const IV_LENGTH = 12;

function getKey(): Buffer {
  // 優先: TOKEN_ENCRYPTION_KEY が明示設定されていればそれを使う
  const explicit = process.env.TOKEN_ENCRYPTION_KEY;
  if (explicit) {
    return Buffer.from(hkdfSync("sha256", explicit, "msgraph-mcp", "encryption-key", 32));
  }

  // 自動導出: MICROSOFT_CLIENT_SECRET + TENANT_ID → HKDF → 32バイト鍵
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
  const tenantId = process.env.MICROSOFT_TENANT_ID;

  if (!clientSecret || !tenantId) {
    throw new Error(
      "暗号化キーを導出できません。\n" +
        "MICROSOFT_CLIENT_SECRET と MICROSOFT_TENANT_ID を設定してください。"
    );
  }

  return Buffer.from(
    hkdfSync("sha256", clientSecret, tenantId, "msgraph-mcp-token-encryption", 32)
  );
}

/**
 * 平文を AES-256-GCM で暗号化する
 * 返り値: hex(iv) + ":" + hex(tag) + ":" + hex(ciphertext)
 */
export function encrypt(plaintext: string): string {
  const key = getKey();
  const iv = randomBytes(IV_LENGTH);
  const cipher = createCipheriv(ALGORITHM, key, iv);

  let encrypted = cipher.update(plaintext, "utf8", "hex");
  encrypted += cipher.final("hex");
  const tag = cipher.getAuthTag();

  return `${iv.toString("hex")}:${tag.toString("hex")}:${encrypted}`;
}

/**
 * AES-256-GCM で暗号化された文字列を復号する
 */
export function decrypt(ciphertext: string): string {
  const key = getKey();
  const parts = ciphertext.split(":");
  if (parts.length !== 3) {
    throw new Error("無効な暗号文フォーマット");
  }

  const iv = Buffer.from(parts[0], "hex");
  const tag = Buffer.from(parts[1], "hex");
  const encrypted = parts[2];

  const decipher = createDecipheriv(ALGORITHM, key, iv);
  decipher.setAuthTag(tag);

  let decrypted = decipher.update(encrypted, "hex", "utf8");
  decrypted += decipher.final("utf8");

  return decrypted;
}
