/**
 * AES-256-GCM 暗号化/復号
 *
 * refresh_token を Redis に保存する際に使用。
 * 暗号化キー (TOKEN_ENCRYPTION_KEY) は Vercel 環境変数に保持し、
 * Redis には暗号文のみが保存される。
 */

import { randomBytes, createCipheriv, createDecipheriv, createHash } from "crypto";

const ALGORITHM = "aes-256-gcm";
const IV_LENGTH = 12;
const TAG_LENGTH = 16;

function getKey(): Buffer {
  const raw = process.env.TOKEN_ENCRYPTION_KEY;
  if (!raw) {
    throw new Error(
      "TOKEN_ENCRYPTION_KEY 環境変数が設定されていません。\n" +
        "以下のコマンドで生成してください:\n" +
        "  node -e \"console.log(require('crypto').randomBytes(32).toString('hex'))\""
    );
  }
  // 任意長の文字列を SHA-256 で 32 バイトに正規化
  return createHash("sha256").update(raw).digest();
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
