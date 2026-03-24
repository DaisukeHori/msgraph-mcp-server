#!/usr/bin/env npx tsx
/**
 * 事前認証スクリプト
 *
 * Claude Desktop で使う前に、ターミナルでこのスクリプトを実行して
 * Microsoft アカウントにサインインしてください。
 *
 * トークンは ~/.msgraph-mcp-token-cache.json にキャッシュされ、
 * Claude Desktop からはこのキャッシュを使って自動的にアクセスします。
 *
 * 使い方:
 *   MICROSOFT_CLIENT_ID=your-id MICROSOFT_TENANT_ID=your-id npx tsx lib/auth-setup.ts
 *
 * または .env ファイルに設定済みなら:
 *   npx tsx lib/auth-setup.ts
 */

import { PublicClientApplication } from "@azure/msal-node";
import * as fs from "fs";
import * as path from "path";

const TOKEN_CACHE_PATH = path.join(
  process.env.HOME || process.env.USERPROFILE || ".",
  ".msgraph-mcp-token-cache.json"
);

const SCOPES = [
  "User.Read",
  "Mail.Read", "Mail.ReadWrite", "Mail.Send",
  "Calendars.Read", "Calendars.ReadWrite",
  "Team.ReadBasic.All", "Channel.ReadBasic.All",
  "ChannelMessage.Read.All", "ChannelMessage.Send",
  "Chat.Read", "Chat.ReadWrite", "ChatMessage.Read", "ChatMessage.Send",
  "Files.Read.All", "Files.ReadWrite.All",
  "Sites.Read.All", "Sites.ReadWrite.All",
  "User.ReadBasic.All",
];

async function main() {
  console.log("");
  console.log("━".repeat(60));
  console.log("  msgraph-mcp-server 事前認証セットアップ");
  console.log("━".repeat(60));
  console.log("");

  // ── 環境変数チェック ──
  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const tenantId = process.env.MICROSOFT_TENANT_ID || "common";

  if (!clientId) {
    console.error("❌ MICROSOFT_CLIENT_ID が設定されていません。");
    console.error("");
    console.error("使い方:");
    console.error("  MICROSOFT_CLIENT_ID=your-id MICROSOFT_TENANT_ID=your-id npx tsx lib/auth-setup.ts");
    console.error("");
    console.error("Azure Portal でアプリを登録してクライアント ID を取得してください:");
    console.error("  https://portal.azure.com → Microsoft Entra ID → アプリの登録");
    process.exit(1);
  }

  console.log(`📋 クライアント ID: ${clientId}`);
  console.log(`📋 テナント ID: ${tenantId}`);
  console.log(`📋 キャッシュ先: ${TOKEN_CACHE_PATH}`);
  console.log("");

  // ── MSAL 初期化 ──
  const pca = new PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: {
      cachePlugin: {
        beforeCacheAccess: async (ctx) => {
          if (fs.existsSync(TOKEN_CACHE_PATH)) {
            ctx.tokenCache.deserialize(fs.readFileSync(TOKEN_CACHE_PATH, "utf-8"));
          }
        },
        afterCacheAccess: async (ctx) => {
          if (ctx.cacheHasChanged) {
            fs.writeFileSync(TOKEN_CACHE_PATH, ctx.tokenCache.serialize(), "utf-8");
          }
        },
      },
    },
  });

  // ── 既存キャッシュの確認 ──
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    console.log(`🔄 既存のキャッシュが見つかりました (${accounts[0].username})`);
    console.log("   サイレント更新を試みます...");
    console.log("");

    try {
      const result = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: SCOPES,
      });

      console.log("✅ サインイン済み！トークンは有効です。");
      console.log("");
      console.log(`   ユーザー: ${result.account?.name || result.account?.username}`);
      console.log(`   メール: ${result.account?.username}`);
      console.log(`   有効期限: ${result.expiresOn?.toLocaleString("ja-JP")}`);
      console.log("");
      console.log("━".repeat(60));
      console.log("  Claude Desktop からそのまま使えます。");
      console.log("━".repeat(60));
      console.log("");
      return;
    } catch {
      console.log("   ⚠️ サイレント更新に失敗しました。再認証します。");
      console.log("");
    }
  }

  // ── Device Code Flow ──
  console.log("🔐 Microsoft アカウントへのサインインを開始します。");
  console.log("   以下の手順に従ってください:");
  console.log("");

  const result = await pca.acquireTokenByDeviceCode({
    scopes: SCOPES,
    deviceCodeCallback: (response) => {
      console.log("┌─────────────────────────────────────────────────┐");
      console.log("│                                                 │");
      console.log("│  1. ブラウザで以下の URL を開く:                │");
      console.log("│     https://microsoft.com/devicelogin           │");
      console.log("│                                                 │");
      console.log(`│  2. コードを入力: ${response.userCode.padEnd(30)}│`);
      console.log("│                                                 │");
      console.log("│  3. Microsoft アカウントでサインイン             │");
      console.log("│                                                 │");
      console.log("│  4. 権限を許可                                  │");
      console.log("│                                                 │");
      console.log("└─────────────────────────────────────────────────┘");
      console.log("");
      console.log("⏳ ブラウザでの認証を待っています...");
    },
  });

  if (!result) {
    console.error("❌ 認証に失敗しました。");
    process.exit(1);
  }

  console.log("");
  console.log("✅ 認証成功！");
  console.log("");
  console.log(`   ユーザー: ${result.account?.name || result.account?.username}`);
  console.log(`   メール: ${result.account?.username}`);
  console.log(`   有効期限: ${result.expiresOn?.toLocaleString("ja-JP")}`);
  console.log(`   キャッシュ: ${TOKEN_CACHE_PATH}`);
  console.log("");
  console.log("━".repeat(60));
  console.log("  セットアップ完了！");
  console.log("");
  console.log("  Claude Desktop を起動（または再起動）すれば、");
  console.log("  Microsoft 365 のツールが使えるようになります。");
  console.log("");
  console.log("  トークンは自動的に更新されるため、");
  console.log("  通常はこのスクリプトを再実行する必要はありません。");
  console.log("━".repeat(60));
  console.log("");
}

main().catch((error) => {
  console.error("❌ エラー:", error.message);
  process.exit(1);
});
