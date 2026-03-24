#!/usr/bin/env npx tsx
/**
 * stdio トランスポート（ローカル用）
 *
 * 本人として Microsoft 365 を操作するためのエントリポイント。
 * 初回実行時にブラウザで Microsoft アカウントにサインインし、
 * 以降はトークンキャッシュから自動リフレッシュされます。
 *
 * 使い方:
 *   AUTH_MODE=delegated \
 *   MICROSOFT_CLIENT_ID=your-client-id \
 *   MICROSOFT_TENANT_ID=your-tenant-id \
 *   npx tsx lib/stdio.ts
 *
 * Claude Desktop 設定 (claude_desktop_config.json):
 *   {
 *     "mcpServers": {
 *       "msgraph": {
 *         "command": "npx",
 *         "args": ["tsx", "/path/to/msgraph-mcp-server/lib/stdio.ts"],
 *         "env": {
 *           "AUTH_MODE": "delegated",
 *           "MICROSOFT_CLIENT_ID": "your-client-id",
 *           "MICROSOFT_TENANT_ID": "your-tenant-id"
 *         }
 *       }
 *     }
 *   }
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerAllTools } from "./mcp/server.js";

async function main() {
  // デフォルトは delegated モード
  if (!process.env.AUTH_MODE) {
    process.env.AUTH_MODE = "delegated";
  }

  const server = new McpServer({
    name: "msgraph-mcp-server",
    version: "1.1.0",
  });

  registerAllTools(server);

  const transport = new StdioServerTransport();
  await server.connect(transport);

  console.error("[msgraph-mcp-server] stdio トランスポートで起動しました");
  console.error(`[msgraph-mcp-server] AUTH_MODE: ${process.env.AUTH_MODE}`);
  console.error("[msgraph-mcp-server] 登録ツール数: 45");
  console.error("[msgraph-mcp-server] 初回のツール呼び出し時にサインインが求められます");
}

main().catch((error) => {
  console.error("起動エラー:", error);
  process.exit(1);
});
