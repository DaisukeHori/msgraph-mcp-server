#!/usr/bin/env npx tsx
/**
 * stdio トランスポート
 *
 * Claude Desktop / Claude Code からローカルで使う場合のエントリポイント。
 * Vercel デプロイとは独立して、ローカル stdio モードで動作する。
 *
 * 使い方:
 *   npx tsx lib/stdio.ts
 *
 * Claude Desktop 設定 (claude_desktop_config.json):
 *   {
 *     "mcpServers": {
 *       "msgraph": {
 *         "command": "npx",
 *         "args": ["tsx", "/path/to/msgraph-mcp-server/lib/stdio.ts"],
 *         "env": {
 *           "AUTH_MODE": "client_credentials",
 *           "MICROSOFT_CLIENT_ID": "your-client-id",
 *           "MICROSOFT_CLIENT_SECRET": "your-client-secret",
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
  const server = new McpServer({
    name: "msgraph-mcp-server",
    version: "1.1.0",
  });

  registerAllTools(server);

  const transport = new StdioServerTransport();
  await server.connect(transport);

  console.error("[msgraph-mcp-server] stdio トランスポートで起動しました");
  console.error("[msgraph-mcp-server] 登録ツール数: 45");
  console.error(`[msgraph-mcp-server] AUTH_MODE: ${process.env.AUTH_MODE || "graph_token"}`);
}

main().catch((error) => {
  console.error("起動エラー:", error);
  process.exit(1);
});
