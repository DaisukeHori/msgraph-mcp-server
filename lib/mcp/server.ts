/**
 * MCP サーバー初期化
 * 全ツールを一括登録する（83 tools）
 */

import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

import { registerMailTools } from "./tools/mail";
import { registerCalendarTools } from "./tools/calendar";
import { registerTeamsTools } from "./tools/teams";
import { registerOneDriveTools } from "./tools/onedrive";
import { registerSharePointTools } from "./tools/sharepoint";
import { registerUserTools } from "./tools/user";
import { registerWorkbookTools } from "./tools/workbook";

export function registerAllTools(server: McpServer): void {
  registerMailTools(server);       // 8 tools
  registerCalendarTools(server);   // 5 tools
  registerTeamsTools(server);      // 8 tools
  registerOneDriveTools(server);   // 12 tools
  registerSharePointTools(server); // 12 tools
  registerUserTools(server);       // 3 tools
  registerWorkbookTools(server);   // 35 tools (Excel Workbook API)
  // Total: 83 tools
}
