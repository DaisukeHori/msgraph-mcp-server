#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";

import { registerMailTools } from "./tools/mail.js";
import { registerCalendarTools } from "./tools/calendar.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerOneDriveTools } from "./tools/onedrive.js";
import { registerSharePointTools } from "./tools/sharepoint.js";
import { registerUserTools } from "./tools/user.js";

const server = new McpServer({
  name: "msgraph-mcp-server",
  version: "1.0.0",
});

// Register all tool domains
registerMailTools(server);
registerCalendarTools(server);
registerTeamsTools(server);
registerOneDriveTools(server);
registerSharePointTools(server);
registerUserTools(server);

// -------------------------------------------------------
// Transport: stdio (default) or HTTP
// -------------------------------------------------------
async function runStdio(): Promise<void> {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("[msgraph-mcp-server] Running on stdio transport");
  console.error(
    "[msgraph-mcp-server] Tools registered: mail(8), calendar(5), teams(8), onedrive(8), sharepoint(11), user(3) = 43 total"
  );
}

async function runHTTP(): Promise<void> {
  const { default: express } = await import("express");
  const app = express();
  app.use(express.json());

  app.post("/mcp", async (req, res) => {
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true,
    });
    res.on("close", () => transport.close());
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  });

  // Health check
  app.get("/health", (_req, res) => {
    res.json({ status: "ok", server: "msgraph-mcp-server", version: "1.0.0" });
  });

  const port = parseInt(process.env.PORT || "3100");
  app.listen(port, () => {
    console.error(`[msgraph-mcp-server] Running on HTTP transport at http://localhost:${port}/mcp`);
  });
}

// Choose transport
const transport = process.env.TRANSPORT || "stdio";
if (transport === "http") {
  runHTTP().catch((error) => {
    console.error("Server error:", error);
    process.exit(1);
  });
} else {
  runStdio().catch((error) => {
    console.error("Server error:", error);
    process.exit(1);
  });
}
