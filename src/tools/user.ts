import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphGet, handleToolError } from "../services/graph-client.js";
import { isAuthenticated, clearTokenCache } from "../services/auth.js";

interface UserProfile {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle?: string;
  officeLocation?: string;
  mobilePhone?: string;
  businessPhones?: string[];
}

export function registerUserTools(server: McpServer): void {
  // -------------------------------------------------------
  // user_get_profile
  // -------------------------------------------------------
  server.registerTool(
    "user_get_profile",
    {
      title: "Get User Profile",
      description: `Get the signed-in user's profile information.
Also triggers authentication if not yet authenticated (device code flow).

Returns: User profile with name, email, job title, etc.`,
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async () => {
      try {
        const profile = await graphGet<UserProfile>("/me", {
          $select: "id,displayName,mail,userPrincipalName,jobTitle,officeLocation,mobilePhone,businessPhones",
        });
        return {
          content: [{ type: "text", text: JSON.stringify(profile, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // user_search_users
  // -------------------------------------------------------
  server.registerTool(
    "user_search_users",
    {
      title: "Search Users",
      description: `Search for users in the organization directory.

Args:
  - query (required): Search query (name, email, etc.)
  - top: Max results (1-50, default 10)

Returns: List of matching users`,
      inputSchema: {
        query: z.string().min(1).describe("Search query"),
        top: z.number().int().min(1).max(50).default(10).describe("Max results"),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async (params) => {
      try {
        const data = await graphGet<{ value: UserProfile[] }>("/users", {
          $search: `"displayName:${params.query}" OR "mail:${params.query}"`,
          $top: params.top,
          $select: "id,displayName,mail,userPrincipalName,jobTitle,officeLocation",
          $orderby: "displayName",
          ConsistencyLevel: "eventual",
        });

        return {
          content: [{ type: "text", text: JSON.stringify({ count: data.value.length, users: data.value }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // auth_status
  // -------------------------------------------------------
  server.registerTool(
    "auth_status",
    {
      title: "Authentication Status",
      description: `Check authentication status and clear token cache if needed.

Args:
  - action: "check" to check status, "logout" to clear tokens

Returns: Authentication status`,
      inputSchema: {
        action: z.enum(["check", "logout"]).default("check").describe("Action"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        if (params.action === "logout") {
          await clearTokenCache();
          return {
            content: [{ type: "text", text: JSON.stringify({ authenticated: false, message: "Token cache cleared. You will need to re-authenticate on next request." }) }],
          };
        }

        const authenticated = await isAuthenticated();
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                authenticated,
                message: authenticated
                  ? "Authenticated. Token cache exists."
                  : "Not authenticated. Call user_get_profile or any tool to trigger authentication.",
              }),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );
}
