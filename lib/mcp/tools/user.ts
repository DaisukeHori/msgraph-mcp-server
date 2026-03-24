import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphGet, handleToolError } from "@/lib/msgraph/graph-client";
import { getAuthMode } from "@/lib/msgraph/auth-context";

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
      title: "ユーザープロフィール取得",
      description: `サインイン中のユーザーのプロフィール情報を取得する。
認証テストとしても使用可能。

Returns: ユーザープロフィール（名前、メール、役職等）`,
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
      title: "ユーザー検索",
      description: `組織ディレクトリのユーザーを検索する。

Args:
  - query (必須): 検索クエリ（名前、メール等）
  - top: 最大件数 (1-50, デフォルト 10)

Returns: マッチしたユーザーの一覧`,
      inputSchema: {
        query: z.string().min(1).describe("検索クエリ"),
        top: z.number().int().min(1).max(50).default(10).describe("最大件数"),
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
      title: "認証ステータス",
      description: `現在の認証モードとステータスを確認する。

Returns: 認証モード、設定状況`,
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
        const mode = getAuthMode();
        const status: Record<string, unknown> = {
          authMode: mode,
          description:
            mode === "graph_token"
              ? "Bearer Token で Microsoft Graph アクセストークンを直接渡すモード"
              : mode === "client_credentials"
                ? "Azure AD クライアント資格情報フローで自動取得するモード"
                : "API キーで MCP サーバー認証 + クライアント資格情報フロー",
        };

        if (mode === "client_credentials" || mode === "api_key") {
          status.clientIdConfigured = !!process.env.MICROSOFT_CLIENT_ID;
          status.clientSecretConfigured = !!process.env.MICROSOFT_CLIENT_SECRET;
          status.tenantIdConfigured = !!process.env.MICROSOFT_TENANT_ID;
        }
        if (mode === "api_key") {
          status.mcpApiKeyConfigured = !!process.env.MCP_API_KEY;
        }

        // テスト: /me を呼んでトークンが有効か確認
        try {
          const me = await graphGet<UserProfile>("/me", { $select: "displayName,mail" });
          status.authenticated = true;
          status.user = { displayName: me.displayName, mail: me.mail };
        } catch (e) {
          status.authenticated = false;
          status.error = e instanceof Error ? e.message : String(e);
        }

        return {
          content: [{ type: "text", text: JSON.stringify(status, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );
}
