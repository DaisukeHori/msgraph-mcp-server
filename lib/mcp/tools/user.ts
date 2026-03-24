import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphGet, handleToolError } from "@/lib/msgraph/graph-client";

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
  server.registerTool(
    "user_get_profile",
    {
      title: "ユーザープロフィール取得",
      description: `サインイン中のユーザー（本人）のプロフィール情報を取得する。

Returns: ユーザープロフィール（名前、メール、役職等）`,
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async () => {
      try {
        const profile = await graphGet<UserProfile>("/me", {
          $select: "id,displayName,mail,userPrincipalName,jobTitle,officeLocation,mobilePhone,businessPhones",
        });
        return { content: [{ type: "text", text: JSON.stringify(profile, null, 2) }] };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

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
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true },
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
        return { content: [{ type: "text", text: JSON.stringify({ count: data.value.length, users: data.value }, null, 2) }] };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  server.registerTool(
    "auth_status",
    {
      title: "認証ステータス",
      description: `現在の認証ステータスを確認する。/me にアクセスしてトークンの有効性をテスト。

Returns: 認証状態、ユーザー情報`,
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async () => {
      try {
        const profile = await graphGet<UserProfile>("/me", {
          $select: "displayName,mail,userPrincipalName",
        });
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              authenticated: true,
              user: {
                displayName: profile.displayName,
                mail: profile.mail,
                userPrincipalName: profile.userPrincipalName,
              },
            }, null, 2),
          }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );
}
