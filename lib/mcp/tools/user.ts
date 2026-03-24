import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphGet, handleToolError } from "@/lib/msgraph/graph-client";
import { getAuthMode, clearAuthCache } from "@/lib/msgraph/auth-context";

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
delegated モードでは初回呼び出し時に Device Code Flow 認証が発動する。

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
      title: "認証ステータス確認・ログアウト",
      description: `現在の認証モードとステータスを確認する。
action="logout" でキャッシュをクリアして再認証を強制する。

Args:
  - action: "check" で確認、"logout" でキャッシュクリア

Returns: 認証モード、設定状況、現在のユーザー`,
      inputSchema: {
        action: z.enum(["check", "logout"]).default("check").describe("check: 確認 / logout: 再認証"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (params) => {
      try {
        if (params.action === "logout") {
          await clearAuthCache();
          return { content: [{ type: "text", text: JSON.stringify({
            success: true,
            message: "認証キャッシュをクリアしました。次のリクエストで再認証が必要になります。",
          }, null, 2) }] };
        }

        const mode = getAuthMode();
        const status: Record<string, unknown> = {
          authMode: mode,
          description: {
            delegated: "本人として操作（MSAL Device Code Flow / ローカル向け）",
            token: "本人として操作（Bearer Token 渡し / Vercel 向け）",
            client_credentials: "管理者アプリとして操作（/me/ 使用不可）",
          }[mode],
        };

        if (mode === "delegated" || mode === "client_credentials") {
          status.clientIdConfigured = !!process.env.MICROSOFT_CLIENT_ID;
          status.tenantIdConfigured = !!process.env.MICROSOFT_TENANT_ID;
        }
        if (mode === "client_credentials") {
          status.clientSecretConfigured = !!process.env.MICROSOFT_CLIENT_SECRET;
        }

        try {
          const me = await graphGet<UserProfile>("/me", { $select: "displayName,mail,userPrincipalName" });
          status.authenticated = true;
          status.user = { displayName: me.displayName, mail: me.mail, userPrincipalName: me.userPrincipalName };
        } catch (e) {
          status.authenticated = false;
          status.error = e instanceof Error ? e.message : String(e);
        }

        return { content: [{ type: "text", text: JSON.stringify(status, null, 2) }] };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );
}
