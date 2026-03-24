import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  graphGet,
  graphPost,
  truncateResponse,
  handleToolError,
  GraphPagedResponse,
} from "@/lib/msgraph/graph-client";
import { Team, Channel, ChatMessage, Chat } from "@/lib/msgraph/types";
import { DEFAULT_PAGE_SIZE } from "@/lib/config";
import { userBase, USER_ID_DESCRIPTION } from "./shared-helpers";

export function registerTeamsTools(server: McpServer): void {
  // -------------------------------------------------------
  // teams_list_joined_teams
  // -------------------------------------------------------
  server.registerTool(
    "teams_list_joined_teams",
    {
      title: "List Joined Teams",
      description: `List all Teams that the signed-in user is a member of.

Returns: List of teams with id, displayName, description, isArchived`,
      inputSchema: {
        user_id: z.string().optional().describe(USER_ID_DESCRIPTION),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const data = await graphGet<GraphPagedResponse<Team>>(`${userBase(params.user_id)}/joinedTeams`);
        const output = data.value.map((t) => ({
          id: t.id,
          displayName: t.displayName,
          description: t.description,
          isArchived: t.isArchived,
        }));
        return {
          content: [{ type: "text", text: JSON.stringify(output, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // teams_list_channels
  // -------------------------------------------------------
  server.registerTool(
    "teams_list_channels",
    {
      title: "List Team Channels",
      description: `List channels in a specific Team.

Args:
  - team_id (required): Team ID

Returns: List of channels with id, displayName, description, membershipType`,
      inputSchema: {
        team_id: z.string().min(1).describe("Team ID"),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const data = await graphGet<GraphPagedResponse<Channel>>(
          `/teams/${params.team_id}/channels`
        );
        const output = data.value.map((c) => ({
          id: c.id,
          displayName: c.displayName,
          description: c.description,
          membershipType: c.membershipType,
        }));
        return {
          content: [{ type: "text", text: JSON.stringify(output, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // teams_list_channel_messages
  // -------------------------------------------------------
  server.registerTool(
    "teams_list_channel_messages",
    {
      title: "List Channel Messages",
      description: `List messages in a Team channel.

Args:
  - team_id (required): Team ID
  - channel_id (required): Channel ID
  - top: Number of messages (1-50, default 25)

Returns: List of messages with sender, body, datetime, attachments`,
      inputSchema: {
        team_id: z.string().min(1).describe("Team ID"),
        channel_id: z.string().min(1).describe("Channel ID"),
        top: z.number().int().min(1).max(50).default(DEFAULT_PAGE_SIZE).describe("Number of messages"),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const data = await graphGet<GraphPagedResponse<ChatMessage>>(
          `/teams/${params.team_id}/channels/${params.channel_id}/messages`,
          { $top: params.top }
        );

        const messages = data.value.map((m) => ({
          id: m.id,
          from: m.from?.user?.displayName || m.from?.application?.displayName || "unknown",
          createdDateTime: m.createdDateTime,
          subject: m.subject,
          bodyPreview: m.body?.content?.replace(/<[^>]*>/g, "").slice(0, 300),
          importance: m.importance,
          attachments: m.attachments?.length || 0,
        }));

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: messages.length, messages }, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // teams_send_channel_message
  // -------------------------------------------------------
  server.registerTool(
    "teams_send_channel_message",
    {
      title: "Send Channel Message",
      description: `Send a message to a Team channel.

Args:
  - team_id (required): Team ID
  - channel_id (required): Channel ID
  - content (required): Message content (HTML supported)
  - content_type: "html" or "text" (default "html")
  - subject: Optional message subject
  - importance: "normal" or "urgent" (default "normal")

Returns: Created message details`,
      inputSchema: {
        team_id: z.string().min(1).describe("Team ID"),
        channel_id: z.string().min(1).describe("Channel ID"),
        content: z.string().min(1).describe("Message content"),
        content_type: z.enum(["html", "text"]).default("html").describe("Content type"),
        subject: z.string().optional().describe("Message subject"),
        importance: z.enum(["normal", "urgent"]).default("normal").describe("Importance"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    async (params) => {
      try {
        const body: Record<string, unknown> = {
          body: { contentType: params.content_type, content: params.content },
          importance: params.importance,
        };
        if (params.subject) body.subject = params.subject;

        const msg = await graphPost<ChatMessage>(
          `/teams/${params.team_id}/channels/${params.channel_id}/messages`,
          body
        );

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: msg.id, createdDateTime: msg.createdDateTime }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // teams_reply_to_channel_message
  // -------------------------------------------------------
  server.registerTool(
    "teams_reply_to_channel_message",
    {
      title: "Reply to Channel Message",
      description: `Reply to a message in a Team channel.

Args:
  - team_id (required): Team ID
  - channel_id (required): Channel ID
  - message_id (required): Parent message ID to reply to
  - content (required): Reply content
  - content_type: "html" or "text" (default "html")

Returns: Created reply details`,
      inputSchema: {
        team_id: z.string().min(1).describe("Team ID"),
        channel_id: z.string().min(1).describe("Channel ID"),
        message_id: z.string().min(1).describe("Parent message ID"),
        content: z.string().min(1).describe("Reply content"),
        content_type: z.enum(["html", "text"]).default("html").describe("Content type"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    async (params) => {
      try {
        const msg = await graphPost<ChatMessage>(
          `/teams/${params.team_id}/channels/${params.channel_id}/messages/${params.message_id}/replies`,
          { body: { contentType: params.content_type, content: params.content } }
        );

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: msg.id }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // teams_list_chats
  // -------------------------------------------------------
  server.registerTool(
    "teams_list_chats",
    {
      title: "List Chats",
      description: `List 1:1 and group chats for the signed-in user.

Args:
  - top: Number of chats (1-50, default 25)
  - filter: OData filter (e.g. "chatType eq 'oneOnOne'")

Returns: List of chats with id, topic, chatType, lastUpdatedDateTime`,
      inputSchema: {
        user_id: z.string().optional().describe(USER_ID_DESCRIPTION),
        top: z.number().int().min(1).max(50).default(DEFAULT_PAGE_SIZE).describe("Number of chats"),
        filter: z.string().optional().describe("OData filter"),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const queryParams: Record<string, string | number | boolean | undefined> = {
          $top: params.top,
          $expand: "members",
        };
        if (params.filter) queryParams.$filter = params.filter;

        const data = await graphGet<GraphPagedResponse<Chat>>(`${userBase(params.user_id)}/chats`, queryParams);

        const chats = data.value.map((c) => ({
          id: c.id,
          topic: c.topic,
          chatType: c.chatType,
          lastUpdatedDateTime: c.lastUpdatedDateTime,
          memberCount: (c.members || []).length,
        }));

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: chats.length, chats }, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // teams_list_chat_messages
  // -------------------------------------------------------
  server.registerTool(
    "teams_list_chat_messages",
    {
      title: "List Chat Messages",
      description: `List messages in a specific 1:1 or group chat.

Args:
  - chat_id (required): Chat ID
  - top: Number of messages (1-50, default 25)

Returns: List of chat messages`,
      inputSchema: {
        user_id: z.string().optional().describe(USER_ID_DESCRIPTION),
        chat_id: z.string().min(1).describe("Chat ID"),
        top: z.number().int().min(1).max(50).default(DEFAULT_PAGE_SIZE).describe("Number of messages"),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const data = await graphGet<GraphPagedResponse<ChatMessage>>(
          `${userBase(params.user_id)}/chats/${params.chat_id}/messages`,
          { $top: params.top }
        );

        const messages = data.value.map((m) => ({
          id: m.id,
          from: m.from?.user?.displayName || "unknown",
          createdDateTime: m.createdDateTime,
          bodyPreview: m.body?.content?.replace(/<[^>]*>/g, "").slice(0, 300),
          importance: m.importance,
        }));

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: messages.length, messages }, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // teams_send_chat_message
  // -------------------------------------------------------
  server.registerTool(
    "teams_send_chat_message",
    {
      title: "Send Chat Message",
      description: `Send a message in a 1:1 or group chat.

Args:
  - chat_id (required): Chat ID
  - content (required): Message content
  - content_type: "html" or "text" (default "html")

Returns: Created message details`,
      inputSchema: {
        user_id: z.string().optional().describe(USER_ID_DESCRIPTION),
        chat_id: z.string().min(1).describe("Chat ID"),
        content: z.string().min(1).describe("Message content"),
        content_type: z.enum(["html", "text"]).default("html").describe("Content type"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    async (params) => {
      try {
        const msg = await graphPost<ChatMessage>(
          `${userBase(params.user_id)}/chats/${params.chat_id}/messages`,
          { body: { contentType: params.content_type, content: params.content } }
        );

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: msg.id, createdDateTime: msg.createdDateTime }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );
}
