import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  graphGet,
  graphPost,
  graphPatch,
  graphDelete,
  truncateResponse,
  handleToolError,
  GraphPagedResponse,
} from "@/lib/msgraph/graph-client";
import { MailMessage, MailFolder, SendMailPayload } from "@/lib/msgraph/types";
import { DEFAULT_PAGE_SIZE } from "@/lib/config";

export function registerMailTools(server: McpServer): void {
  // -------------------------------------------------------
  // mail_list_messages
  // -------------------------------------------------------
  server.registerTool(
    "mail_list_messages",
    {
      title: "List Mail Messages",
      description: `List email messages from the signed-in user's mailbox.
Supports filtering by folder, search query, importance, read status, etc.
Returns subject, sender, date, preview, and metadata.

Args:
  - folder_id: Mail folder ID or well-known name (inbox, sentitems, drafts, deleteditems, junkemail). Default: inbox
  - search: Search query string (KQL supported, e.g. "subject:report from:john")
  - filter: OData filter expression (e.g. "isRead eq false", "importance eq 'high'")
  - top: Number of messages to return (1-50, default 25)
  - skip: Number of messages to skip for pagination
  - orderby: Sort order (e.g. "receivedDateTime desc")
  - select: Comma-separated fields to return

Returns: List of messages with id, subject, from, receivedDateTime, bodyPreview, isRead, importance`,
      inputSchema: {
        folder_id: z.string().default("inbox").describe("Folder ID or well-known name"),
        search: z.string().optional().describe("KQL search query"),
        filter: z.string().optional().describe("OData $filter expression"),
        top: z.number().int().min(1).max(50).default(DEFAULT_PAGE_SIZE).describe("Number of messages"),
        skip: z.number().int().min(0).default(0).describe("Pagination offset"),
        orderby: z.string().default("receivedDateTime desc").describe("Sort order"),
        select: z.string().optional().describe("Comma-separated fields"),
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
        const endpoint = `/me/mailFolders/${params.folder_id}/messages`;
        const queryParams: Record<string, string | number | boolean | undefined> = {
          $top: params.top,
          $skip: params.skip,
          $orderby: params.orderby,
          $select: params.select || "id,subject,bodyPreview,from,toRecipients,receivedDateTime,isRead,importance,hasAttachments",
        };
        if (params.search) queryParams.$search = `"${params.search}"`;
        if (params.filter) queryParams.$filter = params.filter;

        const data = await graphGet<GraphPagedResponse<MailMessage>>(endpoint, queryParams);
        const messages = data.value || [];

        const output = {
          count: messages.length,
          messages: messages.map((m) => ({
            id: m.id,
            subject: m.subject,
            from: m.from?.emailAddress
              ? `${m.from.emailAddress.name} <${m.from.emailAddress.address}>`
              : "unknown",
            receivedDateTime: m.receivedDateTime,
            bodyPreview: m.bodyPreview?.slice(0, 200),
            isRead: m.isRead,
            importance: m.importance,
            hasAttachments: m.hasAttachments,
          })),
          hasMore: !!data["@odata.nextLink"],
        };

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify(output, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // mail_get_message
  // -------------------------------------------------------
  server.registerTool(
    "mail_get_message",
    {
      title: "Get Mail Message",
      description: `Get a specific email message by ID, including full body content.

Args:
  - message_id (required): The message ID
  - include_body: Whether to include the full HTML/text body (default true)

Returns: Full message details including body, recipients, attachments info`,
      inputSchema: {
        message_id: z.string().min(1).describe("Message ID"),
        include_body: z.boolean().default(true).describe("Include full body"),
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
        const select = params.include_body
          ? undefined
          : "id,subject,bodyPreview,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,importance,hasAttachments,webLink,conversationId";
        const queryParams: Record<string, string | number | boolean | undefined> = {};
        if (select) queryParams.$select = select;

        const msg = await graphGet<MailMessage>(`/me/messages/${params.message_id}`, queryParams);

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify(msg, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // mail_send_message
  // -------------------------------------------------------
  server.registerTool(
    "mail_send_message",
    {
      title: "Send Mail Message",
      description: `Send an email message from the signed-in user.

Args:
  - to (required): Array of recipient email addresses
  - subject (required): Email subject
  - body (required): Email body content
  - body_type: "HTML" or "Text" (default "HTML")
  - cc: Array of CC recipient email addresses
  - save_to_sent: Save to Sent Items folder (default true)

Returns: Confirmation of sent message`,
      inputSchema: {
        to: z.array(z.string().email()).min(1).describe("Recipient email addresses"),
        subject: z.string().min(1).describe("Email subject"),
        body: z.string().min(1).describe("Email body content"),
        body_type: z.enum(["HTML", "Text"]).default("HTML").describe("Body content type"),
        cc: z.array(z.string().email()).optional().describe("CC recipients"),
        save_to_sent: z.boolean().default(true).describe("Save to Sent Items"),
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
        const payload: SendMailPayload = {
          message: {
            subject: params.subject,
            body: { contentType: params.body_type, content: params.body },
            toRecipients: params.to.map((addr) => ({
              emailAddress: { address: addr },
            })),
          },
          saveToSentItems: params.save_to_sent,
        };
        if (params.cc && params.cc.length > 0) {
          payload.message.ccRecipients = params.cc.map((addr) => ({
            emailAddress: { address: addr },
          }));
        }

        await graphPost<void>("/me/sendMail", payload);

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, message: `Email sent to ${params.to.join(", ")}` }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // mail_reply_message
  // -------------------------------------------------------
  server.registerTool(
    "mail_reply_message",
    {
      title: "Reply to Mail Message",
      description: `Reply to an email message. Can reply to sender or reply-all.

Args:
  - message_id (required): The message ID to reply to
  - comment (required): Reply body content
  - reply_all: Reply to all recipients (default false)

Returns: Confirmation`,
      inputSchema: {
        message_id: z.string().min(1).describe("Message ID to reply to"),
        comment: z.string().min(1).describe("Reply body content"),
        reply_all: z.boolean().default(false).describe("Reply to all recipients"),
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
        const action = params.reply_all ? "replyAll" : "reply";
        await graphPost<void>(`/me/messages/${params.message_id}/${action}`, {
          comment: params.comment,
        });

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, message: `Reply sent (${action})` }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // mail_update_message
  // -------------------------------------------------------
  server.registerTool(
    "mail_update_message",
    {
      title: "Update Mail Message",
      description: `Update properties of an email message (e.g., mark as read/unread, change importance, move to folder).

Args:
  - message_id (required): The message ID
  - is_read: Mark as read (true) or unread (false)
  - importance: Set importance ("low", "normal", "high")
  - categories: Set categories (array of strings)

Returns: Updated message`,
      inputSchema: {
        message_id: z.string().min(1).describe("Message ID"),
        is_read: z.boolean().optional().describe("Mark as read/unread"),
        importance: z.enum(["low", "normal", "high"]).optional().describe("Importance level"),
        categories: z.array(z.string()).optional().describe("Categories"),
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
        const body: Record<string, unknown> = {};
        if (params.is_read !== undefined) body.isRead = params.is_read;
        if (params.importance !== undefined) body.importance = params.importance;
        if (params.categories !== undefined) body.categories = params.categories;

        const msg = await graphPatch<MailMessage>(`/me/messages/${params.message_id}`, body);

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: msg.id, subject: msg.subject }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // mail_delete_message
  // -------------------------------------------------------
  server.registerTool(
    "mail_delete_message",
    {
      title: "Delete Mail Message",
      description: `Delete an email message (moves to Deleted Items).

Args:
  - message_id (required): The message ID to delete

Returns: Confirmation`,
      inputSchema: {
        message_id: z.string().min(1).describe("Message ID to delete"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        await graphDelete(`/me/messages/${params.message_id}`);
        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, message: "Message deleted (moved to Deleted Items)" }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // mail_move_message
  // -------------------------------------------------------
  server.registerTool(
    "mail_move_message",
    {
      title: "Move Mail Message",
      description: `Move an email message to a different folder.

Args:
  - message_id (required): The message ID to move
  - destination_folder (required): Destination folder ID or well-known name

Returns: Moved message details`,
      inputSchema: {
        message_id: z.string().min(1).describe("Message ID"),
        destination_folder: z.string().min(1).describe("Destination folder ID or well-known name"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const result = await graphPost<MailMessage>(
          `/me/messages/${params.message_id}/move`,
          { destinationId: params.destination_folder }
        );
        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: result.id, subject: result.subject }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // mail_list_folders
  // -------------------------------------------------------
  server.registerTool(
    "mail_list_folders",
    {
      title: "List Mail Folders",
      description: `List mail folders in the signed-in user's mailbox.

Returns: List of folders with ID, name, unread count, total count`,
      inputSchema: {
        top: z.number().int().min(1).max(100).default(50).describe("Number of folders"),
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
        const data = await graphGet<GraphPagedResponse<MailFolder>>("/me/mailFolders", {
          $top: params.top,
        });

        const output = data.value.map((f) => ({
          id: f.id,
          displayName: f.displayName,
          unreadItemCount: f.unreadItemCount,
          totalItemCount: f.totalItemCount,
          childFolderCount: f.childFolderCount,
        }));

        return {
          content: [{ type: "text", text: JSON.stringify(output, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );
}
