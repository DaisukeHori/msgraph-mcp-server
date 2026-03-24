import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  graphGet,
  graphPost,
  graphPatch,
  graphDelete,
  graphUploadSmallFile,
  graphDownloadFile,
  truncateResponse,
  handleToolError,
  GraphPagedResponse,
} from "@/lib/msgraph/graph-client";
import { DriveItem, Drive } from "@/lib/msgraph/types";
import { DEFAULT_PAGE_SIZE } from "@/lib/config";

export function registerOneDriveTools(server: McpServer): void {
  // -------------------------------------------------------
  // onedrive_get_drive
  // -------------------------------------------------------
  server.registerTool(
    "onedrive_get_drive",
    {
      title: "Get OneDrive Info",
      description: `Get information about the signed-in user's OneDrive, including quota usage.

Returns: Drive details with id, name, quota (total, used, remaining)`,
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
        const drive = await graphGet<Drive>("/me/drive");
        return {
          content: [{ type: "text", text: JSON.stringify(drive, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // onedrive_list_items
  // -------------------------------------------------------
  server.registerTool(
    "onedrive_list_items",
    {
      title: "List OneDrive Items",
      description: `List files and folders in a OneDrive location.

Args:
  - path: Folder path (e.g. "/Documents/Projects"). Default: root
  - folder_id: Folder ID (alternative to path)
  - drive_id: Drive ID (default: user's OneDrive)
  - top: Number of items (1-100, default 25)
  - orderby: Sort (e.g. "name", "lastModifiedDateTime desc")
  - filter: OData filter (e.g. "file ne null" for files only)

Returns: List of items with name, size, type, dates, webUrl`,
      inputSchema: {
        path: z.string().optional().describe("Folder path from root"),
        folder_id: z.string().optional().describe("Folder ID"),
        drive_id: z.string().optional().describe("Drive ID"),
        top: z.number().int().min(1).max(100).default(DEFAULT_PAGE_SIZE).describe("Number of items"),
        orderby: z.string().optional().describe("Sort order"),
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
        let endpoint: string;
        const drivePrefix = params.drive_id
          ? `/drives/${params.drive_id}`
          : "/me/drive";

        if (params.folder_id) {
          endpoint = `${drivePrefix}/items/${params.folder_id}/children`;
        } else if (params.path) {
          const cleanPath = params.path.startsWith("/") ? params.path : `/${params.path}`;
          endpoint = `${drivePrefix}/root:${cleanPath}:/children`;
        } else {
          endpoint = `${drivePrefix}/root/children`;
        }

        const queryParams: Record<string, string | number | boolean | undefined> = {
          $top: params.top,
          $select: "id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,parentReference",
        };
        if (params.orderby) queryParams.$orderby = params.orderby;
        if (params.filter) queryParams.$filter = params.filter;

        const data = await graphGet<GraphPagedResponse<DriveItem>>(endpoint, queryParams);

        const items = data.value.map((item) => ({
          id: item.id,
          name: item.name,
          type: item.folder ? "folder" : "file",
          size: item.size,
          mimeType: item.file?.mimeType,
          childCount: item.folder?.childCount,
          lastModifiedDateTime: item.lastModifiedDateTime,
          webUrl: item.webUrl,
        }));

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: items.length, items, hasMore: !!data["@odata.nextLink"] }, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // onedrive_get_item
  // -------------------------------------------------------
  server.registerTool(
    "onedrive_get_item",
    {
      title: "Get OneDrive Item",
      description: `Get metadata of a specific file or folder.

Args:
  - item_id: Item ID (required if no path)
  - path: Item path (alternative to item_id, e.g. "/Documents/report.xlsx")
  - drive_id: Drive ID (default: user's OneDrive)

Returns: Full item metadata`,
      inputSchema: {
        item_id: z.string().optional().describe("Item ID"),
        path: z.string().optional().describe("Item path"),
        drive_id: z.string().optional().describe("Drive ID"),
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
        const drivePrefix = params.drive_id
          ? `/drives/${params.drive_id}`
          : "/me/drive";
        let endpoint: string;

        if (params.item_id) {
          endpoint = `${drivePrefix}/items/${params.item_id}`;
        } else if (params.path) {
          const cleanPath = params.path.startsWith("/") ? params.path : `/${params.path}`;
          endpoint = `${drivePrefix}/root:${cleanPath}`;
        } else {
          throw new Error("Either item_id or path is required");
        }

        const item = await graphGet<DriveItem>(endpoint);
        return {
          content: [{ type: "text", text: JSON.stringify(item, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // onedrive_download_file
  // -------------------------------------------------------
  server.registerTool(
    "onedrive_download_file",
    {
      title: "Download OneDrive File",
      description: `Download the content of a file from OneDrive.
For text files, returns the text. For binary files, returns base64-encoded content.

Args:
  - item_id: File ID (required if no path)
  - path: File path (alternative)
  - drive_id: Drive ID (default: user's OneDrive)

Returns: File content (text or base64) and content type`,
      inputSchema: {
        item_id: z.string().optional().describe("Item ID"),
        path: z.string().optional().describe("File path"),
        drive_id: z.string().optional().describe("Drive ID"),
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
        const drivePrefix = params.drive_id
          ? `/drives/${params.drive_id}`
          : "/me/drive";
        let endpoint: string;

        if (params.item_id) {
          endpoint = `${drivePrefix}/items/${params.item_id}/content`;
        } else if (params.path) {
          const cleanPath = params.path.startsWith("/") ? params.path : `/${params.path}`;
          endpoint = `${drivePrefix}/root:${cleanPath}:/content`;
        } else {
          throw new Error("Either item_id or path is required");
        }

        const result = await graphDownloadFile(endpoint);
        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify(result, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // onedrive_upload_file
  // -------------------------------------------------------
  server.registerTool(
    "onedrive_upload_file",
    {
      title: "Upload File to OneDrive",
      description: `Upload a file to OneDrive (< 4MB). For larger files, use upload session.

Args:
  - path (required): Destination path including filename (e.g. "/Documents/report.txt")
  - content (required): File content (text or base64 for binary)
  - content_type: MIME type (default "text/plain")
  - drive_id: Drive ID (default: user's OneDrive)
  - conflict_behavior: "rename"|"replace"|"fail" (default "replace")

Returns: Created/updated item details`,
      inputSchema: {
        path: z.string().min(1).describe("Destination path with filename"),
        content: z.string().min(1).describe("File content"),
        content_type: z.string().default("text/plain").describe("MIME type"),
        drive_id: z.string().optional().describe("Drive ID"),
        conflict_behavior: z.enum(["rename", "replace", "fail"]).default("replace").describe("Conflict behavior"),
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
        const drivePrefix = params.drive_id
          ? `/drives/${params.drive_id}`
          : "/me/drive";
        const cleanPath = params.path.startsWith("/") ? params.path : `/${params.path}`;
        const endpoint = `${drivePrefix}/root:${cleanPath}:/content`;

        // Determine if content is base64 binary or plain text
        let buffer: Buffer;
        if (params.content_type !== "text/plain" && !params.content_type.startsWith("text/")) {
          buffer = Buffer.from(params.content, "base64");
        } else {
          buffer = Buffer.from(params.content, "utf-8");
        }

        const item = await graphUploadSmallFile<DriveItem>(endpoint, buffer, params.content_type);

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: item.id, name: item.name, webUrl: item.webUrl, size: item.size }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // onedrive_create_folder
  // -------------------------------------------------------
  server.registerTool(
    "onedrive_create_folder",
    {
      title: "Create OneDrive Folder",
      description: `Create a new folder in OneDrive.

Args:
  - name (required): Folder name
  - parent_path: Parent folder path (default: root)
  - parent_id: Parent folder ID (alternative to parent_path)
  - drive_id: Drive ID
  - conflict_behavior: "rename"|"replace"|"fail" (default "rename")

Returns: Created folder details`,
      inputSchema: {
        name: z.string().min(1).describe("Folder name"),
        parent_path: z.string().optional().describe("Parent folder path"),
        parent_id: z.string().optional().describe("Parent folder ID"),
        drive_id: z.string().optional().describe("Drive ID"),
        conflict_behavior: z.enum(["rename", "replace", "fail"]).default("rename").describe("Conflict behavior"),
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
        const drivePrefix = params.drive_id
          ? `/drives/${params.drive_id}`
          : "/me/drive";
        let endpoint: string;

        if (params.parent_id) {
          endpoint = `${drivePrefix}/items/${params.parent_id}/children`;
        } else if (params.parent_path) {
          const cleanPath = params.parent_path.startsWith("/") ? params.parent_path : `/${params.parent_path}`;
          endpoint = `${drivePrefix}/root:${cleanPath}:/children`;
        } else {
          endpoint = `${drivePrefix}/root/children`;
        }

        const body = {
          name: params.name,
          folder: {},
          "@microsoft.graph.conflictBehavior": params.conflict_behavior,
        };

        const folder = await graphPost<DriveItem>(endpoint, body);

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: folder.id, name: folder.name, webUrl: folder.webUrl }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // onedrive_delete_item
  // -------------------------------------------------------
  server.registerTool(
    "onedrive_delete_item",
    {
      title: "Delete OneDrive Item",
      description: `Delete a file or folder from OneDrive (moves to recycle bin).

Args:
  - item_id (required): Item ID to delete
  - drive_id: Drive ID

Returns: Confirmation`,
      inputSchema: {
        item_id: z.string().min(1).describe("Item ID"),
        drive_id: z.string().optional().describe("Drive ID"),
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
        const drivePrefix = params.drive_id
          ? `/drives/${params.drive_id}`
          : "/me/drive";
        await graphDelete(`${drivePrefix}/items/${params.item_id}`);
        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, message: "Item deleted (moved to recycle bin)" }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // onedrive_move_item
  // -------------------------------------------------------
  server.registerTool(
    "onedrive_move_item",
    {
      title: "Move/Rename OneDrive Item",
      description: `Move a file or folder to a new location, and/or rename it.

Args:
  - item_id (required): Item ID to move
  - new_name: New name for the item
  - destination_folder_id: Destination folder ID
  - destination_drive_id: Destination drive ID (for cross-drive moves)
  - drive_id: Source drive ID

Returns: Updated item details`,
      inputSchema: {
        item_id: z.string().min(1).describe("Item ID"),
        new_name: z.string().optional().describe("New name"),
        destination_folder_id: z.string().optional().describe("Destination folder ID"),
        destination_drive_id: z.string().optional().describe("Destination drive ID"),
        drive_id: z.string().optional().describe("Source drive ID"),
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
        const drivePrefix = params.drive_id
          ? `/drives/${params.drive_id}`
          : "/me/drive";

        const body: Record<string, unknown> = {};
        if (params.new_name) body.name = params.new_name;
        if (params.destination_folder_id) {
          const parentRef: Record<string, string> = { id: params.destination_folder_id };
          if (params.destination_drive_id) parentRef.driveId = params.destination_drive_id;
          body.parentReference = parentRef;
        }

        const item = await graphPatch<DriveItem>(
          `${drivePrefix}/items/${params.item_id}`,
          body
        );

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: item.id, name: item.name, webUrl: item.webUrl }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // onedrive_search
  // -------------------------------------------------------
  server.registerTool(
    "onedrive_search",
    {
      title: "Search OneDrive",
      description: `Search for files and folders in OneDrive.

Args:
  - query (required): Search query string
  - drive_id: Drive ID (default: user's OneDrive)
  - top: Max results (1-50, default 25)

Returns: List of matching items`,
      inputSchema: {
        query: z.string().min(1).describe("Search query"),
        drive_id: z.string().optional().describe("Drive ID"),
        top: z.number().int().min(1).max(50).default(DEFAULT_PAGE_SIZE).describe("Max results"),
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
        const drivePrefix = params.drive_id
          ? `/drives/${params.drive_id}`
          : "/me/drive";

        const data = await graphGet<GraphPagedResponse<DriveItem>>(
          `${drivePrefix}/root/search(q='${encodeURIComponent(params.query)}')`,
          { $top: params.top, $select: "id,name,size,lastModifiedDateTime,webUrl,file,folder,parentReference" }
        );

        const items = data.value.map((item) => ({
          id: item.id,
          name: item.name,
          type: item.folder ? "folder" : "file",
          size: item.size,
          mimeType: item.file?.mimeType,
          lastModifiedDateTime: item.lastModifiedDateTime,
          webUrl: item.webUrl,
          path: item.parentReference?.path,
        }));

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: items.length, items }, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // onedrive_shared_with_me
  // -------------------------------------------------------
  server.registerTool(
    "onedrive_shared_with_me",
    {
      title: "List Files Shared With Me",
      description: `他のユーザーから共有されたファイルやフォルダの一覧を取得する。
OneDrive for Business で他人が「共有」したアイテムが表示される。

Args:
  - top: 最大件数 (1-100, default 25)
  - filter: OData filter

Returns: 共有アイテム一覧（名前、共有者、ドライブID、リモートアイテムID）`,
      inputSchema: {
        top: z.number().int().min(1).max(100).default(25).describe("Max results"),
        filter: z.string().optional().describe("OData filter"),
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
        const queryParams: Record<string, string | number> = {
          $top: params.top,
        };
        if (params.filter) queryParams.$filter = params.filter;

        const data = await graphGet<GraphPagedResponse<DriveItem>>(
          "/me/drive/sharedWithMe",
          queryParams
        );

        const items = data.value.map((item) => ({
          id: item.id,
          name: item.name,
          type: item.folder ? "folder" : "file",
          size: item.size,
          mimeType: item.file?.mimeType,
          lastModifiedDateTime: item.lastModifiedDateTime,
          webUrl: item.webUrl,
          // 共有元の情報
          sharedBy: item.shared
            ? {
                owner: item.shared.owner,
                sharedBy: item.shared.sharedBy,
                sharedDateTime: item.shared.sharedDateTime,
                scope: item.shared.scope,
              }
            : undefined,
          // リモートアイテム情報（他人のドライブ上のファイルへのアクセスに使う）
          remoteItem: item.remoteItem
            ? {
                id: item.remoteItem.id,
                name: item.remoteItem.name,
                driveId: item.remoteItem.parentReference?.driveId,
                webUrl: item.remoteItem.webUrl,
                size: item.remoteItem.size,
                file: item.remoteItem.file,
                folder: item.remoteItem.folder,
                lastModifiedBy: item.remoteItem.lastModifiedBy,
              }
            : undefined,
        }));

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: items.length, items }, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // onedrive_shared_item_browse
  // -------------------------------------------------------
  server.registerTool(
    "onedrive_shared_item_browse",
    {
      title: "Browse Shared Drive/Folder",
      description: `他人の OneDrive for Business のドライブやフォルダの中身を閲覧する。
sharedWithMe で取得した remoteItem.driveId と remoteItem.id を使ってアクセス。

Args:
  - drive_id (必須): 共有元のドライブ ID（remoteItem.driveId）
  - item_id: フォルダ ID（remoteItem.id）。省略時はドライブのルート
  - top: 最大件数 (1-100, default 25)

Returns: ファイル/フォルダ一覧`,
      inputSchema: {
        drive_id: z.string().min(1).describe("Shared drive ID (from remoteItem.driveId)"),
        item_id: z.string().optional().describe("Folder item ID (from remoteItem.id)"),
        top: z.number().int().min(1).max(100).default(25).describe("Max results"),
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
        const endpoint = params.item_id
          ? `/drives/${params.drive_id}/items/${params.item_id}/children`
          : `/drives/${params.drive_id}/root/children`;

        const data = await graphGet<GraphPagedResponse<DriveItem>>(
          endpoint,
          { $top: params.top }
        );

        const items = data.value.map((item) => ({
          id: item.id,
          name: item.name,
          type: item.folder ? "folder" : "file",
          size: item.size,
          mimeType: item.file?.mimeType,
          lastModifiedDateTime: item.lastModifiedDateTime,
          webUrl: item.webUrl,
          lastModifiedBy: item.lastModifiedBy,
        }));

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: items.length, items }, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );
}
