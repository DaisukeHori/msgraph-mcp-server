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
import {
  SharePointSite,
  SharePointList,
  SharePointListItem,
  SharePointColumn,
  Drive,
  DriveItem,
} from "@/lib/msgraph/types";
import { DEFAULT_PAGE_SIZE } from "@/lib/config";

export function registerSharePointTools(server: McpServer): void {
  // -------------------------------------------------------
  // sharepoint_search_sites
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_search_sites",
    {
      title: "Search SharePoint Sites",
      description: `Search for SharePoint sites accessible to the signed-in user.

Args:
  - query: Search query (partial site name). If empty, returns recent/popular sites.
  - top: Max results (1-50, default 25)

Returns: List of sites with id, displayName, webUrl, description`,
      inputSchema: {
        query: z.string().optional().describe("Search query"),
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
        let endpoint: string;
        const queryParams: Record<string, string | number | boolean | undefined> = {
          $top: params.top,
        };

        if (params.query) {
          endpoint = `/sites`;
          queryParams.$search = `"${params.query}"`;
        } else {
          // Get followed sites or root site
          endpoint = `/sites`;
          queryParams.$search = `"*"`;
        }

        const data = await graphGet<GraphPagedResponse<SharePointSite>>(endpoint, queryParams);

        const sites = data.value.map((s) => ({
          id: s.id,
          name: s.name,
          displayName: s.displayName,
          webUrl: s.webUrl,
          description: s.description,
          lastModifiedDateTime: s.lastModifiedDateTime,
        }));

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: sites.length, sites }, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // sharepoint_get_site
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_get_site",
    {
      title: "Get SharePoint Site",
      description: `Get details of a specific SharePoint site.

Args:
  - site_id: Site ID (e.g. "contoso.sharepoint.com,guid1,guid2")
  - hostname: Site hostname (e.g. "contoso.sharepoint.com")
  - server_relative_path: Server relative path (e.g. "/sites/Marketing")

Use site_id OR (hostname + server_relative_path).

Returns: Site details`,
      inputSchema: {
        site_id: z.string().optional().describe("Site ID"),
        hostname: z.string().optional().describe("Site hostname"),
        server_relative_path: z.string().optional().describe("Server relative path"),
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
        if (params.site_id) {
          endpoint = `/sites/${params.site_id}`;
        } else if (params.hostname) {
          const path = params.server_relative_path || "";
          const cleanPath = path.startsWith(":") ? path : `:${path.startsWith("/") ? path : `/${path}`}`;
          endpoint = `/sites/${params.hostname}${cleanPath}`;
        } else {
          throw new Error("Either site_id or hostname is required");
        }

        const site = await graphGet<SharePointSite>(endpoint);
        return {
          content: [{ type: "text", text: JSON.stringify(site, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // sharepoint_list_drives
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_list_drives",
    {
      title: "List SharePoint Document Libraries",
      description: `List document libraries (drives) in a SharePoint site.

Args:
  - site_id (required): Site ID

Returns: List of drives/document libraries`,
      inputSchema: {
        site_id: z.string().min(1).describe("Site ID"),
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
        const data = await graphGet<GraphPagedResponse<Drive>>(
          `/sites/${params.site_id}/drives`
        );
        const drives = data.value.map((d) => ({
          id: d.id,
          name: d.name,
          driveType: d.driveType,
          webUrl: d.webUrl,
          quota: d.quota,
        }));
        return {
          content: [{ type: "text", text: JSON.stringify(drives, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // sharepoint_list_drive_items
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_list_drive_items",
    {
      title: "List SharePoint Drive Items",
      description: `List files and folders in a SharePoint document library.

Args:
  - drive_id (required): Drive/document library ID
  - folder_id: Folder ID (default: root)
  - path: Folder path (alternative to folder_id)
  - top: Max results (1-100, default 25)

Returns: List of items in the library`,
      inputSchema: {
        drive_id: z.string().min(1).describe("Drive ID"),
        folder_id: z.string().optional().describe("Folder ID"),
        path: z.string().optional().describe("Folder path"),
        top: z.number().int().min(1).max(100).default(DEFAULT_PAGE_SIZE).describe("Max results"),
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
        if (params.folder_id) {
          endpoint = `/drives/${params.drive_id}/items/${params.folder_id}/children`;
        } else if (params.path) {
          const cleanPath = params.path.startsWith("/") ? params.path : `/${params.path}`;
          endpoint = `/drives/${params.drive_id}/root:${cleanPath}:/children`;
        } else {
          endpoint = `/drives/${params.drive_id}/root/children`;
        }

        const data = await graphGet<GraphPagedResponse<DriveItem>>(endpoint, {
          $top: params.top,
          $select: "id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,parentReference",
        });

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
          content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: items.length, items }, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // sharepoint_get_lists
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_get_lists",
    {
      title: "List SharePoint Lists",
      description: `Get all SharePoint lists available on a site.

Args:
  - site_id (required): Site ID
  - filter: OData filter (e.g. "list/hidden eq false")

Returns: List of SharePoint lists with id, name, description`,
      inputSchema: {
        site_id: z.string().min(1).describe("Site ID"),
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
        const queryParams: Record<string, string | number | boolean | undefined> = {};
        if (params.filter) queryParams.$filter = params.filter;

        const data = await graphGet<GraphPagedResponse<SharePointList>>(
          `/sites/${params.site_id}/lists`,
          queryParams
        );

        const lists = data.value.map((l) => ({
          id: l.id,
          name: l.name,
          displayName: l.displayName,
          description: l.description,
          webUrl: l.webUrl,
          hidden: l.list?.hidden,
          template: l.list?.template,
          lastModifiedDateTime: l.lastModifiedDateTime,
        }));

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: lists.length, lists }, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // sharepoint_get_list_columns
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_get_list_columns",
    {
      title: "Get SharePoint List Columns",
      description: `Get column definitions for a SharePoint list. Useful for understanding the schema before querying items.

Args:
  - site_id (required): Site ID
  - list_id (required): List ID

Returns: List of columns with name, displayName, type, readOnly, required`,
      inputSchema: {
        site_id: z.string().min(1).describe("Site ID"),
        list_id: z.string().min(1).describe("List ID"),
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
        const data = await graphGet<GraphPagedResponse<SharePointColumn>>(
          `/sites/${params.site_id}/lists/${params.list_id}/columns`
        );

        const columns = data.value.map((c) => ({
          id: c.id,
          name: c.name,
          displayName: c.displayName,
          description: c.description,
          readOnly: c.readOnly,
          required: c.required,
        }));

        return {
          content: [{ type: "text", text: JSON.stringify({ count: columns.length, columns }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // sharepoint_get_list_items
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_get_list_items",
    {
      title: "Get SharePoint List Items",
      description: `Get items (rows) from a SharePoint list.

Args:
  - site_id (required): Site ID
  - list_id (required): List ID
  - filter: OData filter (e.g. "fields/Status eq 'Active'")
  - expand: Expand fields (default "fields")
  - select: Fields to select from expanded fields (e.g. "fields/Title,fields/Status")
  - top: Max items (1-100, default 25)
  - orderby: Sort order

Returns: List items with their field values`,
      inputSchema: {
        site_id: z.string().min(1).describe("Site ID"),
        list_id: z.string().min(1).describe("List ID"),
        filter: z.string().optional().describe("OData filter on fields"),
        expand: z.string().default("fields").describe("Expand expression"),
        select: z.string().optional().describe("Field selection"),
        top: z.number().int().min(1).max(100).default(DEFAULT_PAGE_SIZE).describe("Max items"),
        orderby: z.string().optional().describe("Sort order"),
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
          $expand: params.expand,
          $top: params.top,
        };
        if (params.filter) queryParams.$filter = params.filter;
        if (params.select) queryParams.$select = params.select;
        if (params.orderby) queryParams.$orderby = params.orderby;

        const data = await graphGet<GraphPagedResponse<SharePointListItem>>(
          `/sites/${params.site_id}/lists/${params.list_id}/items`,
          queryParams
        );

        const items = data.value.map((item) => ({
          id: item.id,
          fields: item.fields,
          createdDateTime: item.createdDateTime,
          lastModifiedDateTime: item.lastModifiedDateTime,
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
  // sharepoint_get_list_item
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_get_list_item",
    {
      title: "Get SharePoint List Item",
      description: `Get a single item from a SharePoint list by its ID.

Args:
  - site_id (required): Site ID
  - list_id (required): List ID
  - item_id (required): Item ID

Returns: Item with all fields`,
      inputSchema: {
        site_id: z.string().min(1).describe("Site ID"),
        list_id: z.string().min(1).describe("List ID"),
        item_id: z.string().min(1).describe("Item ID"),
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
        const item = await graphGet<SharePointListItem>(
          `/sites/${params.site_id}/lists/${params.list_id}/items/${params.item_id}`,
          { $expand: "fields" }
        );
        return {
          content: [{ type: "text", text: JSON.stringify(item, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // sharepoint_create_list_item
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_create_list_item",
    {
      title: "Create SharePoint List Item",
      description: `Create a new item in a SharePoint list.

Args:
  - site_id (required): Site ID
  - list_id (required): List ID
  - fields (required): Key-value pairs for field values (JSON object).
    The "Title" field is typically required.
    Example: {"Title": "New Item", "Status": "Active", "Priority": "High"}

Returns: Created item with ID and fields`,
      inputSchema: {
        site_id: z.string().min(1).describe("Site ID"),
        list_id: z.string().min(1).describe("List ID"),
        fields: z.record(z.unknown()).describe("Field key-value pairs"),
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
        const item = await graphPost<SharePointListItem>(
          `/sites/${params.site_id}/lists/${params.list_id}/items`,
          { fields: params.fields }
        );

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: item.id, fields: item.fields }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // sharepoint_update_list_item
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_update_list_item",
    {
      title: "Update SharePoint List Item",
      description: `Update an existing item in a SharePoint list. Only specified fields are updated.

Args:
  - site_id (required): Site ID
  - list_id (required): List ID
  - item_id (required): Item ID
  - fields (required): Key-value pairs to update
    Example: {"Status": "Completed", "Notes": "Done on time"}

Returns: Updated item`,
      inputSchema: {
        site_id: z.string().min(1).describe("Site ID"),
        list_id: z.string().min(1).describe("List ID"),
        item_id: z.string().min(1).describe("Item ID"),
        fields: z.record(z.unknown()).describe("Field key-value pairs to update"),
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
        const item = await graphPatch<SharePointListItem>(
          `/sites/${params.site_id}/lists/${params.list_id}/items/${params.item_id}/fields`,
          params.fields
        );

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: params.item_id, updatedFields: item }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // sharepoint_delete_list_item
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_delete_list_item",
    {
      title: "Delete SharePoint List Item",
      description: `Delete an item from a SharePoint list.

Args:
  - site_id (required): Site ID
  - list_id (required): List ID
  - item_id (required): Item ID

Returns: Confirmation`,
      inputSchema: {
        site_id: z.string().min(1).describe("Site ID"),
        list_id: z.string().min(1).describe("List ID"),
        item_id: z.string().min(1).describe("Item ID"),
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
        await graphDelete(
          `/sites/${params.site_id}/lists/${params.list_id}/items/${params.item_id}`
        );
        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, message: "List item deleted" }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // sharepoint_create_list
  // -------------------------------------------------------
  server.registerTool(
    "sharepoint_create_list",
    {
      title: "Create SharePoint List",
      description: `Create a new SharePoint list on a site.

Args:
  - site_id (required): Site ID
  - display_name (required): List display name
  - description: List description
  - template: List template ("genericList", "documentLibrary", "events", etc.) Default: "genericList"

Returns: Created list details`,
      inputSchema: {
        site_id: z.string().min(1).describe("Site ID"),
        display_name: z.string().min(1).describe("Display name"),
        description: z.string().optional().describe("Description"),
        template: z.string().default("genericList").describe("List template"),
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
        const body: Record<string, unknown> = {
          displayName: params.display_name,
          list: { template: params.template },
        };
        if (params.description) body.description = params.description;

        const list = await graphPost<SharePointList>(
          `/sites/${params.site_id}/lists`,
          body
        );

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: list.id, displayName: list.displayName, webUrl: list.webUrl }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );
}
