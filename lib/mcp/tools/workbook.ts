/**
 * Excel Workbook MCP Tools (35 tools)
 *
 * Microsoft Graph Excel Workbook API をラップする。OneDrive / SharePoint
 * 上の .xlsx ファイルに対して、ワークシート・テーブル・行・列・範囲・
 * グラフ・関数の CRUD を提供する。
 *
 * カテゴリ:
 *   A. Sessions      (2)  - workbook_create_session / workbook_close_session
 *   B. Worksheets    (5)  - list / get / add / update / delete
 *   C. Tables        (6)  - list / get / create / update / delete / convert_to_range
 *   D. Table Rows    (5)  - add_rows / list_rows / get_row / update_row / delete_row
 *   E. Table Columns (4)  - list / add / update / delete
 *   F. Range         (8)  - get / update / clear / get_used / insert / delete / merge / unmerge
 *   G. Functions     (1)  - call_function
 *   H. Charts        (4)  - list / create / get_image / delete
 *
 * 共通入力（全ツール）:
 *   - item_id / path        : Excel ファイル指定（どちらか必須）
 *   - drive_id              : 他人ドライブ・SharePoint ドライブ用（optional）
 *   - user_id               : 委任ユーザー（共有メールボックス等、optional）
 *   - workbook_session_id   : 永続セッションID。複数操作の前に
 *                             workbook_create_session で取得して渡すと
 *                             パフォーマンス向上 + 即時反映（optional）
 *
 * セッション運用:
 *   1. workbook_create_session で session_id を取得
 *   2. 各ツールに workbook_session_id を渡しつつ操作
 *   3. workbook_close_session で閉じる（5分間アイドルで自動破棄）
 *
 * Note:
 *   - 行追加 (workbook_table_add_rows) は **テーブル化されたシート** のみ対応。
 *     生のシートにはまず workbook_table_create でテーブル化してから使う。
 *   - 生シートへ直接書き込みたい場合は workbook_range_update を使う。
 */

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
  WorkbookSessionInfo,
  WorkbookWorksheet,
  WorkbookTable,
  WorkbookTableRow,
  WorkbookTableColumn,
  WorkbookRange,
  WorkbookChart,
} from "@/lib/msgraph/types";
import {
  workbookBase,
  encodeWorksheetRef,
  encodeTableRef,
  WorkbookFileLocator,
} from "./workbook-helpers";
import { USER_ID_DESCRIPTION } from "./shared-helpers";

// ============================================================
// 共通 Zod スキーマ（再利用するため変数に切り出し）
// ============================================================

const fileLocatorSchema = {
  item_id: z.string().optional().describe("Excel ファイルの DriveItem ID"),
  path: z
    .string()
    .optional()
    .describe('ファイルパス（例: "/Documents/data.xlsx"）。item_id と排他'),
  drive_id: z
    .string()
    .optional()
    .describe(
      "Drive ID（SharePoint や他人の OneDrive の場合に指定）。省略時は本人 OneDrive"
    ),
  user_id: z.string().optional().describe(USER_ID_DESCRIPTION),
  workbook_session_id: z
    .string()
    .optional()
    .describe(
      "永続セッションID（workbook_create_session で取得）。" +
        "複数の操作をまとめる場合に必須相当。省略時はセッションレス（変更は即保存だが他者反映に最大2分のラグあり）"
    ),
};

/** params から WorkbookFileLocator + workbookSessionId を抽出 */
function locFromParams(params: Record<string, unknown>): {
  loc: WorkbookFileLocator;
  workbookSessionId?: string;
} {
  return {
    loc: {
      item_id: params.item_id as string | undefined,
      path: params.path as string | undefined,
      drive_id: params.drive_id as string | undefined,
      user_id: params.user_id as string | undefined,
    },
    workbookSessionId: params.workbook_session_id as string | undefined,
  };
}

// ============================================================
// メイン登録関数
// ============================================================

export function registerWorkbookTools(server: McpServer): void {
  // ##########################################################
  // A. Sessions (2 tools)
  // ##########################################################

  // ----------------------------------------------------------
  // A1. workbook_create_session
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_create_session",
    {
      title: "Create Workbook Session",
      description: `Excel ブックの永続/非永続セッションを作成する。
セッションを使うと API 呼び出しのパフォーマンスが向上し、変更が他のクライアントに即時反映される。
セッションを使わないと書き込みが他者に反映されるまで最大2分のラグが発生する。

Args:
  - item_id / path (どちらか必須): Excel ファイル指定
  - persist_changes (default: true): true=永続セッション（変更を保存）、false=非永続（一時的な計算用）
  - drive_id / user_id: optional

Returns: { id: string, persistChanges: boolean }
  → この id を以降のツール呼び出しで workbook_session_id として渡す`,
      inputSchema: {
        ...fileLocatorSchema,
        // workbook_session_id はこのツールでは使わないが、スキーマ統一のため許容
        persist_changes: z
          .boolean()
          .default(true)
          .describe("true=永続セッション（変更を保存）、false=非永続"),
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
        const { loc } = locFromParams(params);
        const base = workbookBase(loc);
        const result = await graphPost<WorkbookSessionInfo>(
          `${base}/createSession`,
          { persistChanges: params.persist_changes }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  workbook_session_id: result.id,
                  persistChanges: result.persistChanges,
                  hint:
                    "以降の Workbook ツール呼び出しで workbook_session_id にこの値を渡してください。" +
                    "5分間アイドルで自動破棄されます。明示的に閉じるには workbook_close_session。",
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // A2. workbook_close_session
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_close_session",
    {
      title: "Close Workbook Session",
      description: `Excel ブックのセッションを明示的に閉じる。
通常は呼ばなくても5分のアイドルタイムアウトで自動破棄されるが、即時に解放したい場合に使う。

Args:
  - item_id / path (どちらか必須)
  - workbook_session_id (必須): 閉じるセッションのID
  - drive_id / user_id: optional

Returns: 確認メッセージ`,
      inputSchema: {
        ...fileLocatorSchema,
        workbook_session_id: z
          .string()
          .min(1)
          .describe("閉じるセッションのID（必須）"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        await graphPost(
          `${base}/closeSession`,
          {},
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                message: "セッションを閉じました",
                closed_session_id: workbookSessionId,
              }),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ##########################################################
  // B. Worksheets (5 tools)
  // ##########################################################

  // ----------------------------------------------------------
  // B1. workbook_list_worksheets
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_list_worksheets",
    {
      title: "List Worksheets",
      description: `Excel ブック内のすべてのワークシート（シート）を一覧する。

Args:
  - item_id / path (どちらか必須)
  - drive_id / user_id / workbook_session_id: optional

Returns: シート配列（id, name, position, visibility）`,
      inputSchema: { ...fileLocatorSchema },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const data = await graphGet<GraphPagedResponse<WorkbookWorksheet>>(
          `${base}/worksheets`,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: truncateResponse(
                JSON.stringify(
                  { count: data.value.length, worksheets: data.value },
                  null,
                  2
                )
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // B2. workbook_get_worksheet
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_get_worksheet",
    {
      title: "Get Worksheet",
      description: `特定のワークシートのメタデータを取得する。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - drive_id / user_id / workbook_session_id: optional

Returns: シート詳細（id, name, position, visibility）`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const ws = await graphGet<WorkbookWorksheet>(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}`,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [{ type: "text", text: JSON.stringify(ws, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // B3. workbook_add_worksheet
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_add_worksheet",
    {
      title: "Add Worksheet",
      description: `新しいワークシートをブックに追加する。

Args:
  - item_id / path (どちらか必須)
  - name: 追加するシート名（省略時は自動命名）
  - drive_id / user_id / workbook_session_id: optional

Returns: 作成されたシート情報`,
      inputSchema: {
        ...fileLocatorSchema,
        name: z.string().optional().describe("シート名（省略時は自動命名）"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const body: Record<string, unknown> = {};
        if (params.name) body.name = params.name;
        const ws = await graphPost<WorkbookWorksheet>(
          `${base}/worksheets/add`,
          body,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  success: true,
                  id: ws.id,
                  name: ws.name,
                  position: ws.position,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // B4. workbook_update_worksheet
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_update_worksheet",
    {
      title: "Update Worksheet",
      description: `シート名・位置・可視性を変更する（PATCH）。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - new_name: 新しい名前
  - position: 並び順（0始まり）
  - visibility: "Visible" | "Hidden" | "VeryHidden"
  - drive_id / user_id / workbook_session_id: optional

Returns: 更新後のシート情報`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        new_name: z.string().optional().describe("新しいシート名"),
        position: z
          .number()
          .int()
          .min(0)
          .optional()
          .describe("並び順（0始まり）"),
        visibility: z
          .enum(["Visible", "Hidden", "VeryHidden"])
          .optional()
          .describe("可視性"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const body: Record<string, unknown> = {};
        if (params.new_name !== undefined) body.name = params.new_name;
        if (params.position !== undefined) body.position = params.position;
        if (params.visibility !== undefined) body.visibility = params.visibility;
        const ws = await graphPatch<WorkbookWorksheet>(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}`,
          body,
          { workbookSessionId }
        );
        return {
          content: [{ type: "text", text: JSON.stringify(ws, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // B5. workbook_delete_worksheet
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_delete_worksheet",
    {
      title: "Delete Worksheet",
      description: `指定したワークシートを削除する。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - drive_id / user_id / workbook_session_id: optional

Returns: 確認メッセージ`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        await graphDelete(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}`,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                deleted_worksheet: params.worksheet,
              }),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ##########################################################
  // C. Tables (6 tools)
  // ##########################################################

  // ----------------------------------------------------------
  // C1. workbook_list_tables
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_list_tables",
    {
      title: "List Tables",
      description: `ブック全体、または特定シート内のテーブル一覧を取得する。

Args:
  - item_id / path (どちらか必須)
  - worksheet: シート ID または名前。省略時はブック全体のテーブル一覧
  - drive_id / user_id / workbook_session_id: optional

Returns: テーブル配列（id, name, showHeaders, showTotals 等）`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z
          .string()
          .optional()
          .describe("シート ID または名前。省略時はブック全体"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const endpoint = params.worksheet
          ? `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/tables`
          : `${base}/tables`;
        const data = await graphGet<GraphPagedResponse<WorkbookTable>>(
          endpoint,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: truncateResponse(
                JSON.stringify(
                  { count: data.value.length, tables: data.value },
                  null,
                  2
                )
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // C2. workbook_get_table
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_get_table",
    {
      title: "Get Table",
      description: `特定のテーブルの詳細を取得する。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - drive_id / user_id / workbook_session_id: optional

Returns: テーブル詳細`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const tbl = await graphGet<WorkbookTable>(
          `${base}/tables/${encodeTableRef(params.table)}`,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [{ type: "text", text: JSON.stringify(tbl, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // C3. workbook_create_table
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_create_table",
    {
      title: "Create Table",
      description: `指定シート上の範囲をテーブル化する。
生のシート（テーブル化されていない）に行追加 (workbook_table_add_rows) を使えるようにするための前段として使う。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - address (必須): テーブル化する範囲（例: "A1:D10" または "Sheet1!A1:D10"）
  - has_headers (default: true): 1行目を見出し行として扱うか
  - drive_id / user_id / workbook_session_id: optional

Returns: 作成されたテーブル情報`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        address: z
          .string()
          .min(1)
          .describe('テーブル化する範囲（例: "A1:D10"）'),
        has_headers: z
          .boolean()
          .default(true)
          .describe("1行目を見出し行として扱うか"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const tbl = await graphPost<WorkbookTable>(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/tables/add`,
          {
            address: params.address,
            hasHeaders: params.has_headers,
          },
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                { success: true, id: tbl.id, name: tbl.name },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // C4. workbook_update_table
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_update_table",
    {
      title: "Update Table",
      description: `テーブルの設定を更新する（PATCH）。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - new_name: 新しいテーブル名
  - show_headers: 見出し行を表示するか
  - show_totals: 合計行を表示するか
  - show_banded_rows: 縞模様（行）
  - show_banded_columns: 縞模様（列）
  - show_filter_button: フィルタボタン表示
  - highlight_first_column / highlight_last_column: 端列の強調
  - style: テーブルスタイル名
  - drive_id / user_id / workbook_session_id: optional

Returns: 更新後のテーブル情報`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
        new_name: z.string().optional(),
        show_headers: z.boolean().optional(),
        show_totals: z.boolean().optional(),
        show_banded_rows: z.boolean().optional(),
        show_banded_columns: z.boolean().optional(),
        show_filter_button: z.boolean().optional(),
        highlight_first_column: z.boolean().optional(),
        highlight_last_column: z.boolean().optional(),
        style: z.string().optional(),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const body: Record<string, unknown> = {};
        if (params.new_name !== undefined) body.name = params.new_name;
        if (params.show_headers !== undefined) body.showHeaders = params.show_headers;
        if (params.show_totals !== undefined) body.showTotals = params.show_totals;
        if (params.show_banded_rows !== undefined) body.showBandedRows = params.show_banded_rows;
        if (params.show_banded_columns !== undefined) body.showBandedColumns = params.show_banded_columns;
        if (params.show_filter_button !== undefined) body.showFilterButton = params.show_filter_button;
        if (params.highlight_first_column !== undefined) body.highlightFirstColumn = params.highlight_first_column;
        if (params.highlight_last_column !== undefined) body.highlightLastColumn = params.highlight_last_column;
        if (params.style !== undefined) body.style = params.style;
        const tbl = await graphPatch<WorkbookTable>(
          `${base}/tables/${encodeTableRef(params.table)}`,
          body,
          { workbookSessionId }
        );
        return {
          content: [{ type: "text", text: JSON.stringify(tbl, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // C5. workbook_delete_table
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_delete_table",
    {
      title: "Delete Table",
      description: `テーブルを削除する（テーブル定義のみ削除、セル値は残る）。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - drive_id / user_id / workbook_session_id: optional

Returns: 確認メッセージ`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        await graphDelete(`${base}/tables/${encodeTableRef(params.table)}`, {
          workbookSessionId,
        });
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                deleted_table: params.table,
              }),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // C6. workbook_table_convert_to_range
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_table_convert_to_range",
    {
      title: "Convert Table to Range",
      description: `テーブルをただのセル範囲に戻す（テーブル定義を解除）。
セル値・書式は維持される。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - drive_id / user_id / workbook_session_id: optional

Returns: 変換後の Range 情報`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const result = await graphPost<WorkbookRange>(
          `${base}/tables/${encodeTableRef(params.table)}/convertToRange`,
          {},
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: truncateResponse(JSON.stringify(result, null, 2)),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ##########################################################
  // D. Table Rows (5 tools) ★メイン
  // ##########################################################

  // ----------------------------------------------------------
  // D1. workbook_table_add_rows ★ ユーザー指定の主目的
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_table_add_rows",
    {
      title: "Add Rows to Table",
      description: `テーブルの末尾（または指定位置）に1行以上のレコードを追加する。
**ベストプラクティス: 1行ずつではなく、複数行をまとめて1回のAPI呼び出しで追加すること**（パフォーマンス劣化を避けるため）。

行追加は **テーブル化されたシートのみ** 対応。生のシートに対しては
まず workbook_create_table でテーブル化してから使う。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - values (必須): 2次元配列（外側=行、内側=各セル）
                   例: [["田中", "tanaka@example.com"], ["佐藤", "sato@example.com"]]
  - index: 挿入位置（0始まり、null=末尾）。途中に挿入するとその下が下方向にシフト
  - drive_id / user_id / workbook_session_id: optional

Returns: 作成された行情報（index, values）`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
        values: z
          .array(z.array(z.unknown()))
          .min(1)
          .describe("2次元配列（外側=行、内側=各セル）"),
        index: z
          .number()
          .int()
          .min(0)
          .nullable()
          .optional()
          .describe("挿入位置（0始まり、null/省略=末尾追加）"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const body = {
          index: params.index ?? null,
          values: params.values,
        };
        const result = await graphPost<WorkbookTableRow>(
          `${base}/tables/${encodeTableRef(params.table)}/rows/add`,
          body,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: truncateResponse(
                JSON.stringify(
                  {
                    success: true,
                    inserted_at_index: result.index,
                    rows_added: params.values.length,
                    values: result.values,
                  },
                  null,
                  2
                )
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // D2. workbook_table_list_rows
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_table_list_rows",
    {
      title: "List Table Rows",
      description: `テーブル内の行一覧を取得する。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - top: 最大件数（1-200, default 100）
  - skip: スキップ件数（ページネーション）
  - drive_id / user_id / workbook_session_id: optional

Returns: 行配列（index, values）`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
        top: z.number().int().min(1).max(200).default(100).describe("最大件数"),
        skip: z.number().int().min(0).optional().describe("スキップ件数"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const queryParams: Record<string, string | number | undefined> = {
          $top: params.top,
        };
        if (params.skip !== undefined) queryParams.$skip = params.skip;
        const data = await graphGet<GraphPagedResponse<WorkbookTableRow>>(
          `${base}/tables/${encodeTableRef(params.table)}/rows`,
          queryParams,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: truncateResponse(
                JSON.stringify(
                  { count: data.value.length, rows: data.value },
                  null,
                  2
                )
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // D3. workbook_table_get_row
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_table_get_row",
    {
      title: "Get Table Row",
      description: `テーブルの特定行を index で取得する（0始まり）。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - index (必須): 行のインデックス（0始まり、ヘッダー行は含まない）
  - drive_id / user_id / workbook_session_id: optional

Returns: 行情報（index, values）`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
        index: z.number().int().min(0).describe("行インデックス（0始まり）"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const row = await graphGet<WorkbookTableRow>(
          `${base}/tables/${encodeTableRef(params.table)}/rows/itemAt(index=${params.index})`,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [{ type: "text", text: JSON.stringify(row, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // D4. workbook_table_update_row
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_table_update_row",
    {
      title: "Update Table Row",
      description: `テーブルの特定行の値を更新する（PATCH）。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - index (必須): 行のインデックス（0始まり）
  - values (必須): 1行分の値配列の2次元化（例: [["田中", "更新後"]]）
                   2次元なのは API 仕様で1行のみでも [[..]] にする必要があるため
  - drive_id / user_id / workbook_session_id: optional

Returns: 更新後の行情報`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
        index: z.number().int().min(0).describe("行インデックス"),
        values: z
          .array(z.array(z.unknown()))
          .min(1)
          .max(1)
          .describe("1行分の値（2次元配列で渡す: [[col1, col2, ...]]）"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const row = await graphPatch<WorkbookTableRow>(
          `${base}/tables/${encodeTableRef(params.table)}/rows/itemAt(index=${params.index})`,
          { values: params.values },
          { workbookSessionId }
        );
        return {
          content: [{ type: "text", text: JSON.stringify(row, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // D5. workbook_table_delete_row
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_table_delete_row",
    {
      title: "Delete Table Row",
      description: `テーブルの特定行を削除する。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - index (必須): 削除する行のインデックス（0始まり）
  - drive_id / user_id / workbook_session_id: optional

Returns: 確認メッセージ`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
        index: z.number().int().min(0).describe("削除する行のインデックス"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        await graphDelete(
          `${base}/tables/${encodeTableRef(params.table)}/rows/itemAt(index=${params.index})`,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                deleted_row_index: params.index,
              }),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ##########################################################
  // E. Table Columns (4 tools)
  // ##########################################################

  // ----------------------------------------------------------
  // E1. workbook_table_list_columns
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_table_list_columns",
    {
      title: "List Table Columns",
      description: `テーブルの列一覧を取得する。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - drive_id / user_id / workbook_session_id: optional

Returns: 列配列（id, name, index, values）`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const data = await graphGet<GraphPagedResponse<WorkbookTableColumn>>(
          `${base}/tables/${encodeTableRef(params.table)}/columns`,
          undefined,
          { workbookSessionId }
        );
        // values は重いので各列の最初の数行だけにトリムする
        const columns = data.value.map((c) => ({
          id: c.id,
          name: c.name,
          index: c.index,
        }));
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                { count: columns.length, columns },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // E2. workbook_table_add_column
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_table_add_column",
    {
      title: "Add Table Column",
      description: `テーブルに列を追加する。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - name: 列見出し名（任意、values の最初の要素で代替可能）
  - values: 列のデータ（2次元配列、外側=各行、内側=セル値1個ずつ。先頭が見出し）
            例: [["Status"], ["Open"], ["Closed"]]
  - index: 挿入位置（0始まり、省略時は末尾）
  - drive_id / user_id / workbook_session_id: optional

Returns: 作成された列情報`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
        name: z.string().optional().describe("列見出し名"),
        values: z
          .array(z.array(z.unknown()))
          .optional()
          .describe("列の値（2次元配列、先頭は見出し）"),
        index: z.number().int().min(0).optional().describe("挿入位置（0始まり）"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const body: Record<string, unknown> = {};
        if (params.name !== undefined) body.name = params.name;
        if (params.values !== undefined) body.values = params.values;
        if (params.index !== undefined) body.index = params.index;
        const col = await graphPost<WorkbookTableColumn>(
          `${base}/tables/${encodeTableRef(params.table)}/columns/add`,
          body,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                { success: true, id: col.id, name: col.name, index: col.index },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // E3. workbook_table_update_column
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_table_update_column",
    {
      title: "Update Table Column",
      description: `テーブルの列の名前または値を更新する（PATCH）。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - column (必須): 列 ID または名前
  - new_name: 新しい列名
  - values: 列の値全体（2次元配列、先頭は見出し）
  - drive_id / user_id / workbook_session_id: optional

Returns: 更新後の列情報`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
        column: z.string().min(1).describe("列 ID または名前"),
        new_name: z.string().optional(),
        values: z.array(z.array(z.unknown())).optional(),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const body: Record<string, unknown> = {};
        if (params.new_name !== undefined) body.name = params.new_name;
        if (params.values !== undefined) body.values = params.values;
        const col = await graphPatch<WorkbookTableColumn>(
          `${base}/tables/${encodeTableRef(params.table)}/columns/${encodeURIComponent(params.column)}`,
          body,
          { workbookSessionId }
        );
        return {
          content: [{ type: "text", text: JSON.stringify(col, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // E4. workbook_table_delete_column
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_table_delete_column",
    {
      title: "Delete Table Column",
      description: `テーブルの列を削除する。

Args:
  - item_id / path (どちらか必須)
  - table (必須): テーブル ID または名前
  - column (必須): 列 ID または名前
  - drive_id / user_id / workbook_session_id: optional

Returns: 確認メッセージ`,
      inputSchema: {
        ...fileLocatorSchema,
        table: z.string().min(1).describe("テーブル ID または名前"),
        column: z.string().min(1).describe("列 ID または名前"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        await graphDelete(
          `${base}/tables/${encodeTableRef(params.table)}/columns/${encodeURIComponent(params.column)}`,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                deleted_column: params.column,
              }),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ##########################################################
  // F. Range (8 tools) ★ 生シートの読み書き
  // ##########################################################

  // ----------------------------------------------------------
  // F1. workbook_range_get
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_range_get",
    {
      title: "Get Range Values",
      description: `指定した範囲のセル値・数式・書式を取得する。テーブル化されていない生のシートに対して使える。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - address (必須): A1 形式の範囲（例: "A1:C10"）
  - drive_id / user_id / workbook_session_id: optional

Returns: Range 情報（values, text, formulas, numberFormat 等）`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        address: z.string().min(1).describe('範囲（例: "A1:C10"）'),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const range = await graphGet<WorkbookRange>(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/range(address='${encodeURIComponent(params.address)}')`,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: truncateResponse(JSON.stringify(range, null, 2)),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // F2. workbook_range_update
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_range_update",
    {
      title: "Update Range Values",
      description: `指定範囲のセル値・数式・書式を更新する（PATCH）。テーブル化されていない生のシートに対しても直接書き込める。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - address (必須): A1 形式の範囲（例: "A1:C3"）。values の次元と一致させること
  - values: 2次元配列（外側=行、内側=列）。null を渡したセルは更新されない
  - formulas: 数式の2次元配列（"=SUM(A1:A10)" 等）
  - number_format: 数値書式の2次元配列（"0.00", "yyyy/m/d" 等）
  - drive_id / user_id / workbook_session_id: optional

Returns: 更新後の Range 情報`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        address: z.string().min(1).describe('範囲（例: "A1:C3"）'),
        values: z
          .array(z.array(z.unknown()))
          .optional()
          .describe("値の2次元配列"),
        formulas: z
          .array(z.array(z.unknown()))
          .optional()
          .describe("数式の2次元配列"),
        number_format: z
          .array(z.array(z.unknown()))
          .optional()
          .describe("数値書式の2次元配列"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const body: Record<string, unknown> = {};
        if (params.values !== undefined) body.values = params.values;
        if (params.formulas !== undefined) body.formulas = params.formulas;
        if (params.number_format !== undefined)
          body.numberFormat = params.number_format;
        const range = await graphPatch<WorkbookRange>(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/range(address='${encodeURIComponent(params.address)}')`,
          body,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: truncateResponse(
                JSON.stringify(
                  {
                    success: true,
                    address: range.address,
                    rowCount: range.rowCount,
                    columnCount: range.columnCount,
                    values: range.values,
                  },
                  null,
                  2
                )
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // F3. workbook_range_clear
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_range_clear",
    {
      title: "Clear Range",
      description: `指定範囲をクリア（値削除、書式削除など）。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - address (必須): A1 形式の範囲
  - apply_to: "All" | "Formats" | "Contents" | "Hyperlinks" | "RemoveHyperlinks"（default "All"）
  - drive_id / user_id / workbook_session_id: optional

Returns: 確認メッセージ`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        address: z.string().min(1).describe("範囲"),
        apply_to: z
          .enum(["All", "Formats", "Contents", "Hyperlinks", "RemoveHyperlinks"])
          .default("All")
          .describe("クリア対象"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        await graphPost(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/range(address='${encodeURIComponent(params.address)}')/clear`,
          { applyTo: params.apply_to },
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                cleared: params.address,
                applyTo: params.apply_to,
              }),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // F4. workbook_range_get_used
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_range_get_used",
    {
      title: "Get Used Range",
      description: `シートの使用範囲（usedRange）を取得する。
データが入っているセルの最小バウンディングボックス。
シート全体の内容を一気に取りたい時に便利。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - values_only (default: false): true=空セル除外して値だけ
  - drive_id / user_id / workbook_session_id: optional

Returns: Range 情報（address, rowCount, columnCount, values 等）`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        values_only: z.boolean().default(false).describe("空セル除外して値だけ"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const fn = params.values_only ? "usedRange(valuesOnly=true)" : "usedRange";
        const range = await graphGet<WorkbookRange>(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/${fn}`,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: truncateResponse(JSON.stringify(range, null, 2)),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // F5. workbook_range_insert
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_range_insert",
    {
      title: "Insert Cells (Shift)",
      description: `指定範囲にセルを挿入し、既存のセルをシフトする。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - address (必須): 挿入位置の範囲
  - shift (必須): "Down" | "Right"（既存セルのシフト方向）
  - drive_id / user_id / workbook_session_id: optional

Returns: 挿入後の Range 情報`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        address: z.string().min(1).describe("挿入位置の範囲"),
        shift: z.enum(["Down", "Right"]).describe("既存セルのシフト方向"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const range = await graphPost<WorkbookRange>(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/range(address='${encodeURIComponent(params.address)}')/insert`,
          { shift: params.shift },
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: truncateResponse(JSON.stringify(range, null, 2)),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // F6. workbook_range_delete
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_range_delete",
    {
      title: "Delete Cells (Shift)",
      description: `指定範囲のセルを削除し、残りのセルをシフトする。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - address (必須): 削除する範囲
  - shift (必須): "Up" | "Left"（残りセルのシフト方向）
  - drive_id / user_id / workbook_session_id: optional

Returns: 確認メッセージ`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        address: z.string().min(1).describe("削除する範囲"),
        shift: z.enum(["Up", "Left"]).describe("残りセルのシフト方向"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        await graphPost(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/range(address='${encodeURIComponent(params.address)}')/delete`,
          { shift: params.shift },
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                deleted: params.address,
                shift: params.shift,
              }),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // F7. workbook_range_merge
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_range_merge",
    {
      title: "Merge Range",
      description: `指定範囲のセルを結合する。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - address (必須): 結合する範囲
  - across (default: false): true=各行ごとに結合、false=全体を1セルに結合
  - drive_id / user_id / workbook_session_id: optional

Returns: 確認メッセージ`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        address: z.string().min(1).describe("結合する範囲"),
        across: z
          .boolean()
          .default(false)
          .describe("true=各行ごとに結合、false=全体を1セルに"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const across = (params.across as boolean | undefined) ?? false;
        await graphPost(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/range(address='${encodeURIComponent(params.address)}')/merge`,
          { across },
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                merged: params.address,
                across,
              }),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // F8. workbook_range_unmerge
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_range_unmerge",
    {
      title: "Unmerge Range",
      description: `指定範囲のセル結合を解除する。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - address (必須): 結合解除する範囲
  - drive_id / user_id / workbook_session_id: optional

Returns: 確認メッセージ`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        address: z.string().min(1).describe("結合解除する範囲"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        await graphPost(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/range(address='${encodeURIComponent(params.address)}')/unmerge`,
          {},
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                unmerged: params.address,
              }),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ##########################################################
  // G. Functions (1 tool)
  // ##########################################################

  // ----------------------------------------------------------
  // G1. workbook_call_function
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_call_function",
    {
      title: "Call Excel Function",
      description: `Excel の組み込み関数を任意に呼び出す（VLOOKUP, SUM, IF, INDEX, MATCH 等の300以上の関数）。
Excel の強力な計算エンジンを API から直接利用できる。

Args:
  - item_id / path (どちらか必須)
  - function_name (必須): 関数名（"vlookup", "sum", "pmt" 等、小文字でも大文字でも可）
  - arguments (必須): 関数の引数オブジェクト（関数ごとに異なる）
                      例 vlookup: { lookupValue: "pear", tableArray: { Address: "Sheet1!B2:C7" }, colIndexNum: 2, rangeLookup: false }
                      例 sum:     { values: [1, 2, 3, 4, 5] }
                      例 pmt:     { rate: 0.005, nper: 360, pv: -200000 }
  - drive_id / user_id / workbook_session_id: optional

Returns: 関数の実行結果（{ value, error? }）`,
      inputSchema: {
        ...fileLocatorSchema,
        function_name: z
          .string()
          .min(1)
          .describe('Excel 関数名（例: "vlookup", "sum", "pmt"）'),
        arguments: z
          .record(z.unknown())
          .describe("関数の引数オブジェクト（関数ごとに異なる）"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const fnName = params.function_name.toLowerCase();
        const result = await graphPost<{ value: unknown; error?: string }>(
          `${base}/functions/${encodeURIComponent(fnName)}`,
          params.arguments,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ##########################################################
  // H. Charts (4 tools)
  // ##########################################################

  // ----------------------------------------------------------
  // H1. workbook_list_charts
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_list_charts",
    {
      title: "List Charts",
      description: `シート上のチャート（グラフ）一覧を取得する。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - drive_id / user_id / workbook_session_id: optional

Returns: チャート配列（id, name, height, width, top, left）`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const data = await graphGet<GraphPagedResponse<WorkbookChart>>(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/charts`,
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: truncateResponse(
                JSON.stringify(
                  { count: data.value.length, charts: data.value },
                  null,
                  2
                )
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // H2. workbook_create_chart
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_create_chart",
    {
      title: "Create Chart",
      description: `指定範囲のデータを元にチャートを作成する。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - chart_type (必須): "ColumnClustered", "ColumnStacked", "Line", "Pie", "Bar",
                       "Area", "Scatter", "Doughnut" 等
  - source_data (必須): データ範囲（例: "A1:C10"）
  - series_by (default: "Auto"): "Auto" | "Columns" | "Rows"（系列の方向）
  - drive_id / user_id / workbook_session_id: optional

Returns: 作成されたチャート情報`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        chart_type: z
          .string()
          .min(1)
          .describe('チャート種別（例: "ColumnClustered", "Line", "Pie"）'),
        source_data: z.string().min(1).describe('データ範囲（例: "A1:C10"）'),
        series_by: z
          .enum(["Auto", "Columns", "Rows"])
          .default("Auto")
          .describe("系列の方向"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const seriesBy = (params.series_by as string | undefined) ?? "Auto";
        const chart = await graphPost<WorkbookChart>(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/charts/add`,
          {
            type: params.chart_type,
            sourceData: params.source_data,
            seriesBy,
          },
          undefined,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                { success: true, id: chart.id, name: chart.name },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // H3. workbook_get_chart_image
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_get_chart_image",
    {
      title: "Get Chart Image",
      description: `チャートの画像を base64 PNG で取得する。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - chart (必須): チャート ID または名前
  - width: 画像の幅（pixel、optional）
  - height: 画像の高さ（pixel、optional）
  - fitting_mode (default: "Fit"): "Fit" | "FitAndCenter" | "Fill"
  - drive_id / user_id / workbook_session_id: optional

Returns: { value: "base64-encoded-png" }`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        chart: z.string().min(1).describe("チャート ID または名前"),
        width: z.number().int().min(1).optional().describe("画像幅（pixel）"),
        height: z.number().int().min(1).optional().describe("画像高（pixel）"),
        fitting_mode: z
          .enum(["Fit", "FitAndCenter", "Fill"])
          .default("Fit")
          .describe("フィッティングモード"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        const queryParams: Record<string, string | number> = {};
        if (params.width !== undefined) queryParams.width = params.width as number;
        if (params.height !== undefined) queryParams.height = params.height as number;
        queryParams.fittingMode = (params.fitting_mode as string | undefined) ?? "Fit";
        const result = await graphGet<{ value: string }>(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/charts/${encodeURIComponent(params.chart)}/image`,
          queryParams,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(
                {
                  chart: params.chart,
                  format: "base64-png",
                  size_bytes: result.value?.length || 0,
                  // base64 が長いとレスポンスを肥大化させるので先頭を表示
                  base64_preview: result.value
                    ? result.value.slice(0, 100) + "..."
                    : null,
                  base64: result.value,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // ----------------------------------------------------------
  // H4. workbook_delete_chart
  // ----------------------------------------------------------
  server.registerTool(
    "workbook_delete_chart",
    {
      title: "Delete Chart",
      description: `チャートを削除する。

Args:
  - item_id / path (どちらか必須)
  - worksheet (必須): シート ID または名前
  - chart (必須): チャート ID または名前
  - drive_id / user_id / workbook_session_id: optional

Returns: 確認メッセージ`,
      inputSchema: {
        ...fileLocatorSchema,
        worksheet: z.string().min(1).describe("シート ID または名前"),
        chart: z.string().min(1).describe("チャート ID または名前"),
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
        const { loc, workbookSessionId } = locFromParams(params);
        const base = workbookBase(loc);
        await graphDelete(
          `${base}/worksheets/${encodeWorksheetRef(params.worksheet)}/charts/${encodeURIComponent(params.chart)}`,
          { workbookSessionId }
        );
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                success: true,
                deleted_chart: params.chart,
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
