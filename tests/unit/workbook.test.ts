/**
 * lib/mcp/tools/workbook.ts の単体テスト
 *
 * registerWorkbookTools() で35ツールが登録されること、
 * 各ツールが正しいエンドポイントを叩くこと、
 * workbook_session_id がヘッダーとして渡ること、
 * エラーハンドリングが効くこと、を検証する。
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { clearMockRedis } from "../setup";
import { saveRefreshToken } from "@/lib/redis/token-store";
import { clearAccessTokenCache } from "@/lib/msgraph/auth-context";
import { registerWorkbookTools } from "@/lib/mcp/tools/workbook";

// ── McpServer モック ──
// registerTool() の呼び出しを記録し、後でハンドラを取り出して直接呼ぶ
interface CapturedTool {
  name: string;
  config: {
    title: string;
    description: string;
    inputSchema: Record<string, unknown>;
    annotations?: Record<string, unknown>;
  };
  handler: (params: Record<string, unknown>) => Promise<{
    content: Array<{ type: string; text: string }>;
    isError?: boolean;
  }>;
}

function createMockServer(): {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  server: { registerTool: any };
  captured: CapturedTool[];
} {
  const captured: CapturedTool[] = [];
  const server = {
    registerTool: (
      name: string,
      config: CapturedTool["config"],
      handler: CapturedTool["handler"]
    ) => {
      captured.push({ name, config, handler });
    },
  };
  return { server, captured };
}

/** access_token を取得する fetch 呼び出しをモックし、その後 Graph API 呼び出しを mockResolvedValueOnce する */
function setupFetchMock(
  ...graphResponses: Array<{ ok: boolean; status: number; data: unknown; headers?: HeadersInit }>
) {
  const mockFetch = vi.fn();
  // まず token endpoint
  mockFetch.mockResolvedValueOnce({
    ok: true,
    json: () =>
      Promise.resolve({
        access_token: "at_test",
        refresh_token: "rt_new",
        expires_in: 3600,
      }),
  });
  // その後 Graph API レスポンス群
  for (const r of graphResponses) {
    mockFetch.mockResolvedValueOnce({
      ok: r.ok,
      status: r.status,
      statusText: r.ok ? "OK" : "Error",
      headers: new Headers(r.headers || {}),
      json: () => Promise.resolve(r.data),
    });
  }
  vi.stubGlobal("fetch", mockFetch);
  return mockFetch;
}

// ============================================================
// 登録テスト（35ツールすべて存在することを確認）
// ============================================================

describe("registerWorkbookTools: 全35ツール登録", () => {
  it("WT01: 35個のツールがすべて登録される", () => {
    const { server, captured } = createMockServer();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    registerWorkbookTools(server as any);
    expect(captured).toHaveLength(35);
  });

  it("WT02: 期待される名前すべてが揃う", () => {
    const { server, captured } = createMockServer();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    registerWorkbookTools(server as any);
    const names = captured.map((t) => t.name).sort();
    const expected = [
      // Sessions (2)
      "workbook_create_session",
      "workbook_close_session",
      // Worksheets (5)
      "workbook_list_worksheets",
      "workbook_get_worksheet",
      "workbook_add_worksheet",
      "workbook_update_worksheet",
      "workbook_delete_worksheet",
      // Tables (6)
      "workbook_list_tables",
      "workbook_get_table",
      "workbook_create_table",
      "workbook_update_table",
      "workbook_delete_table",
      "workbook_table_convert_to_range",
      // Table Rows (5)
      "workbook_table_add_rows",
      "workbook_table_list_rows",
      "workbook_table_get_row",
      "workbook_table_update_row",
      "workbook_table_delete_row",
      // Table Columns (4)
      "workbook_table_list_columns",
      "workbook_table_add_column",
      "workbook_table_update_column",
      "workbook_table_delete_column",
      // Range (8)
      "workbook_range_get",
      "workbook_range_update",
      "workbook_range_clear",
      "workbook_range_get_used",
      "workbook_range_insert",
      "workbook_range_delete",
      "workbook_range_merge",
      "workbook_range_unmerge",
      // Functions (1)
      "workbook_call_function",
      // Charts (4)
      "workbook_list_charts",
      "workbook_create_chart",
      "workbook_get_chart_image",
      "workbook_delete_chart",
    ].sort();
    expect(names).toEqual(expected);
  });

  it("WT03: 全ツールに共通入力（item_id/path/drive_id/user_id/workbook_session_id）が含まれる", () => {
    const { server, captured } = createMockServer();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    registerWorkbookTools(server as any);
    for (const tool of captured) {
      expect(tool.config.inputSchema).toHaveProperty("item_id");
      expect(tool.config.inputSchema).toHaveProperty("path");
      expect(tool.config.inputSchema).toHaveProperty("drive_id");
      expect(tool.config.inputSchema).toHaveProperty("user_id");
      // workbook_session_id は全ツールに存在（close_session では required）
      expect(tool.config.inputSchema).toHaveProperty("workbook_session_id");
    }
  });
});

// ============================================================
// 主要ツールの動作テスト（カテゴリごとに代表）
// ============================================================

describe("Workbook Tools: 各カテゴリの代表ツール動作", () => {
  let captured: CapturedTool[];

  beforeEach(async () => {
    clearMockRedis();
    clearAccessTokenCache();
    vi.restoreAllMocks();
    await saveRefreshToken("rt_test");

    const mock = createMockServer();
    captured = mock.captured;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    registerWorkbookTools(mock.server as any);
  });

  function getTool(name: string): CapturedTool {
    const tool = captured.find((t) => t.name === name);
    if (!tool) throw new Error(`Tool not found: ${name}`);
    return tool;
  }

  // ── A. Sessions ──

  it("A1-T1: workbook_create_session が POST /workbook/createSession を叩く", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 201,
      data: { id: "sess-abc-123", persistChanges: true },
    });
    const tool = getTool("workbook_create_session");
    const result = await tool.handler({
      item_id: "01XXXXX",
      persist_changes: true,
    });

    const url = fetchMock.mock.calls[1][0] as string;
    const body = JSON.parse(fetchMock.mock.calls[1][1].body as string);
    expect(url).toContain("/me/drive/items/01XXXXX/workbook/createSession");
    expect(fetchMock.mock.calls[1][1].method).toBe("POST");
    expect(body.persistChanges).toBe(true);
    const text = result.content[0].text;
    expect(text).toContain("sess-abc-123");
  });

  it("A2-T1: workbook_close_session が POST /workbook/closeSession + session-id ヘッダー", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 204,
      data: {},
    });
    const tool = getTool("workbook_close_session");
    await tool.handler({
      item_id: "01XXXXX",
      workbook_session_id: "sess-abc-123",
    });

    const url = fetchMock.mock.calls[1][0] as string;
    const headers = fetchMock.mock.calls[1][1].headers as Record<string, string>;
    expect(url).toContain("/workbook/closeSession");
    expect(headers["workbook-session-id"]).toBe("sess-abc-123");
  });

  // ── B. Worksheets ──

  it("B1-T1: workbook_list_worksheets が GET /workbook/worksheets", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { value: [{ id: "ws1", name: "Sheet1", position: 0 }] },
    });
    const tool = getTool("workbook_list_worksheets");
    const result = await tool.handler({ item_id: "01XXXXX" });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain("/me/drive/items/01XXXXX/workbook/worksheets");
    expect(fetchMock.mock.calls[1][1].method ?? "GET").toBe("GET");
    expect(result.content[0].text).toContain("Sheet1");
  });

  it("B3-T1: workbook_add_worksheet が POST /worksheets/add", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 201,
      data: { id: "ws-new", name: "新規シート", position: 1 },
    });
    const tool = getTool("workbook_add_worksheet");
    const result = await tool.handler({
      item_id: "01XXXXX",
      name: "新規シート",
    });

    const url = fetchMock.mock.calls[1][0] as string;
    const body = JSON.parse(fetchMock.mock.calls[1][1].body as string);
    expect(url).toContain("/workbook/worksheets/add");
    expect(body.name).toBe("新規シート");
    expect(result.content[0].text).toContain("ws-new");
  });

  // ── C. Tables ──

  it("C3-T1: workbook_create_table が POST /worksheets/{name}/tables/add", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 201,
      data: { id: "T1", name: "Table1" },
    });
    const tool = getTool("workbook_create_table");
    await tool.handler({
      item_id: "01XXXXX",
      worksheet: "Sheet1",
      address: "A1:D10",
      has_headers: true,
    });

    const url = fetchMock.mock.calls[1][0] as string;
    const body = JSON.parse(fetchMock.mock.calls[1][1].body as string);
    expect(url).toContain("/worksheets/Sheet1/tables/add");
    expect(body.address).toBe("A1:D10");
    expect(body.hasHeaders).toBe(true);
  });

  it("C1-T1: workbook_list_tables (worksheet 省略時はブック全体)", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { value: [{ id: "T1", name: "Table1" }] },
    });
    const tool = getTool("workbook_list_tables");
    await tool.handler({ item_id: "01XXXXX" });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toMatch(/\/workbook\/tables(\?|$)/);
  });

  it("C1-T2: workbook_list_tables (worksheet 指定時はそのシートのみ)", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { value: [] },
    });
    const tool = getTool("workbook_list_tables");
    await tool.handler({ item_id: "01XXXXX", worksheet: "Sheet1" });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain("/worksheets/Sheet1/tables");
  });

  // ── D. Table Rows (★ メイン) ──

  it("D1-T1: workbook_table_add_rows が POST /tables/{name}/rows/add", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 201,
      data: { index: 5, values: [["田中", "tanaka@example.com"]] },
    });
    const tool = getTool("workbook_table_add_rows");
    const result = await tool.handler({
      item_id: "01XXXXX",
      table: "Table1",
      values: [["田中", "tanaka@example.com"]],
    });

    const url = fetchMock.mock.calls[1][0] as string;
    const body = JSON.parse(fetchMock.mock.calls[1][1].body as string);
    expect(url).toContain("/workbook/tables/Table1/rows/add");
    expect(body.values).toEqual([["田中", "tanaka@example.com"]]);
    expect(body.index).toBeNull(); // 省略時は末尾追加（null）
    expect(fetchMock.mock.calls[1][1].method).toBe("POST");
    expect(result.content[0].text).toContain("rows_added");
  });

  it("D1-T2: workbook_table_add_rows 複数行を一括追加", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 201,
      data: { index: 10, values: [[1, 2], [3, 4], [5, 6]] },
    });
    const tool = getTool("workbook_table_add_rows");
    await tool.handler({
      item_id: "01XXXXX",
      table: "Table1",
      values: [
        [1, 2],
        [3, 4],
        [5, 6],
      ],
    });

    const body = JSON.parse(fetchMock.mock.calls[1][1].body as string);
    expect(body.values).toHaveLength(3);
  });

  it("D1-T3: workbook_table_add_rows index 指定で途中挿入", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 201,
      data: { index: 2, values: [["挿入", "値"]] },
    });
    const tool = getTool("workbook_table_add_rows");
    await tool.handler({
      item_id: "01XXXXX",
      table: "Table1",
      values: [["挿入", "値"]],
      index: 2,
    });

    const body = JSON.parse(fetchMock.mock.calls[1][1].body as string);
    expect(body.index).toBe(2);
  });

  it("D1-T4: workbook_table_add_rows + workbook_session_id でヘッダー付与", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 201,
      data: { index: 0, values: [[1, 2]] },
    });
    const tool = getTool("workbook_table_add_rows");
    await tool.handler({
      item_id: "01XXXXX",
      table: "Table1",
      values: [[1, 2]],
      workbook_session_id: "sess-write-fast",
    });

    const headers = fetchMock.mock.calls[1][1].headers as Record<string, string>;
    expect(headers["workbook-session-id"]).toBe("sess-write-fast");
  });

  it("D1-T5: workbook_table_add_rows path 指定（item_id の代わり）", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 201,
      data: { index: 0, values: [[1]] },
    });
    const tool = getTool("workbook_table_add_rows");
    await tool.handler({
      path: "/Documents/data.xlsx",
      table: "Table1",
      values: [[1]],
    });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain(
      "/me/drive/root:/Documents/data.xlsx:/workbook/tables/Table1/rows/add"
    );
  });

  it("D1-T6: workbook_table_add_rows SharePoint (drive_id 指定)", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 201,
      data: { index: 0, values: [[1]] },
    });
    const tool = getTool("workbook_table_add_rows");
    await tool.handler({
      drive_id: "b!sharepoint-drive",
      item_id: "01ABC",
      table: "Sales",
      values: [[1]],
    });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain(
      "/drives/b!sharepoint-drive/items/01ABC/workbook/tables/Sales/rows/add"
    );
  });

  it("D2-T1: workbook_table_list_rows ($top, $skip)", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { value: [] },
    });
    const tool = getTool("workbook_table_list_rows");
    await tool.handler({
      item_id: "01XXXXX",
      table: "Table1",
      top: 50,
      skip: 100,
    });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain("/tables/Table1/rows");
    expect(url).toContain("%24top=50");
    expect(url).toContain("%24skip=100");
  });

  it("D3-T1: workbook_table_get_row が itemAt(index=N) を使う", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { index: 3, values: [["a", "b"]] },
    });
    const tool = getTool("workbook_table_get_row");
    await tool.handler({
      item_id: "01XXXXX",
      table: "Table1",
      index: 3,
    });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain("/rows/itemAt(index=3)");
  });

  it("D5-T1: workbook_table_delete_row", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 204,
      data: {},
    });
    const tool = getTool("workbook_table_delete_row");
    await tool.handler({
      item_id: "01XXXXX",
      table: "Table1",
      index: 5,
    });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain("/rows/itemAt(index=5)");
    expect(fetchMock.mock.calls[1][1].method).toBe("DELETE");
  });

  // ── E. Table Columns ──

  it("E2-T1: workbook_table_add_column が POST /columns/add", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 201,
      data: { id: "col-new", name: "Status", index: 3 },
    });
    const tool = getTool("workbook_table_add_column");
    await tool.handler({
      item_id: "01XXXXX",
      table: "Table1",
      name: "Status",
      values: [["Status"], ["Open"], ["Closed"]],
      index: 2,
    });

    const url = fetchMock.mock.calls[1][0] as string;
    const body = JSON.parse(fetchMock.mock.calls[1][1].body as string);
    expect(url).toContain("/tables/Table1/columns/add");
    expect(body.name).toBe("Status");
    expect(body.values).toHaveLength(3);
    expect(body.index).toBe(2);
  });

  // ── F. Range（生シートの読み書き） ──

  it("F1-T1: workbook_range_get が range(address='A1:C10')", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { address: "Sheet1!A1:C10", rowCount: 10, columnCount: 3, values: [] },
    });
    const tool = getTool("workbook_range_get");
    await tool.handler({
      item_id: "01XXXXX",
      worksheet: "Sheet1",
      address: "A1:C10",
    });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain("/worksheets/Sheet1/range(address='A1%3AC10')");
  });

  it("F2-T1: workbook_range_update PATCH で values 渡し", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { address: "Sheet1!A1:B2", rowCount: 2, columnCount: 2, values: [[1, 2], [3, 4]] },
    });
    const tool = getTool("workbook_range_update");
    await tool.handler({
      item_id: "01XXXXX",
      worksheet: "Sheet1",
      address: "A1:B2",
      values: [
        [1, 2],
        [3, 4],
      ],
    });

    const body = JSON.parse(fetchMock.mock.calls[1][1].body as string);
    expect(fetchMock.mock.calls[1][1].method).toBe("PATCH");
    expect(body.values).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  it("F2-T2: workbook_range_update formulas + number_format 同時指定", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { address: "A1:A2", rowCount: 2, columnCount: 1, values: [] },
    });
    const tool = getTool("workbook_range_update");
    await tool.handler({
      item_id: "01XXXXX",
      worksheet: "Sheet1",
      address: "A1:A2",
      formulas: [["=SUM(B1:B10)"], ["=AVERAGE(B1:B10)"]],
      number_format: [["#,##0"], ["0.00"]],
    });

    const body = JSON.parse(fetchMock.mock.calls[1][1].body as string);
    expect(body.formulas).toEqual([["=SUM(B1:B10)"], ["=AVERAGE(B1:B10)"]]);
    expect(body.numberFormat).toEqual([["#,##0"], ["0.00"]]);
  });

  it("F4-T1: workbook_range_get_used (valuesOnly=false)", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { address: "A1:Z100", rowCount: 100, columnCount: 26, values: [] },
    });
    const tool = getTool("workbook_range_get_used");
    await tool.handler({
      item_id: "01XXXXX",
      worksheet: "Sheet1",
    });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain("/worksheets/Sheet1/usedRange");
    expect(url).not.toContain("valuesOnly");
  });

  it("F4-T2: workbook_range_get_used (values_only=true)", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { address: "A1:C5", rowCount: 5, columnCount: 3, values: [] },
    });
    const tool = getTool("workbook_range_get_used");
    await tool.handler({
      item_id: "01XXXXX",
      worksheet: "Sheet1",
      values_only: true,
    });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain("usedRange(valuesOnly=true)");
  });

  it("F7-T1: workbook_range_merge across=false (全体結合)", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: {},
    });
    const tool = getTool("workbook_range_merge");
    await tool.handler({
      item_id: "01XXXXX",
      worksheet: "Sheet1",
      address: "A1:C1",
    });

    const body = JSON.parse(fetchMock.mock.calls[1][1].body as string);
    expect(body.across).toBe(false);
  });

  // ── G. Functions ──

  it("G1-T1: workbook_call_function vlookup", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { value: 42 },
    });
    const tool = getTool("workbook_call_function");
    const result = await tool.handler({
      item_id: "01XXXXX",
      function_name: "vlookup",
      arguments: {
        lookupValue: "pear",
        tableArray: { Address: "Sheet1!B2:C7" },
        colIndexNum: 2,
        rangeLookup: false,
      },
    });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain("/workbook/functions/vlookup");
    expect(result.content[0].text).toContain("42");
  });

  it("G1-T2: workbook_call_function 大文字でも小文字化される", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { value: 100 },
    });
    const tool = getTool("workbook_call_function");
    await tool.handler({
      item_id: "01XXXXX",
      function_name: "PMT",
      arguments: { rate: 0.005, nper: 360, pv: -200000 },
    });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain("/functions/pmt");
    expect(url).not.toContain("/PMT");
  });

  // ── H. Charts ──

  it("H2-T1: workbook_create_chart", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 201,
      data: { id: "ch1", name: "Chart 1" },
    });
    const tool = getTool("workbook_create_chart");
    await tool.handler({
      item_id: "01XXXXX",
      worksheet: "Sheet1",
      chart_type: "ColumnClustered",
      source_data: "A1:C5",
    });

    const url = fetchMock.mock.calls[1][0] as string;
    const body = JSON.parse(fetchMock.mock.calls[1][1].body as string);
    expect(url).toContain("/worksheets/Sheet1/charts/add");
    expect(body.type).toBe("ColumnClustered");
    expect(body.sourceData).toBe("A1:C5");
    expect(body.seriesBy).toBe("Auto");
  });

  it("H3-T1: workbook_get_chart_image (width/height/fittingMode をクエリに含める)", async () => {
    const fetchMock = setupFetchMock({
      ok: true,
      status: 200,
      data: { value: "iVBORw0KGgoAAAANS..." },
    });
    const tool = getTool("workbook_get_chart_image");
    await tool.handler({
      item_id: "01XXXXX",
      worksheet: "Sheet1",
      chart: "Chart 1",
      width: 800,
      height: 600,
    });

    const url = fetchMock.mock.calls[1][0] as string;
    expect(url).toContain("/worksheets/Sheet1/charts/Chart%201/image");
    expect(url).toContain("width=800");
    expect(url).toContain("height=600");
    expect(url).toContain("fittingMode=Fit");
  });
});

// ============================================================
// エラーハンドリング
// ============================================================

describe("Workbook Tools: エラーハンドリング", () => {
  let captured: CapturedTool[];

  beforeEach(async () => {
    clearMockRedis();
    clearAccessTokenCache();
    vi.restoreAllMocks();
    await saveRefreshToken("rt_test");

    const mock = createMockServer();
    captured = mock.captured;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    registerWorkbookTools(mock.server as any);
  });

  it("E01: ファイル指定なしでエラーメッセージ", async () => {
    const tool = captured.find((t) => t.name === "workbook_list_tables")!;
    const result = await tool.handler({});
    expect(result.content[0].text).toContain("item_id");
  });

  it("E02: Graph API 404 がエラーメッセージとして返る", async () => {
    setupFetchMock({
      ok: false,
      status: 404,
      data: { error: { code: "ItemNotFound", message: "Excel ファイルがありません" } },
    });
    const tool = captured.find((t) => t.name === "workbook_list_worksheets")!;
    const result = await tool.handler({ item_id: "nonexistent" });
    expect(result.content[0].text).toContain("ItemNotFound");
  });

  it("E03: Graph API 400 BadRequest（テーブルじゃないシートに rows/add）", async () => {
    setupFetchMock({
      ok: false,
      status: 400,
      data: {
        error: {
          code: "BadRequest",
          message: "Resource not found for the segment 'rows'.",
        },
      },
    });
    const tool = captured.find((t) => t.name === "workbook_table_add_rows")!;
    const result = await tool.handler({
      item_id: "01XXXXX",
      table: "NotATable",
      values: [["x"]],
    });
    expect(result.content[0].text).toContain("BadRequest");
  });
});
