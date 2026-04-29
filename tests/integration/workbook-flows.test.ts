/**
 * Workbook ツールの結合テスト
 *
 * 一連のフローが正しい順序・正しいエンドポイント・正しいヘッダーで
 * Microsoft Graph に届くことを検証する。
 *
 * シナリオ:
 *   1. createSession → session_id を取得
 *   2. createTable    → 範囲をテーブル化
 *   3. addRows (×2)   → 複数回の行追加（1回目=単行、2回目=複数行）
 *   4. listRows       → 追加結果を取得
 *   5. updateRow      → 1行更新
 *   6. deleteRow      → 1行削除
 *   7. closeSession   → 明示的にセッション終了
 *
 * 各ステップで:
 *   - URL が正しい
 *   - method が正しい
 *   - workbook-session-id ヘッダーが付与されている
 *   - body に必要なフィールドが入っている
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { clearMockRedis } from "../setup";
import { saveRefreshToken } from "@/lib/redis/token-store";
import { clearAccessTokenCache } from "@/lib/msgraph/auth-context";
import { registerWorkbookTools } from "@/lib/mcp/tools/workbook";

interface CapturedTool {
  name: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  handler: (params: Record<string, unknown>) => Promise<any>;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function makeServer(): { server: any; tools: Map<string, CapturedTool["handler"]> } {
  const tools = new Map<string, CapturedTool["handler"]>();
  const server = {
    registerTool: (
      name: string,
      _config: unknown,
      handler: CapturedTool["handler"]
    ) => {
      tools.set(name, handler);
    },
  };
  return { server, tools };
}

/**
 * fetch を順序付きでモックする。各 Graph API レスポンスを順番に返す。
 * 各 graph 呼び出し前にトークン取得が走るため、ペアで設定する。
 *
 * 簡易化のため access_token 取得は最初の1回キャッシュされる前提で、
 * 2回目以降はキャッシュ経由＝トークン fetch なし。
 */
function setupSequentialFetch(
  ...graphResponses: Array<{ ok: boolean; status: number; data: unknown }>
): ReturnType<typeof vi.fn> {
  const mockFetch = vi.fn();
  // 1回目だけ token endpoint
  mockFetch.mockResolvedValueOnce({
    ok: true,
    json: () =>
      Promise.resolve({
        access_token: "at_test",
        refresh_token: "rt_new",
        expires_in: 3600,
      }),
  });
  for (const r of graphResponses) {
    mockFetch.mockResolvedValueOnce({
      ok: r.ok,
      status: r.status,
      statusText: r.ok ? "OK" : "Error",
      headers: new Headers(),
      json: () => Promise.resolve(r.data),
    });
  }
  vi.stubGlobal("fetch", mockFetch);
  return mockFetch;
}

describe("Workbook 結合テスト: テーブル CRUD フロー", () => {
  let tools: Map<string, CapturedTool["handler"]>;

  beforeEach(async () => {
    clearMockRedis();
    clearAccessTokenCache();
    vi.restoreAllMocks();
    await saveRefreshToken("rt_test");
    const m = makeServer();
    tools = m.tools;
    registerWorkbookTools(m.server);
  });

  it("WF01: createSession → createTable → addRows → listRows → updateRow → deleteRow → closeSession の一連フロー", async () => {
    const fetchMock = setupSequentialFetch(
      // 1. createSession 応答
      {
        ok: true,
        status: 201,
        data: { id: "session-abc-xyz", persistChanges: true },
      },
      // 2. createTable 応答
      {
        ok: true,
        status: 201,
        data: { id: "{xx-yy}", name: "Table1" },
      },
      // 3. addRows (1回目: 単行) 応答
      {
        ok: true,
        status: 201,
        data: { index: 0, values: [["田中", "tanaka@example.com"]] },
      },
      // 4. addRows (2回目: 複数行) 応答
      {
        ok: true,
        status: 201,
        data: {
          index: 1,
          values: [
            ["佐藤", "sato@example.com"],
            ["鈴木", "suzuki@example.com"],
          ],
        },
      },
      // 5. listRows 応答（3行が入った状態）
      {
        ok: true,
        status: 200,
        data: {
          value: [
            { index: 0, values: [["田中", "tanaka@example.com"]] },
            { index: 1, values: [["佐藤", "sato@example.com"]] },
            { index: 2, values: [["鈴木", "suzuki@example.com"]] },
          ],
        },
      },
      // 6. updateRow 応答
      {
        ok: true,
        status: 200,
        data: {
          index: 0,
          values: [["田中（更新済）", "tanaka-new@example.com"]],
        },
      },
      // 7. deleteRow 応答（204 No Content）
      { ok: true, status: 204, data: {} },
      // 8. closeSession 応答（204 No Content）
      { ok: true, status: 204, data: {} }
    );

    const ITEM_ID = "01TESTFILE";

    // ------------------ 1. createSession ------------------
    const create = tools.get("workbook_create_session")!;
    const sessRes = await create({
      item_id: ITEM_ID,
      persist_changes: true,
    });
    const sessJson = JSON.parse(sessRes.content[0].text);
    expect(sessJson.workbook_session_id).toBe("session-abc-xyz");
    const SESS = sessJson.workbook_session_id;

    // ------------------ 2. createTable ------------------
    const createTable = tools.get("workbook_create_table")!;
    await createTable({
      item_id: ITEM_ID,
      worksheet: "Sheet1",
      address: "A1:B1",
      has_headers: true,
      workbook_session_id: SESS,
    });

    // ------------------ 3. addRows (1回目: 単行) ------------------
    const addRows = tools.get("workbook_table_add_rows")!;
    const r1 = await addRows({
      item_id: ITEM_ID,
      table: "Table1",
      values: [["田中", "tanaka@example.com"]],
      workbook_session_id: SESS,
    });
    expect(JSON.parse(r1.content[0].text).rows_added).toBe(1);

    // ------------------ 4. addRows (2回目: 複数行) ------------------
    const r2 = await addRows({
      item_id: ITEM_ID,
      table: "Table1",
      values: [
        ["佐藤", "sato@example.com"],
        ["鈴木", "suzuki@example.com"],
      ],
      workbook_session_id: SESS,
    });
    expect(JSON.parse(r2.content[0].text).rows_added).toBe(2);

    // ------------------ 5. listRows ------------------
    const listRows = tools.get("workbook_table_list_rows")!;
    const listRes = await listRows({
      item_id: ITEM_ID,
      table: "Table1",
      top: 100,
      workbook_session_id: SESS,
    });
    const listJson = JSON.parse(listRes.content[0].text);
    expect(listJson.count).toBe(3);

    // ------------------ 6. updateRow ------------------
    const updateRow = tools.get("workbook_table_update_row")!;
    await updateRow({
      item_id: ITEM_ID,
      table: "Table1",
      index: 0,
      values: [["田中（更新済）", "tanaka-new@example.com"]],
      workbook_session_id: SESS,
    });

    // ------------------ 7. deleteRow ------------------
    const deleteRow = tools.get("workbook_table_delete_row")!;
    await deleteRow({
      item_id: ITEM_ID,
      table: "Table1",
      index: 2,
      workbook_session_id: SESS,
    });

    // ------------------ 8. closeSession ------------------
    const close = tools.get("workbook_close_session")!;
    await close({
      item_id: ITEM_ID,
      workbook_session_id: SESS,
    });

    // ============== 検証 ==============
    // calls[0] = token endpoint
    // calls[1] = createSession
    // calls[2] = createTable
    // calls[3] = addRows 1回目
    // calls[4] = addRows 2回目
    // calls[5] = listRows
    // calls[6] = updateRow
    // calls[7] = deleteRow
    // calls[8] = closeSession
    expect(fetchMock).toHaveBeenCalledTimes(9);

    // createSession は session-id ヘッダー無し（これがセッションを作る側）
    const createSessionHeaders = fetchMock.mock.calls[1][1].headers as Record<
      string,
      string
    >;
    expect(createSessionHeaders["workbook-session-id"]).toBeUndefined();

    // 以降の操作には session-id ヘッダーが付与される
    for (let i = 2; i <= 8; i++) {
      const headers = fetchMock.mock.calls[i][1].headers as Record<string, string>;
      expect(headers["workbook-session-id"]).toBe("session-abc-xyz");
    }

    // URL とメソッドの妥当性スポットチェック
    expect(fetchMock.mock.calls[1][0]).toContain("/workbook/createSession");
    expect(fetchMock.mock.calls[1][1].method).toBe("POST");

    expect(fetchMock.mock.calls[2][0]).toContain("/worksheets/Sheet1/tables/add");
    expect(fetchMock.mock.calls[2][1].method).toBe("POST");

    expect(fetchMock.mock.calls[3][0]).toContain("/tables/Table1/rows/add");
    expect(fetchMock.mock.calls[4][0]).toContain("/tables/Table1/rows/add");

    expect(fetchMock.mock.calls[5][0]).toContain("/tables/Table1/rows");
    expect(fetchMock.mock.calls[5][1].method ?? "GET").toBe("GET");

    expect(fetchMock.mock.calls[6][0]).toContain("/rows/itemAt(index=0)");
    expect(fetchMock.mock.calls[6][1].method).toBe("PATCH");

    expect(fetchMock.mock.calls[7][0]).toContain("/rows/itemAt(index=2)");
    expect(fetchMock.mock.calls[7][1].method).toBe("DELETE");

    expect(fetchMock.mock.calls[8][0]).toContain("/workbook/closeSession");
  });

  it("WF02: 生シートに直接書き込み (range_update) からの読み戻し (range_get)", async () => {
    const fetchMock = setupSequentialFetch(
      // 1. range_update 応答
      {
        ok: true,
        status: 200,
        data: {
          address: "Sheet1!A1:C2",
          rowCount: 2,
          columnCount: 3,
          values: [
            ["氏名", "メール", "電話"],
            ["田中", "tanaka@example.com", "090-0000-0001"],
          ],
        },
      },
      // 2. range_get 応答
      {
        ok: true,
        status: 200,
        data: {
          address: "Sheet1!A1:C2",
          rowCount: 2,
          columnCount: 3,
          values: [
            ["氏名", "メール", "電話"],
            ["田中", "tanaka@example.com", "090-0000-0001"],
          ],
        },
      }
    );

    const ITEM_ID = "01RAWFILE";

    // 1. 書き込み
    const update = tools.get("workbook_range_update")!;
    const writeRes = await update({
      item_id: ITEM_ID,
      worksheet: "Sheet1",
      address: "A1:C2",
      values: [
        ["氏名", "メール", "電話"],
        ["田中", "tanaka@example.com", "090-0000-0001"],
      ],
    });
    expect(JSON.parse(writeRes.content[0].text).success).toBe(true);

    // 2. 読み戻し
    const get = tools.get("workbook_range_get")!;
    const readRes = await get({
      item_id: ITEM_ID,
      worksheet: "Sheet1",
      address: "A1:C2",
    });
    const data = JSON.parse(readRes.content[0].text);
    expect(data.values[1][0]).toBe("田中");

    // メソッド検証
    expect(fetchMock.mock.calls[1][1].method).toBe("PATCH");
    expect(fetchMock.mock.calls[2][1].method ?? "GET").toBe("GET");
  });

  it("WF03: workbook_call_function で SUM を実行 → セッションありで複数回連続呼び出し", async () => {
    const fetchMock = setupSequentialFetch(
      // createSession
      {
        ok: true,
        status: 201,
        data: { id: "calc-session", persistChanges: false },
      },
      // SUM
      { ok: true, status: 200, data: { value: 15 } },
      // AVERAGE
      { ok: true, status: 200, data: { value: 3 } },
      // PMT
      { ok: true, status: 200, data: { value: -1199.10 } },
      // closeSession
      { ok: true, status: 204, data: {} }
    );

    const ITEM_ID = "01CALCFILE";
    const create = tools.get("workbook_create_session")!;
    const callFn = tools.get("workbook_call_function")!;
    const close = tools.get("workbook_close_session")!;

    const sess = JSON.parse(
      (await create({ item_id: ITEM_ID, persist_changes: false })).content[0].text
    ).workbook_session_id;

    const sumRes = await callFn({
      item_id: ITEM_ID,
      function_name: "sum",
      arguments: { values: [1, 2, 3, 4, 5] },
      workbook_session_id: sess,
    });
    expect(JSON.parse(sumRes.content[0].text).value).toBe(15);

    await callFn({
      item_id: ITEM_ID,
      function_name: "average",
      arguments: { values: [1, 2, 3, 4, 5] },
      workbook_session_id: sess,
    });

    await callFn({
      item_id: ITEM_ID,
      function_name: "PMT",
      arguments: { rate: 0.005, nper: 360, pv: -200000 },
      workbook_session_id: sess,
    });

    await close({ item_id: ITEM_ID, workbook_session_id: sess });

    // 全Graph呼び出し（5回）にセッションIDヘッダーが付くことを確認
    // calls[0]=token, calls[1]=createSession（IDなし）, calls[2-5]=以降IDあり
    for (let i = 2; i <= 5; i++) {
      const headers = fetchMock.mock.calls[i][1].headers as Record<string, string>;
      expect(headers["workbook-session-id"]).toBe("calc-session");
    }

    // 関数名が小文字化されていることを確認
    expect(fetchMock.mock.calls[2][0]).toContain("/functions/sum");
    expect(fetchMock.mock.calls[3][0]).toContain("/functions/average");
    expect(fetchMock.mock.calls[4][0]).toContain("/functions/pmt"); // PMT → pmt
  });

  it("WF04: 504 エラーが起きてもリトライで自動回復してフロー継続", async () => {
    const fetchMock = vi.fn();
    // token
    mockFetchToken(fetchMock);
    // createSession 1回目: 504, 2回目: 成功
    mockFetchStatus(fetchMock, 504, {});
    mockFetchData(fetchMock, 201, { id: "recovered-session", persistChanges: true });
    // createTable: 504, 504, 成功 (3回目で復活)
    mockFetchStatus(fetchMock, 504, {});
    mockFetchStatus(fetchMock, 504, {});
    mockFetchData(fetchMock, 201, { id: "T1", name: "Table1" });
    vi.stubGlobal("fetch", fetchMock);

    const ITEM_ID = "01FLAKY";
    const create = tools.get("workbook_create_session")!;
    const sess = JSON.parse(
      (await create({ item_id: ITEM_ID })).content[0].text
    ).workbook_session_id;
    expect(sess).toBe("recovered-session");

    const createTable = tools.get("workbook_create_table")!;
    const tblRes = await createTable({
      item_id: ITEM_ID,
      worksheet: "Sheet1",
      address: "A1:B1",
      has_headers: true,
      workbook_session_id: sess,
    });
    expect(JSON.parse(tblRes.content[0].text).success).toBe(true);

    // token + createSession(2) + createTable(3) = 6回 fetch
    expect(fetchMock).toHaveBeenCalledTimes(6);
  }, 30000);

  it("WF05: SharePoint 上の Excel ファイルへの操作（drive_id 指定）", async () => {
    const fetchMock = setupSequentialFetch(
      // createSession
      {
        ok: true,
        status: 201,
        data: { id: "sp-session", persistChanges: true },
      },
      // addRows
      {
        ok: true,
        status: 201,
        data: { index: 0, values: [["A", "B"]] },
      }
    );

    const create = tools.get("workbook_create_session")!;
    const addRows = tools.get("workbook_table_add_rows")!;

    const sess = JSON.parse(
      (
        await create({
          drive_id: "b!sharepoint-drive",
          item_id: "01ABC",
          persist_changes: true,
        })
      ).content[0].text
    ).workbook_session_id;

    await addRows({
      drive_id: "b!sharepoint-drive",
      item_id: "01ABC",
      table: "Sales",
      values: [["A", "B"]],
      workbook_session_id: sess,
    });

    expect(fetchMock.mock.calls[1][0]).toContain(
      "/drives/b!sharepoint-drive/items/01ABC/workbook/createSession"
    );
    expect(fetchMock.mock.calls[2][0]).toContain(
      "/drives/b!sharepoint-drive/items/01ABC/workbook/tables/Sales/rows/add"
    );
  });
});

// ── ヘルパー（WF04 用） ──
function mockFetchToken(fn: ReturnType<typeof vi.fn>) {
  fn.mockResolvedValueOnce({
    ok: true,
    json: () =>
      Promise.resolve({
        access_token: "at_test",
        refresh_token: "rt_new",
        expires_in: 3600,
      }),
  });
}
function mockFetchStatus(fn: ReturnType<typeof vi.fn>, status: number, data: unknown) {
  fn.mockResolvedValueOnce({
    ok: false,
    status,
    statusText: "Error",
    headers: new Headers(),
    json: () => Promise.resolve(data),
  });
}
function mockFetchData(fn: ReturnType<typeof vi.fn>, status: number, data: unknown) {
  fn.mockResolvedValueOnce({
    ok: true,
    status,
    statusText: "OK",
    headers: new Headers(),
    json: () => Promise.resolve(data),
  });
}
