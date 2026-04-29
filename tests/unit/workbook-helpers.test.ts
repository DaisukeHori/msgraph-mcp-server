/**
 * workbook-helpers.ts の単体テスト
 *
 * Excel Workbook API のエンドポイント構築が正しいことを検証する。
 * 4種類の指定方法 × 4種類のロケーション（me/users/drives）の組み合わせ
 * が想定通りの URL になることを確認する。
 */

import { describe, it, expect } from "vitest";
import {
  workbookBase,
  encodeWorksheetRef,
  encodeTableRef,
} from "@/lib/mcp/tools/workbook-helpers";

describe("workbookBase: ファイル指定パターン", () => {
  // ── /me/drive 系（本人 OneDrive） ──

  it("WB01: 本人 OneDrive + item_id", () => {
    expect(workbookBase({ item_id: "01XXXXX" })).toBe(
      "/me/drive/items/01XXXXX/workbook"
    );
  });

  it("WB02: 本人 OneDrive + path（先頭スラッシュあり）", () => {
    expect(workbookBase({ path: "/Documents/data.xlsx" })).toBe(
      "/me/drive/root:/Documents/data.xlsx:/workbook"
    );
  });

  it("WB03: 本人 OneDrive + path（先頭スラッシュなし → 自動で付加）", () => {
    expect(workbookBase({ path: "Documents/data.xlsx" })).toBe(
      "/me/drive/root:/Documents/data.xlsx:/workbook"
    );
  });

  // ── /users/{user_id}/drive 系（共有メールボックス・委任） ──

  it("WB04: 委任ユーザー + item_id", () => {
    expect(
      workbookBase({ user_id: "user@example.com", item_id: "01YYYYY" })
    ).toBe("/users/user@example.com/drive/items/01YYYYY/workbook");
  });

  it("WB05: 委任ユーザー + path", () => {
    expect(
      workbookBase({ user_id: "user@example.com", path: "/data.xlsx" })
    ).toBe("/users/user@example.com/drive/root:/data.xlsx:/workbook");
  });

  // ── /drives/{drive_id}/items 系（SharePoint・他人ドライブ） ──

  it("WB06: drive_id + item_id（SharePoint ドキュメントライブラリ）", () => {
    expect(
      workbookBase({ drive_id: "b!abc123", item_id: "01ZZZZZ" })
    ).toBe("/drives/b!abc123/items/01ZZZZZ/workbook");
  });

  it("WB07: drive_id + path", () => {
    expect(
      workbookBase({ drive_id: "b!abc123", path: "/Shared/budget.xlsx" })
    ).toBe("/drives/b!abc123/root:/Shared/budget.xlsx:/workbook");
  });

  it("WB08: drive_id 指定時は user_id を無視", () => {
    // drive_id があれば /drives/... が優先される（共有ドライブ直接アクセス）
    expect(
      workbookBase({
        drive_id: "b!shared",
        user_id: "ignored@example.com",
        item_id: "01AAA",
      })
    ).toBe("/drives/b!shared/items/01AAA/workbook");
  });

  // ── エラー系 ──

  it("WB09: item_id も path も無いとエラー", () => {
    expect(() => workbookBase({})).toThrow("item_id か path");
  });

  it("WB10: drive_id だけ指定でファイル未指定もエラー", () => {
    expect(() => workbookBase({ drive_id: "b!only" })).toThrow(
      "item_id か path"
    );
  });

  // ── path の特殊文字 ──

  it("WB11: 日本語パスもそのまま（呼び出し側でエンコード不要）", () => {
    expect(workbookBase({ path: "/業務/集計表.xlsx" })).toBe(
      "/me/drive/root:/業務/集計表.xlsx:/workbook"
    );
  });
});

describe("encodeWorksheetRef / encodeTableRef", () => {
  it("EN01: 英数字はそのまま", () => {
    expect(encodeWorksheetRef("Sheet1")).toBe("Sheet1");
    expect(encodeTableRef("Table1")).toBe("Table1");
  });

  it("EN02: スペース含む名前を URL エンコード", () => {
    expect(encodeWorksheetRef("Q3 Sales")).toBe("Q3%20Sales");
    expect(encodeTableRef("Sales Table")).toBe("Sales%20Table");
  });

  it("EN03: 日本語シート名を URL エンコード", () => {
    const encoded = encodeWorksheetRef("売上");
    expect(encoded).toMatch(/^%[0-9A-F]+/);
    expect(decodeURIComponent(encoded)).toBe("売上");
  });

  it("EN04: 特殊記号を URL エンコード", () => {
    // # は URL でフラグメント扱いになるのでエンコード必須
    expect(encodeWorksheetRef("Top#1")).toBe("Top%231");
  });

  it("EN05: ID 形式（{guid}）はそのままパススルー（中括弧はエンコードされる）", () => {
    const id = "{00000000-0001-0000-0100-000000000000}";
    const encoded = encodeWorksheetRef(id);
    expect(decodeURIComponent(encoded)).toBe(id);
  });
});
