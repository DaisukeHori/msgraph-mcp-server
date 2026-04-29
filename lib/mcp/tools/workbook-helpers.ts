/**
 * Workbook 関連の共通ヘルパー
 *
 * Excel Workbook API のエンドポイントは以下のいずれかをベースにする：
 *   /me/drive/items/{item_id}/workbook/...
 *   /me/drive/root:/{path}:/workbook/...
 *   /users/{user_id}/drive/items/{item_id}/workbook/...
 *   /drives/{drive_id}/items/{item_id}/workbook/...
 *
 * このヘルパーは item_id / path / drive_id / user_id を受け取り、
 * 適切な /workbook プレフィックスを返す。
 */

import { userBase } from "@/lib/mcp/tools/shared-helpers";

export interface WorkbookFileLocator {
  item_id?: string;
  path?: string;
  drive_id?: string;
  user_id?: string;
}

/**
 * ファイル指定から workbook エンドポイントの "ベース" を返す。
 *
 * 戻り値の例:
 *   "/me/drive/items/01XXXX/workbook"
 *   "/drives/b!abc/items/01YYYY/workbook"
 *   "/me/drive/root:/Documents/data.xlsx:/workbook"
 *
 * @throws item_id も path も無い場合
 */
export function workbookBase(loc: WorkbookFileLocator): string {
  const drivePrefix = loc.drive_id
    ? `/drives/${loc.drive_id}`
    : `${userBase(loc.user_id)}/drive`;

  if (loc.item_id) {
    return `${drivePrefix}/items/${loc.item_id}/workbook`;
  }
  if (loc.path) {
    const cleanPath = loc.path.startsWith("/") ? loc.path : `/${loc.path}`;
    return `${drivePrefix}/root:${cleanPath}:/workbook`;
  }
  throw new Error(
    "Excel ファイルを指定してください: item_id か path のどちらかが必須です"
  );
}

/**
 * Worksheet ID または名前を URL セーフにエンコードする。
 * 名前にスペースや日本語が含まれていても安全に渡せるようにする。
 */
export function encodeWorksheetRef(idOrName: string): string {
  return encodeURIComponent(idOrName);
}

/** Table ID または名前を URL セーフにエンコード */
export function encodeTableRef(idOrName: string): string {
  return encodeURIComponent(idOrName);
}
