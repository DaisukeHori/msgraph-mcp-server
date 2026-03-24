/**
 * Microsoft Graph API クライアント
 *
 * AsyncLocalStorage ベースの認証コンテキストからトークンを取得し、
 * Graph API へのリクエストを実行する。
 */

import { GRAPH_BASE_URL, CHARACTER_LIMIT } from "@/lib/config";
import { getGraphToken } from "./auth-context";

// ── 型定義 ──

export interface GraphPagedResponse<T> {
  "@odata.context"?: string;
  "@odata.count"?: number;
  "@odata.nextLink"?: string;
  value: T[];
}

interface GraphErrorBody {
  error?: { code?: string; message?: string };
}

interface RequestOptions {
  method?: string;
  body?: unknown;
  headers?: Record<string, string>;
  queryParams?: Record<string, string | number | boolean | undefined>;
}

// ── ユーティリティ ──

function buildUrl(
  endpoint: string,
  queryParams?: Record<string, string | number | boolean | undefined>
): string {
  const url = new URL(`${GRAPH_BASE_URL}${endpoint}`);
  if (queryParams) {
    for (const [key, value] of Object.entries(queryParams)) {
      if (value !== undefined && value !== null) {
        url.searchParams.set(key, String(value));
      }
    }
  }
  return url.toString();
}

function formatError(body: GraphErrorBody): string {
  const code = body.error?.code || "UnknownError";
  const message = body.error?.message || "不明なエラーが発生しました";
  return `Microsoft Graph API エラー [${code}]: ${message}`;
}

// ── メインリクエスト ──

export async function graphRequest<T>(
  endpoint: string,
  options: RequestOptions = {}
): Promise<T> {
  const { method = "GET", body, headers = {}, queryParams } = options;
  const token = await getGraphToken();
  const url = buildUrl(endpoint, queryParams);

  const fetchHeaders: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
    ...headers,
  };

  const fetchOptions: RequestInit = { method, headers: fetchHeaders };

  if (body && method !== "GET" && method !== "DELETE") {
    fetchOptions.body = JSON.stringify(body);
  }

  const response = await fetch(url, fetchOptions);

  if (response.status === 204) return undefined as T;

  const data = await response.json();
  if (!response.ok) throw new Error(formatError(data as GraphErrorBody));
  return data as T;
}

// ── CRUD ヘルパー ──

export async function graphGet<T>(
  endpoint: string,
  queryParams?: Record<string, string | number | boolean | undefined>
): Promise<T> {
  return graphRequest<T>(endpoint, { queryParams });
}

export async function graphPost<T>(
  endpoint: string,
  body: unknown,
  queryParams?: Record<string, string | number | boolean | undefined>
): Promise<T> {
  return graphRequest<T>(endpoint, { method: "POST", body, queryParams });
}

export async function graphPatch<T>(
  endpoint: string,
  body: unknown
): Promise<T> {
  return graphRequest<T>(endpoint, { method: "PATCH", body });
}

export async function graphDelete(endpoint: string): Promise<void> {
  await graphRequest<void>(endpoint, { method: "DELETE" });
}

// ── ファイルアップロード (< 4MB) ──

export async function graphUploadSmallFile<T>(
  endpoint: string,
  content: Buffer | string,
  contentType = "application/octet-stream"
): Promise<T> {
  const token = await getGraphToken();
  const url = `${GRAPH_BASE_URL}${endpoint}`;

  const response = await fetch(url, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": contentType },
    body: new Uint8Array(
      content instanceof Buffer ? content : Buffer.from(content as string)
    ),
  });

  const data = await response.json();
  if (!response.ok) throw new Error(formatError(data as GraphErrorBody));
  return data as T;
}

// ── ファイルダウンロード ──

export async function graphDownloadFile(
  endpoint: string
): Promise<{ content: string; contentType: string }> {
  const token = await getGraphToken();
  const response = await fetch(`${GRAPH_BASE_URL}${endpoint}`, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(formatError(err as GraphErrorBody));
  }

  const ct = response.headers.get("content-type") || "application/octet-stream";
  const buffer = Buffer.from(await response.arrayBuffer());

  if (ct.startsWith("text/") || ct.includes("json") || ct.includes("xml")) {
    return { content: buffer.toString("utf-8"), contentType: ct };
  }
  return { content: buffer.toString("base64"), contentType: ct };
}

// ── レスポンスユーティリティ ──

export function truncateResponse(text: string): string {
  if (text.length <= CHARACTER_LIMIT) return text;
  return (
    text.slice(0, CHARACTER_LIMIT) +
    "\n\n--- レスポンスが上限を超えたため切り詰められました。フィルターやページネーションで絞り込んでください。 ---"
  );
}

export function handleToolError(error: unknown): string {
  if (error instanceof Error) return `エラー: ${error.message}`;
  return `エラー: ${String(error)}`;
}
