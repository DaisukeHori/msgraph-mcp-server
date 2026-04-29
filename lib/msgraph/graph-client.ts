/**
 * Microsoft Graph API クライアント
 *
 * Redis の refresh_token → access_token で Graph API を呼ぶ。
 * Excel Workbook API 用に workbook-session-id ヘッダーと
 * 一時的エラー（429/503/504）の自動リトライをサポート。
 */

import { GRAPH_BASE_URL, CHARACTER_LIMIT } from "@/lib/config";
import { getGraphTokenFromRedis } from "./auth-context";

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
  /** Excel Workbook API で workbook-session-id ヘッダーを付与する */
  workbookSessionId?: string;
  /** Prefer: respond-async などの Prefer ヘッダーを付与する */
  prefer?: string;
  /** リトライの最大回数（デフォルト: 3） */
  maxRetries?: number;
}

// ── 定数 ──

const RETRY_STATUS_CODES = new Set([429, 503, 504]);
const DEFAULT_MAX_RETRIES = 3;
const BASE_BACKOFF_MS = 500;

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

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** Retry-After ヘッダー値を ms に変換（秒数または HTTP 日付） */
function parseRetryAfter(retryAfter: string | null): number | null {
  if (!retryAfter) return null;
  const seconds = parseInt(retryAfter, 10);
  if (!isNaN(seconds)) return seconds * 1000;
  const date = Date.parse(retryAfter);
  if (!isNaN(date)) return Math.max(0, date - Date.now());
  return null;
}

// ── メインリクエスト ──

export async function graphRequest<T>(
  endpoint: string,
  options: RequestOptions = {}
): Promise<T> {
  const {
    method = "GET",
    body,
    headers = {},
    queryParams,
    workbookSessionId,
    prefer,
    maxRetries = DEFAULT_MAX_RETRIES,
  } = options;
  const token = await getGraphTokenFromRedis();
  const url = buildUrl(endpoint, queryParams);

  const fetchHeaders: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
    ...headers,
  };
  if (workbookSessionId) {
    fetchHeaders["workbook-session-id"] = workbookSessionId;
  }
  if (prefer) {
    fetchHeaders["Prefer"] = prefer;
  }

  const fetchOptions: RequestInit = { method, headers: fetchHeaders };

  if (body !== undefined && method !== "GET" && method !== "DELETE") {
    fetchOptions.body = JSON.stringify(body);
  }

  let lastError: Error | null = null;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const response = await fetch(url, fetchOptions);

    if (response.status === 204) return undefined as T;

    // 一時的エラー → リトライ
    if (RETRY_STATUS_CODES.has(response.status) && attempt < maxRetries) {
      const retryAfterHeader = response.headers.get("Retry-After");
      const retryAfterMs = parseRetryAfter(retryAfterHeader);
      // Retry-After 指定があればそれに従い、なければ指数バックオフ
      const backoffMs =
        retryAfterMs ?? BASE_BACKOFF_MS * Math.pow(2, attempt);
      lastError = new Error(
        `Graph API ${response.status}（${attempt + 1}/${maxRetries + 1}回目）: ${response.statusText}`
      );
      await sleep(backoffMs);
      continue;
    }

    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      // リトライ対象ステータスでリトライ上限到達した場合は status を含めたメッセージにする
      if (RETRY_STATUS_CODES.has(response.status)) {
        throw new Error(
          `Graph API ${response.status}（リトライ${maxRetries}回後も復旧せず）: ${response.statusText || formatError(data as GraphErrorBody)}`
        );
      }
      throw new Error(formatError(data as GraphErrorBody));
    }
    return data as T;
  }

  // ここに到達するのはリトライ上限到達のみ
  throw lastError ?? new Error("Graph API リクエスト失敗（リトライ上限到達）");
}

// ── CRUD ヘルパー ──

export async function graphGet<T>(
  endpoint: string,
  queryParams?: Record<string, string | number | boolean | undefined>,
  options: { workbookSessionId?: string } = {}
): Promise<T> {
  return graphRequest<T>(endpoint, {
    queryParams,
    workbookSessionId: options.workbookSessionId,
  });
}

export async function graphPost<T>(
  endpoint: string,
  body: unknown,
  queryParams?: Record<string, string | number | boolean | undefined>,
  options: { workbookSessionId?: string; prefer?: string } = {}
): Promise<T> {
  return graphRequest<T>(endpoint, {
    method: "POST",
    body,
    queryParams,
    workbookSessionId: options.workbookSessionId,
    prefer: options.prefer,
  });
}

export async function graphPatch<T>(
  endpoint: string,
  body: unknown,
  options: { workbookSessionId?: string } = {}
): Promise<T> {
  return graphRequest<T>(endpoint, {
    method: "PATCH",
    body,
    workbookSessionId: options.workbookSessionId,
  });
}

export async function graphDelete(
  endpoint: string,
  options: { workbookSessionId?: string } = {}
): Promise<void> {
  await graphRequest<void>(endpoint, {
    method: "DELETE",
    workbookSessionId: options.workbookSessionId,
  });
}

// ── ファイルアップロード (< 4MB) ──

export async function graphUploadSmallFile<T>(
  endpoint: string,
  content: Buffer | string,
  contentType = "application/octet-stream"
): Promise<T> {
  const token = await getGraphTokenFromRedis();
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
  const token = await getGraphTokenFromRedis();
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
