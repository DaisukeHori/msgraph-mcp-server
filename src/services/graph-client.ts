import { GRAPH_BASE_URL, CHARACTER_LIMIT } from "../constants.js";
import { GraphErrorResponse, GraphPagedResponse } from "../types.js";
import { getAccessToken } from "./auth.js";

interface RequestOptions {
  method?: string;
  body?: unknown;
  headers?: Record<string, string>;
  queryParams?: Record<string, string | number | boolean | undefined>;
  rawResponse?: boolean;
  beta?: boolean;
}

function buildUrl(
  endpoint: string,
  queryParams?: Record<string, string | number | boolean | undefined>,
  beta?: boolean
): string {
  const base = beta
    ? "https://graph.microsoft.com/beta"
    : GRAPH_BASE_URL;
  const url = new URL(`${base}${endpoint}`);
  if (queryParams) {
    for (const [key, value] of Object.entries(queryParams)) {
      if (value !== undefined && value !== null) {
        url.searchParams.set(key, String(value));
      }
    }
  }
  return url.toString();
}

function formatGraphError(error: GraphErrorResponse): string {
  const code = error.error?.code || "UnknownError";
  const message = error.error?.message || "An unknown error occurred";
  return `Microsoft Graph API Error [${code}]: ${message}`;
}

/**
 * Make a request to Microsoft Graph API
 */
export async function graphRequest<T>(
  endpoint: string,
  options: RequestOptions = {}
): Promise<T> {
  const { method = "GET", body, headers = {}, queryParams, beta } = options;
  const token = await getAccessToken();
  const url = buildUrl(endpoint, queryParams, beta);

  const fetchHeaders: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
    ...headers,
  };

  const fetchOptions: RequestInit = {
    method,
    headers: fetchHeaders,
  };

  if (body && method !== "GET" && method !== "DELETE") {
    fetchOptions.body = JSON.stringify(body);
  }

  const response = await fetch(url, fetchOptions);

  // Handle 204 No Content
  if (response.status === 204) {
    return undefined as T;
  }

  // Handle binary responses (file downloads)
  if (options.rawResponse) {
    return response as unknown as T;
  }

  const data = await response.json();

  if (!response.ok) {
    throw new Error(formatGraphError(data as GraphErrorResponse));
  }

  return data as T;
}

/**
 * GET request helper
 */
export async function graphGet<T>(
  endpoint: string,
  queryParams?: Record<string, string | number | boolean | undefined>,
  beta?: boolean
): Promise<T> {
  return graphRequest<T>(endpoint, { queryParams, beta });
}

/**
 * POST request helper
 */
export async function graphPost<T>(
  endpoint: string,
  body: unknown,
  queryParams?: Record<string, string | number | boolean | undefined>
): Promise<T> {
  return graphRequest<T>(endpoint, { method: "POST", body, queryParams });
}

/**
 * PATCH request helper
 */
export async function graphPatch<T>(
  endpoint: string,
  body: unknown
): Promise<T> {
  return graphRequest<T>(endpoint, { method: "PATCH", body });
}

/**
 * DELETE request helper
 */
export async function graphDelete(endpoint: string): Promise<void> {
  await graphRequest<void>(endpoint, { method: "DELETE" });
}

/**
 * PUT request helper (for file uploads etc.)
 */
export async function graphPut<T>(
  endpoint: string,
  body: unknown,
  headers?: Record<string, string>
): Promise<T> {
  return graphRequest<T>(endpoint, { method: "PUT", body, headers });
}

/**
 * Upload small file content (< 4MB) via PUT
 */
export async function graphUploadSmallFile<T>(
  endpoint: string,
  content: Buffer | string,
  contentType: string = "application/octet-stream"
): Promise<T> {
  const token = await getAccessToken();
  const url = `${GRAPH_BASE_URL}${endpoint}`;

  const response = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": contentType,
    },
    body: new Uint8Array(content instanceof Buffer ? content : Buffer.from(content as string)),
  });

  const data = await response.json();
  if (!response.ok) {
    throw new Error(formatGraphError(data as GraphErrorResponse));
  }
  return data as T;
}

/**
 * Download file content
 */
export async function graphDownloadFile(
  endpoint: string
): Promise<{ content: string; contentType: string }> {
  const token = await getAccessToken();
  const url = `${GRAPH_BASE_URL}${endpoint}`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(
      formatGraphError(errorData as GraphErrorResponse)
    );
  }

  const contentType = response.headers.get("content-type") || "application/octet-stream";
  const buffer = Buffer.from(await response.arrayBuffer());

  // Return base64 for binary, text for text
  if (
    contentType.startsWith("text/") ||
    contentType.includes("json") ||
    contentType.includes("xml")
  ) {
    return { content: buffer.toString("utf-8"), contentType };
  }
  return { content: buffer.toString("base64"), contentType };
}

/**
 * Paged GET - fetches all pages or up to maxPages
 */
export async function graphGetAllPages<T>(
  endpoint: string,
  queryParams?: Record<string, string | number | boolean | undefined>,
  maxPages: number = 5
): Promise<T[]> {
  const allItems: T[] = [];
  let currentUrl: string | null = buildUrl(endpoint, queryParams);
  let pageCount = 0;
  const token = await getAccessToken();

  while (currentUrl && pageCount < maxPages) {
    const response = await fetch(currentUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(
        formatGraphError(errorData as GraphErrorResponse)
      );
    }

    const data = (await response.json()) as GraphPagedResponse<T>;
    allItems.push(...data.value);
    currentUrl = data["@odata.nextLink"] || null;
    pageCount++;
  }

  return allItems;
}

/**
 * Truncate text if it exceeds CHARACTER_LIMIT
 */
export function truncateResponse(text: string): string {
  if (text.length <= CHARACTER_LIMIT) return text;
  return (
    text.slice(0, CHARACTER_LIMIT) +
    "\n\n--- Response truncated. Use filters or pagination to narrow results. ---"
  );
}

/**
 * Common error handler for tool implementations
 */
export function handleToolError(error: unknown): string {
  if (error instanceof Error) {
    return `Error: ${error.message}`;
  }
  return `Error: ${String(error)}`;
}
