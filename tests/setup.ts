/**
 * テストセットアップ
 * 全テストの前に環境変数とモックを初期化
 */

import { vi } from "vitest";

// ── テスト用環境変数 ──
process.env.ADMIN_SECRET = "test-admin-secret-123";
process.env.MICROSOFT_CLIENT_ID = "test-client-id";
process.env.MICROSOFT_CLIENT_SECRET = "test-client-secret";
process.env.MICROSOFT_TENANT_ID = "test-tenant-id";
process.env.TOKEN_ENCRYPTION_KEY = "a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2";
process.env.CRON_SECRET = "test-cron-secret";
process.env.KV_REST_API_URL = "https://test.upstash.io";
process.env.KV_REST_API_TOKEN = "test-redis-token";

// ── Upstash Redis モック ──
const mockRedisStore = new Map<string, { value: string; ttl?: number }>();

vi.mock("@upstash/redis", () => ({
  Redis: class MockRedis {
    async get<T>(key: string): Promise<T | null> {
      const entry = mockRedisStore.get(key);
      return entry ? (entry.value as unknown as T) : null;
    }
    async set(key: string, value: unknown, options?: { ex?: number }): Promise<void> {
      mockRedisStore.set(key, {
        value: typeof value === "string" ? value : JSON.stringify(value),
        ttl: options?.ex,
      });
    }
    async del(key: string): Promise<void> {
      mockRedisStore.delete(key);
    }
    async incr(key: string): Promise<number> {
      const entry = mockRedisStore.get(key);
      const newVal = entry ? parseInt(entry.value) + 1 : 1;
      mockRedisStore.set(key, { value: String(newVal), ttl: entry?.ttl });
      return newVal;
    }
    async expire(key: string, seconds: number): Promise<void> {
      const entry = mockRedisStore.get(key);
      if (entry) entry.ttl = seconds;
    }
    async ttl(key: string): Promise<number> {
      const entry = mockRedisStore.get(key);
      return entry?.ttl ?? -1;
    }
  },
}));

// Redis ストアをリセットするヘルパー
export function clearMockRedis() {
  mockRedisStore.clear();
}

export function getMockRedisStore() {
  return mockRedisStore;
}

// ── fetch モック用ヘルパー ──
export function mockFetchResponse(data: unknown, status = 200) {
  return vi.fn().mockResolvedValue({
    ok: status >= 200 && status < 300,
    status,
    statusText: status === 200 ? "OK" : "Error",
    json: () => Promise.resolve(data),
    text: () => Promise.resolve(JSON.stringify(data)),
    headers: new Headers({ "content-type": "application/json" }),
    arrayBuffer: () => Promise.resolve(Buffer.from(JSON.stringify(data))),
  });
}

export function mockFetchSequence(responses: Array<{ data: unknown; status?: number }>) {
  const fn = vi.fn();
  responses.forEach((r, i) => {
    fn.mockResolvedValueOnce({
      ok: (r.status || 200) >= 200 && (r.status || 200) < 300,
      status: r.status || 200,
      json: () => Promise.resolve(r.data),
      headers: new Headers({ "content-type": "application/json" }),
      arrayBuffer: () => Promise.resolve(Buffer.from(JSON.stringify(r.data))),
    });
  });
  return fn;
}
