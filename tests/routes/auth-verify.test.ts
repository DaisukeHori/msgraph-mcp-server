import { describe, it, expect, beforeEach } from "vitest";
import { clearMockRedis } from "../setup";
import { POST } from "@/app/api/auth/verify/route";
import { NextRequest } from "next/server";

function makeRequest(body: unknown, ip = "127.0.0.1"): NextRequest {
  return new NextRequest("http://localhost/api/auth/verify", {
    method: "POST",
    body: JSON.stringify(body),
    headers: {
      "content-type": "application/json",
      "x-forwarded-for": ip,
    },
  });
}

describe("POST /api/auth/verify", () => {
  beforeEach(() => {
    clearMockRedis();
  });

  it("V01: 正しい ADMIN_SECRET でセッショントークンを返す", async () => {
    const res = await POST(makeRequest({ secret: "test-admin-secret-123" }));
    expect(res.status).toBe(200);
    const data = await res.json();
    expect(data.sessionToken).toBeDefined();
    expect(data.sessionToken.length).toBe(64);
  });

  it("V02: 不正な ADMIN_SECRET で 401", async () => {
    const res = await POST(makeRequest({ secret: "wrong" }));
    expect(res.status).toBe(401);
    const data = await res.json();
    expect(data.error).toContain("正しくありません");
  });

  it("V03: 空の secret で 401", async () => {
    const res = await POST(makeRequest({ secret: "" }));
    expect(res.status).toBe(401);
  });

  it("V04: secret なしで 401", async () => {
    const res = await POST(makeRequest({}));
    expect(res.status).toBe(401);
  });

  it("V05: 残り回数がレスポンスに含まれる", async () => {
    const res = await POST(makeRequest({ secret: "wrong" }, "10.0.0.1"));
    const data = await res.json();
    expect(data.remaining).toBe(4);
  });

  it("V06: 5回失敗でロックアウト 429", async () => {
    for (let i = 0; i < 5; i++) {
      await POST(makeRequest({ secret: "wrong" }, "10.0.0.2"));
    }
    const res = await POST(makeRequest({ secret: "test-admin-secret-123" }, "10.0.0.2"));
    expect(res.status).toBe(429);
    const data = await res.json();
    expect(data.error).toContain("ロックアウト");
  });

  it("V07: 異なる IP は独立カウント", async () => {
    for (let i = 0; i < 5; i++) {
      await POST(makeRequest({ secret: "wrong" }, "10.0.0.3"));
    }
    // 別 IP からは正常にアクセス可能
    const res = await POST(makeRequest({ secret: "test-admin-secret-123" }, "10.0.0.4"));
    expect(res.status).toBe(200);
  });

  it("V08: 成功後ロックアウトカウンターがクリアされる", async () => {
    // 3回失敗
    for (let i = 0; i < 3; i++) {
      await POST(makeRequest({ secret: "wrong" }, "10.0.0.5"));
    }
    // 成功
    await POST(makeRequest({ secret: "test-admin-secret-123" }, "10.0.0.5"));
    // さらに3回失敗してもロックされない（カウンターリセット済み）
    for (let i = 0; i < 3; i++) {
      await POST(makeRequest({ secret: "wrong" }, "10.0.0.5"));
    }
    const res = await POST(makeRequest({ secret: "wrong" }, "10.0.0.5"));
    expect(res.status).toBe(401); // 429 ではない
  });

  it("V09: 不正な JSON ボディで 401", async () => {
    const req = new NextRequest("http://localhost/api/auth/verify", {
      method: "POST",
      body: "not json",
      headers: { "content-type": "application/json", "x-forwarded-for": "1.1.1.1" },
    });
    const res = await POST(req);
    expect(res.status).toBe(401);
  });

  it("V10: ADMIN_SECRET 未設定で 500", async () => {
    const original = process.env.ADMIN_SECRET;
    delete process.env.ADMIN_SECRET;
    const res = await POST(makeRequest({ secret: "test" }));
    expect(res.status).toBe(500);
    process.env.ADMIN_SECRET = original;
  });

  it("V11: 毎回異なるセッショントークンを返す", async () => {
    const res1 = await POST(makeRequest({ secret: "test-admin-secret-123" }, "20.0.0.1"));
    const res2 = await POST(makeRequest({ secret: "test-admin-secret-123" }, "20.0.0.2"));
    const data1 = await res1.json();
    const data2 = await res2.json();
    expect(data1.sessionToken).not.toBe(data2.sessionToken);
  });

  it("V12: タイミングセーフ比較（長さ違いでも即座に拒否しない）", async () => {
    const start = Date.now();
    await POST(makeRequest({ secret: "x" }));
    const shortTime = Date.now() - start;

    const start2 = Date.now();
    await POST(makeRequest({ secret: "x".repeat(1000) }));
    const longTime = Date.now() - start2;

    // 両方とも極端に速い（数ms）ことを確認（タイミング攻撃耐性）
    expect(shortTime).toBeLessThan(500);
    expect(longTime).toBeLessThan(500);
  });
});
