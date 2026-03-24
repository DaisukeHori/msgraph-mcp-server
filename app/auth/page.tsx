"use client";

import { useState, useEffect, useCallback } from "react";

type Phase = "password" | "authenticated" | "setup_complete";

interface AuthStatus {
  authenticated: boolean;
  user: { name: string; email: string } | null;
  tokenUpdatedAt: string | null;
  tokenCreatedAt: string | null;
  mcpApiKeyConfigured: boolean;
  mcpApiKeyHint: string | null;
  lastCronExecution: string | null;
  cronSchedule: string;
}

export default function AuthPage() {
  const [phase, setPhase] = useState<Phase>("password");
  const [secret, setSecret] = useState("");
  const [sessionToken, setSessionToken] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [newApiKey, setNewApiKey] = useState("");
  const [status, setStatus] = useState<AuthStatus | null>(null);
  const [copied, setCopied] = useState(false);

  // URL パラメータを読む
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const errorParam = params.get("error");
    const successParam = params.get("success");
    const sessionParam = params.get("session");
    const newKeyParam = params.get("newKey");

    if (errorParam) {
      setError(errorParam);
    }
    if (successParam === "true" && sessionParam) {
      setSessionToken(sessionParam);
      setPhase("setup_complete");
      if (newKeyParam) {
        setNewApiKey(newKeyParam);
      }
      // URL をクリーンアップ
      window.history.replaceState({}, "", "/auth");
    }
  }, []);

  // ステータス取得
  const fetchStatus = useCallback(async () => {
    if (!sessionToken) return;
    try {
      const res = await fetch("/api/auth/status", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ session: sessionToken }),
      });
      if (res.ok) {
        setStatus(await res.json());
      }
    } catch {
      // ignore
    }
  }, [sessionToken]);

  useEffect(() => {
    if (phase === "authenticated" || phase === "setup_complete") {
      fetchStatus();
    }
  }, [phase, fetchStatus]);

  // ADMIN_SECRET 検証
  async function handleVerify(e: React.FormEvent) {
    e.preventDefault();
    setError("");
    setLoading(true);

    try {
      const res = await fetch("/api/auth/verify", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ secret }),
      });

      const data = await res.json();

      if (!res.ok) {
        setError(data.error || "認証失敗");
        if (data.remaining !== undefined) {
          setError(`${data.error} (残り${data.remaining}回)`);
        }
        return;
      }

      setSessionToken(data.sessionToken);
      setPhase("authenticated");
    } catch {
      setError("サーバーに接続できません");
    } finally {
      setLoading(false);
    }
  }

  // Microsoft ログイン開始
  function handleMicrosoftLogin() {
    window.location.href = `/api/auth/login?session=${sessionToken}`;
  }

  // API キー ローテーション
  async function handleRotateKey() {
    if (!confirm("新しい MCP API キーを発行しますか？\n古いキーは即座に無効化されます。")) {
      return;
    }
    setLoading(true);
    try {
      const res = await fetch("/api/auth/rotate-key", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ session: sessionToken }),
      });
      const data = await res.json();
      if (res.ok) {
        setNewApiKey(data.mcpApiKey);
        setCopied(false);
        fetchStatus();
      } else {
        setError(data.error);
      }
    } catch {
      setError("キーのローテーションに失敗しました");
    } finally {
      setLoading(false);
    }
  }

  // コピー
  function handleCopy() {
    navigator.clipboard.writeText(newApiKey);
    setCopied(true);
    setTimeout(() => setCopied(false), 3000);
  }

  return (
    <div style={styles.container}>
      <div style={styles.card}>
        <div style={styles.logo}>
          <svg width="20" height="20" viewBox="0 0 16 16" fill="none">
            <rect width="7" height="7" fill="#F25022" />
            <rect x="9" width="7" height="7" fill="#7FBA00" />
            <rect y="9" width="7" height="7" fill="#00A4EF" />
            <rect x="9" y="9" width="7" height="7" fill="#FFB900" />
          </svg>
          <span style={{ fontWeight: 600, fontSize: 16 }}>msgraph-mcp-server</span>
        </div>
        <h1 style={styles.title}>管理パネル</h1>

        {error && <div style={styles.error}>{error}</div>}

        {/* Phase 1: パスワード入力 */}
        {phase === "password" && (
          <form onSubmit={handleVerify}>
            <p style={styles.desc}>管理パスワードを入力してください</p>
            <input
              type="password"
              value={secret}
              onChange={(e) => setSecret(e.target.value)}
              placeholder="ADMIN_SECRET"
              style={styles.input}
              autoFocus
              required
            />
            <button type="submit" disabled={loading} style={styles.btnPrimary}>
              {loading ? "検証中..." : "ログイン"}
            </button>
          </form>
        )}

        {/* Phase 2: Microsoft ログイン */}
        {phase === "authenticated" && (
          <div>
            {status?.authenticated ? (
              <div>
                <div style={styles.statusBox}>
                  <div style={styles.statusRow}>
                    <span style={styles.statusLabel}>ステータス</span>
                    <span style={{ ...styles.badge, background: "#DFF6DD", color: "#107C10" }}>認証済み</span>
                  </div>
                  <div style={styles.statusRow}>
                    <span style={styles.statusLabel}>ユーザー</span>
                    <span>{status.user?.name} ({status.user?.email})</span>
                  </div>
                  <div style={styles.statusRow}>
                    <span style={styles.statusLabel}>MCP API キー</span>
                    <span>{status.mcpApiKeyHint || "未設定"}</span>
                  </div>
                  <div style={styles.statusRow}>
                    <span style={styles.statusLabel}>最終更新</span>
                    <span>{status.tokenUpdatedAt ? new Date(status.tokenUpdatedAt).toLocaleString("ja-JP") : "-"}</span>
                  </div>
                  <div style={styles.statusRow}>
                    <span style={styles.statusLabel}>最終 Cron</span>
                    <span>{status.lastCronExecution ? new Date(status.lastCronExecution).toLocaleString("ja-JP") : "未実行"}</span>
                  </div>
                </div>
                <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
                  <button onClick={handleMicrosoftLogin} style={styles.btnOutline}>
                    再認証（トークン更新）
                  </button>
                  <button onClick={handleRotateKey} style={styles.btnDanger} disabled={loading}>
                    APIキーを再発行
                  </button>
                </div>
              </div>
            ) : (
              <div>
                <p style={styles.desc}>
                  Microsoft アカウントでログインして、Graph API への<br />
                  アクセスを許可してください。
                </p>
                <button onClick={handleMicrosoftLogin} style={styles.btnPrimary}>
                  Microsoft でログイン
                </button>
              </div>
            )}
          </div>
        )}

        {/* Phase 3: セットアップ完了 */}
        {phase === "setup_complete" && (
          <div>
            <div style={{ ...styles.statusBox, borderColor: "#107C10" }}>
              <p style={{ margin: 0, fontWeight: 600, color: "#107C10" }}>✅ セットアップ完了</p>
            </div>

            {newApiKey && (
              <div style={styles.keyBox}>
                <p style={{ margin: "0 0 8px", fontWeight: 600, fontSize: 14 }}>
                  🔑 MCP API キー（この画面でのみ表示されます）
                </p>
                <div style={styles.keyDisplay}>
                  <code style={styles.keyCode}>{newApiKey}</code>
                  <button onClick={handleCopy} style={styles.btnCopy}>
                    {copied ? "✓ コピー済み" : "コピー"}
                  </button>
                </div>
                <p style={{ margin: "8px 0 0", fontSize: 12, color: "#616161" }}>
                  このキーを MCP クライアントの Bearer Token に設定してください。<br />
                  ページを閉じると再表示できません（再発行は可能です）。
                </p>
              </div>
            )}

            {status && (
              <div style={{ ...styles.statusBox, marginTop: 16 }}>
                <div style={styles.statusRow}>
                  <span style={styles.statusLabel}>ユーザー</span>
                  <span>{status.user?.name} ({status.user?.email})</span>
                </div>
                <div style={styles.statusRow}>
                  <span style={styles.statusLabel}>MCP API キー</span>
                  <span>{status.mcpApiKeyHint}</span>
                </div>
                <div style={styles.statusRow}>
                  <span style={styles.statusLabel}>Cron</span>
                  <span>{status.cronSchedule}</span>
                </div>
              </div>
            )}

            <div style={{ marginTop: 16 }}>
              <p style={{ fontSize: 13, color: "#616161", lineHeight: 1.7 }}>
                <strong>MCP クライアントの設定:</strong><br />
                URL: <code style={styles.inlineCode}>{typeof window !== "undefined" ? window.location.origin : ""}/api/mcp</code><br />
                ヘッダー: <code style={styles.inlineCode}>Authorization: Bearer &lt;上記のキー&gt;</code>
              </p>
            </div>

            <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
              <button onClick={() => { setPhase("authenticated"); fetchStatus(); }} style={styles.btnOutline}>
                管理画面に戻る
              </button>
              <button onClick={handleRotateKey} style={styles.btnDanger} disabled={loading}>
                APIキーを再発行
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  container: {
    minHeight: "100vh",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    background: "#fafafa",
    fontFamily: "'Segoe UI', -apple-system, sans-serif",
    padding: 20,
  },
  card: {
    background: "#fff",
    border: "1px solid #e0e0e0",
    borderRadius: 12,
    padding: 32,
    maxWidth: 520,
    width: "100%",
    boxShadow: "0 4px 8px rgba(0,0,0,0.04)",
  },
  logo: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    marginBottom: 24,
  },
  title: {
    fontSize: 22,
    fontWeight: 700,
    margin: "0 0 16px",
    color: "#242424",
  },
  desc: {
    fontSize: 14,
    color: "#616161",
    margin: "0 0 16px",
    lineHeight: 1.6,
  },
  input: {
    width: "100%",
    padding: "10px 14px",
    border: "1px solid #e0e0e0",
    borderRadius: 4,
    fontSize: 14,
    marginBottom: 12,
    outline: "none",
    fontFamily: "inherit",
  },
  btnPrimary: {
    width: "100%",
    padding: "10px 20px",
    background: "#0078D4",
    color: "#fff",
    border: "none",
    borderRadius: 4,
    fontSize: 14,
    fontWeight: 600,
    cursor: "pointer",
    fontFamily: "inherit",
  },
  btnOutline: {
    flex: 1,
    padding: "8px 16px",
    background: "#fff",
    color: "#242424",
    border: "1px solid #e0e0e0",
    borderRadius: 4,
    fontSize: 13,
    fontWeight: 600,
    cursor: "pointer",
    fontFamily: "inherit",
  },
  btnDanger: {
    flex: 1,
    padding: "8px 16px",
    background: "#fff",
    color: "#D83B01",
    border: "1px solid #D83B01",
    borderRadius: 4,
    fontSize: 13,
    fontWeight: 600,
    cursor: "pointer",
    fontFamily: "inherit",
  },
  btnCopy: {
    padding: "6px 14px",
    background: "#0078D4",
    color: "#fff",
    border: "none",
    borderRadius: 4,
    fontSize: 12,
    fontWeight: 600,
    cursor: "pointer",
    whiteSpace: "nowrap",
    fontFamily: "inherit",
  },
  error: {
    background: "#FDE7E9",
    color: "#A4262C",
    padding: "10px 14px",
    borderRadius: 4,
    fontSize: 13,
    marginBottom: 16,
    border: "1px solid #F1BBBC",
  },
  statusBox: {
    background: "#f5f5f5",
    border: "1px solid #e0e0e0",
    borderRadius: 8,
    padding: 16,
  },
  statusRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "6px 0",
    fontSize: 13,
    borderBottom: "1px solid #ebebeb",
  },
  statusLabel: {
    color: "#8A8A8A",
    fontWeight: 500,
  },
  badge: {
    padding: "2px 8px",
    borderRadius: 100,
    fontSize: 11,
    fontWeight: 600,
  },
  keyBox: {
    background: "#FFF8F0",
    border: "1px solid #FFE0C2",
    borderRadius: 8,
    padding: 16,
    marginTop: 16,
  },
  keyDisplay: {
    display: "flex",
    gap: 8,
    alignItems: "center",
  },
  keyCode: {
    flex: 1,
    background: "#1e1e1e",
    color: "#4EC9B0",
    padding: "8px 12px",
    borderRadius: 4,
    fontSize: 11,
    fontFamily: "Consolas, 'Fira Code', monospace",
    wordBreak: "break-all" as const,
    lineHeight: 1.4,
  },
  inlineCode: {
    background: "#f5f5f5",
    padding: "2px 6px",
    borderRadius: 3,
    fontSize: 12,
    fontFamily: "Consolas, 'Fira Code', monospace",
  },
};
