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

/* ── コピーボタン付きコードブロック ── */
function CopyBlock({ label, value, mono = true }: { label: string; value: string; mono?: boolean }) {
  const [copied, setCopied] = useState(false);
  const go = () => { navigator.clipboard.writeText(value); setCopied(true); setTimeout(() => setCopied(false), 2000); };
  return (
    <div style={{ marginBottom: 12 }}>
      <div style={{ fontSize: 11, color: "#8A8A8A", marginBottom: 4, fontWeight: 600 }}>{label}</div>
      <div style={{ display: "flex", gap: 6, alignItems: "stretch" }}>
        <code style={{
          flex: 1, background: "#1e1e1e", color: "#4EC9B0", padding: "8px 10px", borderRadius: 4,
          fontSize: mono ? 11 : 12, fontFamily: mono ? "Consolas,'Fira Code',monospace" : "inherit",
          wordBreak: "break-all" as const, lineHeight: 1.5, whiteSpace: "pre-wrap", display: "block", overflowX: "auto" as const,
        }}>{value}</code>
        <button onClick={go} style={{
          padding: "0 12px", background: copied ? "#107C10" : "#0078D4", color: "#fff", border: "none",
          borderRadius: 4, fontSize: 11, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" as const, fontFamily: "inherit", minWidth: 60,
        }}>{copied ? "✓" : "コピー"}</button>
      </div>
    </div>
  );
}

/* ── アコーディオンカード ── */
function ClientCard({ icon, name, desc, children, open: defaultOpen = false }: {
  icon: string; name: string; desc: string; children: React.ReactNode; open?: boolean;
}) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div style={{ border: "1px solid #e0e0e0", borderRadius: 8, marginBottom: 8, overflow: "hidden" }}>
      <button onClick={() => setOpen(!open)} style={{
        width: "100%", display: "flex", alignItems: "center", gap: 10, padding: "12px 14px",
        background: open ? "#f5f5f5" : "#fff", border: "none", cursor: "pointer", fontFamily: "inherit", textAlign: "left" as const,
      }}>
        <span style={{ fontSize: 20 }}>{icon}</span>
        <div style={{ flex: 1 }}>
          <div style={{ fontWeight: 600, fontSize: 14, color: "#242424" }}>{name}</div>
          <div style={{ fontSize: 12, color: "#8A8A8A" }}>{desc}</div>
        </div>
        <span style={{ fontSize: 18, color: "#8A8A8A", transition: "transform .2s", transform: open ? "rotate(180deg)" : "rotate(0)" }}>▾</span>
      </button>
      {open && <div style={{ padding: "12px 14px", borderTop: "1px solid #e0e0e0", background: "#fafafa" }}>{children}</div>}
    </div>
  );
}

const IC: React.CSSProperties = { background: "#f0f0f0", padding: "1px 5px", borderRadius: 3, fontSize: 11, fontFamily: "Consolas,'Fira Code',monospace" };

export default function AuthPage() {
  const [phase, setPhase] = useState<Phase>("password");
  const [secret, setSecret] = useState("");
  const [sessionToken, setSessionToken] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [newApiKey, setNewApiKey] = useState("");
  const [status, setStatus] = useState<AuthStatus | null>(null);
  const [keyCopied, setKeyCopied] = useState(false);

  const baseUrl = typeof window !== "undefined" ? window.location.origin : "";

  useEffect(() => {
    const p = new URLSearchParams(window.location.search);
    if (p.get("error")) setError(p.get("error")!);
    if (p.get("success") === "true" && p.get("session")) {
      setSessionToken(p.get("session")!);
      setPhase("setup_complete");
      if (p.get("newKey")) setNewApiKey(p.get("newKey")!);
      window.history.replaceState({}, "", "/auth");
    }
  }, []);

  const fetchStatus = useCallback(async () => {
    if (!sessionToken) return;
    try { const r = await fetch("/api/auth/status", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ session: sessionToken }) }); if (r.ok) setStatus(await r.json()); } catch {}
  }, [sessionToken]);

  useEffect(() => { if (phase === "authenticated" || phase === "setup_complete") fetchStatus(); }, [phase, fetchStatus]);

  async function handleVerify(e: React.FormEvent) {
    e.preventDefault(); setError(""); setLoading(true);
    try {
      const r = await fetch("/api/auth/verify", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ secret }) });
      const d = await r.json();
      if (!r.ok) { setError(d.remaining !== undefined ? `${d.error} (残り${d.remaining}回)` : d.error || "認証失敗"); return; }
      setSessionToken(d.sessionToken); setPhase("authenticated");
    } catch { setError("サーバーに接続できません"); } finally { setLoading(false); }
  }

  function handleMicrosoftLogin() { window.location.href = `/api/auth/login?session=${sessionToken}`; }

  async function handleRotateKey() {
    if (!confirm("新しい MCP API キーを発行しますか？\n古いキーは即座に無効化されます。")) return;
    setLoading(true);
    try {
      const r = await fetch("/api/auth/rotate-key", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ session: sessionToken }) });
      const d = await r.json();
      if (r.ok) { setNewApiKey(d.mcpApiKey); setKeyCopied(false); fetchStatus(); } else { setError(d.error); }
    } catch { setError("キーのローテーションに失敗しました"); } finally { setLoading(false); }
  }

  /* ── 設定値 ── */
  const mcpUrl = `${baseUrl}/api/mcp`;
  const mcpUrlKey = `${mcpUrl}?key=${newApiKey || "<APIキー>"}`;
  const claudeDesktopJson = JSON.stringify({ mcpServers: { microsoft365: { type: "url", url: mcpUrlKey } } }, null, 2);
  const cursorJson = JSON.stringify({ mcpServers: { microsoft365: { url: mcpUrlKey } } }, null, 2);
  const vscodeJson = JSON.stringify({ servers: { microsoft365: { type: "http", url: mcpUrlKey } } }, null, 2);
  const claudeCodeCmd = `claude mcp add --transport http microsoft365 "${mcpUrlKey}"`;

  /* ── MCP 接続ガイド ── */
  function McpGuide() {
    return (
      <div style={{ marginTop: 20 }}>
        <h2 style={{ fontSize: 16, fontWeight: 700, margin: "0 0 4px", color: "#242424" }}>MCP クライアント接続ガイド</h2>
        <p style={{ fontSize: 12, color: "#8A8A8A", margin: "0 0 12px" }}>お使いのクライアントを選んで、設定をコピーしてください</p>

        <ClientCard icon="🌐" name="Claude.ai（Web / モバイル）" desc="ブラウザまたはモバイルアプリから利用" open={true}>
          <p style={{ fontSize: 12, color: "#616161", margin: "0 0 10px", lineHeight: 1.6 }}>
            設定 → <strong>コネクタ</strong> → <strong>カスタムコネクタを追加</strong> から以下を入力:
          </p>
          <CopyBlock label="名前" value="Microsoft 365" mono={false} />
          <CopyBlock label="リモートMCPサーバーURL" value={mcpUrlKey} />
          <p style={{ fontSize: 11, color: "#8A8A8A", margin: "4px 0 0" }}>
            ※ OAuth Client ID / クライアントシークレット は<strong>空のまま</strong> →「追加」<br />
            ※ 追加後、<strong>新しいチャット</strong>を開いてください（既存チャットには反映されません）
          </p>
        </ClientCard>

        <ClientCard icon="🖥️" name="Claude Desktop" desc="macOS / Windows デスクトップアプリ">
          <p style={{ fontSize: 12, color: "#616161", margin: "0 0 10px", lineHeight: 1.6 }}>
            <code style={IC}>claude_desktop_config.json</code> に以下を追加して再起動:
          </p>
          <CopyBlock label="claude_desktop_config.json" value={claudeDesktopJson} />
          <p style={{ fontSize: 11, color: "#8A8A8A", margin: "4px 0 0" }}>
            macOS: <code style={IC}>~/Library/Application Support/Claude/</code> / Windows: <code style={IC}>%APPDATA%\Claude\</code>
          </p>
        </ClientCard>

        <ClientCard icon="⌨️" name="Claude Code（CLI）" desc="ターミナルから利用">
          <CopyBlock label="コマンド" value={claudeCodeCmd} />
          <p style={{ fontSize: 11, color: "#8A8A8A", margin: "4px 0 0" }}>追加後 <code style={IC}>/mcp</code> で接続確認</p>
        </ClientCard>

        <ClientCard icon="📝" name="Cursor" desc="AI コードエディタ">
          <p style={{ fontSize: 12, color: "#616161", margin: "0 0 10px", lineHeight: 1.6 }}>
            Settings → MCP → Add new MCP server、または <code style={IC}>.cursor/mcp.json</code>:
          </p>
          <CopyBlock label=".cursor/mcp.json" value={cursorJson} />
        </ClientCard>

        <ClientCard icon="💻" name="VS Code（Copilot Chat MCP）" desc="GitHub Copilot MCP 対応">
          <p style={{ fontSize: 12, color: "#616161", margin: "0 0 10px", lineHeight: 1.6 }}>
            <code style={IC}>.vscode/mcp.json</code> に追加:
          </p>
          <CopyBlock label=".vscode/mcp.json" value={vscodeJson} />
        </ClientCard>

        <ClientCard icon="🤖" name="ChatGPT / その他の MCP クライアント" desc="汎用 HTTP MCP 接続">
          <CopyBlock label="MCP サーバー URL（キー付き）" value={mcpUrlKey} />
          <CopyBlock label="MCP サーバー URL（Bearer Token 方式）" value={mcpUrl} />
          <p style={{ fontSize: 11, color: "#8A8A8A", margin: "4px 0 0" }}>
            Bearer Token 方式: ヘッダーに <code style={IC}>Authorization: Bearer {newApiKey || "<APIキー>"}</code><br />
            Bearer が使えない場合は <code style={IC}>?key=</code> パラメータで認証
          </p>
        </ClientCard>
      </div>
    );
  }

  return (
    <div style={S.container}>
      <div style={S.card}>
        <div style={S.logo}>
          <svg width="20" height="20" viewBox="0 0 16 16" fill="none"><rect width="7" height="7" fill="#F25022"/><rect x="9" width="7" height="7" fill="#7FBA00"/><rect y="9" width="7" height="7" fill="#00A4EF"/><rect x="9" y="9" width="7" height="7" fill="#FFB900"/></svg>
          <span style={{ fontWeight: 600, fontSize: 16 }}>msgraph-mcp-server</span>
        </div>
        <h1 style={S.title}>管理パネル</h1>
        {error && <div style={S.error}>{error}</div>}

        {phase === "password" && (
          <form onSubmit={handleVerify}>
            <p style={S.desc}>管理パスワードを入力してください</p>
            <input type="password" value={secret} onChange={e => setSecret(e.target.value)} placeholder="ADMIN_SECRET" style={S.input} autoFocus required />
            <button type="submit" disabled={loading} style={S.btnPrimary}>{loading ? "検証中..." : "ログイン"}</button>
          </form>
        )}

        {phase === "authenticated" && (
          <div>
            {status?.authenticated ? (
              <div>
                <div style={S.statusBox}>
                  <div style={S.statusRow}><span style={S.sLabel}>ステータス</span><span style={{ ...S.badge, background: "#DFF6DD", color: "#107C10" }}>認証済み</span></div>
                  <div style={S.statusRow}><span style={S.sLabel}>ユーザー</span><span>{status.user?.name} ({status.user?.email})</span></div>
                  <div style={S.statusRow}><span style={S.sLabel}>MCP API キー</span><span>{status.mcpApiKeyHint || "未設定"}</span></div>
                  <div style={S.statusRow}><span style={S.sLabel}>最終更新</span><span>{status.tokenUpdatedAt ? new Date(status.tokenUpdatedAt).toLocaleString("ja-JP") : "-"}</span></div>
                  <div style={S.statusRow}><span style={S.sLabel}>最終 Cron</span><span>{status.lastCronExecution ? new Date(status.lastCronExecution).toLocaleString("ja-JP") : "未実行"}</span></div>
                </div>
                <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
                  <button onClick={handleMicrosoftLogin} style={S.btnOutline}>再認証</button>
                  <button onClick={handleRotateKey} style={S.btnDanger} disabled={loading}>APIキーを再発行</button>
                </div>
                {newApiKey && (
                  <div style={S.keyBox}>
                    <p style={{ margin: "0 0 8px", fontWeight: 600, fontSize: 14 }}>🔑 MCP API キー（この画面でのみ表示されます）</p>
                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                      <code style={S.keyCode}>{newApiKey}</code>
                      <button onClick={() => { navigator.clipboard.writeText(newApiKey); setKeyCopied(true); setTimeout(() => setKeyCopied(false), 3000); }} style={S.btnCopy}>{keyCopied ? "✓ コピー済み" : "キーをコピー"}</button>
                    </div>
                    <p style={{ margin: "8px 0 0", fontSize: 12, color: "#616161" }}>ページを閉じると再表示できません（再発行は可能です）。</p>
                  </div>
                )}
                {newApiKey && <McpGuide />}
              </div>
            ) : (
              <div>
                <p style={S.desc}>Microsoft アカウントでログインして、Graph API への<br/>アクセスを許可してください。</p>
                <button onClick={handleMicrosoftLogin} style={S.btnPrimary}>Microsoft でログイン</button>
              </div>
            )}
          </div>
        )}

        {phase === "setup_complete" && (
          <div>
            <div style={{ ...S.statusBox, borderColor: "#107C10" }}><p style={{ margin: 0, fontWeight: 600, color: "#107C10" }}>✅ セットアップ完了</p></div>
            {newApiKey && (
              <div style={S.keyBox}>
                <p style={{ margin: "0 0 8px", fontWeight: 600, fontSize: 14 }}>🔑 MCP API キー（この画面でのみ表示されます）</p>
                <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                  <code style={S.keyCode}>{newApiKey}</code>
                  <button onClick={() => { navigator.clipboard.writeText(newApiKey); setKeyCopied(true); setTimeout(() => setKeyCopied(false), 3000); }} style={S.btnCopy}>{keyCopied ? "✓ コピー済み" : "キーをコピー"}</button>
                </div>
                <p style={{ margin: "8px 0 0", fontSize: 12, color: "#616161" }}>ページを閉じると再表示できません（再発行は可能です）。</p>
              </div>
            )}
            {status && (
              <div style={{ ...S.statusBox, marginTop: 16 }}>
                <div style={S.statusRow}><span style={S.sLabel}>ユーザー</span><span>{status.user?.name} ({status.user?.email})</span></div>
                <div style={S.statusRow}><span style={S.sLabel}>MCP API キー</span><span>{status.mcpApiKeyHint}</span></div>
                <div style={S.statusRow}><span style={S.sLabel}>Cron</span><span>{status.cronSchedule}</span></div>
              </div>
            )}
            <McpGuide />
            <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
              <button onClick={() => { setPhase("authenticated"); fetchStatus(); }} style={S.btnOutline}>管理画面に戻る</button>
              <button onClick={handleRotateKey} style={S.btnDanger} disabled={loading}>APIキーを再発行</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  container: { minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#fafafa", fontFamily: "'Segoe UI',-apple-system,sans-serif", padding: 20 },
  card: { background: "#fff", border: "1px solid #e0e0e0", borderRadius: 12, padding: 32, maxWidth: 620, width: "100%", boxShadow: "0 4px 8px rgba(0,0,0,0.04)" },
  logo: { display: "flex", alignItems: "center", gap: 10, marginBottom: 24 },
  title: { fontSize: 22, fontWeight: 700, margin: "0 0 16px", color: "#242424" },
  desc: { fontSize: 14, color: "#616161", margin: "0 0 16px", lineHeight: 1.6 },
  input: { width: "100%", padding: "10px 14px", border: "1px solid #e0e0e0", borderRadius: 4, fontSize: 14, marginBottom: 12, outline: "none", fontFamily: "inherit", boxSizing: "border-box" as const },
  btnPrimary: { width: "100%", padding: "10px 20px", background: "#0078D4", color: "#fff", border: "none", borderRadius: 4, fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" },
  btnOutline: { flex: 1, padding: "8px 16px", background: "#fff", color: "#242424", border: "1px solid #e0e0e0", borderRadius: 4, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" },
  btnDanger: { flex: 1, padding: "8px 16px", background: "#fff", color: "#D83B01", border: "1px solid #D83B01", borderRadius: 4, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" },
  btnCopy: { padding: "6px 14px", background: "#0078D4", color: "#fff", border: "none", borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" as const, fontFamily: "inherit" },
  error: { background: "#FDE7E9", color: "#A4262C", padding: "10px 14px", borderRadius: 4, fontSize: 13, marginBottom: 16, border: "1px solid #F1BBBC" },
  statusBox: { background: "#f5f5f5", border: "1px solid #e0e0e0", borderRadius: 8, padding: 16 },
  statusRow: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", fontSize: 13, borderBottom: "1px solid #ebebeb" },
  sLabel: { color: "#8A8A8A", fontWeight: 500 },
  badge: { padding: "2px 8px", borderRadius: 100, fontSize: 11, fontWeight: 600 },
  keyBox: { background: "#FFF8F0", border: "1px solid #FFE0C2", borderRadius: 8, padding: 16, marginTop: 16 },
  keyCode: { flex: 1, background: "#1e1e1e", color: "#4EC9B0", padding: "8px 12px", borderRadius: 4, fontSize: 11, fontFamily: "Consolas,'Fira Code',monospace", wordBreak: "break-all" as const, lineHeight: 1.4 },
};
