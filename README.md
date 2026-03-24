# msgraph-mcp-server

**本人として Microsoft 365 を AI エージェントから操作する MCP サーバー**

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https%3A%2F%2Fgithub.com%2FDaisukeHori%2Fmsgraph-mcp-server&env=ADMIN_SECRET%2CMICROSOFT_CLIENT_ID%2CMICROSOFT_CLIENT_SECRET%2CMICROSOFT_TENANT_ID&envDescription=ADMIN_SECRET%3A+%2Fauth%E7%AE%A1%E7%90%86%E3%83%91%E3%82%B9%E3%83%AF%E3%83%BC%E3%83%89+%7C+MICROSOFT_CLIENT_ID%2FSECRET%2FTENANT_ID%3A+Azure+AD%E3%82%A2%E3%83%97%E3%83%AA%EF%BC%88README%E3%81%AE%E3%82%B9%E3%83%86%E3%83%83%E3%83%971%E5%8F%82%E7%85%A7%EF%BC%89&envLink=https%3A%2F%2Fgithub.com%2FDaisukeHori%2Fmsgraph-mcp-server%23%E3%82%B9%E3%83%86%E3%83%83%E3%83%97-1-azure-ad-%E3%81%AB%E3%82%A2%E3%83%97%E3%83%AA%E3%82%92%E7%99%BB%E9%8C%B2%E3%81%99%E3%82%8B5%E5%88%86&project-name=msgraph-mcp-server&repository-name=msgraph-mcp-server&integration-ids=oac_V3R1GIpkoJorr6fqyiwdhl17&skippable-integrations=1)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

> **LP:** [daisukehori.github.io/msgraph-mcp-server](https://daisukehori.github.io/msgraph-mcp-server/)

Exchange・Teams・OneDrive・SharePoint の **45 MCP ツール**を提供。
すべて `/me/` エンドポイントを使い、**操作者本人のデータ**にアクセスします。

---

## なぜこれが必要か

| 既存の方法 | この MCP サーバー |
|:--|:--|
| Graph Explorer のトークンは1時間で切れる | **refresh_token を Redis に暗号化保存 → 自動更新で実質無期限** |
| client_credentials は管理者権限で全員のデータが見える | **委任アクセスで本人のデータだけ** |
| 認証のたびにブラウザ操作が必要 | **初回1回だけ。以降は Cron が毎日トークン更新** |

---

## アーキテクチャ

```
┌─── Vercel ──────────────────────────────────────┐
│                                                  │
│  /auth          管理画面（ADMIN_SECRET + OAuth）  │
│    → Microsoft ログイン（初回のみ）               │
│    → refresh_token を AES-256-GCM 暗号化         │
│    → Upstash Redis に保存                        │
│    → MCP API キーを自動発行                      │
│                                                  │
│  /api/mcp       MCP エンドポイント                │
│    → API キー検証                                │
│    → Redis から refresh_token → access_token     │
│    → Graph API 呼び出し（/me/ = 本人として）      │
│                                                  │
│  /api/cron/keep-alive   毎日 03:00 UTC           │
│    → refresh_token を自動更新                    │
│    → 90日カウンターを毎日リセット = 実質無期限    │
│                                                  │
│  Upstash Redis (Vercel Marketplace / 自動追加)    │
│    → 暗号化された refresh_token                   │
│    → MCP API キー                                │
│    → ブルートフォース対策カウンター               │
│                                                  │
└──────────────────────────────────────────────────┘
```

## セキュリティ

| 脅威 | 対策 |
|:--|:--|
| /auth への不正アクセス | ADMIN_SECRET（パスワード）+ Microsoft OAuth テナント制限の二重ロック |
| ADMIN_SECRET 総当たり | 5回失敗で15分ロックアウト（Redis カウント） |
| Redis の中身が見られた | refresh_token は AES-256-GCM 暗号化。暗号化キーは Vercel 環境変数のみ |
| MCP API キー漏洩 | /auth から即座にローテーション可能（古いキー即無効化） |
| refresh_token 期限切れ | 毎日 Cron で更新。90日問題を完全回避 |
| Cron 不正呼び出し | CRON_SECRET ヘッダー検証（Vercel 自動生成） |

---

## クイックスタート（5ステップ）

### ステップ 1: Azure AD にアプリを登録する（5分）

1. [Azure Portal](https://portal.azure.com) にサインイン
2. **Microsoft Entra ID** → **アプリの登録** → **＋ 新規登録**
   - 名前: `msgraph-mcp-server`
   - アカウントの種類: **この組織ディレクトリのみ**
3. 登録完了画面で以下をメモ:
   - **アプリケーション (クライアント) ID** → `MICROSOFT_CLIENT_ID`
   - **ディレクトリ (テナント) ID** → `MICROSOFT_TENANT_ID`
4. 左メニュー → **証明書とシークレット** → **新しいクライアント シークレット** → 値をメモ → `MICROSOFT_CLIENT_SECRET`
5. 左メニュー → **API のアクセス許可** → **アクセス許可の追加** → **Microsoft Graph** → **委任されたアクセス許可**:

   `User.Read` `User.ReadBasic.All` `Mail.Read` `Mail.ReadWrite` `Mail.Send` `Calendars.Read` `Calendars.ReadWrite` `Team.ReadBasic.All` `Channel.ReadBasic.All` `ChannelMessage.Read.All` `ChannelMessage.Send` `Chat.Read` `Chat.ReadWrite` `ChatMessage.Read` `ChatMessage.Send` `Files.Read.All` `Files.ReadWrite.All` `Sites.Read.All` `Sites.ReadWrite.All` `offline_access`

6. 左メニュー → **認証** → **＋ プラットフォームを追加** → **「Web」を選択**
   - **リダイレクト URI**: `https://your-app.vercel.app/api/auth/callback`（Vercel デプロイ後に URL が確定してから入力。後からでも追加可能）
   - **フロントチャネルのログアウト URL**: 空のまま
   - **アクセストークン（暗黙的なフローに使用）**: **チェックしない**
   - **ID トークン（暗黙的およびハイブリッド フローに使用）**: **チェックしない**
   - 「構成」をクリックして保存
7. **「[組織名] に管理者の同意を与えます」** をクリック

### ステップ 2: Vercel にデプロイ（2分）

上の **Deploy with Vercel** ボタンをクリック。Upstash Redis のセットアップが自動で促されます（無料枠で十分）。

環境変数 4 つを入力:

| 変数 | 値 |
|:--|:--|
| `ADMIN_SECRET` | 自分で決めた管理パスワード |
| `MICROSOFT_CLIENT_ID` | ステップ 1 でメモしたクライアント ID |
| `MICROSOFT_CLIENT_SECRET` | ステップ 1 でメモしたシークレット |
| `MICROSOFT_TENANT_ID` | ステップ 1 でメモしたテナント ID |

### ステップ 3: リダイレクト URI を更新

デプロイ完了後、Vercel の URL（例: `https://msgraph-mcp-server-xxx.vercel.app`）が確定したら Azure Portal に戻って:

1. アプリの **認証** → Web プラットフォームの **リダイレクト URI** に以下を追加:
   ```
   https://あなたのアプリ名.vercel.app/api/auth/callback
   ```
2. 「保存」をクリック

> ⚠️ ステップ 1 で既に入力済みなら、URL 部分を実際の Vercel URL に置き換えるだけです。

### ステップ 4: /auth で認証（1分）

1. `https://your-app.vercel.app/auth` にアクセス
2. ADMIN_SECRET（管理パスワード）を入力
3. 「Microsoft でログイン」→ サインイン → 権限許可
4. **MCP API キーが画面に表示される** → コピーして保存

### ステップ 5: MCP クライアントに接続して使い始める

Claude.ai / Claude Desktop で MCP サーバーを追加:

- URL: `https://your-app.vercel.app/api/mcp`
- ヘッダー: `Authorization: Bearer <ステップ4のAPIキー>`

**以降一切何もしなくていい。** Cron が毎日トークンを自動更新し続けます。

---

## /auth 管理画面

| 機能 | 説明 |
|:--|:--|
| 初回セットアップ | Microsoft ログイン → API キー自動発行 |
| ステータス確認 | ユーザー名、トークン最終更新日、最終 Cron 実行日時 |
| API キー確認 | 末尾4文字のヒント表示 |
| キー ローテーション | 「再発行」ボタン → 古いキー即無効化 → 新キーを表示 |
| 再認証 | refresh_token 失効時に再ログイン |

---

## ツール一覧（45 ツール）

| カテゴリ | ツール数 | 主な操作 |
|:--|:--|:--|
| **Exchange / メール** | 8 | 一覧・取得・送信・返信・更新・削除・移動・フォルダ |
| **カレンダー** | 5 | 一覧・取得・作成・更新・削除（Asia/Tokyo 既定） |
| **Teams** | 8 | チーム/チャネル/チャット一覧・メッセージ取得/送信/返信 |
| **OneDrive** | 9 | ドライブ情報・一覧/取得/DL/UL/フォルダ作成/削除/移動/検索 |
| **SharePoint** | 12 | サイト検索/取得・ライブラリ・リスト/カラム/アイテム CRUD |
| **ユーザー / 認証** | 3 | プロフィール・ユーザー検索・認証ステータス |

---

## 同じ作者のプロジェクト

### [HubSpot MA MCP Server](https://github.com/DaisukeHori/hubspot-ma-mcp)

128 ツール + Knowledge Store + Claude Skill の 3 層構造。
「来月セミナーやるからよろしく」で AI がキャンペーン〜ワークフローまで一貫実行。

---

## ライセンス

MIT
