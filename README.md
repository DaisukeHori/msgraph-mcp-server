# msgraph-mcp-server

**Microsoft 365 を AI エージェントから操作する MCP サーバー**

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https%3A%2F%2Fgithub.com%2FDaisukeHori%2Fmsgraph-mcp-server&env=AUTH_MODE%2CMICROSOFT_CLIENT_ID%2CMICROSOFT_CLIENT_SECRET%2CMICROSOFT_TENANT_ID&envDescription=AUTH_MODE%3A+graph_token%28%E3%83%87%E3%83%95%E3%82%A9%E3%83%AB%E3%83%88%29+or+client_credentials+or+api_key+%7C+Azure+AD+%E3%82%A2%E3%83%97%E3%83%AA%E8%A8%AD%E5%AE%9A&envLink=https%3A%2F%2Fgithub.com%2FDaisukeHori%2Fmsgraph-mcp-server%23%E8%AA%8D%E8%A8%BC%E3%83%A2%E3%83%BC%E3%83%89&project-name=msgraph-mcp-server&repository-name=msgraph-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

> **エンドポイント:** `https://your-app.vercel.app/api/mcp`
> **LP:** [daisukehori.github.io/msgraph-mcp-server](https://daisukehori.github.io/msgraph-mcp-server/)

Exchange・Teams・OneDrive・SharePoint の **45 MCP ツール**を提供。メール送信、予定作成、ファイル管理、SharePoint リスト操作 — すべて自然言語で AI エージェントから実行可能。

## 2 つの使い方

### 🌐 Vercel にデプロイして使う（推奨）

上の「Deploy with Vercel」ボタンをワンクリック。デプロイ後、MCP クライアントから接続するだけ。

```
「受信トレイの未読メール一覧を教えて」
「明日 14:00 に営業会議を作成して」
「Teams の #general チャネルに進捗報告を送って」
「OneDrive の /Documents/報告書.xlsx をダウンロードして」
「SharePoint サイトのタスクリストに新しいアイテムを追加して」
```

### 💻 ローカルで使う（stdio）

Claude Desktop / Claude Code からローカル実行。

```bash
git clone https://github.com/DaisukeHori/msgraph-mcp-server.git
cd msgraph-mcp-server
npm install
npx tsx lib/stdio.ts
```

## アーキテクチャ

```
┌──────────────────────────────────────────┐
│  Next.js App Router (Vercel)             │
│  /api/mcp  → Streamable HTTP            │
│  /api/sse  → SSE (後方互換)              │
├──────────────────────────────────────────┤
│  認証コンテキスト (AsyncLocalStorage)      │
│  3 モード: graph_token / client_creds /  │
│            api_key                       │
├──────────────────────────────────────────┤
│  MCP Tools (45 ツール)                    │
│  Mail(8) / Calendar(5) / Teams(8) /     │
│  OneDrive(9) / SharePoint(12) / User(3) │
├──────────────────────────────────────────┤
│  Microsoft Graph API Client              │
│  graph.microsoft.com/v1.0               │
└──────────────────────────────────────────┘
```

## セットアップ

### ステップ 1: Azure AD アプリを作る

1. [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **アプリの登録** → **新規登録**
2. 名前: `msgraph-mcp-server`（任意）
3. サポートされているアカウントの種類: **この組織ディレクトリのみ**
4. **登録** をクリック
5. **アプリケーション (クライアント) ID** をメモ → `MICROSOFT_CLIENT_ID`
6. **ディレクトリ (テナント) ID** をメモ → `MICROSOFT_TENANT_ID`
7. **証明書とシークレット** → **新しいクライアント シークレット** を作成 → `MICROSOFT_CLIENT_SECRET`

### ステップ 2: API アクセス許可を付与

**API のアクセス許可** → **アクセス許可の追加** → **Microsoft Graph** → **アプリケーションのアクセス許可**:

```
Mail.Read, Mail.ReadWrite, Mail.Send
Calendars.Read, Calendars.ReadWrite
Team.ReadBasic.All, Channel.ReadBasic.All
ChannelMessage.Read.All
Chat.Read.All
Files.Read.All, Files.ReadWrite.All
Sites.Read.All, Sites.ReadWrite.All
User.Read.All
```

→ **[組織名] に管理者の同意を与えます** をクリック

### ステップ 3: デプロイまたはローカル起動

#### Vercel（推奨）

上の **Deploy with Vercel** ボタンをクリック → 環境変数を入力 → デプロイ完了

#### ローカル（stdio）

```bash
export AUTH_MODE=client_credentials
export MICROSOFT_CLIENT_ID=your-client-id
export MICROSOFT_CLIENT_SECRET=your-client-secret
export MICROSOFT_TENANT_ID=your-tenant-id

npx tsx lib/stdio.ts
```

### ステップ 4: MCP クライアントに接続

#### Claude.ai Web（Vercel デプロイ後）

設定 → 接続 → MCP サーバーを追加:
```
URL: https://your-app.vercel.app/api/mcp
```

#### Claude Desktop（ローカル）

`claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "msgraph": {
      "command": "npx",
      "args": ["tsx", "/path/to/msgraph-mcp-server/lib/stdio.ts"],
      "env": {
        "AUTH_MODE": "client_credentials",
        "MICROSOFT_CLIENT_ID": "your-client-id",
        "MICROSOFT_CLIENT_SECRET": "your-client-secret",
        "MICROSOFT_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

#### Claude Code

```bash
claude mcp add msgraph \
  -e AUTH_MODE=client_credentials \
  -e MICROSOFT_CLIENT_ID=your-client-id \
  -e MICROSOFT_CLIENT_SECRET=your-client-secret \
  -e MICROSOFT_TENANT_ID=your-tenant-id \
  -- npx tsx /path/to/msgraph-mcp-server/lib/stdio.ts
```

## 認証モード

| モード | 用途 | トークン渡し方 |
|:--|:--|:--|
| `graph_token`（デフォルト） | ユーザーが自分のトークンを渡す | Bearer Token or `?token=` |
| `client_credentials` | サーバー対サーバー認証 | 環境変数で Azure AD 設定 |
| `api_key` | MCP サーバーへのアクセス制限 + client_credentials | Bearer Token or `?key=` |

## ツール一覧（45 ツール）

### Exchange / メール（8 ツール）

| ツール | 操作 |
|:--|:--|
| `mail_list_messages` | メール一覧（KQL 検索対応） |
| `mail_get_message` | メール詳細取得 |
| `mail_send_message` | メール送信 |
| `mail_reply_message` | 返信 / 全員に返信 |
| `mail_update_message` | 既読/未読・重要度変更 |
| `mail_delete_message` | 削除 |
| `mail_move_message` | フォルダ移動 |
| `mail_list_folders` | フォルダ一覧 |

### カレンダー（5 ツール）

| ツール | 操作 |
|:--|:--|
| `calendar_list_events` | 予定一覧（日付範囲指定可） |
| `calendar_get_event` | 予定詳細 |
| `calendar_create_event` | 予定作成（オンライン会議対応、既定 Asia/Tokyo） |
| `calendar_update_event` | 予定更新 |
| `calendar_delete_event` | 予定削除 |

### Teams（8 ツール）

| ツール | 操作 |
|:--|:--|
| `teams_list_joined_teams` | 参加チーム一覧 |
| `teams_list_channels` | チャネル一覧 |
| `teams_list_channel_messages` | チャネルメッセージ一覧 |
| `teams_send_channel_message` | チャネルにメッセージ送信 |
| `teams_reply_to_channel_message` | メッセージに返信 |
| `teams_list_chats` | チャット一覧 |
| `teams_list_chat_messages` | チャットメッセージ一覧 |
| `teams_send_chat_message` | チャットにメッセージ送信 |

### OneDrive（9 ツール）

| ツール | 操作 |
|:--|:--|
| `onedrive_get_drive` | ドライブ情報/容量 |
| `onedrive_list_items` | ファイル/フォルダ一覧 |
| `onedrive_get_item` | アイテム詳細 |
| `onedrive_download_file` | ファイルダウンロード |
| `onedrive_upload_file` | ファイルアップロード（< 4MB） |
| `onedrive_create_folder` | フォルダ作成 |
| `onedrive_delete_item` | 削除（ゴミ箱へ） |
| `onedrive_move_item` | 移動/リネーム |
| `onedrive_search` | 検索 |

### SharePoint（12 ツール）

| ツール | 操作 |
|:--|:--|
| `sharepoint_search_sites` | サイト検索 |
| `sharepoint_get_site` | サイト詳細 |
| `sharepoint_list_drives` | ドキュメントライブラリ一覧 |
| `sharepoint_list_drive_items` | ライブラリ内ファイル一覧 |
| `sharepoint_get_lists` | リスト一覧 |
| `sharepoint_get_list_columns` | リストカラム定義 |
| `sharepoint_get_list_items` | リストアイテム取得 |
| `sharepoint_get_list_item` | リストアイテム詳細 |
| `sharepoint_create_list_item` | リストアイテム作成 |
| `sharepoint_update_list_item` | リストアイテム更新 |
| `sharepoint_delete_list_item` | リストアイテム削除 |
| `sharepoint_create_list` | リスト作成 |

### ユーザー / 認証（3 ツール）

| ツール | 操作 |
|:--|:--|
| `user_get_profile` | サインインユーザーのプロフィール |
| `user_search_users` | 組織内ユーザー検索 |
| `auth_status` | 認証モード・ステータス確認 |

## 技術スタック

- **Next.js 15** App Router
- **mcp-handler** (Streamable HTTP + SSE)
- **@modelcontextprotocol/sdk** (MCP SDK)
- **AsyncLocalStorage** (リクエストスコープ認証)
- **Zod** (入力バリデーション)
- **Microsoft Graph API v1.0**
- **Vercel** / **stdio** デュアルトランスポート

## ライセンス

MIT
