# msgraph-mcp-server

**本人として Microsoft 365 を AI エージェントから操作する MCP サーバー**

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https%3A%2F%2Fgithub.com%2FDaisukeHori%2Fmsgraph-mcp-server&env=AUTH_MODE&envDescription=AUTH_MODE%3A+token+%28Vercel%E6%8E%A8%E5%A5%A8%29&project-name=msgraph-mcp-server&repository-name=msgraph-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

> **LP:** [daisukehori.github.io/msgraph-mcp-server](https://daisukehori.github.io/msgraph-mcp-server/)

Exchange・Teams・OneDrive・SharePoint の **45 MCP ツール**を提供。すべてのツールは `/me/` エンドポイントを使い、**操作者本人のデータ**にアクセスします。

---

## 重要: 「本人として」操作する

このサーバーは**本人として**自分のメール・予定・ファイル・Teams を操作します。
「謎の管理者アプリ」としてではありません。

| 認証モード | 誰として動く | `/me/` | 主な用途 |
|:--|:--|:--|:--|
| **`delegated`** (推奨) | **あなた本人** | ✅ | ローカル (Claude Desktop / Code) |
| **`token`** | **あなた本人** | ✅ | Vercel (アクセストークンを渡す) |
| `client_credentials` | 管理者アプリ | ❌ | 自動化 (/users/{id}/ が必要) |

---

## クイックスタート: ローカルで本人として使う

### ステップ 1: Azure AD にアプリを登録する（5分）

1. **[Azure Portal](https://portal.azure.com) にサインイン**
   - あなたの Microsoft 365 アカウントでサインインします

2. **Microsoft Entra ID を開く**
   - 左メニュー → 「Microsoft Entra ID」（旧 Azure Active Directory）

3. **アプリを登録**
   - 左メニュー → 「アプリの登録」→「＋ 新規登録」
   - **名前**: `msgraph-mcp-server`（任意）
   - **サポートされているアカウントの種類**: 「この組織ディレクトリのみに含まれるアカウント」
   - **リダイレクト URI**: 空のまま
   - 「登録」をクリック

4. **2 つの値をメモ**
   - 登録完了画面で以下をコピー:
     - **アプリケーション (クライアント) ID** → `MICROSOFT_CLIENT_ID`
     - **ディレクトリ (テナント) ID** → `MICROSOFT_TENANT_ID`

5. **パブリッククライアントフローを有効化**
   - 左メニュー → 「認証」
   - 一番下の「詳細設定」セクション
   - **「パブリック クライアント フローを許可する」を「はい」**に設定
   - 「保存」

6. **API アクセス許可を追加**
   - 左メニュー → 「API のアクセス許可」
   - 「＋ アクセス許可の追加」→「Microsoft Graph」→ **「委任されたアクセス許可」**
   - 以下をすべて追加:

   | カテゴリ | スコープ |
   |:--|:--|
   | User | `User.Read`, `User.ReadBasic.All` |
   | Mail | `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` |
   | Calendar | `Calendars.Read`, `Calendars.ReadWrite` |
   | Teams | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `Chat.Read`, `Chat.ReadWrite`, `ChatMessage.Read`, `ChatMessage.Send` |
   | Files (OneDrive) | `Files.Read.All`, `Files.ReadWrite.All` |
   | Sites (SharePoint) | `Sites.Read.All`, `Sites.ReadWrite.All` |
   | その他 | `offline_access`（トークン自動更新用） |

   - 管理者の場合: **「[組織名] に管理者の同意を与えます」** をクリック
   - 管理者でない場合: テナント管理者に同意を依頼してください

### ステップ 2: クローンして起動（2分）

```bash
git clone https://github.com/DaisukeHori/msgraph-mcp-server.git
cd msgraph-mcp-server
npm install
```

### ステップ 3: 事前認証（ターミナルで 1 回だけ）

⚠️ **この手順が重要です。** Claude Desktop は MCP サーバーをバックグラウンドで起動するため、
認証画面が表示されません。**先にターミナルで認証を済ませる**必要があります。

```bash
cd msgraph-mcp-server

MICROSOFT_CLIENT_ID=ステップ1のクライアントID \
MICROSOFT_TENANT_ID=ステップ1のテナントID \
npx tsx lib/auth-setup.ts
```

以下の手順が表示されます:

```
┌─────────────────────────────────────────────────┐
│                                                 │
│  1. ブラウザで以下の URL を開く:                │
│     https://microsoft.com/devicelogin           │
│                                                 │
│  2. コードを入力: ABCD1234                      │
│                                                 │
│  3. Microsoft アカウントでサインイン             │
│                                                 │
│  4. 権限を許可                                  │
│                                                 │
└─────────────────────────────────────────────────┘
```

1. ブラウザで https://microsoft.com/devicelogin を開く
2. 表示されたコードを入力
3. あなたの Microsoft アカウントでサインイン
4. 権限を許可
5. ターミナルに「✅ 認証成功！」と表示されれば完了

トークンは `~/.msgraph-mcp-token-cache.json` にキャッシュされ、自動的に更新されます。
**通常この手順は 1 回だけ**で、以降は Claude Desktop が自動的にキャッシュを使います。

### ステップ 4: Claude Desktop に設定

`claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "msgraph": {
      "command": "npx",
      "args": ["tsx", "/path/to/msgraph-mcp-server/lib/stdio.ts"],
      "env": {
        "AUTH_MODE": "delegated",
        "MICROSOFT_CLIENT_ID": "ステップ1でメモしたクライアントID",
        "MICROSOFT_TENANT_ID": "ステップ1でメモしたテナントID"
      }
    }
  }
}
```

Claude Desktop を再起動すれば、Microsoft 365 のツールが使えます。

---

## Vercel でリモートデプロイする場合

### ステップ 1: デプロイ

上の **Deploy with Vercel** ボタンをクリック → 環境変数:
- `AUTH_MODE`: `token`

### ステップ 2: アクセストークンを取得

Vercel の場合、MCP クライアントから Bearer Token でアクセストークンを渡す必要があります。

**トークン取得方法:**
1. https://developer.microsoft.com/graph/graph-explorer にアクセス
2. 「Sign in to Graph Explorer」でサインイン
3. 左上の「Access token」タブからトークンをコピー

### ステップ 3: MCP クライアントに設定

Claude.ai Web:
```
URL: https://your-app.vercel.app/api/mcp
ヘッダー: Authorization: Bearer <コピーしたトークン>
```

> ⚠️ Graph Explorer のトークンは約 1 時間で期限切れになります。
> 本格運用には delegated モード（ローカル）が推奨です。

---

## アーキテクチャ

```
┌──────────────────────────────────────────┐
│  Next.js App Router (Vercel)             │
│  /api/mcp  → Streamable HTTP            │
│  /api/sse  → SSE (後方互換)              │
├──────────────────────────────────────────┤
│  lib/stdio.ts (ローカル)                  │
│  MSAL Device Code Flow → 本人認証        │
├──────────────────────────────────────────┤
│  認証コンテキスト (AsyncLocalStorage)      │
│  delegated / token / client_credentials  │
├──────────────────────────────────────────┤
│  MCP Tools (45 ツール)                    │
│  すべて /me/ エンドポイントを使用          │
│  = 操作者本人のデータにアクセス           │
├──────────────────────────────────────────┤
│  Microsoft Graph API v1.0                │
│  graph.microsoft.com                     │
└──────────────────────────────────────────┘
```

---

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
| `user_get_profile` | 本人のプロフィール取得（認証テスト兼用） |
| `user_search_users` | 組織内ユーザー検索 |
| `auth_status` | 認証ステータス確認 / ログアウト |

---

## ライセンス

MIT
