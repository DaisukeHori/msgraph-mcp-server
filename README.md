# msgraph-mcp-server

Microsoft Graph API MCP Server — Exchange / Teams / OneDrive / SharePoint の CRUD 操作を提供する MCP サーバー

## 概要

Microsoft Graph API をラップし、以下の Microsoft 365 サービスに対して完全な CRUD 操作を提供します:

| スコープ | ツール数 | 主な操作 |
|---------|---------|---------|
| **Exchange (Mail)** | 8 | メール一覧・取得・送信・返信・更新・削除・移動・フォルダ一覧 |
| **Calendar** | 5 | 予定一覧・取得・作成・更新・削除 |
| **Teams** | 8 | チーム一覧・チャネル一覧・メッセージ取得/送信/返信・チャット一覧・チャットメッセージ |
| **OneDrive** | 8 | ドライブ情報・ファイル一覧/取得/ダウンロード/アップロード・フォルダ作成・削除・移動/リネーム・検索 |
| **SharePoint** | 11 | サイト検索/取得・ドキュメントライブラリ一覧・ファイル一覧・リスト一覧/カラム取得・アイテムCRUD・リスト作成 |
| **User** | 3 | プロフィール取得・ユーザー検索・認証状態 |
| **合計** | **43ツール** | |

## セットアップ

### 1. Azure AD (Entra ID) アプリ登録

1. [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**
2. 名前: `msgraph-mcp-server` (任意)
3. サポートされているアカウントの種類: **この組織ディレクトリのみ** (シングルテナント)
4. **Register** をクリック
5. **Application (client) ID** をメモ → これが `MICROSOFT_CLIENT_ID`
6. **Directory (tenant) ID** をメモ → これが `MICROSOFT_TENANT_ID`

### 2. パブリッククライアントフローを有効化

1. 登録したアプリの **Authentication** → **Advanced settings**
2. **Allow public client flows** を **Yes** に設定
3. **Save**

### 3. API permissions の追加

1. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
2. 以下のスコープを追加:

```
Mail.Read, Mail.ReadWrite, Mail.Send
Calendars.Read, Calendars.ReadWrite
Team.ReadBasic.All, Channel.ReadBasic.All
ChannelMessage.Read.All, ChannelMessage.Send
Chat.Read, Chat.ReadWrite, ChatMessage.Read, ChatMessage.Send
Files.Read.All, Files.ReadWrite.All
Sites.Read.All, Sites.ReadWrite.All
User.Read, User.ReadBasic.All
```

3. **Grant admin consent for [Organization]** をクリック

### 4. インストール・ビルド

```bash
cd msgraph-mcp-server
npm install
npm run build
```

### 5. 環境変数

```bash
export MICROSOFT_CLIENT_ID="your-application-client-id"
export MICROSOFT_TENANT_ID="your-tenant-id"  # optional, defaults to "common"
```

### 6. Claude Desktop 設定

`claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "msgraph": {
      "command": "node",
      "args": ["/path/to/msgraph-mcp-server/dist/index.js"],
      "env": {
        "MICROSOFT_CLIENT_ID": "your-client-id",
        "MICROSOFT_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

### 7. Claude Code 設定

```bash
claude mcp add msgraph -- node /path/to/msgraph-mcp-server/dist/index.js
```

環境変数は `.env` または事前に export してください。

## 認証フロー

初回アクセス時、**Device Code Flow** が発動します:

1. MCP ツールを呼ぶ（例: `user_get_profile`）
2. stderr に認証URL とコードが表示される
3. ブラウザで https://microsoft.com/devicelogin にアクセス
4. コードを入力 → サインイン → 権限許可
5. トークンが `~/.msgraph-mcp-token-cache.json` にキャッシュされる
6. 以降はサイレント更新（refresh token で自動更新）

## ツール一覧

### Mail (Exchange)
- `mail_list_messages` - メール一覧 (KQL検索対応)
- `mail_get_message` - メール詳細取得
- `mail_send_message` - メール送信
- `mail_reply_message` - 返信/全員に返信
- `mail_update_message` - 既読/未読・重要度変更
- `mail_delete_message` - 削除
- `mail_move_message` - フォルダ移動
- `mail_list_folders` - フォルダ一覧

### Calendar
- `calendar_list_events` - 予定一覧 (日付範囲指定可)
- `calendar_get_event` - 予定詳細
- `calendar_create_event` - 予定作成 (オンライン会議対応)
- `calendar_update_event` - 予定更新
- `calendar_delete_event` - 予定削除

### Teams
- `teams_list_joined_teams` - 参加チーム一覧
- `teams_list_channels` - チャネル一覧
- `teams_list_channel_messages` - チャネルメッセージ一覧
- `teams_send_channel_message` - チャネルにメッセージ送信
- `teams_reply_to_channel_message` - メッセージに返信
- `teams_list_chats` - チャット一覧
- `teams_list_chat_messages` - チャットメッセージ一覧
- `teams_send_chat_message` - チャットにメッセージ送信

### OneDrive
- `onedrive_get_drive` - ドライブ情報/容量
- `onedrive_list_items` - ファイル/フォルダ一覧
- `onedrive_get_item` - アイテム詳細
- `onedrive_download_file` - ファイルダウンロード
- `onedrive_upload_file` - ファイルアップロード (< 4MB)
- `onedrive_create_folder` - フォルダ作成
- `onedrive_delete_item` - 削除 (ゴミ箱へ)
- `onedrive_move_item` - 移動/リネーム
- `onedrive_search` - 検索

### SharePoint
- `sharepoint_search_sites` - サイト検索
- `sharepoint_get_site` - サイト詳細
- `sharepoint_list_drives` - ドキュメントライブラリ一覧
- `sharepoint_list_drive_items` - ライブラリ内ファイル一覧
- `sharepoint_get_lists` - リスト一覧
- `sharepoint_get_list_columns` - リストカラム定義
- `sharepoint_get_list_items` - リストアイテム取得
- `sharepoint_get_list_item` - リストアイテム詳細
- `sharepoint_create_list_item` - リストアイテム作成
- `sharepoint_update_list_item` - リストアイテム更新
- `sharepoint_delete_list_item` - リストアイテム削除
- `sharepoint_create_list` - リスト作成

### User / Auth
- `user_get_profile` - サインインユーザーのプロフィール
- `user_search_users` - 組織内ユーザー検索
- `auth_status` - 認証状態確認/ログアウト

## HTTP モード

リモートサーバーとして動かす場合:

```bash
TRANSPORT=http PORT=3100 node dist/index.js
```

エンドポイント: `POST http://localhost:3100/mcp`

## ライセンス

MIT
