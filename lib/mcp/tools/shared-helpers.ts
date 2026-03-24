/**
 * 共有リソースアクセス用ヘルパー
 *
 * user_id が指定された場合: /users/{user_id}/... (共有メールボックス、委任カレンダー等)
 * 未指定の場合: /me/... (本人)
 *
 * 使い方:
 *   const base = userBase(params.user_id);  // "/me" or "/users/someone@example.com"
 *   const endpoint = `${base}/messages`;
 */

/**
 * ユーザーベースパスを返す
 * @param userId - ユーザー ID、UPN (email)、または共有メールボックスのメールアドレス。未指定なら "/me"
 */
export function userBase(userId?: string): string {
  if (userId && userId.trim().length > 0) {
    return `/users/${userId.trim()}`;
  }
  return "/me";
}

/**
 * user_id パラメータの共通 Zod スキーマ説明
 */
export const USER_ID_DESCRIPTION =
  "対象ユーザーの ID またはメールアドレス（UPN）。" +
  "共有メールボックスや委任カレンダーにアクセスする場合に指定。" +
  "省略すると本人（/me/）のリソースにアクセス。" +
  "例: 'shared-mailbox@revol.co.jp', 'alex@contoso.com'";
