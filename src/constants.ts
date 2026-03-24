export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
export const GRAPH_BETA_URL = "https://graph.microsoft.com/beta";
export const CHARACTER_LIMIT = 25000;
export const DEFAULT_PAGE_SIZE = 25;
export const MAX_PAGE_SIZE = 100;

// Microsoft Graph API permission scopes
export const SCOPES = [
  // Mail
  "Mail.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  // Calendar
  "Calendars.Read",
  "Calendars.ReadWrite",
  // Teams
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "ChannelMessage.Send",
  "Chat.Read",
  "Chat.ReadWrite",
  "ChatMessage.Read",
  "ChatMessage.Send",
  // OneDrive
  "Files.Read.All",
  "Files.ReadWrite.All",
  // SharePoint
  "Sites.Read.All",
  "Sites.ReadWrite.All",
  // User (for context)
  "User.Read",
  "User.ReadBasic.All",
  // Offline access
  "offline_access",
];
