// ============================================================
// Mail / Exchange
// ============================================================
export interface MailMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  body?: { contentType: string; content: string };
  from?: { emailAddress: { name: string; address: string } };
  toRecipients?: { emailAddress: { name: string; address: string } }[];
  ccRecipients?: { emailAddress: { name: string; address: string } }[];
  receivedDateTime: string;
  sentDateTime?: string;
  isRead: boolean;
  importance: string;
  hasAttachments: boolean;
  webLink?: string;
  conversationId?: string;
}

export interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId?: string;
  childFolderCount: number;
  unreadItemCount: number;
  totalItemCount: number;
}

export interface SendMailPayload {
  message: {
    subject: string;
    body: { contentType: string; content: string };
    toRecipients: { emailAddress: { address: string; name?: string } }[];
    ccRecipients?: { emailAddress: { address: string; name?: string } }[];
  };
  saveToSentItems?: boolean;
}

// ============================================================
// Calendar
// ============================================================
export interface CalendarEvent {
  id: string;
  subject: string;
  bodyPreview?: string;
  body?: { contentType: string; content: string };
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location?: { displayName: string };
  organizer?: { emailAddress: { name: string; address: string } };
  attendees?: {
    emailAddress: { name: string; address: string };
    status?: { response: string };
    type: string;
  }[];
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
  webLink?: string;
  isAllDay?: boolean;
  recurrence?: unknown;
  showAs?: string;
  importance?: string;
}

// ============================================================
// Teams
// ============================================================
export interface Team {
  id: string;
  displayName: string;
  description?: string;
  isArchived?: boolean;
  webUrl?: string;
}

export interface Channel {
  id: string;
  displayName: string;
  description?: string;
  membershipType?: string;
  webUrl?: string;
}

export interface ChatMessage {
  id: string;
  messageType: string;
  createdDateTime: string;
  lastModifiedDateTime?: string;
  body: { contentType: string; content: string };
  from?: {
    user?: { displayName: string; id: string };
    application?: { displayName: string; id: string };
  };
  subject?: string;
  importance?: string;
  webUrl?: string;
  attachments?: { id: string; contentType: string; name: string }[];
}

export interface Chat {
  id: string;
  topic?: string;
  chatType: string;
  createdDateTime: string;
  lastUpdatedDateTime?: string;
  members?: { displayName: string; email?: string }[];
  webUrl?: string;
}

// ============================================================
// OneDrive / SharePoint Drive Items
// ============================================================
export interface DriveItem {
  id: string;
  name: string;
  size?: number;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl?: string;
  file?: { mimeType: string; hashes?: unknown };
  folder?: { childCount: number };
  parentReference?: {
    driveId: string;
    id: string;
    path?: string;
  };
  "@microsoft.graph.downloadUrl"?: string;
}

export interface Drive {
  id: string;
  name: string;
  driveType: string;
  owner?: { user?: { displayName: string } };
  quota?: {
    total: number;
    used: number;
    remaining: number;
    state: string;
  };
  webUrl?: string;
}

// ============================================================
// SharePoint
// ============================================================
export interface SharePointSite {
  id: string;
  name: string;
  displayName: string;
  webUrl: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  description?: string;
}

export interface SharePointList {
  id: string;
  name: string;
  displayName: string;
  description?: string;
  webUrl?: string;
  list?: {
    hidden: boolean;
    template: string;
    contentTypesEnabled: boolean;
  };
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

export interface SharePointListItem {
  id: string;
  fields: Record<string, unknown>;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  webUrl?: string;
  createdBy?: { user: { displayName: string } };
  lastModifiedBy?: { user: { displayName: string } };
}

export interface SharePointColumn {
  id: string;
  name: string;
  displayName: string;
  description?: string;
  type?: string;
  readOnly?: boolean;
  required?: boolean;
}

