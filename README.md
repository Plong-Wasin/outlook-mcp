# outlook-mcp

A Model Context Protocol (MCP) server for sending and managing Microsoft Outlook email via the Microsoft Graph API.

## Features

- **Send email** — To, CC, BCC, plain text or HTML body, file attachments (max 3 MB each)
- **Save draft** — Save a message to Drafts without sending
- **Update draft** — Partially edit a saved Draft (only supplied fields change)
- **Send draft** — Transmit a saved Draft to its recipients
- **List mail folders** — Discover all folders including custom ones
- **List messages** — Browse any folder with full-text search or OData filter
- **Get message** — Read full message body (plain text or HTML) + attachment metadata
- **Reply / Forward** — Reply to or forward an existing message
- **Mark message** — Mark as read or unread
- **Delete message** — Move to Deleted Items (recoverable)
- **List contact folders** — Browse personal contact lists (folders) and their contacts
- **List contacts** — Search personal Outlook address book
- **List M365 Group members** — Find everyone in a Microsoft 365 Group
- **List distribution groups** — Find mail-enabled distribution lists in the organisation
- **List users** — Search or browse Azure AD users across the organisation

## Authentication

Uses OAuth 2.0 **Device Code Flow** — no browser redirect required. Tokens are stored locally at `.access-token.txt` and refreshed automatically.

```
auth_start  →  visit the URL, enter the code  →  auth_poll
```

> **Note:** `auth_start` should only be called when another tool returns an authentication error. Do not call it proactively — all tools load the stored token automatically.

## Setup

### 1. Install dependencies and build

```bash
npm install
npm run build
```

### 2. Register in Claude Code

Add to your project's `.mcp.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "type": "stdio",
      "command": "node",
      "args": ["C:\\Users\\staff\\Documents\\mcp\\outlook-mcp\\build\\index.js"]
    }
  }
}
```

### 3. Authenticate

Call `auth_start` once, follow the instructions, then call `auth_poll` to complete login.

## Tools

### Authentication
| Tool | Description |
|------|-------------|
| `auth_start` | Start the Microsoft authentication flow |
| `auth_poll` | Complete authentication after signing in |

### Sending & Drafts
| Tool | Description |
|------|-------------|
| `send_email` | Send an email immediately |
| `save_draft` | Save an email as a draft |
| `update_draft` | Partially edit a saved draft (only supplied fields change) |
| `send_draft` | Send a previously saved draft by its ID |

### Reading Mail
| Tool | Description |
|------|-------------|
| `list_mail_folders` | List all mail folders including custom ones |
| `list_messages` | List messages from any folder; supports search and filter |
| `get_message` | Read full message content (body + attachment metadata) |

### Mail Actions
| Tool | Description |
|------|-------------|
| `reply_message` | Reply to a message (thread history preserved) |
| `forward_message` | Forward a message to new recipients |
| `mark_message` | Mark a message as read or unread |
| `delete_message` | Move a message to Deleted Items |

### Address Book & Directory
| Tool | Description |
|------|-------------|
| `list_contact_folders` | List personal contact folders; pass `folderId` to list contacts inside |
| `list_contacts` | List personal Outlook contacts (supports search) |
| `list_group_members` | List members of a Microsoft 365 Group |
| `list_distribution_groups` | List mail-enabled distribution groups in the org |
| `list_users` | Search or list Azure AD users in the organisation |

## send_email / save_draft Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `to` | `string[]` | ✅ | Recipient email addresses |
| `subject` | `string` | ✅ | Email subject |
| `body` | `string` | ✅ | Email body content |
| `cc` | `string[]` | — | CC recipients |
| `bcc` | `string[]` | — | BCC recipients |
| `contentType` | `"text" \| "html"` | — | Body format (default: `"text"`) |
| `attachments` | `string[]` | — | Local file paths to attach (max 3 MB each) |

## Scopes

This MCP requests the following Microsoft Graph permissions:

- `Mail.Send` — Send email
- `Mail.ReadWrite` — Save drafts
- `Group.Read.All` — List group members and distribution groups
- `GroupMember.Read.All` — Read group membership
- `Directory.Read.All` — Browse the directory
- `Contacts.Read` — Read personal contacts
- `User.Read.All` — Search/list organisation users
- `offline_access` — Keep the session alive with a refresh token

## Token Storage

Tokens are stored at `outlook-mcp/.access-token.txt` (excluded from git via `.gitignore`).

## Attachment Limits

Attachments are encoded inline (base64). Files larger than **3 MB** are rejected with a clear error before any API call is made. For larger files, share a link instead.
