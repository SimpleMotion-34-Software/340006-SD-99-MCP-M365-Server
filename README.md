# MCP M365 Server

A Model Context Protocol (MCP) server for Microsoft 365 email operations using the Microsoft Graph API.

## Features

- **Authentication**: OAuth 2.0 with PKCE for secure browser-based authentication
- **Message Operations**: List, search, read emails with attachment support
- **Send Operations**: Compose, reply, and forward emails
- **Draft Management**: Create, update, and delete drafts
- **Folder Management**: List folders, move and delete messages

## Installation

```bash
# From the project directory
pip install -e .

# Or install directly
pip install mcp-m365
```

## Azure AD App Registration

Before using this server, you need to register an application in Azure AD:

### Step 1: Create App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click "New registration"
3. Enter a name (e.g., "MCP M365 Server")
4. Select "Accounts in this organizational directory only" (or as needed)
5. Set Redirect URI to: `http://localhost:8743/callback` (type: Web)
6. Click "Register"

### Step 2: Configure API Permissions

In your app registration, go to "API permissions" and add:

| Permission | Type | Description |
|------------|------|-------------|
| `Mail.Read` | Delegated | Read user mail |
| `Mail.ReadWrite` | Delegated | Read and write user mail |
| `Mail.Send` | Delegated | Send mail as user |
| `User.Read` | Delegated | Sign in and read user profile |
| `offline_access` | Delegated | Maintain access to data |

Click "Grant admin consent" if required by your organization.

### Step 3: Get Credentials

1. Go to "Overview" and copy the **Application (client) ID**
2. Go to "Certificates & secrets" → "New client secret"
3. Copy the secret value (you won't see it again)

### Step 4: Store Credentials

Store your credentials in the system keychain:

**macOS:**
```bash
security add-generic-password -a "m365-mcp" -s "m365-client-id" -w "your-client-id"
security add-generic-password -a "m365-mcp" -s "m365-client-secret" -w "your-client-secret"
security add-generic-password -a "m365-mcp" -s "m365-tenant-id" -w "your-tenant-id"
```

**Linux (using secret-tool):**
```bash
secret-tool store --label="M365 MCP Client ID" service m365-mcp attribute M365_CLIENT_ID
secret-tool store --label="M365 MCP Client Secret" service m365-mcp attribute M365_CLIENT_SECRET
secret-tool store --label="M365 MCP Tenant ID" service m365-mcp attribute M365_TENANT_ID
```

**Environment Variables (fallback):**
```bash
export M365_CLIENT_ID="your-client-id"
export M365_CLIENT_SECRET="your-client-secret"
export M365_TENANT_ID="your-tenant-id"
```

## Claude Code Configuration

Add to your Claude Code settings (`~/.claude.json` or project settings):

```json
{
  "mcpServers": {
    "m365": {
      "command": "python",
      "args": ["-m", "mcp_m365"],
      "cwd": "/path/to/340006-SD-99-MCP-M365-Server/src"
    }
  }
}
```

Or if installed as a package:

```json
{
  "mcpServers": {
    "m365": {
      "command": "mcp-m365"
    }
  }
}
```

## Available Tools

### Authentication Tools

| Tool | Description |
|------|-------------|
| `m365_auth_status` | Check connection status and show user info |
| `m365_connect` | Initiate OAuth flow, opens browser for authentication |
| `m365_disconnect` | Clear stored tokens and disconnect |

### Message Tools

| Tool | Description |
|------|-------------|
| `m365_list_messages` | List emails with pagination, filter by folder |
| `m365_search_messages` | Search emails by subject, sender, date range |
| `m365_get_message` | Get full email content by ID |
| `m365_get_thread` | Get all messages in a conversation thread |
| `m365_get_attachment` | Download attachment by message and attachment ID |

### Send Tools

| Tool | Description |
|------|-------------|
| `m365_send_message` | Compose and send new email |
| `m365_reply` | Reply to an existing email |
| `m365_forward` | Forward an email to new recipients |

### Draft Tools

| Tool | Description |
|------|-------------|
| `m365_list_drafts` | List all draft emails |
| `m365_create_draft` | Create a new draft |
| `m365_update_draft` | Modify an existing draft |
| `m365_delete_draft` | Delete a draft |

### Folder Tools

| Tool | Description |
|------|-------------|
| `m365_list_folders` | List mail folders |
| `m365_create_folder` | Create new folder |
| `m365_move_message` | Move message to folder |
| `m365_delete_message` | Move message to Deleted Items |

## Usage Examples

### Check Authentication Status

```
Use m365_auth_status to check if connected
```

### Connect to Microsoft 365

```
Use m365_connect to authenticate - this will open your browser
```

### List Recent Emails

```
Use m365_list_messages with limit=10 to see recent emails
```

### Search for Emails

```
Use m365_search_messages with query="from:john@example.com subject:meeting"
```

### Send an Email

```
Use m365_send_message with:
- to: ["recipient@example.com"]
- subject: "Hello"
- body: "Email content here"
```

## Token Storage

Tokens are stored encrypted at `~/.m365/tokens.enc` using machine-specific encryption (Fernet with PBKDF2 key derivation). Tokens are automatically refreshed when expired.

## Rate Limiting

The server implements automatic rate limiting to stay within Microsoft Graph API limits (10,000 requests per 10 minutes). Requests are queued if limits are approached.

## License

MIT License - see LICENSE.md
