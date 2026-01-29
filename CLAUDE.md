# CLAUDE.md

## Project Overview

MCP server for Microsoft 365 email operations (Project ID: 340006-SD-99).

This server provides Claude Code integration with Microsoft 365 via the Microsoft Graph API, enabling email reading, sending, draft management, and folder operations.

## Architecture

```
src/mcp_m365/
├── __init__.py           # Entry point (main function)
├── __main__.py           # CLI entry (python -m mcp_m365)
├── server.py             # MCP server class with stdio transport
├── auth/
│   ├── oauth.py          # Azure AD OAuth2 with PKCE
│   └── token_store.py    # Encrypted token persistence
├── graph/
│   └── client.py         # Microsoft Graph API client
└── tools/
    ├── __init__.py       # Aggregates ALL_TOOLS
    ├── auth.py           # Auth status, connect, disconnect
    ├── messages.py       # Read, list, search emails
    ├── send.py           # Send emails, reply, forward
    ├── drafts.py         # Draft management
    └── folders.py        # Folder operations
```

## Key Configuration

- **OAuth callback port**: 8743 (different from Xero's 8742)
- **Token storage**: `~/.m365/tokens.enc`
- **Microsoft Graph base URL**: `https://graph.microsoft.com/v1.0`

## Development

```bash
# Install in development mode
pip install -e .

# Run directly
python -m mcp_m365

# Or via entry point
mcp-m365
```

## Tool Naming Convention

All tools use the prefix `m365_` followed by action and resource:
- `m365_auth_status` - Authentication status
- `m365_list_messages` - List messages
- `m365_send_message` - Send message
- `m365_create_draft` - Create draft

## Related Files

- Xero MCP server (pattern reference): `340005-SD-99-MCP-Xero-Server/`
