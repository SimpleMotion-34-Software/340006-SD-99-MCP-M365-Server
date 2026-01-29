"""Authentication tools for M365 MCP server."""

from typing import Any

from mcp.types import Tool

from ..auth import M365OAuth

AUTH_TOOLS = [
    Tool(
        name="m365_auth_status",
        description="Check the current Microsoft 365 authentication status. Returns whether you're connected, if credentials are configured, and user details.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="m365_connect",
        description="Connect to Microsoft 365 using client credentials (app-only authentication). Requires m365-client-id, m365-client-secret, m365-tenant-id, and m365-user-id in keychain.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="m365_disconnect",
        description="Disconnect from Microsoft 365 by removing stored authentication tokens.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
]


async def handle_auth_tool(
    name: str, arguments: dict[str, Any], oauth: M365OAuth
) -> dict[str, Any]:
    """Handle authentication tool calls.

    Args:
        name: Tool name
        arguments: Tool arguments
        oauth: OAuth handler

    Returns:
        Tool result
    """
    if name == "m365_auth_status":
        return oauth.get_status()

    elif name == "m365_connect":
        if not oauth.is_configured:
            return {
                "error": "Microsoft credentials not configured",
                "message": "Store m365-client-id and m365-client-secret in keychain",
            }

        try:
            tokens = await oauth.authenticate_client_credentials()
            return {
                "success": True,
                "message": "Successfully connected to Microsoft 365",
                "user_email": tokens.user_email,
                "user_name": tokens.user_name,
                "scopes": tokens.scope,
            }
        except ValueError as e:
            return {"error": str(e)}

    elif name == "m365_disconnect":
        oauth.disconnect()
        return {
            "success": True,
            "message": "Disconnected from Microsoft 365. Stored tokens have been removed.",
        }

    return {"error": f"Unknown auth tool: {name}"}
