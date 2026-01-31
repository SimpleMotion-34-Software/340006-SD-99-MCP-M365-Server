"""Authentication tools for M365 MCP Server."""

from typing import Any, Dict, List

from mcp.types import Tool

from ..auth import M365OAuth, get_active_profile, set_active_profile, CREDENTIAL_PROFILES
from ..graph import GraphClient


AUTH_TOOLS: List[Tool] = [
    Tool(
        name="m365_auth_status",
        description="Check Microsoft 365 authentication status for the current profile",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="m365_connect",
        description="Authenticate with Microsoft 365 using device code flow. Returns a code to enter at microsoft.com/devicelogin, then polls for completion.",
        inputSchema={
            "type": "object",
            "properties": {
                "timeout": {
                    "type": "integer",
                    "description": "Maximum time to wait for authentication in seconds (default: 300)",
                    "default": 300,
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_disconnect",
        description="Disconnect from Microsoft 365 by clearing stored tokens",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="m365_set_profile",
        description="Switch to a different Microsoft 365 tenant profile",
        inputSchema={
            "type": "object",
            "properties": {
                "profile": {
                    "type": "string",
                    "description": f"Profile name. Available: {', '.join(CREDENTIAL_PROFILES.keys())}",
                    "enum": list(CREDENTIAL_PROFILES.keys()),
                },
            },
            "required": ["profile"],
        },
    ),
    Tool(
        name="m365_list_profiles",
        description="List available Microsoft 365 tenant profiles and their status",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
]


async def handle_auth_status(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_auth_status tool call."""
    return oauth.get_status()


async def handle_connect(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_connect tool call using device code flow."""
    if not oauth.is_configured:
        return {
            "error": "Credentials not configured",
            "instructions": [
                "Add credentials to keychain:",
                f'security add-generic-password -a "m365-mcp" -s "m365-{oauth.profile}-client-id" -w "YOUR_CLIENT_ID"',
                f'security add-generic-password -a "m365-mcp" -s "m365-{oauth.profile}-client-secret" -w "YOUR_SECRET"',
                f'security add-generic-password -a "m365-mcp" -s "m365-{oauth.profile}-tenant-id" -w "YOUR_TENANT_ID"',
            ],
        }

    timeout = arguments.get("timeout", 300)

    try:
        # Start device code flow
        result = await oauth.authenticate_device_code()

        # Return the device code info - the polling happens automatically
        # The user needs to complete authentication at the URL
        device_code = result["_device_code"]
        interval = result["_interval"]

        # Return initial status with instructions
        initial_response = {
            "status": "awaiting_authentication",
            "profile": oauth.profile,
            "user_code": result["user_code"],
            "verification_uri": result["verification_uri"],
            "message": f"Go to {result['verification_uri']} and enter code: {result['user_code']}",
            "expires_in": result["expires_in"],
        }

        # Poll for completion
        tokens = await oauth.poll_device_code(
            device_code=device_code,
            interval=interval,
            timeout=timeout,
        )

        return {
            "status": "connected",
            "profile": oauth.profile,
            "user_email": tokens.user_email,
            "user_name": tokens.user_name,
            "message": "Successfully authenticated with Microsoft 365",
        }
    except Exception as e:
        return {
            "error": str(e),
            "status": "failed",
        }


async def handle_disconnect(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_disconnect tool call."""
    oauth.disconnect()
    return {
        "status": "disconnected",
        "profile": oauth.profile,
        "message": "Tokens cleared",
    }


async def handle_set_profile(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_set_profile tool call."""
    profile = arguments.get("profile", "SM")

    if profile not in CREDENTIAL_PROFILES:
        return {
            "error": f"Invalid profile: {profile}",
            "available": list(CREDENTIAL_PROFILES.keys()),
        }

    set_active_profile(profile)

    # Check new profile status
    new_oauth = M365OAuth(profile)
    status = new_oauth.get_status()

    return {
        "status": "profile_changed",
        "profile": profile,
        "configured": status["configured"],
        "connected": status["connected"],
        "user_email": status.get("user_email"),
    }


async def handle_list_profiles(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_list_profiles tool call."""
    active = get_active_profile()
    profiles = []

    for profile in CREDENTIAL_PROFILES.keys():
        profile_oauth = M365OAuth(profile)
        status = profile_oauth.get_status()
        profiles.append({
            "profile": profile,
            "active": profile == active,
            "configured": status["configured"],
            "connected": status["connected"],
            "user_email": status.get("user_email"),
        })

    return {
        "active_profile": active,
        "profiles": profiles,
    }


AUTH_HANDLERS = {
    "m365_auth_status": handle_auth_status,
    "m365_connect": handle_connect,
    "m365_disconnect": handle_disconnect,
    "m365_set_profile": handle_set_profile,
    "m365_list_profiles": handle_list_profiles,
}
