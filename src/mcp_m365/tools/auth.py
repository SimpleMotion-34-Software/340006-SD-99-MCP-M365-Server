"""Authentication tools for M365 MCP Server."""

import subprocess
import sys
from typing import Any, Dict, List

from mcp.types import Tool

from ..auth import M365OAuth, get_active_profile, set_active_profile, CREDENTIAL_PROFILES
from ..auth.cert_utils import (
    generate_self_signed_certificate,
    import_to_keychain,
    save_certificate_file,
    delete_certificate_from_keychain,
    certificate_exists_in_keychain,
)
from ..graph import GraphClient


def _keychain_set(service: str, account: str, password: str) -> bool:
    """Set a credential in macOS Keychain."""
    if sys.platform != "darwin":
        return False

    # Delete existing entry first
    subprocess.run(
        ["security", "delete-generic-password", "-s", service, "-a", account],
        capture_output=True,
    )

    # Add new entry
    result = subprocess.run(
        ["security", "add-generic-password", "-s", service, "-a", account, "-w", password, "-U"],
        capture_output=True,
        text=True,
    )
    return result.returncode == 0


def _keychain_delete(service: str, account: str) -> bool:
    """Delete a credential from macOS Keychain."""
    if sys.platform != "darwin":
        return False

    result = subprocess.run(
        ["security", "delete-generic-password", "-s", service, "-a", account],
        capture_output=True,
    )
    return result.returncode == 0


def _keychain_exists(service: str, account: str) -> bool:
    """Check if a credential exists in macOS Keychain."""
    if sys.platform != "darwin":
        return False

    result = subprocess.run(
        ["security", "find-generic-password", "-s", service, "-a", account],
        capture_output=True,
    )
    return result.returncode == 0


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
        description="Authenticate with Microsoft 365 using client credentials flow. No browser required - uses Application permissions with admin consent.",
        inputSchema={
            "type": "object",
            "properties": {},
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
    Tool(
        name="m365_set_credential",
        description="Set a Microsoft 365 credential in macOS Keychain. Use this to configure client_id, tenant_id, or user_id for a profile.",
        inputSchema={
            "type": "object",
            "properties": {
                "credential": {
                    "type": "string",
                    "description": "Credential type: 'client_id', 'tenant_id', or 'user_id'",
                    "enum": ["client_id", "tenant_id", "user_id"],
                },
                "value": {
                    "type": "string",
                    "description": "The credential value to store",
                },
                "profile": {
                    "type": "string",
                    "description": "Profile to set credential for (e.g., 'SM', 'SG'). Defaults to active profile.",
                },
            },
            "required": ["credential", "value"],
        },
    ),
    Tool(
        name="m365_delete_credential",
        description="Delete a Microsoft 365 credential from macOS Keychain.",
        inputSchema={
            "type": "object",
            "properties": {
                "credential": {
                    "type": "string",
                    "description": "Credential type: 'client_id', 'tenant_id', or 'user_id'",
                    "enum": ["client_id", "tenant_id", "user_id"],
                },
                "profile": {
                    "type": "string",
                    "description": "Profile to delete credential for (e.g., 'SM', 'SG'). Defaults to active profile.",
                },
            },
            "required": ["credential"],
        },
    ),
    Tool(
        name="m365_list_credentials",
        description="List which Microsoft 365 credentials are configured in Keychain (does not show values).",
        inputSchema={
            "type": "object",
            "properties": {
                "profile": {
                    "type": "string",
                    "description": "Profile to check (e.g., 'SM', 'SG'). If not specified, checks all profiles.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_delete_tokens",
        description="Delete stored OAuth tokens from Keychain for a profile. Use this to force re-authentication.",
        inputSchema={
            "type": "object",
            "properties": {
                "profile": {
                    "type": "string",
                    "description": "Profile to delete tokens for (e.g., 'SM', 'SG'). Defaults to active profile.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_generate_certificate",
        description="Generate a self-signed certificate for Azure AD certificate-based authentication. More secure than client secrets.",
        inputSchema={
            "type": "object",
            "properties": {
                "profile": {
                    "type": "string",
                    "description": "Profile to generate certificate for (e.g., 'SM', 'SG'). Defaults to active profile.",
                },
                "validity_days": {
                    "type": "integer",
                    "description": "Certificate validity in days (default: 730 = 2 years).",
                    "default": 730,
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="m365_delete_certificate",
        description="Delete certificate credentials from Keychain for a profile.",
        inputSchema={
            "type": "object",
            "properties": {
                "profile": {
                    "type": "string",
                    "description": "Profile to delete certificate for (e.g., 'SM', 'SG'). Defaults to active profile.",
                },
            },
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
    """Handle m365_connect tool call using client credentials flow."""
    if not oauth.is_configured:
        return {
            "error": "Credentials not configured",
            "instructions": [
                "Certificate authentication required:",
                "1. Generate certificate: m365_generate_certificate",
                "2. Upload certificate to Azure AD (App registrations > Certificates & secrets)",
                f'3. security add-generic-password -a "m365-mcp" -s "m365{oauth.suffix}-client-id" -w "YOUR_CLIENT_ID"',
                f'4. security add-generic-password -a "m365-mcp" -s "m365{oauth.suffix}-tenant-id" -w "YOUR_TENANT_ID"',
                f'5. security add-generic-password -a "m365-mcp" -s "m365{oauth.suffix}-user-id" -w "user@domain.com"',
            ],
        }

    try:
        # Authenticate with client credentials
        tokens = await oauth.authenticate()

        return {
            "status": "connected",
            "profile": oauth.profile,
            "user_email": tokens.user_email,
            "message": "Successfully authenticated with Microsoft 365 (client credentials)",
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


async def handle_set_credential(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_set_credential tool call."""
    credential = arguments.get("credential")
    value = arguments.get("value")
    profile = arguments.get("profile") or get_active_profile()

    if not credential or not value:
        return {"error": "credential and value are required"}

    valid_creds = ["client_id", "tenant_id", "user_id"]
    if credential not in valid_creds:
        return {"error": f"credential must be one of: {', '.join(valid_creds)}"}

    suffix = CREDENTIAL_PROFILES.get(profile, f"-{profile}")
    service_name = f"m365{suffix}-{credential.replace('_', '-')}"

    success = _keychain_set(service_name, "m365-mcp", value)
    if success:
        return {
            "success": True,
            "profile": profile,
            "credential": credential,
            "message": f"Credential '{credential}' set for profile {profile}",
        }
    else:
        return {"error": "Failed to set credential in Keychain"}


async def handle_delete_credential(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_delete_credential tool call."""
    credential = arguments.get("credential")
    profile = arguments.get("profile") or get_active_profile()

    if not credential:
        return {"error": "credential is required"}

    suffix = CREDENTIAL_PROFILES.get(profile, f"-{profile}")
    service_name = f"m365{suffix}-{credential.replace('_', '-')}"

    success = _keychain_delete(service_name, "m365-mcp")
    if success:
        return {
            "success": True,
            "profile": profile,
            "credential": credential,
            "message": f"Credential '{credential}' deleted for profile {profile}",
        }
    else:
        return {"error": "Failed to delete credential (may not exist)"}


async def handle_list_credentials(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_list_credentials tool call."""
    profile_arg = arguments.get("profile")
    profiles_to_check = [profile_arg] if profile_arg else list(CREDENTIAL_PROFILES.keys())

    results = {}
    for prof in profiles_to_check:
        suffix = CREDENTIAL_PROFILES.get(prof, f"-{prof}")
        has_cert = certificate_exists_in_keychain(prof)

        results[prof] = {
            "client_id": _keychain_exists(f"m365{suffix}-client-id", "m365-mcp"),
            "tenant_id": _keychain_exists(f"m365{suffix}-tenant-id", "m365-mcp"),
            "user_id": _keychain_exists(f"m365{suffix}-user-id", "m365-mcp"),
            "certificate": has_cert,
            "auth_mode": "certificate" if has_cert else "none",
            "tokens": _keychain_exists("m365-mcp-tokens", prof),
        }

    return {
        "profiles": results,
        "message": "Credential status (does not show values). Certificate authentication required.",
    }


async def handle_delete_tokens(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_delete_tokens tool call."""
    profile = arguments.get("profile") or get_active_profile()
    success = _keychain_delete("m365-mcp-tokens", profile)

    if success:
        return {
            "success": True,
            "profile": profile,
            "message": f"Tokens deleted for profile {profile}. Re-authentication required.",
        }
    else:
        return {"error": "Failed to delete tokens (may not exist)"}


async def handle_generate_certificate(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_generate_certificate tool call."""
    profile = arguments.get("profile") or get_active_profile()
    validity_days = arguments.get("validity_days", 730)

    # Check if certificate already exists
    if certificate_exists_in_keychain(profile):
        return {
            "error": "Certificate already exists for this profile",
            "profile": profile,
            "hint": "Use m365_delete_certificate first to remove the existing certificate",
        }

    try:
        # Generate certificate
        common_name = f"M365 MCP Server - {profile}"
        private_key_pem, cert_pem, thumbprint = generate_self_signed_certificate(
            common_name=common_name,
            validity_days=validity_days,
        )

        # Store in keychain
        if not import_to_keychain(profile, private_key_pem, cert_pem, thumbprint):
            return {
                "error": "Failed to store certificate in keychain",
                "profile": profile,
            }

        # Save .cer file for Azure upload
        cert_file = save_certificate_file(profile, cert_pem)

        return {
            "success": True,
            "profile": profile,
            "common_name": common_name,
            "validity_days": validity_days,
            "thumbprint": thumbprint,
            "cert_file": str(cert_file),
            "message": f"Certificate generated and stored in keychain",
            "next_steps": [
                f"1. Upload the certificate to Azure AD:",
                f"   - Go to: https://portal.azure.com",
                f"   - Navigate to: App registrations > Your App > Certificates & secrets",
                f"   - Click 'Upload certificate' and select: {cert_file}",
                f"2. Optionally remove client secret (certificate is more secure)",
                f"3. Run /sm-mcp-m365 --auth --tenant {profile} to authenticate",
            ],
        }
    except Exception as e:
        return {
            "error": f"Failed to generate certificate: {str(e)}",
            "profile": profile,
        }


async def handle_delete_certificate(
    arguments: Dict[str, Any],
    oauth: M365OAuth,
    client: GraphClient,
) -> Dict[str, Any]:
    """Handle m365_delete_certificate tool call."""
    profile = arguments.get("profile") or get_active_profile()

    key_deleted, cert_deleted, thumb_deleted = delete_certificate_from_keychain(profile)

    if not any([key_deleted, cert_deleted, thumb_deleted]):
        return {
            "error": "No certificate credentials found to delete",
            "profile": profile,
        }

    return {
        "success": True,
        "profile": profile,
        "deleted": {
            "private_key": key_deleted,
            "certificate": cert_deleted,
            "thumbprint": thumb_deleted,
        },
        "message": f"Certificate credentials deleted for profile {profile}",
        "note": "Remember to also remove the certificate from Azure AD if no longer needed",
    }


AUTH_HANDLERS = {
    "m365_auth_status": handle_auth_status,
    "m365_connect": handle_connect,
    "m365_disconnect": handle_disconnect,
    "m365_set_profile": handle_set_profile,
    "m365_list_profiles": handle_list_profiles,
    "m365_set_credential": handle_set_credential,
    "m365_delete_credential": handle_delete_credential,
    "m365_list_credentials": handle_list_credentials,
    "m365_delete_tokens": handle_delete_tokens,
    "m365_generate_certificate": handle_generate_certificate,
    "m365_delete_certificate": handle_delete_certificate,
}
