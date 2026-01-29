"""OAuth client credentials flow handler for Microsoft Azure AD authentication."""

import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import aiohttp

from .token_store import TokenSet, TokenStore


# Microsoft OAuth endpoints
MICROSOFT_TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"

# Scope for client credentials (application permissions)
# Uses .default to request all configured application permissions
CLIENT_CREDENTIALS_SCOPE = "https://graph.microsoft.com/.default"


def _get_keychain_password_macos(service: str) -> str | None:
    """Retrieve password from macOS Keychain.

    Args:
        service: Keychain service name

    Returns:
        Password if found, None otherwise
    """
    try:
        result = subprocess.run(
            ["security", "find-generic-password", "-s", service, "-w"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass

    return None


def _get_secret_tool_password_linux(name: str) -> str | None:
    """Retrieve password from Linux secret storage using secret-tool (libsecret).

    Args:
        name: Secret name (e.g., 'm365-client-id')

    Returns:
        Password if found, None otherwise
    """
    try:
        result = subprocess.run(
            ["secret-tool", "lookup", "service", "m365-mcp", "name", name],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass

    return None


def _get_secure_credential(name: str) -> str | None:
    """Retrieve credential from platform-specific secure storage.

    Args:
        name: Credential name (e.g., 'm365-client-id')

    Returns:
        Credential value if found, None otherwise

    Platform support:
        - macOS: Keychain (security command)
        - Linux: libsecret via secret-tool (GNOME Keyring, KDE Wallet)
    """
    if sys.platform == "darwin":
        return _get_keychain_password_macos(name)
    elif sys.platform.startswith("linux"):
        return _get_secret_tool_password_linux(name)
    return None


class M365OAuth:
    """Handle Microsoft Azure AD OAuth 2.0 client credentials authentication."""

    def __init__(
        self,
        client_id: str | None = None,
        client_secret: str | None = None,
        tenant_id: str | None = None,
        token_store: TokenStore | None = None,
    ):
        """Initialize OAuth handler.

        Args:
            client_id: Azure AD app client ID (defaults to keychain or M365_CLIENT_ID env var)
            client_secret: Azure AD app client secret (defaults to keychain or M365_CLIENT_SECRET env var)
            tenant_id: Azure AD tenant ID (defaults to keychain or M365_TENANT_ID env var)
            token_store: Token storage handler

        Credential lookup order:
            1. Explicit parameter
            2. Platform secure storage:
               - macOS: Keychain (m365-client-id, m365-client-secret, m365-tenant-id)
               - Linux: libsecret via secret-tool
            3. Environment variable (M365_CLIENT_ID, M365_CLIENT_SECRET, M365_TENANT_ID)
        """
        self.client_id = (
            client_id
            or _get_secure_credential("m365-client-id")
            or os.environ.get("M365_CLIENT_ID", "")
        )
        self.client_secret = (
            client_secret
            or _get_secure_credential("m365-client-secret")
            or os.environ.get("M365_CLIENT_SECRET", "")
        )
        self.tenant_id = (
            tenant_id
            or _get_secure_credential("m365-tenant-id")
            or os.environ.get("M365_TENANT_ID", "")
        )

        # Token storage
        token_path = Path.home() / ".m365" / "tokens.enc"
        self.token_store = token_store or TokenStore(storage_path=token_path)

    @property
    def is_configured(self) -> bool:
        """Check if OAuth credentials are configured."""
        return bool(self.client_id and self.client_secret)

    async def authenticate_client_credentials(self, user_id: str | None = None) -> TokenSet:
        """Authenticate using client credentials grant (app-only, no browser).

        This is used for Azure AD apps with application permissions that have
        been granted admin consent. No user interaction required.

        Args:
            user_id: User ID or email to impersonate for mail access.
                     If not provided, checks keychain for 'm365-user-id'.

        Returns:
            Token set with access token

        Raises:
            ValueError: If credentials not configured or auth fails

        Note:
            Requires:
            - Application permissions (not delegated) in Azure AD app
            - Admin consent granted for the tenant
            - Specific tenant_id (not "common")
        """
        if not self.is_configured:
            raise ValueError("M365_CLIENT_ID and M365_CLIENT_SECRET must be set")

        if self.tenant_id == "common":
            raise ValueError(
                "Client credentials flow requires a specific tenant_id. "
                "Set m365-tenant-id in keychain or M365_TENANT_ID env var."
            )

        # Get user_id from parameter, keychain, or env
        target_user = (
            user_id
            or _get_secure_credential("m365-user-id")
            or os.environ.get("M365_USER_ID")
        )

        token_url = MICROSOFT_TOKEN_URL.format(tenant=self.tenant_id)

        async with aiohttp.ClientSession() as session:
            async with session.post(
                token_url,
                data={
                    "client_id": self.client_id,
                    "client_secret": self.client_secret,
                    "scope": CLIENT_CREDENTIALS_SCOPE,
                    "grant_type": "client_credentials",
                },
                headers={"Content-Type": "application/x-www-form-urlencoded"},
            ) as response:
                if response.status != 200:
                    error = await response.text()
                    raise ValueError(f"Client credentials auth failed: {error}")

                data = await response.json()

        # For client credentials, we store the target user for API calls
        tokens = TokenSet(
            access_token=data["access_token"],
            refresh_token="",  # No refresh token with client credentials
            expires_at=datetime.now().timestamp() + data["expires_in"],
            token_type=data["token_type"],
            scope=data.get("scope", CLIENT_CREDENTIALS_SCOPE).split(),
            user_email=target_user,
            user_name=target_user,  # Use email as display name
        )

        self.token_store.save(tokens)
        return tokens

    async def get_valid_tokens(self) -> TokenSet | None:
        """Get valid tokens, re-authenticating if expired.

        Returns:
            Valid token set or None if not authenticated
        """
        tokens = self.token_store.load()
        if not tokens:
            return None

        if tokens.is_expired:
            try:
                # Re-authenticate using client credentials
                tokens = await self.authenticate_client_credentials(tokens.user_email)
            except ValueError:
                return None

        return tokens

    def disconnect(self) -> None:
        """Disconnect from Microsoft by removing stored tokens."""
        self.token_store.delete()

    def get_status(self) -> dict[str, Any]:
        """Get current authentication status.

        Returns:
            Status dictionary with connection info
        """
        if not self.is_configured:
            return {
                "connected": False,
                "configured": False,
                "message": "Microsoft credentials not configured. Set M365_CLIENT_ID and M365_CLIENT_SECRET.",
            }

        tokens = self.token_store.load()
        if not tokens:
            return {
                "connected": False,
                "configured": True,
                "message": "Not connected to Microsoft 365. Use m365_connect to begin authentication.",
            }

        status = {
            "connected": True,
            "configured": True,
            "expired": tokens.is_expired,
            "user_email": tokens.user_email,
            "user_name": tokens.user_name,
            "scopes": tokens.scope,
            "message": f"Connected as {tokens.user_name or tokens.user_email or 'Unknown'}"
            + (" (token expired, will refresh)" if tokens.is_expired else ""),
        }

        return status
