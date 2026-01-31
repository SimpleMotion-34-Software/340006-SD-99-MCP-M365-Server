"""OAuth 2.0 client credentials flow for Microsoft 365 authentication."""

import subprocess
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

import aiohttp

from .token_store import TokenStore, Tokens


# Multi-tenant profile configuration
CREDENTIAL_PROFILES = {
    "SM": "-SM",  # SimpleMotion (@simplemotion.com)
    "SG": "-SG",  # SG tenant (@simplemotion.global)
}

DEFAULT_PROFILE = "SM"

# Microsoft OAuth endpoints
TOKEN_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

# Graph API scope for client credentials
CLIENT_CREDENTIALS_SCOPE = "https://graph.microsoft.com/.default"


def get_active_profile() -> str:
    """Get the currently active profile."""
    profile_file = Path.home() / ".m365" / "active_profile"
    if profile_file.exists():
        return profile_file.read_text().strip()
    return DEFAULT_PROFILE


def set_active_profile(profile: str) -> None:
    """Set the active profile."""
    if profile not in CREDENTIAL_PROFILES:
        raise ValueError(f"Invalid profile: {profile}. Must be one of: {list(CREDENTIAL_PROFILES.keys())}")

    profile_dir = Path.home() / ".m365"
    profile_dir.mkdir(mode=0o700, exist_ok=True)

    profile_file = profile_dir / "active_profile"
    profile_file.write_text(profile)


def _get_keychain_credential(name: str) -> Optional[str]:
    """Get a credential from the macOS keychain.

    Args:
        name: The credential name (e.g., 'm365-SM-client-id')

    Returns:
        The credential value if found, None otherwise.
    """
    if sys.platform != "darwin":
        return None

    try:
        result = subprocess.run(
            ["security", "find-generic-password", "-a", "m365-mcp", "-s", name, "-w"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass
    return None


class M365OAuth:
    """Microsoft 365 OAuth 2.0 authentication using client credentials flow."""

    def __init__(self, profile: Optional[str] = None):
        """Initialize OAuth handler.

        Args:
            profile: The credential profile to use. If None, uses the active profile.
        """
        self.profile = profile or get_active_profile()
        self.suffix = CREDENTIAL_PROFILES.get(self.profile, f"-{self.profile}")

        # Load credentials from keychain
        self.client_id = _get_keychain_credential(f"m365{self.suffix}-client-id")
        self.client_secret = _get_keychain_credential(f"m365{self.suffix}-client-secret")
        self.tenant_id = _get_keychain_credential(f"m365{self.suffix}-tenant-id")
        self.user_id = _get_keychain_credential(f"m365{self.suffix}-user-id")

        self.token_store = TokenStore(self.profile)

    @property
    def is_configured(self) -> bool:
        """Check if credentials are configured."""
        return all([self.client_id, self.client_secret, self.tenant_id])

    def get_status(self) -> dict:
        """Get the current authentication status.

        Returns:
            Dictionary with status information.
        """
        tokens = self.token_store.load()
        return {
            "profile": self.profile,
            "configured": self.is_configured,
            "has_tokens": tokens is not None,
            "connected": tokens is not None and not tokens.is_expired(),
            "user_email": tokens.user_email if tokens else self.user_id,
            "user_name": tokens.user_name if tokens else None,
            "tenant_id": self.tenant_id[:8] + "..." if self.tenant_id else None,
        }

    async def get_valid_tokens(self) -> Optional[Tokens]:
        """Get valid tokens, re-authenticating if necessary.

        Returns:
            Valid tokens if available, None otherwise.
        """
        tokens = self.token_store.load()

        if not tokens or tokens.is_expired():
            # Re-authenticate using client credentials
            tokens = await self.authenticate()
            if tokens:
                self.token_store.save(tokens)

        return tokens

    async def authenticate(self) -> Tokens:
        """Authenticate using client credentials grant.

        This flow is used for application-level access without user interaction.
        Requires Application permissions (not Delegated) with admin consent.

        Returns:
            The obtained tokens.

        Raises:
            RuntimeError: If authentication fails.
        """
        if not self.is_configured:
            raise RuntimeError(
                "Credentials not configured. Add to keychain:\n"
                f"  security add-generic-password -a m365-mcp -s m365{self.suffix}-client-id -w YOUR_ID\n"
                f"  security add-generic-password -a m365-mcp -s m365{self.suffix}-client-secret -w YOUR_SECRET\n"
                f"  security add-generic-password -a m365-mcp -s m365{self.suffix}-tenant-id -w YOUR_TENANT"
            )

        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "client_credentials",
            "scope": CLIENT_CREDENTIALS_SCOPE,
        }

        async with aiohttp.ClientSession() as session:
            async with session.post(
                TOKEN_URL.format(tenant_id=self.tenant_id),
                data=data,
            ) as resp:
                if resp.status != 200:
                    error = await resp.text()
                    raise RuntimeError(f"Authentication failed: {error}")

                result = await resp.json()

        # Calculate expiry
        expires_in = result.get("expires_in", 3600)
        expires_at = datetime.now(timezone.utc) + timedelta(seconds=expires_in - 300)

        tokens = Tokens(
            access_token=result["access_token"],
            refresh_token="",  # Client credentials doesn't return refresh token
            token_type=result.get("token_type", "Bearer"),
            expires_at=expires_at.isoformat(),
            scope=result.get("scope", CLIENT_CREDENTIALS_SCOPE),
            user_email=self.user_id,  # Use configured user_id
            user_name=None,
        )

        return tokens

    def disconnect(self) -> None:
        """Clear stored tokens."""
        self.token_store.clear()
