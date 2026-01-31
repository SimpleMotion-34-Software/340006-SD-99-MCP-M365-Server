"""OAuth 2.0 device code flow for Microsoft 365 authentication."""

import asyncio
import subprocess
from dataclasses import dataclass
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
DEVICE_CODE_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode"

# Required scopes
SCOPES = [
    "offline_access",
    "User.Read",
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Contacts.Read",
    "Contacts.ReadWrite",
]


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
    try:
        result = subprocess.run(
            ["security", "find-generic-password", "-a", "m365-mcp", "-s", name, "-w"],
            capture_output=True,
            text=True,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return None


@dataclass
class DeviceCodeResponse:
    """Device code flow response from Microsoft."""

    device_code: str
    user_code: str
    verification_uri: str
    expires_in: int
    interval: int
    message: str


class M365OAuth:
    """Microsoft 365 OAuth 2.0 authentication with device code flow."""

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
            "user_email": tokens.user_email if tokens else None,
            "user_name": tokens.user_name if tokens else None,
            "tenant_id": self.tenant_id[:8] + "..." if self.tenant_id else None,
        }

    async def get_valid_tokens(self) -> Optional[Tokens]:
        """Get valid tokens, refreshing if necessary.

        Returns:
            Valid tokens if available, None otherwise.
        """
        tokens = self.token_store.load()
        if not tokens:
            return None

        if tokens.is_expired():
            # Try to refresh
            tokens = await self._refresh_tokens(tokens.refresh_token)
            if tokens:
                self.token_store.save(tokens)

        return tokens

    async def authenticate_device_code(self) -> dict:
        """Start device code authentication flow.

        Returns:
            Dictionary with device code info for user display.

        Raises:
            RuntimeError: If authentication fails.
        """
        if not self.is_configured:
            raise RuntimeError("Credentials not configured. Add credentials to keychain.")

        # Request device code
        data = {
            "client_id": self.client_id,
            "scope": " ".join(SCOPES),
        }

        async with aiohttp.ClientSession() as session:
            async with session.post(
                DEVICE_CODE_URL.format(tenant_id=self.tenant_id),
                data=data,
            ) as resp:
                if resp.status != 200:
                    error = await resp.text()
                    raise RuntimeError(f"Device code request failed: {error}")

                result = await resp.json()

        device_code = DeviceCodeResponse(
            device_code=result["device_code"],
            user_code=result["user_code"],
            verification_uri=result["verification_uri"],
            expires_in=result["expires_in"],
            interval=result.get("interval", 5),
            message=result["message"],
        )

        return {
            "user_code": device_code.user_code,
            "verification_uri": device_code.verification_uri,
            "message": device_code.message,
            "expires_in": device_code.expires_in,
            "_device_code": device_code.device_code,
            "_interval": device_code.interval,
        }

    async def poll_device_code(self, device_code: str, interval: int = 5, timeout: int = 300) -> Tokens:
        """Poll for device code authentication completion.

        Args:
            device_code: The device code from authenticate_device_code
            interval: Polling interval in seconds
            timeout: Maximum time to wait in seconds

        Returns:
            The obtained tokens.

        Raises:
            RuntimeError: If authentication fails or times out.
        """
        if not self.is_configured:
            raise RuntimeError("Credentials not configured.")

        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": device_code,
        }

        elapsed = 0
        while elapsed < timeout:
            await asyncio.sleep(interval)
            elapsed += interval

            async with aiohttp.ClientSession() as session:
                async with session.post(
                    TOKEN_URL.format(tenant_id=self.tenant_id),
                    data=data,
                ) as resp:
                    result = await resp.json()

                    if resp.status == 200:
                        # Success - got tokens
                        expires_in = result.get("expires_in", 3600)
                        expires_at = datetime.now(timezone.utc) + timedelta(seconds=expires_in - 300)

                        tokens = Tokens(
                            access_token=result["access_token"],
                            refresh_token=result.get("refresh_token", ""),
                            token_type=result.get("token_type", "Bearer"),
                            expires_at=expires_at.isoformat(),
                            scope=result.get("scope", ""),
                        )

                        tokens = await self._fetch_user_info(tokens)
                        self.token_store.save(tokens)
                        return tokens

                    error = result.get("error", "")

                    if error == "authorization_pending":
                        # User hasn't authenticated yet, keep polling
                        continue
                    elif error == "slow_down":
                        # Increase interval
                        interval += 5
                        continue
                    elif error == "expired_token":
                        raise RuntimeError("Device code expired. Please try again.")
                    elif error == "authorization_declined":
                        raise RuntimeError("User declined authorization.")
                    else:
                        error_desc = result.get("error_description", error)
                        raise RuntimeError(f"Authentication failed: {error_desc}")

        raise RuntimeError("Device code authentication timed out.")

    async def _refresh_tokens(self, refresh_token: str) -> Optional[Tokens]:
        """Refresh the access token.

        Args:
            refresh_token: The refresh token.

        Returns:
            New tokens if successful, None otherwise.
        """
        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": " ".join(SCOPES),
        }

        try:
            async with aiohttp.ClientSession() as session:
                async with session.post(
                    TOKEN_URL.format(tenant_id=self.tenant_id),
                    data=data,
                ) as resp:
                    if resp.status != 200:
                        return None

                    result = await resp.json()

            expires_in = result.get("expires_in", 3600)
            expires_at = datetime.now(timezone.utc) + timedelta(seconds=expires_in - 300)

            tokens = Tokens(
                access_token=result["access_token"],
                refresh_token=result.get("refresh_token", refresh_token),
                token_type=result.get("token_type", "Bearer"),
                expires_at=expires_at.isoformat(),
                scope=result.get("scope", ""),
            )

            tokens = await self._fetch_user_info(tokens)
            return tokens

        except Exception:
            return None

    async def _fetch_user_info(self, tokens: Tokens) -> Tokens:
        """Fetch user info and add to tokens.

        Args:
            tokens: The tokens to update.

        Returns:
            Updated tokens with user info.
        """
        try:
            async with aiohttp.ClientSession() as session:
                async with session.get(
                    "https://graph.microsoft.com/v1.0/me",
                    headers={"Authorization": f"Bearer {tokens.access_token}"},
                ) as resp:
                    if resp.status == 200:
                        user = await resp.json()
                        tokens.user_email = user.get("mail") or user.get("userPrincipalName")
                        tokens.user_name = user.get("displayName")
        except Exception:
            pass

        return tokens

    def disconnect(self) -> None:
        """Clear stored tokens."""
        self.token_store.clear()
