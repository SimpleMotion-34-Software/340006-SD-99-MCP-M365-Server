"""OAuth 2.0 client credentials flow for Microsoft 365 authentication."""

import subprocess
import sys
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

import aiohttp
import jwt
from cryptography.hazmat.primitives import serialization

from .cert_utils import get_private_key_from_keychain, get_thumbprint_from_keychain
from .token_store import TokenStore, Tokens


# Multi-tenant profile configuration
# Each profile has its own Azure AD app registration and credentials
CREDENTIAL_PROFILES = {
    "SM": "SimpleMotion (@simplemotion.com)",
    "SG": "SimpleMotion Global (@simplemotion.global)",
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


KEYCHAIN_ACCOUNT = "m365-mcp"


def _get_keychain_credential(service: str) -> Optional[str]:
    """Get a credential from the macOS keychain.

    Args:
        service: The service name (e.g., 'SM-M365-Client-ID')

    Returns:
        The credential value if found, None otherwise.
    """
    if sys.platform != "darwin":
        return None

    try:
        result = subprocess.run(
            ["security", "find-generic-password", "-a", KEYCHAIN_ACCOUNT, "-s", service, "-w"],
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

        # Load credentials from keychain (format: {Profile}-M365-{Key-Name})
        # Example: SM-M365-Client-ID, SM-M365-Tenant-ID
        self.client_id = _get_keychain_credential(f"{self.profile}-M365-Client-ID")
        self.tenant_id = _get_keychain_credential(f"{self.profile}-M365-Tenant-ID")
        self.user_id = _get_keychain_credential(f"{self.profile}-M365-User-ID")

        # Load certificate credentials
        self.cert_thumbprint = get_thumbprint_from_keychain(self.profile)
        self._private_key_pem: Optional[bytes] = None  # Lazy-loaded

        self.token_store = TokenStore(self.profile)

    def _has_private_key(self) -> bool:
        """Check if a private key is available in keychain."""
        if self._private_key_pem is not None:
            return True
        key = get_private_key_from_keychain(self.profile)
        if key:
            self._private_key_pem = key
            return True
        return False

    @property
    def auth_mode(self) -> str:
        """Return the authentication mode: 'certificate' or 'none'.

        Only certificate-based authentication is supported.
        """
        if self.cert_thumbprint and self._has_private_key():
            return "certificate"
        return "none"

    @property
    def is_configured(self) -> bool:
        """Check if credentials are configured.

        Requires client_id, tenant_id, and a certificate.
        """
        has_base = all([self.client_id, self.tenant_id])
        return has_base and (self.auth_mode == "certificate")

    def get_status(self) -> dict:
        """Get the current authentication status.

        Returns:
            Dictionary with status information.
        """
        tokens = self.token_store.load()
        return {
            "profile": self.profile,
            "configured": self.is_configured,
            "auth_mode": self.auth_mode,
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

    def _create_jwt_assertion(self) -> str:
        """Create a signed JWT assertion for certificate-based authentication.

        Returns:
            Signed JWT string.

        Raises:
            RuntimeError: If private key is not available.
        """
        if not self._private_key_pem:
            self._private_key_pem = get_private_key_from_keychain(self.profile)
        if not self._private_key_pem:
            raise RuntimeError("Private key not found in keychain")

        # Load the private key
        private_key = serialization.load_pem_private_key(
            self._private_key_pem,
            password=None,
        )

        # Token endpoint for this tenant
        token_url = TOKEN_URL.format(tenant_id=self.tenant_id)

        # Current time
        now = datetime.now(timezone.utc)

        # JWT header with x5t#S256 (certificate thumbprint)
        headers = {
            "alg": "RS256",
            "typ": "JWT",
            "x5t#S256": self.cert_thumbprint,
        }

        # JWT payload
        payload = {
            "iss": self.client_id,  # Issuer: the app's client ID
            "sub": self.client_id,  # Subject: same as issuer for client credentials
            "aud": token_url,       # Audience: the token endpoint
            "jti": str(uuid.uuid4()),  # Unique token ID
            "nbf": int(now.timestamp()),  # Not before
            "exp": int((now + timedelta(minutes=10)).timestamp()),  # Expires in 10 min
        }

        # Sign the JWT with the private key
        assertion = jwt.encode(
            payload,
            private_key,
            algorithm="RS256",
            headers=headers,
        )

        return assertion

    async def authenticate(self) -> Tokens:
        """Authenticate using client credentials grant with certificate.

        This flow is used for application-level access without user interaction.
        Requires Application permissions (not Delegated) with admin consent.

        Uses certificate-based authentication with a signed JWT assertion.

        Returns:
            The obtained tokens.

        Raises:
            RuntimeError: If authentication fails.
        """
        if not self.is_configured:
            raise RuntimeError(
                f"M365 credentials not configured for profile '{self.profile}'.\n\n"
                f"Required steps:\n"
                f"  1. Set credentials in keychain:\n"
                f"     security add-generic-password -a {KEYCHAIN_ACCOUNT} -s {self.profile}-M365-Client-ID -w YOUR_APP_ID\n"
                f"     security add-generic-password -a {KEYCHAIN_ACCOUNT} -s {self.profile}-M365-Tenant-ID -w YOUR_TENANT_ID\n"
                f"     security add-generic-password -a {KEYCHAIN_ACCOUNT} -s {self.profile}-M365-User-ID -w user@domain.com\n\n"
                f"  2. Generate certificate:\n"
                f"     Use m365_generate_certificate tool or /sm-mcp-m365 --cert-gen\n\n"
                f"  3. Upload certificate to Azure AD:\n"
                f"     Azure Portal → App registrations → Your App → Certificates & secrets\n\n"
                f"Current status:\n"
                f"  - Client ID: {'Found' if self.client_id else 'NOT FOUND'}\n"
                f"  - Tenant ID: {'Found' if self.tenant_id else 'NOT FOUND'}\n"
                f"  - User ID: {'Found' if self.user_id else 'NOT FOUND'}\n"
                f"  - Certificate: {'Found' if self.cert_thumbprint else 'NOT FOUND'}"
            )

        # Certificate-based authentication using JWT assertion
        assertion = self._create_jwt_assertion()
        data = {
            "client_id": self.client_id,
            "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "client_assertion": assertion,
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
                    raise RuntimeError(f"Authentication failed (certificate): {error}")

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
