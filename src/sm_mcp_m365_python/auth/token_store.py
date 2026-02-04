"""Secure token storage for Microsoft 365 OAuth tokens using macOS Keychain."""

import json
import subprocess
import sys
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from typing import Optional


@dataclass
class Tokens:
    """OAuth tokens and metadata."""

    access_token: str
    refresh_token: str
    token_type: str
    expires_at: str  # ISO format datetime
    scope: str
    user_email: Optional[str] = None
    user_name: Optional[str] = None

    def is_expired(self) -> bool:
        """Check if the access token is expired."""
        try:
            expires = datetime.fromisoformat(self.expires_at.replace("Z", "+00:00"))
            # Consider expired 5 minutes before actual expiry
            return datetime.now(timezone.utc) >= expires
        except (ValueError, AttributeError):
            return True

    def to_dict(self) -> dict:
        """Convert to dictionary."""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "Tokens":
        """Create from dictionary."""
        return cls(
            access_token=data.get("access_token", ""),
            refresh_token=data.get("refresh_token", ""),
            token_type=data.get("token_type", "Bearer"),
            expires_at=data.get("expires_at", ""),
            scope=data.get("scope", ""),
            user_email=data.get("user_email"),
            user_name=data.get("user_name"),
        )


class TokenStore:
    """Secure storage for OAuth tokens using macOS Keychain."""

    def __init__(self, profile: str = "SM"):
        """Initialize token store for a specific profile.

        Args:
            profile: The credential profile (SM, SG, etc.)
        """
        self.profile = profile.upper()
        # Keychain service name: {Profile}-M365 (e.g., SM-M365, SG-M365)
        self.keychain_service = f"{self.profile}-M365"

    def _keychain_save(self, data: str) -> bool:
        """Save data to macOS Keychain.

        Args:
            data: JSON string to save

        Returns:
            True if successful
        """
        if sys.platform != "darwin":
            raise RuntimeError("Keychain storage only supported on macOS")

        # Delete existing entry first (ignore errors if doesn't exist)
        subprocess.run(
            [
                "security", "delete-generic-password",
                "-s", self.keychain_service,
                "-a", self.profile,
            ],
            capture_output=True,
        )

        # Add new entry
        result = subprocess.run(
            [
                "security", "add-generic-password",
                "-s", self.keychain_service,
                "-a", self.profile,
                "-w", data,
                "-U",  # Update if exists
            ],
            capture_output=True,
            text=True,
        )
        return result.returncode == 0

    def _keychain_load(self) -> Optional[str]:
        """Load data from macOS Keychain.

        Returns:
            JSON string if found, None otherwise
        """
        if sys.platform != "darwin":
            return None

        try:
            result = subprocess.run(
                [
                    "security", "find-generic-password",
                    "-s", self.keychain_service,
                    "-a", self.profile,
                    "-w",
                ],
                capture_output=True,
                text=True,
                timeout=5,
            )
            if result.returncode == 0:
                return result.stdout.strip()
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass
        return None

    def _keychain_delete(self) -> bool:
        """Delete entry from macOS Keychain.

        Returns:
            True if successful
        """
        if sys.platform != "darwin":
            return False

        result = subprocess.run(
            [
                "security", "delete-generic-password",
                "-s", self.keychain_service,
                "-a", self.profile,
            ],
            capture_output=True,
        )
        return result.returncode == 0

    def save(self, tokens: Tokens) -> None:
        """Save tokens to Keychain.

        Args:
            tokens: The tokens to save.
        """
        data = json.dumps(tokens.to_dict())
        if not self._keychain_save(data):
            raise RuntimeError("Failed to save tokens to Keychain")

    def load(self) -> Optional[Tokens]:
        """Load tokens from Keychain.

        Returns:
            The tokens if found and valid, None otherwise.
        """
        data = self._keychain_load()
        if not data:
            return None

        try:
            return Tokens.from_dict(json.loads(data))
        except (json.JSONDecodeError, KeyError):
            return None

    def clear(self) -> None:
        """Clear stored tokens."""
        self._keychain_delete()

    def exists(self) -> bool:
        """Check if tokens exist in storage."""
        return self._keychain_load() is not None
