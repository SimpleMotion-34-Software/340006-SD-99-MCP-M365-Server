"""Encrypted token storage for Microsoft 365 OAuth tokens."""

import json
import os
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from cryptography.fernet import Fernet


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
    """Encrypted storage for OAuth tokens."""

    def __init__(self, profile: str = "SM"):
        """Initialize token store for a specific profile.

        Args:
            profile: The credential profile (SM, SG, etc.)
        """
        self.profile = profile
        self.base_dir = Path.home() / ".m365"
        self.token_file = self.base_dir / f"tokens-{profile}.enc"
        self.key_file = self.base_dir / f".key-{profile}"

    def _ensure_dir(self) -> None:
        """Ensure the storage directory exists with proper permissions."""
        self.base_dir.mkdir(mode=0o700, exist_ok=True)

    def _get_or_create_key(self) -> bytes:
        """Get or create the encryption key."""
        self._ensure_dir()

        if self.key_file.exists():
            return self.key_file.read_bytes()

        key = Fernet.generate_key()
        self.key_file.write_bytes(key)
        os.chmod(self.key_file, 0o600)
        return key

    def _get_fernet(self) -> Fernet:
        """Get the Fernet instance for encryption/decryption."""
        key = self._get_or_create_key()
        return Fernet(key)

    def save(self, tokens: Tokens) -> None:
        """Save tokens to encrypted storage.

        Args:
            tokens: The tokens to save.
        """
        self._ensure_dir()
        fernet = self._get_fernet()

        data = json.dumps(tokens.to_dict())
        encrypted = fernet.encrypt(data.encode())

        self.token_file.write_bytes(encrypted)
        os.chmod(self.token_file, 0o600)

    def load(self) -> Optional[Tokens]:
        """Load tokens from encrypted storage.

        Returns:
            The tokens if found and valid, None otherwise.
        """
        if not self.token_file.exists():
            return None

        try:
            fernet = self._get_fernet()
            encrypted = self.token_file.read_bytes()
            data = json.loads(fernet.decrypt(encrypted).decode())
            return Tokens.from_dict(data)
        except Exception:
            return None

    def clear(self) -> None:
        """Clear stored tokens."""
        if self.token_file.exists():
            self.token_file.unlink()

    def exists(self) -> bool:
        """Check if tokens exist in storage."""
        return self.token_file.exists()
