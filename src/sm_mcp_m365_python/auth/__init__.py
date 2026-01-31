"""Authentication module for Microsoft 365 MCP Server."""

from .oauth import M365OAuth, get_active_profile, set_active_profile, CREDENTIAL_PROFILES
from .token_store import TokenStore, Tokens

__all__ = [
    "M365OAuth",
    "TokenStore",
    "Tokens",
    "get_active_profile",
    "set_active_profile",
    "CREDENTIAL_PROFILES",
]
