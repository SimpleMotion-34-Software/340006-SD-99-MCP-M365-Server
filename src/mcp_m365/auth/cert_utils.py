"""Certificate utilities for M365 OAuth certificate-based authentication."""

import base64
import hashlib
import subprocess
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional, Tuple

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID


def generate_self_signed_certificate(
    common_name: str,
    validity_days: int = 730,
) -> Tuple[bytes, bytes, str]:
    """Generate a self-signed X.509 certificate for Azure AD authentication.

    Args:
        common_name: The CN for the certificate (e.g., "M365 MCP Server - SM")
        validity_days: Certificate validity period in days (default: 730 = 2 years)

    Returns:
        Tuple of (private_key_pem, cert_pem, thumbprint)
        - private_key_pem: PEM-encoded RSA private key
        - cert_pem: PEM-encoded X.509 certificate
        - thumbprint: Base64url-encoded SHA-256 thumbprint of the DER-encoded certificate
    """
    # Generate RSA 2048-bit key pair
    private_key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
    )

    # Build certificate
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, common_name),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "SimpleMotion"),
    ])

    now = datetime.now(timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(private_key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now)
        .not_valid_after(now + timedelta(days=validity_days))
        .add_extension(
            x509.BasicConstraints(ca=False, path_length=None),
            critical=True,
        )
        .add_extension(
            x509.KeyUsage(
                digital_signature=True,
                key_encipherment=False,
                content_commitment=False,
                data_encipherment=False,
                key_agreement=False,
                key_cert_sign=False,
                crl_sign=False,
                encipher_only=False,
                decipher_only=False,
            ),
            critical=True,
        )
        .sign(private_key, hashes.SHA256())
    )

    # Serialize to PEM
    private_key_pem = private_key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    )

    cert_pem = cert.public_bytes(serialization.Encoding.PEM)

    # Calculate thumbprint (base64url of SHA-256 of DER-encoded cert)
    cert_der = cert.public_bytes(serialization.Encoding.DER)
    thumbprint_bytes = hashlib.sha256(cert_der).digest()
    thumbprint = base64.urlsafe_b64encode(thumbprint_bytes).rstrip(b"=").decode("ascii")

    return private_key_pem, cert_pem, thumbprint


def _keychain_set(service: str, account: str, value: str) -> bool:
    """Set a credential in macOS Keychain."""
    if sys.platform != "darwin":
        return False

    # Delete existing entry first (ignore errors)
    subprocess.run(
        ["security", "delete-generic-password", "-s", service, "-a", account],
        capture_output=True,
    )

    # Add new entry
    result = subprocess.run(
        ["security", "add-generic-password", "-s", service, "-a", account, "-w", value, "-U"],
        capture_output=True,
        text=True,
    )
    return result.returncode == 0


def _keychain_get(service: str, account: str) -> Optional[str]:
    """Get a credential from macOS Keychain."""
    if sys.platform != "darwin":
        return None

    try:
        result = subprocess.run(
            ["security", "find-generic-password", "-s", service, "-a", account, "-w"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass
    return None


def _keychain_delete(service: str, account: str) -> bool:
    """Delete a credential from macOS Keychain."""
    if sys.platform != "darwin":
        return False

    result = subprocess.run(
        ["security", "delete-generic-password", "-s", service, "-a", account],
        capture_output=True,
    )
    return result.returncode == 0


def import_to_keychain(
    profile: str,
    private_key_pem: bytes,
    cert_pem: bytes,
    thumbprint: str,
) -> bool:
    """Store certificate, private key, and thumbprint in macOS Keychain.

    Args:
        profile: The credential profile (e.g., "SM", "SG")
        private_key_pem: PEM-encoded private key
        cert_pem: PEM-encoded certificate
        thumbprint: Base64url-encoded SHA-256 thumbprint

    Returns:
        True if all items were stored successfully.

    Note:
        PEM data is base64-encoded before storage because the macOS keychain
        hex-encodes values containing newlines, which corrupts the PEM format.
    """
    account = "m365-mcp"
    suffix = f"-{profile}"

    # Store private key (base64-encoded to avoid keychain hex encoding issues)
    key_service = f"m365{suffix}-cert-key"
    key_b64 = base64.b64encode(private_key_pem).decode("ascii")
    if not _keychain_set(key_service, account, key_b64):
        return False

    # Store certificate (base64-encoded)
    cert_service = f"m365{suffix}-cert"
    cert_b64 = base64.b64encode(cert_pem).decode("ascii")
    if not _keychain_set(cert_service, account, cert_b64):
        return False

    # Store thumbprint (already base64url, no newlines)
    thumb_service = f"m365{suffix}-cert-thumbprint"
    if not _keychain_set(thumb_service, account, thumbprint):
        return False

    return True


def get_private_key_from_keychain(profile: str) -> Optional[bytes]:
    """Retrieve the PEM-encoded private key from macOS Keychain.

    Args:
        profile: The credential profile (e.g., "SM", "SG")

    Returns:
        PEM-encoded private key bytes, or None if not found.
    """
    suffix = f"-{profile}"
    key_service = f"m365{suffix}-cert-key"
    key_b64 = _keychain_get(key_service, "m365-mcp")
    if key_b64:
        try:
            return base64.b64decode(key_b64)
        except Exception:
            # Fall back to raw encoding for backwards compatibility
            return key_b64.encode("utf-8")
    return None


def get_certificate_from_keychain(profile: str) -> Optional[bytes]:
    """Retrieve the PEM-encoded certificate from macOS Keychain.

    Args:
        profile: The credential profile (e.g., "SM", "SG")

    Returns:
        PEM-encoded certificate bytes, or None if not found.
    """
    suffix = f"-{profile}"
    cert_service = f"m365{suffix}-cert"
    cert_b64 = _keychain_get(cert_service, "m365-mcp")
    if cert_b64:
        try:
            return base64.b64decode(cert_b64)
        except Exception:
            # Fall back to raw encoding for backwards compatibility
            return cert_b64.encode("utf-8")
    return None


def get_thumbprint_from_keychain(profile: str) -> Optional[str]:
    """Retrieve the certificate thumbprint from macOS Keychain.

    Args:
        profile: The credential profile (e.g., "SM", "SG")

    Returns:
        Base64url-encoded SHA-256 thumbprint, or None if not found.
    """
    suffix = f"-{profile}"
    thumb_service = f"m365{suffix}-cert-thumbprint"
    return _keychain_get(thumb_service, "m365-mcp")


def delete_certificate_from_keychain(profile: str) -> Tuple[bool, bool, bool]:
    """Delete certificate, private key, and thumbprint from macOS Keychain.

    Args:
        profile: The credential profile (e.g., "SM", "SG")

    Returns:
        Tuple of (key_deleted, cert_deleted, thumbprint_deleted) booleans.
    """
    account = "m365-mcp"
    suffix = f"-{profile}"

    key_deleted = _keychain_delete(f"m365{suffix}-cert-key", account)
    cert_deleted = _keychain_delete(f"m365{suffix}-cert", account)
    thumb_deleted = _keychain_delete(f"m365{suffix}-cert-thumbprint", account)

    return key_deleted, cert_deleted, thumb_deleted


def save_certificate_file(profile: str, cert_pem: bytes) -> Path:
    """Save the certificate to a .cer file for upload to Azure AD.

    Args:
        profile: The credential profile (e.g., "SM", "SG")
        cert_pem: PEM-encoded certificate

    Returns:
        Path to the saved .cer file.
    """
    # Create certs directory
    certs_dir = Path.home() / ".m365" / "certs"
    certs_dir.mkdir(mode=0o700, parents=True, exist_ok=True)

    # Save certificate file
    cert_file = certs_dir / f"m365-{profile}-cert.cer"
    cert_file.write_bytes(cert_pem)
    cert_file.chmod(0o600)

    return cert_file


def certificate_exists_in_keychain(profile: str) -> bool:
    """Check if certificate credentials exist in keychain.

    Args:
        profile: The credential profile (e.g., "SM", "SG")

    Returns:
        True if thumbprint and private key are both present.
    """
    thumbprint = get_thumbprint_from_keychain(profile)
    private_key = get_private_key_from_keychain(profile)
    return thumbprint is not None and private_key is not None
