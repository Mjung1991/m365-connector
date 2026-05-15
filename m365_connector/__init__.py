"""m365_connector — Wiederverwendbares Microsoft 365 Authentifizierungs-Modul."""

from .client import M365Client
from .oauth import build_admin_consent_url, parse_consent_callback
from .permissions import DEFAULT_PERMISSIONS, PERMISSIONS, Permission
from .storage import CallbackStorage, CredentialStorage, EnvFileStorage
from .subscriptions import validate_subscription_token

__all__ = [
    "M365Client",
    "Permission",
    "PERMISSIONS",
    "DEFAULT_PERMISSIONS",
    "validate_subscription_token",
    "build_admin_consent_url",
    "parse_consent_callback",
    "CredentialStorage",
    "EnvFileStorage",
    "CallbackStorage",
]
