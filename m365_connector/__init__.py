"""m365_connector — Wiederverwendbares Microsoft 365 Authentifizierungs-Modul."""

from .client import M365Client
from .permissions import DEFAULT_PERMISSIONS, PERMISSIONS, Permission
from .subscriptions import validate_subscription_token

__all__ = [
    "M365Client",
    "Permission",
    "PERMISSIONS",
    "DEFAULT_PERMISSIONS",
    "validate_subscription_token",
]
