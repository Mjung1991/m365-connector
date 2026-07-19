"""m365_connector — Wiederverwendbares Microsoft 365 Authentifizierungs-Modul."""

from .client import M365Client
from .exceptions import (
    M365AuthError,
    M365Error,
    M365NotFoundError,
    M365RateLimitError,
    M365ServiceError,
    M365ValidationError,
)
from .files import (
    FRAGMENT_EINHEIT,
    SITZUNG_AB_BYTES,
    STANDARD_FRAGMENT,
    FilesService,
    M365PfadFehler,
    fragmente,
    pfad_pruefen,
)
from .oauth import build_admin_consent_url, parse_consent_callback
from .permissions import DEFAULT_PERMISSIONS, PERMISSIONS, Permission
from .storage import CallbackStorage, CredentialStorage, EnvFileStorage
from .subscriptions import validate_subscription_token

__all__ = [
    "M365Client",
    "FilesService",
    "M365PfadFehler",
    "pfad_pruefen",
    "fragmente",
    "FRAGMENT_EINHEIT",
    "STANDARD_FRAGMENT",
    "SITZUNG_AB_BYTES",
    "Permission",
    "PERMISSIONS",
    "DEFAULT_PERMISSIONS",
    "validate_subscription_token",
    "build_admin_consent_url",
    "parse_consent_callback",
    "CredentialStorage",
    "EnvFileStorage",
    "CallbackStorage",
    "M365Error",
    "M365ServiceError",
    "M365AuthError",
    "M365NotFoundError",
    "M365RateLimitError",
    "M365ValidationError",
]
