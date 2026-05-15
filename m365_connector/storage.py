"""Pluggable Credential-Storage für den Setup-Wizard.

Der Wizard ruft `storage.save(...)` mit den final ermittelten Credentials auf.
Projekte können eigene Storage-Implementierungen mitbringen (DB, Secret-Manager,
HashiCorp Vault, etc.) — solange sie das Protocol `CredentialStorage` erfüllen.
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Protocol, runtime_checkable


@runtime_checkable
class CredentialStorage(Protocol):
    """Protocol für Credential-Storage-Implementierungen."""

    def save(
        self,
        *,
        tenant_id: str,
        client_id: str,
        client_secret: str,
    ) -> None:
        """Persistiert die Credentials. Darf Exceptions werfen wenn fehlschlägt."""
        ...


class EnvFileStorage:
    """Default-Storage: schreibt eine .env-Datei (chmod 600)."""

    def __init__(self, path: str | Path = ".env") -> None:
        self.path = Path(path)

    def save(
        self,
        *,
        tenant_id: str,
        client_id: str,
        client_secret: str,
    ) -> None:
        from .credentials import save_credentials

        save_credentials(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            env_path=self.path,
        )


class CallbackStorage:
    """Storage der eine beliebige Callable aufruft.

    Useful für DB-Writer, Secret-Manager-Push, etc.

    Example:
        def write_to_db(tenant_id, client_id, client_secret):
            db.tenants.update(...).set(m365_credentials=...)

        storage = CallbackStorage(write_to_db)
    """

    def __init__(self, callback: Callable[[str, str, str], None]) -> None:
        self.callback = callback

    def save(
        self,
        *,
        tenant_id: str,
        client_id: str,
        client_secret: str,
    ) -> None:
        self.callback(tenant_id, client_id, client_secret)
