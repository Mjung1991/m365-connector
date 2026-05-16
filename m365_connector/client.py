"""M365Client — Haupt-Einstiegspunkt für alle Projekte."""

from __future__ import annotations

from pathlib import Path
from typing import Callable

import aiohttp

from .auth import M365Auth
from .calendar import CalendarService
from .credentials import load_credentials
from .mail import MailService
from .subscriptions import SubscriptionService


class M365Client:
    """
    Connects to Microsoft 365 via App-only (Client Credentials).
    Use as async context manager to ensure the HTTP session is properly closed.

    Example:
        async with M365Client.from_env() as client:
            messages = await client.mail.list_inbox(mailbox="bot@firma.de")
    """

    def __init__(self, auth: M365Auth) -> None:
        self._auth = auth
        self._session: aiohttp.ClientSession | None = None
        self.mail = MailService(auth, self._get_session)
        self.calendar = CalendarService(auth, self._get_session)
        self.subscriptions = SubscriptionService(auth, self._get_session)

    def _get_session(self) -> aiohttp.ClientSession:
        """Gibt die gemeinsame HTTP-Session zurück — wird bei Bedarf erstellt.
        Total-Timeout 60s verhindert haengende Calls bei grossen Attachments."""
        if self._session is None or self._session.closed:
            from ._http import make_session
            self._session = make_session()
        return self._session

    @classmethod
    def from_env(cls, env_path: str | Path = ".env") -> "M365Client":
        """Erstellt einen Client aus einer .env-Datei."""
        creds = load_credentials(Path(env_path))
        auth = M365Auth(
            tenant_id=creds["tenant_id"],
            client_id=creds["client_id"],
            client_secret=creds["client_secret"],
        )
        return cls(auth)

    @classmethod
    def from_credentials(
        cls,
        tenant_id: str,
        client_id: str,
        client_secret: str,
    ) -> "M365Client":
        """Erstellt einen Client direkt mit Credentials (z.B. aus Secrets Manager)."""
        auth = M365Auth(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
        )
        return cls(auth)

    async def verify_connection(self) -> dict:
        """Testet die Verbindung — gibt Tenant-Info zurück."""
        token = await self._auth.get_token()
        session = self._get_session()
        async with session.get(
            "https://graph.microsoft.com/v1.0/organization",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "id,displayName,verifiedDomains"},
        ) as resp:
            if resp.status != 200:
                from ._http import to_typed as _to_typed
                raise _to_typed(resp.status, "client.verify_connection")
            data = await resp.json()
            org = data["value"][0] if data.get("value") else {}
            return {
                "tenant_id": org.get("id"),
                "display_name": org.get("displayName"),
                "domains": [d["name"] for d in org.get("verifiedDomains", [])],
            }

    async def close(self) -> None:
        if self._session and not self._session.closed:
            await self._session.close()
        await self._auth.close()

    async def __aenter__(self) -> "M365Client":
        return self

    async def __aexit__(self, *_) -> None:
        await self.close()
