"""Microsoft Graph — Mail-Folder Operationen."""

from __future__ import annotations

from typing import Callable

import aiohttp

from .auth import M365Auth
from ._http import to_typed as _to_typed

_GRAPH = "https://graph.microsoft.com/v1.0"


class MailFolderService:
    """Mail-Folder Operations (mailFolders endpoint).

    Access via `client.mail.folders.X` — list, create, ensure.
    Operates on the user's mailFolders namespace; no calendar/contact folders.
    """

    def __init__(self, auth: M365Auth, get_session: Callable[[], aiohttp.ClientSession]) -> None:
        self._auth = auth
        self._get_session = get_session

    async def list(self, mailbox: str, parent_id: str | None = None) -> list[dict]:
        """Lists top-level folders or child folders of a given parent.

        Args:
            mailbox: Postfach-Adresse oder User-ID.
            parent_id: Wenn gesetzt, listet Kindordner; sonst Top-Level-Ordner.

        Returns:
            Liste der Folder-Objekte mit id, displayName, parentFolderId, totalItemCount, childFolderCount.
        """
        if parent_id:
            url = f"{_GRAPH}/users/{mailbox}/mailFolders/{parent_id}/childFolders"
        else:
            url = f"{_GRAPH}/users/{mailbox}/mailFolders"

        token = await self._auth.get_token()
        async with self._get_session().get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            params={"$top": 100},
        ) as resp:
            if resp.status != 200:
                raise _to_typed(resp.status, "folders.list")
            data = await resp.json()
            return data.get("value", [])

    async def create(self, mailbox: str, name: str, parent_id: str | None = None) -> dict:
        """Creates a new mail folder.

        Args:
            mailbox: Postfach-Adresse oder User-ID.
            name: displayName des neuen Ordners.
            parent_id: Parent-Folder-ID für verschachtelten Ordner; None = Top-Level.

        Returns:
            Folder-Objekt mit id und Metadaten.
        """
        if parent_id:
            url = f"{_GRAPH}/users/{mailbox}/mailFolders/{parent_id}/childFolders"
        else:
            url = f"{_GRAPH}/users/{mailbox}/mailFolders"

        token = await self._auth.get_token()
        async with self._get_session().post(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"displayName": name},
        ) as resp:
            if resp.status != 201:
                raise _to_typed(resp.status, "folders.create")
            return await resp.json()

    async def ensure(self, mailbox: str, name: str, parent_id: str | None = None) -> str:
        """Idempotent: returns existing folder id if a folder with this name exists,
        otherwise creates it and returns the new id.

        Args:
            mailbox: Postfach-Adresse oder User-ID.
            name: displayName des Ordners (case-sensitive Match).
            parent_id: Parent-Folder-ID; None = Top-Level-Suche.

        Returns:
            Folder-ID (string).
        """
        existing = await self.list(mailbox, parent_id=parent_id)
        for folder in existing:
            if folder.get("displayName") == name:
                return folder["id"]
        created = await self.create(mailbox, name, parent_id=parent_id)
        return created["id"]
