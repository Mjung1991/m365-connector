"""Microsoft Graph — Delta-Polling für Mail-Folder.

Delta API liefert nur was sich seit dem letzten Token geändert hat. Konsument persistiert
den deltaLink (z.B. in DB) und übergibt ihn beim nächsten Lauf an `next()`.

Typischer Ablauf:
    messages, link, done = await client.mail.delta.initial(mailbox, folder="inbox")
    while not done:
        more, link, done = await client.mail.delta.next(link)
        messages.extend(more)
    # persistiere `link` als delta_token für nächsten Sync
"""

from __future__ import annotations

from typing import Callable

import aiohttp

from .auth import M365Auth

_GRAPH = "https://graph.microsoft.com/v1.0"


class MailDeltaService:
    """Delta-Polling für Mail-Folder.

    Access via `client.mail.delta.X` — initial, next.
    """

    def __init__(self, auth: M365Auth, get_session: Callable[[], aiohttp.ClientSession]) -> None:
        self._auth = auth
        self._get_session = get_session

    async def initial(
        self,
        mailbox: str,
        folder: str = "inbox",
    ) -> tuple[list[dict], str, bool]:
        """Starts a new delta sync for a folder.

        Returns:
            (messages, link, is_complete):
                - messages: list of changed message objects on this page
                - link: URL for next call (next page) or final delta link (for next sync run)
                - is_complete: True if `link` is the final deltaLink (sync done),
                               False if `link` is a nextLink (more pages to fetch).
        """
        url = f"{_GRAPH}/users/{mailbox}/mailFolders/{folder}/messages/delta"
        return await self._fetch(url)

    async def next(self, link: str) -> tuple[list[dict], str, bool]:
        """Fetches the next page using a nextLink, OR resumes from a previously
        stored deltaLink to get only changes since then.

        Args:
            link: URL from a previous initial()/next() call — either a @odata.nextLink
                  (mid-pagination) or a @odata.deltaLink (resume from last sync).

        Returns:
            (messages, link, is_complete) — see initial() for semantics.
        """
        return await self._fetch(link)

    async def _fetch(self, url: str) -> tuple[list[dict], str, bool]:
        token = await self._auth.get_token()
        async with self._get_session().get(
            url,
            headers={"Authorization": f"Bearer {token}"},
        ) as resp:
            if resp.status != 200:
                raise RuntimeError(f"mail.delta failed ({resp.status})")
            data = await resp.json()
            messages = data.get("value", [])
            if "@odata.nextLink" in data:
                return messages, data["@odata.nextLink"], False
            if "@odata.deltaLink" in data:
                return messages, data["@odata.deltaLink"], True
            raise RuntimeError(
                "mail.delta response missing both @odata.nextLink and @odata.deltaLink"
            )
